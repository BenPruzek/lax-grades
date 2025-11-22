'
' This macro allows the user to generate different types of animations/movies.
'
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
'-----------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' 16-May-2023    : Show 'Delete results' dialog box for a parameter animation with existing results.
' 30-Aug-2022 ube,fde: allow time animation together with structure rotation
' 01-Mar-2021 thn: change logo default position to NorthWest
' 15-Oct-2020 pge: Assure that output files are always named consecutively, otherwise video contains only a single frame if step > 1
' 08-Feb-2020 mwi: moved all ffmpeg-specific functionality into video_creation.lib to allow for a more easy re-use
' 08-Feb-2020 mwi: removed dependency on ImageMagick completely: all functionality can be realized much more efficiently using ffmpeg and ffplay
' 07-Feb-2020 mwi: removed all references to old ImageMagick to simplify the code
' 07-Feb-2020 mwi: removed mng codec - nobody can play this nowadays, added theora codec, correctly set width and height to a multiple of two
'                  or four for some codecs
' 14-May-2019    : new logo
' 06-Sep-2018 mwi: use .bmp instead of .png in palettegen for animated .gif, as ffmpeg crashes reproducably with .png
' 05-Sep-2018 mwi: updated for ImageMagick versions >= 7.0: convert.exe is called magick.exe now (CST-52246)
' 03-Aug-2017 ube: include visible hint on newer Imagemagick version
' 19-Jul-2017 mwi: fixed "All Codecs" option if the old ImageMagick installation was used: windows "move" does not like double double quotes
' 02-Feb-2017 tsi: fixed bash script generation (missing crlf)
' 09-Nov-2016 mwi: added Linux support
' 21-Oct-2016 mwi: removed bmp2avi support as agreed with ube & fwo - those who need the .avi codecs should use the new ImageMagick;
'                  added support for switching between old and new ImageMagick; added special settings dialog to point to own ImageMagick/
'                  FFMPEG installations; allowing more file formats for log file
' 06-Sep-2016 ebu: removed deprecated .Plot command
' 04-Aug-2016 mwi: added several additional codecs; added option to export into all supported codecs for convenience; improved output quality
'                  of videos; merged rescaling and logo stamping step; using ffmpeg (included in newer ImageMagick versions) for almost all codecs;
'                  some general refactoring; added 4K resolution; added option to set video framerate
' 29-Jul-2016 fwo: added initial mp4 export implementation, requires updated ImageMagick ImageMagick-7.0.2-5
' 25-Jul-2016 mwi: added "Rescale" option to export images directly in the requested size to avoid any rescaling during the video generation
' 22-Jul-2016 mwi: fixed code for generating videos from pic phase space monitors, set fixed axis limits to avoid flicker
' 14-Dec-2015 ube: macro was limited to 999 images, now increased to 99999 images per video
' 29-Oct-2015 ama,ube: make sure, files to be deleted, are not read-only
' 24-Oct-2015 ube: subroutine Start renamed to Start_LIB
' 21-Oct-2015 ube: Preview button did not work for spaces in the project path or filename
' 28-Aug-2015 tgl: Added chcp to batch file to be more robust with umlauts
' 08-Jul-2015 fsr: Replaced obsolete "AutoScale" with "AutoRange"
' 04-Jun-2015 fsr: Certain file names would prevent the macro from starting in rare instances, fixed
' 04-Feb-2015 fsr: Bugfixes; rotation option now uses drop down list instead of radio buttons;
'					new "Rotation step ratio" option for slower/faster rotation than phase animation
' 17-Dec-2014 fsr: Fixed an issue with files located on unmapped network drives (\\server\...)
' 11-Aug-2014 fwo,ube: runandwait (filecopy, etc) did not work with special filenames
' 18-Nov-2013 fsr: Using project temp folder now; clarified where frames are stored for "Frames Only" option
' 10-Oct-2013 fsr: Some GUI clean-up and improvements
' 09-Oct-2013 jwa: Option to create a video of a marker moving along a 1D plot (e.g., for correlation with transient 2d/3d plot)
' 09-Oct-2013 fsr: PIC phase space frames can now be exported into a video
' 22-Jan-2013 fde: added white background as option
' 12-Jul-2012 msc: new feature of storing .bmp files only (useful for external movie generation, e.g. HD video)
' 14-Mar-2012 ube: adding option  +dither  for improved look of legend and colorbar
' 23-Mar-2011 ube: compressed dialogue, so that dialogue also fits on laptop with beamer in 1024x768 resolution
' 23-Mar-2011 fde: Fixed/Check phase/time input concerning consistency and added Revert Video
' 09-Nov-2009 fde: Fixed crash when no parameter is defined and parameter animation selected (PR 12325)
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 17-Jun-2008 ube,ala: new logo CST_logo.png
' 26-Jun-2007 fde: included "wait 1" for time animations (should remove timing problems)
' 18-Jun-2007 ube: created graphics file stored at same level as cst-file
' 02-Jan-2007 fde: Included structure animation
' 21-Oct-2005 imu: Included into Online Help
' 19-Nov-2004 ube,btr: temporary replacement of enum-construction by constants
' 12-Dec-2003 ube: ability to switch off "inserting logo" via public constant
' 21-Oct-2003 ube: bmp2avi has many problems with crashes, therefore disable those 3 formats with Public Constant
' 29-Jul-2003 ube: bugfix for MPEG-time movie
' 29-Jul-2003 ube: bmp2avi.exe link now to CST's support page  MWS_FAQs
' 05-Jun-2003 ube: new www-link for bmp2avi.exe
' 15-Nov-2002 ube: time movies included
'-----------------------------------------------------------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_Graphics_Save_Video"
Const cst_macroname = "Graphics~Save_Video"

'#include "vba_globals_all.lib"
'#include "video_creation.lib"

'-----------------------------------------------------------------------------------------------------------------------------------
Public cst_DoInsertLogo As Boolean
Public cst_DoAllCodecs As Boolean
Public cst_projectdir As String
Public cst_projectname As String
Public cst_tempdir As String
Public cst_templatedir As String
Public cst_ffmpeg As String         ' full path to ffmpeg.exe (including file name)
Public cst_ffplay As String         ' full path to ffplay.exe (including file name)
Public cst_parameter() As String
Public cst_phi As Single, cst_phi0 As Single, cst_phi1 As Single, cst_dphi As Single
Public cst_paralow As Double, cst_parahigh As Double, cst_parastep As Double, cst_para As Double
Public cst_pbparameter As Boolean
Public cst_revert_video As Boolean
Public cst_white_background As Boolean
Public cst_GradientSetting As Boolean
Public cst_rc As String, cst_gc As String, cst_bc As String
Public cst_codec As Integer, cst_size As Integer, cst_resize As Integer, cst_framerate As String
Public cst_logo_filename As String, cst_logo_position As Integer, cst_logo_border As Integer
Public cst_videofile As String
Public cst_NoProcess As Boolean, cst_KeepTemp As Boolean
Public cst_sPlotDomain As String
Public cst_bPlotDomain3D As Boolean
Public cst_sPlotParameter As String
Public cst_it As Integer, cst_itlow As Integer, cst_ithigh As Integer, cst_itstep As Integer
Public cst_NumberOfPICFrames As Long, cst_NumberOfCurvesInPlot As Long
Public cst_OKWasPressed As Boolean

Const RotationOptions = Array("None", "X or U axis", "Y or V axis", "Z or W axis", "Left", "Right", "Up", "Down")

'-----------------------------------------------------------------------------------------------------------------------------

Sub Main ()

	cst_templatedir	 = MakeWindowsPath(GetInstallPath() + "\Library\Misc\templates")

    GetFFMPegPath(cst_ffmpeg, cst_ffplay)

	cst_tempdir = MakeWindowsPath(GetProjectPath("Temp") + "Save_Video\") ' see CST-46286
	If (Dir$(cst_tempdir, vbDirectory) = "") Then
		MkDir(cst_tempdir)
	Else
		DeleteFilesWithPattern(cst_tempdir, "*.*")
	End If

	' Fill Parameter List
	ReDim cst_parameter(GetNumberOfParameters())
	Dim cst_index As Integer
	For cst_index = 0 To (GetNumberOfParameters -1)
		cst_parameter(cst_index) = GetParameterName(cst_index)
	Next cst_index

	Dim s1 As String
	s1 = MakeWindowsPath(GetProjectPath("Project")) ' see CST-46286
	cst_projectdir	= DirName(s1)
	cst_projectname = ShortName(s1)

	Dim cst_videofile_basename As String
	cst_videofile_basename = CreateFileBaseName(cst_projectdir, cst_projectname)

	cst_sPlotDomain = IdentifyPlotDomain(cst_NumberOfPICFrames, cst_NumberOfCurvesInPlot, cst_bPlotDomain3D)
	cst_sPlotParameter = ""

	Begin Dialog UserDialog 0,0,540,465,"Save Video",.DialogFunc ' %GRID:10,3,1,1
		GroupBox 10,6,520,204,"Video Settings",.BoxAnimation
		CheckBox 420,66,90,15,"All Codecs",.AllCodecs
		Text 20,21,80,14,"Filename",.LabelTarget
		TextBox 20,36,340,21,.Target
		PushButton 370,36,90,21,"Browse...",.BrowseTarget
		Text 20,68,90,14,"Video Codec:",.LabelCodec
		DropListBox 140,63,270,120,cst_codec_description(),.Codec
		Text 20,90,90,15,"Video Size:",.SizeLabel
		DropListBox 140,87,270,120,cst_video_default_format(),.Size
		OptionGroup .Resize
			OptionButton 30,112,70,14,"Crop",.OptionButton1
			OptionButton 110,112,70,14,"Fit",.OptionButton2
			OptionButton 190,112,70,14,"Distort",.OptionButton3
			OptionButton 270,112,70,14,"Ignore",.OptionButton4
			OptionButton 350,111,80,15,"Rescale",.OptionButton5
		Text 30,162,100,15,"Framerate [1/s]:",.Text3
		TextBox 140,159,90,21,.FrameRate
		CheckBox 240,161,210,14,"Bounce Playback At End",.BounceVideoPlaybackAtEnd
		Text 30,186,100,15,"Expert options:",.Text1
		TextBox 140,183,380,21,.ExpertOptions
		GroupBox 10,213,350,111,"",.BoxLogo
		CheckBox 20,213,60,15,"Logo",.EnableLogo
		Text 20,234,30,15,"File:",.LabelSource
		TextBox 50,231,200,21,.Source
		PushButton 260,231,90,21,"Browse...",.BrowseSource
		GroupBox 20,252,100,66,"Position",.BoxPosition
		OptionGroup .Position
			OptionButton 30,267,20,15,"",.NorthWest
			OptionButton 30,282,20,15,"",.West
			OptionButton 30,297,20,15,"",.SouthWest
			OptionButton 60,297,20,15,"",.South
			OptionButton 90,297,20,15,"",.SouthEast
			OptionButton 90,282,20,15,"",.East
			OptionButton 90,267,20,15,"",.NorthEast
			OptionButton 60,267,20,15,"",.North
		Text 130,274,130,15,"Distance to border:",.LabelBorder
		TextBox 260,271,90,21,.Border
		GroupBox 10,330,520,105,"Phase / Time / x-Value / Frame number / Parameter",.GroupBox1
		Text 20,357,90,15,"Start [deg]",.LabelStart
		TextBox 120,351,130,21,.StartValue
		Text 20,384,90,15,"Stop [deg]",.LabelStop
		TextBox 120,378,130,21,.StopValue
		Text 20,411,90,15,"Step [deg]",.LabelStep
		TextBox 120,407,130,21,.StepValue
		GroupBox 370,213,160,111,"3D Model Rotation",.GroupBox3
		DropListBox 380,231,140,168,RotationOptions(),.RotationDLB
		Text 380,258,120,24,"Step ratio to phase/time step:",.Text2
		TextBox 380,288,140,21,.RotationStepRatioT
		DropListBox 300,378,220,78,cst_parameter(),.ParameterList
		CheckBox 280,357,160,15,"Animate parameter:",.Parameter_CheckBox1
		CheckBox 420,90,100,15,"White Bkg.",.GenwhiteBk
		CheckBox 300,402,190,15,"Include reverse playback",.PBParameter_CheckBox1
		CheckBox 30,138,120,15,"Reverse Video",.RevertVideo
		CheckBox 190,138,150,15,"Frame Export Only",.GenFramesOnly
		CheckBox 360,138,160,15,"Keep Temporary Files",.KeepTemp
		OKButton 200,438,80,21
		PushButton 360,438,80,21,"Preview",.Preview
		CancelButton 280,438,80,21
		PushButton 440,438,80,21,"Help",.Help
	End Dialog

	cst_codec	         = Max(Min(GetRegDouble("CST STUDIO SUITE", "Animation", "codec", GIF), LASTCODEC), GIF)
	cst_videofile        = cst_videofile_basename + "." + cst_codec_default_extension(cst_codec)
	cst_DoAllCodecs      = GetString("CST STUDIO SUITE", "Animation", "all_codecs", "False") = "True"
	cst_size	         = GetRegDouble("CST STUDIO SUITE", "Animation", "size", 3)
	cst_resize	         = GetRegDouble("CST STUDIO SUITE", "Animation", "resize", 3)
	cst_framerate        = GetString("CST STUDIO SUITE", "Animation", "framerate", "24")
	cst_DoInsertLogo     = GetString("CST STUDIO SUITE", "Animation", "insertLogo", "True") = "True"
	cst_white_background = GetString("CST STUDIO SUITE", "Animation", "whiteBKG", "False") = "True"
	cst_revert_video     = GetString("CST STUDIO SUITE", "Animation", "revertVideo", "False") = "True"
	cst_NoProcess        = GetString("CST STUDIO SUITE", "Animation", "noProcess", "False") = "True"
	cst_KeepTemp         = GetString("CST STUDIO SUITE", "Animation", "keepTemp", "False") = "True"
	cst_logo_filename    = GetString("CST STUDIO SUITE", "Animation", "logo_filename", "SIMULIA_CST_Studio_Suite.png")
	cst_logo_position    = GetRegDouble("CST STUDIO SUITE", "Animation", "position", 0)
	cst_logo_border	     = GetRegDouble("CST STUDIO SUITE", "Animation", "border", 10)
	cst_phi0	         = GetRegDouble("CST STUDIO SUITE", "Animation", "StartValue", 0)
	cst_phi1	         = GetRegDouble("CST STUDIO SUITE", "Animation", "StopValue", 360)
	cst_dphi	         = GetRegDouble("CST STUDIO SUITE", "Animation", "StepValue", 5)

	Dim dlg As UserDialog
	If (Dialog(dlg) >= 0) Then Exit All

	cst_videofile        = MakeWindowsPath(FullPath(dlg.Target, cst_projectdir))
	cst_codec	         = dlg.Codec
	cst_size	         = dlg.Size
	cst_resize	         = dlg.Resize
	cst_framerate        = dlg.FrameRate
	cst_DoInsertLogo     = dlg.EnableLogo
	cst_white_background = dlg.GenwhiteBk
	cst_NoProcess        = dlg.GenFramesOnly
	cst_KeepTemp         = dlg.KeepTemp
	cst_revert_video     = dlg.RevertVideo
	cst_DoAllCodecs      = dlg.AllCodecs
	cst_logo_filename    = MakeWindowsPath(FullPath(dlg.Source, cst_templatedir))
	cst_logo_position    = dlg.position
	cst_logo_border	     = CInt(dlg.border)

	SaveInteger("CST STUDIO SUITE", "Animation", "codec", dlg.codec)
	SaveString ("CST STUDIO SUITE", "Animation", "all_codecs", IIf(dlg.AllCodecs, "True", "False"))
	SaveInteger("CST STUDIO SUITE", "Animation", "size", dlg.size)
	SaveInteger("CST STUDIO SUITE", "Animation", "resize", dlg.resize)
	SaveInteger("CST STUDIO SUITE", "Animation", "framerate", dlg.FrameRate)
	SaveString ("CST STUDIO SUITE", "Animation", "insertLogo", IIf(dlg.EnableLogo, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "whiteBKG", IIf(dlg.GenwhiteBk, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "revertVideo", IIf(dlg.RevertVideo, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "noProcess", IIf(dlg.GenFramesOnly, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "keepTemp", IIf(dlg.KeepTemp, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "logo_filename", dlg.Source)
	SaveInteger("CST STUDIO SUITE", "Animation", "position", dlg.position)
	SaveString ("CST STUDIO SUITE", "Animation", "border", dlg.border)
	If (InStr("PICPhaseSpace|1DResult", cst_sPlotDomain)=0) Then ' do not save these for 1D result plots or PIC phase space monitors
		SaveString("CST STUDIO SUITE", "Animation", "StartValue", dlg.startvalue)
		SaveString("CST STUDIO SUITE", "Animation", "StopValue", dlg.stopvalue)
		SaveString("CST STUDIO SUITE", "Animation", "StepValue", dlg.stepvalue)
	End If
    
	If cst_NoProcess Then
		ReportInformationToWindow("Save Video: Saved frames in '" + cst_tempdir + "'.")
    Else
        CreateVideoFromImageSequence(MakeNativePath(cst_tempdir), _
                      IIf(IsWindows(), "image_%%05d.bmp", "image_%05d.bmp"), _
                      MakeNativePath(cst_ffmpeg), _
                      MakeNativePath(cst_ffplay), _
                      cst_video_default_format(cst_size), _
                      cst_resize, _
                      cst_framerate, _
                      cst_DoInsertLogo, _
                      MakeNativePath(cst_logo_filename), _
                      cst_logo_position, _
                      cst_logo_border, _
                      cst_codec, _
                      cst_DoAllCodecs, _
                      MakeNativePath(cst_videofile), _
                      cst_revert_video, _
                      dlg.BounceVideoPlaybackAtEnd, _
                      dlg.ExpertOptions, _
                      False)
		ReportInformationToWindow("Finished creating video(s).")
    End If

    If Not cst_KeepTemp Then
        DeleteFilesWithPattern(cst_tempdir, "*.*")
    End If

    If Not cst_DoAllCodecs Then
		Start_LIB(MakeNativePath(cst_videofile))
	End If
End Sub

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

	Dim cst_filename As String, cst_extension As String, cst_index As Integer
	Dim cst_iii As Integer
	Dim oTempObject As Object
	Dim nFrame As Long
	Dim dRotAngle As Double

    Dim cst_videowidth As Integer, cst_videoheight As Integer
    ParseSize(DlgText("Size"), cst_videowidth, cst_videoheight)

	Select Case Action
		Case 1 ' Dialog box initialization
			cst_OKWasPressed = False
			DlgText("Target", cst_videofile)
			cst_codec = Max(Min(cst_codec, UBound(cst_codec_description)), LBound(cst_codec_description))
			DlgValue("Codec", cst_codec)
			DialogFunc("Codec", 2, cst_codec)
			DlgValue("Size", cst_size)
			DlgValue("Resize", cst_resize)
			DlgText("FrameRate", cst_framerate)
			DlgValue("EnableLogo", cst_DoInsertLogo)
			DialogFunc("EnableLogo", 2, cst_DoInsertLogo)
			DlgValue("AllCodecs", cst_DoAllCodecs)
			DialogFunc("AllCodecs", 2, cst_DoInsertLogo)
			DlgValue("GenwhiteBk", cst_white_background)
			DlgValue("RevertVideo", cst_revert_video)
			DlgValue("GenFramesOnly", cst_NoProcess)
			If cst_sPlotDomain="time" Then
				DlgText("StartValue", "1")
				DlgText("StopValue", CStr(Plot2D3D.GetNumberOfSamples))
				DlgText("StepValue", "1")
			Else
				DlgText("StartValue", CStr(cst_phi0))
				DlgText("StopValue", CStr(cst_phi1))
				DlgText("StepValue", CStr(cst_dphi))
			End If
			DlgText("Source", cst_logo_filename)
			DlgValue("Position", cst_logo_position)
			DlgText("Border", CStr(cst_logo_border))
		    DlgValue("RotationDLB", 0)
		    DlgValue("Parameter_CheckBox1", 0)
			DlgText("ExpertOptions", cst_codec_default_option(CInt(DlgValue("Codec"))))
			DlgEnable("Size", IIf(DlgValue("Resize") = 3, 0, 1))
            DlgEnable("RotationStepRatioT", IIf(DlgValue("RotationDLB") = 0, 0, 1))
            DlgText("RotationStepRatioT", "0.2")
            cst_phi0 = 0
            cst_phi1 = 0
            cst_dphi = 5
			cst_rc = Plot.GetBackgroundColorR()
			cst_gc = Plot.GetBackgroundColorG()
			cst_bc = Plot.GetBackgroundColorB()
			cst_GradientSetting = Plot.GetGradientBackground() 'check and store if gradient

			If GetNumberOfParameters = 0 Then
				DlgEnable("Parameter_CheckBox1", 0)
			End If
			If cst_sPlotDomain="time" Then
				DlgText("LabelStart", "Start sample")
				DlgText("LabelStop", "Stop sample")
				DlgText("LabelStep", "Step")
			ElseIf cst_sPlotDomain = "PICPhaseSpace" Then
				DlgText("LabelStart", "Start frame")
				DlgText("StartValue", "1")
				DlgText("LabelStop", "Stop frame")
				DlgText("StopValue", CStr(cst_NumberOfPICFrames-1))
				DlgText("LabelStep", "Step")
				DlgText("StepValue", "1")
			ElseIf cst_sPlotDomain = "1DResult" Then
				If (GetFileType(Plot1D.GetCurveFileName(0))="complex") Then
					Set oTempObject = Result1DComplex(Plot1D.GetCurveFileName(0))
				Else
					Set oTempObject = Result1D(Plot1D.GetCurveFileName(0))
				End If
				DlgText("LabelStart", "x-Start")
				DlgText("StartValue", CStr(oTempObject.GetX(0))) ' assumes sorted by x value
				DlgText("LabelStop", "x-Stop")
				DlgText("StopValue", CStr(oTempObject.GetX(oTempObject.GetN()-1))) ' assumes sorted by x value
				DlgText("LabelStep", "x-Step")
				DlgText("StepValue", CStr((oTempObject.GetX(oTempObject.GetN()-1)-oTempObject.GetX(0))/100))
			Else
				'
			End If
			DlgEnable("ParameterList", 0)
			DlgEnable("PBParameter_Checkbox1", 0)
			cst_pbparameter = False

		Case 2 ' Value changing or button pressed
			DialogFunc = True

			Select Case Item
 	          	Case "RotationDLB"
            		DlgEnable("RotationStepRatioT", IIf(DlgValue("RotationDLB") = 0, 0, 1))
 				Case "BrowseTarget"
					cst_extension = cst_codec_default_extension(DlgValue("Codec"))
					cst_filename  = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
					cst_filename  = MakeWindowsPath(GetFilePath(ShortName(cst_filename), cst_extension, DirName(cst_filename), "Save Animation As", 3))
					If (cst_filename <> "") Then
						DlgText("Target", ShortPath(cst_filename, cst_projectdir))
					End If
				Case "Codec"
					cst_extension = IIf(cst_DoAllCodecs, "*", cst_codec_default_extension(Value))
					cst_filename  = DlgText("Target")
					cst_index	  = InStrRev(cst_filename, ".")
					DlgText("Target", Left$(cst_filename, cst_index) + cst_extension)
					DlgText("ExpertOptions", cst_codec_default_option(CInt(DlgValue("Codec"))))
				Case "Resize"
					DlgEnable("Size", IIf(DlgValue("Resize") = 3, 0, 1))
				Case "AllCodecs"
					cst_DoAllCodecs = DlgValue("AllCodecs")
					DlgEnable("Codec", Not cst_DoAllCodecs)
					DlgEnable("LabelCodec", Not cst_DoAllCodecs)
					DlgEnable("Text1", Not cst_DoAllCodecs)
					DlgEnable("ExpertOptions", Not cst_DoAllCodecs)
					DialogFunc("Codec", 2, cst_codec)
				Case "EnableLogo"
					cst_DoInsertLogo = DlgValue("EnableLogo")
					DlgEnable("BoxLogo", cst_DoInsertLogo)
					DlgEnable("LabelSource", cst_DoInsertLogo)
					DlgEnable("Source", cst_DoInsertLogo)
					DlgEnable("BrowseSource", cst_DoInsertLogo)
					DlgEnable("BoxPosition", cst_DoInsertLogo)
					DlgEnable("Position", cst_DoInsertLogo)
					DlgEnable("LabelBorder", cst_DoInsertLogo)
					DlgEnable("Border", cst_DoInsertLogo)
				Case "BrowseSource"
					cst_extension = "Image Files|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tga;*.tiff"
					cst_filename  = MakeWindowsPath(FullPath(DlgText("Source"), cst_templatedir))
					cst_filename  = MakeWindowsPath(GetFilePath(ShortName(cst_filename), cst_extension, DirName(cst_filename), "Choose logo file", 0))
					If (cst_filename <> "") Then
						DlgText("Source", ShortPath(cst_filename, cst_templatedir))
					End If
				Case "Parameter_CheckBox1"
					If (DlgValue "Parameter_CheckBox1") Then
						DlgEnable "ParameterList", 1
						DlgEnable "PBParameter_Checkbox1", 1
						DialogFunc("ParameterList", 2, DlgValue("ParameterList"))
					Else
						DlgEnable "ParameterList", 0
						cst_sPlotParameter = ""
						DlgEnable "PBParameter_Checkbox1", 0
					End If

				Case "ParameterList"
				   cst_sPlotParameter = cst_parameter(DlgValue("ParameterList"))
				   DlgText("StartValue", RestoreParameter(cst_sPlotParameter))
				   DlgText("StepValue", CStr(0.5*RealVal(RestoreParameter(cst_sPlotParameter))))
				   DlgText("StopValue", CStr(5*RealVal(RestoreParameter(cst_sPlotParameter))))

				Case "Help"
					StartHelp(HelpFileName)
				Case "Cancel"
					If cst_white_background Then 'reset before exit
						Plot.SetGradientBackground(cst_GradientSetting)
						Plot.SetBackgroundColor(cst_rc, cst_gc, cst_bc)
					End If
					Plot2D3D.PhaseValue 0
					DialogFunc = False ' close dialog
				Case "Preview"
					Dim cst_tempfile As String, cst_resize As Integer
					cst_tempfile = cst_tempdir + "preview.bmp"
                    cst_videofile = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
                    cst_codec = DlgValue("codec")
                    cst_DoAllCodecs = DlgValue("AllCodecs")
					cst_resize = DlgValue("resize")
                    ExportImage(cst_resize, cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
                    CreateVideoFromImageSequence(MakeNativePath(cst_tempdir), _
                                  MakeNativePath(cst_tempfile), _
                                  MakeNativePath(cst_ffmpeg), _
                                  MakeNativePath(cst_ffplay), _
                                  DlgText("size"), _
                                  cst_resize, _
                                  DlgText("FrameRate"), _
                                  cst_DoInsertLogo, _
                                  MakeNativePath(FullPath(DlgText("Source"), cst_templatedir)), _
                                  DlgValue("position"), _
                                  CInt(DlgText("Border")), _
                                  cst_codec, _
                                  cst_DoAllCodecs, _
                                  MakeNativePath(cst_videofile), _
                                  False, _
                                  False, _
                                  DlgText("ExpertOptions"), _
                                  True)

				Case "OK"
					Dim Input_ok As Boolean
					Input_ok = False

					cst_revert_video = DlgValue("RevertVideo")
					cst_white_background = DlgValue("GenwhiteBk")

					If cst_sPlotParameter = "" Then
						If cst_sPlotDomain="time" Then
							cst_itlow = CInt(DlgText("StartValue"))
							cst_ithigh= CInt(DlgText("StopValue"))
							cst_itstep= CInt(DlgText("StepValue"))
							'Check input
							If ((cst_ithigh <= cst_itlow) Or (cst_itstep <=0)) Then
								MsgBox "Check Time Settings"
								cst_ithigh = cst_itlow
								Input_ok = False
							Else
								cst_it = cst_itlow
								Input_ok = True
							End If
						Else
							cst_phi0 = RealVal(DlgText("StartValue"))
							cst_phi1 = RealVal(DlgText("StopValue"))
							cst_dphi = RealVal(DlgText("StepValue"))
							'Check input
							If ((cst_phi1 <= cst_phi0) Or (cst_dphi <= 0)) Then
								MsgBox "Check Phase Settings"
								cst_phi1 = cst_phi0
								Input_ok = False
							Else
								cst_phi	 = cst_phi0
								Input_ok = True
							End If
						End If
					Else
						cst_paralow = RealVal(DlgText("StartValue"))
						cst_parahigh = RealVal(DlgText("StopValue"))
						cst_parastep = RealVal(DlgText("StepValue"))
						'Check input
						If ((cst_parahigh <= cst_paralow) Or (cst_parastep <= 0)) Then
							MsgBox "Check Parameter Settings"
							cst_parahigh = cst_paralow
							Input_ok = False
						Else
							cst_para = cst_paralow
							Input_ok = True
						End If
						cst_pbparameter = DlgValue("PBParameter_CheckBox1")
						
						DIM bparameter As Boolean
						bparameter = DlgValue("Parameter_CheckBox1")
						if (bparameter) Then
							If (AskForDeleteResultsOnParameterChange()) Then
								Input_ok = True
							Else
								Input_ok = False
							End If
						End If
					End If

					If Input_ok Then
						' ---Set Bk to white--
						If cst_white_background Then 'set to white
							Plot.SetGradientBackground(False)
							Plot.SetBackgroundColor("1", "1", "1")
						End If

						' ---Disable all dialog items (except for the Cancel button)---
						For cst_index = 0 To DlgCount()-1
							On Error Resume Next
								DlgEnable cst_index, IIf(DlgText(cst_index) = "Cancel", 1, 0)
							On Error GoTo 0
						Next

						cst_OKWasPressed = True ' this triggers the work in the dialog's idle event
					End If
			End Select
		Case 3 ' ComboBox or TextBox Value changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			If Not cst_OKWasPressed Then
				Wait(0.2) ' required on Linux to prevent bogging the system
				DialogFunc = True
				Exit Function
			End If

			If cst_sPlotParameter = "" Then 'parameter animation?
				If cst_sPlotDomain="time" Then
					If (cst_ithigh > cst_itlow) Then
						If (cst_it <= cst_ithigh) Then
							cst_index = (cst_it - cst_itlow)/cst_itstep
							cst_tempfile = cst_tempdir + "image_" + Format(cst_index, "00000") + ".bmp"
							DlgText "StartValue", CStr(cst_it)
							Plot2D3D.SetSample cst_it

							dRotAngle = CDBl(DlgText("StepValue"))*CDbl(DlgText("RotationStepRatioT"))
							HandleStructureRotation(DlgText("RotationDLB"), dRotAngle)

							Plot.Update
							ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
							cst_it = cst_it + cst_itstep
						ElseIf (cst_it > cst_ithigh) Then
							DlgText "StartValue", CStr(cst_itlow)
							If cst_white_background Then 'reset before exit
                                Plot.SetGradientBackground(cst_GradientSetting)
                                Plot.SetBackgroundColor(cst_rc, cst_gc, cst_bc)
							End If
							DlgEnd -1
						End If
					End If
				ElseIf (cst_sPlotDomain = "PICPhaseSpace") Then
					Dim nStartFrame As Integer, nEndFrame As Integer, nFrameStep As Integer
					Dim sResultPath As String

					If (cst_phi1 > cst_phi0) Then

						nStartFrame = DlgText("StartValue")
						nEndFrame = DlgText("StopValue")
						nFrameStep = Fix(DlgText("StepValue")+0.1)
						If (nFrameStep <= 0) Then nFrameStep = 1

						sResultPath = GetSelectedTreeItem()

						' identify min/max x and y values across all graphs
						Dim minX As Double, maxX As Double, minY As Double, maxY As Double
						minX = +1.e127
						minY = +1.e127
						maxX = -1.e127
						maxY = -1.e127

						DlgEnable("Cancel", False)

						nFrame = nStartFrame
						While (nFrame <= nEndFrame)
                            DlgText("StartValue", CStr(nFrame) + " search min/max")

							Dim sFileName As String
							sFileName = Resulttree.GetFileFromTreeItem( sResultPath + "\Frame " + Format(nFrame, "0000") )

							Dim oResult As Object
							Set oResult = Result1D( sFileName )
							oResult.SortByX()

							minX = Min(oResult.GetX(0), minX)
							maxX = Max(oResult.GetX(oResult.GetN()-1), minX)

							minY = Min(oResult.GetY(oResult.GetGlobalMinimum()), minY)
							maxY = Max(oResult.GetY(oResult.GetGlobalMaximum()), maxY)

							nFrame = nFrame + nFrameStep
						Wend

						nFrame = nStartFrame
						While (nFrame <= nEndFrame And SelectTreeItem(sResultPath + "\Frame " + Format(nFrame, "0000")))
							cst_tempfile = cst_tempdir + "image_" + Format((nFrame - nStartFrame)/nFrameStep, "00000") + ".bmp"

							With Plot1D
								.SetLineColor(0, 0, 0, 255) ' set points to blue
								.XRange(minX, maxX)
								.YRange(minY, maxY)
								ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
								.ResetView
								.RemoveLineColor(0)
								.XAutoRange(True)
								.YAutoRange(True)
							End With

							DlgText "StartValue", CStr(nFrame)
							nFrame = nFrame + nFrameStep
						Wend

						DlgEnd -1
					End If
				ElseIf (cst_sPlotDomain = "1DResult") Then
					Dim dXValue As Double, dXStart As Double, dXStop As Double, dXStep As Double, nCurveIndex As Long

					If (cst_phi1 > cst_phi0) Then

						dXStart = DlgText("StartValue")
						dXStop = DlgText("StopValue")
						dXStep = DlgText("StepValue")
						If (dXStep <= 0) Then dXStep = (dXStop-dXStart)/100

						dXValue = dXStart
						nFrame = 0

						DlgEnable("Cancel", False)

						While (dXValue <= dXStop+0.1*dXStep)
							Plot1D.DeleteAllMarker
							For nCurveIndex = 0 To cst_NumberOfCurvesInPlot-1
								Plot1D.AddMarkerToCurve(dXValue, nCurveIndex)
							Next
							cst_tempfile = cst_tempdir + "image_" + Format(nFrame, "00000") + ".bmp"
							ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
							DlgText "StartValue", CStr(dXValue)
							dXValue = dXValue + dXStep
							nFrame = nFrame + 1
						Wend

						DlgEnd -1
					End If
				Else  '	 If sPlotDomain="frequency" Or DlgValue("RotationDLB")=0
					If (cst_phi1 > cst_phi0) Then
						If (cst_phi < cst_phi1) Then
							cst_index = CInt(0.1 + (cst_phi - cst_phi0) / cst_dphi) ' Warning: Bankers rounding !!!
							cst_tempfile = cst_tempdir + "image_" + Format(cst_index, "00000") + ".bmp"
							DlgText "StartValue", CStr(cst_phi)
							Plot2D3D.PhaseValue cst_phi-Fix(cst_phi/360)*360

							dRotAngle = CDBl(DlgText("StepValue"))*CDbl(DlgText("RotationStepRatioT"))
							HandleStructureRotation(DlgText("RotationDLB"), dRotAngle)

							Plot.Update
							ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
							cst_phi = cst_phi + cst_dphi
						ElseIf (cst_phi >= cst_phi1) Then
							DlgText "StartValue", CStr(cst_phi0)
							Plot2D3D.PhaseValue 0
							If cst_white_background Then 'reset before exit
							Plot.SetGradientBackground(cst_GradientSetting)
							Plot.SetBackgroundColor(cst_rc, cst_gc, cst_bc)
							End If
							DlgEnd -1
						End If
					End If	 'If (cst_phi1 > cst_phi0) Then
				End If '  If sPlotDomain="frequency" Or DlgValue("RotationDLB")>0
				Else	'Parameter Sweep Animation
				If (cst_parahigh > cst_paralow) Then
					If (cst_para < cst_parahigh) Then
						StoreParameter(cst_sPlotParameter, CStr(cst_para))
						RebuildForParametricChange
						cst_index = CInt((cst_para-cst_paralow)/cst_parastep)
						'cst_index = CInt(0.1 + cst_para /cst_parastep) ' Warning: Bankers rounding !!!
						cst_tempfile = cst_tempdir + "image_" + Format(cst_index, "00000") + ".bmp"
						Plot.Update
						ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, cst_bPlotDomain3D)
						cst_para = cst_para + cst_parastep
						DlgText "StartValue", CStr(cst_para)
					ElseIf (cst_para >= cst_parahigh) Then
						DlgText "StartValue", CStr(cst_para)
					End If

				If (cst_para >= cst_parahigh) Then
					cst_index = CInt((cst_para-cst_paralow)/cst_parastep)
					'Copy Files for Playback..
					If cst_pbparameter Then
						For cst_iii = 1 To cst_index
							FileCopy cst_tempdir + "image_" + Format(cst_index-cst_iii, "00000") + ".bmp", _
										cst_tempdir + "image_" + Format(cst_index+cst_iii, "00000") + ".bmp"
						Next cst_iii
					End If
					If cst_white_background Then 'reset before exit
						Plot.SetGradientBackground(cst_GradientSetting)
						Plot.SetBackgroundColor(cst_rc, cst_gc, cst_bc)
					End If
					DlgEnd -1
				End If
			End If

		End If	'Parameter Sweep Animation
		Wait(0.1)
		DialogFunc = True
	End Select

End Function
Sub HandleStructureRotation(sRotType As String, dRotAngle As Double)

	Select Case sRotType
		Case "None"
			' Do nothing
		Case "Up"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "up"
		Case "Y or V axis"		'Yaxis
			Plot.RotationAngle 20
			Plot.Rotate "up"
			Plot.RotationAngle 30
			Plot.Rotate "right"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "right"
			Plot.RotationAngle 30
			Plot.Rotate "left"
			Plot.RotationAngle 20
			Plot.Rotate "down"
		Case "Left"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "left"
		Case "Right"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "right"
		Case "Z or W axis" 		'Zaxis
			Plot.RotationAngle 20
			Plot.Rotate "up"
			Plot.RotationAngle 30
			Plot.Rotate "right"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "counterclockwise"
			Plot.RotationAngle 30
			Plot.Rotate "left"
			Plot.RotationAngle 20
			Plot.Rotate "down"
		Case "Down"
			Plot.RotationAngle 20: Plot.Rotate "up"
			Plot.Rotate "down"
		Case "X or U axis"		'Xaxis
			Plot.RotationAngle 20
			Plot.Rotate "up"
			Plot.RotationAngle 30
			Plot.Rotate "right"
			Plot.RotationAngle dRotAngle
			Plot.Rotate "down"
			Plot.RotationAngle 30
			Plot.Rotate "left"
			Plot.RotationAngle 20
			Plot.Rotate "down"
	End Select

End Sub

Function IdentifyPlotDomain(ByRef NumberOfPICFrames As Long, NumberOfCurvesInPlot As Long, ByRef Is3DPlot As Boolean) As String
	If (InStr(GetSelectedTreeItem(), "PIC Phase Space Monitor")>0) Then
		IdentifyPlotDomain = "PICPhaseSpace"
        Is3DPlot = False

		' If an actual frame is selected, go up one level in the tree to select the folder
		If (InStr(GetSelectedTreeItem, "Frame")>0) Then
			SelectTreeItem(Left(GetSelectedTreeItem, InStrRev(GetSelectedTreeItem, "\")))
		End If

		NumberOfPICFrames = 0
		Dim sFrameLabel As String
		sFrameLabel = Resulttree.GetFirstChildName(GetSelectedTreeItem)
		If sFrameLabel = "" Then
			ReportError("No frames found, exiting.")
		Else
			NumberOfPICFrames = 1
			While(Resulttree.GetNextItemName(sFrameLabel) <> "")
				NumberOfPICFrames = NumberOfPICFrames + 1
				sFrameLabel = Resulttree.GetNextItemName(sFrameLabel)
			Wend
		End If
	ElseIf (InStr(GetSelectedTreeItem(), "1D Results")) Then
		IdentifyPlotDomain = "1DResult"
        Is3DPlot = False

		NumberOfCurvesInPlot = Plot1D.GetNumberOfCurves
		If (NumberOfCurvesInPlot <= 0) Then
			ReportError("No result curves found, exiting. Please note that this feature currently does not support multiselections or parametric results. If multiple curves are needed, please copy them into a single folder and select the folder, then restart this macro.")
		End If
	Else
		IdentifyPlotDomain = Plot2D3D.GetDomain
        Is3DPlot = True
	End If
End Function
