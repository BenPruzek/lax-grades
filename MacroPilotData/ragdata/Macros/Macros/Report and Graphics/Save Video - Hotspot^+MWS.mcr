'#Language "WWB-COM"

' *Graphics / HotSpot Video
' !!! Do not change the line above !!!

' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
'-------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 01-Mar-2021 thn: change logo default position to NorthWest
' 12-Feb-2020 mwl: use video_creation.lib instead of doing everything manually
' 14-May-2019     : new logo
' 09-Jan-2018 rsh: added some error handlers for invalid input
' 28-Nov-2017 rsh: improved user interface
' 11-Jul-2017 fde : negative theta is now allowed
' 30 Nov-2016 FDE : change for new start command start_lib 
' 30-Jun-2014 fde : using project temp folder now
' 10-Jul-2013 fde : first Version
'-------------------------------------------------------------------------------------------------

Option Explicit

Const HelpFileName = "common_preloadedmacro_Graphics_Save_Video"
Const cst_macroname = "Graphics~Save_Video"

'#include "vba_globals_all.lib"
'#include "video_creation.lib"

'--------------------------------------------------------------------------------------------------------------
Public cst_doInsertLogo As Boolean
Public cst_projectdir As String
Public cst_projectname As String
Public cst_tempdir As String
Public cst_templatedir As String
Public cst_ffmpeg As String         ' full path to ffmpeg.exe (including file name)
Public cst_ffplay As String         ' full path to ffplay.exe (including file name)
Public cst_codec As Integer, cst_size As Integer, cst_resize As Integer, cst_framerate As String
Public cst_logo_filename As String, cst_logo_position As Integer, cst_logo_border As Integer
Public cst_videofile As String

Dim cst_vary As Integer
Dim cst_theta As Double
Dim cst_phi As Double
Dim cst_fix As Double
Dim cst_start As Double
Dim cst_stop As Double
Dim cst_step As Double
Dim cst_current As Double
Dim cst_pindex As Integer

Const VARY_NOTHING=-1, VARY_THETA=0, VARY_PHI=1
'--------------------------------------------------------------------------------------------------------------

Sub Main

	cst_templatedir	 = MakeWindowsPath(GetInstallPath() + "\Library\Misc\templates")

    GetFFMPegPath(cst_ffmpeg, cst_ffplay)

	cst_tempdir = MakeWindowsPath(GetProjectPath("Temp") + "Save_Video\") ' see CST-46286
	If (Dir$(cst_tempdir, vbDirectory) = "") Then
		MkDir(cst_tempdir)
	Else
		DeleteFilesWithPattern(cst_tempdir, "*.*")
	End If

	Dim s1 As String
	s1 = MakeWindowsPath(GetProjectPath("Project")) ' see CST-46286
	cst_projectdir	= DirName(s1)
	cst_projectname = ShortName(s1)

	Dim cst_videofile_basename As String
	cst_videofile_basename = CreateFileBaseName(cst_projectdir, cst_projectname)
    

	Begin Dialog UserDialog 0,0,480,476,"Save Video - Hotspot",.DialogFunc ' %GRID:10,7,1,1

		GroupBox 10,7,460,203,"Video Settings",.BoxAnimation
		Text 20,21,80,14,"Filename",.LabelTarget
		TextBox 20,36,340,21,.Target
		PushButton 370,35,90,21,"Browse...",.BrowseTarget
		Text 20,68,90,14,"Video Codec:",.LabelCodec
		DropListBox 140,63,320,119,cst_codec_description(),.Codec
		Text 20,90,90,15,"Video Size:",.SizeLabel
		DropListBox 140,84,320,119,cst_video_default_format(),.Size
		OptionGroup .Resize
			OptionButton 30,112,70,14,"Crop",.OptionButton1
			OptionButton 110,112,70,14,"Fit",.OptionButton2
			OptionButton 190,112,70,14,"Distort",.OptionButton3
			OptionButton 270,112,70,14,"Ignore",.OptionButton4
			OptionButton 350,111,80,15,"Rescale",.OptionButton5
		Text 30,162,100,15,"Framerate [1/s]:",.Text5
		TextBox 140,159,90,21,.FrameRate
		CheckBox 240,161,210,14,"Bounce Playback At End",.BounceVideoPlaybackAtEnd
		Text 30,186,100,15,"Expert options:",.Text4
		TextBox 140,182,320,21,.ExpertOptions
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

		PushButton 380,448,80,21,"Help",.Help
		CancelButton 200,448,80,21
		PushButton 290,448,80,21,"Preview",.Preview
		OKButton 110,448,80,21

		GroupBox 10,329,220,105,"Fixed Angle",.GroupBox1
		OptionGroup .FixedAngle
			OptionButton 30,350,60,14,"Phi",.varytheta
			OptionButton 30,378,70,14,"Theta",.varyphi
		GroupBox 250,329,220,105,"Vary",.GroupBox2
		TextBox 110,350,100,21,.fixphi
		TextBox 110,378,100,21,.fixtheta
		TextBox 350,350,100,21,.startangle
		TextBox 350,378,100,21,.stopangle
		TextBox 350,406,100,21,.stepangle
		Text 270,350,50,14,"Start:",.Text1
		Text 270,378,50,14,"Stop:",.Text2
		Text 270,406,60,14,"Step:",.Text3
	End Dialog

	cst_codec	       = Max(Min(GetRegDouble("CST STUDIO SUITE", "Animation", "codec", GIF), LASTCODEC), GIF)
	cst_videofile      = cst_videofile_basename + "." + cst_codec_default_extension(cst_codec)
    cst_size           = GetRegDouble("CST STUDIO SUITE", "Animation", "size", 3)
    cst_resize         = GetRegDouble("CST STUDIO SUITE", "Animation", "resize", 3)
	cst_framerate      = GetString("CST STUDIO SUITE", "Animation", "framerate", "24")
	cst_doInsertLogo   = GetString("CST STUDIO SUITE", "Animation", "insertLogo", "True") = "True"
    cst_logo_filename  = GetString("CST STUDIO SUITE", "Animation", "logo_filename", "SIMULIA_CST_Studio_Suite.png")
    cst_logo_position  = GetRegDouble("CST STUDIO SUITE", "Animation", "position", 0)
    cst_logo_border    = GetRegDouble("CST STUDIO SUITE", "Animation", "border", 10)

	cst_vary = VARY_NOTHING ' this will be set to some other value if OK is pressed in the dialog to enable processing in the dialogs Idle handler

	Dim dlg As UserDialog
    If (Dialog(dlg) >= 0) Then Exit All

	cst_videofile        = MakeWindowsPath(FullPath(dlg.Target, cst_projectdir))
	cst_codec	         = dlg.Codec
	cst_size	         = dlg.Size
	cst_resize	         = dlg.Resize
	cst_framerate        = dlg.FrameRate
	cst_doInsertLogo     = dlg.EnableLogo
	cst_logo_filename    = MakeWindowsPath(FullPath(dlg.Source, cst_templatedir))
	cst_logo_position    = dlg.position
	cst_logo_border	     = CInt(dlg.border)

	SaveInteger("CST STUDIO SUITE", "Animation", "codec", dlg.codec)
	SaveInteger("CST STUDIO SUITE", "Animation", "size", dlg.size)
	SaveInteger("CST STUDIO SUITE", "Animation", "resize", dlg.resize)
	SaveInteger("CST STUDIO SUITE", "Animation", "framerate", dlg.FrameRate)
	SaveString ("CST STUDIO SUITE", "Animation", "insertLogo", IIf(dlg.EnableLogo, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "logo_filename", dlg.Source)
	SaveInteger("CST STUDIO SUITE", "Animation", "position", dlg.position)
	SaveString ("CST STUDIO SUITE", "Animation", "border", dlg.border)

    CreateVideoFromImageSequence(MakeNativePath(cst_tempdir), _
                  IIf(IsWindows(), "image_%%05d.bmp", "image_%05d.bmp"), _
                  MakeNativePath(cst_ffmpeg), _
                  MakeNativePath(cst_ffplay), _
                  cst_video_default_format(cst_size), _
                  cst_resize, _
                  cst_framerate, _
                  cst_doInsertLogo, _
                  MakeNativePath(cst_logo_filename), _
                  cst_logo_position, _
                  cst_logo_border, _
                  cst_codec, _
                  False, _
                  MakeNativePath(cst_videofile), _
                  False, _
    			  dlg.BounceVideoPlaybackAtEnd, _
                  dlg.ExpertOptions, _
                  False)

    DeleteFilesWithPattern(cst_tempdir, "*.*")

	Start_LIB(MakeNativePath(cst_videofile))
End Sub


Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

    Dim cst_filename As String, cst_extension As String, cst_index As Integer, cst_mode As Integer
	Dim cst_tempfile As String

    Dim cst_videowidth As Integer, cst_videoheight As Integer
    ParseSize(DlgText("Size"), cst_videowidth, cst_videoheight)

    Select Case Action
	    Case 1 ' Dialog box initialization
		    Dim cst_mon_name As String
		    cst_mon_name = Right$(GetSelectedTreeItem,(Len(GetSelectedTreeItem)-(InStrRev(GetSelectedTreeItem,"\"))))
		    If ((Left$(GetSelectedTreeItem,35) <> "2D/3D Results\Hotspots\Polarization") Or (Left$(cst_mon_name,7) <> "thetaEx")) Then
		        MsgBox("Please select a hotspot monitor first")
		        Exit All
		    End If

			DlgText("Target", cst_videofile)
			cst_codec = Max(Min(cst_codec, UBound(cst_codec_description)), LBound(cst_codec_description))
			DlgValue("Codec", cst_codec)
			DialogFunc("Codec", 2, cst_codec)
			DlgValue("Size", cst_size)
			DlgValue("Resize", cst_resize)
			DlgText("FrameRate", cst_framerate)
			DlgValue("EnableLogo", cst_doInsertLogo)
			DialogFunc("EnableLogo", 2, cst_doInsertLogo)
			DlgText("Source", cst_logo_filename)
			DlgValue("Position", cst_logo_position)
			DlgText("Border", CStr(cst_logo_border))
			DlgText("ExpertOptions", cst_codec_default_option(CInt(DlgValue("Codec"))))
			DlgEnable("Size", IIf(DlgValue("Resize") = 3, 0, 1))

			DlgText("fixtheta", "90")
			DlgText("fixphi", "0")
	        DlgText("startangle", "0")
	        DlgText("stopangle", "180")
	        DlgText("stepangle", "10")
			DialogFunc("FixedAngle", 2, 0)

	    Case 2 ' Value changing or button pressed
			DialogFunc = True

			Select Case Item
                Case "BrowseTarget"
					cst_extension = cst_codec_default_extension(DlgValue("Codec"))
					cst_filename  = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
					cst_filename  = MakeWindowsPath(GetFilePath(ShortName(cst_filename), cst_extension, DirName(cst_filename), "Save Animation As", 3))
					If (cst_filename <> "") Then
						DlgText("Target", ShortPath(cst_filename, cst_projectdir))
					End If
				Case "Codec"
					cst_extension = cst_codec_default_extension(Value)
					cst_filename  = DlgText("Target")
					cst_index	  = InStrRev(cst_filename, ".")
					DlgText("Target", Left$(cst_filename, cst_index) + cst_extension)
					DlgText("ExpertOptions", cst_codec_default_option(CInt(DlgValue("Codec"))))
                Case "Resize"
					DlgEnable("Size", IIf(DlgValue("Resize") = 3, 0, 1))
				Case "EnableLogo"
					cst_doInsertLogo = DlgValue("EnableLogo")
					DlgEnable("BoxLogo", cst_doInsertLogo)
					DlgEnable("LabelSource", cst_doInsertLogo)
					DlgEnable("Source", cst_doInsertLogo)
					DlgEnable("BrowseSource", cst_doInsertLogo)
					DlgEnable("BoxPosition", cst_doInsertLogo)
					DlgEnable("Position", cst_doInsertLogo)
					DlgEnable("LabelBorder", cst_doInsertLogo)
					DlgEnable("Border", cst_doInsertLogo)
                Case "BrowseSource"
					cst_extension = "Image Files|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tga;*.tiff"
					cst_filename  = MakeWindowsPath(FullPath(DlgText("Source"), cst_templatedir))
					cst_filename  = MakeWindowsPath(GetFilePath(ShortName(cst_filename), cst_extension, DirName(cst_filename), "Choose logo file", 0))
					If (cst_filename <> "") Then
						DlgText("Source", ShortPath(cst_filename, cst_templatedir))
					End If
				Case "FixedAngle"
					cst_mode = DlgValue("FixedAngle")
					DlgEnable("fixphi", cst_mode <> VARY_PHI)
					DlgEnable("fixtheta", cst_mode <> VARY_THETA)
					DlgText("GroupBox2", IIf(cst_mode = VARY_THETA, "Vary Theta", "Vary Phi"))
                Case "Help"
					StartHelp HelpFileName
                Case "Cancel"
					DialogFunc = False ' close dialog
                Case "Preview"
					Dim cst_resize As Integer
					cst_tempfile = cst_tempdir + "preview.bmp"
                    cst_videofile = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
                    cst_codec = DlgValue("codec")
					cst_resize = DlgValue("resize")
                    ExportImage(cst_resize, cst_tempfile, cst_videowidth, cst_videoheight, True)
                    CreateVideoFromImageSequence(MakeNativePath(cst_tempdir), _
                                  MakeNativePath(cst_tempfile), _
                                  MakeNativePath(cst_ffmpeg), _
                                  MakeNativePath(cst_ffplay), _
                                  DlgText("size"), _
                                  cst_resize, _
                                  DlgText("FrameRate"), _
                                  cst_doInsertLogo, _
                                  MakeNativePath(FullPath(DlgText("Source"), cst_templatedir)), _
                                  DlgValue("position"), _
                                  CInt(DlgText("Border")), _
                                  cst_codec, _
                                  False, _
                                  MakeNativePath(cst_videofile), _
                                  False, _
                    			  False, _
                                  DlgText("ExpertOptions"), _
                                  True)
				Case "OK"
					Dim cst_full As String, full_tree_name As String
					cst_full = Left$(GetSelectedTreeItem, InStrRev(GetSelectedTreeItem,"\"))

					full_tree_name = GenerateResultName(cst_full, DlgValue("FixedAngle"), _
														CDbl(DlgText(IIf(DlgValue("FixedAngle") = VARY_PHI, "fixtheta", "fixphi"))), _
														CStr(DlgText("startangle")))

					If Not Resulttree.DoesTreeItemExist(full_tree_name) Then
						MsgBox("First plot does not exist:" + Chr(13) + full_tree_name + "," + Chr(13) + "check fixed angle settings.")
					ElseIf (DlgText("startangle")>DlgText("stopangle")) Then
						MsgBox("Start angle has to be smaller than stop angle.")
					ElseIf (DlgText("stepangle")<=0) Then
						MsgBox("Step angle has to be a positive value")
					ElseIf ((DlgText("stopangle")-(DlgText("startangle"))) Mod DlgText("stepangle")<>0) Then
						MsgBox("Stop angle not accessible with given step angle")
					Else
						' ---Disable all dialog items (except for the Cancel button)---
						For cst_index = 0 To DlgCount()-1
							On Error Resume Next
								DlgEnable cst_index, IIf(DlgText(cst_index) = "Cancel", 1, 0)
							On Error GoTo 0
						Next

						' all is fine - do your work
						cst_vary = DlgValue("FixedAngle")
				        cst_fix = CDbl(DlgText(IIf(cst_vary = VARY_PHI, "fixtheta", "fixphi")))
						cst_step = CDbl(DlgText("stepangle"))
						cst_start = CDbl(DlgText("startangle"))
						cst_stop = CDbl(DlgText("stopangle"))
					    cst_pindex = 0
					    cst_current = cst_start
					End If
			End Select

	   	Case 3 ' TextBox or ComboBox text changed
	    Case 4 ' Focus changed
	    Case 5 ' Idle
			If cst_vary <> VARY_NOTHING Then
				If cst_current <= cst_stop Then
					DlgText("startangle", CStr(cst_current))

					cst_full = Left$(GetSelectedTreeItem, InStrRev(GetSelectedTreeItem,"\"))
					full_tree_name = GenerateResultName(cst_full, cst_vary, cst_fix, cst_current)

					If SelectTreeItem(full_tree_name) Then
						cst_tempfile = cst_tempdir + "image_" + Format(cst_pindex, "00000") + ".bmp"
						Plot.Update
						ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, True)
						cst_pindex = cst_pindex + 1
					Else
						ReportWarning "Tree item does not exist: " + full_tree_name
					End If

					cst_current = cst_current + cst_step
				Else
					DlgText("startangle", CStr(cst_start))
					DlgEnd(-1)
				End If
			End If
			DialogFunc = True
	    Case 6 ' Function key
    End Select
End Function

Function GenerateResultName(path As String, varytype As Integer, fixedValue As Double, varyingValue As Double) As String
	If varytype = VARY_PHI Then
		GenerateResultName = path + "thetaEx=" + CStr(fixedValue) +",phiEx=" + CStr(varyingValue)
	ElseIf varytype = VARY_THETA Then
		GenerateResultName = path + "thetaEx=" + CStr(varyingValue) +",phiEx=" + CStr(fixedValue)
	Else
		GenerateResultName = ""
	End If
End Function
