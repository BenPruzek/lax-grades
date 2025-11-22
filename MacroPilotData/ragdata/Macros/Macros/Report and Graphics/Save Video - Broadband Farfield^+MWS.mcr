' *Graphics / Save Broadband Farfield Video
' !!! Do not change the line above !!!
' macro.566
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
'--------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 09-Aug-2022 set: Use FarfieldPlot.Redraw instead of reselecting in the tree
' 04-Jan-2022 set: Fix plot update
' 01-Mar-2021 thn: change logo default position to NorthWest
' 11-Feb-2020 mwl: use video_creation.lib instead of doing everything manually
' 14-May-2019    : new logo
' 24-Oct-2015 ube: subroutine Start renamed to Start_LIB
' 28-Aug-2015 tgl: Added chcp to batch file to be more robust with umlauts
' 18-Dec-2014 ube,fsr: Using project temp folder now, also working for unmapped network drives (\\server\...)
' 30-Dec-2013 ube: point button page to save video page
' 14-Mar-2012 ube: adding option  +dither  for improved look of legend and colorbar
' 02-Aug-2011 fsr: replaced obsolete 'vba_globals.lib' with 'vba_globals_all.lib' and 'vba_globals_3d.lib'
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 19-Jun-2008 fde: Removes MSGBOX for time step adjustment dor 2009
' 17-Jun-2008 ube,ala: new logo CST_logo.png
' 06-Nov-2007 fde: added warning if selected tree Item is not a braodband farfield monitor
' 05-Nov-2007 fde: wrong MSgbox Fixed, added MSGBOX to time step adjustment. Solved roundig error problem
' 22-Sep-2007 fde: some bugs fixed
' 04-Jul-2007 fde: adopted for Broadband Farfield (still a problem with the flow control. Sometimes the last value is not plotted
' 18-Jun-2007 ube: created graphics file stored at same level as cst-file
' 21-Oct-2005 imu: Included into Online Help
' 19-Nov-2004 ube,btr: temporary replacement of enum-construction by constants
' 12-Dec-2003 ube: ability to switch off "inserting logo" via public constant
' 21-Oct-2003 ube: bmp2avi has many problems with crashes, therefore disable those 3 formats with Public Constant
' 29-Jul-2003 ube: bugfix for MPEG-time movie
' 29-Jul-2003 ube: bmp2avi.exe link now to CST's support page  MWS_FAQs
' 05-Jun-2003 ube: new www-link for bmp2avi.exe
' 15-Nov-2002 ube: time movies included
'--------------------------------------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_Graphics_Save_Video"
Const cst_macroname = "Graphics~Save_Video"

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
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

Public cst_pindex As Integer
Public sPlotDomain As String
Public cst_dt As Double, cst_dtlow As Double, cst_dthigh As Double, cst_dtstep As Double
Public cst_df As Double, cst_dflow As Double, cst_dfhigh As Double, cst_dfstep As Double

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

	Dim s1 As String
	s1 = MakeWindowsPath(GetProjectPath("Project")) ' see CST-46286
	cst_projectdir	= DirName(s1)
	cst_projectname = ShortName(s1)

	Dim cst_videofile_basename As String
	cst_videofile_basename = CreateFileBaseName(cst_projectdir, cst_projectname)


	Begin Dialog UserDialog 0,0,480,476,"Save Broadband Farfield Video",.DialogFunc ' %GRID:10,7,1,1
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
		Text 30,162,100,15,"Framerate [1/s]:",.Text3
		TextBox 140,159,90,21,.FrameRate
		CheckBox 240,161,210,14,"Bounce Playback At End",.BounceVideoPlaybackAtEnd
		Text 30,186,100,15,"Expert options:",.Text1
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
		GroupBox 10,336,220,105,"",.GroupBox2
		GroupBox 250,336,220,105,"",.GroupBox4
		Text 20,357,60,14,"Start",.t0
		Text 260,357,60,14,"Start",.t2
		Text 20,385,50,14,"Stop",.t1
		Text 260,385,50,14,"Stop",.t3
		Text 20,413,50,14,"Step",.dt
		Text 260,413,50,14,"Step",.dt2
		TextBox 100,357,100,21,.dtlow
		TextBox 330,357,110,21,.dfreqlow
		TextBox 100,385,100,21,.dthigh
		TextBox 330,385,110,21,.dfreqhigh
		TextBox 100,413,100,21,.dtstep
		TextBox 330,413,110,21,.dfreqstep
		OptionGroup .TFSelect_group
			OptionButton 20,336,120,14,"Time Animation",.TimeSelect
			OptionButton 260,336,160,14,"Frequency Animation",.FrequencySelect
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

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

    Dim cst_filename As String, cst_extension As String, cst_index As Integer
	Dim cst_tempfile As String

    Dim cst_videowidth As Integer, cst_videoheight As Integer
    ParseSize(DlgText("Size"), cst_videowidth, cst_videoheight)

    Select Case Action
        Case 1 ' Dialog box initialization
            If Left$(GetSelectedTreeItem, 30) <> "Farfields\farfield (broadband)" Then
            	MsgBox("Please select broadband farfield monitor in tree first")
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
            
            DlgEnable("dfreqhigh", True)
            DlgEnable("dfreqlow", True)
            DlgEnable("dfreqstep", True)
            DlgText("dfreqhigh", CStr(Solver.getfmax))
            DlgText("dfreqlow", CStr(Solver.getfmin))
            DlgText("dfreqstep", CStr((Solver.getfmax-Solver.getfmin)*0.1))
            DlgEnable("dthigh", False)
            DlgEnable("dtlow", False)
            DlgEnable("dtstep", False)
            DlgText("dthigh", CStr(1))
            DlgText("dtlow", CStr(0))
            DlgText("dtstep", CStr(0.1))

            DlgValue("TFSelect_group", 1)
            sPlotDomain = "Frequency"

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
                Case "TFSelect_group"
                    If(DlgValue("TFSelect_group") = 1) Then
                        DlgEnable("dfreqhigh", True)
                        DlgEnable("dfreqlow", True)
                        DlgEnable("dfreqstep", True)
                        DlgEnable("dtstep", False)
                        DlgEnable("dthigh", False)
                        DlgEnable("dtlow", False)
                        sPlotDomain = "Frequency"
                    End If
                    If(DlgValue("TFSelect_group") = 0) Then
                        DlgEnable("dfreqhigh", False)
                        DlgEnable("dfreqlow", False)
                        DlgEnable("dfreqstep", False)
                        DlgEnable("dtstep", True)
                        DlgEnable("dthigh", True)
                        DlgEnable("dtlow", True)
                        sPlotDomain = "Time"
                    End If

                Case "Help"
					StartHelp HelpFileName
                Case "Cancel"
                    Plot2D3D.PhaseValue 0
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
		            If sPlotDomain="Time" Then
	    	        	cst_dtlow = CDbl(DlgText("dtlow"))
	        	    	cst_dthigh = CDbl(DlgText("dthigh"))
	            		cst_dtstep = CDbl(DlgText("dtstep"))
	            		cst_dt    = cst_dtlow
	            	Else
		   	         	cst_dflow = CDbl(DlgText("dfreqlow"))
	    	        	cst_dfhigh = CDbl(DlgText("dfreqhigh"))
	        	    	cst_dfstep = CDbl(DlgText("dfreqstep"))
	            		cst_df = cst_dflow
					End If

					cst_pindex = 0

					' ---Disable all dialog items (except for the Cancel button)---
					For cst_index = 0 To DlgCount()-1
						On Error Resume Next
							DlgEnable cst_index, IIf(DlgText(cst_index) = "Cancel", 1, 0)
						On Error GoTo 0
					Next
            End Select

        Case 3 ' ComboBox or TextBox Value changed
        Case 4 ' Focus changed
        Case 5 ' Idle
            If sPlotDomain="Time" Then
				If (cst_dthigh > cst_dtlow) Then
					If (cst_dt*0.9999 <= cst_dthigh) Then 'trip to reduce rounding error
						cst_tempfile = cst_tempdir + "image_" + Format(cst_pindex, "00000") + ".bmp"
						DlgText("dtlow", CStr(cst_dt))
                        FarfieldPlot.SetTimeDomainFF True
                        FarfieldPlot.SetTime CStr(cst_dt)
						FarfieldPlot.Redraw()
						ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, True)
						cst_pindex = cst_pindex + 1
						cst_dt = cst_dt + cst_dtstep
					ElseIf (cst_dt > cst_dthigh) Then
						DlgText("dtlow", CStr(cst_dtlow))
						DlgEnd(-1)
					End If
				End If
			ElseIf sPlotDomain="Frequency" Then
				If (cst_dfhigh > cst_dflow) Then
					If ((cst_df*0.9999999) <= cst_dfhigh) Then 'trick to remove rounding error
						cst_tempfile = cst_tempdir + "image_" + Format(cst_pindex, "00000") + ".bmp"
						DlgText("dfreqlow", CStr(cst_df))
   	                    FarfieldPlot.SetTimeDomainFF False
       	                FarfieldPlot.Setfrequency CStr(cst_df)
						FarfieldPlot.Redraw()
						ExportImage(DlgValue("Resize"), cst_tempfile, cst_videowidth, cst_videoheight, True)
						cst_pindex = cst_pindex + 1
						cst_df = cst_df + cst_dfstep
					ElseIf (cst_df > cst_dfhigh) Then
						DlgText("dfreqlow", CStr(cst_dflow))
						DlgEnd(-1)
					End If
				End If
			End If

			DialogFunc = True
          End Select
End Function
