' *Graphics / Save Image
' !!! Do not change the line above !!!
' macro.565
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
'--------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 01-Mar-2021 thn: change logo default position to NorthWest
' 11-Feb-2020 mwl: use video_creation.lib instead of doing everything manually
' 14-May-2019    : new logo
' 24-Oct-2015 ube: subroutine Start renamed to Start_LIB
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 17-Jun-2008 ube,ala: new logo CST_logo.png
' 18-Jun-2007 ube: created graphics file stored at same level as cst-file
' 21-Oct-2005 imu: Included into Online Help
' 12-Dec-2003 ube: ability to switch off "inserting logo" via public constant
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_graphics_save_image"
Const cst_macroname = "Graphics~Save_Image"

'#include "vba_globals_all.lib"
'#include "video_creation.lib"

Public cst_doInsertLogo As Boolean
Public cst_DoAllCodecs As Boolean
Public cst_projectdir As String
Public cst_projectname As String
Public cst_tempdir As String
Public cst_templatedir As String
Public cst_ffmpeg As String         ' full path to ffmpeg.exe (including file name)
Public cst_ffplay As String         ' full path to ffplay.exe (including file name)
Public cst_codec As Integer, cst_size As Integer, cst_resize As Integer
Public cst_logo_filename As String, cst_logo_position As Integer, cst_logo_border As Integer
Public cst_outputfile As String
'-----------------------------------------------------------------------------------------------------------------------------

Sub Main ()

	cst_templatedir	 = MakeWindowsPath(GetInstallPath() + "\Library\Misc\templates")

    GetFFMPegPath(cst_ffmpeg, cst_ffplay)

	cst_tempdir = MakeWindowsPath(GetProjectPath("Temp") + "Save_Image\") ' see CST-46286
	If (Dir$(cst_tempdir, vbDirectory) = "") Then
		MkDir(cst_tempdir)
	Else
		DeleteFilesWithPattern(cst_tempdir, "*.*")
	End If

	Dim s1 As String
	s1 = MakeWindowsPath(GetProjectPath("Project")) ' see CST-46286
	cst_projectdir	= DirName(s1)
	cst_projectname = ShortName(s1)

	Dim cst_outputfile_basename As String
	cst_outputfile_basename = CreateFileBaseName(cst_projectdir, cst_projectname)

	Begin Dialog UserDialog 0,0,480,315,"Save Image",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,460,154,"Image Settings",.BoxAnimation
		CheckBox 370,63,90,14,"All Formats",.AllCodecs
		Text 20,21,80,14,"Filename",.LabelTarget
		TextBox 20,36,340,21,.Target
		PushButton 370,35,90,21,"Browse...",.BrowseTarget
		Text 20,68,90,14,"Image Format:",.LabelCodec
		DropListBox 140,63,220,119,cst_image_description(),.Codec
		Text 20,90,90,15,"Image Size:",.SizeLabel
		DropListBox 140,84,320,119,cst_video_default_format(),.Size
		OptionGroup .Resize
			OptionButton 30,112,70,14,"Crop",.OptionButton1
			OptionButton 110,112,70,14,"Fit",.OptionButton2
			OptionButton 190,112,70,14,"Distort",.OptionButton3
			OptionButton 270,112,70,14,"Ignore",.OptionButton4
			OptionButton 350,111,80,15,"Rescale",.OptionButton5
		Text 30,140,100,14,"Expert options:",.Text1
		TextBox 140,84,320,21,.ExpertOptions
		GroupBox 10,164,350,111,"",.BoxLogo
		CheckBox 20,164,60,15,"Logo",.EnableLogo
		Text 20,185,30,15,"File:",.LabelSource
		TextBox 50,182,200,21,.Source
		PushButton 260,182,90,21,"Browse...",.BrowseSource
		GroupBox 20,203,100,66,"Position",.BoxPosition
		OptionGroup .Position
			OptionButton 30,218,20,15,"",.NorthWest
			OptionButton 30,233,20,15,"",.West
			OptionButton 30,248,20,15,"",.SouthWest
			OptionButton 60,248,20,15,"",.South
			OptionButton 90,248,20,15,"",.SouthEast
			OptionButton 90,233,20,15,"",.East
			OptionButton 90,218,20,15,"",.NorthEast
			OptionButton 60,218,20,15,"",.North
		Text 130,225,130,15,"Distance to border:",.LabelBorder
		TextBox 260,222,90,21,.Border

		PushButton 380,287,80,21,"Help",.Help
		CancelButton 200,287,80,21
		PushButton 290,287,80,21,"Preview",.Preview
		OKButton 110,287,80,21
	End Dialog

	cst_codec	       = Max(Min(GetRegDouble("CST STUDIO SUITE", "Animation", "image_format", GIF_), LASTIMAGECODEC), GIF_)
	cst_outputfile     = cst_outputfile_basename + "." + cst_image_default_extension(cst_codec)
	cst_DoAllCodecs      = GetString("CST STUDIO SUITE", "Animation", "all_image_formats", "False") = "True"
    cst_size           = GetRegDouble("CST STUDIO SUITE", "Animation", "size", 3)
    cst_resize         = GetRegDouble("CST STUDIO SUITE", "Animation", "resize", 3)
	cst_doInsertLogo   = GetString("CST STUDIO SUITE", "Animation", "insertLogo", "True") = "True"
    cst_logo_filename  = GetString("CST STUDIO SUITE", "Animation", "logo_filename", "SIMULIA_CST_Studio_Suite.png")
    cst_logo_position  = GetRegDouble("CST STUDIO SUITE", "Animation", "position", 0)
    cst_logo_border    = GetRegDouble("CST STUDIO SUITE", "Animation", "border", 10)

    Dim dlg As UserDialog
    If (Dialog(dlg) >= 0) Then Exit All

	cst_outputfile       = MakeWindowsPath(FullPath(dlg.Target, cst_projectdir))
	cst_codec	         = dlg.Codec
	cst_size	         = dlg.Size
	cst_resize	         = dlg.Resize
	cst_doInsertLogo     = dlg.EnableLogo
	cst_DoAllCodecs      = dlg.AllCodecs
	cst_logo_filename    = MakeWindowsPath(FullPath(dlg.Source, cst_templatedir))
	cst_logo_position    = dlg.position
	cst_logo_border	     = CInt(dlg.border)

	SaveInteger("CST STUDIO SUITE", "Animation", "image_format", dlg.codec)
	SaveString ("CST STUDIO SUITE", "Animation", "all_image_formats", IIf(dlg.AllCodecs, "True", "False"))
	SaveInteger("CST STUDIO SUITE", "Animation", "size", dlg.size)
	SaveInteger("CST STUDIO SUITE", "Animation", "resize", dlg.resize)
	SaveString ("CST STUDIO SUITE", "Animation", "insertLogo", IIf(dlg.EnableLogo, "True", "False"))
	SaveString ("CST STUDIO SUITE", "Animation", "logo_filename", dlg.Source)
	SaveInteger("CST STUDIO SUITE", "Animation", "position", dlg.position)
	SaveString ("CST STUDIO SUITE", "Animation", "border", dlg.border)

	Dim cst_tempfile As String
	cst_tempfile = cst_tempdir + "source_image.bmp"

    Dim cst_imagewidth As Integer, cst_imageheight As Integer
    ParseSize(cst_video_default_format(cst_size), cst_imagewidth, cst_imageheight)

	ExportImage(cst_resize, cst_tempfile, cst_imagewidth, cst_imageheight, True)

    CreateImageFromImage(MakeNativePath(cst_tempdir), _
                  "source_image.bmp", _
                  MakeNativePath(cst_ffmpeg), _
                  MakeNativePath(cst_ffplay), _
                  cst_video_default_format(cst_size), _
                  cst_resize, _
                  cst_doInsertLogo, _
                  MakeNativePath(cst_logo_filename), _
                  cst_logo_position, _
                  cst_logo_border, _
                  cst_codec, _
                  cst_DoAllCodecs, _
                  MakeNativePath(cst_outputfile), _
                  dlg.ExpertOptions, _
                  False)

    DeleteFilesWithPattern(cst_tempdir, "*.*")

    If Not cst_DoAllCodecs Then
		Start_LIB(MakeNativePath(cst_outputfile))
	End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

    Dim cst_filename As String, cst_extension As String, cst_index As Integer
	Dim cst_tempfile As String

    Dim cst_imagewidth As Integer, cst_imageheight As Integer
    ParseSize(DlgText("Size"), cst_imagewidth, cst_imageheight)

    Select Case Action
        Case 1 ' Dialog box initialization
			DlgText("Target", cst_outputfile)
			cst_codec = Max(Min(cst_codec, UBound(cst_codec_description)), LBound(cst_codec_description))
			DlgValue("Codec", cst_codec)
			DialogFunc("Codec", 2, cst_codec)
			DlgValue("Size", cst_size)
			DlgValue("Resize", cst_resize)
			DlgValue("EnableLogo", cst_doInsertLogo)
			DialogFunc("EnableLogo", 2, cst_doInsertLogo)
			DlgValue("AllCodecs", cst_DoAllCodecs)
			DialogFunc("AllCodecs", 2, cst_DoInsertLogo)
			DlgText("Source", cst_logo_filename)
			DlgValue("Position", cst_logo_position)
			DlgText("Border", CStr(cst_logo_border))
			DlgText("ExpertOptions", cst_image_default_option(CInt(DlgValue("Codec"))))
			DlgEnable("Size", IIf(DlgValue("Resize") = 3, 0, 1))

        Case 2 ' Value changing or button pressed
			DialogFunc = True

            Select Case Item
                Case "BrowseTarget"
					cst_extension = cst_image_default_extension(DlgValue("Codec"))
					cst_filename  = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
					cst_filename  = MakeWindowsPath(GetFilePath(ShortName(cst_filename), cst_extension, DirName(cst_filename), "Save Image As", 3))
					If (cst_filename <> "") Then
						DlgText("Target", ShortPath(cst_filename, cst_projectdir))
					End If
				Case "Codec"
					cst_extension = cst_image_default_extension(Value)
					cst_filename  = DlgText("Target")
					cst_index	  = InStrRev(cst_filename, ".")
					DlgText("Target", Left$(cst_filename, cst_index) + cst_extension)
					DlgText("ExpertOptions", cst_image_default_option(CInt(DlgValue("Codec"))))
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
                Case "Help"
					StartHelp HelpFileName
                Case "Cancel"
					DialogFunc = False ' close dialog
                Case "Preview" ' TODO
					Dim cst_resize As Integer
					cst_tempfile = cst_tempdir + "preview.bmp"
                    cst_outputfile = MakeWindowsPath(FullPath(DlgText("Target"), cst_projectdir))
                    cst_codec = DlgValue("codec")
					cst_resize = DlgValue("resize")
                    ExportImage(cst_resize, cst_tempfile, cst_imagewidth, cst_imageheight, True)
				    CreateImageFromImage(MakeNativePath(cst_tempdir), _
                                  MakeNativePath(cst_tempfile), _
				                  MakeNativePath(cst_ffmpeg), _
				                  MakeNativePath(cst_ffplay), _
				                  DlgText("size"), _
				                  cst_resize, _
				                  cst_doInsertLogo, _
				                  MakeNativePath(FullPath(DlgText("Source"), cst_templatedir)), _
                                  DlgValue("position"), _
                                  CInt(DlgText("Border")), _
                                  cst_codec, _
				                  cst_DoAllCodecs, _
				                  MakeNativePath(cst_outputfile), _
				                  DlgText("ExpertOptions"), _
				                  True)
                Case "OK"
					DlgEnd(-1)
            End Select
        Case 3 ' ComboBox or TextBox Value changed
        Case 4 ' Focus changed
        Case 5 ' Idle
    End Select
End Function

