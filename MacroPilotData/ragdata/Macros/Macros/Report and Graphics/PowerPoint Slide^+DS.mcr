' *Report / PowerPoint Slide
' !!! Do not change the line above !!!
' macro.561
' ================================================================================================
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 30-Jun-2021 ube: replaced old functions StoreViewInClipboard / StoreViewInBmpFile by ExportImageToClipboard / ExportImageToFile
' 06-Jun-2019    : new template CST_default.pptx
' 11-Oct-2017    : Fixed issue that paste did not work if the presentation was already open (Office 2013 or higher)
' 09-Jun-2011 rsj: Set comment to SetLock and added the command: DS.Getinstallpath
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 23-Jul-2007 ube: common macro for MWS + DS
' 03-Jul-2007 ube: created report file stored at same level as cst-file
' 21-Oct-2005 imu: Included into Online Help
' 27-Jan-2005 ube: small spelling fix Prsentation -> Presentation
' 25-Jun-2004 ube: new command GetSelectedTreeItem used
' 02-Jun-2004 ube: one common template for all Studios (CST_default.ppt)
' 30-Sep-2002 ube: check for correct active presentation (with txt and image frames in the master)
' 24-Sep-2002 ube: global lib included (better RealVal Function)
'--------------------------------------------------------------------------------------------

' **********************************************
' ube 25-Jun commented out !!! Option Explicit
' **********************************************

'#include "vba_globals_all.lib"

Const HelpFileName = "common_preloadedmacro_Report_PowerPoint_Slide"


Public projectdir As String
Public templatedir As String
Const MacroName = "Report~PowerPoint_Presentation"
'--------------------------------------------------------------------------------------------

Sub Main ()

        Dim ppt_fallback As String, ppt_default As String, ppt_template As String
        Dim projectdir As String, projectname As String, ppt_file As String
        Dim version As Integer, filename As String, presentation As Variant
        Dim title As String

        With CreateObject("PowerPoint.Application").Presentations.Application
                .Visible = True
                .WindowState = 2
                On Error Resume Next
                        Set presentation = .ActivePresentation
                On Error GoTo 0
        End With

		Dim sTreeEntry As String

		sTreeEntry = "Screenshot"

		If ( Not bDS ) Then
			sTreeEntry = GetSelectedTreeItem

			' only consider the last 2 levels in the slide-title

			Dim i1cst_count As Integer, i2cst_count As Integer, i3cst_count As Integer
			i1cst_count = InStrRev(sTreeEntry, "\")
			If i1cst_count > 0 Then
				i2cst_count = InStrRev(sTreeEntry, "\",  i1cst_count-1 )
				If i2cst_count > 0 Then
					i3cst_count = InStrRev(sTreeEntry, "\", i2cst_count-1 )
					If i3cst_count > 0 Then
						sTreeEntry = Mid(sTreeEntry, 1+i3cst_count )
					End If
				End If
			End If

			sTreeEntry = Replace(sTreeEntry,"\"," > ")
		End If

        If (Not IsEmpty(presentation)) Then

				If (bPPT_has_TextAndImageFrame(presentation)) Then

	                ExportImageToClipboard 0,0
	                ppt_AddSlide presentation, sTreeEntry
	                Set presentation = Nothing
	                Exit All

				Else

					Dim cst_answer As Integer
					cst_answer = MsgBox( _
							"The currently active ppt Presentation does not contain correct Frames." + vbCrLf + _
							"It is not possible to add CST-slides to this presentation." + vbCrLf + vbCrLf + _
							"The active ppt presentation is:   " + presentation.Name + vbCrLf + vbCrLf + _
							"Would you like to open a new Presentation to add this slide?", _
							vbYesNo+vbExclamation,"Active Presentation without frames")

			        If (cst_answer = vbNo) Then
						Exit All
					End If
	                Set presentation = Nothing

				End If

        Else
                Set presentation = Nothing
        End If

        '######################################################################
        'Select a PowerPoint template
        '        * Create template directory, if necessary
        '        * Create default template, if necessary
        '######################################################################

        templatedir = DS.GetInstallPath + "\Library\Misc\templates\"
        If (Dir$(templatedir, vbDirectory) = "") Then
                MkDir templatedir
        End If

        ppt_fallback = "CST_default.pptx"
        ppt_default = GetString("CST STUDIO SUITE", "Presentation", "pptx_template", ppt_fallback)

        If (Dir$(templatedir + ppt_fallback, vbNormal) = "") Then
                ppt_CreateTemplate templatedir + ppt_fallback
        End If

        '######################################################################
        'Create a new PowerPoint presentation from template
        '        * Suggest an automatically created new filename
        '        * Let the user select a filename
        '        * Make sure that the file doesn't already exist
        '        * Let the user select a template
        '        * Check selected template for validity
        '######################################################################

		Dim s1 As String
		s1 = GetProjectPath("Project")
		projectdir  = DirName(s1)
		projectname = ShortName(s1)

		Dim newversion As Integer
		version  = 1
        filename = FindFirstFile(projectdir, projectname + "_##.*", False)
        While (filename <> "")
                newversion  = CInt(Right$(BaseName(filename), 2)) + 1
                If newversion > version Then version = newversion
                filename = FindNextFile
        Wend
        ppt_file = projectdir + "\" + projectname + Format(version, "\_00\.\p\p\t")

        Begin Dialog UserDialog 390,231,"Powerpoint Slide",.DialogFunc
                GroupBox 10,0,370,189,"",.GroupBox1
                        Text 30,21,70,14,"Filename",.LabelFile
                                TextBox 40,35,210,21,.FileName
                                PushButton 270,35,90,21,"Browse...",.BrowseFile
                        Text 30,70,130,14,"Template",.LabelTemplate
                                TextBox 40,84,210,21,.TemplateName
                                PushButton 270,84,90,21,"Browse...",.BrowseTemplate
                        CheckBox 30,119,250,14,"Keep PowerPoint Open",.Keep
                        CheckBox 40,140,250,14,"Show Presentation",.Show
                        CheckBox 50,161,250,14,"Run Presentation",.Run
                OKButton 20,203,100,21
                CancelButton 145,203,100,21
                PushButton 270,203,100,21,"Help",.Help
        End Dialog
        Dim dlg As UserDialog

        Do
                dlg.FileName     = ShortName(ppt_file)
                dlg.TemplateName = ppt_default
                dlg.Keep         = 1
                dlg.Show         = 0
                dlg.Run          = 0
                If (Dialog(dlg) >= 0) Then Exit All
        Loop Until (dlg.FileName <> "" And dlg.TemplateName <> "")

        ppt_template = IIf(ShortName(dlg.TemplateName) = dlg.TemplateName, templatedir + dlg.TemplateName, dlg.TemplateName)
        
        SaveString  "CST STUDIO SUITE", "Presentation", "pptx_template", ShortName(ppt_template)

        ppt_file = FullPath(dlg.FileName, projectdir)
        
        Dim Ftmp As String
		Ftmp$ = Dir$(ppt_template)
		If Ftmp$ = "" Then
			MsgBox "Template "+ShortName(ppt_template)+" not found." + vbCrLf + vbCrLf + _
					"Please restart macro and choose an existing template.",vbExclamation
			Exit All
		End If

        ppt_CreateFromTemplate presentation, ppt_file, ppt_template


        '######################################################################
        'Create new slides with screenshots in PowerPoint presentation
        '        * Try to get the filtered results (AR)
        '        * else take the raw results
        '######################################################################

        ExportImageToClipboard 0,0
        ppt_AddSlide presentation, sTreeEntry


        '######################################################################
        'Save PowerPoint presentation
        '######################################################################

        With presentation
                .Save
        End With
        Set presentation = Nothing

'        ScreenUpdating True
       ' SetLock False

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Sub ppt_CreateTemplate(filepath As String)

        Dim SlideWidth As Single, SlideHeight As Single
        Dim i As Integer

        With CreateObject("PowerPoint.Application")
                With .Presentations.Add
                        SlideWidth = .PageSetup.SlideWidth
                        SlideHeight = .PageSetup.SlideHeight
                        With .SlideMaster
                                For i = .Shapes.Count To 1 Step -1
                                        .Shapes(i).Delete
                                Next
                                With .Shapes.AddLabel(1, 100, 100, 100, 100)
                                        .TextFrame.TextRange = "CST_TextFrame"
                                        .Width = SlideWidth * 0.8
                                        .Height = SlideHeight * 0.075
                                        .Left = SlideWidth * 0.1
                                        .Top = SlideHeight * 0.05
                                        .Fill.ForeColor.RGB = RGB(192, 192, 192)
                                        .Fill.Solid
                                        .Fill.Visible = True
                                End With
                                With .Shapes.AddLabel(1, 100, 100, 100, 100)
                                        .TextFrame.TextRange = "CST_ImageFrame"
                                        .TextFrame.HorizontalAnchor = 2
                                        .TextFrame.VerticalAnchor = 3
                                        .TextFrame.WordWrap = False
                                        .Width = SlideWidth * 0.8
                                        .Height = SlideHeight * 0.7
                                        .Left = SlideWidth * 0.1
                                        .Top = SlideHeight * 0.15
                                        .Fill.ForeColor.RGB = RGB(192, 192, 192)
                                        .Fill.Solid
                                        .Fill.Visible = True
                                End With
                                With .Background.Fill
                                        .Visible = True
                                        .ForeColor.RGB = RGB(0, 0, 255)
                                        .Transparency = 0
                                        .OneColorGradient 3, 4, 0.23
                                End With
                        End With
                        .SaveAs(filepath)
                        .Close
                End With
        End With

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function ppt_CheckTemplate (ppt_template As String)

        Dim SlideWidth As Single, SlideHeight As Single
        Dim Flag1, Flag2 As Boolean
        Dim index As Integer, version As Integer
        Dim pattern As String, Caption As String
        Dim Base As String, filename As String

        With CreateObject("PowerPoint.Application")
                .Visible = True
                .WindowState = 2
                With .Presentations.Open(filename:=ppt_template)
                        SlideWidth = .PageSetup.SlideWidth
                        SlideHeight = .PageSetup.SlideHeight
                        Flag1 = False
                        Flag2 = False
                        With .SlideMaster
                                For index = .Shapes.Count To 1 Step -1
                                        With .Shapes(index)
                                                If (.HasTextFrame) Then
                                                        Caption = .TextFrame.TextRange
                                                        If (Caption = "CST_ImageFrame") Then Flag1 = True
                                                        If (Caption = "CST_TextFrame") Then Flag2 = True
                                                End If
                                        End With
                                Next
                        End With
                        If Not Flag1 Then
                                With .SlideMaster.Shapes.AddLabel(1, 100, 100, 100, 100)
                                        .TextFrame.TextRange = "CST_TextFrame"
                                        .Width = SlideWidth * 0.8
                                        .Height = SlideHeight * 0.075
                                        .Left = SlideWidth * 0.1
                                        .Top = SlideHeight * 0.05
                                        .Fill.ForeColor.RGB = RGB(192, 192, 192)
                                        .Fill.Solid
                                        .Fill.Visible = True
                                End With
                        End If
                        If Not Flag2 Then
                                With .SlideMaster.Shapes.AddLabel(1, 100, 100, 100, 100)
                                        .TextFrame.TextRange = "CST_ImageFrame"
                                        .TextFrame.HorizontalAnchor = 2
                                        .TextFrame.VerticalAnchor = 3
                                        .TextFrame.WordWrap = False
                                        .Width = SlideWidth * 0.8
                                        .Height = SlideHeight * 0.7
                                        .Left = SlideWidth * 0.1
                                        .Top = SlideHeight * 0.15
                                        .Fill.ForeColor.RGB = RGB(192, 192, 192)
                                        .Fill.Solid
                                        .Fill.Visible = True
                                End With
                        End If
                        If Not (Flag1 And Flag2) Then
                                Base = BaseName(ppt_template)

                                For index = 99 To 1 STEP -1
                                        pattern = templatedir + Base + Format(index, "\_00\.\*")
                                        If (Dir$(pattern, vbNormal) <> "") Then Exit For
                                        version = index
                                Next
                                ppt_template = Base + Format(version, "\_00\.\p\p\t")

                                filename = GetFilePath(ppt_template, "ppt", templatedir, "Save Modified Template As", 3)
                                If (filename <> "") Then
                                        ppt_template = filename
                                        .SaveAs(filename:=ppt_template)
                                End If
                        End If
                        .Close
                End With
        End With

        ppt_CheckTemplate = ppt_template

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Sub ppt_CreateFromTemplate (presentation As Object, ppt_file As String, ppt_template As String)

        With CreateObject("PowerPoint.Application")
                .Visible = True
                .WindowState = 2
                With .Presentations.Open(filename:=ppt_template)
                        .SaveAs(filename:=ppt_file)
                        Set presentation = .Application.ActivePresentation
                End With
        End With

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Sub ppt_AddSlide (presentation As Object, title As String)

        Dim index_slide As Integer, i As Integer
        Dim Caption As String
        Dim FrameWidth As Single, FrameHeight As Single
        Dim FrameLeft As Single, FrameTop As Single
        Dim SizeFactor As Single
        Dim ImageWidth As Single, ImageHeight As Single
        Dim ImageLeft As Single, ImageTop As Single
        Dim TargetWidth As Single, TargetHeight As Single
        Dim TargetLeft As Single, TargetTop As Single

        With presentation
                index_slide = .Slides.Count + 1
                .Slides.Add index_slide, 12
                With .SlideMaster
                        For i = .Shapes.Count To 1 STEP -1
                                With .Shapes(i)
                                        If (.HasTextFrame) Then
                                                Caption = .TextFrame.TextRange
                                                If (Caption = "CST_ImageFrame") Then
                                                        FrameWidth  = .Width
                                                        FrameHeight = .Height
                                                        FrameLeft   = .Left
                                                        FrameTop    = .Top
                                                        .Visible = False
                                                End If
                                        End If
                                End With
                        Next
                End With
				' Switch to normal mode to paste content
        		.Application.ActiveWindow.ViewType = 1
                With .Slides(index_slide).Shapes.Paste(1)
                        SizeFactor = IIf(FrameWidth/.Width < FrameHeight/.Height, FrameWidth/.Width, FrameHeight/.Height)
                        TargetWidth  = SizeFactor * .Width
                        TargetHeight = SizeFactor * .Height
                        TargetLeft   = FrameLeft + (FrameWidth - TargetWidth) / 2
                        TargetTop    = FrameTop + (FrameHeight - TargetHeight) / 2
                        .Width  = TargetWidth
                        .Height = TargetHeight
                        .Left   = TargetLeft
                        .Top    = TargetTop
                End With
                With .SlideMaster
                        For i = .Shapes.Count To 1 STEP -1
                                With .Shapes(i)
                                        If (.HasTextFrame) Then
                                                Caption = .TextFrame.TextRange
                                                If (Caption = "CST_TextFrame") Then
                                                        .Visible = True
                                                        .Copy
                                                        .Visible = False
                                                End If
                                        End If
                                End With
                        Next
                End With
                With .Slides(index_slide).Shapes.Paste(1)
                        .Visible = True
                        .TextFrame.TextRange = title
                End With
				' Switch to slide sorting mode
				.Application.ActiveWindow.ViewType = 7
        End With

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function bPPT_has_TextAndImageFrame (presentation As Object) As Boolean

        Dim i As Integer
        Dim Caption As String

		Dim b_ImageFrame_found As Boolean
		Dim b_TextFrame_found As Boolean

		b_ImageFrame_found = False
		b_TextFrame_found = False

        With presentation
                With .SlideMaster
                        For i = .Shapes.Count To 1 STEP -1
                                With .Shapes(i)
                                        If (.HasTextFrame) Then
                                                Caption = .TextFrame.TextRange
                                                If (Caption = "CST_ImageFrame") Then
                                                        b_ImageFrame_found = True
                                                End If
                                                If (Caption = "CST_TextFrame") Then
                                                        b_TextFrame_found = True
                                                End If
                                        End If
                                End With
                        Next
                End With
        End With

        bPPT_has_TextAndImageFrame = b_ImageFrame_found And b_TextFrame_found

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

        Dim dummy As String, filename As String, extension As String
        Dim index As Integer
        Dim poerpoint As Object

        Select Case Action
                Case 1 ' Dialog box initialization
                        DlgEnable "Keep", 0
                        DlgEnable "Show", 0
                        DlgEnable "Run", 0
                Case 2 ' Value changing or button pressed
                        Select Case Item
                                Case "BrowseFile"
                                        extension = "ppt"
                                        filename = FullPath(DlgText("Filename"), projectdir)
                                        filename = GetFilePath(ShortName(filename), extension, DirName(filename), "Save Presentation As", 3)
                                        If (filename <> "") Then
                                                DlgText "Filename", ShortPath(filename, projectdir)
                                        End If
                                        DialogFunc = True
                                Case "BrowseTemplate"
                                        extension = "ppt"
                                        filename = FullPath(DlgText("TemplateName"), templatedir)
                                        filename = GetFilePath(ShortName(filename), extension, DirName(filename), "Load Template From", 0)
                                        If (filename <> "") Then
                                                filename = ppt_CheckTemplate(filename)
                                                DlgText "TemplateName", ShortPath(filename, templatedir)
                                        End If
                                        DialogFunc = True
                                Case "Keep"
                                        DlgEnable "Run", DlgValue("Keep")
                                        DialogFunc = True
                                Case "Help"
										StartHelp HelpFileName
                                        DialogFunc = True
                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select
End Function
