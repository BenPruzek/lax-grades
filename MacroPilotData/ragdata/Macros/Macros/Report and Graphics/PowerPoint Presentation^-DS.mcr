' *Report / PowerPoint Presentation
' !!! Do not change the line above !!!
' macro.560
'--------------------------------------------------------------------------------------------
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 06-Sep-2023 ube: typo corrected
' 30-Jun-2021 ube: replaced old functions StoreViewInClipboard / StoreViewInBmpFile by ExportImageToClipboard / ExportImageToFile
' 07-Jul-2020 ube: use SelectModelView instead of SelectTreeItem "Components" (to ensure, 2d3dplot windows become inactive)
' 06-Jun-2019    : new template CST_default.pptx
' 13-Dec-2017 rsh,ube: added support for 1D results in EMS
' 05-Dec-2017 rsh,ube: fixed support for B-Field and Potential and Magnetic Energy Density for EMS
' 01-Apr-2011 apr,ube: now also working for IFX viewer
' 24-Feb-2011 ube: added complex S-parameters (db + smithchart)
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 21-Sep-2008 mru: change background to white before making slide if gradient background enabled
' 03-Jul-2007 wko,ube: include screenshots of the Floquet modes (additional dialogue included)
' 18-Jun-2007 ube: created report file stored at same level as cst-file
' 21-Oct-2005 ube: Included into Online Help
' 27-Jan-2005 ube: small spelling fix Prsentation -> Presentation
' 02-Jun-2004 ube: one common template for all Studios (CST_default.ppt)
' 05-Dec-2003 ube: new Tree entries of EMS-Solver included
' 19-Nov-2003 ube: Layers replaced by Components
' 14-May-2003 ube: 1d- + 0d-templates
' 17-Nov-2002 ube: small changes (create slides for MWS-1DResults only if not empty)
' 13-Nov-2002 ube: same ppt-macro now for EMS and MWS
' 24-Sep-2002 ube: global lib included (better RealVal Function)
' 27-Jan-2002 ube: version 4 - small changes in the command SelectTreeItem
'--------------------------------------------------------------------------------------------
Option Explicit

'#include "vba_globals_all.lib"

Public projectdir As String
Public templatedir As String
Const MacroName = "Report~PowerPoint_Presentation"
'--------------------------------------------------------------------------------------------

Sub Main ()

        Dim ppt_fallback As String, ppt_default As String, ppt_template As String
        Dim projectname As String, ppt_file As String
        Dim version As Integer, filename As String, presentation As Object
        Dim title As String
		Dim iCount As Integer

        '######################################################################
        'Select a PowerPoint template
        '        * Create template directory, if necessary
        '        * Create default template, if necessary
        '######################################################################

        templatedir = GetInstallPath + "\Library\Misc\templates\"
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

        Begin Dialog UserDialog 390,231,"Powerpoint Presentation",.DialogFunc
                OKButton 20,203,100,21
                CancelButton 145,203,100,21
                PushButton 270,203,100,21,"Help",.Help
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
        End Dialog
        Dim dlg As UserDialog

        Do
                dlg.FileName     = ShortName(ppt_file)
                dlg.TemplateName = ppt_default
                dlg.Keep         = GetRegDouble("CST STUDIO SUITE", "Presentation", "keep", 1)
                dlg.Show         = GetRegDouble("CST STUDIO SUITE", "Presentation", "show", 1)
                dlg.Run          = GetRegDouble("CST STUDIO SUITE", "Presentation", "run", 1)
                If (Dialog(dlg) >= 0) Then Exit All
        Loop Until (dlg.FileName <> "" And dlg.TemplateName <> "")

        ppt_template = IIf(ShortName(dlg.TemplateName) = dlg.TemplateName, templatedir + dlg.TemplateName, dlg.TemplateName)

        SaveString  "CST STUDIO SUITE", "Presentation", "pptx_template", ShortName(ppt_template)

        SaveInteger "CST STUDIO SUITE", "Presentation", "keep", dlg.Keep
        SaveInteger "CST STUDIO SUITE", "Presentation", "show", dlg.Show
        SaveInteger "CST STUDIO SUITE", "Presentation", "run",  dlg.Run

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

        Dim n_port As Integer, n_mode As Integer, mode_type As String
        Dim wave_impedance As Double, line_impedance As Double
        Dim mode_text As String, impedance_text As String
        Dim ss_short As String, ss As String, ss1 As String, ilen2 As Integer

        ' ube 13-nov-2002    bug:switching wcs moves structure
        ' WCS.ActivateWCS "global"

		Mesh.ViewMeshMode  False

        title = CaptureTreeItem("Components", "Benchmark")
        ppt_AddSlide presentation, title

		If bMWS Then
	        ' ube 8-may-2001: only one port/one mode (faster)  For n_port = 1 To 10
	        For n_port = 1 To 1
	                On Error Resume Next
	                        mode_type = Port.GetModeType(n_port, 1)
	                On Error GoTo 0
	                If (mode_type = "") Then Exit For
	        ' ube 8-may-2001: only one port/one mode (faster)  For n_mode = 1 To 10
	                For n_mode = 1 To 1
	                        mode_type = ""
	                        On Error Resume Next
	                                mode_type = Port.GetModeType(n_port, n_mode)
	                        On Error GoTo 0
	                        If (mode_type = "") Then Exit For
	                        If (mode_type = "UNDEF") Then mode_type = "HYBRID"
	                        wave_impedance = Port.GetWaveImpedance(n_port, n_mode)
	                        line_impedance = Port.GetLineImpedance(n_port, n_mode)
	                        mode_text      = mode_type + "-mode"
	                        impedance_text = IIf(line_impedance > 0, _
	                                Format(line_impedance, """Line Imp. ""0"" Ohms"""), _
	                                Format(wave_impedance, """Wave Imp. ""0"" Ohms"""))
	                        title = CaptureTreeItem( _
	                                "2D/3D Results\Port Modes\Port" + CStr(n_port) + "\e" + CStr(n_mode), _
	                                "E(" + CStr(n_port) + ", " + CStr(n_mode) + "): " + mode_text + ", " + impedance_text)
	                        ppt_AddSlide presentation, title
	                        title = CaptureTreeItem( _
	                                "2D/3D Results\Port Modes\Port" + CStr(n_port) + "\h" + CStr(n_mode), _
	                                "H(" + CStr(n_port) + ", " + CStr(n_mode) + "): " + mode_text + ", " + impedance_text)
	                        ppt_AddSlide presentation, title
	                Next
	        Next

			Dim sFloquetPorts(2)
			Dim nFloquetPorts As Integer
			nFloquetPorts = 0

	        If (FloquetPort.IsPortAtZmin) Then
				sFloquetPorts(nFloquetPorts) = "Zmin"
				nFloquetPorts = nFloquetPorts + 1
	        End If

	        If (FloquetPort.IsPortAtZmax) Then
				sFloquetPorts(nFloquetPorts) = "Zmax"
				nFloquetPorts = nFloquetPorts + 1
	        End If

	        Dim bIncludeFloquet As Boolean, nMaxMode As Integer
	        bIncludeFloquet = False
	        nMaxMode = 9999

	        If nFloquetPorts > 0 Then
				Begin Dialog UserDialog 370,105,"Handling of Floquet Modes" ' %GRID:10,7,1,1
					GroupBox 10,0,350,70,"",.GroupBox1
					CheckBox 20,21,320,14,"Include Screenshots of Floquet-Modes (E,H,P)",.floquet
					Text 20,49,250,14,"Highest Floquet Mode, to be included:",.Text1
					TextBox 280,42,60,21,.sMaxMode
					PushButton 10,77,90,21,"Continue",.Continue
					PushButton 110,77,90,21,"Abort",.Abort
				End Dialog
				Dim dlg2 As UserDialog
				dlg2.smaxmode = "2"
                If (Dialog(dlg2) = 2) Then Exit All
				bIncludeFloquet = dlg2.floquet
				nMaxMode = Evaluate(dlg2.sMaxMode)
	        End If

			If (bIncludeFloquet) Then
		        For n_port = 0 To nFloquetPorts - 1

	        		Dim sPort As String
	        		sPort = sFloquetPorts(n_port)

					FloquetPort.Port(sPort)

	        		Dim nFloquetModes As Integer
	        		nFloquetModes = FloquetPort.GetNumberOfModesConsidered
	        		If (nFloquetModes > nMaxMode) Then nFloquetModes = nMaxMode

					n_mode = 1

	        		If (FloquetPort.FirstMode = True) Then
	        		Do
        				Dim nOrder1 As Long
        				Dim nOrder2 As Long

                        If (FloquetPort.GetMode(mode_type, nOrder1, nOrder2) = True) Then

	                        mode_text = mode_type + "(" + CStr(nOrder1) + "," + CStr(nOrder2) + ")-mode"

                        	Dim sSub(2)
							sSub(0) = "In"
							sSub(1) = "Out"

							Dim nSub As Integer
							For nSub = 0 To 1
		                        title = CaptureTreeItem( _
		                                "2D/3D Results\Port Modes\" + sPort + "\" + sSub(nSub) + "\e" + CStr(n_mode), _
		                                "E(" + sPort + "(" + CStr(n_mode) + ")): " + mode_text + "(" + sSub(nSub) + ")")
		                        ppt_AddSlide presentation, title
		                        title = CaptureTreeItem( _
		                                "2D/3D Results\Port Modes\" + sPort + "\" + sSub(nSub) + "\h" + CStr(n_mode), _
		                                "H(" + sPort + "(" + CStr(n_mode) + ")): " + mode_text + "(" + sSub(nSub) + ")")
		                        ppt_AddSlide presentation, title
		                        title = CaptureTreeItem( _
		                                "2D/3D Results\Port Modes\" + sPort + "\" + sSub(nSub) + "\p" + CStr(n_mode), _
		                                "P(" + sPort + "(" + CStr(n_mode) + ")): " + mode_text + "(" + sSub(nSub) + ")")
		                        ppt_AddSlide presentation, title
		                    Next
		                End If
						n_mode = n_mode + 1
	                Loop While (FloquetPort.NextMode And n_mode <= nFloquetModes)
	        		End If
		        Next
		    End If ' (bIncludeFloquet)

	        Const MWS_1DResults = Array("Port signals", "S-Parameters", "|S| linear", "|S| dB", "arg(S)", "S polar", "Smith Chart", "Energy", "Balance")

			On Error Resume Next
        	iCount = 0
			Do
				ss = MWS_1DResults(iCount)
				ss1 = "1D Results\" + ss
				If (Resulttree.GetFirstChildName(ss1) <> "") Then
					If ss = "S-Parameters" Then
						Plot1D.PlotView "magnitudedb"
						Plot1D.Plot
					End If

					Wait 0.02
					title = CaptureTreeItem(ss1,ss)
			        ppt_AddSlide presentation, title

					If ss = "S-Parameters" Then
						Plot1D.PlotView "smith"
						Plot1D.Plot
						Wait 0.02
						title = CaptureTreeItem(ss1,ss)
				        ppt_AddSlide presentation, title
					End If
			    End If
				iCount = iCount + 1
			Loop While MWS_1DResults(iCount) <> ""
			On Error GoTo 0
			
		End If ' of bMWS - loop

        Const EMS_TreeEntries = Array("Current Paths", "Voltage Paths", "Coils", "Permanent Magnets", _
        			"Potentials", "Charges", "Current Ports", "Particle Sources")

		If bEMS Or bPS Then
			On Error Resume Next
        	iCount = 0
			Do
				ss = EMS_TreeEntries(iCount)
				If (Resulttree.GetFirstChildName(ss) <> "") Then
					title = CaptureTreeItem(ss,ss)
			        ppt_AddSlide presentation, title
			    End If	
				iCount = iCount + 1
			Loop While EMS_TreeEntries(iCount) <> ""
			On Error GoTo 0

			' --- make slides for selected 1D Results
			Const EMS_1DResults = Array("Current","Voltage","Losses","Torque","Incremental Inductance Matrix","Inductance Matrix", "Motion Induced Voltage","Iron Losses")

			Dim treepaths As Variant, resulttypes As Variant, filenames As Variant, resultinfo As Variant
        	Dim nResults As Long
        	Dim startEntryShortName As Integer
			nResults = Resulttree.GetTreeResults("1D Results","0D/1D folder recursive","",treepaths,resulttypes,filenames,resultinfo)

			For iCount=0 To nResults-1
				startEntryShortName = InStrRev(treepaths(iCount),"\") 'make shortname if not the root
				If startEntryShortName = 0 Then
					ss_short$ = treepaths(iCount)
				Else
					ss_short$ = Mid(treepaths(iCount), startEntryShortName+1)
				End If
				If FindListIndex(EMS_1DResults,ss_short)<>-1 Then 'when the current short name is found in the EMS_1DResults array, the result is added in powerpoint
					If ss_short = "Iron Losses" Then
						ss1 = Resulttree.GetFirstChildName(treepaths(iCount))
						While ss1 <> ""
							title = CaptureTreeItem(ss1,ss1)
						    ppt_AddSlide presentation, title
							ss1=Resulttree.GetNextItemName (ss1)
						Wend
					Else
						title = CaptureTreeItem(treepaths(iCount),treepaths(iCount))
				    	ppt_AddSlide presentation, title
			    	End If
				End If
			Next iCount

			' --- make slides for all 2D/3D Results

			ss = Resulttree.GetFirstChildName("2D/3D Results")   ' length=13
			While ss <> ""
				If ( ss="2D/3D Results\B-Field" Or ss="2D/3D Results\Potential" Or ss="2D/3D Results\Magnetic Energy Dens. [MQSTD]" ) Then
					' this is a treefolder, go through all children
					ilen2 = Len(ss)+2
					ss1 = Resulttree.GetFirstChildName(ss)   ' length=21 / 23
					While ss1 <> ""
						ss_short$ = Mid(ss1, ilen2)
						title = CaptureTreeItem(ss1,ss_short)
					    ppt_AddSlide presentation, title
						ss1=Resulttree.GetNextItemName (ss1)
					Wend
				ElseIf (ss="2D/3D Results\Iron Losses") Then
					'skip them
				Else
					ss_short$ = Mid(ss, 15)
					title = CaptureTreeItem(ss,ss_short)
				    ppt_AddSlide presentation, title
				End If
				ss=Resulttree.GetNextItemName (ss)
			Wend
			
		End If

		' --- make slides for all Tables
		
		ss = Resulttree.GetFirstChildName("Tables")   ' length=6
		While ss <> ""
			If ( ss="Tables\1D Results" Or ss="Tables\0D Results" ) Then
				ss1 = Resulttree.GetFirstChildName(ss)   ' length=17
				While ss1 <> ""
					ss_short$ = Mid(ss1, 19)
					title = CaptureTreeItem(ss1,ss_short)
				    ppt_AddSlide presentation, title
					ss1=Resulttree.GetNextItemName (ss1)
				Wend
			Else
				ss_short$ = Mid(ss, 8)
				title = CaptureTreeItem(ss,ss_short)
			    ppt_AddSlide presentation, title
			End If
			ss=Resulttree.GetNextItemName (ss)
		Wend
			

        '######################################################################
        'Save PowerPoint presentation, show slide selection
        '######################################################################

        With presentation
                .Application.ActiveWindow.ViewType = 1
                .Save
                If (dlg.Keep = 0) Then
                        If .Application.Presentations.Count = 1 Then
                                .Application.Quit
                        Else
                                .Close
                        End If
                Else
                        .Application.ActiveWindow.ViewType = 7
                        If (dlg.Show <> 0) Then
                                .Application.WindowState = 3
                                If (dlg.Run <> 0) Then
                                        .SlideShowSettings.Run
                                End If
                        End If
                End If
        End With
        ReportInformationToWindow(CStr(Time)+": The presentation is stored in: " + ppt_file)
        Set presentation = Nothing

        Exit Sub

        NO_PORT:
                n_port = -99
        Resume Next

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function CaptureTreeItem (Item As String, Text As String)

	Dim title As String

	If SelectTreeItem(Item + " (AR)") Then
		title = Text + " (AR)"
	Else 
		If SelectTreeItem(Item) Then
			title = Text
			If Item = "Components" Then
				SelectModelView
			End If
		Else
			title = " "
		End If
	End If
	Wait 0.3
	
    'Added by mru: sets background to white before taking slide.
	Dim r, g, b As String
	r = Plot.GetBackgroundColorR()
	g = Plot.GetBackgroundColorG()
	b = Plot.GetBackgroundColorB()
	Dim bGrad As Boolean
	bGrad = Plot.GetGradientBackground()
	If bGrad Then
		'set to white
		Plot.SetGradientBackground(False)
		Plot.SetBackgroundColor("1", "1", "1")

		'store view
		ExportImageToClipboard 0,0

		'restore background
		Plot.SetGradientBackground(bGrad)
		Plot.SetBackgroundColor(r, g, b)
		Plot.SetGradientBackground(True)
	Else
		Plot.SetBackgroundColor("1", "1", "1")

		'store view
		ExportImageToClipboard 0,0

		'restore background
		Plot.SetBackgroundColor(r, g, b)
    End If

    CaptureTreeItem = title

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Sub ppt_CreateTemplate(filepath As String)

        Dim SlideWidth As Single, SlideHeight As Single
        Dim i As Integer

        With CreateObject("PowerPoint.Application")
                With .Presentations.Add
                        SlideWidth = .PageSetup.SlideWidth
                        SlideHeight = .PageSetup.SlideHeight
                        With .SlideMaster
                                For i = .Shapes.Count To 1 STEP -1
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
                                For index = .Shapes.Count To 1 STEP -1
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
        End With

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

        Dim dummy As String, filename As String, extension As String
        Dim index As Integer
        Dim powerpoint As Object

        Select Case Action
                Case 1 ' Dialog box initialization
                        DlgEnable "Show", DlgValue("Keep")
                        DlgEnable "Run", DlgValue("Show")
                Case 2 ' Value changing or button pressed
                        DlgEnable "Show", DlgValue("Keep")
                        DlgEnable "Run", DlgValue("Show")
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
										StartHelp "common_preloadedmacro_report_powerpoint_presentation"
                                        ' ShowHelp(MacroName)
                                        DialogFunc = True
                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select
End Function

