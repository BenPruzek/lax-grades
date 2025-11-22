' *Report / Word Report
' !!! Do not change the line above !!!
' macro.562
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 30-Jun-2021 ube: correct calculation of max S11 dB  (old method d1(1)1(1).sig no longer supported)
' 18-Dec-2013 ube: Time Signals excluded, if not existing
' 03-Jun-2011 ube: complex S-parameters (db) were not properly handled
' 03-Jul-2007 ube: created report file stored at same level as cst-file
' 21-Oct-2005 imu: Included into Online Help
' 10-Nov-2003 ube: example for word-Report
'--------------------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_Report_Word_Report"

'#Uses "WordReport.cls"
'#include "vba_globals_all.lib"

Sub Main

	Begin Dialog UserDialog 430,84,"Create Word Report",.DialogFunc ' %GRID:10,7,1,1
		Text 40,21,390,14,"This will create a MS-Word report on the current simulation",.Text1
		CancelButton 160,56,100,21
		OKButton 40,56,100,21
		PushButton 280,56,100,21,"Help",.Help
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg)=0   Then Exit All

    Dim projectname As String, ppt_file As String
    Dim version As Integer, filename As String, presentation As Object
    Dim title As String, projectdir As String
	Dim iCount As Integer
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

	Dim report As New WordReport, sdocfile As String

	sdocfile = projectdir + "\" + projectname + Format(version, "\_00\.\d\o\c")

	With report
		
		.NewFile sdocfile
		
		.NewLine
		.NewLine
		.WriteTitle "Example For Automatic Report Generation"
		.WriteTitle "by Using Microsoft(R) Word 2000"
		.NewLine
		.NewLine
		
		.WriteParagraph "This simple example demonstrates how reports can be generated automatically " & _
							  "by using the WordReport VBA macro class."
		.NewLine
							  
		.WriteHeading "1. Structure Visualisation"							  
		.WriteParagraph "The following plot shows the structure investigated in this example. " & _
		 				"The defined parameters are listed in the table below."

		.WriteStructurePlot
		.NewLine
		.WriteParameterTable
		
		.NewPage

		If bEMS Then

			' for CST EM STUDIO no default 1D Results exist

			.WriteHeading "2. Simulation Results"
			.WriteParagraph "Here some existing 1D Signals or Tables could be displayed " & _
							"using the routine WriteResultPlot"
			.NewLine
		End If

		If bMWS Then

			If (Resulttree.GetFirstChildName("1D Results\Port signals") <> "") Then

				.WriteHeading "Simulation Time Signals"
				.WriteParagraph "The following plot shows the time signals which describe the mode amplitudes at " & _
								"the waveguide ports."
				.NewLine

				.WriteResultPlot "1D Results\Port signals"

				.NewPage

			End If

			.WriteHeading "2. S-Parameter Results"
			.WriteParagraph "The following plot shows the S-parameters as a function of frequency."
			.NewLine

			If (Resulttree.GetFirstChildName("1D Results\S-Parameters") <> "") Then
				Plot1D.PlotView "magnitudedb"
				Plot1D.Plot
				.WriteResultPlot "1D Results\S-Parameters"
			Else
				.WriteResultPlot "1D Results\|S| dB"
			End If

			Dim oComplex1DC As Object, s11 As Object
			Dim maxSlinear As Double, maxSdB As Double

			On Error GoTo NoS11
				Set oComplex1DC = Result1DComplex("cS1(1)1(1)")

				On Error GoTo 0

				Set s11 = oComplex1DC.Magnitude

				maxSlinear = s11.GetY(s11.GetGlobalMaximum())
				maxSdB = 20*Log(maxSlinear)/Log(10)

				' btw: logarithmic factor (we know it is 20 for S-Parameters) could also be obtained from   "oComplex1DC.GetLogarithmicFactor"

				.WriteHeading "4. Remarks"
				.WriteParagraph "The maximum return loss within the simulated frequency range is " & CStr(maxSdB) & " dB."

			NoS11:

		End If ' of bMWS

		.CloseFile
	
	End With

	MsgBox "The report has successfully been generated. " & vbCr & sdocfile
	
End Sub
'--------------------------------------------------------------------------------------------
Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "Help"
			StartHelp HelpFileName
			DialogFunc = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
