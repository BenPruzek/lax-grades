Option Explicit

' Export port mode as near field source

'#include "vba_globals_all.lib"
'#include "mws_ports.lib"

' ================================================================================================
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 18-Feb-2016 fsr: Added support for ports aligned with any major axis; tetrahedral mesh now supported
' 22-Apr-2015 fsr: File export path copied to clipboard
' 21-Apr-2015 fsr: Added GUI
' 16-Apr-2015 fsr: Initial version
' ---------------------------------------------------------------------------------------------------------------------------

Sub Main ()

	' Initialize the global arrays first

	FillPortNumberArray

	FillModeNumberArray PortNumberArray(0)
	ModeNumberArray_OUT = ModeNumberArray

	FillModeNumberArray PortNumberArray(0)
	ModeNumberArray_IN = ModeNumberArray

	Begin Dialog UserDialog 440,196,"Export Port Mode as Near Field Data",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 20,7,400,147,"Settings",.GroupBox3
		Text 40,35,90,14,"Port:",.Text3
		DropListBox 100,28,90,112,PortNumberArray(),.Port_IN
		Text 220,35,50,14,"Mode:",.Text4
		DropListBox 290,28,90,112,ModeNumberArray_IN(),.Mode_IN
		Text 40,70,170,14,"Number of samples (width):",.Text1
		Text 40,98,180,14,"Number of samples (height):",.Text2
		TextBox 220,63,60,21,.XSamplesT
		TextBox 220,91,60,21,.YSamplesT
		CheckBox 290,98,120,14,"Same as width",.YSameAsXCB
		CheckBox 40,126,200,14,"Set source center to 0/0/0",.SourceCenterToOriginCB
		OKButton 220,161,90,21
		CancelButton 320,161,90,21
	End Dialog
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		Exit All
	End If

End Sub

Function ExportFieldOnSheet(dXPosition() As Double, dYPosition() As Double, dZPosition() As Double, _
								dXShift As Double, dYShift As Double, dZShift As Double, _
								sComponent As String, dFrequency As Double, dRealValues() As Double, dImagValues() As Double, _
								sExportPath As String, sExportPrefix As String) As Integer

	' Export the complex data dRealValues and dImagValues at the given positions dX/Y/ZPosition
	' sComponent describes the field component (Ex/Ey/Ez/Hx/Hy/Hz)
	' Field source data will be placed in sExportPath

	Dim i As Long, j As Long
	Dim dGeometryUnitsToSI As Double
	Dim sOutputXMLFileName As String, sOutputDATFileName As String, sOutputDATFileNameShort As String
	Dim iOutputXMLFileID As Long, iOutputDATFileID As Long

	dGeometryUnitsToSI = Units.GetGeometryUnitToSI

	sOutputXMLFileName = sExportPath & "\" & sExportPrefix & sComponent & ".xml"
	sOutputDATFileNameShort = sExportPrefix & sComponent & ".dat"
	sOutputDATFileName = sExportPath & "\" & sOutputDATFileNameShort

	iOutputXMLFileID = FreeFile()
	Open sOutputXMLFileName For Output As iOutputXMLFileID

	Print #iOutputXMLFileID, "<EmissionScan>"
	Print #iOutputXMLFileID, "<Nfs_ver>1.0</Nfs_ver>"
	Print #iOutputXMLFileID, "<Filename>" & sOutputXMLFileName & "</Filename>"
	Print #iOutputXMLFileID, "<Probe><Field>" & sComponent & "</Field></Probe>"
	Print #iOutputXMLFileID, "<Data>"
	Print #iOutputXMLFileID, "<Coordinates>xyz</Coordinates>"
	Print #iOutputXMLFileID, "<Frequencies>"
	Print #iOutputXMLFileID, "<List>" & Cstr(dFrequency*Units.GetFrequencyUnitToSI) & "</List>"
	Print #iOutputXMLFileID, "</Frequencies>"
	Print #iOutputXMLFileID, "<Measurement>"
	Print #iOutputXMLFileID, "<Format>ri</Format>"
	Print #iOutputXMLFileID, "<Data_files>" & sOutputDATFileNameShort & "</Data_files>"
	Print #iOutputXMLFileID, "</Measurement>"
	Print #iOutputXMLFileID, "</Data>"
	Print #iOutputXMLFileID, "</EmissionScan>"

	Close iOutputXMLFileID

	iOutputDATFileID = FreeFile()
	Open sOutputDATFileName For Output As iOutputDATFileID

	For i = 0 To UBound(dXPosition)
		Print #iOutputDATFileID, CStr((dXPosition(i)-dXShift)*dGeometryUnitsToSI) & " " & CStr((dYPosition(i)-dYShift)*dGeometryUnitsToSI) & " " & CStr((dZPosition(i)-dZShift)*dGeometryUnitsToSI) & " " & CStr(dRealValues(i)) & " " & CStr(dImagValues(i))
	Next

	Close iOutputDATFileID

End Function

Function StartExport()

	Dim i As Long, j As Long

	Dim sSelectedPortField As String
	Dim iPortNumber As Long, iModeNumber As Long
	Dim iPortOrientation As Long, dPortXMin As Double, dPortXMax As Double, dPortYMin As Double, dPortYMax As Double, dPortZMin As Double, dPortZMax As Double
	Dim bFlipEField As Boolean

	Dim nXSamples As Long, nYSamples As Long, nTotalSamples As Long
	Dim dXPosition() As Double, dYPosition() As Double, dZPosition() As Double, dAxis1Step As Double, dAxis2Step As Double
	Dim dAxis1Position() As Double, dAxis2Position() As Double, dPortPosition() As Double, dReTempPosition() As Double, dImTempPosition() As Double
	Dim sAxis1 As String, sAxis2 As String, sAxis3 As String
	Dim dAxis1Min As Double, dAxis1Max As Double, dAxis2Min As Double, dAxis2Max As Double, dPortPlane As Double
	Dim dXShift As Double, dYShift As Double, dZShift As Double
	Dim dEXReValues() As Double, dEXImValues() As Double, dEYReValues() As Double, dEYImValues() As Double, dEZReValues() As Double, dEZImValues() As Double
	Dim dHXReValues() As Double, dHXImValues() As Double, dHYReValues() As Double, dHYImValues() As Double, dHZReValues() As Double, dHZImValues() As Double
	Dim dFrequency As Double
	Dim sExportPath As String

	iPortNumber = DlgText("Port_IN")
	iModeNumber = DlgText("Mode_IN")

	dFrequency = Port.GetFrequency(iPortNumber, iModeNumber)
	sExportPath = GetProjectPath("Root")

	If (Mesh.GetMeshType = "Tetrahedral") Then
		Port.GetPortLocation(iPortNumber, iPortOrientation, dPortXMin, dPortXMax, dPortYMin, dPortYMax, dPortZMin, dPortZMax)
		' port orientation does not work as of 1/5/16, determine it manually
		If dPortXMin = dPortXMax Then
			iPortOrientation = 1
		ElseIf dPortYMin = dPortYMax Then
			iPortOrientation = 3
		ElseIf dPortZMin = dPortZMax Then
			iPortOrientation = 5
		Else
			ReportError("Export port mode as field source: Oblique port orientation is currently not supported.")
		End If
	Else
		Port.GetPortMeshCoordinates (iPortNumber, iPortOrientation, dPortXMin, dPortXMax, dPortYMin, dPortYMax, dPortZMin, dPortZMax)
	End If

	Select Case iPortOrientation
		Case -1
			ReportError("Export port mode as field source: The port needs to be aligned with one of the major axes.")
		Case 0,1 ' x direction
			' determine integration axes
			sAxis1 = "z"
			sAxis2 = "y"
			sAxis3 = "x"
			dPortPlane = dPortXMin
			dAxis1Min = dPortZMin
			dAxis1Max = dPortZMax
			dAxis2Min = dPortYMin
			dAxis2Max = dPortYMax
		Case 2,3 ' y direction
			sAxis1 = "x"
			sAxis2 = "z"
			sAxis3 = "y"
			dAxis1Min = dPortXMin
			dAxis1Max = dPortXMax
			dPortPlane = dPortYMin
			dAxis2Min = dPortZMin
			dAxis2Max = dPortZMax
		Case 4,5 ' z direction
			sAxis1 = "x"
			sAxis2 = "y"
			sAxis3 = "z"
			dAxis1Min = dPortXMin
			dAxis1Max = dPortXMax
			dAxis2Min = dPortYMin
			dAxis2Max = dPortYMax
			dPortPlane = dPortZMin
	End Select

	dXShift = IIf(DlgValue("SourceCenterToOriginCB") = 0, 0, (dAxis1Max+dAxis1Min)/2)
	dYShift = IIf(DlgValue("SourceCenterToOriginCB") = 0, 0, (dAxis2Max+dAxis2Min)/2)
	dZShift = IIf(DlgValue("SourceCenterToOriginCB") = 0, 0, dPortPlane)

	bFlipEField = False
	If (iPortOrientation = 0) Or (iPortOrientation = 2) Or (iPortOrientation = 5) Then
		bFlipEField = True ' flip x axis for negative port orientation
	End If

	sSelectedPortField = SelectTreeItem("2D/3D Results\Port Modes\Port" & CStr(iPortNumber) & "\e" & CStr(iModeNumber))

	nXSamples = Evaluate(DlgText("XSamplesT"))
	If (DlgValue("YSameAsXCB") = 1) Then
		nYSamples = nXSamples
	Else
		nYSamples = Evaluate(DlgText("YSamplesT"))
	End If

	nTotalSamples = nXSamples * nYSamples
	ReDim dAxis1Position(nTotalSamples-1)
	ReDim dAxis2Position(nTotalSamples-1)
	ReDim dPortPosition(nTotalSamples-1)

	dAxis1Step = (dAxis1Max-dAxis1Min)/(nXSamples-1)
	dAxis2Step = (dAxis2Max-dAxis2Min)/(nYSamples-1)

	' Prepare list to evaluate fields
	For i = 0 To nXSamples - 1
		For j = 0 To nYSamples - 1
			dAxis1Position(j + i*nYSamples) = dAxis1Min + i * dAxis1Step
			dAxis2Position(j + i*nYSamples) = dAxis2Min + j * dAxis2Step
			dPortPosition(j + i*nYSamples) = dPortPlane
		Next
	Next
	' Map axis1/2/port to x/y/z
	Select Case iPortOrientation
		Case 0,1 ' x direction
			dXPosition = dPortPosition
			dYPosition = dAxis2Position
			dZPosition = dAxis1Position
		Case 2,3 ' y direction
			dXPosition = dAxis1Position
			dYPosition = dPortPosition
			dZPosition = dAxis2Position
		Case 4,5 ' z direction
			dXPosition = dAxis1Position
			dYPosition = dAxis2Position
			dZPosition = dPortPosition
	End Select

	VectorPlot3D.SetPoints(dXPosition, dYPosition, dZPosition)
	VectorPlot3D.CalculateList()

	dEXReValues = VectorPlot3D.GetList(sAxis1 & "re")
	dEXImValues = VectorPlot3D.GetList(sAxis1 & "im")
	dEYReValues = VectorPlot3D.GetList(sAxis2 & "re")
	dEYImValues = VectorPlot3D.GetList(sAxis2 & "im")
	dEZReValues = VectorPlot3D.GetList(sAxis3 & "re")
	dEZImValues = VectorPlot3D.GetList(sAxis3 & "im")

	If bFlipEField Then
		' flip efield so that the exported field source is pointing in +z direction
		dReTempPosition = dEXReValues
		dImTempPosition = dEXImValues
		For i = 0 To UBound(dEXReValues)
			dEXReValues(i) = -dReTempPosition(i)
			dEXImValues(i) = -dImTempPosition(i)
		Next
		dReTempPosition = dEYReValues
		dImTempPosition = dEYImValues
		For i = 0 To UBound(dEXReValues)
			dEYReValues(i) = -dReTempPosition(i)
			dEYImValues(i) = -dImTempPosition(i)
		Next
	End If

	' ex and ey
	ExportFieldOnSheet(dAxis1Position(), dAxis2Position(), dPortPosition(), _
						dXShift, dYShift, dZShift, _
						"Ex", dFrequency, dEXReValues(), dEXImValues(), _
						sExportPath, "port" & CStr(iPortNumber) & "mode" & CStr(iModeNumber))
	ExportFieldOnSheet(dAxis1Position(), dAxis2Position(), dPortPosition(), _
						dXShift, dYShift, dZShift, _
						"Ey", dFrequency, dEYReValues(), dEYImValues(), _
						sExportPath, "port" & CStr(iPortNumber) & "mode" & CStr(iModeNumber))

	SelectTreeItem("2D/3D Results\Port Modes\Port" & CStr(iPortNumber) & "\h" & CStr(iModeNumber))
	VectorPlot3D.CalculateList()

	dHXReValues = VectorPlot3D.GetList(sAxis1 & "re")
	dHXImValues = VectorPlot3D.GetList(sAxis1 & "im")
	dHYReValues = VectorPlot3D.GetList(sAxis2 & "re")
	dHYImValues = VectorPlot3D.GetList(sAxis2 & "im")
	dHZReValues = VectorPlot3D.GetList(sAxis3 & "re")
	dHZImValues = VectorPlot3D.GetList(sAxis3 & "im")

	' hx and hy
	ExportFieldOnSheet(dAxis1Position(), dAxis2Position(), dPortPosition(), _
						dXShift, dYShift, dZShift, _
						"Hx", dFrequency, dHXReValues(), dHXImValues(), _
						sExportPath, "port" & CStr(iPortNumber) & "mode" & CStr(iModeNumber))
	ExportFieldOnSheet(dAxis1Position(), dAxis2Position(), dPortPosition(), _
						dXShift, dYShift, dZShift, _
						"Hy", dFrequency, dHYReValues(), dHYImValues(), _
						sExportPath, "port" & CStr(iPortNumber) & "mode" & CStr(iModeNumber))


	MsgBox("Export finished. Files are located in: " & GetProjectPath("Root") & vbNewLine & "File path has been copied to clipboard.", "Success")
	Clipboard(sExportPath)

End Function

Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean

	Dim nIndex As Long

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgValue("YSameAsXCB", 1)
		DlgEnable("YSamplesT", Not (DlgValue("YSameAsXCB") = 1))
		DlgText("XSamplesT", "101")
		DlgText("YSamplesT", "101")
		DlgValue("SourceCenterToOriginCB", 1)
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box
		If (DlgItem = "Port_IN") Then

			nIndex = DlgValue("Port_IN")
			FillModeNumberArray PortNumberArray(nIndex)

			ModeNumberArray_IN = ModeNumberArray

			DlgListBoxArray "Mode_IN", ModeNumberArray_IN
			DlgValue("Mode_IN", 0)

		ElseIf (DlgItem = "OK") Then
			DlgEnable("OK", False)
			DlgEnable("Cancel", False)
			StartExport()
		End If

		DlgEnable("YSamplesT", Not (DlgValue("YSameAsXCB") = 1))

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
