'#include "vba_globals_all.lib"

' This template generates a 3D "radar range equation" pattern, i.e. at which distance an object of a given size
' can still be detected, assuming a given antenna/receiver configuration.

' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 16-Jan-2012 fsr: Dialog window remains open after calculation
' 27-Oct-2011 fsr: Initial version

Sub Main

	Dim sFFResult As String, sFFResultList() As String

	ReDim sFFResultList(0)
	sFFResult = Resulttree.GetFirstChildName("Farfields")
	If sFFResult = "" Then
		MsgBox("No farfield results found.")
	Else
		While sFFResult <> ""
			If (sFFResult <> "Farfields\Radar Range Pattern [m]") Then
				If (sFFResultList(UBound(sFFResultList)) <> "") Then ReDim Preserve sFFResultList(UBound(sFFResultList)+1)
				sFFResultList(UBound(sFFResultList)) = sFFResult
			End If
			sFFResult = Resulttree.GetNextItemName(sFFResult)
		Wend
	End If

	Begin Dialog UserDialog 400,196,"Generate 3D Radar Range Pattern",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 140,14,240,192,sFFResultList(),.TransmitAntennaDLB
		Text 20,21,110,14,"Antenna pattern:",.Text1
		Text 20,49,160,14,"Frequency ("+Units.GetUnit("Frequency")+"):",.Text2
		TextBox 200,42,180,21,.FrequencyT
		Text 20,77,160,14,"RCS of target (m^2):",.Text4
		TextBox 200,70,180,21,.RCST
		Text 20,105,160,14,"Transmitted power (W):",.Text3
		TextBox 200,98,180,21,.TransmitPowerT
		Text 20,133,170,14,"Min. detectable signal (W):",.Text5
		TextBox 200,126,180,21,.MinDetectSigT
		OKButton 200,161,90,21
		CancelButton 300,161,80,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		If InStr(DlgText("TransmitAntennaDLB"), "(f=") Then
			DlgText("FrequencyT",Split(Split(DlgText("TransmitAntennaDLB"),"(f=")(1),")")(0))
		Else
			DlgText("FrequencyT", "1")
		End If
		DlgText("RCST","1e-3")
		DlgText("TransmitPowerT","1e3")
		DlgText("MinDetectSigT","1")
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "TransmitAntennaDLB"
				If InStr(DlgText("TransmitAntennaDLB"), "(f=") Then
					DlgText("FrequencyT",Split(Split(DlgText("TransmitAntennaDLB"),"(f=")(1),")")(0))
				End If
			Case "Cancel"
				DialogFunc = False
				Exit All
			Case "OK"
				DlgEnable("OK", False)
				DlgEnable("Cancel", False)
				Generate3DMonostaticRadarRangePattern(DlgText("TransmitAntennaDLB"), Evaluate(DlgText("FrequencyT"))*Units.GetFrequencyUnitToSI, Evaluate(DlgText("TransmitPowerT")), Evaluate(DlgText("MinDetectSigT")), Evaluate(DlgText("RCST")))
				MsgBox("Done, plot units are 'meters'.", "Radar Range Pattern")
				DlgEnable("OK", True)
				DlgEnable("Cancel", True)
				DialogFunc = True
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function Generate3DMonostaticRadarRangePattern(sAntennaPattern As String, _
												dFrequency As Double, _
												dTransmitPower As Double, _
												dMinDetectSignal As Double, _
												dTargetRCS As Double, _
												Optional dThetaStep As Double, _
												Optional dPhiStep As Double, _
												Optional dThetaMin As Double, _
												Optional dThetaMax As Double, _
												Optional dPhiMin As Double, _
												Optional dPhiMax As Double)

	Dim dTheta As Double, dPhi As Double, i As Long
	Dim nPhiSamples As Integer, nThetaSamples As Integer
	Dim sOutputFileName As String, iOutputFile As Integer, sFarfieldName As String, dMonitorFrequency As Double

	Dim dGainTheta As Double, dGainPhi As Double, dRadarRangeTheta As Double, dRadarRangePhi As Double

	sOutputFileName = GetProjectPath("Result")+"RadarRange.ffs"

	' Set up parameters
	dThetaMin = 0
	dThetaMax = 180
	dThetaStep = 5
	dPhiMin = 0
	dPhiMax = 360
	dPhiStep = 5
	nThetaSamples = Fix((dThetaMax-dThetaMin)/dThetaStep) + 1
	nPhiSamples = Fix((dPhiMax-dPhiMin)/dPhiStep) + 1

	iOutputFile = FreeFile()

	' Open farfield source file
	Open sOutputFileName For Output As iOutputFile
	Print #iOutputFile, "// CST Farfield Source File"
	Print #iOutputFile, "// Version:"
	Print #iOutputFile, "2.1"
	Print #iOutputFile, "// Data Type"
	Print #iOutputFile, "Farfield"
	Print #iOutputFile, "// Position"
	Print #iOutputFile, "0 0 0"
	Print #iOutputFile, "// z-Axis"
	Print #iOutputFile, "0 0 1"
	Print #iOutputFile, "// x-Axis"
	Print #iOutputFile, "1 0 0"
	Print #iOutputFile, "// Radiated Power"
	Print #iOutputFile, "-1"
	Print #iOutputFile, "// Accepted Power"
	Print #iOutputFile, "-1"
	Print #iOutputFile, "// Stimulated Power"
	Print #iOutputFile, "-1"
	Print #iOutputFile, "// Frequency"
	Print #iOutputFile, CStr(dFrequency)
	Print #iOutputFile, "// Total number of phi and theta samples"
	Print #iOutputFile, CStr(nPhiSamples)+" "+CStr(nThetaSamples)
	Print #iOutputFile, "// phi, theta, Re(Etheta), Im(Etheta), Re(Ephi), Im(Ephi)"

	SelectTreeItem(sAntennaPattern)
	FarfieldPlot.Reset
	FarfieldPlot.PlotType("3D")
	FarfieldPlot.SetPlotMode("gain")
	FarfieldPlot.IncludeUnitCellSideWalls(True)
	FarfieldPlot.SetScaleLinear(True)
	FarfieldPlot.Plot
	For dPhi = dPhiMin To dPhiMax STEP dPhiStep
		For dTheta = dThetaMin To dThetaMax STEP dThetaStep
			' Evaluate farfield in correct position
			FarfieldPlot.AddListItem(dTheta, dPhi, 1)
		Next dTheta
	Next dPhi
	FarfieldPlot.CalculateList(sAntennaPattern)

	For i = 0 To (nPhiSamples*nThetaSamples-1)
			' Evaluate farfield in correct position
			dGainTheta = FarfieldPlot.GetListItem(i,"spherical linear theta abs")
			dGainPhi = FarfieldPlot.GetListItem(i,"spherical linear phi abs")
			dRadarRangeTheta = Sqr(Sqr(dTransmitPower*dGainTheta^2*CLight^2*dTargetRCS/dFrequency^2/(4*Pi)^3/dMinDetectSignal))
			dRadarRangePhi = Sqr(Sqr(dTransmitPower*dGainPhi^2*CLight^2*dTargetRCS/dFrequency^2/(4*Pi)^3/dMinDetectSignal))
			If (FarfieldPlot.GetListItem(i,"Point_P")=0 And FarfieldPlot.GetListItem(i,"Point_T")=0) Then
			'	ReportInformationToWindow(dTransmitPower)
			'	ReportInformationToWindow(dTargetRCS)
			'	ReportInformationToWindow(dFrequency)
			'	ReportInformationToWindow(dMinDetectSignal)
			'	ReportInformationToWindow(dGainTheta)
			'	ReportInformationToWindow(dGainPhi)
			'	ReportInformationToWindow(dRadarRangeTheta)
			'	ReportInformationToWindow(dRadarRangePhi)
			End If
			Print #iOutputFile, CStr(FarfieldPlot.GetListItem(i,"Point_P"))+" "+CStr(FarfieldPlot.GetListItem(i,"Point_T"))+" "+CStr(dRadarRangeTheta)+" "+CStr(0)+" "+CStr(dRadarRangePhi)+" "+CStr(0)
	Next i

	Close iOutputFile
	'Shell "notepad " & sOutputFileName, 3

	' Add to tree
	With Resulttree
		.Name "Farfields\Radar Range Pattern [m]"
		.File sOutputFileName
		.Type "Farfield"
		.Add
	End With

	SelectTreeItem("Farfields\Radar Range Pattern [m]")
	FarfieldPlot.PlotType("3D")
	FarfieldPlot.SetPlotMode("efield")
	FarfieldPlot.Plot

End Function
