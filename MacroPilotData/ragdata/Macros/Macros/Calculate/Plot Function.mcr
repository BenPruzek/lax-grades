'#Language "WWB-COM"

'#include "vba_globals_all.lib"

' This macro plots a function from a user-entered expression and adds the plot to the tree

' Version history:
' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' --------------------------------------------------------------------------------------------------------------------
' 16-Nov-2018 fsr: Fixed a problem with negative X values
' 16-Jul-2018 fsr: Replace Exit button with Cancel button to enable windows 'X'
' 31-Mar-2017 tgl: Added support for log x-axis
' 17-Nov-2014 fsr: Added DS compatibility; singularities are now skipped; option to enter plot name
' 30-Oct-2014 fsr: Initial version
' --------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main

	Begin Dialog UserDialog 530,189,"Function Plotter",.DialogFunction ' %GRID:10,7,1,1
		Text 20,14,210,14,"Function to be plotted (over $X):",.Text1
		TextBox 20,35,490,21,.sFunctionT
		Text 20,77,70,14,"From $X =",.Text2
		TextBox 150,70,160,21,.sXMinT
		Text 320,77,30,14,"to",.Text3
		TextBox 350,70,160,21,.sXMaxT
		Text 20,105,120,14,"Number of samples:",.Text4
		TextBox 150,98,160,21,.sNSamplesT
		CheckBox 350,105,130,14,"Log Sampling",.bMakeLog
		Text 20,133,120,14,"Plot name:",.Text5
		TextBox 150,126,360,21,.sPlotNameT
		PushButton 220,154,90,21,"Plot",.PlotPB
		PushButton 320,154,90,21,"Show All",.ShowAllPB
		CancelButton 420,154,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.sFunctionT = "Exp(-$X)*Sin(10*Pi*$X)"
	dlg.sXMinT = "0"
	dlg.sXMaxT = "1"
	dlg.sNSamplesT = "101"
	dlg.sPlotNameT = "Auto"
	dlg.bMakeLog = 0

	If Dialog(dlg, 1) = 0 Then Exit All

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "Cancel"
				Exit All
			Case "PlotPB"
				DialogFunction = True

				Dim sFunction As String, sPlotName As String
				Dim i As Long, nSamples As Long
				Dim dXMin As Double, dXMax As Double, dXValue As Double, dYValue As Double
				Dim oResultPlot As Object
				Dim bMakeLog As Boolean

				sFunction = DlgText("sFunctionT")
				dXMin = Evaluate(DlgText("sXMinT"))
				dXMax = Evaluate(DlgText("sXMaxT"))
				nSamples = Evaluate(DlgText("sNSamplesT"))
				bMakeLog = CBool(DlgValue("bMakeLog"))
				sPlotName = DlgText("sPlotNameT")

				If (Left(GetApplicationName(), 2) = "DS") Then
					Set oResultPlot = DS.Result1D("")
				Else
					Set oResultPlot = Result1D("")
				End If
				If bMakeLog Then
					If dXMin <= 0 Then
						MsgBox("log plotting does not allow xMin <= 0.",vbCritical)
						Exit Function
					End If
					dXMin = Log(dXMin)/Log(10)
					dXMax = Log(dXMax)/Log(10)
				End If
				For i = 1 To nSamples
					On Error Resume Next ' to skip potential division by zero and continue
					If bMakeLog Then
						dXValue = 10^(dXMin + (i-1)*(dXMax-dXMin)/(nSamples-1))
					Else
						dXValue = dXMin + (i-1)*(dXMax-dXMin)/(nSamples-1)
					End If
					dYValue = Evaluate(Replace(sFunction, "$X", "(" & CStr(dXValue) & ")"))
					oResultPlot.AppendXY(dXValue, dYValue)
				Next
				If oResultPlot.GetN() = 0 Then
					MsgBox("Plotting failed, please check your settings.")
				Else
					oResultPlot.XLabel("X")
					If(sPlotName = "Auto") Then sPlotName =  Cstr(sFunction)+" ("+Cstr(nSamples)+" samples)"
					AddPlotToTree_LIB(oResultPlot, "Plot Function\" + sPlotName, True)
				End If

			Case "ShowAllPB"
				DialogFunction = True
				If (Left(GetApplicationName(), 2) = "DS") Then
					DS.SelectTreeItem("Results\1D Results")
				Else
					SelectTreeItem("1D Results\Plot Function")
				End If
			Case "bMakeLog"
				
			Case Else
				ReportError("Unknown dialog item.")
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
