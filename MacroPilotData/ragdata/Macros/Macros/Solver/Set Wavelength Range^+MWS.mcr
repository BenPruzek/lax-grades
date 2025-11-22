'#Language "WWB-COM"

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 07-Nov-2014 ube: First version
' ================================================================================================
Option Explicit

Sub Main

	Dim dFMin As Double, dFMax As Double

	Begin Dialog UserDialog 330,112,"Wavelength Range Settings" ' %GRID:10,7,1,1
		Text 20,14,90,14,"Lambda min:",.Text1
		Text 20,56,90,14,"Lambda max:",.Text2
		TextBox 20,28,190,21,.LambdaMinT
		TextBox 20,70,190,21,.LambdaMaxT
		OKButton 230,14,90,21
		CancelButton 230,42,90,21
	End Dialog
	Dim dlg As UserDialog

	dFMin = Solver.GetFmin
	dFMax = Solver.GetFmax

	If(dFMax <> 0) Then
		dlg.LambdaMinT = CStr(Round(CLight/(dFMax*Units.GetFrequencyUnitToSI)/Units.GetGeometryUnitToSI*1e6)/1e6)
	Else
		dlg.LambdaMinT = "999999"
	End If

	If(dFMin <> 0) Then
		dlg.LambdaMaxT = CStr(Round(CLight/(dFMin*Units.GetFrequencyUnitToSI)/Units.GetGeometryUnitToSI*1e6)/1e6)
	Else
		dlg.LambdaMaxT = "999999"
	End If

	If Dialog(dlg, -1) = 0 Then
		Exit All
	Else
		If(Evaluate(dlg.LambdaMinT) <> 0) Then
			dFMax = CLight/(Evaluate(dlg.LambdaMinT)*Units.GetGeometryUnitToSI)/Units.GetFrequencyUnitToSI
		Else
			dFMax = 999999
		End If

		If(Evaluate(dlg.LambdaMaxT) <> 0) Then
			dFMin = CLight/(Evaluate(dlg.LambdaMaxT)*Units.GetGeometryUnitToSI)/Units.GetFrequencyUnitToSI
		Else
			dFMin = 999999
		End If
		AddToHistory("define frequency range", "Solver.FrequencyRange " + Chr(34) + CStr(dFMin) + Chr(34) + ", " + Chr(34) + CStr(dFMax) + Chr(34))
	End If


End Sub
