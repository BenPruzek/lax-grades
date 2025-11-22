'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Copyright 2022-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 25-Jun-2024 fmr: First version
' ================================================================================================

Function bi2s(i As Integer) As String

	If 0 = i Then
		bi2s = "false"
	Else
		bi2s = "true"
	End If

End Function

Function getIndex(strVal As String ) As Integer
	getIndex = 0

	If(strVal = "Minmax") Then
		getIndex = 1
	ElseIf(strVal = "Maxmin") Then
		getIndex = 2
	End If

End Function



Sub Main

	Dim ObjectiveFormulationList(2) As String

	ObjectiveFormulationList$(0) = "Weighted sum"
	ObjectiveFormulationList$(1) = "Minmax"
	ObjectiveFormulationList$(2) = "Maxmin"

	Begin Dialog UserDialog 360,238,"Non-Parametric Optimization Settings" ' %GRID:10,7,1,1
		GroupBox 10,161,330,49,"General",.GroupBox1
		GroupBox 10,63,330,91,"Topology Optimization",.GroupBox2
		GroupBox 10,7,330,49,"Optimization Settings",.GroupBox3
		CheckBox 30,182,240,14,"Show Design Response Curves",.CheckBoxDRESP1D
		CheckBox 30,126,240,14,"Show material distribution monitor",.CheckBoxMatDistr
		Text 30,91,130,14,"Isocut Threshold:",.Text1
		TextBox 180,88,90,21,.IsocutThreshhold

		Text 30,32,170,14,"Objective Formulation:",.Text2
		DropListBox 180,29,140,70,ObjectiveFormulationList(),.iobjF

		OKButton 150,214,90,21
		CancelButton 250,214,90,21
	End Dialog

	Dim dlg As UserDialog

	Dim oldIsocut As String
	Dim showDRESP1D As String
	Dim showMateralDistr As String
	Dim old_objFormulation As String

	Dim key_word As String
	key_word = "NPOSetting;"

	Dim settingCut As String
	Dim dresp1DOut As String
	Dim matdistOut As String
	Dim objFormulation As String

	settingCut = "Isocut Threshold"
	dresp1DOut = "Show Design Response Curves"
	matdistOut = "Show Material Distribution Monitor"
	objFormulation = "Objective Formulation"

	With Optimizer
		oldIsocut = .GetSetting(key_word+settingCut )
		showDRESP1D = .GetSetting(key_word+dresp1DOut)
		showMateralDistr = .GetSetting(key_word+matdistOut)
		old_objFormulation = .GetSetting(key_word+objFormulation)
	End With

	dlg.IsocutThreshhold = oldIsocut
    dlg.CheckBoxDRESP1D = (showDRESP1D = "true")
    dlg.CheckBoxMatDistr = (showMateralDistr = "true")
    dlg.iobjF = getIndex(old_objFormulation)
	' Run the dialog
	Dim iDlg As Integer
	iDlg = Dialog(dlg)

	' On OK pressed
	If (-1 = iDlg) Then

		' Solver Settings

		With Optimizer

			If Not dlg.IsocutThreshhold = oldIsocut Then
				.SetSetting(key_word+settingCut, dlg.IsocutThreshhold)
			End If

			Dim helper As String
			helper = bi2s(dlg.CheckBoxDRESP1D)
			If Not helper = showDRESP1D Then
				.SetSetting(key_word+dresp1DOut, helper)
			End If

			helper = bi2s(dlg.CheckBoxMatDistr)
			If Not helper = showMateralDistr Then
				.SetSetting(key_word+matdistOut, helper)
			End If

			helper = ObjectiveFormulationList(dlg.iobjF)
			If Not helper = old_objFormulation Then
				.SetSetting(key_word+objFormulation, helper)
			End If


		End With

	End If

End Sub
