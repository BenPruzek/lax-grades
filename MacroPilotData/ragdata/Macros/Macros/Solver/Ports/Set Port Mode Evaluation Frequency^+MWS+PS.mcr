'#Language "WWB-COM"

'#include "vba_globals_all.lib"

' This macro can be used to set the port mode calculation frequency for PIC simulations. It allows to enter arbitrary frequency values,
' which might be difficult or impossible to set up using the slider in the solver dialog. By default, it also disables the feature that
' TE/TM modes are always calculated at the center frequency, instead of the specified frequency.
'
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' -----------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 27-Jan-2017 mbk: Fixed a problem with non-US locale
' 26-Jan-2017 fsr: Initial version
' -----------------------------------------------------------------------------------------------


Option Explicit

Sub Main

	ActivateScriptSettings True
	' ClearScriptSettings

	Begin Dialog UserDialog 450,161,"Port Mode Calculation Frequency",.DlgFunc ' %GRID:10,7,1,1
		Text 20,28,250,14,"Set port mode calculation frequency to:",.Text1
		TextBox 290,21,90,21,.FrequencyT
		Text 390,28,40,14,"",.UnitsL
		CheckBox 20,56,370,14,"Use center frequency instead for TE/TM modes",.UseCenterFreqCB
		Text 20,84,410,28,"Note: This macro needs to be re-run if the frequency range of the project is changed",.Text2
		OKButton 240,126,90,21
		CancelButton 340,126,90,21
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then
		Exit All
	End If

	ActivateScriptSettings False

End Sub

Rem See DialogFunc help topic for more information.
Private Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim sCommand As String
	Dim dFMin As Double, dFMax As Double, dFrequencyFactor As Double

	dFMin = Solver.GetFmin
	dFMax = Solver.GetFmax

	Select Case Action%
	Case 1 ' Dialog box initialization
		If ((dFMax = 0) And (dFMin = 0)) Then
			MsgBox("Frequency range is undefined. Please set the frequency range before running this macro.", vbCritical, "Error")
			Exit All
		End If
		DlgText("UnitsL", Units.GetUnit("Frequency"))
		DlgText("FrequencyT", CStr(dFMin + (dFMax - dFMin) * Evaluate(GetScriptSetting("FrequencyFactor", "0.5"))))
		DlgValue("UseCenterFreqCB", GetScriptSetting("UseCenterFreq", "0"))
	Case 2 ' Value changing or button pressed
		Rem DlgFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "OK"
				If ((Evaluate(DlgText("FrequencyT")) > dFMax) Or (Evaluate(DlgText("FrequencyT")) < dFMin)) Then
					MsgBox("Selected frequency is outside of the project frequency range. Please check your settings. ", vbExclamation, "Check Settings")
					DlgFunc = True
				Else
					dFrequencyFactor = (Evaluate(DlgText("FrequencyT")) - dFMin)/(dFMax- dFMin)
					StoreScriptSetting("FrequencyFactor", CStr(dFrequencyFactor))
					StoreScriptSetting("UseCenterFreq", CStr(DlgValue("UseCenterFreqCB")))
					sCommand = ""
					sCommand = AppendHistoryLine_LIB(sCommand, "Solver.SetModeFreqFactor", dFrequencyFactor)
					sCommand = AppendHistoryLine_LIB(sCommand, "Solver.ScaleTETMModeToCenterFrequency", IIf(DlgValue("UseCenterFreqCB") = 0, "False", "True"))
					AddToHistory("define special time domain solver parameters", sCommand)
					DlgFunc = False ' close the dialog
				End If
			Case "Cancel"
				DlgFunc = False ' close the dialog
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DlgFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
