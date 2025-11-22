'#Language "WWB-COM"

'----------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2021-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 11-Jun-2021 ube,fmo: initial version
'----------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main ()

	Begin Dialog UserDialog 670,182,"FD-TET-ROM - Allow Sensitivities for Structure Parameters" ' %GRID:10,7,1,1
		Text 30,56,610,77,"This macro enables Sensitivity Analysis for all structure parameters. Please note, that for every activated parameter a little computational overhead is required." + vbCrLf  + vbCrLf +  "Note: When activating this feature, by default ALL structure parameters are selected. You may want to deselect some manually. The macro also activates the FD-TET ROM solver as active solver and will trigger an automatic 'Save'.",.Text5
		OKButton 20,147,90,21
		CancelButton 120,147,90,21
		CheckBox 40,21,330,14,"Activate Sensitivities for Structure Parameters",.CheckBox1
	End Dialog
	Dim dlg As UserDialog

	dlg.CheckBox1 = IIf(SensitivityAnalysis.GetAllowStructureParameters(), 1, 0)

	If (Dialog(dlg) = -1) Then

		Dim sCommand As String
		sCommand = ""

		If dlg.CheckBox1 Then
			' 1) select FD-solver
			' ChangeSolverType "HF Frequency Domain"
			sCommand = "ChangeSolverType ""HF Frequency Domain""" + vbCrLf

			' 2) Activate the structure sensitivities
			' SensitivityAnalysis.SetAllowStructureParameters "True"
			sCommand = sCommand + "SensitivityAnalysis.SetAllowStructureParameters ""True""" + vbCrLf

			' 3) activate sensitivity and the featured method in F-Solver
			' FDSolver.SetMethod "Tetrahedral", "Fast reduced order model"
			' FDSolver.UseSensitivityAnalysis "True"
			sCommand = sCommand + "FDSolver.SetMethod ""Tetrahedral"", ""Fast reduced order model""" + vbCrLf
			sCommand = sCommand + "FDSolver.UseSensitivityAnalysis ""True""" + vbCrLf

		Else
			' just deactivate the structure sensitivities
			' SensitivityAnalysis.SetAllowStructureParameters "False"
			sCommand = sCommand + "SensitivityAnalysis.SetAllowStructureParameters ""False""" + vbCrLf
		End If

		' MsgBox sCommand
		AddToHistory "(*) FD-TET-ROM - Allow Sensitivities for Structure Parameters", sCommand

	End If
	
	Save  ' to prevent that feature is used in unsaved untitled project

End Sub
