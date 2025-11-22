
' ================================================================================================
' This macro enables/disables support of geometric symmetry in the I-solver
' ================================================================================================
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 19-Jun-2023 jfl: initial version
' ================================================================================================

Option Explicit


Sub Main

	Begin Dialog UserDialog 570,168,"Allow non-symmetric excitation",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,540,63,"INFO",.GroupBox4
		Text 30,28,510,35,"This script enables/disables the calculation of geometric symmetry even if excitation is not symmetric in the I-Solver.",.Text1
		GroupBox 10,77,540,56,"Currently active settings",.GroupBox1
		OKButton 20,140,100,21
		CancelButton 130,140,100,21
		CheckBox 30,98,290,28,"Allow non-symmetric excitations",.CheckBox1
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then Exit All

	If (dlg.CheckBox1 = 1) Then
		AddToHistory("Support Geometric Symmetry","FDSolver.AllowGeometricSymmetry ""True""" )
		FDSolver.AllowGeometricSymmetry "True"
	Else
		AddToHistory("Don't Support Geometric Symmetry","FDSolver.AllowGeometricSymmetry ""False""" )
		FDSolver.AllowGeometricSymmetry "False"
	End If
	' Debug.Print dlg.CheckBox1

End Sub

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	Select Case Action%
	    Case 1 ' Dialog box initialization
	    	Debug.Print FDSolver.IsGeometricSymmetrySupported
			If FDSolver.IsGeometricSymmetrySupported = True Then
				DlgValue "CheckBox1", 1
			Else
				DlgValue "CheckBox1", 0
			End If

    	Case 2 ' Value changing or button pressed
    	Case 3 ' TextBox or ComboBox text changed
    	Case 4 ' Focus changed
    	Case 5 ' Idle
    	Case 6 ' Function key

    End Select

End Function
