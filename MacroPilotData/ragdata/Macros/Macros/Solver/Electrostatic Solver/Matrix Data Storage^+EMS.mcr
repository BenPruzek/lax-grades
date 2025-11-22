' move-mesh

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 21-Jun-2018 ube: First version
' ================================================================================================
Sub Main () 

	Begin Dialog UserDialog 520,140,"Electrostatic Solver" ' %GRID:10,7,1,1
		GroupBox 20,7,480,98,"",.GroupBox1
		CheckBox 40,28,440,14,"Enable fast Matrix Storage for C Matrix coefficients",.UseMatrixValuedSQLStorage
		OKButton 30,112,90,21
		CancelButton 130,112,90,21
		Text 70,56,410,42,"Note: Access in Result Templates via Template"+vbCrLf+vbCrLf+"          Statics and Low Frequency -> Get Matrix Coefficient",.Text1
	End Dialog
	Dim dlg As UserDialog

	dlg.UseMatrixValuedSQLStorage = EStaticSolver.GetUseMatrixValuedSQLStorage

	If (Dialog(dlg) = 0) Then Exit All

	' Delete Result Tree, otherwise old entries may still be displayed
	DeleteResults

	Dim sCommand As String
	sCommand = ""

	If dlg.UseMatrixValuedSQLStorage Then
		sCommand = sCommand + "EStaticSolver.UseMatrixValuedSQLStorage ""true""" + vbCrLf
	Else
		sCommand = sCommand + "EStaticSolver.UseMatrixValuedSQLStorage ""false""" + vbCrLf
	End If

	' MsgBox sCommand
	AddToHistory "(*) define Matrix Storage for C Matrix", sCommand

End Sub
