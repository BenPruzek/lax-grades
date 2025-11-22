
' ================================================================================================
' Purpose: Enable settiing of more accuracy settings
' ================================================================================================
' Copyright 2020-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' -------------------
' 23-Jan-2020 jfl: initial version
' ================================================================================================
'#include "vba_globals_all.lib"

Sub Main

	' Dialog definition
	Begin Dialog UserDialog 360,147,"Additional Settings for M-Solver",.DialogFunc ' %GRID:10,7,1,1
		OKButton 20,112,90,21
		CancelButton 130,112,90,21
		GroupBox 10,14,340,84,"",.GroupBox3
		CheckBox 30,35,160,14,"Use 2nd Order",.UseHigherOrderForML
		CheckBox 30,70,240,14,"Use Previous Deembedding",.UsePreviousDeembedding
	End Dialog
	Dim dlg As UserDialog

	'Restore ==================================================
	dlg.UseHigherOrderForML = False
	dlg.UsePreviousDeembedding = False

	If RestoreGlobalDataValue("UseHigherOrderForML")="1" Then
		dlg.UseHigherOrderForML=True
	End If

	If RestoreGlobalDataValue("UseVerticalDeembeddingForML") = "0" Then
		dlg.UsePreviousDeembedding=True
	End If

	If RestoreGlobalDataValue("UseVerticalDeembeddingForML") = "1" Then
		dlg.UsePreviousDeembedding=False
	End If

	' Rund dialog =============================================
	If (Dialog(dlg) = 0) Then Exit All

	' Store ===================================================
	If dlg.UseHigherOrderForML = 1 Then
		StoreGlobalDataValue("UseHigherOrderForML", "1")
	Else
		StoreGlobalDataValue("UseHigherOrderForML", "0")
	End If

	If dlg.UsePreviousDeembedding = 1 Then
		StoreGlobalDataValue("UseVerticalDeembeddingForML", "0")
	Else
		StoreGlobalDataValue("UseVerticalDeembeddingForML", "1")
	End If

End Sub

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	Select Case Action%
	    Case 1 ' Dialog box initialization
    	Case 2 ' Value changing or button pressed
    	Case 3 ' TextBox or ComboBox text changed
    	Case 4 ' Focus changed
    	Case 5 ' Idle
    	Case 6 ' Function key
    End Select

End Function
