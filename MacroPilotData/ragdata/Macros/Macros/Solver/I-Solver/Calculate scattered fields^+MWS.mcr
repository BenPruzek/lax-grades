
' ================================================================================================
' This macro enables/disables the calculation of scattered fields in the I-solver
' ================================================================================================
' Copyright 2013-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 25-Sep-2023 dta: modified dialog box and added field source as 3rd option. Enhanced nfs imprint is enabled to calculate correct scattered field with the I-Solver.
' 29-Jul-2013 dta: initial version
' ================================================================================================


Option Explicit


Sub Main

	Dim cst_ff_calc_type As String, cst_nf_calc_type As String
	Dim sCommand As String
	Dim scattered_calc As Boolean




	'If RestoreGlobalDataValue("Calc_Scattered_Farfields") = "1" Then
	'	MsgBox("Scattered farfield calculation is active")
	'Else
	'	MsgBox("Total farfield calculation is active")
	'End If


	'If RestoreGlobalDataValue("Calc_Scattered_Nearfields") = "1" Then
	'	MsgBox("Scattered nearfield calculation is active")
	'Else
	'	MsgBox("Total nearfield calculation is active")
	'End If


	Begin Dialog UserDialog 560,308,"Calculate scattered fields in the I-Solver",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,540,147,"INFO",.GroupBox4
		Text 30,28,510,112,"This script enables/disables the calculation of scattered fields in the I-Solver."+vbLf+"All the following simulations performed on this project will use these settings."+vbLf+vbLf+"These settings apply only to the following excitations:"+vbLf+vbLf+"1) Plane Wave --> The displayed farfield results are always scattered."+vbLf+"2) Farfield source"+vbLf+"3) Field source",.Text1
		GroupBox 10,161,540,112,"Currently active settings",.GroupBox1
		GroupBox 280,189,220,77,"Farfield",.GroupBox3
		GroupBox 40,189,220,77,"Nearfield",.GroupBox
		OptionGroup .Group_NF
			OptionButton 60,210,90,14,"Total Field",.Total_Field_NF
			OptionButton 60,231,120,14,"Scattered Field",.Scatt_Field_NF
		OptionGroup .Group_FF
			OptionButton 310,210,150,14,"Total Field",.Total_Field_FF
			OptionButton 310,231,150,21,"Scattered Field",.Scatt_Field_FF
		OKButton 50,280,100,21
		CancelButton 170,280,100,21
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then Exit All


cst_nf_calc_type=CStr(dlg.Group_NF)
cst_ff_calc_type=CStr(dlg.Group_FF)


scattered_calc=False

 Select Case cst_nf_calc_type
			Case 0
				StoreGlobalDataValue("Calc_Scattered_Nearfields", "0")
			Case 1
				StoreGlobalDataValue("Calc_Scattered_Nearfields", "1")

				'add command to the History
				sCommand = ""
				sCommand = sCommand + "With FDSolver " + vbLf
				sCommand = sCommand + ".UseEnhancedNFSImprint ""True""" + vbLf
				sCommand = sCommand + "End With"
				AddToHistory "enable enhanced NFS Imprint", sCommand

				scattered_calc=True
 End Select

 Select Case cst_ff_calc_type
			Case 0
				StoreGlobalDataValue("Calc_Scattered_Farfields", "0")
			Case 1
				StoreGlobalDataValue("Calc_Scattered_Farfields", "1")

				If scattered_calc=False Then

				'add command to the History
				sCommand = ""
				sCommand = sCommand + "With FDSolver " + vbLf
				sCommand = sCommand + ".UseEnhancedNFSImprint ""True""" + vbLf
				sCommand = sCommand + "End With"
				AddToHistory "enable enhanced NFS Imprint", sCommand

				End If
 End Select






End Sub

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	Select Case Action%
	    Case 1 ' Dialog box initialization
			If RestoreGlobalDataValue("Calc_Scattered_Farfields")= "1" Then
				DlgValue "Group_FF", 1
			Else
				DlgValue "Group_FF", 0
			End If

			If RestoreGlobalDataValue("Calc_Scattered_Nearfields")= "1"  Then
				DlgValue "Group_NF", 1
			Else
				DlgValue "Group_NF", 0
			End If
    	Case 2 ' Value changing or button pressed
    	Case 3 ' TextBox or ComboBox text changed
    	Case 4 ' Focus changed
    	Case 5 ' Idle
    	Case 6 ' Function key

    End Select

End Function
