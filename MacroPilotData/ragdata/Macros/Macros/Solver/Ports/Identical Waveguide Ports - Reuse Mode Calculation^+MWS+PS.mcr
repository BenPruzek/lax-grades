' Port Adaptation - Use Reference

' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 26-Sep-2011 ube: First version
' ================================================================================================
Sub Main () 
	Begin Dialog UserDialog 340,224,"Line Impedance Adaptation",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 10,7,310,147,"",.GroupBox1
		Text 20,24,110,14,"Reference Port:",.Text1
		TextBox 160,21,130,21,.PortRef
		Text 20,56,260,14,"Results of Reference Port are used for ",.Text3
		OptionGroup .Group1
			OptionButton 40,77,200,14,"all other waveguide ports",.OptionButton1
			OptionButton 40,98,200,14,"following ports  (e.g. 2;4;7)",.OptionButton2
		TextBox 70,119,220,21,.selected
		OKButton 10,196,90,21
		CancelButton 110,196,90,21
		CheckBox 20,168,300,14,"Reset/Deactivate this feature for all ports",.CheckReset
	End Dialog
	Dim dlg As UserDialog
	dlg.PortRef = "1"
	dlg.selected = "2;4;6"

	If (Dialog(dlg)<>0) Then

		Dim sCommand As String

		If dlg.CheckReset = 1 Then
			sCommand = "Port.ResetReferencePort ""all"""
		Else
			If dlg.Group1 = 0 Then
				sCommand = "Port.UseReferencePort """ + dlg.PortRef+""",""all""
			Else
				sCommand = "Port.UseReferencePort """ + dlg.PortRef+""","""+dlg.selected+""""
			End If
		End If

		AddToHistory "define special time domain solver parameters", sCommand

	End If

End Sub
Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean

	If (Action%=1 Or Action%=2) Then
			' Action%=1: The dialog box is initialized
			' Action%=2: The user changes a value or presses a button

		Dim bReset As Boolean
		bReset = (DlgValue("CheckReset")=1)

		DlgEnable "PortRef", Not bReset
		DlgEnable "Group1", Not bReset
		DlgEnable "selected", ((DlgValue("Group1")=1) And Not bReset)


		If (DlgItem = "OK") Then

		    ' The user pressed the Ok button. Check the settings and display an error message if some required
		    ' fields have been left blank.

		    If (DlgText("selected") = "") Then
				MsgBox "Please enter port numbers, separated by semicolon (e.g 2;3;6;8).", vbCritical
				DialogFunction = True
									' There is an error in the settings -> Don't close the dialog box.
			End If
		End If

	End If
End Function
