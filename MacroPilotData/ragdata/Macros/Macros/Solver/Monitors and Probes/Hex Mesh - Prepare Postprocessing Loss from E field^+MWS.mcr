' Hex Mesh - Prepare Loss from E

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 18-Jul-2018 ube: First version
' ================================================================================================
Sub Main ()

	Begin Dialog UserDialog 590,231,"Hex Mesh - Prepare Loss from E" ' %GRID:10,7,1,1
		GroupBox 20,7,550,49,"",.GroupBox1
		CheckBox 40,28,440,14,"Force calculation of conductivity matrix, even without loss monitor",.AllowPowerLossPP
		OKButton 30,196,90,21
		CancelButton 130,196,90,21
		Text 30,77,540,84,"The loss calculation requires a conductivity matrix, which is created automatically, when a loss monitor is defined. "+vbCrLf+vbCrLf+"Now it is also possible to calculate loss in the postprocessing from an E-Field monitor to save disc space and run time. For this new use case, the matrix computation has to be triggered upfront, using this setting. ",.Text1
		Text 30,168,320,14,"Note: OK forces deletion of current results.",.Text2
	End Dialog
	Dim dlg As UserDialog

	dlg.AllowPowerLossPP = Mesh.IsGenericUserFlag("AllowPowerLossPP")

	If (Dialog(dlg) = 0) Then Exit All

	' Delete Results !!!
	DeleteResults

	Dim sCommand As String
	sCommand = ""

	If dlg.AllowPowerLossPP Then
		sCommand = sCommand + "  Mesh.SetGenericUserFlag(""AllowPowerLossPP"", True)" + vbCrLf
	Else
		sCommand = sCommand + "  Mesh.SetGenericUserFlag(""AllowPowerLossPP"", False)" + vbCrLf
	End If

	' Mesh.SetGenericUserFlag("AllowPowerLossPP", True)

	' MsgBox sCommand
	AddToHistory "(*) Hex Mesh - Prepare Loss from E", sCommand

End Sub
