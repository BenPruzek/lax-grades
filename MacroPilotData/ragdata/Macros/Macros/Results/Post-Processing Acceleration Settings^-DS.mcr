' ================================================================================================
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 10-May-2023 ube: First version
' ================================================================================================

Sub Main () 

	Begin Dialog UserDialog 520,105,"Post-Processing Acceleration Settings" ' %GRID:10,7,1,1
		GroupBox 20,7,480,56,"",.GroupBox1
		CheckBox 40,28,440,14,"Enable GPU acceleration for selected post-processing steps",.GPU_Acceleration
		OKButton 30,77,90,21
		CancelButton 130,77,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.GPU_Acceleration = IIf (LCase(GetPostprocessingAcceleration("GPU"))="on",1,0)

	If (Dialog(dlg) = 0) Then Exit All

	Dim sCommand As String
	sCommand = ""

	If dlg.GPU_Acceleration Then
		sCommand = sCommand + "  SetPostprocessingAcceleration ""GPU"",  ""ON""" + vbCrLf
	Else
		sCommand = sCommand + "  SetPostprocessingAcceleration ""GPU"",  ""OFF""" + vbCrLf
	End If

	'MsgBox sCommand
	AddToHistory "(*) PP Acceleration Settings", sCommand

End Sub
