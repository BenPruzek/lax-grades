' ================================================================================================
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 09-Apr-2023 tlu: First version
' ================================================================================================

Dim caption, id, description As String
Sub Main ()
    id = "CheckSourceInsideDielectric"
    caption = "Check Source Inside Dielectric Object"
    description = "If True, a check if a source is inside a dielectric is performed and a warning is potentially emitted."

	Begin Dialog UserDialog 500,110,caption
		CheckBox 30,11,400,15,caption,.checkBox
		Text 30,35,450,25,description,.Text1
		OKButton 30,80,90,21
		CancelButton 140,80,90,21
	End Dialog

    Dim dlg As UserDialog
    dlg.checkBox = AsymptoticSolver.Get(id)
    If Dialog(dlg) = -1 Then	'OK button
        AsymptoticSolver.Set id, CStr(dlg.checkBox)
    End If
End Sub
