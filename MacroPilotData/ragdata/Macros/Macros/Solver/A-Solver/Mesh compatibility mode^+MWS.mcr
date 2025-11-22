' ================================================================================================
' Copyright 2021-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 09-Nov-2021 bmz: First version
' ================================================================================================

Dim caption, id, description As String
Sub Main ()
    id = "MeshCompatibilityMode"
    caption = "Mesh compatibility mode"
    description = "If True, the compatibility mode will be used for facetting to obtain maximum consistency with previous versions."

	Begin Dialog UserDialog 500,110,caption
		CheckBox 30,11,400,15,caption,.checkBox
		Text 30,35,450,25,description,.Text1
		OKButton 30,80,90,21
		CancelButton 140,80,90,21
	End Dialog

    Dim dlg As UserDialog
    dlg.checkBox = AsymptoticSolver.Get(id)
    If Dialog(dlg) = -1 Then	'OK button
        AddToHistory "asymptotic solver: set " + caption, "AsymptoticSolver.Set """ + id + """, " + CStr(dlg.checkBox)
    End If
End Sub
