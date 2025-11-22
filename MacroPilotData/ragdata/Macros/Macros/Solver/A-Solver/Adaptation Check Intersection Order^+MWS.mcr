
' ==============================================================================
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ==============================================================================
' History of Changes
' ------------------------------------------------------------------------------
' 06-Feb-2024 bmz: First version
' ==============================================================================

Dim caption, id, description As String
Sub Main ()
    id = "AdaptationCheckNhits"
    caption = "Set maximum intersection order considered for adaptation"
    description = "Set maximum intersection order considered for adaptation checks."

	Begin Dialog UserDialog 500,85,caption
		TextBox 30,11,40,17,.input
		Text 30,35,450,50,description,.Text1
		OKButton 30,55,90,21
		CancelButton 140,55,90,21
	End Dialog

    Dim dlg As UserDialog
    dlg.input = AsymptoticSolver.Get(id)
    If Dialog(dlg) = -1 Then	'OK button
        AddToHistory "asymptotic solver: set " + caption, "AsymptoticSolver.Set """ + id + """, " + CStr(dlg.input)
    End If
End Sub


