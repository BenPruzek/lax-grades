
' ================================================================================================
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 27-Oct-2023 tlu: initial version
' 16-Feb-2024 tlu: fix bug writing commands to history list
' ================================================================================================

Dim caption, id, description As String
Sub Main ()
    id_activate = "BroadBandSweepActive"
	id_range = "BroadBandSweepRange"
    caption = "Broad Band Sweep RADAR Configuration"
    description = "Maximal RADAR Range in meter"

	Begin Dialog UserDialog 480,119,caption ' %GRID:10,7,1,1
		Text 30,42,450,49,description,.Text1
		TextBox 30,63,90,14,.BroadBandSweepRange
		OKButton 30,98,90,21
		CancelButton 140,98,90,21
		CheckBox 30,7,260,14,"Activate Broad Band Sweep",.BroadBandSweepActive
	End Dialog

    Dim dlg As UserDialog
    dlg.BroadBandSweepActive = AsymptoticSolver.Get(id_activate)
    dlg.BroadBandSweepRange = AsymptoticSolver.Get(id_range)

    If Dialog(dlg) = -1 Then	'OK button
    	Dim sCommand As String
	    sCommand = "AsymptoticSolver.Set """ + id_activate + """, " + CStr(dlg.BroadBandSweepActive) + vbLf
		sCommand = sCommand + "AsymptoticSolver.Set """ + id_range + """, " + CStr(dlg.BroadBandSweepRange)
        AddToHistory "asymptotic solver: set broad band sweep", sCommand
    End If
End Sub
