'#Language "WWB-COM"
'#include "vba_globals_all.lib"

Option Explicit

' This macro removes all parts that fit in a cube of specified size
'
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' -----------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 22-Sep-2016 fsr: Initial version
' -----------------------------------------------------------------------------------------------------------------------

Sub Main

	Begin Dialog UserDialog 330,84,"Remove or Exclude Small Parts",.DialogFunc ' %GRID:10,7,1,1
		Text 20,21,170,14,"Min. box size around solid:"
		TextBox 200,14,60,21,.BoxSizeT
		Text 270,21,30,14,"LengthUnits",.LengthUnitsL
		PushButton 20,49,90,21,"Remove",.RemovePB
		PushButton 120,49,90,21,"Exclude",.ExcludePB
		PushButton 220,49,90,21,"Exit",.ExitPB
	End Dialog

	Dim dlg As UserDialog
	If Dialog(dlg, 1) = 3 Then
		Exit All
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim i As Long
	Dim dBoxSize As Double, nSolids As Long, sListOfSolids() As String, nAffectedSolids As Long
	Dim sHistoryString As String, dLargestExtension As Double
	Dim dXMin As Double, dXMax As Double, dYMin As Double, dYMax As Double, dZMin As Double, dZMax As Double

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("BoxSizeT","0.01")
		DlgText("LengthUnitsL", Units.GetUnit("Length"))
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		dBoxSize = Evaluate(DlgText("BoxSizeT"))
		nSolids = Solid.GetNumberOfShapes
		ReDim sListOfSolids(nSolids - 1)
		For i = 0 To nSolids - 1
			sListOfSolids(i) = Solid.GetNameOfShapeFromIndex(i)
		Next

		sHistoryString = ""
		nAffectedSolids = 0

    	Select Case DlgItem$
			Case "RemovePB"
				If (MsgBox("This will permanently remove the specified parts from the project. Do you want to proceed?", vbYesNo, "Confirm Part Removal") = vbYes) Then
					For i = 0 To nSolids-1
						Solid.GetLooseBoundingBoxOfShape(sListOfSolids(i), dXMin, dXMax, dYMin, dYMax, dZMin, dZMax)
						dLargestExtension = Max_LIB(Array(dXMax - dXMin, dYMax - dYMin, dZMax - dZMin))
						If (dLargestExtension < dBoxSize) Then
							nAffectedSolids = nAffectedSolids + 1
							sHistoryString = sHistoryString & "Solid.Delete(" & Chr(34) & sListOfSolids(i) & Chr(34) & ")" & vbNewLine
						End If
					Next
					If (nAffectedSolids > 0) Then
						FastAddToHistory("Remove Parts Smaller Than " + CStr(dBoxSize), sHistoryString)
					End If
				End If
			Case "ExcludePB"
				For i = 0 To nSolids-1
					Solid.GetLooseBoundingBoxOfShape(sListOfSolids(i), dXMin, dXMax, dYMin, dYMax, dZMin, dZMax)
					dLargestExtension = Max_LIB(Array(dXMax - dXMin, dYMax - dYMin, dZMax - dZMin))
					If (dLargestExtension < dBoxSize) Then
						nAffectedSolids = nAffectedSolids + 1
						sHistoryString = sHistoryString & "Group.AddItem(" & Chr(34) & "solid$" & sListOfSolids(i) & Chr(34) & ", " & Chr(34) & "Excluded from Simulation" & Chr(34) & ")" & vbNewLine
						sHistoryString = sHistoryString & "Group.AddItem(" & Chr(34) & "solid$" & sListOfSolids(i) & Chr(34) & ", " & Chr(34) & "Excluded from Bounding Box" & Chr(34) & ")" & vbNewLine
					End If
				Next
				If (nAffectedSolids > 0) Then
					FastAddToHistory("Exclude Parts Smaller Than " & CStr(dBoxSize), sHistoryString)
				End If
	        Case "ExitPB"
	        	Exit All
        End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function FastAddToHistory(sEntryLabel As String, sCommand As String) As Integer

	' Disable screen updating for faster history execution
	ScreenUpdating("False")
	AddToHistory(sEntryLabel, sCommand)
	ScreenUpdating("True")

End Function
