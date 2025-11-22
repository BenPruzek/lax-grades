'#Language "WWB-COM"
'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

Option Explicit

' Create a mesh group for a volume as defined by one or more dummy objects; macro will calculate intersections of dummy object with all objects in the volume, and include those in the mesh group

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 24-Sep-2015 fsr: Mesh refinement setting now applied correctly
' 23-Sep-2015 fsr: Updated format to 2015 version
' 07-Jan-2015 fsr: Improved GUI
' 05-Dec-2014 fsr: Initial version

Public sDummyObject As String, sTempStringArray() As String, nDummyObjects As Integer, sIntersectionObjects() As String, nIntersectionObjects As Integer, nTemp As Integer

Sub Main

	Dim i As Long

	' Preload sTargetObjects with all solids
	nIntersectionObjects = Solid.GetNumberOfShapes()
	ReDim sIntersectionObjects(nIntersectionObjects - 1)
	For i = 0 To nIntersectionObjects - 1
		sIntersectionObjects(i) = Solid.GetNameOfShapeFromIndex(i)
	Next

	Begin Dialog UserDialog 490,238,"Apply Mesh Refinement To Dummy Object",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,470,189,"",.GroupBox1
		Text 20,84,450,14,"Step 1:   Create dummy object(s) (recommended material: background)",.Text1
		Text 20,112,50,14,"Step 2:",.Text2
		PushButton 80,105,380,21,"Select dummy object(s)",.SelectDummyPB
		Text 20,21,450,56,"This macro allows a volumetric mesh refinement for TET meshes. The volume is defined by one or more dummy objects. The macro will perform all necessary boolean operations with other objects and create a new mesh group for the defined volume.",.Text3
		Text 20,140,50,14,"Step 3:",.Text4
		PushButton 80,133,380,21,"Select solids to be affected (skip = ALL)",.SelectObjectsPB
		Text 20,168,390,14,"Step 4:  Enter desired mesh step size (can be changed later)",.Text5
		TextBox 410,161,50,21,.MeshSizeT
		OKButton 280,203,90,21
		CancelButton 380,203,90,21
	End Dialog
	Dim dlg As UserDialog
	dlg.MeshSizeT = "0.1"

	Dialog dlg

End Sub

Sub ApplyMeshRefinement(dStepSize As Double)

	Dim sDummyObjectCopy As String, sListOfCopies() As String
	Dim sOriginalComponent As String, sTempComponent As String, sTempSolid As String, sTempSolidCopy As String, sTempSolidCopy2 As String
	Dim sMeshGroupName As String
	Dim i As Long
	Dim sHistoryCommand As String

	ReDim sListOfCopies(0)

	' Create mesh group
	sMeshGroupName = Replace(sDummyObject, ":", "_")
	sHistoryCommand = "Group.Add " + Chr(34) + sMeshGroupName + Chr(34) + ", " + Chr(34) + "mesh" + Chr(34)
	AddToHistory("create group: " + sMeshGroupName, sHistoryCommand)

	sHistoryCommand = "With MeshSettings" & vbNewLine
    sHistoryCommand = sHistoryCommand & "   With .ItemMeshSettings (" & Chr(34) & "group$" & sMeshGroupName & Chr(34) & ")" & vbNewLine
    sHistoryCommand = sHistoryCommand & "      .SetMeshType " & Chr(34) & "Hex" & Chr(34) & vbNewLine
    sHistoryCommand = sHistoryCommand & "      .Set " & Chr(34) & "Step" & Chr(34) & ", " & Chr(34) & CStr(dStepSize) & Chr(34) & ", " & Chr(34) & CStr(dStepSize) & Chr(34) & ", " & Chr(34) & CStr(dStepSize) & Chr(34) & vbNewLine
	sHistoryCommand = sHistoryCommand & "      .SetMeshType " & Chr(34) & "Tet" & Chr(34) & vbNewLine
    sHistoryCommand = sHistoryCommand & "      .Set " & Chr(34) & "Size" & Chr(34) & ", " & Chr(34) & CStr(dStepSize) & Chr(34) & vbNewLine
	sHistoryCommand = sHistoryCommand & "   End With" & vbNewLine
	sHistoryCommand = sHistoryCommand & "End With" & vbNewLine
	AddToHistory("set local mesh properties for: " + sMeshGroupName, sHistoryCommand)

	For i = 0 To nIntersectionObjects-1
		sTempSolid = sIntersectionObjects(i)
		sOriginalComponent = Left(sTempSolid, InStrRev(sTempSolid, ":")-1)
		sTempComponent = Replace(sTempSolid, ":", "_")
		If (Mid(GetApplicationVersion, 9, 7) < 2015) Then
			sTempSolidCopy = sTempComponent+":"+Mid(sTempSolid, InStrRev(sTempSolid, ":")+1) + "_1"
			sDummyObjectCopy = sTempComponent+":"+Mid(sDummyObject, InStrRev(sDummyObject, ":")+1) + "_1"
		Else
			' version 2015 does not add _1 when creating a copy in a new component folder, except if name already exists
			sDummyObjectCopy = sTempComponent+":"+Mid(sDummyObject, InStrRev(sDummyObject, ":")+1)
			sTempSolidCopy = sTempComponent+":"+Mid(sTempSolid, InStrRev(sTempSolid, ":")+1)
			If sTempSolidCopy = sDummyObjectCopy Then sTempSolidCopy = sTempSolidCopy + "_1"
		End If

		sHistoryCommand = "Component.New " + Chr(34) + sTempComponent + Chr(34)
		AddToHistory("new component: " + sTempComponent, sHistoryCommand)

		' Create copy of dummy object and each object
		sHistoryCommand = ""
		sHistoryCommand = sHistoryCommand & "With Transform" & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Reset" & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Name " & Chr(34) & sDummyObject & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .MultipleObjects " & Chr(34) & "True" & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Destination " & Chr(34) & sTempComponent & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Material " & Chr(34) & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Transform " & Chr(34) & "Shape" & Chr(34) & ", " & Chr(34) & "Translate" & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "End With" & vbNewLine
		AddToHistory("transform: translate " & sDummyObject, sHistoryCommand)

		sHistoryCommand = ""
		sHistoryCommand = sHistoryCommand & "With Transform" & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Reset" & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Name " & Chr(34) & sTempSolid & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .MultipleObjects " & Chr(34) & "True" & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Destination " & Chr(34) & sTempComponent & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Material " & Chr(34) & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "     .Transform " & Chr(34) & "Shape" & Chr(34) & ", " & Chr(34) & "Translate" & Chr(34) & vbNewLine
		sHistoryCommand = sHistoryCommand & "End With" & vbNewLine
		AddToHistory("transform: translate " & sTempSolid, sHistoryCommand)

		' Calculate intersection of the two copies
		sHistoryCommand = "Solid.Intersect " & Chr(34) & sTempSolidCopy & Chr(34) & ", " & Chr(34) & sDummyObjectCopy & Chr(34)
		AddToHistory("boolean intersect shapes: " & sTempSolidCopy & ", " & sDummyObjectCopy, sHistoryCommand)

		' Insert intersection into the original object and the dummy, if copy exists (i.e., if there was intersection) and move copy to original component, add to group
		If (SelectTreeItem("Components\"+Replace(sTempSolidCopy, ":", "\")) And (sTempSolid <> sDummyObject)) Then
			sHistoryCommand = "Solid.Insert " & Chr(34) & sTempSolid & Chr(34) & ", " & Chr(34) & sTempSolidCopy & Chr(34)
			AddToHistory("boolean insert shapes: " & sTempSolid & ", " & sTempSolidCopy, sHistoryCommand)

			sTempSolidCopy2 = Replace(sTempSolidCopy, sTempComponent, sOriginalComponent)
			While SelectTreeItem("Components\" & Replace(sTempSolidCopy2, ":", "\")) ' While there is an object with this name already, append _1
				sTempSolidCopy2 = sTempSolidCopy2 & "_1"
			Wend

			sHistoryCommand = "Solid.Rename " & Chr(34) & sTempSolidCopy & Chr(34) & ", " & Chr(34) & sTempSolidCopy2 & Chr(34)
			AddToHistory("rename block: " & sTempSolidCopy & " to: " & sTempSolidCopy2, sHistoryCommand)

			sHistoryCommand = "Group.AddItem " & Chr(34) & "solid$" & sTempSolidCopy2 & Chr(34) & ", " & Chr(34) & sMeshGroupName & Chr(34)
			AddToHistory("add items to group: " & Chr(34) & sMeshGroupName & Chr(34), sHistoryCommand)

			If sListOfCopies(0) <> "" Then ReDim Preserve sListOfCopies(UBound(sListOfCopies)+1)
			sListOfCopies(UBound(sListOfCopies)) = sTempSolidCopy2
		End If
		' Remove copies
		sHistoryCommand = "Component.Delete " + Chr(34) + sTempComponent + Chr(34)
		AddToHistory("delete component: " + sTempComponent, sHistoryCommand)
	Next

	' Insert copies in dummy, add dummy to mesh group
	For i = 0 To UBound(sListOfCopies)
		sHistoryCommand = "Solid.Insert " & Chr(34) & sDummyObject & Chr(34) & ", " & Chr(34) & sListOfCopies(i) & Chr(34)
		AddToHistory("boolean insert shapes: " & sDummyObject & ", " & sListOfCopies(i), sHistoryCommand)

		' If dummy does not exist anymore, it has been completely replaced by the copies
		If Not SelectTreeItem("Components\"+Replace(sDummyObject, ":", "\")) Then Exit For
	Next
	sHistoryCommand = "Group.AddItem " & Chr(34) & "solid$" & sDummyObject & Chr(34) & ", " & Chr(34) & sMeshGroupName & Chr(34)
	AddToHistory("add items to group: " & Chr(34) & sMeshGroupName & Chr(34), sHistoryCommand)

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "SelectDummyPB"
				SelectSolids_LIB(sTempStringArray, nDummyObjects)
				If (nDummyObjects <> 1) Then
					MsgBox("Please select exactly one solid.", "Error")
				Else
					sDummyObject = sTempStringArray(0)
				End If
				DialogFunc = True
			Case "SelectObjectsPB"
				SelectSolids_LIB(sIntersectionObjects, nIntersectionObjects)
				DialogFunc = True
			Case "OK"
				If (nDummyObjects = 1) And (nIntersectionObjects > 0) Then
					ApplyMeshRefinement(Evaluate(DlgText("MeshSizeT")))
					DialogFunc = False
				Else
					MsgBox("Error: This macro requires exactly 1 dummy object and N>0 target objects to be selected. Please check your settings.", "Please check your settings.")
					DialogFunc = True
				End If
			Case "Cancel"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
