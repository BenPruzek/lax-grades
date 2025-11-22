'-------------------------------------------------------------------------------------------------------
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-------------------------------------------------------------------------------------------------------
' 31-Aug-2023 ube: read current hexmeshtype first, now support Hex and HexTLM
' 25-Aug-2023 ube: first version
'-------------------------------------------------------------------------------------------------------

'#Language "WWB-COM"

Option Explicit

Sub Main
	Dim sTree As String
	sTree = GetSelectedTreeItem

	If Left(sTree,19) <> "Groups\Mesh Groups\" Then
		MsgBox "Please first select a meshgroup in the tree, before executing this macro.", vbOkOnly+vbExclamation, "Hex Mesh - extend range"
		Exit All
	End If

	Dim meshType As String
	With Mesh
        meshType = .GetHexMeshType()
	End With
	Dim title As String
	If meshType = "Hex" Then
		title = "Hex Mesh - "
	ElseIf meshType = "HexTLM" Then
		title = "Hex TLM Mesh - "
	Else
        Exit All
	End If

	Dim sNameMeshgroup As String
	sNameMeshgroup = Mid(sTree,20)

	With MeshSettings
		With .ItemMeshSettings ("group$"+sNameMeshgroup)
			.SetMeshType meshType

	Begin Dialog UserDialog 520,182,title+"Extend mesh refinement range - "+sNameMeshgroup,.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,500,140,"",.GroupBox1
		OptionGroup .Group1
			OptionButton 30,21,310,14,"Number of cells in x, y, z direction",.OptionButton1
			OptionButton 30,84,240,14,"Absolute value in x, y, z direction",.OptionButton2
		OKButton 20,154,90,21
		CancelButton 120,154,90,21
		TextBox 50,49,140,21,.nx
		TextBox 200,49,140,21,.ny
		TextBox 350,49,140,21,.nz
		TextBox 50,112,140,21,.vx
		TextBox 200,112,140,21,.vy
		TextBox 350,112,140,21,.vz
	End Dialog
			Dim dlg As UserDialog

			dlg.Group1 = IIf(.GetStr("VolumeRefinementExtentType")="ABS_VALUE",1,0)

			dlg.nx = .GetComponentStr("VolumeRefinementExtentNumSteps",0)
			dlg.ny = .GetComponentStr("VolumeRefinementExtentNumSteps",1)
			dlg.nz = .GetComponentStr("VolumeRefinementExtentNumSteps",2)

			dlg.vx = .GetComponentStr("VolumeRefinementExtentStep",0)
			dlg.vy = .GetComponentStr("VolumeRefinementExtentStep",1)
			dlg.vz = .GetComponentStr("VolumeRefinementExtentStep",2)

			If (Dialog(dlg) = 0) Then Exit All

		End With
	End With

	Dim sCommand As String
	sCommand = ""

	sCommand = sCommand + "With MeshSettings" + vbCrLf
	sCommand = sCommand + " With .ItemMeshSettings (""group$"+sNameMeshgroup+""")" + vbCrLf
	sCommand = sCommand + "  .SetMeshType """ + meshType + """" + vbCrLf
	sCommand = sCommand + "  .Set ""VolumeRefinementExtentValueUseSameXYZ"", 0" + vbCrLf

	If dlg.Group1 = 0 Then
		sCommand = sCommand + "  .Set ""VolumeRefinementExtentType"", ""STEPS_PER_DIM""" + vbCrLf
		sCommand = sCommand + "  .Set ""VolumeRefinementExtentNumSteps""," + dlg.nx + "," + dlg.ny + "," + dlg.nz + vbCrLf
	Else
		sCommand = sCommand + "  .Set ""VolumeRefinementExtentType"", ""ABS_VALUE""" + vbCrLf
		sCommand = sCommand + "  .Set ""VolumeRefinementExtentStep""," + dlg.vx + "," + dlg.vy + "," + dlg.vz + vbCrLf
	End If

	sCommand = sCommand + " End With" + vbCrLf
	sCommand = sCommand + "End With" + vbCrLf

	'MsgBox sCommand
	AddToHistory "(*) Hex Mesh - extend range "+sNameMeshgroup, sCommand

End Sub
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean

	Dim bAbsValue As Boolean

	If (Action%=1) Or (Action%=2) Then

		bAbsValue = (DlgValue("Group1")=1)

		DlgEnable "nx", Not bAbsValue
		DlgEnable "ny", Not bAbsValue
		DlgEnable "nz", Not bAbsValue

		DlgEnable "vx", bAbsValue
		DlgEnable "vy", bAbsValue
		DlgEnable "vz", bAbsValue

	End If
End Function
