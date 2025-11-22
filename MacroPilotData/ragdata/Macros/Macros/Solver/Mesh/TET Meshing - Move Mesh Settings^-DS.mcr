' move-mesh

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 14-Jun-2018 ube: First version
' ================================================================================================
Sub Main () 

	Begin Dialog UserDialog 520,126,"Move Mesh Settings" ' %GRID:10,7,1,1
		GroupBox 20,7,480,77,"",.GroupBox1
		CheckBox 40,28,440,14,"Move Mesh with careful smoothing, preserving local cell size",.MoveMeshSizePreservingSmoothing
		CheckBox 40,56,320,14,"Move Mesh without Mesh Quality improvement",.MoveMeshTryWithoutImprovements
		OKButton 40,91,90,21
		CancelButton 140,91,90,21
	End Dialog
	Dim dlg As UserDialog

	With MeshSettings
		.SetMeshType "Unstr"
		dlg.MoveMeshSizePreservingSmoothing = .Get("MoveMeshSizePreservingSmoothing")
		dlg.MoveMeshTryWithoutImprovements  = .Get("MoveMeshTryWithoutImprovements")
	End With
	' MsgBox CStr(MeshSettings.Get("MoveMeshSizePreservingSmoothing")) + vbCrLf + CStr(MeshSettings.Get("MoveMeshTryWithoutImprovements"))

	If (Dialog(dlg) = 0) Then Exit All

	Dim sCommand As String
	sCommand = ""
	sCommand = sCommand + "With MeshSettings" + vbCrLf
	sCommand = sCommand + "  .SetMeshType ""Unstr""" + vbCrLf

	If dlg.MoveMeshSizePreservingSmoothing Then
		sCommand = sCommand + "  .Set ""MoveMeshSizePreservingSmoothing"", ""true""" + vbCrLf
	Else
		sCommand = sCommand + "  .Set ""MoveMeshSizePreservingSmoothing"", ""false""" + vbCrLf
	End If

	If dlg.MoveMeshTryWithoutImprovements Then
		sCommand = sCommand + "  .Set ""MoveMeshTryWithoutImprovements"", ""true""" + vbCrLf
	Else
		sCommand = sCommand + "  .Set ""MoveMeshTryWithoutImprovements"", ""false""" + vbCrLf
	End If

	sCommand = sCommand + "End With" + vbCrLf

	' MsgBox sCommand
	AddToHistory "(*) define move mesh special settings", sCommand

End Sub
