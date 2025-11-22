'#Language "WWB-COM"

' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 07-Jun-2022 mba: First version
' 07-May-2024 : Changed wording
' ================================================================================================
Option Explicit

Sub Main

	If (MsgBox "This macro will copy all Hexahedral TLM mesh settings to FIT. The global mesh settings and all local mesh groups" + vbCrLf + vbCrLf + _
			"Do you want to overwrite existing FIT Mesh Settings?",vbQuestion+vbYesNo,"Copy Mesh Settings From TLM to FIT") = vbNo Then
		Exit All
	End If

	Mesh.SyncMeshSettings "HexTLM", "Hex", "1", "1", "1", "PBA", "High Frequency"
	Mesh.Update

	MsgBox "Hexahedral TLM mesh settings have been successfully copied to FIT and mesh type has been switched to FIT.",vbInformation+vbOkOnly,"Copy Mesh Settings From TLM to FIT"

	
End Sub
