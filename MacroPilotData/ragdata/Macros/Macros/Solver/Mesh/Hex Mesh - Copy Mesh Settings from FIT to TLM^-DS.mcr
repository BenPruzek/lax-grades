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

	If (MsgBox "This macro will copy all Hexahedral Fit mesh settings to TLM. The global mesh settings and all local mesh groups." + vbCrLf + vbCrLf + _
			"Do you want to overwrite existing TLM mesh settings?",vbQuestion+vbYesNo,"Copy Mesh Settings From FIT to TLM") = vbNo Then
		Exit All
	End If

	Mesh.SyncMeshSettings "Hex", "HexTLM", "1", "1", "1", "HexahedralTLM", "High Frequency"
	Mesh.Update

	MsgBox "Hexahedral FIT mesh settings have been successfully copied to TLM and mesh type has been switched to TLM.",vbInformation+vbOkOnly,"Copy Mesh Settings From FIT to TLM"


End Sub
