'#Language "WWB-COM"

' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 29-Jul-2022 mba: First version
' ================================================================================================
Option Explicit

Sub Main

	If (MsgBox "This macro will apply the current Mesh Setting defaults to the project." + vbCrLf + vbCrLf + _
			"Mesh Settings that have already been created will not be changed. " + _
			"Any Mesh Settings that are created after applying this macro will use the current defaults." + vbCrLf + vbCrLf + _
			"Do you really want to continue?",vbYesNo,"Reset Settings to Current Defaults") = vbNo Then
		Exit All
	End If

	Mesh.ResetMeshSettingsDefaultsSelectionVersionToLatest
	Mesh.Update
	
	MsgBox "Mesh Setting defaults have been reset",vbInformation+vbOkOnly,""

End Sub
