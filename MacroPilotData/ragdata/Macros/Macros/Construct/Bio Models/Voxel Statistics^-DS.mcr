' Voxel Statistics

' ================================================================================================
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 25-Apr-2017 apr: First version
' ================================================================================================

Sub Main ()

	Dim logFile As String
	logFile = GetProjectPath("Result") + "VoxelModelStatistics.txt"
	HumanModel.WriteLogFile(logFile)

	If Dir(logFile) <> "" Then
		Shell "notepad " & logFile,1
	Else
		MsgBox "Error writing to " & logFile
	End If

End Sub
