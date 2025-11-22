' ================================================================================================
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 27-Jun-2007 ube: First version
' ================================================================================================

Sub Main ()

	Dim system_id As String
	Dim path As String
	Dim filename As String
	Dim location As String

	path = ""
	filename = ""
	location = ""

	If IsWindows Then
		system_id = GetSystemId
		
		If system_id = "IA32" Then
			'Windows 32bit found - start HWAccDiagnostics in MWS root folder
			path = GetInstallPath & "\"
			filename = "HWAccDiagnostics.exe"
		End If

		If system_id = "AMD64" Then
			'Windows 64bit on AMD or Intel architecture found - start HWAccDiagnostics_AMD64 in 'MWS root folder'\AMD64
			path = GetInstallPath & "\" & GetSystemId & "\"
			filename = "HWAccDiagnostics_AMD64.exe"
		End If
		
		'define location where to store the results
		location = GetProjectPath ("Temp")
		location = " -f=" & Chr(34) & Left(location,Len(location)-1) & Chr(34)

		If (path <> "") Then
			Shell(path + filename + location, 1)
		End If
	Else
		If IsArm Then
			MsgBox "Studio Suite front end is not yet available for Arm, please start HWAccDiagnostics_ARM64 from " & GetInstallPath & "\LinuxARM64\",vbExclamation
		Else
			path = GetInstallPath & "/LinuxAMD64/"
			filename = "HWAccDiagnostics_AMD64"
			location = MakeNativePath(GetProjectPath ("Temp"))
			fullpath = MakeNativePath(path + filename + " -f=" + location)
			RunCommand(fullpath)
			ReportInformation("Hardware Check done, results can be found in: " & location)
		End If
	End If

End Sub


