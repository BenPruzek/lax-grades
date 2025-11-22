Option Explicit

' --------------------------------------------------------------------------------------------------------
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 08-Jul-2015 ube: made key handling version independent (2015, 2016, etc)
' 22-May-2015 ube: first version
' ---------------------------------------------------------------------------------------------------------

Sub Main
	Begin Dialog UserDialog 650,140,"HSPICE Settings" ' %GRID:10,7,1,1
		Text 20,14,410,14,"Specify location of hspice.exe:",.Text1
		TextBox 20,35,600,21,.hspicedir
		OKButton 20,105,90,21
		CancelButton 120,105,90,21
		Text 30,63,470,14,"Note: Full path has to be entered including hspice.exe",.Text2
		Text 60,84,450,14,"Example: c:\Synopsis\Hspice-xxx\WIN64\hspice.exe",.Text3
	End Dialog
	Dim dlg As UserDialog

	Dim myWS As Object, sRegKey As String
	Set myWS = CreateObject("WScript.Shell")

	Dim sKeyName As String
	sKeyName = "HKEY_CURRENT_USER\Software\CST AG\CST DESIGN ENVIRONMENT " + Mid(GetApplicationVersion,9,4) + "\Usersettings\HSPICEExePath"

	On Error Resume Next
	sRegKey = myWS.RegRead(sKeyName)
	dlg.hspicedir = sRegKey
	On Error GoTo 0

	If dlg.hspicedir = "" Then dlg.hspicedir = "c:\synopsys\Hspice-xxx\WIN64\hspice.exe"

	If (Dialog(dlg)<>0) Then
		myWS.RegWrite(sKeyName, dlg.hspicedir, "REG_SZ")
	End If
End Sub

' GetFilePath$("hspice.exe", , , "Browse hspice.exe")
