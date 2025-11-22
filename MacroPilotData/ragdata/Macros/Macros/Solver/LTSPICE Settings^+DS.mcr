Option Explicit

' --------------------------------------------------------------------------------------------------------
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 29-Apr-2019 cri: first version
' 03-Jun-2019 cri: Corrected wrong heading and macro text from HSPICE to LTSPICE.
' 29-Aug-2022 cri: Introduce subcircuit name separator as new setting
' ---------------------------------------------------------------------------------------------------------

Sub Main

 Dim lists$(3)
    lists$(0) = "~"
    lists$(1) = "|"
    lists$(2) = ":"
    lists$(3) = "?"

	Begin Dialog UserDialog 750,245,"LTSPICE Settings" ' %GRID:10,7,1,1
		Text 20,14,410,14,"Specify location of LTSPICE executable:",.Text1
		TextBox 20,35,700,21,.ltspicedir
		Text 40,63,470,14,"Note: Full path has to be entered including the LTSPICE executable",.Text2
		Text 80,84,450,14,"Example: C:\Program Files\LTC\LTspiceVersion\Versionx64.exe",.Text3
		Text 20,104,210,14,"Subcircuit name separator:"
      	DropListBox 220,104,50,15,lists$(),.separator
		Text 40,125,700,14,"Note: A subcircuit name separator character is used by CST Design Studio to make subcircuit names unique.",.Text4
		Text 80,146,700,14,"It needs to be a forbidden character in CST Design Studio and an allowed character in LTspice.",.Text5
		Text 80,167,700,14,"Allowed characters for subcircuit names changed between LTspice versions.", .Text6
		Text 80,188,700,14,"Therefore, the separator has been made configurable.",.Text7
		OKButton 20,213,90,21
		CancelButton 120,213,90,21
	End Dialog
	Dim dlg As UserDialog

	Dim myWS As Object, sRegKey As String, sep As String
	Set myWS = CreateObject("WScript.Shell")

	Dim sKeyName As String, sKeyName2 As String
	sKeyName = "HKEY_CURRENT_USER\Software\CST AG\CST DESIGN ENVIRONMENT " + Mid(GetApplicationVersion,9,4) + "\Usersettings\LTSPICEExePath"
	sKeyName2 = "HKEY_CURRENT_USER\Software\CST AG\CST DESIGN ENVIRONMENT " + Mid(GetApplicationVersion,9,4) + "\Usersettings\LTSPICESeparator"
	On Error Resume Next
	sRegKey = myWS.RegRead(sKeyName)
	dlg.ltspicedir = sRegKey
	sep = myWS.RegRead(sKeyName2)
	If sep = lists(1) Then
		dlg.separator = 1
    ElseIf sep = lists(2) Then
    	dlg.separator = 2
    ElseIf sep = lists(3) Then
    	dlg.separator = 3
    Else
		dlg.separator = 0
	End If

	On Error GoTo 0

	If dlg.ltspicedir = "" Then dlg.ltspicedir = "C:\Program Files\LTC\LTspiceVersion\Versionx64.exe"

	If (Dialog(dlg)<>0) Then
		myWS.RegWrite(sKeyName, dlg.ltspicedir, "REG_SZ")
		myWS.RegWrite(sKeyName2, lists(dlg.separator), "REG_SZ")
	End If
End Sub

' GetFilePath$("ltspice exe", , , "Browse LTSpice executable")
