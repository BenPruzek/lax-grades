' --------------------------------------------------------------------------------------------------------
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 07-Apr-2017 ssu: Added SolverTrace key
' 04-Nov-2016 tsi:  The Linux registry is case sensitive. I adjusted the names accordingly.
' 29-Jun-2016 ube,fwo: adjusted text in dialogue
' 20-Jun-2016 ube,fwo: added guiinfo and SolverServiceLog
' 30-Dec-2015 ube:  first version
' ---------------------------------------------------------------------------------------------------------

Option Explicit

Public Const SKeyArray = Array( _
									"DebugStart", _
									"DELog", _
									"ModelerLog", _
									"SolverServiceLog", _
									"SolverTrace", _
									"DSLog", _
									"CSTSettingsLog", _
									"GUITrace", _
									"GUIInfo", _
									"ModelerFileloadTiming", _
									"ModelerStartupTiming", _
									"ModelerFilesaveTiming", _
									"LicenseTrace")

Sub Main ()
	Begin Dialog UserDialog 630,189,"Debug Logging" ' %GRID:10,7,1,1
		GroupBox 10,7,610,140,"",.GroupBox1
		Text 20,21,590,77,"Please switch on logging for CST STUDIO SUITE using this macro when your CST support engineer asks you to do so. The information stored in the log files can help CST to find setup or compatibility issues of CST STUDIO SUITE on your system. The log files can be found in the main installation folder of CST STUDIO SUITE (all files with the file extension "".log""). Please restart the CST frontend after activating/deactivating the logging such that the setting takes effect.",.Text3
		OptionGroup .Group1
			OptionButton 30,98,320,14,"Activate debug logging",.OptionButton1
			OptionButton 30,119,310,14,"Deactivate debug logging",.OptionButton2
		OKButton 20,161,90,21
		CancelButton 120,161,90,21
	End Dialog
	Dim dlg As UserDialog
	Dim i As Integer, sfolder As String

	If (Dialog(dlg)<>0) Then

		'"HKEY_CURRENT_USER\Software\CST AG" will be added automatically

		sfolder = "CST Debug\"+Mid(GetApplicationVersion,9,4)+"\"

		Select Case dlg.Group1
		Case 0 ' activate
			For i = 0 To UBound(SKeyArray)
				RegWriteHKCU sfolder+Cstr(SKeyArray(i)), 1
			Next
		Case 1 ' de-activate
			For i = 0 To UBound(SKeyArray)
				RegDeleteHKCU sfolder+Cstr(SKeyArray(i))
			Next
		End Select

	End If
End Sub
