'#Language "WWB-COM"

'#include "vba_globals_all.lib"

' This macro exports every history entry of the currently open model after 'sIPKISSAdditionalCommandsTag' to a file. This file can then be used as "additional settings" for the CST/IPKISS link
'
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 23-Oct-2017 fsr: Additional settings are now appended, such that previous additial settings are retained
' 24-Jul-2017 fsr: Output file path is recovered from IPKISS via GlobalDataValue. Added some error checks.
' 05-Jun-2017 fsr: Initial version
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Const sIPKISSModelCreatedTag = "'@ --- End Of IPKISS Model Creation ---"
Private Const sIPKISSAdditionalCommandsTag = "'@ --- End Of CST Additional Commands ---"

Sub Main

	Begin Dialog UserDialog 230,168,"Export to IPKISS",.DlgFunction ' %GRID:10,7,1,1
		Picture 20,7,190,126,GetInstallPath() & "\Library\Macros\Results\- Import and Export\luceda_logo.bmp",0,.LucedaLogo
		PushButton 20,140,90,21, "Export"
		CancelButton 120,140,90,21
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
		Exit All
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DlgFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DlgFunction = True ' Prevent button press from closing the dialog box
		Dim sListOfSettings As String, sOutputFileName As String, nOutputFileID As Integer
		Select Case DlgItem
			Case "Export"
				sOutputFileName = ReStoreGlobalDataValue("IPKISS_Additional_Commands_Path")
				If (sOutputFileName = "") Then
					ReportError("Export to IPKISS: Export file name is undefined. Please make sure that the model was created using the latest version of Luceda IPKISS.")
				End If

				' Save project and then load mod file
				Save()
				sListOfSettings = TextFileToString_LIB(GetProjectPath("Model3D") & "\Model.mod")

				If (InStr(sListOfSettings, sIPKISSAdditionalCommandsTag) < 1) Then
					ReportError("Export to IPKISS: This model does not seem to have been created by Luceda IPKISS.")
				End If

				sListOfSettings = Mid(sListOfSettings, InStr(sListOfSettings, sIPKISSAdditionalCommandsTag))
				sListOfSettings = Replace(sListOfSettings, sIPKISSAdditionalCommandsTag, "" & vbNewLine) & vbNewLine & ""

				nOutputFileID = FreeFile()
				Open sOutputFileName For Append As nOutputFileID
					Print #nOutputFileID, sListOfSettings
				Close nOutputFileID

				MsgBox("Settings exported successfully to " & sOutputFileName & ".", "Success")
			Case "Cancel"
				DlgFunction = False
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DlgFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
