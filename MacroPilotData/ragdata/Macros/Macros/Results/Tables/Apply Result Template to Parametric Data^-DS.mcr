'#Language "WWB-COM"

'Option Explicit
'#include "vba_globals_all.lib"

' This macro converts a template (selected by the user) to a macro and adds it to the macro tree. When executed, this macro performs the template
' functionality for all parameter/table entries of the selected input.
'
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' --------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 14-Aug-2018 ube: Msgbox to inform about new functionality in v2019
' 03-Aug-2016 fsr: Added NoForbiddenFilenameCharacters to Save command for bwc with storage option "SQL and ASCII"
' 09-Oct-2014 fsr: Fixed a bug that created incorrect reference object for 'MakeCompatibleTo'
' 20-Jun-2014 fsr: Fixed a bug that caused the template to fail if the number of run IDs was lower than the number of parameters
' 30-Jan-2014 fsr: Added GUI
' 23-Jan-2014 fsr: Tables are now supported, all code is now contained in single file
' 02-Nov-2013 fsr: Initial version
' --------------------------------------------------------------------------------------------------------------------------------------------------

Sub Main

	If (MsgBox "This macro is no longer required in conjunction with the result template ""0D or 1D Result from 1D Result "". Instead, please just define the result template in the normal way and when evaluating, you will be asked to evaluate for all existing RunIDs." + vbCrLf + vbCrLf + "Do you like to continue(=YES) or abort(=NO)?",vbYesNo,"Evaluate for existing RunIDs") = vbNo Then Exit All

	Dim sTemplateFolderList() As String, sTemplateNameList() As String

	Begin Dialog UserDialog 630,140,"Apply Result Template to Parametric Results",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 20,35,580,170,sTemplateFolderList(),.TemplateFolderDLB
		DropListBox 20,70,580,170,sTemplateNameList(),.TemplateDLB
		OKButton 410,105,90,21
		CancelButton 510,105,90,21
		Text 20,14,210,14,"Please select a result template:",.Text1
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then
		Exit All
	End If

End Sub

Sub GenerateMacroFromTemplate(sTemplateFileName As String)

	Dim sTemplateFileContents As String
	Dim nTemplateFileID As Integer

	Dim sMainFileName As String, sMainFileContents As String
	Dim nMainFileID As Integer

	Dim sScriptFileName As String
	Dim nScriptFileID As Integer

	sMainFileName = GetInstallPath+"\Library\Macros\Results\Tables\Apply Result Template to Parametric Data^-DS.mcr" ' path to this file
	' sTemplateFileName = GetFilePath("*.rtp", "*", GetInstallPath+"\Library\Result Templates", "Select a template", 0)
	sScriptFileName = GetProjectPath("Model3D")+"\"+Replace(Split(sTemplateFileName, "\")(UBound(Split(sTemplateFileName,"\"))), ".rtp", ".mcr")

	nTemplateFileID = FreeFile()
	Open sTemplateFileName For Input As nTemplateFileID
		sTemplateFileContents = Input(LOF(nTemplateFileID), nTemplateFileID)
	Close nTemplateFileID
	sTemplateFileContents = Replace(sTemplateFileContents, "StoreTemplateSetting", "StoreScriptSetting")
	sTemplateFileContents = Replace(sTemplateFileContents, "TemplateType", "Script_TemplateType")

	nMainFileID = FreeFile()
	Open sMainFileName For Input As nMainFileID
		sMainFileContents = Input(LOF(nMainFileID), nMainFileID)
		sMainFileContents = Split(sMainFileContents, "'#'#SCRIPT PATTERN START")(2) ' throw away the first part of the code
		sMainFileContents = Replace(sMainFileContents, "'#'#", "") ' remove all special comments to unlock script code
		sMainFileContents = Replace(sMainFileContents, "TEMPLATENAME", Split(Mid(sTemplateFileName, InStrRev(sTemplateFileName, "\")), ".")(0))
	Close nMainFileID
	If (InStr(sTemplateFileContents, "Evaluate0D(")=0) Then sMainFileContents = Replace(sMainFileContents, "Evaluate0D", "vBogusFunction")
	If (InStr(sTemplateFileContents, "Evaluate1D(")=0) Then sMainFileContents = Replace(sMainFileContents, "Evaluate1D", "vBogusFunction")
	If (InStr(sTemplateFileContents, "Evaluate1DComplex(")=0) Then sMainFileContents = Replace(sMainFileContents, "Evaluate1DComplex", "vBogusFunction")

	sTemplateFileContents = Replace(sTemplateFileContents, "GetLastResult_LIB(", "GetResultByRunID_LIB(GetScriptSetting("+Chr(34)+"CSTRUNID"+Chr(34)+","+Chr(34)+Chr(34)+"), ")
	sTemplateFileContents = Replace(sTemplateFileContents, "lib_rundef", "-1.23456e27")

	nScriptFileID = FreeFile()
	Open sScriptFileName For Output As nScriptFileID
		Print #nScriptFileID, sTemplateFileContents
	Close nScriptFileID
	Open sScriptFileName For Append As nScriptFileID
		Print #nScriptFileID, sMainFileContents
	Close nScriptFileID

	MsgBox("Macro generated, please start it from the macro menu.")
	'Shell("notepad " + sScriptFileName)
	'RunScript(sScriptFileName)

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
			Dim i As Long

			UpdateFolderList_LIB(GetInstallPath+"\Library\Result Templates", False, False)
			ReDim sTemplateFolderList(UBound(folderContents_LIB)-4)
			For i = 4 To UBound(folderContents_LIB()) ' first 2 entries are . and .., next 2 are obsolete 0D and 1D
				sTemplateFolderList(i-4) = Replace(Mid(folderContents_LIB(i), 3), "\", "")
			Next
			DlgListBoxArray("TemplateFolderDLB", sTemplateFolderList)
			DlgValue("TemplateFolderDLB", 0)

			' fill list with currently selected folder and add to drop list box; set value to 0
			UpdateTemplateList(sTemplateFolderList(0))
			DlgListBoxArray("TemplateDLB", sTemplateNameList)
			DlgValue("TemplateDLB", 0)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "TemplateFolderDLB"
				DialogFunc = True
				' User changed folder, update template name list
				UpdateTemplateList(sTemplateFolderList(SuppValue))
				DlgListBoxArray("TemplateDLB", sTemplateNameList)
				DlgValue("TemplateDLB", 0)
			Case "OK"
				GenerateMacroFromTemplate(sTemplateFileList(DlgValue("TemplateDLB")))
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Dim sTemplateFolderList() As String
Dim sTemplateNameList() As String
Dim sTemplateFileList() As String

Sub UpdateTemplateList(sTemplateFolder As String)
	' Returns an array containing all template files inside the provided template folder
	Dim sTempFileName As String, sApplicationName As String
	Dim iTemplateFileID As Long, sTemplateFileContents As String
	Dim i As Long

	sApplicationName = GetApplicationName()
	If (Left(sApplicationName, 2) = "DS") Then sApplicationName = "DS"

	i = 0
	ReDim sTemplateNameList(0)
	sTemplateNameList(0) = "No templates supported in this group."
	sTempFileName = Dir(GetInstallPath+"\Library\Result Templates\"+sTemplateFolder+"\*.rtp")
	While (sTempFileName <> "")
		If (((Left(sTempFileName, 1) <> "-") And (InStr(sTempFileName, "-"+sApplicationName)=0))And _
			((InStr(sTempFileName, "^")=0) Or (InStr(sTempFileName, "+"+sApplicationName)>0))) Then
			' take a quick look into template file. It has to contain 'GetLastResult_LIB', or we can discard it
			iTemplateFileID = FreeFile
			Open GetInstallPath+"\Library\Result Templates\"+sTemplateFolder+"\"+sTempFileName For Input As iTemplateFileID
				sTemplateFileContents = Input(LOF(iTemplateFileID), iTemplateFileID)
			Close iTemplateFileID
			If (InStr(sTemplateFileContents, "GetLastResult_LIB(")>0) Then
				ReDim Preserve sTemplateNameList(i)
				ReDim Preserve sTemplateFileList(i)
				sTemplateNameList(i) = Split(Split(Split(sTempFileName, "\")(UBound(Split(sTempFileName, "\"))), "^")(0), ".rtp")(0)
				sTemplateFileList(i) = GetInstallPath+"\Library\Result Templates\"+sTemplateFolder+"\"+sTempFileName
				i = i + 1
			End If
		End If
		sTempFileName = Dir()
	Wend

End Sub


'#'#SCRIPT PATTERN START
'#'#
'#'#Sub Main
'#'#
'#'#	Dim sName As String, bCreate As Boolean, bNameChanged As Boolean
'#'#	Dim d0DResult As Double, o1DResult As Object, o1DCResult As Object
'#'#	Dim sRunIDList() As String
'#'#	Dim sReferenceTreeEntry As String
'#'#	Dim i As Long, j As Long
'#'#	Dim vParameterNames As Variant, vParameterValues As Variant, oParameterPlots() As Object, nValidRunIDs As Long
'#'#
'#'#	ActivateScriptSettings(True)
'#'#	ClearScriptSettings
'#'#
'#'#	If (Define(sName, bCreate, bNameChanged)) Then
'#'#
'#'#		sReferenceTreeEntry = "1D Results"
'#'#		While (Resulttree.GetFirstChildName(sReferenceTreeEntry) <> "")
'#'#			sReferenceTreeEntry = Resulttree.GetFirstChildName(sReferenceTreeEntry)
'#'#		Wend
'#'#		sRunIDList = GetListOfRunIDs_LIB(sReferenceTreeEntry)
'#'#
'#'#		Set o1DResult = Result1D("")
'#'#		GetParameterCombination(sRunIDList(i), vParameterNames, vParameterValues)
'#'#		' Store parameter values
'#'#		nValidRunIDs = 0
'#'#		For i = 0 To UBound(sRunIDList)
'#'#			vParameterNames = Empty
'#'#			GetParameterCombination(sRunIDList(i), vParameterNames, vParameterValues)
'#'#			If (Not IsEmpty(vParameterNames)) Then
'#'#				nValidRunIDs = nValidRunIDs + 1
'#'#				If (nValidRunIDs=1) Then
'#'#					ReDim oParameterPlots(UBound(vParameterNames))
'#'#					For j = 0 To UBound(vParameterNames)
'#'#						Set oParameterPlots(j) = Result1D("")
'#'#						oParameterPlots(j).AppendXY(1, CDbl(vParameterValues(j)))
'#'#					Next
'#'#				Else
'#'#					For j = 0 To UBound(vParameterNames)
'#'#						oParameterPlots(j).AppendXY(nValidRunIDs, CDbl(vParameterValues(j)))
'#'#					Next
'#'#				End If
'#'#			End If
'#'#		Next
'#'#		If (nValidRunIDs = 0) Then
'#'#			ReportError("No RunIDs with parameter combinations found, exiting.")
'#'#		Else
'#'#			For j = 0 To UBound(vParameterNames)
'#'#				oParameterPlots(j).Save(NoForbiddenFilenameCharacters(GetProjectPath("Result")+"TemplateToMacroResultParameterValues_"+CStr(vParameterNames(j))))
'#'#				oParameterPlots(j).AddToTree("1D Results\Template to Macro\TEMPLATENAME\Parameter values\"+CStr(vParameterNames(j)))
'#'#			Next
'#'#		End If
'#'#		For i = 0 To UBound(sRunIDList)
'#'#			vParameterNames = Empty
'#'#			GetParameterCombination(sRunIDList(i), vParameterNames, vParameterValues)
'#'#			If (Not IsEmpty(vParameterNames)) Then
'#'#				StoreScriptSetting("CSTRUNID", sRunIDList(i))
'#'#				Select Case(GetScriptSetting("Script_TemplateType", ""))
'#'#					Case "0D"
'#'#						d0DResult = Evaluate0D()
'#'#						o1DResult.AppendXY(i, d0DResult)
'#'#						o1DResult.Save(NoForbiddenFilenameCharacters(GetProjectPath("Result")+"TemplateToMacroResult0D"))
'#'#						o1DResult.AddToTree("1D Results\Template to Macro\TEMPLATENAME\Result values")
'#'#					Case "1D"
'#'#						Set o1DResult = Evaluate1D()
'#'#						o1DResult.Save(NoForbiddenFilenameCharacters(GetProjectPath("Result")+"TemplateToMacroResult1D_"+sRunIDList(i)))
'#'#						o1DResult.AddToTree("1D Results\Template to Macro\TEMPLATENAME\"+sRunIDList(i))
'#'#					Case "1DC"
'#'#						Set o1DCResult = Evaluate1DComplex()
'#'#						o1DCResult.Save(NoForbiddenFilenameCharacters(GetProjectPath("Result")+"TemplateToMacroResult1DC_"+sRunIDList(i)))
'#'#						o1DCResult.AddToTree("1D Results\Template to Macro\TEMPLATENAME\"+sRunIDList(i))
'#'#					Case Else
'#'#						ReportError("Could not determine template type.")
'#'#				End Select
'#'#			End If
'#'#		Next
'#'#
'#'#	End If
'#'#
'#'#	ActivateScriptSettings(False)
'#'#
'#'#End Sub
'#'#
'#'#Function vBogusFunction() As Variant
'#'#	' an empty bogus function
'#'#End Function
