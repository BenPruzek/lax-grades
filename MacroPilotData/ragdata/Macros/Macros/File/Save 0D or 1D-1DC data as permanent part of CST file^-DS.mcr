'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"

' ================================================================================================
' Macro: Stores given result file as a permanent part of the cst file.
'
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 31-Jan-2023 mha: first version
' ================================================================================================

' *** global variables
	Dim sResults()     As String  ' Result array - only contains results from the current run
	Dim sResultsDel()  As String  ' Result array - results from "Permanent storage" folder
	Dim vResultInfo    As Variant ' Additional information as specified by infoType
	Dim lNumResults    As Long    ' Number of results that can be used.
	Dim lNumResultsDel As Long    ' Number of results that can be used.

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Populate array with 0D, 1D and 1DC results:
	Dim lRunningIndex01      As Long    ' Running index...
	Dim sResultIDs()         As String  ' Array which contains result ids - userd for filtering out results
	Dim lNumResultsPre       As Long    ' Preliminary number of results that can be used.
	Dim lNumResultsPreDel    As Long  ' Number of results which can be deleted...
	Dim vTreePaths           As Variant ' Treepaths
	Dim vResultTypes         As Variant ' Type of result
	Dim vFileNames           As Variant ' Name of file
	Dim vTreePathsDel        As Variant ' Treepaths
	Dim vResultTypesDel      As Variant ' Type of result
	Dim vFileNamesDel        As Variant ' Name of file
	Dim vResultInfoDel       As Variant ' Additional information as specified by infoType
	Dim sResultDelIDs()      As String  ' Array which contains result ids - userd for filtering out results
	Dim sTitle               As String  ' Title as specified in dialog
	Dim bTitle               As Boolean ' Indicates whether or not a title has been specified by the user
	Dim sXLabel              As String  ' x-label as specified in dialog
	Dim bXLabel              As Boolean ' Indicates whether or not a label for the x-axis has been specified by the user
	Dim sYLabel              As String  ' y-label as specified in dialog
	Dim bYLabel              As Boolean ' Indicates whether or not a label for the y-axis has been specified by the user
	Dim sXUnit               As String  ' x-unit as specified in dialog
	Dim bXUnit               As Boolean ' Indicates whether or not a unit for the x-axis has been specified by the user
	Dim sYUnit               As String  ' y-unit as specified in dialog
	Dim bYUnit               As Boolean ' Indicates whether or not a unit for the y-axis has been specified by the user
	Dim bUseExFilePathName   As Boolean ' Determines whether or not original file path / name should be used within new containing folder
	Dim bAddFile             As Boolean ' True -> file is to be added, False -> file is to be deleted
	Dim sAddToHistoryContent As String
	Dim sAddToHistoryCaption As String
	Dim sSelectedResult      As String
	Dim sSelectedResultDel   As String
	Dim sFolderName          As String
	Dim sFileName            As String

	' Create array with results which can be stored
	lNumResultsPre = Resulttree.GetTreeResults("1D Results", "0D/1D recursive", "filetype0D1D", vTreePaths, vResultTypes, vFileNames, vResultInfo)
	' Filter result item entries for current run...
	If lNumResultsPre > 0 Then
		lNumResults = 0 ' Initialize
		ReDim sResults(lNumResultsPre-1)
		For lRunningIndex01 = 0 To lNumResultsPre-1 STEP 1
			sResultIDs = Resulttree.GetResultIDsFromTreeItem(vTreePaths(lRunningIndex01))
			If DoesContainCurrentRun(sResultIDs) Then
				sResults(lNumResults) = vTreePaths(lRunningIndex01)
				lNumResults           = lNumResults + 1
			End If
		Next lRunningIndex01
	Else
		lNumResults = lNumResultsPre
		ReDim sResults(0)
		sResults(0) = "No results available..."
	End If

	' Create array with results which can be deleted
	If Resulttree.DoesTreeItemExist("1D Results\Permanent storage") Then
		lNumResultsPreDel = Resulttree.GetTreeResults("1D Results\Permanent storage", "0D/1D recursive", "filetype0D1D", vTreePathsDel, vResultTypesDel, vFileNamesDel, vResultInfoDel)
		' Filter result item entries for current run (shouldn't be necessary, but why not)...
		If lNumResultsPreDel > 0 Then
			lNumResultsDel = 0 ' Initialize
			ReDim sResultsDel(lNumResultsPreDel-1)
			For lRunningIndex01 = 0 To lNumResultsPreDel-1 STEP 1
				sResultDelIDs = Resulttree.GetResultIDsFromTreeItem(vTreePathsDel(lRunningIndex01))
				If DoesContainCurrentRun(sResultDelIDs) Then
					sResultsDel(lNumResultsDel) = vTreePathsDel(lRunningIndex01)
					lNumResultsDel              = lNumResultsDel + 1
				End If
			Next lRunningIndex01
		Else
			lNumResultsDel = lNumResultsPreDel
			ReDim sResultsDel(0)
			sResultsDel(0) = "No results available (for deletion)..."
		End If
	Else
		lNumResultsDel = 0
		ReDim sResultsDel(0)
		sResultsDel(0) = "No results available (for deletion)..."
	End If

	' Initialize default settings
    StoreScriptSetting("SSFolderName",  "First subfolder")
    StoreScriptSetting("SSFileName",    "My filename")

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, sResultsDel, sSelectedResultDel, lNumResultsDel, sResults, sSelectedResult, lNumResults, sFolderName, sFileName, bUseExFilePathName, bAddFile, sTitle, bTitle, sXLabel, bXLabel, sYLabel, bYLabel, sXUnit, bXUnit, sYUnit, bYUnit) ) Then
		If bAddFile Then
			If lNumResults > 0 Then
				Dim sFilePath As String
				If CreateHistoryListEntryForAddingFile(sAddToHistoryContent, sAddToHistoryCaption, sSelectedResult, sFolderName, sFileName, bUseExFilePathName, sTitle, bTitle, sXLabel, bXLabel, sYLabel, bYLabel, sXUnit, bXUnit, sYUnit, bYUnit, sFilePath) Then
					If AddToHistory (sAddToHistoryCaption, sAddToHistoryContent) Then
						ReportInformationToWindow("The result item has been successfully added as a permanent part of the CST Studio Suite file as: " & sFilePath & ".")
					Else
						ReportWarningToWindow("Unfortunately the result item could not be successfully added as a permanent part of the CST Studio Suite file.")
					End If
				End If
			End If
		Else
			CreateHistoryListEntryForDeletingFile(sAddToHistoryContent, sAddToHistoryCaption, sSelectedResultDel)
			AddToHistory(sAddToHistoryCaption, sAddToHistoryContent)
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, sResultsDel() As String, sSelectedResultDel As String, lNumResultsDel As Long, sResults() As String, sSelectedResult As String, lNumResults As Long, sFolderName As String, sFileName As String, bUseExFilePathName As Boolean, bAddFile As Boolean, sTitle As String, bTitle As Boolean, sXLabel As String, bXLabel As Boolean, sYLabel As String, bYLabel As Boolean, sXUnit As String, bXUnit As Boolean, sYUnit As String, bYUnit As Boolean) As Boolean
	If lNumResults = 0 Then
		ReDim sResults(0)
		sResults(0)   = "No corresponding result item has been found."
	End If

	Begin Dialog UserDialog 690, 402, "Embed 0D / 1D / 1DC data as part of CST file or delete correspondingly created file...", .DialogFunc ' %GRID:3,3,1,1
		GroupBox       9,   6, 671, 105, "Please take note",                           .GBPleaseTakeNote
		Text          20,  25, 651,  14, "Only result items in the ""1D Results"" with a ""current run"" are considered for permanent storage... These", .TExplanation1
		Text          20,  40, 651,  14, "files will be stored in ""1D Results\Permanent storage\..."" Further subfolders and resultname can be",        .TExplanation2
		Text          20,  55, 546,  14, "specified in the macro. To create subfolders, please use ""\"".",                                              .TExplanation3
		Text          20,  75, 630,  14, "Only results in ""1D Results\Permanent storage\..."" can be deleted with this macro. Deleting via the",        .TExplanation4
		Text          20,  90, 630,  14, "context menu will not be sufficient.",                                                                         .TExplanation5

		GroupBox       9, 116, 671, 253, "Select result item and associated settings",  .GBSelectFile
		OptionGroup                                                                     .OGDeleteOrSave
		OptionButton  20, 135, 268,  14, "Permanently embed data file in CST file",     .OG0Save
		OptionButton 319, 135, 142,  14, "Delete selected file",                        .OG1Delete

		DropListBox   20, 159, 650,  21, sResults(),                                    .DLBSelectResultFile
		OptionGroup                                                                     .OGOverwriteOrNot
		OptionButton  20, 189, 284,  14, "Use existing folder structure and file name", .OG0Overwrite
		OptionButton 319, 189, 234,  14, "Specify folder name and file name",           .OG1SpecifyFolderFileName
		Text          20, 215, 237,  14, "Specify folder name for storage of file:",    .TSpecifyFolderName
		TextBox      267, 212, 403,  21,                                                .TBFolderName
		Text          20, 245, 220,  14, "Specify file name for storage of file:",      .TSpecifyFileName
		TextBox      267, 242, 403,  21,                                                .TBFileName

		CheckBox      20, 277, 117,  21, "Specify x-label:",                            .CBSpecifyXLabel
		TextBox      162, 277, 150,  21,                                                .TBXLabel
		CheckBox     379, 277, 116,  21, "Specify x units:",                            .CBSpecifyXUnits
		TextBox      520, 277, 150,  21,                                                .TBXUnits
		CheckBox      20, 307, 117,  21, "Specify y-label:",                            .CBSpecifyYLabel
		TextBox      162, 307, 150,  21,                                                .TBYLabel
		CheckBox     379, 307, 116,  21, "Specify y units:",                            .CBSpecifyYUnits
		TextBox      520, 307, 150,  21,                                                .TBYUnits
		CheckBox      20, 337,  97,  21, "Specify title:",                              .CBSpecifyTitle
		TextBox      162, 337, 150,  21,                                                .TBTitle

		' OK and Cancel buttons
		OKButton      15, 375,  90,  21
		CancelButton 115, 375,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		If dlg.OGDeleteOrSave = 0 Then
			' Add result
			sSelectedResult    = sResults(dlg.DLBSelectResultFile)
			sFolderName        = dlg.TBFolderName
			sFileName          = dlg.TBFileName
			bUseExFilePathName = Not CBool(dlg.OGOverwriteOrNot)
			bAddFile           = Not CBool(dlg.OGDeleteOrSave)
			sTitle             = dlg.TBTitle
			bTitle             = CBool(dlg.CBSpecifyTitle)
			sXLabel            = dlg.TBXLabel
			bXLabel            = CBool(dlg.CBSpecifyXLabel)
			sYLabel            = dlg.TBYLabel
			bYLabel            = CBool(dlg.CBSpecifyYLabel)
			sXUnit             = dlg.TBXUnits
			bXUnit             = CBool(dlg.CBSpecifyXUnits)
			sYUnit             = dlg.TBYUnits
			bYUnit             = CBool(dlg.CBSpecifyYUnits)
		Else
			sSelectedResultDel = sResultsDel(dlg.DLBSelectResultFile)
		End If
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunc: This function defines the dialog box behaviour. It is automatically called
'             whenever the user changes some settings in the dialog box, presses any button
'             or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim vResultIDs           As Long
	Dim sNoForbiddenChars    As String
	Dim sAddToHistoryContent As String
	Dim sAddToHistoryCaption As String
	Dim sSelectedResult      As String

	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			If lNumResults = 0 Then
				DlgEnable "DLBSelectResultFile", False
				DlgEnable "OGOverwriteOrNot",    False
				DlgEnable "TBFolderName",        False
				DlgEnable "TBFileName",          False
				DlgText   "TBFolderName",        GetScriptSetting("SSFolderName", "First subfolder")
				DlgText   "TBFileName",          GetScriptSetting("SSFileName",   "My filename")
				DlgEnable "CBSpecifyXLabel",     False
				DlgEnable "TBXLabel",            False
				DlgEnable "CBSpecifyXUnits",     False
				DlgEnable "TBXUnits",            False
				DlgEnable "CBSpecifyYLabel",     False
				DlgEnable "TBYLabel",            False
				DlgEnable "CBSpecifyYUnits",     False
				DlgEnable "TBYUnits",            False
				DlgEnable "CBSpecifyTitle",      False
				DlgEnable "TBTitle",             False
			Else
				DlgEnable "DLBSelectResultFile", True
				DlgEnable "OGOverwriteOrNot",    True
				DlgEnable "TBFolderName",        False
				DlgEnable "TBFileName",          False
				DlgText   "TBFolderName",        GetScriptSetting("SSFolderName", "First subfolder")
				DlgText   "TBFileName",          GetScriptSetting("SSFileName",   "My filename")
				If InStr(LCase(CStr(vResultInfo(0))), LCase("0D")) Then
					DlgEnable "CBSpecifyXLabel", False
					DlgEnable "TBXLabel",        False
					DlgEnable "CBSpecifyXUnits", False
					DlgEnable "TBXUnits",        False
				Else
					DlgEnable "CBSpecifyXLabel", True
					DlgEnable "TBXLabel",        False
					DlgEnable "CBSpecifyXUnits", True
					DlgEnable "TBXUnits",        False
				End If
				DlgEnable "CBSpecifyYLabel",     True
				DlgEnable "TBYLabel",            False
				DlgEnable "CBSpecifyYUnits",     True
				DlgEnable "TBYUnits",            False
				DlgEnable "CBSpecifyTitle",      True
				DlgEnable "TBTitle",             False
			End If

		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "OGDeleteOrSave" ) Then
				If ( lSuppValue = 0 ) Then
					DlgEnable "OGOverwriteOrNot",   True
					DlgEnable "TSpecifyFolderName", True
					DlgEnable "TSpecifyFileName",   True
					If DlgValue("OGOverwriteOrNot") = 0 Then
						DlgEnable "TBFolderName",   False
						DlgEnable "TBFileName",     False
					Else
						DlgEnable "TBFolderName",   True
						DlgEnable "TBFileName",     True
					End If
					DlgEnable "CBSpecifyXLabel",    True
					DlgEnable "TBXLabel",           IIf(DlgValue("CBSpecifyXLabel") <> 0, True, False)
					DlgEnable "CBSpecifyXUnits",    True
					DlgEnable "TBXUnits",           IIf(DlgValue("CBSpecifyXUnits") <> 0, True, False)
					DlgEnable "CBSpecifyYLabel",    True
					DlgEnable "TBYLabel",           IIf(DlgValue("CBSpecifyYLabel") <> 0, True, False)
					DlgEnable "CBSpecifyYUnits",    True
					DlgEnable "TBYUnits",           IIf(DlgValue("CBSpecifyYUnits") <> 0, True, False)
					DlgEnable "CBSpecifyTitle",     True
					DlgEnable "TBTitle",            IIf(DlgValue("CBSpecifyTitle") <> 0, True, False)
					' Update array in dropdown list...
					DlgListBoxArray "DLBSelectResultFile", sResults()
					If lNumResults = 0 Then
						DlgEnable "DLBSelectResultFile", False
						DlgValue  "DLBSelectResultFile", 0
					Else
						DlgEnable "DLBSelectResultFile", True
						DlgValue  "DLBSelectResultFile", 0
					End If
				Else
					DlgEnable "OGOverwriteOrNot",   False
					DlgEnable "TSpecifyFolderName", False
					DlgEnable "TBFolderName",       False
					DlgEnable "TSpecifyFileName",   False
					DlgEnable "TBFileName",         False
					DlgEnable "CBSpecifyXLabel",    False
					DlgEnable "TBXLabel",           False
					DlgEnable "CBSpecifyXUnits",    False
					DlgEnable "TBXUnits",           False
					DlgEnable "CBSpecifyYLabel",    False
					DlgEnable "TBYLabel",           False
					DlgEnable "CBSpecifyYUnits",    False
					DlgEnable "TBYUnits",           False
					DlgEnable "CBSpecifyTitle",     False
					DlgEnable "TBTitle",            False
					' Update array in dropdown list...
					DlgListBoxArray "DLBSelectResultFile", sResultsDel()
					If lNumResultsDel = 0 Then
						DlgEnable "DLBSelectResultFile", False
						DlgValue  "DLBSelectResultFile", 0
					Else
						DlgEnable "DLBSelectResultFile", True
						DlgValue  "DLBSelectResultFile", 0
					End If
				End If

			ElseIf ( sDlgItem = "OGOverwriteOrNot" ) Then
				If ( lSuppValue = 0 ) Then
					DlgEnable "TBFolderName", False
					DlgEnable "TBFileName",   False
				Else
					DlgEnable "TBFolderName", True
					DlgEnable "TBFileName",   True
				End If

			ElseIf ( sDlgItem = "CBSpecifyXLabel" ) Then
				DlgEnable "TBXLabel", IIf( ( lSuppValue = 0 ), False, True)

			ElseIf ( sDlgItem = "CBSpecifyXUnits" ) Then
				DlgEnable "TBXUnits", IIf( ( lSuppValue = 0 ), False, True)

			ElseIf ( sDlgItem = "CBSpecifyYLabel" ) Then
				DlgEnable "TBYLabel", IIf( ( lSuppValue = 0 ), False, True)

			ElseIf ( sDlgItem = "CBSpecifyYUnits" ) Then
				DlgEnable "TBYUnits", IIf( ( lSuppValue = 0 ), False, True)

			ElseIf ( sDlgItem = "CBSpecifyTitle" ) Then
				DlgEnable "TBTitle", IIf( ( lSuppValue = 0 ), False, True)

			ElseIf ( sDlgItem = "DLBSelectResultFile" ) Then
				If DlgValue("OGDeleteOrSave") = 0 Then
					If InStr(LCase(CStr(vResultInfo(DlgValue("DLBSelectResultFile")))), LCase("0D")) Then
						DlgEnable "CBSpecifyXLabel", False
						DlgEnable "TBXLabel",        False
						DlgEnable "CBSpecifyXUnits", False
						DlgEnable "TBXUnits",        False
					Else
						DlgEnable "CBSpecifyXLabel", True
						DlgEnable "TBXLabel",        IIf(DlgValue("CBSpecifyXLabel")<>0, True, False)
						DlgEnable "CBSpecifyXUnits", True
						DlgEnable "TBXUnits",        IIf(DlgValue("CBSpecifyXUnits")<>0, True, False)
					End If
				End If
			End If

		Case 3 ' TextBox or ComboBox text changed
			Dim sValueSet As String
			If ( sDlgItem = "TBFileName" ) Then
				sValueSet = DlgText(sDlgItem)
				If sValueSet = "" Then
					DlgText sDlgItem, GetScriptSetting("SSFileName",  "My filename")
					MsgBox "Please be aware that you seem to have entered an empty string for the file name. As such, the entry has been replaced with the original default value.", vbExclamation
				Else
					' Remove any forbidden characters and eliminate leading "."
					sNoForbiddenChars = NoForbiddenFilenameCharacters(sValueSet)
					sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
					If sNoForbiddenChars = "" Then
						DlgText sDlgItem, GetScriptSetting("SSFileName",  "My filename")
						MsgBox "Please be aware that the file name you have entered consisted only of characters not allowed to be used for filenames or resulted in leading periods. As such, the entry has been replaced with the original default value.", vbExclamation
					ElseIf sNoForbiddenChars <> sValueSet Then
						DlgText sDlgItem, sNoForbiddenChars
						MsgBox "Please be aware that the file name you have entered contained some characters not allowed to be used for filenames. As such, the entry has been slightly altered.", vbExclamation
					End If
				End If

			ElseIf ( sDlgItem = "TBFolderName" ) Then
				sValueSet = DlgText(sDlgItem)
				If sValueSet = "" Then
					DlgText sDlgItem, GetScriptSetting("SSFolderName",  "First subfolder")
					MsgBox "Please be aware that you seem to have entered an empty string for the folder name. As such, the entry has been replaced with the original default value.", vbExclamation
				Else
					' Remove any forbidden characters and eliminate leading "."
					sNoForbiddenChars = NoForbiddenFilenameCharactersExceptBackslashInMiddle(sValueSet)
					sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
					If sNoForbiddenChars = "" Then
						DlgText sDlgItem, GetScriptSetting("SSFolderName",  "First subfolder")
						MsgBox "Please be aware that the folder name you have entered consisted only of characters not allowed to be used for filenames or resulted in leading periods. As such, the entry has been replaced with the original default value.", vbExclamation
					ElseIf sNoForbiddenChars <> sValueSet Then
						DlgText sDlgItem, sNoForbiddenChars
						MsgBox "Please be aware that the folder name you have entered contained some characters not allowed to be used for filenames. As such, the entry has been slightly altered.", vbExclamation
					End If
				End If

			ElseIf ( sDlgItem = "TBXLabel" Or sDlgItem = "TBYLabel"  Or sDlgItem = "TBTitle" ) Then
				sValueSet = DlgText(sDlgItem)
				If sValueSet <> "" Then
					' Remove any forbidden characters and eliminate leading "."
					sNoForbiddenChars = NoForbiddenFilenameCharacters(sValueSet)
					sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
					If sNoForbiddenChars <> sValueSet Then
						DlgText sDlgItem, sNoForbiddenChars
						MsgBox "Please be aware that the label or title you have entered contained forbidden characters and has been changed accordingly.", vbExclamation
					End If
				End If
			End If

		Case 4 ' Focus changed
			' Nothing to do in this case...

		Case 5 ' Idle
			' Nothing to do in this case...
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions

		Case 6 ' Function key
	End Select
End Function



' Returns true if array contains entry GetLastResultID(), otherwise returns false
Private Function DoesContainCurrentRun(sCurrentRunIDs() As String) As Boolean
	Dim bIsContained
	Dim lRunningIndex As Long

	bIsContained = False

	For lRunningIndex = 0 To UBound(sCurrentRunIDs) STEP 1
		If sCurrentRunIDs(lRunningIndex) = GetLastResultID() Then
			bIsContained = True
			Exit For
		End If
	Next
	DoesContainCurrentRun = bIsContained
End Function



' Removes leading periods in string
Private Function RemoveLeadingPeriods(sMyString As String) As String
	Dim sRemainingStr As String

	If Left(sMyString, 1) = "." Then
		sRemainingStr = Right(sMyString, Len(sMyString)-1)
		While Left(sRemainingStr, 1) = "."
			sRemainingStr = Right(sRemainingStr, Len(sRemainingStr)-1)
		Wend
	Else
		sRemainingStr = sMyString
	End If

	RemoveLeadingPeriods = sRemainingStr
End Function



' Construct string for history list entry
' sAddToHistoryContent: String containing history list entry
' sAddToHistoryCaption: String containing caption of history list entry
' sSelectedResult:      Result which was selected
' sFolderName:          Name of folder in 1D Results where file should be saved to (if bUseExFilePathName is false)
' sFileName:            Name of file to be used for saving it (if bUseExFilePathName is false)
' bUseExFilePathName:   Indicates whether original file path should be used (inside new containing folder)
' sTitleSpec:           Title as specified in dialog
' bTitle:               Indicates whether or not a title has been specified by the user
' sXLabelSpec:          x-label as specified in dialog
' bXLabel:              Indicates whether or not a label for the x-axis has been specified by the user
' sYLabelSpec:          y-label as specified in dialog
' bYLabel:              Indicates whether or not a label for the y-axis has been specified by the user
' sXUnitSpec:           x-unit as specified in dialog
' bXUnit:               Indicates whether or not a unit for the x-axis has been specified by the user
' sYUnitSpec:           y-unit as specified in dialog
' bYUnit:               Indicates whether or not a unit for the y-axis has been specified by the user
' sFilePathUsed:        Full path of file which is to be saved...
Private Function CreateHistoryListEntryForAddingFile(sAddToHistoryContent As String, sAddToHistoryCaption As String, sSelectedResult As String, sFolderName As String, sFileName As String, bUseExFilePathName As Boolean, sTitleSpec As String, bTitle As Boolean, sXLabelSpec As String, bXLabel As Boolean, sYLabelSpec As String, bYLabel As Boolean, sXUnitSpec As String, bXUnit As Boolean, sYUnitSpec As String, bYUnit As Boolean, sFilePathUsed As String) As Boolean
	Dim sFilePath     As String ' path for saving of file (in tree)
	Dim sFullFileName As String ' filename for saving file before adding to the tree
	Dim oResult       As Object ' original result object...
	Dim lLength       As Long   ' number of entries in result object
	Dim sXLabel       As String ' x-label of new object
	Dim sXUnit        As String ' unit of x-axis of new object
	Dim sYLabel       As String ' y-label of new object
	Dim sYUnit        As String ' unit of y-axis of new object
	Dim sTitle        As String ' title of new object
	Dim dDataRe       As Double ' data in object (real)
	Dim dDataIm       As Double ' data in object (imaginary)
	Dim dLogFac       As Double ' only concerns Result1DComplex
	Dim sPlotView     As String ' only concerns Result1DComplex
	Dim lRunningIndex01 As Long ' running index

	If Not bUseExFilePathName Then
		sFilePath     = "1D Results\Permanent storage\" & sFolderName & "\" & sFileName
		sFullFileName = Replace(Replace(sFilePath, "/", "-"), "\", "-")
	Else
		sFilePath     = sSelectedResult
		sFileName     = Right(sSelectedResult, Len(sSelectedResult) - InStrRev(sSelectedResult, "\"))
		sFilePath     = Right(sFilePath, Len(sFilePath) - Len("1D Results\"))
		sFilePath     = "1D Results\Permanent storage\" & sFilePath
		sFullFileName = Replace(Replace(sFilePath, "/", "-"), "\", "-")
	End If
	sFilePathUsed = sFullFileName

	' Initialize
	CreateHistoryListEntryForAddingFile = True

	' Attempt to load result item
	On Error Resume Next
		Set oResult = Resulttree.GetResultFromTreeItem(sSelectedResult, GetLastResultID())
	If Err.Number <> 0 Then
		ReportWarningToWindow("Unfortunately the result item in question could not be saved succesfully.")
		CreateHistoryListEntryForAddingFile = False
	ElseIf oResult Is Nothing Then
		ReportWarningToWindow("Unfortunately the result item in question seems to be empty and could not be saved succesfully.")
		CreateHistoryListEntryForAddingFile = False
	Else
		If LCase(oResult.GetResultObjectType) = LCase("0D") Then
			' Resultobject is of type 0D
			If oResult.GetN() = 1 Then
				' Gather necessary data
				sTitle = oResult.GetTitle
				sXUnit = oResult.GetDataUnit
				If bTitle  Then sTitle  = sTitleSpec
				If bXUnit  Then sXUnit  = sXUnitSpec
				If bXLabel Then sXLabel = sXLabelSpec
				oResult.GetData(dDataRe)

				' Set up history list string
				sAddToHistoryContent = "" & vbCrLf & _
				"Dim oMyResult0DObject As Object" & vbCrLf & _
				"Set oMyResult0DObject = Result0D("""")" & vbCrLf
				sAddToHistoryContent = sAddToHistoryContent & _
				"oMyResult0DObject.SetData(" & CStr(dDataRe) & ")" & vbCrLf & _
				"oMyResult0DObject.Title(""" & sTitle & """)" & vbCrLf & _
				"oMyResult0DObject.SetDataLabelAndUnit(""" & sXLabelSpec & """, """ & sXUnit & """)" & vbCrLf & _
				"oMyResult0DObject.SetFileName(""" & sFullFileName & """)" & vbCrLf & _
				"oMyResult0DObject.Save()" & vbCrLf & _
				"oMyResult0DObject.AddToTree(""" & sFilePath & """)" & vbCrLf

				' Set up caption string
				sAddToHistoryCaption = "" & vbCrLf & _
				"Add 0D result: " & sFileName
			Else
				ReportWarningToWindow("Unfortunately the result item in question seems to be empty and could not be saved succesfully.")
				CreateHistoryListEntryForAddingFile = False
			End If
		ElseIf LCase(oResult.GetResultObjectType) = LCase("0DC") Then
			' Resultobject is of type 0DC
			If oResult.GetN() = 1 Then
				' Gather necessary data
				sTitle = oResult.GetTitle
				sXUnit = oResult.GetDataUnit
				If bTitle Then sTitle = sTitleSpec
				If bXUnit Then sXUnit = sXUnitSpec
				If bXLabel Then sXLabel = sXLabelSpec
				oResult.GetDataComplex(dDataRe, dDataIm)

				' Set up history list string
				sAddToHistoryContent = "" & vbCrLf & _
				"Dim oMyResult0DCObject As Object" & vbCrLf & _
				"Set oMyResult0DCObject = Result0D("""")" & vbCrLf
				sAddToHistoryContent = sAddToHistoryContent & _
				"oMyResult0DCObject.SetDataComplex(" & CStr(dDataRe) & ", " & CStr(dDataIm) & ")" & vbCrLf & _
				"oMyResult0DCObject.Title(""" & sTitle & """)" & vbCrLf & _
				"oMyResult0DObject.SetDataLabelAndUnit(""" & sXLabelSpec & """, """ & sXUnit & """)" & vbCrLf & _
				"oMyResult0DCObject.SetFileName(""" & sFullFileName & """)" & vbCrLf & _
				"oMyResult0DCObject.Save()" & vbCrLf & _
				"oMyResult0DCObject.AddToTree(""" & sFilePath & """)" & vbCrLf

				' Set up caption string
				sAddToHistoryCaption = "" & vbCrLf & _
				"Add 0D result: " & sFileName
			Else
				ReportWarningToWindow("Unfortunately the result item in question seems to be empty and could not be saved succesfully.")
				CreateHistoryListEntryForAddingFile = False
			End If
		ElseIf LCase(oResult.GetResultObjectType) = LCase("1D") Then
			' Resultobject is of type 1D
			If oResult.GetN() > 0 Then
				' Gather necessary data
				lLength = oResult.GetN()
				sTitle  = oResult.GetTitle
				oResult.GetXLabelAndUnit(sXLabel, sXUnit)
				oResult.GetYLabelAndUnit(sYLabel, sYUnit)
				If bTitle  Then sTitle  = sTitleSpec
				If bXUnit  Then sXUnit  = sXUnitSpec
				If bXLabel Then sXLabel = sXLabelSpec
				If bYUnit  Then sYUnit  = sYUnitSpec
				If bYLabel Then sYLabel = sYLabelSpec

				sAddToHistoryContent = "" & vbCrLf & _
				"' Set up arrays containing the values..." & vbCrLf & _
				"Dim dXValues(" & CStr(lLength-1) & ") As Double" & vbCrLf & _
				"Dim dYValues(" & CStr(lLength-1) & ") As Double" & vbCrLf & _
				"Dim lRunningIndex01 As Long" & vbCrLf

				For lRunningIndex01 = 0 To lLength-1 STEP 1
					sAddToHistoryContent = sAddToHistoryContent & "    dXValues(" & CStr(lRunningIndex01) & ") = " & CStr(oResult.GetX(lRunningIndex01)) & vbCrLf
					sAddToHistoryContent = sAddToHistoryContent & "    dYValues(" & CStr(lRunningIndex01) & ") = " & CStr(oResult.GetY(lRunningIndex01)) & vbCrLf
				Next lRunningIndex01

				sAddToHistoryContent = sAddToHistoryContent & vbCrLf & _
				"' Define result object" & vbCrLf & _
				"Dim oMyResult1DObject As Object" & vbCrLf & _
				"Set oMyResult1DObject = Result1D("""")" & vbCrLf & _
				"oMyResult1DObject.Initialize(" & CStr(lLength) & ")" & vbCrLf & _
				"oMyResult1DObject.SetArray(dXValues, ""x"")" & vbCrLf & _
				"oMyResult1DObject.SetArray(dYValues, ""y"")" & vbCrLf & _
				"oMyResult1DObject.Title(""" & sTitle & """)" & vbCrLf & _
				"oMyResult1DObject.SetXLabelAndUnit(""" & sXLabel & """, """ & sXUnit & """)" & vbCrLf & _
				"oMyResult1DObject.SetYLabelAndUnit(""" & sYLabel & """, """ & sYUnit & """)" & vbCrLf & _
				"oMyResult1DObject.Save(""" & sFullFileName & """)" & vbCrLf & _
				"oMyResult1DObject.AddToTree(""" & sFilePath & """)"

				' Set up caption string
				sAddToHistoryCaption = "" & vbCrLf & _
				"Add 1D result: " & sFileName
			Else
				ReportWarningToWindow("Unfortunately the result item in question seems to be empty and could not be saved succesfully.")
				CreateHistoryListEntryForAddingFile = False
			End If
		ElseIf LCase(oResult.GetResultObjectType) = LCase("1DC") Then
			' Resultobject is of type 1DC
			If oResult.GetN() > 0 Then
				' Gather necessary data
				lLength = oResult.GetN()
				sTitle  = oResult.GetTitle
				oResult.GetXLabelAndUnit(sXLabel, sXUnit)
				oResult.GetYLabelAndUnit(sYLabel, sYUnit)
				dLogFac = oResult.GetLogarithmicFactor
				If Not ( dLogFac = 10 Or dLogFac = 20 Or dLogFac = -1 ) Then
					dLogFac = -1
				End If
				sPlotView = oResult.GetDefaultPlotView
				If Not ( sPlotView = "real" Or sPlotView = "imaginary" Or sPlotView = "magnitude" Or sPlotView = "magnitudedb" Or sPlotView = "phase" Or sPlotView = "polar" ) Then
					sPlotView = ""
				End If
				If bTitle  Then sTitle  = sTitleSpec
				If bXUnit  Then sXUnit  = sXUnitSpec
				If bXLabel Then sXLabel = sXLabelSpec
				If bYUnit  Then sYUnit  = sYUnitSpec
				If bYLabel Then sYLabel = sYLabelSpec

				sAddToHistoryContent = "" & vbCrLf & _
				"' Set up arrays containing the values..." & vbCrLf & _
				"Dim dXValues(" & CStr(lLength-1) & ") As Double" & vbCrLf & _
				"Dim dYReValues(" & CStr(lLength-1) & ") As Double" & vbCrLf & _
				"Dim dYImValues(" & CStr(lLength-1) & ") As Double" & vbCrLf & _
				"Dim lRunningIndex01 As Long" & vbCrLf

				For lRunningIndex01 = 0 To lLength-1 STEP 1
				sAddToHistoryContent = sAddToHistoryContent & "    dXValues(" & CStr(lRunningIndex01) & ") = " & CStr(oResult.GetX(lRunningIndex01)) & vbCrLf
				sAddToHistoryContent = sAddToHistoryContent & "    dYReValues(" & CStr(lRunningIndex01) & ") = " & CStr(oResult.GetYRe(lRunningIndex01)) & vbCrLf
				sAddToHistoryContent = sAddToHistoryContent & "    dYImValues(" & CStr(lRunningIndex01) & ") = " & CStr(oResult.GetYIm(lRunningIndex01)) & vbCrLf
				Next lRunningIndex01

				sAddToHistoryContent = sAddToHistoryContent & _
				"' Define result object" & vbCrLf & _
				"Dim oMyResult1DCObject As Object" & vbCrLf & _
				"Set oMyResult1DCObject = Result1DComplex("""")" & vbCrLf & _
				"oMyResult1DCObject.Initialize(" & CStr(lLength) & ")" & vbCrLf & _
				"oMyResult1DCObject.SetArray(dXValues, ""x"")" & vbCrLf & _
				"oMyResult1DCObject.SetArray(dYReValues, ""yre"")" & vbCrLf & _
				"oMyResult1DCObject.SetArray(dYImValues, ""yim"")" & vbCrLf & _
				"oMyResult1DCObject.Title(""" & sTitle & """)" & vbCrLf & _
				"oMyResult1DCObject.SetXLabelAndUnit(""" & sXLabel & """, """ & sXUnit & """)" & vbCrLf & _
				"oMyResult1DCObject.SetYLabelAndUnit(""" & sYLabel & """, """ & sYUnit & """)" & vbCrLf & _
				"oMyResult1DCObject.SetLogarithmicFactor(" & CStr(dLogFac) & ")" & vbCrLf & _
				"oMyResult1DCObject.SetDefaultPlotView(""" & sPlotView & """)" & vbCrLf & _
				"oMyResult1DCObject.Save(""" & sFullFileName & """)" & vbCrLf & _
				"oMyResult1DCObject.AddToTree(""" & sFilePath & """)"

				' Set up caption string
				sAddToHistoryCaption = "" & vbCrLf & _
				"Add 1D result: " & sFileName
			Else
				ReportWarningToWindow("Unfortunately the result item in question seems to be empty and could not be saved succesfully.")
				CreateHistoryListEntryForAddingFile = False
			End If
		Else
			ReportWarningToWindow("Unfortunately the result item in question could not be saved succesfully.")
			CreateHistoryListEntryForAddingFile = False
		End If
	End If
End Function



' Creates entry for history list for deleting given result.
Private Function CreateHistoryListEntryForDeletingFile(sAddToHistoryContent As String, sAddToHistoryCaption As String, sSelectedResult As String) As Boolean
	sAddToHistoryContent = "" & _
	"Resulttree.Name(""" & sSelectedResult & """)" & vbCrLf & _
	"Resulttree.Delete"

	sAddToHistoryCaption = "" & _
	"Delete result: " & sSelectedResult

	CreateHistoryListEntryForDeletingFile = True
End Function



' -------------------------------------------------------------------------------------------------
' NoForbiddenFilenameCharactersExceptBackslashInMiddle:
' This function eliminates most characters from a string, which are not allowed to be used for
' filenames within CST Studio: ~^[]:,|*/$"()
' The exception are backlashes. Trailing and leading backslashes are removed and any consecutive
' backslashes are replaced by only one instance. This allows to define subfolders in this macro.
' -------------------------------------------------------------------------------------------------
Private Function NoForbiddenFilenameCharactersExceptBackslashInMiddle(s1 As String) As String
	' Declare And initialize
	Dim s2 As String
	s2 = s1

	s2 = Replace(s2,"~", "")
	s2 = Replace(s2,"^", "")
	' Replace square brackets
	s2 = Replace(s2,"]", ")")
	s2 = Replace(s2,"[", "(")
	s2 = Replace(s2,">", ")")
	s2 = Replace(s2,"<", "(")
	s2 = Replace(s2,":", "")
	s2 = Replace(s2,"|", "")
	s2 = Replace(s2,"*", "")
	s2 = Replace(s2,"/", "")
	s2 = Replace(s2,"$", "")
	s2 = Replace(s2,"""", "")
	s2 = Replace(s2,";", ",")

	' Remove consecutive backslashes
	While InStr("s2", "\\") > 0
		s2 = Replace(s2,"\\", "\")
	Wend

	' Remove leading and trailing white spaces and backslashes
	While ( ( Left(s2,1) = "\" ) Or ( Left(s2,1) = " " ) Or ( Right(s2,1) = "\" ) Or ( Right(s2,1) = " " ) )
		s2 = Trim(s2)
		If Left(s2,1)  = "\" Then s2 = Right(s2, Len(s2)-1)
		If Right(s2,1) = "\" Then s2 = Left(s2, Len(s2)-1)
	Wend

	NoForbiddenFilenameCharactersExceptBackslashInMiddle = s2
End Function
