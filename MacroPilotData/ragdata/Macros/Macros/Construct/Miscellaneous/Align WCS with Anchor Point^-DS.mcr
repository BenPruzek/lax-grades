'#Language "WWB-COM"

' ================================================================================================
' Copyright 2020-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 02-Sep-2020 mha: first version
'------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Populate array with names of anchor points
	Dim sArrayOfAnchorPointNamesLong()  As String	' Array with all anchor point (paths)
	Dim sArrayOfAnchorPointNamesShort() As String	' Array with all (shortened) anchor point paths
	Dim lNumberOfAnchorPoints           As Long		' Number of anchor points
	Dim lSelectedAnchorPointIndex       As Long		' Index of selected anchor point in array sArrayOfAnchorPointNamesLong
	Dim sSelectedAnchorPointName        As String	' Name of selected anchor point (without path)
	Dim bSaceWCSAtNewPosition           As Boolean	' Indicates whether user wanted to save wcs at new position of anchor point
	Dim sNameOfWCSAtNewPosition         As String	' If WCS should be changed, name is stored here.
	Dim lSaveWCSForAllAnchorPoints      As Long		' If WCS should be saved for all anchor points, this is indicated in this option (0 -> all anchor points saved; 1 -> only selected anchor point is saved)
	lNumberOfAnchorPoints = PopulateStringArrayWithNamesOfAnchorPoints(sArrayOfAnchorPointNamesLong)
	ReDim sArrayOfAnchorPointNamesShort(lNumberOfAnchorPoints)
	ShortenStringArrayWithNamesOfAnchorPoints(sArrayOfAnchorPointNamesLong, sArrayOfAnchorPointNamesShort, 37, lNumberOfAnchorPoints)

	' Call the define method and check whether it is completed successfully
	If (Define("test", True, False, sArrayOfAnchorPointNamesShort, lSelectedAnchorPointIndex, bSaceWCSAtNewPosition, sNameOfWCSAtNewPosition, lSaveWCSForAllAnchorPoints)) Then
		' If the define method is executed properly, call AlignWCS (as often as necessary, depending on options)
		' Skip everything in case no anchor points exist...
		If lNumberOfAnchorPoints > 0 Then
			Dim sAddToHistoryString1 As String
			Dim sAddToHistoryString2 As String
			Dim sWCSNameForCaption   As String
			If lSaveWCSForAllAnchorPoints = 0 Then ' Save WCS for all anchor points...
				For lSelectedAnchorPointIndex = 0 To UBound(sArrayOfAnchorPointNamesLong) STEP 1
					sAddToHistoryString1 = AlignWCS(sArrayOfAnchorPointNamesLong(lSelectedAnchorPointIndex), bSaceWCSAtNewPosition, sArrayOfAnchorPointNamesLong(lSelectedAnchorPointIndex), lSaveWCSForAllAnchorPoints, sAddToHistoryString2, sWCSNameForCaption)
					' Get name of anchor point
					sSelectedAnchorPointName = ConvertTreeItemPathToAPName(sArrayOfAnchorPointNamesLong(lSelectedAnchorPointIndex))
					If InStr(sSelectedAnchorPointName, ":") > 0 Then
						sSelectedAnchorPointName = Right(sSelectedAnchorPointName, Len(sSelectedAnchorPointName) - InStrRev(sSelectedAnchorPointName, ":"))
					End If
					AddToHistory("align WCS with anchor point: " & sSelectedAnchorPointName, sAddToHistoryString1)
					If sAddToHistoryString2 <> "" Then
						AddToHistory("save WCS as " & sWCSNameForCaption, sAddToHistoryString2)
					End If
				Next lSelectedAnchorPointIndex
			Else
				sAddToHistoryString1 = AlignWCS(sArrayOfAnchorPointNamesLong(lSelectedAnchorPointIndex), bSaceWCSAtNewPosition, sNameOfWCSAtNewPosition, lSaveWCSForAllAnchorPoints, sAddToHistoryString2, sWCSNameForCaption)
				' Get name of anchor point
				sSelectedAnchorPointName = ConvertTreeItemPathToAPName(sArrayOfAnchorPointNamesLong(lSelectedAnchorPointIndex))
				If InStr(sSelectedAnchorPointName, ":") > 0 Then
					sSelectedAnchorPointName = Right(sSelectedAnchorPointName, Len(sSelectedAnchorPointName) - InStrRev(sSelectedAnchorPointName, ":"))
				End If
				AddToHistory("align WCS with anchor point: " & sSelectedAnchorPointName, sAddToHistoryString1)
				If sAddToHistoryString2 <> "" Then
					AddToHistory("save WCS as " & sWCSNameForCaption, sAddToHistoryString2)
				End If
			End If
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False

End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, sArrayOfAnchorPointNames() As String, lSelectedAnchorPointIndex As Long, bSaceWCSAtNewPosition As Boolean, sNameOfWCSAtNewPosition As String, lSaveWCSForAllAnchorPoints As Long) As Boolean
	' Define the graphical dialog of the macro.

	Begin Dialog UserDialog 360, 216, "Align WCCS with anchor point", .DialogFunc ' %GRID:10,7,1,1
		' Groupbox
		GroupBox      15,  10, 330, 96, "Align WCS with selected anchor point(s)", .GBAlignWCSWithSelectedAnchorPoint
		OptionGroup                                                                .OGAllOrIndividual
		OptionButton  30,  28, 220,  14, "Save WCS for all anchor points",         .OBSaveWCSForAllAnchorPoints
		OptionButton  30,  48, 209,  14, "Select individual anchor point:",        .OBSaveWCSForSpecificAnchorPoint
		DropListBox   30,  71, 300,  21, sArrayOfAnchorPointNames(),               .DLBArrayOfAnchorPointNames
		' Groupbox
		GroupBox      15, 105, 330,  71, "Save new WCS position",                  .SaveCurrentWorkingCoordinateSystemPosition
		CheckBox      30, 126, 191,  14, "Save WCS at new position",               .CBSaveCurrentWCS
		Text          30, 150,  91,  21, "Specify name:",                          .TBSpecifyName
		TextBox      128, 147, 202,  21,                                           .TBNameOfNewlySavedWCS
		' OK and Cancel buttons
		OKButton      20, 186,  90,  21
		CancelButton 130, 186,  90,  21
	End Dialog

	' Initialize
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		' In case of a result template settings would be stored as ScriptSettings, to be retreived again later. Can also be used for macros if settings should be stored!
		lSelectedAnchorPointIndex  = dlg.DLBArrayOfAnchorPointNames
		bSaceWCSAtNewPosition      = dlg.CBSaveCurrentWCS
		sNameOfWCSAtNewPosition    = dlg.TBNameOfNewlySavedWCS
		lSaveWCSForAllAnchorPoints = dlg.OGAllOrIndividual

	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			DlgValue  "OGAllOrIndividual",          1
			DlgEnable "DLBArrayOfAnchorPointNames", True
			DlgEnable "CBSaveCurrentWCS",           True
			DlgValue  "CBSaveCurrentWCS",           0
			DlgEnable "TBNameOfNewlySavedWCS",      False
			DlgText   "TBNameOfNewlySavedWCS",      ""
		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "OGAllOrIndividual" ) Then
				If ( DlgValue("OGAllOrIndividual") = 0 ) Then	' WCSs should be saved for all anchor points, grey out all individual options
					DlgEnable "DLBArrayOfAnchorPointNames", False
					DlgEnable "CBSaveCurrentWCS",           False
					DlgEnable "TBNameOfNewlySavedWCS",      False
				Else 	' Only individual WCSs should be saved... Start off, just like for initialization, but don't reset values...
					DlgEnable "DLBArrayOfAnchorPointNames", True
					DlgEnable "CBSaveCurrentWCS",           True
					If DlgValue("CBSaveCurrentWCS") <> 0 Then
						DlgEnable "TBNameOfNewlySavedWCS", True
					Else
						DlgEnable "TBNameOfNewlySavedWCS", False
					End If
				End If
			ElseIf ( sDlgItem = "CBSaveCurrentWCS" ) Then
				If ( DlgValue("CBSaveCurrentWCS") = 0 ) Then	' WCS should not be saved, grey out textbox specifying name
					DlgEnable "TBNameOfNewlySavedWCS", False
				Else
					DlgEnable "TBNameOfNewlySavedWCS", True
				End If
			ElseIf ( sDlgItem = "OK" ) Then
				' The user pressed the Ok button. Check the settings and display an error message if some required fields have been left blank.
				' Verify that an expression has been entered into the textbox
				' A check whether the expression is valid or not is performed in Case 3...
				If ( DlgValue("CBSaveCurrentWCS") <> 0 ) And ( DlgValue("OGAllOrIndividual") = 1) Then	' WCS should be saved, make sure that a proper name has been specified
					If ( DlgText("TBNameOfNewlySavedWCS") = "" ) Then
						MsgBox "You have check marked ""Save WCS at New position"", but not entered a name for the WCS yet. Please enter a name for the WCS before pressing ""OK"" or uncheck ""Save WCS at New position"".", vbExclamation
						DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
					End If
				End If
			End If
		Case 3 ' TextBox or ComboBox text changed
			' Check to see whether the entry confirms to CST naming standards (by removing any characters that are not allowed).
			Dim sValueSetOld As String	' string as entered in textbox
			Dim sValueSetNew As String	' string where forbidden characters have been removed
			sValueSetOld = DlgText(sDlgItem)
			sValueSetNew = NoForbiddenFilenameCharacters(sValueSetOld)
			DlgText  "TBNameOfNewlySavedWCS", sValueSetNew
			If ( Len(sValueSetNew) < Len(sValueSetOld) ) Then
				ReportWarningToWindow("Align WCCS with anchor point: Please note that the filename you have entered has been changed slightly, forbidden characters in the string have been removed.")
			End If
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function



' -------------------------------------------------------------------------------------------------
' AlignWCS: This function returns a string which contains the commands to align the working
' coordinate system with a selected anchor point (or save a WCS for all anchor points)
' -------------------------------------------------------------------------------------------------
Private Function AlignWCS(sAnchorPoint As String, bSaceWCSAtNewPosition As Boolean, sNameOfWCSAtNewPosition As String, lSaveWCSForAllAnchorPoints As Long, sHistoryString2 As String, sWCSNameForCaption As String) As String
	Dim sHistoryString         As String
	Dim sMyAnchorPointTreePath As String
	Dim sNewName               As String

	sMyAnchorPointTreePath = ConvertTreeItemPathToAPName(sAnchorPoint)

	' In case WCSs should be stored for all anchor points, set bSaceWCSAtNewPosition to True
	If lSaveWCSForAllAnchorPoints = 0 Then bSaceWCSAtNewPosition = True

	sHistoryString = "' Check if WCS is currently active. If not, activate it." & vbNewLine & _
	"If LCase(WCS.IsWCSActive) = ""global"" Then" & vbNewLine & _
	vbTab & "WCS.ActivateWCS(""local"")" & vbNewLine & _
	"End If" & vbNewLine & _
	"" & vbNewLine & _
	"' Align WCS to anchor point" & vbNewLine & _
	"AnchorPoint.Restore(""" & sMyAnchorPointTreePath & """)" & vbNewLine & _
	"" & vbNewLine

	' Set up a new string sAddToHistoryString2 for saving the WCS.
	' This is done to get around that AddToHistory will replace the last entry if the caption is equal.
	sHistoryString2    = ""
	sWCSNameForCaption = ""
	If bSaceWCSAtNewPosition Then
		' Check to see whether the name contains any slashes or backslashes. If so, the option "Save WCS for all anchor points" must have been used. In this case, replace all slashes / backlashes by hyphens.
		sNewName = Replace(sNameOfWCSAtNewPosition, "\", "-")
		sNewName = Replace(sNewName, "/", "-")
		sHistoryString2 = "' Store WCS at new position" & vbNewLine & _
		"WCS.Store(""" & sNewName & """)" & vbNewLine
		sWCSNameForCaption = sNewName
	End If

	AlignWCS = sHistoryString
End Function



' -------------------------------------------------------------------------------------------------
' PopulateStringArrayWithNamesOfAnchorPoints: This function takes a string array and populates it
' with the paths of all anchor points. It returns the number of entries.
' s_AnchorPointName:            String Array, lists all anchor points in Anchor Points folder, populated by function
' -------------------------------------------------------------------------------------------------
Private Function PopulateStringArrayWithNamesOfAnchorPoints(s_AnchorPointNamesArray() As String) As Long
	Dim lLengthOfAnchorPointNameArrayIncr As Long
	Dim lLengthOfFolderNameArrayIncr      As Long
	Dim lNumberOfAnchorPoints             As Long

	lLengthOfAnchorPointNameArrayIncr = 10
	lLengthOfFolderNameArrayIncr      = 10
	lNumberOfAnchorPoints = ParseFolderIncludingSubfoldersForAnchorPoints("Anchor Points", s_AnchorPointNamesArray, lLengthOfAnchorPointNameArrayIncr, lLengthOfFolderNameArrayIncr)
	PopulateStringArrayWithNamesOfAnchorPoints = lNumberOfAnchorPoints
End Function



' -------------------------------------------------------------------------------------------------
' ShortenStringArrayWithNamesOfAnchorPoints: This function takes a string array and shortens all
' entries down to a specified length. It returns the number of entries in the array.
' sStringArray:      Array which entries are shortened
' sStringArrayShort: Array with shortened entries
' lShortenDown:      Number of characters entries are shortened down to
' lNumberOfAncPoint: Number of anchor points
' -------------------------------------------------------------------------------------------------
Private Function ShortenStringArrayWithNamesOfAnchorPoints(sStringArray() As String, sStringArrayShort() As String, lShortenDown As Long, lNumberOfAncPoint As Long) As Long
	Dim lRunningIndex    As Long
	Dim lNumberOfEntries As Long

	If lNumberOfAncPoint > 0 Then
		lNumberOfEntries = UBound(sStringArray)

		' In the first step, get rid of "Anchor Points\"
		For lRunningIndex = 0 To lNumberOfEntries STEP 1
			sStringArrayShort(lRunningIndex) = Right(sStringArray(lRunningIndex), Len(sStringArray(lRunningIndex)) - Len("Anchor Points\"))
			If Len(sStringArrayShort(lRunningIndex)) > lShortenDown Then
				sStringArrayShort(lRunningIndex) = "..." & Right(sStringArrayShort(lRunningIndex), lShortenDown)
			End If
		Next lRunningIndex
		lNumberOfEntries = lNumberOfEntries + 1
	Else
		lNumberOfEntries = lNumberOfAncPoint
	End If

	ShortenStringArrayWithNamesOfAnchorPoints = lNumberOfEntries
End Function



' -------------------------------------------------------------------------------------------------
' Accepts treeitem as string, will give back designation as "AnchorPoint", "Folder" or "Other"
' -------------------------------------------------------------------------------------------------
Private Function DetermineDesignationOfTreeItemAnchorPointVersion(sPath As String) As String
	Dim sDesignation As String
	Dim sPathAPName  As String

	' Replace slashes or colons with backslashes
	sPath = Replace(sPath, "/", "\")
	sPath = Replace(sPath, ":", "\")

	' Strip trailing slashes, backslashes or colons if existing
	If Right(sPath,1) = "\" Or Right(sPath,1) = "/" Or Right(sPath,1) = ":" Then sPath = Left(sPath, Len(sPath)-1)

	' Check whether string exists as treeitem
	If Resulttree.DoesTreeItemExist(sPath) Then
		' Check whether treeitem exists in the Anchor Points tree
		If Left(sPath, 14) = "Anchor Points\" Then
			' Check if treeitem has children
			If Resulttree.GetFirstChildName(sPath) <> "" Then
				sDesignation = "Folder"
			Else
				' Check if treeitem is an anchor point
				sPathAPName = ConvertTreeItemPathToAPName(sPath)
				If AnchorPoint.DoesExist(sPathAPName) Then
					sDesignation = "AnchorPoint"
				Else
					sDesignation = "Folder"
				End If
			End If
		ElseIf sPath = "Anchor Points" Then
			sDesignation = "Folder"
		Else
			sDesignation = "Other"
		End If
	Else
		sDesignation = "Other"
	End If

	DetermineDesignationOfTreeItemAnchorPointVersion = sDesignation
End Function



' -------------------------------------------------------------------------------------------------
' Takes string, copies it, replaces last backslash in copy with colon, other backslashes with slashes and removes first twelve characters.
' Thereby a valid solid treeitem is returned as a solidname which can be used in the Solid object.
' If the given string has less or equal to 12 characters, an empty string is returned.
' -------------------------------------------------------------------------------------------------
Private Function ConvertTreeItemPathToAPName(sTreeItemPath As String) As String
	Dim sPathAPName    As String
	Dim sAPName        As String
	Dim sFolderNameTmp As String
	Dim sFolderName    As String

	If Len(sTreeItemPath) > 14 Then
		' Isolate anchor point name and folder name
		sAPName        = Right(sTreeItemPath, Len(sTreeItemPath) - InStrRev(sTreeItemPath, "\"))
		sFolderNameTmp = Left(sTreeItemPath, Len(sTreeItemPath) - Len(sAPName) - 1)
		' In case of a folder / anchor point directly in the Anchor Points root folder
		If Len(sFolderNameTmp) < 14 Then
			sFolderName = ""
		Else
			sFolderName = Right(sFolderNameTmp, Len(sFolderNameTmp) - 14)
		End If
		' replace backslashes by slashes and add in colon between component name and solid name
		sFolderName = Replace(sFolderName, "\", "/")
		If sFolderName <> "" Then
			sPathAPName = sFolderName & ":" & sAPName
		Else
			sPathAPName = sAPName
		End If
	Else
		sPathAPName = ""
	End If

	ConvertTreeItemPathToAPName = sPathAPName
End Function



' -------------------------------------------------------------------------------------------------
' ParseFolderIncludingSubfoldersForAnchorPoints: Parses given folder for all anchor points contained therein
' Accepts given folder and transfers list of all anchor points in that folder - including subfolders! - to given array.
' sFolderName:                       Name of folder for which all anchor points are to be listed
' s_AnchorPointNames:                Array containing anchor point names, added to by the function - the assumption is that s_AnchorPointNames might already contain values, which the function adds to
' lLengthOfAnchorPointNameArrayIncr: Length of initial (temp) array, furthermore array is padded by this amount whenever necessary.
' lLengthOfFolderNameArrayIncr:      Length of initial (temp) array, furthermore array is padded by this amount whenever necessary.
' -------------------------------------------------------------------------------------------------
Private Function ParseFolderIncludingSubfoldersForAnchorPoints(sFolderName As String, s_AnchorPointNames() As String, lLengthOfAnchorPointNameArrayIncr As Long, lLengthOfFolderNameArrayIncr As Long) As Long
	Dim s_FolderName()                As String  ' String array of folders
	Dim l_FolderParent()              As Long    ' Index of parent folder
	Dim l_FolderNumChildren()         As Long    ' Number of folder children
	Dim l_FolderNumChildrenSearched() As Long    ' Number of folder children that have already been searched / catalogued
	Dim lCurrentNumFolders            As Long    ' Number of folders that have been catalogued
	Dim lCurrentFolderIndex           As Long    ' Index of current component
	Dim lCurrentParentIndex           As Long    ' Index of current parent component
	Dim lCurrentParentLevel           As Long    ' Level of current parent component
	Dim lRunningIndex01               As Long    ' Running index
	Dim lRunningIndex02               As Long    ' Running index
	Dim lRootNumFolderChildren        As Long    ' Number of folder children of "Anchor Points"
	Dim sCurrentTreeItem              As String
	Dim bGoDownStep                   As Boolean ' Flag
	Dim bSearchingFolderChild         As Boolean ' Flag, true while child of "Anchor Points" should still be searched
	Dim lOverallNumberOfAnchorPoints  As Long    ' Overall number of anchor points

	' Start with lLengthOfFolderNameArrayIncr entries for components, expand as needed, similiar for solids
	ReDim s_FolderName(lLengthOfFolderNameArrayIncr)
	ReDim l_FolderParent(lLengthOfFolderNameArrayIncr)
	ReDim l_FolderNumChildren(lLengthOfFolderNameArrayIncr)
	ReDim l_FolderNumChildrenSearched(lLengthOfFolderNameArrayIncr)
	ReDim l_FolderLevel(lLengthOfFolderNameArrayIncr)
	ReDim s_AnchorPointNames(lLengthOfAnchorPointNameArrayIncr)

	' Initialize
	lCurrentNumFolders = 0

	' Catalogue root folder of search
	sCurrentTreeItem               = sFolderName
	lRootNumFolderChildren         = GetNumberOfFolderChildren(sCurrentTreeItem)
	s_FolderName(0)                = sCurrentTreeItem
	l_FolderParent(0)              = -1
	l_FolderNumChildren(0)         = lRootNumFolderChildren
	l_FolderNumChildrenSearched(0) = 0
	l_FolderLevel(0)               = 0
	lCurrentNumFolders             = lCurrentNumFolders + 1
	' Catalogue solids in root folder of search
	lOverallNumberOfAnchorPoints = ParseFolderWithoutSubfoldersForAnchorPoints(sFolderName, s_AnchorPointNames, lLengthOfAnchorPointNameArrayIncr, 0)

	For lRunningIndex01 = 1 To lRootNumFolderChildren STEP 1
		' Go to lRunningIndex01 child and catalogue item
		GoToNthChild(sFolderName, sCurrentTreeItem, lRunningIndex01)
		s_FolderName(lCurrentNumFolders)                = sCurrentTreeItem
		l_FolderParent(lCurrentNumFolders)              = 0
		l_FolderNumChildren(lCurrentNumFolders)         = GetNumberOfFolderChildren(sCurrentTreeItem)
		l_FolderNumChildrenSearched(lCurrentNumFolders) = 0
		l_FolderLevel(lCurrentNumFolders)               = 1
		lCurrentNumFolders                              = lCurrentNumFolders + 1
		' Catalogue Anchor Points as well...
		lOverallNumberOfAnchorPoints = lOverallNumberOfAnchorPoints + ParseFolderWithoutSubfoldersForAnchorPoints(sCurrentTreeItem, s_AnchorPointNames, lLengthOfAnchorPointNameArrayIncr, lOverallNumberOfAnchorPoints)
		' If folder children exist, start cataloguing
		If l_FolderNumChildren(lCurrentNumFolders-1) > 0 Then
			GoToNthChild(s_FolderName(lCurrentNumFolders-1), sCurrentTreeItem, 1)
			l_FolderNumChildrenSearched(lCurrentNumFolders-1) = 1
			lCurrentParentIndex                               = lCurrentNumFolders - 1
			lCurrentParentLevel                               = 1
			bGoDownStep                                       = True
			bSearchingFolderChild                             = True
		Else
			bSearchingFolderChild = False
		End If
		While bSearchingFolderChild = True
			' First, check if size of arrays has to be expanded.
			If UBound(s_FolderName) - 1 = lCurrentNumFolders Then
				ReDim Preserve s_FolderName(lCurrentNumFolders + lLengthOfFolderNameArrayIncr)
				ReDim Preserve l_FolderParent(lCurrentNumFolders + lLengthOfFolderNameArrayIncr)
				ReDim Preserve l_FolderNumChildren(lCurrentNumFolders + lLengthOfFolderNameArrayIncr)
				ReDim Preserve l_FolderNumChildrenSearched(lCurrentNumFolders + lLengthOfFolderNameArrayIncr)
				ReDim Preserve l_FolderLevel(lCurrentNumFolders + lLengthOfFolderNameArrayIncr)
			End If
			If bGoDownStep = True Then
			' If GoDownStep
			'    -> catalogue folder as indicated by CurrentTreeItem
			'    -> If catalogued item has folder children, set GoDownFlag and indicate child
			'    -> If catalogued item has no folder children, set GoUpFlag and indicate parent
				s_FolderName(lCurrentNumFolders)                = sCurrentTreeItem
				l_FolderParent(lCurrentNumFolders)              = lCurrentParentIndex
				l_FolderNumChildren(lCurrentNumFolders)         = GetNumberOfFolderChildren(sCurrentTreeItem)
				l_FolderNumChildrenSearched(lCurrentNumFolders) = 0
				l_FolderLevel(lCurrentNumFolders)               = lCurrentParentLevel + 1
				lOverallNumberOfAnchorPoints = lOverallNumberOfAnchorPoints + ParseFolderWithoutSubfoldersForAnchorPoints(sCurrentTreeItem, s_AnchorPointNames, lLengthOfAnchorPointNameArrayIncr, lOverallNumberOfAnchorPoints)
				If l_FolderNumChildren(lCurrentNumFolders) > 0 Then
					GoToNthChild(s_FolderName(lCurrentNumFolders), sCurrentTreeItem, 1)
					l_FolderNumChildrenSearched(lCurrentNumFolders) = 1
					lCurrentParentIndex                             = lCurrentNumFolders
					lCurrentParentLevel                             = l_FolderLevel(lCurrentNumFolders)
					bGoDownStep                                     = True
				Else
					lCurrentFolderIndex = l_FolderParent(lCurrentNumFolders)
					bGoDownStep             = False
				End If
				' hold off on updating the number of catalogued components, until all values pertaining to the current component have been set
				lCurrentNumFolders                                 = lCurrentNumFolders + 1
			Else
			' If GoUpStep
			'    If SearchedChildren < ExistingChildren
			'       -> Set GoDownStep Flag and indicate SearchedChildren+1 child
			'    Else If arrived at level 1
			'       -> then stop / break out of while loop
			'    Else If SearchedChildren >= ExistingChildren but not at level 1
			'       -> Set GoUpStep Flag, indicating parent
				' Check whether the component still has children which have not been searched yet
				If l_FolderNumChildrenSearched(lCurrentFolderIndex) < l_FolderNumChildren(lCurrentFolderIndex) Then
					GoToNthChild(s_FolderName(lCurrentFolderIndex), sCurrentTreeItem, l_FolderNumChildrenSearched(lCurrentFolderIndex)+1)
					l_FolderNumChildrenSearched(lCurrentFolderIndex) = l_FolderNumChildrenSearched(lCurrentFolderIndex) + 1
					lCurrentParentIndex                              = lCurrentFolderIndex
					lCurrentParentLevel                              = l_FolderLevel(lCurrentFolderIndex)
					bGoDownStep                                      = True
				Else
					' Check whether component is at level 1 (child of "Components")
					If l_FolderLevel(lCurrentFolderIndex) = 1 Then
						bSearchingFolderChild             = False
					Else
						sCurrentTreeItem    = s_FolderName(l_FolderParent(lCurrentFolderIndex))
						lCurrentFolderIndex = l_FolderParent(lCurrentFolderIndex)
						bGoDownStep         = False
					End If
				End If
			End If
		Wend
	Next lRunningIndex01

	If lOverallNumberOfAnchorPoints > 0 Then
		ReDim Preserve s_AnchorPointNames(lOverallNumberOfAnchorPoints-1)
	Else
		ReDim s_AnchorPointNames(0)
	End If

	ParseFolderIncludingSubfoldersForAnchorPoints = lOverallNumberOfAnchorPoints
End Function



' -------------------------------------------------------------------------------------------------
' Accepts treeitem as string, returns number of children which are folders
' -------------------------------------------------------------------------------------------------
Private Function GetNumberOfFolderChildren(sPath As String) As Long
	Dim lNumFolderChildren As Long
	lNumFolderChildren = 0

	' Replace slashes or colons with backslashes
	sPath = Replace(sPath, "/", "\")
	sPath = Replace(sPath, ":", "\")

	' Strip trailing slashes, backslashes or colons if existing
	If Right(sPath,1) = "\" Or Right(sPath,1) = "/" Or Right(sPath,1) = ":" Then sPath = Left(sPath, Len(sPath)-1)

	' Check whether string exists as treeitem
	If Resulttree.DoesTreeItemExist(sPath) Then
		' Check that treeitem is a folder
		If DetermineDesignationOfTreeItemAnchorPointVersion(sPath) = "Folder" Then
			Dim sCurrentTreeItem As String
			sCurrentTreeItem = Resulttree.GetFirstChildName(sPath)
			If sCurrentTreeItem <> "" Then
				Dim sCurrentTreeItemIsFolder As Boolean
				If DetermineDesignationOfTreeItemAnchorPointVersion(sCurrentTreeItem) = "Folder" Then
					sCurrentTreeItemIsFolder = True
					While sCurrentTreeItemIsFolder = True
						lNumFolderChildren = lNumFolderChildren + 1
						sCurrentTreeItem = Resulttree.GetNextItemName(sCurrentTreeItem)
						If DetermineDesignationOfTreeItemAnchorPointVersion(sCurrentTreeItem) <> "Folder" Then
							sCurrentTreeItemIsFolder = False
						End If
					Wend
				Else
					lNumFolderChildren = 0
				End If
			Else
				lNumFolderChildren = 0
			End If
		Else
			lNumFolderChildren = 0
		End If
	Else
		lNumFolderChildren = 0
	End If

	GetNumberOfFolderChildren = lNumFolderChildren
End Function



' -------------------------------------------------------------------------------------------------
' Accepts given folder and transfers list of all anchor points in that folder - not including subfolders! - to given array.
' sFolderName:                       Name of folder for which all anchor points are to be listed (subfolders are not considered)
' s_AnchorPointNames:                Array containing anchor point names, added to by the function - the assumption is that s_AnchorPointNames might already contain values, which the function adds to
' lLengthOfAnchorPointNameArrayIncr: Length of initial (temp) array, furthermore array is padded by this amount whenever necessary.
' lCurrentNumberOfAnchorPoints:      Current number of anchor points, s_AnchorPointNames is extended from this index forth, regardless of initial length (if there are solids that are present, otherwise the array is not touched in any way)
' -------------------------------------------------------------------------------------------------
Private Function ParseFolderWithoutSubfoldersForAnchorPoints(sFolderName As String, s_AnchorPointNames() As String, lLengthOfAnchorPointNameArrayIncr As Long, lCurrentNumberOfAnchorPoints As Long) As Long
	Dim lRunningIndex01            As Long    ' Running Index
	Dim s_AnchorPointNamesTemp()   As String  ' Set up temporary string array with anchor point names and merge at the end
	Dim sCurrentTreeItem           As String  ' Current tree item...
	Dim lNumberOfAnchorPoints      As Long    ' Total number of anchor points found
	Dim lNumberOfFolders           As Long    ' Number of folder children in folder
	Dim bAllAnchorPointsCatalogued As Boolean ' Flag...

	bAllAnchorPointsCatalogued = False
	' Start with lLengthOfAnchorPointNameArrayIncr entries, expand as needed
	ReDim s_AnchorPointNamesTemp(lLengthOfAnchorPointNameArrayIncr)

	' Determine number of folder children of given folder, then initialize sCurrentTreeItem
	lNumberOfFolders = GetNumberOfFolderChildren(sFolderName)
	If GoToNthChild(sFolderName, sCurrentTreeItem, lNumberOfFolders+1) Then
		s_AnchorPointNamesTemp(0) = sCurrentTreeItem
		lNumberOfAnchorPoints     = 1
		While Not bAllAnchorPointsCatalogued
			For lRunningIndex01 = 1 To lLengthOfAnchorPointNameArrayIncr STEP 1
				sCurrentTreeItem = Resulttree.GetNextItemName(sCurrentTreeItem)
				If sCurrentTreeItem <> "" Then
					s_AnchorPointNamesTemp(lNumberOfAnchorPoints) = sCurrentTreeItem
					lNumberOfAnchorPoints                   = lNumberOfAnchorPoints + 1
				Else
					bAllAnchorPointsCatalogued = True
					Exit For
				End If
			Next lRunningIndex01
			' Expand size of array
			If Not bAllAnchorPointsCatalogued Then
				ReDim Preserve s_AnchorPointNamesTemp(UBound(s_AnchorPointNamesTemp) + lLengthOfAnchorPointNameArrayIncr)
			End If
		Wend
		' Merge arrays and update overall number of anchor points in array

		ReDim Preserve s_AnchorPointNames(lCurrentNumberOfAnchorPoints + lNumberOfAnchorPoints - 1)
		For lRunningIndex01 = 0 To lNumberOfAnchorPoints-1 STEP 1
			s_AnchorPointNames(lCurrentNumberOfAnchorPoints + lRunningIndex01) = s_AnchorPointNamesTemp(lRunningIndex01)
		Next lRunningIndex01
	Else
		' The component does not contain any anchor points
		lNumberOfAnchorPoints = 0
	End If
	ParseFolderWithoutSubfoldersForAnchorPoints = lNumberOfAnchorPoints
End Function



' Saves path of lN'th child of sParentPath in sNthChildPath. Returns True if operation was successfull and False otherwise. First child corresponds to lN = 1, not lN = 0
Private Function GoToNthChild(sParentPath As String, sNthChildPath As String, lN As Long) As Boolean
	Dim bReturnValue     As Boolean
	Dim sCurrentTreeItem As String
	Dim lRunningIndex    As Long

	bReturnValue     = True
	sCurrentTreeItem = Resulttree.GetFirstChildName(sParentPath)
	If lN > 0 Then
		If sCurrentTreeItem <> "" Then
			If lN > 1 Then
				For lRunningIndex = 1 To lN - 1 STEP 1
					sCurrentTreeItem = Resulttree.GetNextItemName(sCurrentTreeItem)
					If sCurrentTreeItem = "" Then
						sNthChildPath = ""
						bReturnValue  = False
						Exit For
					End If
					sNthChildPath = sCurrentTreeItem
				Next lRunningIndex
			Else
				sNthChildPath = sCurrentTreeItem
			End If
		Else
			sNthChildPath = ""
			bReturnValue  = False
		End If
	Else
		sNthChildPath = ""
		bReturnValue  = False
	End If

	GoToNthChild = bReturnValue
End Function
