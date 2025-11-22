'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Macro: deactivates setting "Overwrite material color"
'
' Copyright 2021-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 05-Mar-2021 mha: first version
' ================================================================================================

' *** global variables
' No global variables specified...

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Populate array with component names
	Dim sComponents()     As String  ' Array with all components
	Dim sComponent        As String  ' Specific component which is chosen
	Dim lComponentLevel() As Long    ' Array with level of components
	Dim lDeepestCompLevel As Long    ' Deepest level encountered
	Dim sSolids()         As String  ' Array containing all solids in selected component
	Dim bAllSolids        As Boolean ' Indicates whether setting should be deactivated for all solids
	Dim lNumSolids        As Long    ' Number of solids for which setting is to be deactivated.
	Dim lRunningIndex     As Long    ' Running index.

	ParseComponentsTreeForComponents(sComponents, lComponentLevel, 20, lDeepestCompLevel)

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, sComponents, sComponent, sSolids, bAllSolids, lNumSolids) ) Then
		' If the define method is executed properly, define string for history list to deactivate setting for given solids.
		If bAllSolids Then
			Dim lNumberOfSolidsMin1 As Long
			lNumberOfSolidsMin1 = Solid.GetNumberOfShapes - 1
			ReDim sSolids(lNumberOfSolidsMin1)
			For lRunningIndex = 0 To lNumberOfSolidsMin1 STEP 1
				sSolids(lRunningIndex) = Solid.GetNameOfShapeFromIndex(lRunningIndex)
			Next lRunningIndex
		Else
			' Convert solid names
			For lRunningIndex = 0 To UBound(sSolids) STEP 1
				sSolids(lRunningIndex) = ConvertTreeItemPathToSolidName(sSolids(lRunningIndex))
			Next lRunningIndex
		End If
		' Go through solids array construct string for history list
		Dim sAddToHistoryString As String
		sAddToHistoryString = ""
		For lRunningIndex = 0 To UBound(sSolids) STEP 1
			sAddToHistoryString = sAddToHistoryString + "Solid.SetUseIndividualColor """ & sSolids(lRunningIndex) & """, 0" & vbCrLf
		Next lRunningIndex
		AddToHistory("Deactivate color overwrite for selected solids", sAddToHistoryString)
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, sComponents() As String, sComponent As String, sSolids() As String, bAllSolids As Boolean, lNumSolids As Long) As Boolean

	Begin Dialog UserDialog 420, 135, "Deactivate setting ""Overwrite material color""", .DialogFunc ' %GRID:3,3,1,1

		' Groupbox - select solids
		GroupBox      15,   5, 390, 95,"Selection of solids ...",       .GBSelectComponent
		OptionGroup                                                 .OGWhatKindOfSelection
		OptionButton  30,  25, 167, 14,"Deactivate for all solids", .OBDeactivateForAllSolids
		OptionButton  30,  45, 190, 15,"Select specific component", .OBDeactivateForSpecificComponent
		DropListBox   30,  70, 360, 18, sComponents(),               .DLBComponents
		' OK and Cancel buttons
		OKButton      15, 108,  90, 21
		CancelButton 115, 108,  90, 21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg               As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		' Set variables
		bAllSolids = Not CBool(dlg.OGWhatKindOfSelection)
		If Not bAllSolids Then
			sComponent = sComponents(dlg.DLBComponents)
			lNumSolids = ParseComponentIncludingSubcomponentsForSolids(sComponent, sSolids, 50, 50)
		End If
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunc: This function defines the dialog box behaviour. It is automatically called
'             whenever the user changes some settings in the dialog box, presses any button
'             or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			DlgEnable "OGWhatKindOfSelection", True
			DlgValue  "OGWhatKindOfSelection", 0
			DlgEnable "DLBComponents",         False
		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "OGWhatKindOfSelection" ) Then
				If ( DlgValue("OGWhatKindOfSelection") <> 0 ) Then	' Do no simply deactivate for all solids...
					DlgEnable "DLBComponents", True
				Else  ' Deactivate for all solids...
					DlgEnable "DLBComponents", False
				End If
			End If
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function



' Goes through components tree and transfers list of all components to given array. Returns the number of components that are present (not counting the root "Components")
' s_ComponentName: String array, lists all components in component tree, populated by function
' l_ComponentLevel: Long array, lists level of all components ("Components" -> level 0, child of "Components" is level 1, child of a child of "Components" is level 2,...), populated by function
' lLengthOfComponentNameArrayIncr: Length of initial array, furthermore array is padded by this amount whenever necessary.
' lExistingComponentLevels: Deepest component level encountered
Function ParseComponentsTreeForComponents(s_ComponentName() As String, l_ComponentLevel() As Long, lLengthOfComponentNameArrayIncr As Long, lExistingComponentLevels As Long) As Long
	Dim l_ComponentParent()              As Long    ' Index of parent component
	Dim l_ComponentNumChildren()         As Long    ' Number of component children
	Dim l_ComponentNumChildrenSearched() As Long    ' Number of component children that have already been searched / catalogued
	Dim lCurrentNumComponents            As Long    ' Number of components that have been catalogued
	Dim lCurrentComponentsIndex          As Long    ' Index of current component
	Dim lCurrentParentIndex              As Long    ' Index of current parent component
	Dim lCurrentParentLevel              As Long    ' Level of current parent component
	Dim lRunningIndex01                  As Long    ' Running index
	Dim lRunningIndex02                  As Long    ' Running index
	Dim lRootNumCompChildren             As Long    ' Number of component children of "Components"
	Dim sCurrentTreeItem                 As String
	Dim bGoDownStep                      As Boolean ' Flag
	Dim bSearchingComponentChild         As Boolean ' Flag, true while child of "Components" should still be searched
	Dim lDeepestLevel                    As Long    ' Deepest component level which is present

	' Start with lLengthOfComponentNameArrayIncr entries for components, expand as needed
	ReDim s_ComponentName(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentParent(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentNumChildren(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentNumChildrenSearched(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentLevel(lLengthOfComponentNameArrayIncr)

	' Initialize
	lCurrentNumComponents = 0

	' Catalogue root of components tree
	sCurrentTreeItem                  = "Components"
	lRootNumCompChildren              = GetNumberOfComponentChildren(sCurrentTreeItem)
	s_ComponentName(0)                = sCurrentTreeItem
	l_ComponentParent(0)              = -1
	l_ComponentNumChildren(0)         = lRootNumCompChildren
	l_ComponentNumChildrenSearched(0) = 0
	l_ComponentLevel(0)               = 0
	lDeepestLevel                     = 0
	lCurrentNumComponents             = lCurrentNumComponents + 1

	For lRunningIndex01 = 1 To lRootNumCompChildren STEP 1
		' Go to lRunningIndex01 child and catalogue item
		GoToNthChild("Components", sCurrentTreeItem, lRunningIndex01)
		s_ComponentName(lCurrentNumComponents)                = sCurrentTreeItem
		l_ComponentParent(lCurrentNumComponents)              = 0
		l_ComponentNumChildren(lCurrentNumComponents)         = GetNumberOfComponentChildren(sCurrentTreeItem)
		l_ComponentNumChildrenSearched(lCurrentNumComponents) = 0
		l_ComponentLevel(lCurrentNumComponents)               = 1
		If l_ComponentLevel(lCurrentNumComponents) > lDeepestLevel Then lDeepestLevel = l_ComponentLevel(lCurrentNumComponents)
		lCurrentNumComponents                                 = lCurrentNumComponents + 1
		' If component children exist, start cataloguing
		If l_ComponentNumChildren(lCurrentNumComponents-1) > 0 Then
			GoToNthChild(s_ComponentName(lCurrentNumComponents-1), sCurrentTreeItem, 1)
			l_ComponentNumChildrenSearched(lCurrentNumComponents-1) = 1
			lCurrentParentIndex                                     = lCurrentNumComponents - 1
			lCurrentParentLevel                                     = 1
			bGoDownStep = True
			bSearchingComponentChild = True
		Else
			bSearchingComponentChild = False
		End If
		While bSearchingComponentChild = True
			' First, check if size of arrays has to be expanded.
			If UBound(s_ComponentName) - 1 = lCurrentNumComponents Then
				ReDim Preserve s_ComponentName(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentParent(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentNumChildren(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentNumChildrenSearched(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentLevel(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
			End If
			If bGoDownStep = True Then
			' If GoDownStep
			'    -> catalogue component as indicated by CurrentTreeItem
			'    -> If catalogued item has component children, set GoDownFlag and indicate child
			'    -> If catalogued item has no component children, set GoUpFlag and indicate parent
				s_ComponentName(lCurrentNumComponents)                = sCurrentTreeItem
				l_ComponentParent(lCurrentNumComponents)              = lCurrentParentIndex
				l_ComponentNumChildren(lCurrentNumComponents)         = GetNumberOfComponentChildren(sCurrentTreeItem)
				l_ComponentNumChildrenSearched(lCurrentNumComponents) = 0
				l_ComponentLevel(lCurrentNumComponents)               = lCurrentParentLevel + 1
				If l_ComponentLevel(lCurrentNumComponents) > lDeepestLevel Then lDeepestLevel = l_ComponentLevel(lCurrentNumComponents)
				If l_ComponentNumChildren(lCurrentNumComponents) > 0 Then
					GoToNthChild(s_ComponentName(lCurrentNumComponents), sCurrentTreeItem, 1)
					l_ComponentNumChildrenSearched(lCurrentNumComponents) = 1
					lCurrentParentIndex                                   = lCurrentNumComponents
					lCurrentParentLevel                                   = l_ComponentLevel(lCurrentNumComponents)
					bGoDownStep                                           = True
				Else
					lCurrentComponentsIndex = l_ComponentParent(lCurrentNumComponents)
					bGoDownStep             = False
				End If
				' hold off on updating the number of catalogued components, until all values pertaining to the current component have been set
				lCurrentNumComponents                                 = lCurrentNumComponents + 1
			Else
			' If GoUpStep
			'    If SearchedChildren < ExistingChildren
			'       -> Set GoDownStep Flag and indicate SearchedChildren+1 child
			'    Else If arrived at level 1
			'       -> then stop / break out of while loop
			'    Else If SearchedChildren >= ExistingChildren but not at level 1
			'       -> Set GoUpStep Flag, indicating parent
				' Check whether the component still has children which have not been searched yet
				If l_ComponentNumChildrenSearched(lCurrentComponentsIndex) < l_ComponentNumChildren(lCurrentComponentsIndex) Then
					GoToNthChild(s_ComponentName(lCurrentComponentsIndex), sCurrentTreeItem, l_ComponentNumChildrenSearched(lCurrentComponentsIndex)+1)
					l_ComponentNumChildrenSearched(lCurrentComponentsIndex) = l_ComponentNumChildrenSearched(lCurrentComponentsIndex) + 1
					lCurrentParentIndex                                     = lCurrentComponentsIndex
					lCurrentParentLevel                                     = l_ComponentLevel(lCurrentComponentsIndex)
					bGoDownStep = True
				Else
					' Check whether component is at level 1 (child of "Components")
					If l_ComponentLevel(lCurrentComponentsIndex) = 1 Then
						bSearchingComponentChild = False
					Else
						sCurrentTreeItem        = s_ComponentName(l_ComponentParent(lCurrentComponentsIndex))
						lCurrentComponentsIndex = l_ComponentParent(lCurrentComponentsIndex)
						bGoDownStep = False
					End If
				End If
			End If
		Wend
	Next lRunningIndex01

	lExistingComponentLevels = lDeepestLevel
	ReDim Preserve s_ComponentName(lCurrentNumComponents-1)
	ReDim Preserve l_ComponentLevel(lCurrentNumComponents-1)
	ParseComponentsTreeForComponents = lCurrentNumComponents-1
End Function



' Accepts given component and transfers list of all solids in that component - including subcomponents - to given array.
' sComponentName: Name of component for which all solids are to be listed
' s_SolidNames: Array containing solid names, added to by the function - the assumption is that s_SolidNames might already contain values, which the function adds to
' lLengthOfSolidNameArrayIncr: Length of initial array, furthermore array is padded by this amount whenever necessary (concerning solids).
' lLengthOfComponentNameArrayIncr: Length of initial array, furthermore array is padded by this amount whenever necessary (concerning components).
Function ParseComponentIncludingSubcomponentsForSolids(sComponentName As String, s_SolidNames() As String, lLengthOfSolidNameArrayIncr As Long, lLengthOfComponentNameArrayIncr As Long) As Long
	Dim s_ComponentName()                As String  ' String array of components
	Dim l_ComponentParent()              As Long    ' Index of parent component
	Dim l_ComponentNumChildren()         As Long    ' Number of component children
	Dim l_ComponentNumChildrenSearched() As Long    ' Number of component children that have already been searched / catalogued
	Dim lCurrentNumComponents            As Long    ' Number of components that have been catalogued
	Dim lCurrentComponentsIndex          As Long    ' Index of current component
	Dim lCurrentParentIndex              As Long    ' Index of current parent component
	Dim lCurrentParentLevel              As Long    ' Level of current parent component
	Dim lRunningIndex01                  As Long    ' Running index
	Dim lRunningIndex02                  As Long    ' Running index
	Dim lRootNumCompChildren             As Long    ' Number of component children of "Components"
	Dim sCurrentTreeItem                 As String
	Dim bGoDownStep                      As Boolean ' Flag
	Dim bSearchingComponentChild         As Boolean ' Flag, true while child of "Components" should still be searched
	Dim lOverallNumberOfSolids           As Long    ' Overall number of solids

	' Start with lLengthOfComponentNameArrayIncr entries for components, expand as needed, similiar for solids
	ReDim s_ComponentName(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentParent(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentNumChildren(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentNumChildrenSearched(lLengthOfComponentNameArrayIncr)
	ReDim l_ComponentLevel(lLengthOfComponentNameArrayIncr)
	ReDim s_SolidNames(lLengthOfSolidNameArrayIncr)

	' Initialize
	lCurrentNumComponents = 0

	' Catalogue root component of search
	sCurrentTreeItem                  = sComponentName
	lRootNumCompChildren              = GetNumberOfComponentChildren(sCurrentTreeItem)
	s_ComponentName(0)                = sCurrentTreeItem
	l_ComponentParent(0)              = -1
	l_ComponentNumChildren(0)         = lRootNumCompChildren
	l_ComponentNumChildrenSearched(0) = 0
	l_ComponentLevel(0)               = 0
	lCurrentNumComponents             = lCurrentNumComponents + 1
	' Catalogue solids in root component of search
	lOverallNumberOfSolids = ParseComponentWithoutSubcomponentsForSolids(sComponentName, s_SolidNames, lLengthOfSolidNameArrayIncr, 0)

	For lRunningIndex01 = 1 To lRootNumCompChildren STEP 1
		' Go to lRunningIndex01 child and catalogue item
		GoToNthChild(sComponentName, sCurrentTreeItem, lRunningIndex01)
		s_ComponentName(lCurrentNumComponents)                = sCurrentTreeItem
		l_ComponentParent(lCurrentNumComponents)              = 0
		l_ComponentNumChildren(lCurrentNumComponents)         = GetNumberOfComponentChildren(sCurrentTreeItem)
		l_ComponentNumChildrenSearched(lCurrentNumComponents) = 0
		l_ComponentLevel(lCurrentNumComponents)               = 1
		lCurrentNumComponents                                 = lCurrentNumComponents + 1
		' Catalogue solids as well...
		lOverallNumberOfSolids = lOverallNumberOfSolids + ParseComponentWithoutSubcomponentsForSolids(sCurrentTreeItem, s_SolidNames, lLengthOfSolidNameArrayIncr, lOverallNumberOfSolids)
		' If component children exist, start cataloguing
		If l_ComponentNumChildren(lCurrentNumComponents-1) > 0 Then
			GoToNthChild(s_ComponentName(lCurrentNumComponents-1), sCurrentTreeItem, 1)
			l_ComponentNumChildrenSearched(lCurrentNumComponents-1) = 1
			lCurrentParentIndex                                     = lCurrentNumComponents - 1
			lCurrentParentLevel                                     = 1
			bGoDownStep = True
			bSearchingComponentChild = True
		Else
			bSearchingComponentChild = False
		End If
		While bSearchingComponentChild = True
			' First, check if size of arrays has to be expanded.
			If UBound(s_ComponentName) - 1 = lCurrentNumComponents Then
				ReDim Preserve s_ComponentName(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentParent(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentNumChildren(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentNumChildrenSearched(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
				ReDim Preserve l_ComponentLevel(lCurrentNumComponents + lLengthOfComponentNameArrayIncr)
			End If
			If bGoDownStep = True Then
			' If GoDownStep
			'    -> catalogue component as indicated by CurrentTreeItem
			'    -> If catalogued item has component children, set GoDownFlag and indicate child
			'    -> If catalogued item has no component children, set GoUpFlag and indicate parent
				s_ComponentName(lCurrentNumComponents)                = sCurrentTreeItem
				l_ComponentParent(lCurrentNumComponents)              = lCurrentParentIndex
				l_ComponentNumChildren(lCurrentNumComponents)         = GetNumberOfComponentChildren(sCurrentTreeItem)
				l_ComponentNumChildrenSearched(lCurrentNumComponents) = 0
				l_ComponentLevel(lCurrentNumComponents)               = lCurrentParentLevel + 1
				lOverallNumberOfSolids = lOverallNumberOfSolids + ParseComponentWithoutSubcomponentsForSolids(sCurrentTreeItem, s_SolidNames, lLengthOfSolidNameArrayIncr, lOverallNumberOfSolids)
				If l_ComponentNumChildren(lCurrentNumComponents) > 0 Then
					GoToNthChild(s_ComponentName(lCurrentNumComponents), sCurrentTreeItem, 1)
					l_ComponentNumChildrenSearched(lCurrentNumComponents) = 1
					lCurrentParentIndex                                   = lCurrentNumComponents
					lCurrentParentLevel                                   = l_ComponentLevel(lCurrentNumComponents)
					bGoDownStep                                           = True
				Else
					lCurrentComponentsIndex = l_ComponentParent(lCurrentNumComponents)
					bGoDownStep             = False
				End If
				' hold off on updating the number of catalogued components, until all values pertaining to the current component have been set
				lCurrentNumComponents                                 = lCurrentNumComponents + 1
			Else
			' If GoUpStep
			'    If SearchedChildren < ExistingChildren
			'       -> Set GoDownStep Flag and indicate SearchedChildren+1 child
			'    Else If arrived at level 1
			'       -> then stop / break out of while loop
			'    Else If SearchedChildren >= ExistingChildren but not at level 1
			'       -> Set GoUpStep Flag, indicating parent
				' Check whether the component still has children which have not been searched yet
				If l_ComponentNumChildrenSearched(lCurrentComponentsIndex) < l_ComponentNumChildren(lCurrentComponentsIndex) Then
					GoToNthChild(s_ComponentName(lCurrentComponentsIndex), sCurrentTreeItem, l_ComponentNumChildrenSearched(lCurrentComponentsIndex)+1)
					l_ComponentNumChildrenSearched(lCurrentComponentsIndex) = l_ComponentNumChildrenSearched(lCurrentComponentsIndex) + 1
					lCurrentParentIndex                                     = lCurrentComponentsIndex
					lCurrentParentLevel                                     = l_ComponentLevel(lCurrentComponentsIndex)
					bGoDownStep = True
				Else
					' Check whether component is at level 1 (child of "Components")
					If l_ComponentLevel(lCurrentComponentsIndex) = 1 Then
						bSearchingComponentChild = False
					Else
						sCurrentTreeItem        = s_ComponentName(l_ComponentParent(lCurrentComponentsIndex))
						lCurrentComponentsIndex = l_ComponentParent(lCurrentComponentsIndex)
						bGoDownStep = False
					End If
				End If
			End If
		Wend
	Next lRunningIndex01

	If lOverallNumberOfSolids > 0 Then
		ReDim Preserve s_SolidNames(lOverallNumberOfSolids-1)
	Else
		ReDim s_SolidNames(0)
	End If

	ParseComponentIncludingSubcomponentsForSolids = lOverallNumberOfSolids
End Function



' Accepts treeitem as string, returns number of children which are components
Function GetNumberOfComponentChildren(sPath As String) As Long
	Dim lNumCompChildren As Long
	lNumCompChildren = 0

	' Replace slashes or colons with backslashes
	sPath = Replace(sPath, "/", "\")
	sPath = Replace(sPath, ":", "\")

	' Strip trailing slashes, backslashes or colons if existing
	If Right(sPath,1) = "\" Or Right(sPath,1) = "/" Or Right(sPath,1) = ":" Then sPath = Left(sPath, Len(sPath)-1)

	' Check whether string exists as treeitem
	If Resulttree.DoesTreeItemExist(sPath) Then
		' Check that treeitem is a component
		If DetermineDesignationOfTreeItem(sPath) = "component" Then
			Dim sCurrentTreeItem As String
			sCurrentTreeItem = Resulttree.GetFirstChildName(sPath)
			If sCurrentTreeItem <> "" Then
				Dim sCurrentTreeItemIsComponent As Boolean
				If DetermineDesignationOfTreeItem(sCurrentTreeItem) = "component" Then
					sCurrentTreeItemIsComponent = True
					While sCurrentTreeItemIsComponent = True
						lNumCompChildren = lNumCompChildren + 1
						sCurrentTreeItem = Resulttree.GetNextItemName(sCurrentTreeItem)
						If DetermineDesignationOfTreeItem(sCurrentTreeItem) <> "component" Then
							sCurrentTreeItemIsComponent = False
						End If
					Wend
				Else
					lNumCompChildren = 0
				End If
			Else
				lNumCompChildren = 0
			End If
		Else
			lNumCompChildren = 0
		End If
	Else
		lNumCompChildren = 0
	End If

	GetNumberOfComponentChildren = lNumCompChildren
End Function



' Saves path of lN'th child of sParentPath in sNthChildPath. Returns True if operation was successfull and False otherwise. First child corresponds to lN = 1, not lN = 0
Function GoToNthChild(sParentPath As String, sNthChildPath As String, lN As Long) As Boolean
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



' -------------------------------------------------------------------------------------------------
' Accepts given component and transfers list of all solids in that component - not including subcomponents! - to given array.
' sComponentName: Name of component for which all solids are to be listed (subcomponents are not considered)
' s_SolidNames: Array containing solid names, added to by the function - the assumption is that s_SolidNames might already contain values, which the function adds to
' lLengthOfSolidNameArrayIncr: Length of initial (temp) array, furthermore array is padded by this amount whenever necessary.
' lCurrentNumberOfSolids: Current number of solids, s_SolidNames is extended from this index forth, regardless of initial length (if there are solids that are present, otherwise the array is not touched in any way)
' -------------------------------------------------------------------------------------------------
Function ParseComponentWithoutSubcomponentsForSolids(sComponentName As String, s_SolidNames() As String, lLengthOfSolidNameArrayIncr As Long, lCurrentNumberOfSolids As Long) As Long
	Dim lRunningIndex01      As Long    ' Running Index
	Dim s_SolidNamesTemp()   As String  ' Set up temporary string array with solid names and merge at the end
	Dim sCurrentTreeItem     As String  ' Current tree item...
	Dim lNumberOfSolids      As Long    ' Total number of solids found
	Dim lNumberOfComponents  As Long    ' Number of component children in component
	Dim bAllSolidsCatalogued As Boolean ' Flag...

	bAllSolidsCatalogued = False
	' Start with lLengthOfSolidNameArrayIncr entries, expand as needed
	ReDim s_SolidNamesTemp(lLengthOfSolidNameArrayIncr)

	' Determine number of component children of given component, then initialize sCurrentTreeItem
	lNumberOfComponents = GetNumberOfComponentChildren(sComponentName)
	If GoToNthChild(sComponentName, sCurrentTreeItem, lNumberOfComponents+1) Then
		s_SolidNamesTemp(0) = sCurrentTreeItem
		lNumberOfSolids     = 1
		While Not bAllSolidsCatalogued
			For lRunningIndex01 = 1 To lLengthOfSolidNameArrayIncr STEP 1
				sCurrentTreeItem = Resulttree.GetNextItemName(sCurrentTreeItem)
				If sCurrentTreeItem <> "" Then
					s_SolidNamesTemp(lNumberOfSolids) = sCurrentTreeItem
					lNumberOfSolids                   = lNumberOfSolids + 1
				Else
					bAllSolidsCatalogued = True
					Exit For
				End If
			Next lRunningIndex01
			' Expand size of array
			If Not bAllSolidsCatalogued Then
				ReDim Preserve s_SolidNamesTemp(UBound(s_SolidNamesTemp) + lLengthOfSolidNameArrayIncr)
			End If
		Wend
		' Merge arrays and update overall number of solids in array

		ReDim Preserve s_SolidNames(lCurrentNumberOfSolids + lNumberOfSolids - 1)
		For lRunningIndex01 = 0 To lNumberOfSolids-1 STEP 1
			s_SolidNames(lCurrentNumberOfSolids + lRunningIndex01) = s_SolidNamesTemp(lRunningIndex01)
		Next lRunningIndex01
	Else
		' The component does not contain any solids
		lNumberOfSolids = 0
	End If
	ParseComponentWithoutSubcomponentsForSolids = lNumberOfSolids
End Function


' Accepts treeitem as string, will give back designation as "component", "solid" or "other"
Function DetermineDesignationOfTreeItem(sPath As String) As String
	Dim sDesignation   As String
	Dim sPathSolidName As String

	' Replace slashes or colons with backslashes
	sPath = Replace(sPath, "/", "\")
	sPath = Replace(sPath, ":", "\")

	' Strip trailing slashes, backslashes or colons if existing
	If Right(sPath,1) = "\" Or Right(sPath,1) = "/" Or Right(sPath,1) = ":" Then sPath = Left(sPath, Len(sPath)-1)

	' Check whether string exists as treeitem
	If Resulttree.DoesTreeItemExist(sPath) Then
		' Check whether treeitem exists in the Components tree
		If Left(sPath, 11) = "Components\" Then
			' Check if treeitem has children
			If Resulttree.GetFirstChildName(sPath) <> "" Then
				sDesignation = "component"
			Else
				' Check if treeitem is a solid
				sPathSolidName = ConvertTreeItemPathToSolidName(sPath)
				If Solid.DoesExist(sPathSolidName) Then
					sDesignation = "solid"
				Else
					sDesignation = "component"
				End If
			End If
		ElseIf sPath = "Components" Then
			sDesignation = "component"
		Else
			sDesignation = "other"
		End If
	Else
		sDesignation = "other"
	End If

	DetermineDesignationOfTreeItem = sDesignation
End Function



' Takes string, copies it, replaces last backslash in copy with colon, other backslashes with slashes and removes first twelve characters.
' Thereby a valid solid treeitem is returned as a solidname which can be used in the Solid object.
' If the given string has less or equal to 12 characters, an empty string is returned.
Function ConvertTreeItemPathToSolidName(sTreeItemPath As String) As String
	Dim sPathSolidName As String
	Dim sSolidName     As String
	Dim sComponentName As String

	If Len(sTreeItemPath) > 12 Then
		' Isolate solid name and component name
		sSolidName     = Right(sTreeItemPath, Len(sTreeItemPath) - InStrRev(sTreeItemPath, "\"))
		sComponentName = Right(Left(sTreeItemPath, Len(sTreeItemPath) - Len(sSolidName) - 1), Len(sTreeItemPath) - Len(sSolidName) - 12)

		' replace backslashes by slashes and add in colon between component name and solid name
		sComponentName = Replace(sComponentName, "\", "/")
		sPathSolidName = sComponentName & ":" & sSolidName
	Else
		sPathSolidName = ""
	End If

	ConvertTreeItemPathToSolidName = sPathSolidName
End Function
