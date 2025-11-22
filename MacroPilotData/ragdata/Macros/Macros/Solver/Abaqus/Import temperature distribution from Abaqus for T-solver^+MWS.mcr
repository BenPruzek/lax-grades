'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"

' ================================================================================================
' Macro: Sets up temperature field import from Abaqus.
'
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 27-Apr-2023 mha: Made explicit that macro should only be used for T-solver
' 17-Mar-2023 mha: first version
' ================================================================================================

' *** global variables
Dim sDefaultImportName   As String  ' default name for temperature field import
Dim sFinalImportName     As String  ' final name for temperature field import
Dim sMeshFileList()      As String  ' lists possible mesh files for display in GUI
Dim sMeshFileShortName   As String  ' short name (no path) of mesh file
Dim lNumMeshFiles        As Long    ' number of possibly selected mesh files
Dim sReportFileList()    As String  ' lists possible report files for display in GUI
Dim lNumReportFiles      As Long    ' number of possibly selected mesh files
Dim sReportFileShortName As String  ' short name (no path) of report file
Dim bUseRelPath          As Boolean ' indicates whether or not to use the relative path
Dim sDirectoryPath       As String  ' path of folder where files can be found
Dim sDirectoryPathDefVal As String  ' default value folder textbox
Dim bRelPathPossible     As Boolean ' True if relative path can be defined, False otherwise
Dim sRelPathToDirectory  As String  ' relative path to folder containing data
Dim sFinalPathToDir      As String  ' path actually used for import - can be full path or relative
Dim sTempArray(2)        As String  ' string array, contains temperature units
Dim sGeometryArray(7)    As String  ' string array containing geometry units
Dim sGeometryUnit        As String  ' string, containing user selected geometry unit
Dim sTemperatureUnit     As String  ' string, containing user selected temperature unit
Dim bTSolver             As Boolean ' True if T-solver is selected, False otherwise
Dim bContinue            As Boolean ' True if user wants to continue with execution of this macro

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality. Clear the data in order to
	' provide well defined environment for testing.
	ActivateScriptSettings True
	ClearScriptSettings
	DS.ClearScriptSettings

	' Initialize global variables
	sDefaultImportName   = "Abaqus-Temp-Distribution"
	bUseRelPath          = False
	bRelPathPossible     = True
	sDirectoryPathDefVal = "No file selected yet..."
	ReDim sMeshFileList(0)
	ReDim sReportFileList(0)
	sMeshFileList(0)     = sDirectoryPathDefVal
	sReportFileList(0)   = sDirectoryPathDefVal
	lNumMeshFiles        = 0
	lNumReportFiles      = 0
	sTempArray(0)        = "Celsius"
	sTempArray(1)        = "Kelvin"
	sTempArray(2)        = "Fahrenheit"
	sGeometryArray(0)    = "m"
	sGeometryArray(1)    = "cm"
	sGeometryArray(2)    = "mm"
	sGeometryArray(3)    = "um"
	sGeometryArray(4)    = "nm"
	sGeometryArray(5)    = "ft"
	sGeometryArray(6)    = "in"
	sGeometryArray(7)    = "mil"
	bContinue            = True

	' Check solver type
	Dim sCurrentSolver As String
	sCurrentSolver = GetSolverType
	If LCase(sCurrentSolver) = LCase("HF Time Domain") Then
		bTSolver = True
	Else
		bTSolver = False
	End If

	' Call the define method and check whether it is completed successfully
	If (Define("test", True, False)) Then
		If bContinue Then
			sFinalPathToDir = IIf(bRelPathPossible And bUseRelPath, sRelPathToDirectory, sDirectoryPath)
			FieldSource.CreateFieldImportFromAbaqus(sFinalImportName, sFinalPathToDir, sMeshFileShortName, sReportFileShortName,  sGeometryUnit, sTemperatureUnit)
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean) As Boolean
	Begin Dialog UserDialog 940, 430, "Create temperature field import from Abaqus for T-solver", .DialogFunc ' %GRID:10,7,1,1
		GroupBox      10,  10, 920,  55, "Specify name for the temperature field import",               .GBSpecifyName
		Text          20,  37,  50,  14, "Name:",                                                       .TName
		TextBox       90,  34, 830,  21,                                                                .TBName

		GroupBox      10,  70, 920, 216, "Specify files for temperature import",                        .GBSpecifyFiles
		Text          20,  97,  70,  14, "Directory:",                                                  .TDirectory
		TextBox       90,  94, 830,  21,                                                                .TBDirectory
		CheckBox      20, 125, 130,  14, "Use relative path",                                           .CBUseRelPath
		PushButton    20, 150,  95,  21, "Browse file...",                                              .PBSelectFile
		Text         130, 154, 320,  14, "Specify Nastran file: mesh output, "".nas"" or "".inp""",     .TExplainNastranFile
		ListBox       20, 178, 435, 100, sMeshFileList(),                                               .LBMeshFileNames
		Text         485, 154, 320,  14, "Specify temperature field report file: "".rpt""",             .TExplainReportFile
		ListBox      485, 178, 435, 100, sReportFileList(),                                             .LBReportFileNames

		GroupBox      10, 296, 920,  44, "Specify units for temperature import",                        .GBSpecifyUnits
		Text          20, 315, 296,  14, "Geometry  unit (for interpreting the Nastran file):",         .TGeoUnit
		DropListBox  330, 311,  60,  21, sGeometryArray(),                                              .DLBGeometryArray
		Text         485, 315, 302,  14, "Temperature  unit (for interpreting the report file):",       .TTempUnit
		DropListBox  800, 311, 110,  21, sTempArray(),                                                  .DLBTempArray

		GroupBox      10, 345, 920,  53, "Please note:",                                                .GBNote
		Text          20, 360, 900,  14, "The mesh file ("".nas"", "".inp"") as well as the report" & _
                                         " file ("".rpt"") should both be saved in the same parent" & _
		                                 " folder.",                                                    .TPleaseNote1
		Text          20, 376, 900,  14, "The field import which is generated can currently only be" & _
                                         " utilized together with the T-solver.",                       .TPleaseNote2

		OKButton      15, 403,  90,  21
		CancelButton 123, 403,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg           As UserDialog
	Dim sImportName   As String
	Dim sDirectory    As String
	Dim sNastranFile  As String
	Dim sTempDist     As String
	Dim sGeometryUnit As String
	Dim sTempUnit     As String

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		' Store the script settings into the database for later reuse by either the define function (for modifications) or the evaluate function.
	End If
End Function



Rem See DialogFunc help topic for more information.
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim sValueSet         As String
	Dim sNoForbiddenChars As String
	Dim lRunningIndex     As Long

	Select Case iAction
	Case 1 ' Dialog box initialization
		DlgText   "TBName",             sDefaultImportName
		DlgEnable "TBDirectory",        False
		DlgText   "TBDirectory",        sDirectoryPathDefVal
		DlgEnable "PBSelectFile",       True
		DlgEnable "CBUseRelPath",       bRelPathPossible
		DlgValue  "CBUseRelPath",       IIf(bUseRelPath, 1, 0)
		DlgEnable "LBMeshFileNames",    False
		DlgValue  "LBMeshFileNames",    -1
		DlgEnable "LBReportFileNames",  False
		DlgValue  "LBReportFileNames",  -1
		DlgEnable "DLBGeometryArray",   False
		DlgValue  "DLBGeometryArray",   2
		DlgEnable "DLBTempArray",       False
		DlgValue  "DLBTempArray",       0
		DlgEnable "CBUseRelPath",       False
		DlgValue  "CBUseRelPath",       0
		sFinalImportName = sDefaultImportName
		sGeometryUnit    = sGeometryArray(DlgValue("DLBGeometryArray"))
		sTemperatureUnit = sTempArray(DlgValue("DLBTempArray"))

	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		If ( sDlgItem = "PBSelectFile" ) Then ' "Browse file..." has been pressed
			Dim sExtensions    As String
			Dim sFileName      As String
			sExtensions = "nas;inp"

			sFileName          = GetFilePath("", sExtensions, GetProjectPath("Root"), "Select mesh file ("".nas"" or "".inp"")", 0)
			sMeshFileShortName = ShortName(sFileName)

			If (sFileName <> "") Then
				sDirectoryPath = DirName(sFileName)
				' Update textboxes
				DlgEnable "TBDirectory", False
				DlgText   "TBDirectory", sDirectoryPath
				' Update listbox / sFileList
				lNumMeshFiles   = GetFilesInDirectory(sDirectoryPath, ".nas", sMeshFileList, 0)
				lNumMeshFiles   = lNumMeshFiles + GetFilesInDirectory(sDirectoryPath, ".inp", sMeshFileList, lNumMeshFiles)
				lNumReportFiles = GetFilesInDirectory(sDirectoryPath, ".rpt", sReportFileList, 0)
				' Go through list and shorten down to actual file name:
				For lRunningIndex = 0 To UBound(sMeshFileList) STEP 1
					sMeshFileList(lRunningIndex) = ShortName(sMeshFileList(lRunningIndex))
				Next lRunningIndex
				For lRunningIndex = 0 To UBound(sReportFileList) STEP 1
					sReportFileList(lRunningIndex) = ShortName(sReportFileList(lRunningIndex))
				Next lRunningIndex
				' Determine relative path and assign values to global variables bRelPathPossible and sRelPathToSigDir
				DetermineRelativePathToFolder()
				' Update listboxes
				DlgListBoxArray "LBMeshFileNames",   sMeshFileList
				DlgEnable       "LBMeshFileNames",   True
				DlgListBoxArray "LBReportFileNames", sReportFileList
				DlgEnable       "LBReportFileNames", True
				' Select chosen file in mesh listbox
				For lRunningIndex = 0 To UBound(sMeshFileList) STEP 1
					If sMeshFileShortName = sMeshFileList(lRunningIndex) Then
						DlgValue "LBMeshFileNames", lRunningIndex
						Exit For
					End If
				Next lRunningIndex
				' Update relative path check box
				If bRelPathPossible Then
					DlgEnable "CBUseRelPath", True
				Else
					DlgEnable "CBUseRelPath", False
					DlgValue  "CBUseRelPath", 0
				End If
				DlgEnable "DLBGeometryArray", True
				DlgEnable "DLBTempArray",     True
			Else
				sDirectoryPath = sDirectoryPathDefVal
				DlgEnable "TBDirectory", False
				DlgText   "TBDirectory", sDirectoryPathDefVal
				ReDim sMeshFileList(0)
				sMeshFileList(0) = sDirectoryPathDefVal
				ReDim sReportFileList(0)
				sReportFileList(0) = sDirectoryPathDefVal
				DlgListBoxArray "LBMeshFileNames",   sMeshFileList
				DlgEnable       "LBMeshFileNames",   False
				DlgValue        "LBMeshFileNames",   -1
				DlgListBoxArray "LBReportFileNames", sReportFileList
				DlgEnable       "LBReportFileNames", False
				DlgValue        "LBReportFileNames", -1
				DlgEnable       "DLBGeometryArray",  False
				DlgEnable       "DLBTempArray",      False
				DlgEnable       "CBUseRelPath",      False
				DlgValue        "CBUseRelPath",      0
				lNumMeshFiles   = 0
				lNumReportFiles = 0
			End If
			' keep dialog open
			DialogFunc = True

		ElseIf ( sDlgItem = "CBUseRelPath" ) Then
			bUseRelPath  = CBool(DlgValue("CBUseRelPath"))
			If bUseRelPath Then
				DlgEnable "TBDirectory", False
				DlgText   "TBDirectory", sRelPathToDirectory
			Else
				DlgEnable "TBDirectory", False
				DlgText   "TBDirectory", sDirectoryPath
			End If

		ElseIf ( sDlgItem = "OK" ) Then
			' First, check if all basic necessities are set...
			If ( lNumMeshFiles <= 0 ) Or ( lNumReportFiles <= 0 ) Or ( DlgValue("LBMeshFileNames") < 0 ) Or ( DlgValue("LBReportFileNames") < 0 ) Then
				If lNumMeshFiles <= 0 Then
					MsgBox "Currently, no mesh files are present - to successfully set up the field import both a mesh file and a report file must be selected.", vbExclamation
					DialogFunc = True
				ElseIf lNumReportFiles <= 0 Then
					MsgBox "Currently, no report files are present - to successfully set up the field import both a mesh file and a report file must be selected.", vbExclamation
					DialogFunc = True
				ElseIf DlgValue("LBMeshFileNames") < 0 Then
					MsgBox "Currently, no mesh file is selected - to successfully set up the field import both a mesh file and a report file must be selected.", vbExclamation
					DialogFunc = True
				ElseIf DlgValue("LBReportFileNames") < 0 Then
					MsgBox "Currently, no report file is selected - to successfully set up the field import both a mesh file and a report file must be selected.", vbExclamation
					DialogFunc = True
				End If
			' Next, check if the T-solver is active - if not, then query if the user wants to continue
			Else
				If bTSolver Then
					bContinue = True
				Else
					bContinue = AreYouSureYouWantToContinue()
				End If
			End If

		ElseIf ( sDlgItem = "LBMeshFileNames" ) Then
			sMeshFileShortName = sMeshFileList(DlgValue("LBMeshFileNames"))

		ElseIf ( sDlgItem = "LBReportFileNames" ) Then
			sReportFileShortName = sReportFileList(DlgValue("LBReportFileNames"))

		ElseIf ( sDlgItem = "DLBGeometryArray" ) Then
			sGeometryUnit = sGeometryArray(DlgValue("DLBGeometryArray"))

		ElseIf ( sDlgItem = "DLBTempArray" ) Then
			sTemperatureUnit = sTempArray(DlgValue("DLBTempArray"))
		End If

	Case 3 ' TextBox or ComboBox text changed
		If ( sDlgItem = "TBName" ) Then
			sValueSet    = DlgText(sDlgItem)
			If sValueSet = "" Then
				DlgText sDlgItem, sDefaultImportName
				MsgBox "You seem to have entered an empty string for the name of the field import. As such the name has been reset to the default setting.", vbExclamation
			Else
				' Remove any forbidden characters and eliminate leading "."
				sNoForbiddenChars = NoForbiddenFilenameCharacters(sValueSet)
				sNoForbiddenChars = LTrim(sValueSet)
				sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
				If sNoForbiddenChars = "" Then
					DlgText sDlgItem, sDefaultImportName
					MsgBox "Please be aware that the monitor name you have entered consisted only of characters not allowed to be used or resulted in leading periods. As such, the entry has been reduced to an empty string and been replaced with the default value.", vbExclamation
				ElseIf sNoForbiddenChars <> sValueSet Then
					DlgText sDlgItem, sNoForbiddenChars
					MsgBox "Please be aware that the monitor name you have entered contained some characters not allowed to be used. As such, the entry has been slightly altered.", vbExclamation
				End If
			End If
			sFinalImportName = DlgText("TBName")
		End If

	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
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



' -------------------------------------------------------------------------------------------------
' GetFilesInDirectory: This function returns the number of files in the folder which correspond to
'                      the filter sFilter.
'                      The files will be added to string array sFileNames
'                      lInitPos specifies the position where the first entry should be made
' -------------------------------------------------------------------------------------------------
Private Function GetFilesInDirectory(sPath As String, sFilter As String, sFileNames() As String, lInitPos As Long) As Long
	' Declar string array for file names
	Dim sFile        As String
	Dim lCounter     As Long
	Dim lPadding     As Long

	' Initialize
	lCounter = 0
	lPadding = 10
	ReDim Preserve sFileNames(lInitPos + lPadding)

	' Make sure folder path is terminated by path seperator
	If Right(sPath, 1) <> "\" Then
		sPath = sPath & "\"
	End If

	' Initialize sFilter if empty
	If sFilter = "" Then
		sFilter = "*.*"
	ElseIf Not Left(sFilter, 1) = "*" Then
		sFilter = "*" & sFilter
	End If

	' Call with path initializes the dir function and returns the first file
	sFile = Dir(sPath & sFilter)

	' Call dir function until return value is empty
	Do While sFile <> ""
		' Store file name, then go to next file name
		sFileNames(lInitPos + lCounter) = sFile
		lCounter                        = lCounter + 1
        sFile = Dir

		' Pad array if necessary
		If lCounter = UBound(sFileNames) Then
			ReDim Preserve sFileNames(UBound(sFileNames) + lPadding)
		End If
	Loop

	' Cut array down to size
	If lCounter > 0 Then
		ReDim Preserve sFileNames(lInitPos + lCounter-1)
    End If

    GetFilesInDirectory = lCounter
End Function



' --------------------------------------------------------------------------------
' DetermineRelativePathToFolder
' Determines relative path to folder containing files (given by global string
' variable sDirectoryPath), assignes relative path to global string variable
' sRelPathToDirectory and sets global Boolean bRelPathPossible
' --------------------------------------------------------------------------------
Private Function DetermineRelativePathToFolder() As Boolean
	Dim sFullPath        As String
	Dim sPathProjDir     As String
	Dim sFullPathAr()    As String
	Dim sPathProjDirAr() As String
	Dim sPath            As String ' build up relative path...
	Dim lRunningIndex    As Long

	' basic paths
	sFullPath    = sDirectoryPath
	sPathProjDir = GetProjectPath("Project")

	' convert to array, separating by path seperator
	sFullPathAr()  = Split(Left(sFullPath, Len(sFullPath)-1), "\")
	sPathProjDirAr = Split(sPathProjDir, "\")
	If sFullPathAr(0) <> sPathProjDirAr(0) Then
		' no relative path possible...
		bRelPathPossible    = False
		sRelPathToDirectory = sDirectoryPath
	Else
		' relative path possible, eliminate all entries which are identical from root of path onward
		Do While ( ( UBound(sFullPathAr) > 0 ) And ( UBound(sPathProjDirAr) > 0 ) And ( sFullPathAr(0) = sPathProjDirAr(0) ) )
			RemoveEntryFromStringArrayAtGivenIndex(sFullPathAr, 0)
			RemoveEntryFromStringArrayAtGivenIndex(sPathProjDirAr, 0)
		Loop
		' at this point there is still at least one entry left in both arrays and the first entry might be identical or not..
		' to get the relative path it is necessary to go up the tree from the project directory, then down again the remaining path to the file...
		sPath = ""
		If sFullPathAr(0) = sPathProjDirAr(0) Then
			For lRunningIndex = 1 To UBound(sPathProjDirAr) STEP 1
				sPath = sPath & "..\"
			Next lRunningIndex
			For lRunningIndex = 1 To UBound(sFullPathAr) STEP 1
				sPath = sPath & sFullPathAr(lRunningIndex) & "\"
			Next lRunningIndex
		Else
			For lRunningIndex = 1 To UBound(sPathProjDirAr) STEP 1
				sPath = sPath & "..\"
			Next lRunningIndex
				sPath = sPath & "..\"
			For lRunningIndex = 0 To UBound(sFullPathAr) STEP 1
				sPath = sPath & sFullPathAr(lRunningIndex) & "\"
			Next lRunningIndex
		End If
		bRelPathPossible    = True
		sRelPathToDirectory = sPath
	End If

	DetermineRelativePathToFolder = bRelPathPossible
End Function



' --------------------------------------------------------------------------------
' RemoveEntryFromStringArrayAtGivenIndex
' Removes entry at position lIndex from string array sStringArray unless the array
' contains only one element.
' --------------------------------------------------------------------------------
Private Function RemoveEntryFromStringArrayAtGivenIndex(sStringArray() As String, lIndex As Long) As Boolean
	Dim lRunningIndex01 As Long
	Dim bReturnValue    As Boolean

	' Check that index is not too large...
	If lIndex > UBound(sStringArray) Then
		bReturnValue = False
	Else
		' Check if array contains more then one element
		If UBound(sStringArray) > 0 Then
			For lRunningIndex01 = lIndex+1 To  UBound(sStringArray) STEP 1
				sStringArray(lRunningIndex01-1) = sStringArray(lRunningIndex01)
			Next lRunningIndex01
			ReDim Preserve sStringArray(UBound(sStringArray)-1)
		Else
			bReturnValue = False
		End If
	End If

	RemoveEntryFromStringArrayAtGivenIndex = bReturnValue
End Function



' This function is called if a solver other then the T-solver is active.
Private Function AreYouSureYouWantToContinue() As Boolean
	Dim vReturn As Variant
	vReturn = MsgBox "Please note that this temperature field import can only be used together with the T-solver." & _
	vbNewLine & _
	"Do you still want to continue with the temperature import?", vbYesNo + _
	vbQuestion, "Utilizability of temperature import from Abaqus"

	If vReturn = vbYes Then
		AreYouSureYouWantToContinue = True
	Else
		AreYouSureYouWantToContinue = False
	End If
End Function
