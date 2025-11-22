'#Language "WWB-COM"

' This macro imports a SPICE netlist and generates a SPICE block, representing a subcircuit, from it on the schematic for transient simulation.
' In addition, a default transient task is generated, accommodating a simulation options file which is needed for trasnient simulation.
' The user specifies a subcircuit name and terminals. In addition, he specifies, which statements of the netlist are put into the subcircuit
' and which statements are put into a simulation options file via filters.
' Finally the Spice format needs to be specified by the user.
' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 09-Dec-2020 cri: Created first version. No SPICE block is generated yet. Instead, the options.sp, the circuit.sp, and the logging.txt files are written into the directory of the specified netlist file.
' 05-Jan-2021 cri: Add check for unique block name and use unique block name in netlist segmentation
' 06-Jan-2021 cri: Generate SPICE block on schematic when executing macro.
' 07-Jan-2021 cri: Generate transient task and write simulation options file into task model folder after generating transient task.
' 11-Jan-2021 cri: Add message window output and copy logfile to block model folder.
' 12-Jan-2021 cri: Implement help page and link it to the Help button.
' 13-Jan-2021 cri: Adapted path to the helppage of this macro.
' 14-Jan-2021 cri: Improve message output on block and task generation.
' 29-Mar-2022 cri: Replace undocumented VBA application command by an equivalent, documented ReportInformationToWindow()
' 30-Apr-2024 cri: Support LTSPICE as Spice dialect 
' ---------------------------------------------------------------------------------------------------------------------------------
Option Explicit

Dim sPythonScript As String, sCompleteShellCommand
Dim sPythonExe As String
Const macroPath = "\Library\Macros\Construct\Miscellaneous\"
Const CSTStudioAppdataPath = "\DASSAULT_SYSTEMES\CSTStudioSuite\"
Const optionsFilterConfigFile = "OptionsPositiveFilter.cfg"
Const circuitFilterConfigFile = "CircuitNegativeFilter.cfg"

Public Function GetGlobalOptionsFilterConfigFile() As String
	GetGlobalOptionsFilterConfigFile = GetInstallPath + macroPath + optionsFilterConfigFile
End Function

Public Function GetGlobalCircuitFilterConfigFile() As String
	GetGlobalCircuitFilterConfigFile = GetInstallPath + macroPath + circuitFilterConfigFile
End Function

Public Function GetLocalOptionsFilterConfigFile() As String
	Dim strAppData As String
	strAppData = Environ("LOCALAPPDATA")
	Dim ret As String
	ret = strAppData + CSTStudioAppdataPath + optionsFilterConfigFile
	GetLocalOptionsFilterConfigFile = ret
End Function

Public Function GetLocalCircuitFilterConfigFile() As String
	Dim strAppData As String
	strAppData = Environ("LOCALAPPDATA")
	Dim ret As String
	ret = strAppData + CSTStudioAppdataPath + circuitFilterConfigFile
	GetLocalCircuitFilterConfigFile = ret
End Function

Public Function FileExists(ByVal s As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(s)
End Function

Public Function CreateFolder(ByVal folder As String) As Boolean
	CreateFolder = True
	If (CreateObject("Scripting.FileSystemObject").FolderExists(folder) = True) Then
	    Exit Function
	Else
	    On Error GoTo ErrorHandling
		If (CreateObject("Scripting.FileSystemObject").CreateFolder(folder) = True) Then
	    	Exit Function
	    End If
	End If
	ErrorHandling:
	    CreateFolder = False
End Function

Public Function CreateFolderFullPath(ByVal folder As String) As Boolean
	Dim subPath As String
	CreateFolderFullPath = True
	If (CreateObject("Scripting.FileSystemObject").FolderExists(folder) = False) Then
		subPath = folder
		While (CreateFolder(subPath) = False)
			If (subPath = "") Then
				CreateFolderFullPath = False
				Exit Function
			End If
	        subPath = GetFolderFromPath(Left(subPath, Len(subPath$)-1))
		Wend
		Return CreateFolderFullPath(folder)
	End If
	CreateFolderFullPath = True
End Function

Public Function AreOnSameDrive(ByVal path1 As String, ByVal path2 As String) As Boolean
	Dim drive1 As String
	Dim drive2 As String
	drive1 = Left(path1, InStr(path1, "\"))
	drive2 = Left(path2, InStr(path2, "\"))
	AreOnSameDrive = (drive1 = drive2)
End Function

Public Function GetOptionsFilterConfigFile() As String
    Dim ret As String
	ret = GetLocalOptionsFilterConfigFile()
	If FileExists(ret) = False Then
		ret = GetGlobalOptionsFilterConfigFile()
		If FileExists(ret) = False Then
			ret = ""
		End If
	End If
	GetOptionsFilterConfigFile = ret
End Function

Public Function GetCircuitFilterConfigFile() As String
	Dim ret As String
	ret = GetLocalCircuitFilterConfigFile()
	If FileExists(ret) = False Then
		ret = GetGlobalCircuitFilterConfigFile()
		If FileExists(ret) = False Then
			ret = ""
		End If
	End If
	GetCircuitFilterConfigFile = ret
End Function

Public Function GetFolderFromPath(ByVal strFullPath As String) As String
    GetFolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function

Public Function BrowseNetlistFile() As String
	BrowseNetlistFile = GetFilePath("*.sp; *.cir; *.net; *.txt", "SPICE files|*.*", , "Please select SPICE file to be imported", 0)
End Function

Public Function WriteToFile(ByVal strFilename As String, ByVal s As String)
	Dim iFile As Integer
	iFile = FreeFile
	Open strFilename For Output As #iFile
	Print #iFile,s$
	Close #iFile
End Function

Public Function ReadFromFile(strFilename) As String
	Dim strFileContent As String
	Dim iFile As Integer
	iFile = FreeFile
	Open strFilename For Input As #iFile
	strFileContent = Input(LOF(iFile), iFile)
	Close #iFile
	ReadFromFile = strFileContent
End Function


Public Function IsUniqueName(ByVal blockName As String, ByVal nBlocks As Long, ByVal BlockNameArray() As String) As Boolean
	Dim nIndex As Long
	For nIndex = 0 To nBlocks-1
		If (BlockNameArray(nIndex) = blockName) Then
			IsUniqueName = False
			Exit Function
		End If
	Next nIndex
	IsUniqueName = True
End Function

Public Function GetUniqueBlockName(ByVal blockName As String) As String
	Dim nBlocks As Long, nIndex As Long
	nBlocks = Block.StartBlockNameIteration()
	Dim BlockNameArray() As String
	ReDim BlockNameArray(nBlocks)
	Dim strBlockName As String
	With Block
		.Enable3DCommands(False)
	End With
	For nIndex = 0 To nBlocks-1
		strBlockName = Block.GetNextBlockName()
		BlockNameArray(nIndex) = strBlockName
	Next nIndex
	Dim ret As String
	ret = blockName
	Dim cnt As Integer
	cnt = 1
	While (IsUniqueName(ret, nBlocks, BlockNameArray) = False)
		ret = ret & cnt
		cnt = cnt + 1
	Wend
    GetUniqueBlockName = ret
End Function

Public Function GetUniqueTaskName(ByVal taskName As String) As String
	Dim nTasks As Long, nIndex As Long
	nTasks = SimulationTask.StartTaskNameIteration()
	Dim TaskNameArray() As String
	ReDim TaskNameArray(nTasks)
	Dim strTaskName As String
	For nIndex = 0 To nTasks-1
		strTaskName = SimulationTask.GetNextTaskName()
		TaskNameArray(nIndex) = strTaskName
	Next nIndex
	Dim ret As String
	ret = taskName
	Dim cnt As Integer
	cnt = 1
	While (IsUniqueName(ret, nTasks, TaskNameArray) = False)
		ret = ret & cnt
		cnt = cnt + 1
	Wend
    GetUniqueTaskName = ret
End Function

Public Function GetUniqueFileName(ByVal dirPath As String, ByVal fileName As String) As String
	Dim fname As String, body As String, extension As String
	Dim pos As Variant
	pos = InStrRev(fileName, ".")
	If (pos = 0) Then
		body = fileName
		extension = ""
	Else
		body = Mid$(fileName, 1, pos-1)
		extension = Mid$(fileName, pos)
	End If
	fname = body + extension
	Dim cnt As Integer
	cnt = 1
	While (FileExists(dirPath + fname) = True)
		fname = (body & cnt) + extension
		cnt = cnt + 1
	Wend
	GetUniqueFileName = dirPath + fname
End Function

Public Function CreateSpiceBlock(ByVal blockName As String, ByVal circuitFileName As String, ByVal spiceDialect As String) As String
	With Block
		.Reset
		.Type ("SPICEImport")
		.SetFile(circuitFileName)
		.SetRelativePath(True)
		Dim uniqueBlockName As String
		uniqueBlockName = GetUniqueBlockName(blockName)
		.Name (uniqueBlockName)
		.Create
		If (spiceDialect = "SPICE3f5") Then
			.SetSpiceFilter("SPICE3")
		ElseIf (spiceDialect = "Combined") Then
			.SetSpiceFilter("AUTOMATIC")
		Else
			.SetSpiceFilter(spiceDialect)
		End If
		.ConvertToProjectFileBlock()
		DS.ZoomToFit()
		CreateSpiceBlock = uniqueBlockName
		Exit Function
	End With
	CreateSpiceBlock = ""
End Function

Public Function CreateTransientTask(ByVal taskName As String, ByVal optionsFileContents As String, ByVal spiceDialect As String) As String
	With SimulationTask
		Dim uniqueTaskName As String
		uniqueTaskName = GetUniqueTaskName(taskName)
		.Reset
		.Name (uniqueTaskName)
		.Type ("transient")
		.Create
		Dim taskFolder As String, optionsFileName As String
		taskFolder = .GetModelFolder(True)
		optionsFileName = taskFolder + "\options.sp"
		If (spiceDialect = "HSPICE") Then
			.SetProperty("circuit simulator", "hspice")
		Else
			.SetProperty("circuit simulator", "ltspice")
		End If
		If (CreateFolderFullPath(taskFolder)) Then
			WriteToFile(optionsFileName, optionsFileContents)
		Else
			MsgBox("Options file could not be written.")
			CreateTransientTask = ""
			Exit Function
		End If
		CreateTransientTask = uniqueTaskName
		Exit Function
	End With
	CreateTransientTask = ""
End Function



Public Function ShowLogfileContents(ByVal logFileName As String, ByVal blockName As String)
	If (FileExists(logFileName) = True) Then
		With Block
			.Reset
			.Name (blockName)
			' Copy log file to block model folder
			Dim blockFolder As String, targetName As String
			blockFolder = .GetModelFolder()
			CreateFolderFullPath(blockFolder)
			targetName = blockFolder + "\logging.txt"
			Dim logFileContents As String
			logFileContents = ReadFromFile(logFileName)
			WriteToFile(targetName,logFileContents)
			' Write logfile contents to message window
			DS.ReportInformationToWindow(logFileContents)
		End With
	End If
End Function

Public Function OnOK(ByVal netlistFileName As String, ByVal blockName As String, ByVal nodes As String, ByVal optionFilter As String, ByVal circuitFilter As String, ByVal dialect As String)
	Dim dirPath As String, circuitFileName As String, optionsFileName As String, logFileName As String
	dirPath = GetFolderFromPath(netlistFileName)
	circuitFileName = GetUniqueFileName(dirPath, "circuit.sp")
	optionsFileName = GetUniqueFileName(dirPath, "options.sp")
	logFileName = GetUniqueFileName(dirPath, "logging.txt")
	DS.ReportInformationToWindow("Starting netlist segmentation.")
	Dim projectPath As String
	projectPath = GetProjectPath("Root")
	If (AreOnSameDrive(netlistFileName, projectPath) = False) Then
		MsgBox("SPICE block creation failed. The project and the netlist file have to be on the same drive.")
		Exit Function
	End If
	DS.NetlistSegmentation(dialect, netlistFileName, blockName, nodes, optionFilter, circuitFilter, optionsFileName, circuitFileName, logFileName)
	DS.ReportInformationToWindow("Netlist " + netlistFileName + " has been successfully segmented.")
	Dim uniqueBlockName As String
	uniqueBlockName = CreateSpiceBlock(blockName, circuitFileName, dialect)
	DS.ReportInformationToWindow("SPICE block " + uniqueBlockName + " has been successfully generated. Its content can be viewed and modified via the Edit... command of its context menu.")
	Dim optionsFileContents As String
	optionsFileContents = ReadFromFile(optionsFileName)
	Dim uniqueTaskName As String
	uniqueTaskName = CreateTransientTask("Tran1_"+blockName, optionsFileContents, dialect)

	Dim absFileName As String
	With SimulationTask
		.Reset
		.Name (uniqueTaskName)
		Dim taskFolder As String
		taskFolder = .GetModelFolder(True)
		absFileName = taskFolder + "\options.sp"
	End With
	DS.ReportInformationToWindow("Transient task with simulation options " + uniqueTaskName + " has been successfully generated. The simulation options file is located at " + absFileName)
	ShowLogfileContents(logFileName, uniqueBlockName)
End Function


Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses Any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
	Dim optionsFilterContents As String
	Dim localOptionsFilterConfigFile As String
	Dim globalOptionsFilterConfigFile As String
	Dim circuitFilterContents As String
	Dim localCircuitFilterConfigFile As String
	Dim globalCircuitFilterConfigFile As String
	Dim fileCreated As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable ("filename", False)
		DlgEnable ("optionFilter", False)
		DlgEnable ("circuitFilter", False)
    Case 2 ' Value changing or button pressed'
    	DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
		Case "browse"
 			DlgText("fileName", BrowseNetlistFile())
 		Case "editOption"
			optionsFilterContents = DlgText("optionFilter")
			localOptionsFilterConfigFile = GetLocalOptionsFilterConfigFile()
			fileCreated = False
			If (FileExists(localOptionsFilterConfigFile) = False) Then
				If (CreateFolderFullPath(GetFolderFromPath(localOptionsFilterConfigFile)) = True) Then
					fileCreated = True
					WriteToFile(localOptionsFilterConfigFile, optionsFilterContents)
				End If
			Else
				fileCreated = True
				WriteToFile(localOptionsFilterConfigFile, optionsFilterContents)
			End If
			If (fileCreated = True) Then
				optionsFilterContents = DS.EditTextFile(localOptionsFilterConfigFile, "Options positive filter")
				DlgText("optionFilter", optionsFilterContents)
			End If
		Case "editCircuit"
			circuitFilterContents = DlgText("circuitFilter")
			localCircuitFilterConfigFile = GetLocalCircuitFilterConfigFile()
			fileCreated = False
			If (FileExists(localCircuitFilterConfigFile) = False) Then
				If (CreateFolderFullPath(GetFolderFromPath(localCircuitFilterConfigFile)) = True) Then
					fileCreated = True
					WriteToFile(localCircuitFilterConfigFile, circuitFilterContents)
				End If
			Else
				fileCreated = True
				WriteToFile(localCircuitFilterConfigFile, circuitFilterContents)
			End If
			If (fileCreated = True) Then
				circuitFilterContents = DS.EditTextFile(localCircuitFilterConfigFile, "Circuit negative filter")
				DlgText("circuitFilter", circuitFilterContents)
			End If
		Case "resetFilters"
			globalOptionsFilterConfigFile = GetGlobalOptionsFilterConfigFile()
			If (FileExists(globalOptionsFilterConfigFile) = True) Then
				optionsFilterContents = ReadFromFile(globalOptionsFilterConfigFile)
				localOptionsFilterConfigFile = GetLocalOptionsFilterConfigFile()
				If (FileExists(localOptionsFilterConfigFile) = True) Then
					WriteToFile(localOptionsFilterConfigFile, optionsFilterContents)
				End If
				DlgText("optionFilter", optionsFilterContents)
			End If
			globalCircuitFilterConfigFile = GetGlobalCircuitFilterConfigFile()
			If (FileExists(globalCircuitFilterConfigFile) = True) Then
				circuitFilterContents = ReadFromFile(globalCircuitFilterConfigFile)
				localCircuitFilterConfigFile = GetLocalCircuitFilterConfigFile()
				If (FileExists(localCircuitFilterConfigFile) = True) Then
					WriteToFile(localCircuitFilterConfigFile, circuitFilterContents)
				End If
				DlgText("circuitFilter", circuitFilterContents)
			End If
		Case "OKButtonPushed"
			OnOK(DlgText("filename"), DlgText("blockname"), DlgText("nodes"), DlgText("optionFilter"), DlgText("circuitFilter"), DlgText("dialect"))
			Exit All
		Case "HelpButtonPushed"
			StartDESHelp "macro\common_macro_netlistsegmentation"
		Case "CancelButtonPushed"
			Exit All
	End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    	DialogFunc = True ' Prevent button press from closing the dialog box

    Case 6 ' Function key
    End Select
End Function

Sub Main()
    Dim lists$(4)
    lists$(0) = "SPICE3f5"
    lists$(1) = "PSPICE"
	lists$(2) = "LTSPICE"
    lists$(3) = "HSPICE"
    lists$(4) = "Combined"
    Begin Dialog UserDialog 520,200, "Construct SPICE block from netlist", .DialogFunc
        PushButton 10,15,80,20,"&Browse...", .browse
        PushButton 100,15,100,20,"Reset filters", .resetFilters
        Text 10,40,270,15,"File name", .Text3
        TextBox 95,40,405,15, .filename
        Text 10,60,270,15,"Block name", .Text4
        TextBox 95,60,405,15, .blockname
        Text 10,80,280,15,"Nodes of interest"
        TextBox 125,80,375,15, .nodes
        Text 10,100,280,15,"Options positive filter"
        TextBox 145,100,285,15, .optionFilter
        PushButton 450,100,53,15,"&Edit...", .editOption
        Text 10,120,280,15,"Circuit negative filter"
        TextBox 145,120,285,15, .circuitFilter
        PushButton 450,120,53,15,"&Edit...", .editCircuit
		Text 10,140,170,15,"Spice format"
      	DropListBox 145,140,100,15,lists$(),.dialect
        PushButton 30,175,60,20, "&OK", .OKButtonPushed
        PushButton 100,175,60,20, "&Cancel", .CancelButtonPushed
        PushButton 170,175,60,20, "&Help",.HelpButtonPushed
    End Dialog
    Dim dlg As UserDialog
    dlg.nodes = ""
    dlg.filename = ""
    dlg.blockname = "MySubckt"
    dlg.optionFilter = ""
    Dim optionsFilterConfigFile As String
    optionsFilterConfigFile = GetOptionsFilterConfigFile()
    If (optionsFilterConfigFile <> "") Then
    	dlg.optionFilter = ReadFromFile(optionsFilterConfigFile)
    End If

    dlg.circuitFilter = ""
    Dim circuitFilterConfigFile As String
    circuitFilterConfigFile = GetCircuitFilterConfigFile()
    If (circuitFilterConfigFile <> "") Then
    	dlg.circuitFilter = ReadFromFile(circuitFilterConfigFile)
    End If
    dlg.dialect = 3
    Dialog dlg ' show dialog
End Sub
