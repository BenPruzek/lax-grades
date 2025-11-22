'#Reference {C45D10A1-54E8-420B-A052-719D47EC7C16}#2.0#0#C:\Program Files\Microsoft HPC Pack 2019\Bin\Microsoft.Hpc.Scheduler.tlb#Microsoft.Hpc.Scheduler#Microsoft_Hpc_Scheduler
'#Reference {F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}#1.0#0#C:\Windows\SysWOW64\wshom.ocx#Windows Script Host Object Model#IWshRuntimeLibrary
VERSION 2019.0 Class
Begin
  MultiUse = -1  'True
End
Attribute VB_Name = "CST_Microsoft_HPC_Scheduler_2019"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'CST_Microsoft_HPC_Scheduler.cls
'CST_Microsoft_HPC_Scheduler|New Module|Class Module
'Edit|Properties|Name=CST_Microsoft_HPC_Scheduler_2019
'#Language "WWB-COM"
'#Uses "CSTClusterConf.cls"
Option Explicit


''' The class CST_Microsoft_HPC_Scheduler_2019 handles functions and properties specific for the Microsoft HPC Scheduler 2019.
''' Uses the API of Microsoft HPC Scheduler 2019. Thus it references to the Shell runtime library (for network functions) and the Microsoft HPC Scheduler. From this class the functions configure and submit (in this order) have to be called.
''' @version 20180705

' ----- fixed settings
Private mobjClusterConf As Object ' Handle to parent CSTClusterConfConf Class
Private mobjMSScheduler As Object 'Scheduler

Private Enum SelectionType
	Template
	Nodegroup
End Enum
Private mSelectionType As SelectionType 'can be read from CSTClusterConf!
'Private mstrJobTemplateList() As String
'Private mstrJobTemplate As String ' Name of Template or node group to use
Private Enum AccOptionType
	Single_Node
	Distributed_computing
	MPI_computing
	Distributed_computing_And_MPI_computing 'just for completeness, in principle this is 1+2=3 ;-)
End Enum

Public strRequiredSettings As Variant 'define in Initialization!
Public strSelectionLabel As String

Private mbInitialized As Boolean


' ----- variable settings - set in Class Functions

' ----- Set/Let and Get Functions of fixed Settings
Property Get bInitialized() As Boolean
	Attribute bInitialized.VB_UserMemId = 0
	bInitialized = CBool(mbInitialized)
End Property

''' ReadWrite Property
''' Parent CSTHPCConf Class object. Needs to be set in CSTHPCConf.Scheduler! Set Option checks if project is saved on a network share and replace volume letter with mapped network share (if available)
''' @returns objCSTHPCConf object reference to parent @link #CSTHPCConf HPCConfiguration @endlink
Property Set ClusterConf(objCSTClusterConf As CSTClusterConf)
	Set mobjClusterConf = objCSTClusterConf
End Property

Property Get ClusterConf() As CSTClusterConf
	Set ClusterConf = mobjClusterConf
End Property

' ----- (Let and) Get Functions of variable settings - set in Class Functions (e.g. using GUI)

Property Get strAvailableQueues() As String()
	
	Dim i As Integer
	Dim stringitem As Variant
	
	'Select NodeSelection
	Dim mstrJobTemplateList() As String
	ReDim mstrJobTemplateList(mobjMSScheduler.GetJobTemplateList.Count)
	i = 0
	
	For Each stringitem In mobjMSScheduler.GetJobTemplateList
		mstrJobTemplateList(i) = CStr(stringitem)
		i = i+1
	Next stringitem

	Const strNodeSelection = Array("Template")', "Nodegroup")

	mobjClusterConf.strAvailableQueues = mstrJobTemplateList
	strAvailableQueues = mstrJobTemplateList

	'Dim returnValue As JobTemplateInfo
	'returnValue = instance.GetJobTemplateInfo("Default")
	
End Property


Property Let strQueue(Queue As String)
	mobjClusterConf.Log("Setting queue to: " + Queue)
	mobjClusterConf.strSetQueue = Cstr(Queue)
End Property

''' ReadOnly Property
''' Contacts the Scheduler and determines the number of available nodes. Uses CSTCLusterConf.numNodes to set value in CSTClusterConf
''' @returns Integer with number of available nodes in queue
Property Get available_node_num() As Integer
	'mobjMSScheduler.GetCounters.TotalNodes
End Property


''' ReadOnly Property
''' Creates Jobname from full path of ModelFile
''' @returns String with Jobname
Property Get strJobName() As String
   With mobjClusterConf
      strJobName = Right$(.strModelFile,Len(.strModelFile )-InStrRev(.strModelFile ,"\"))
   End With
End Property

' ---- Class Initialization and Termination

''' Constructor of #CST_Microsoft_HPC_Scheduler
'''
''' @sa Class_Terminate
Private Sub Class_Initialize()
	
	
	mbInitialized = True
	strRequiredSettings = Array("SchedulerNode","CstInstallPath","ConsoleConnect","CopyToLocal")',"Username","Password")
	
	strSelectionLabel = "Template"
	
	
	Debug.Print "Initializing CST_Microsoft_HPC_Scheduler_2019" 
	
	
	On Error GoTo noMSHPCScheduler

	Set mobjMSScheduler = CreateObject("Microsoft.HPC.Scheduler.Scheduler")

	
ExitSub:
	Exit Sub

noMSHPCScheduler:
	Debug.Print("No Microsoft HPC Scheduler found.")
	MsgBox("No Microsoft HPC Scheduler found.  Please contact your system administrator for the installation of Microsoft HPC Pack.")
	mbInitialized = False
	Resume ExitSub
	
End Sub

''' @brief Destructor of #CST_Microsoft_HPC_Scheduler
'''
''' frees the Memory of the reference to the parent class #mobjClusterConf
''' \sa Class_Initialize
Private Sub Class_Terminate()
	mbInitialized = False
	If Not mobjClusterConf Is Nothing Then
			mobjClusterConf.Log("Terminating CST_Microsoft_HPC_Scheduler_2019")
		Set mobjClusterConf = Nothing
	End If
End Sub

' ----- Class Functions
''' @brief Run Job to Microsoft HPC Scheduler
'''
''' Run function (must be called macro!) Questioning for Number of nodes etc.
Sub Run()

	Dim objShell As Object
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run("""C:\Program Files\Microsoft HPC Pack 2019\Bin\job"" submit /scheduler:" + mobjClusterConf.strSchedulerNode + " hostname" ,1, True) 'verify credentials, to test this use hpccred command and delete credentials on hpc cluster

	With mobjClusterConf
	
		If .BadFileName(.strModelFile) Then
			MsgBox "File name contains bad characters. Please rename the model file."
			Exit All
		End If
	
		.strModelFile = MapNetworkDrive(.strModelFile)
	End With
End Sub


''' @brief Submit Job to Microsoft HPC Scheduler
'''
''' Submit function submits job when all settings are set by JobDialogFunc. After submission it closes the current project in the CST Design Environment.
Sub Submit()
	
	mobjClusterConf.Log("Submit Job")

	Dim job As Object 
	Set job = mobjMSScheduler.CreateJob()

	Dim iCorenum as integer
	
	if not(Cbool(mobjClusterConf.intSetThreads)) then
		iCorenum = mobjMSScheduler.GetCounters.TotalCores/mobjMSScheduler.GetCounters.TotalNodes*mobjClusterConf.intSetNodes
	else
		iCorenum = Fix(min(mobjClusterConf.intSetThreads,mobjMSScheduler.GetCounters.TotalCores/mobjMSScheduler.GetCounters.TotalNodes*mobjClusterConf.intSetNodes))
	end if
	
	With job
		.Name = "CST Simulation"
		.MinimumNumberOfNodes = mobjClusterConf.intSetNodes
		.MaximumNumberOfNodes = mobjClusterConf.intSetNodes
		mobjClusterConf.Log("Number of used nodes: " + CStr(mobjClusterConf.intSetNodes))

		.MinimumNumberOfCores = iCorenum
		.MaximumNumberOfCores = iCorenum	
		mobjClusterConf.Log("Number of cores on all nodes: " + CStr(iCorenum))

		
		.SetJobTemplate(mobjClusterConf.strSetQueue)
		mobjClusterConf.Log("Using Job Template " + mobjClusterConf.strSetQueue)
		.SetEnvironmentVariable("CST_WAIT_FOR_LICENSE","1")
				
		If (mobjClusterConf.strConsoleConnect="Attach") Then
			.SetEnvironmentVariable("HPC_ATTACHTOSESSION","True")
		Else
			.SetEnvironmentVariable("HPC_CREATECONSOLE","True")
		End If
		
		.SetEnvironmentVariable("I_MPI_AUTH_METHOD","delegate")
		.UnitType = 0 'UnityTypeEnum 0=core, 1=socket, 2=Node
		.IsExclusive = true
	End With
	
	Dim task As Object 
	Set task = job.CreateTask()

	With task
	
		.Name = Me.strJobName
		.CommandLine = CommandLine
		.MinimumNumberOfNodes = mobjClusterConf.intSetNodes
		.MaximumNumberOfNodes = mobjClusterConf.intSetNodes
		.MinimumNumberOfCores = iCorenum	
		.MaximumNumberOfCores = iCorenum
		
		mobjClusterConf.Log("Using Command Line: " + mobjClusterConf.CommandLine)
		.StdErrFilePath = mobjClusterConf.strModelFile + "\stderr.log"
		.StdOutFilePath = mobjClusterConf.strModelFile + "\stdout.log"
	End With
		
    job.AddTask(task)

	rem wait 60
	mobjMSScheduler.SubmitJob(job, NULL, NULL)
	mobjClusterConf.Log("Submitting.")
	
	'open untitled project
	
	Dim app As Object
	Dim mws As Object
	Set app = CreateObject("CSTStudio.Application."+ Split(Split(Replace(GetApplicationVersion(), "Version ", ""),"-")(0),".")(0)) 'TEST
	Set mws = app.NewMWS
	
	'current project will be closed when terminating @CSTClusterConf

	
End Sub

''' @brief Configuration of Microsoft_Hpc_Scheduler.
'''
''' Configuration of Scheduler system.
Friend Sub Configure()

   With mobjClusterConf
   
	On Error Goto ShowConfigGui
		
		.Log("Connecting to cluster node " + .strSchedulerNode,2)
		mobjMSScheduler.Connect(.strSchedulerNode)

		.intQueueMaxGPUs = 1000
		.Log("Set available GPUs on Cluster to " + Cstr(.intQueueMaxGPUs))
		.intQueueMaxNodes = 1000
		.setAccOption(MPI_computing,True)
		.setAccOption(Distributed_Computing,False)

		.strSetQueue = strAvailableQueues(CInt(.strDefaultQueue))

   End With

   Exit Sub

   ShowConfigGui:
   mobjClusterConf.ConfigGui
   
	
End Sub



' ----- Help Functions

Function CommandLine() As String
   mobjClusterConf.strMachineFile = "%TMPFILE%"
   CommandLine = mobjClusterConf.quote(WriteWrapperScript)
End Function

Function WriteWrapperScript() As String
	' Function writing a wrapper file for DC or MPI calculation
	' collects the assigned Nodes and tells CST to use these for calculation
	' Input: mCstInstallpath .......... Path of installation of CST DE
	'        StartfromWrapper ........ Command Line which should be start from the Wrapper
	' Output: WriteWrapperScript ..... path to output file/ used as command line input for the job xml file
	
	
	' construct the command line with the CST executable and switches depending on Acceleration and GPU
	' if Acceleration is set make a wrapper script with the command line
	
	Dim WriteWrapperScriptfile as String
	With mobjClusterConf
	
	dim copytolocal as boolean
	
	copytolocal = false
	If (mobjClusterConf.strCopyToLocal="True") Then copytolocal = true
	
	
	WriteWrapperScriptfile = .strModelFile + "\runCST.bat"
	Open  WriteWrapperScriptfile For Output As #1

	Print #1, "@echo Off"
	Print #1, "SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION"	
	Print #1, "GoTo :SCRIPTSTART"
	Print #1, vbLf
	Print #1, ":GETTEMPNAME"
	Print #1, "Set TMPFILE=%TMP%\mytempfile-%Random%-%Random%-%Time:~9%"
	Print #1, "If exist "+ .strMachineFile + " GoTo :GETTEMPNAME"
	Print #1, "GoTo :EOF"
	Print #1, vbLf
	Print #1, ":ADDONETOCOUNTER"
	Print #1, "Set /a COUNTER = %COUNTER% + 1"
	Print #1, "GoTo :EOF"
	Print #1, vbLf
	Print #1, ":WRITEMACHINEIFCOUNTER"
	Print #1, "Set /a COUNTERMODTWO = %COUNTER% %% 2"
	Print #1, "If %COUNTERMODTWO% == 0 echo %1 >> " + .strMachineFile
	Print #1, "Call :ADDONETOCOUNTER"
	Print #1, "GoTo :EOF"
	Print #1, vbLf
	
	If copytolocal then
		Print #1, ":COPYDATA"
		Print #1, "echo Constructing unique folder name..."
		Print #1, "Set MDERR=1"
		Print #1, "FOR /L %%i IN (1,1,50) DO ("
		Print #1, "	Set num=0000%%i"
		Print #1, "	Set TARGETFOLDERNAME=%~2" + strJobName + "-!num:~-5!\"
		Print #1, "	echo Trying !TARGETFOLDERNAME!"
		Print #1, "	IF NOT EXIST !TARGETFOLDERNAME! ("
		Print #1, "		echo --- Seems to be valid."
		Print #1, "		xcopy ""%~1" + strJobName + """ ""!TARGETFOLDERNAME!\" + strJobName + "\"" /E /Y /Q /I /C "
		Print #1, "		xcopy ""%~1" + strJobName + ".cst"" ""!TARGETFOLDERNAME!"" /Y /Q /C "
		Print #1, "		Set MDERR=!ERRORLEVEL!"
		Print #1, "		ENDLOCAL"
		Print #1, "	  	Set RUNPATH=!TARGETFOLDERNAME!"
		Print #1, "		GoTo :EOF"
		Print #1, "  )"
		Print #1, "  echo --- Seems to exist already. Trying next one."
		Print #1, ")"
	End If
	
	Print #1, vbLf
	Print #1, ":SCRIPTSTART"
	Print #1, "PING localhost -n 1 -w 6000 > NUL" 'wait for alle the files to be constructed
	Print #1, "PING 1.1.1.1 -n 1 -w 6000 > NUL" 'wait for alle the files to be constructed
	Print #1, "Set RUNPATH=" + Left$(.strModelFile,Len(.strModelFile )-InStr(.strModelFile ,"\")-Len(strJobName)) + "\"
	Print #1, "Set REMOTEPATH=%RUNPATH%"
	
	If copytolocal Then Print #1, "Call :COPYDATA ""%REMOTEPATH%"" %TMP%\"
		
	Print #1, vbLf
	Print #1, "Call :GETTEMPNAME"
	Print #1, "echo # CSTMachineFile 20080509 > " + .strMachineFile
	Print #1, "echo # DefaultInstallationFolder """ + .strCstInstallpath + """ >> " + .strMachineFile
	Print #1, "echo # DefaultTempFolder %TMP% >> " + .strMachineFile
	Print #1, "echo # DefaultMachineArchitecture Windows AMD64 >> " + .strMachineFile
	Print #1, vbLf
	Print #1, "Set COUNTER=1"
	Print #1, "For %%I In (%CCP_NODES%) Do ("
	Print #1, "Call :WRITEMACHINEIFCOUNTER %%I"
	Print #1, ")"
	Print #1, vbLf
	
	Print #1, "Set LOGFILE=%RUNPATH%" + strJobName + "\Temp\hpc_log.txt"
	
	Print #1, ":PRINT_ENVIRONMENT"
	Print #1, "echo ""Running job"" >> ""%LOGFILE%"""
	Print #1, "echo ""Environment"" >> ""%LOGFILE%"""
	Print #1, "set >> ""%LOGFILE%"""
	Print #1, "echo ""HardwareInfo"" >> ""%LOGFILE%"""
	Print #1, .quote(. strCstInstallPath + "\AMD64\CSTHardwareInfo_AMD64.exe") + " --nowait" + " >> ""%LOGFILE%"""
	Print #1, "echo ""DeviceInfo"" >> ""%LOGFILE%"""
	rem Print #1, "START /WAIT " + .quote( .strCstInstallPath + "\AMD64\HWAccDiagnostics_AMD64.exe --deviceinfo") + " >> ""%LOGFILE%"""
	Print #1, "echo ""Machinefile"" >> ""%LOGFILE%"""
	Print #1, "type " + .strMachineFile + " >> ""%LOGFILE%"""

	Print #1, vbLf
	Print #1, ":FILE_CHECK"
	Print #1, "IF NOT EXIST ""%RUNPATH%" + strJobName + "\Model.lok"" GoTo :PROJECTCLOSED"
	Print #1, "ping 127.0.0.1 -n 1 -w 3000 > nul"
	Print #1, "GoTo :FILE_CHECK"
	Print #1, vbLf
	Print #1, ":PROJECTCLOSED"
	Print #1, "start ""Status"" ""%RUNPATH%" + strJobName + "\status.bat"" "
	Print #1, "Set COMMANDLINE=" + .CommandLine
	Print #1, "Set NEWCOMMANDLINE=!COMMANDLINE:%REMOTEPATH%=%RUNPATH%!" 
	Print #1, "start /Wait ""CST"" !COMMANDLINE:%REMOTEPATH%=%RUNPATH%!"
	Print #1, "Job modify %CCP_JOBID% /progress:100 /progressmsg:""Finished simulation."""
	
	If copytolocal Then
		Print #1, "Set LOCALPATH=%RUNPATH%"
		Print #1, "Call :COPYDATA ""%RUNPATH%"" ""%REMOTEPATH%"""
		Print #1, "rmdir /s /q ""%LOCALPATH%"""
	End If
			
	Print #1, "del %TMPFILE%"
	Print #1, ":EOF"

	Close #1
	dim StatusScript as String
	StatusScript = .strModelFile + "\status.bat"
	Open  StatusScript For Output As #2

	Print #2, "@echo Off"
	Print #2, ":start"
	Print #2, "for /f ""usebackq delims=: tokens=1,3"" %%a in (""%~dp0Result\progress.log"") do ("
	Print #2, "If ""%%a""==""Progress"" ("
    Print #2, "	Set PR=%%b "
	Print #2, "	) Else If ""%%a""==""Header"" ("
	Print #2, "		Set MSG_HEAD=%%b"
	Print #2, ")	 Else If ""%%a""==""Text"" ("
	Print #2, "		Set MSG_TEXT=%%b"
	Print #2, "	)"
	Print #2, ")"
	Print #2, "Job modify %CCP_JOBID% /progress:%PR% /progressmsg:""%MSG_HEAD%: %MSG_TEXT%"""
	Print #2, "ping localhost -n 30 > nul"
	Print #2, "goto :start"
	
	Close #2
	' Report successful creation and writing of Wrapper Script
	.Log("Wrote WrapperScript " + CStr(WriteWrapperScriptfile),2)
	WriteWrapperScript = WriteWrapperScriptfile
	
	End With
End Function

' Windows Assistance Function
Function MapNetworkDrive(ByVal sFileName As String) as String
	' Function changes volume of sFileName path to network share if mapped
	' Input: sFileName .... ref to path of ModelFile, changed if function successful
	Dim WshNetwork, oDrives As Object
	Dim i As Integer
	Dim networkpath As String
	
	mobjClusterConf.Log("Mapped File " + sFileName)
	
	' fetch all network mappings using Wscript.network
	Set WshNetwork = CreateObject("WScript.Network")
	Set oDrives = WshNetwork.EnumNetworkDrives

	' search for the network share the project is stored on
	For i = 0 To oDrives.Count-1 STEP 2
		If (oDrives.Item(i) Like Left$(sFileName,2)) Then
			networkpath = oDrives.Item(i+1)
			sFileName = networkpath + Right$(sFileName,Len(sFileName)-2)
		End If
	Next

	' if the model path was not mapped complain that model is not stored on network
	If Left$(sFileName,2) <> "\\" Then
		MsgBox "Model not stored on network path.",vbCritical
	End If
	MapNetworkDrive = sFileName
	mobjClusterConf.Log("to Network filename " + sFileName)

End Function

