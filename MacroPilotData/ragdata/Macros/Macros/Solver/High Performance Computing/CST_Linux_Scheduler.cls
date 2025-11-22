'#Reference {F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}#1.0#0#C:\Windows\SysWOW64\wshom.ocx#Windows Script Host Object Model#IWshRuntimeLibrary
VERSION 2019.0 Class
Begin
  MultiUse = -1  'True
End
Attribute VB_Name = "CST_Linux_Scheduler"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'CST_Linux_Scheduler.cls
'CST_Linux_Scheduler|New Module|Class Module
'Edit|Properties|Name=CST_Linux_Scheduler
'#Language "WWB-COM"
'#Uses "CSTClusterConf.cls"
Option Explicit

''' The class CST_Linux_Scheduler handles functions and properties specific for the Linux Scheduler.
'''
''' From this class the functions configure and submit (in this order) have to be called. It references to the Shell runtime library to allow a better handling of the Shell.
'''
''' @version 20180705

' ----- fixed settings
Private mobjClusterConf As Object ' Handle to parent CSTClusterConf Class
Private objShell As Object 'Object with open shell; set in @link #Configure Configuration @endlink

' ----- variable settings - set in Class Functions

Private mstrCstJobSubmit As String ' Submit Command on Cluster
Private mstrPlinkSetting As String

''' Enumeration defining the possible Acceleration options. see CSTClusteConf for details
''' @todo manage to have this enum definition only in CSTClusterConf
Private Enum AccOptionType
	Single_Node
	Distributed_computing
	MPI_computing
	Distributed_computing_And_MPI_computing 'just for completeness, in principle this is 1+2=3 ;-)
End Enum

'' CHANCED THIS HERE
Public strRequiredSettings As Variant 'define in Initialization!
Public strSelectionLabel as String

Private mbInitialized As Boolean

' ----- Set/Let and Get Functions of fixed Settings
Property Get bInitialized() As Boolean
	Attribute bInitialized.VB_UserMemId = 0
	bInitialized = CBool(mbInitialized)
End Property


' ----- Set/Let and Get Functions of fixed Settings

''' ReadWrite Property
''' Parent CSTClusterConf Class object. Needs to be set in CSTClusterConf.Scheduler!
''' @returns objCSTClustersConf object reference to parent @link #CSTClusterConf Configuration @endlink
Property Set ClusterConf(objCSTClusterConf As CSTClusterConf)

	If mobjClusterConf Is Nothing Then
		Set mobjClusterConf = objCSTClusterConf
	End If
End Property
Property Get ClusterConf() As CSTClusterConf
	Set ClusterConf = mobjClusterConf
End Property

' ----- (Let and) Get Functions of variable settings - set in Class Functions (e.g. using GUI) requesting information from scheduler system
'maybe I can convert all of these to functions or subs
''' ReadOnly Property
''' Contacts the Scheduler and determines available queues. If multiple queues are available the user is asked which one to take; Sets mstrCurrentQueue to (chosen) value
''' @returns String array with available queues
Property Get strAvailableQueues() As String()


	mobjClusterConf.Log("Getting available queues from cluster",2)

	Dim strtmp() As String

	strtmp = Split(clusterInfo("","get-available-queues"),":")


	If Ubound(strtmp) = -1 Then
		'This part will be used when get-available-queues failed
		strtmp = mobjClusterConf.strAvailableQueues
	end if

	mobjClusterConf.strAvailableQueues = strtmp
	strAvailableQueues = strtmp

End	Property

''' Write Property
''' Currently selected queue. If the queue is changed the parameters available acceleration options, node and GPU number and available solvers are determined
''' @returns String with currently set queues
'do I really need this?
Property Let strQueue(Queue As String)
	mobjClusterConf.Log("Setting queue to: " + Queue)
	mobjClusterConf.strSetQueue = Cstr(Queue)

	' If queue is changed, everything needs to change
	strAvailableClusterAcc
	strAvailableSolvers
	intAvailableGpus
	intAvailableNodes
End Property

''' ReadOnly Property
''' Contacts the Scheduler and determines available acceleration types uses CSTClusterConf.setAccOption to set type in CSTClusterConf. Called automatically by queue.
''' @returns String array with available cluster accelerations
Private Property Get strAvailableClusterAcc() As String()


	Dim strtmp() As String
	strtmp = Split(clusterInfo(mobjClusterConf.strSetQueue,"get-available-cluster-acc"),":")

	strAvailableClusterAcc = strtmp

	With mobjClusterConf
	   .setAccOption(Distributed_computing_And_MPI_computing,False)
	Dim i As Integer
	For i = LBound(strtmp) To UBound(strtmp)
	   .setAccOption(eval(strtmp(i)),True)
	Next i

	End With
	mobjClusterConf.Log("Available acceleration methods in queue: " + Join(strtmp),2)
End Property


'''
''' @brief Get available solvers
''' Contacts the Scheduler and determines available Solvers. Calls CSTClusterConf.ChangeSupported to adjust the global solver table of CSTClusterConf
''' @returns String array with available Solvers
Private Property Get strAvailableSolvers() As String()

  	With mobjClusterConf

	   Dim strtmp() As String

	   strtmp = Split(clusterInfo(mobjClusterConf.strSetQueue,"get-available-solvers"),":")
	   strAvailableSolvers = strtmp
	   .ChangeSupported(strtmp)
	   .Log("Available Solvers in queue: " + Join(strtmp))
	End With
End Property

''' @brief Get number of available GPUs
''' Contacts the Scheduler and determines the number of available GPUS. Uses CSTCLusterConf.numGPU to set value in CSTClusterConf
''' @returns Integer with number of available GPUs in queue
Private Property Get intAvailableGpus() As Integer

	With mobjClusterConf
		.intQueueMaxGPUs = Cint(CByte(clusterInfo(mobjClusterConf.strSetQueue,"get-available-gpu-num")))
		.Log("Set available GPUs on Cluster to " + Cstr(.intQueueMaxGPUs))
		intAvailableGpus = .intQueueMaxGPUs
	End With

End Property

''' @brief Get number of available nodes
''' Contacts the Scheduler and determines the number of available nodes. Uses CSTClusterConf.intQueueMaxNodes to set value in CSTClusterConf
''' @returns Integer with number of available nodes in queue
Private Property Get intAvailableNodes() As Integer

   	With mobjClusterConf
	    .intQueueMaxNodes = CInt(CByte(clusterInfo(mobjClusterConf.strSetQueue,"get-available-node-num")))
		.Log("Set available Nodes on Cluster to " + Cstr(.intQueueMaxNodes))
		intAvailableNodes = .intQueueMaxNodes
	End With
End Property

' ---- Class Initialization and Termination

''' Constructor of #CST_Linux_Scheduler
'''
''' Initialization does nothing specific so far
''' @sa Class_Terminate
Private Sub Class_Initialize()

	mbInitialized = True
	strRequiredSettings = Array("SchedulerNode","CstInstallPath","WorkDirectory","Username","Password")

	strSelectionLabel = "Queue"
	Debug.Print "Initializing CST_Linux_Scheduler Class"

	Set objShell = CreateObject("WScript.Shell")

	''Resetting Plink authentication

	mstrPlinkSetting = "1"
	On Error Resume Next
	mstrPlinkSetting= objShell.RegRead("HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\Default%20Settings\AuthGSSAPI")
	On Error GoTo 0

	objShell.RegWrite "HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\Default%20Settings\AuthGSSAPI", "0", "REG_DWORD"


End Sub

''' @brief Destructor of #CSTClusterConf
'''
''' frees the Memory of the reference to the parent class #mobjClusterConf and objShell
''' \sa Class_Initialize
Private Sub Class_Terminate()
	mobjClusterConf.Log("Terminating CST_Linux_Scheduler")

	If Not (objShell Is Nothing) Then

		objShell.RegWrite "HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\Default%20Settings\AuthGSSAPI", mstrPlinkSetting, "REG_DWORD"
		Set objShell = Nothing
	End If

	If Not (mobjClusterConf Is Nothing) Then Set mobjClusterConf = Nothing

End Sub

' ----- Class Functions

Sub Run()

		Begin Dialog UserDialog 410,51,"Choose Action"
		PushButton 20,21,100,21,"Submit Job",.Submit '1
		PushButton 140,21,100,21,"Fetch Job",.Fetch '2
		PushButton 260,21,130,21,"Cluster Settings...",.Cluster '3
		End Dialog
		Dim dlg As UserDialog


		Select Case Dialog(dlg) 'in case of 1, go on
			Case 2: Fetch()
			Case 3: mobjClusterConf.ConfigGui
		End Select

End Sub
''' @brief Submit Job to Linux Cluster
'''
''' Submits job to Linux cluster by first transferring the model to the cluster via scp (in a subdirectory with the project name in the working directory (set in CSTCLusterConf). Afterwards the function calls the script on the cluster to schedule the simulation. After submission the macro is done.
''' <br><b>Submit hast to be called in the Macro explicitly (See Interact with Scheduler)
Sub Submit()

	Dim fso As Object

	Set fso = CreateObject("Scripting.FileSystemObject")

	Dim intReturnValue As Integer
	Dim strRemotepath As String

	Dim strCommandLine As String
	Dim strCommandLineLog As String

	With mobjClusterConf

	Dim AccCommand As String
	AccCommand = ""

	Dim tmpaccoption As AccOptionType
	tmpaccoption = Eval(.strAccOption)

	If CBool(tmpaccoption) Then
		AccCommand = " -a " + .strAccOption
	End If

	Dim GPUCommand As String
	GPUCommand = ""
	If CBool(.intSetGPUs) Then
		GPUCommand = " -g " + .intSetGPUs
	End If

	Dim NodeCommand As String
	NodeCommand = " -n 2"
	If CBool(.intSetNodes) Then
		NodeCommand = " -n " + .intSetNodes
	End If

	Dim ThreadCommand As String		'uses --numthreads
	ThreadCommand = ""
	If CBool(.intSetThreads) Then
		ThreadCommand = " -t " + .intSetThreads
	End If



	Dim straddhop As String

	straddhop = ""


	' not Single node!
	If (clusterInfo("","get-queuesys-name") = "None") Then
		Dim ListArray() As String
		ListArray = Split(clusterInfo("","get-available-node-names"),":")

		Begin Dialog UserDialog 400,203,"Node selection" ' %GRID:10,7,1,1
			Text 40,14,310,21,"Select " + .intSetNodes + " Nodes",.Text1
			MultiListBox 40,42,310,112,ListArray(),.nodenames
			OKButton 30,168,120,21
			CancelButton 220,168,120,21
		End Dialog
		Dim dlg As UserDialog
		Do
			If (Dialog(dlg) = 0) Then Exit All
		Loop While UBound(dlg.nodenames)+1 <> .intSetNodes


		Dim item As Integer

		Dim strtmp As String
		strtmp = ""
		For Each item In dlg.nodenames
			If Not(Len(strtmp) = 0) Then
               strtmp = strtmp + ","
			End If

			strtmp = strtmp + Trim(ListArray(item))
		Next item

		NodeCommand=" -n " + strtmp
	    'straddhop = "ssh " + ListArray(item) + " "
		straddhop = "ssh " + mobjClusterConf.strSchedulerNode + " "

	End If

	End With

	With mobjClusterConf
	strRemotepath = .strWorkDirectory + "/" +  fso.GetFileName(.strModelFile) + "/"

	If (CInt(CByte(sshCommand("/bin/bash -c 'if [ -d " + strRemotepath + " ]; then echo 1; else echo 2; fi'"))) = 1) Then
		.Log("Work folder already existing") ' TODO: is this a problem?
		MsgBox("Project with same name exists on cluster. Please rename your CST model file.",vbOkOnly+vbExclamation)
		Exit All
	Else
		.Log("Creating work folder.")
		sshCommand("mkdir -pv " + strRemotepath)
	End If

	' escape quotes in password
	Dim strPwd As string
	strPwd = Replace$(.strPassword, """", """""")
	strPwd = .quote(strPwd)

	strCommandLine  = .quote(GetInstallPath + "\pscp")
	strCommandLine += " " + "-pw" + " " + strPwd
	strCommandLine += " " + "-C"
	strCommandLine += " " + scpFileList(.strModelFile)
	strCommandLine += " " + .strUsername + "@" + .strSchedulerNode  + ":" + strRemotepath

	strCommandLineLog  = .quote(GetInstallPath + "\pscp")
	strCommandLineLog += " " + "-pw" + " " + "<pw>"
	strCommandLineLog += " " + "-C"
	strCommandLineLog += " " + scpFileList(.strModelFile)
	strCommandLineLog += " " + .strUsername + "@" + .strSchedulerNode  + ":" + strRemotepath

	.Log("Transferring file to cluster.",2)
	.Log("Transfer command: " + strCommandLineLog)
	intReturnValue = objShell.Run(strCommandLine ,1, True)


	.Log("Submitting job, this may take a while.",2)

	strCommandLine = straddhop + mstrCstJobSubmit + " -q " + .strSetQueue + " -m " + strRemotepath + fso.GetFileName(.strModelFile) + ".cst -s "+ .strCurrentSolver + GPUCommand + AccCommand + NodeCommand + ThreadCommand

	.Log("Starting job on scheduler.",2)
	.Log("Sending command to cluster: " + strCommandLine)

	'command -v nohup' to get command of nohup


	 CStr(sshCommand(strCommandLine,1)) 'TODO: Here i can pipe to the stdout file and decouple the shell " $nohup /bin/bash -c ... &" or nohup

	.Log("Submit done.",2)
	End With

	If (clusterInfo("","get-queuesys-name") = "None") Then
		Fetch()
	End If

End Sub

Function scpFileList(modelFile As String) As String
	scpFileList += Chr$(34) + modelFile + ".cst" + Chr$(34)

	' non-parametric optimizer: add <model>.inp file
	If Dir(modelFile + ".inp") <> "" Then
		scpFileList += " " + Chr$(34) + modelFile + ".inp" + Chr$(34)
	End If

	' non-parametric optimizer: add <model>_tosca.par file
	If Dir(modelFile + "_tosca.par") <> "" Then
		scpFileList += " " + Chr$(34) + modelFile + "_tosca.par" + Chr$(34)
	End If
End Function

''' @brief Fetch Job from Linux Cluster or display status if not ready yet
'''
''' Fetches jobs from cluster if ready;  Therefore it checks if the Model.lok file is created; if yes it collects the process; if no either the process.log does not exist so the project is not started yet, or the results are fetched from the Cluster into a new built subdirectory
''' <br><b>Submit hast to be called in the Macro explicitly (See Interact with Scheduler)

Sub Fetch()

	mobjClusterConf.Log("Trying to fetch job data, this may take a while.",2)
	Dim fso As Object

	Set fso = CreateObject("Scripting.FileSystemObject")

	Dim intReturnValue As Integer
	Dim strRemotepath As String
	Dim strRemoteProjectPath As String

	Dim strCommandLine As String
	Dim strCommandLineLog As String
	Dim i As Integer
	Dim newfolder As String

	With mobjClusterConf
	   strRemotepath = .strWorkDirectory + "/" +  fso.GetFileName(.strModelFile) + "/"
	   strRemoteProjectPath = strRemotepath + fso.GetFileName(.strModelFile) + "/"


	Dim Filelocked As Boolean
	Dim bProgress As Boolean
	Filelocked = (CInt(CByte(sshCommand("bash -c 'if [ -f " + strRemoteProjectPath + "Model.lok ]; then echo 1; else echo 2; fi'"))) = 1)
	bProgress = (CInt(CByte(sshCommand("bash -c 'if [ -f " + strRemoteProjectPath + "Result/progress.log ]; then echo 1; else echo 2; fi'"))) = 1)

	If (Filelocked And bProgress) Then
		Dim a() As String

		a = Split(sshCommand("tr '\n'  _  < " + strRemoteProjectPath + "Result/progress.log"),"_")

		Dim progress As Integer
		Dim progressmessage As String
		progress = CInt(Split(a(2),":")(2))

		progressmessage = "Simulation Status: " + Split(a(0),":")(2) + ": " + Split(a(1),":")(2) + vbCrLf + _
			"Progress: [" + String$(Fix(progress/10),"=") + String$(10-Fix(progress/10),"_") + _
			"]          " + CStr(progress) +"%"

		MsgBox(progressmessage,"CurrentProgress")

	ElseIf(Not(Filelocked) And Not(bProgress)) Then
		MsgBox("Project not started, please be patient.",vbOkOnly,"Current Progress")
	ElseIf (Not(Filelocked) And bProgress) Then
		i = 1

		While Len(Dir(.strModelFile + "-" + Format(i,"0000"),vbDirectory)) <> 0
			i = i+1
		Wend

		newfolder = .strModelFile + "-" + Format(i,"0000")
		MkDir newfolder

		' escape quotes in password
		Dim strPwd As string
		strPwd = Replace$(.strPassword, """", """""")
		strPwd = .quote(strPwd)

		strCommandLine  = mobjClusterConf.quote(GetInstallPath + "\pscp")
		strCommandLine += " " + "-pw" + " " + strPwd
		strCommandLine += " " + "-C"
		strCommandLine += " " + "-r"
		strCommandLine += " " + .strUsername + "@" + .strSchedulerNode  + ":" + strRemotepath +  "*"
		strCommandLine += " " + .quote(newfolder)

		strCommandLineLog  = mobjClusterConf.quote(GetInstallPath + "\pscp")
		strCommandLineLog += " " + "-pw" + " " + "<pw>"
		strCommandLineLog += " " + "-C"
		strCommandLineLog += " " + "-r"
		strCommandLineLog += " " + .strUsername + "@" + .strSchedulerNode  + ":" + strRemotepath +  "*"
		strCommandLineLog += " " + .quote(newfolder)

		.Log("Transferring file back.",2)
		.Log("Transfer command: " + strCommandLineLog)
		intReturnValue = objShell.Run(strCommandLine ,1, True)
		.Log("Transferred back to " + newfolder + ".",2)

		If (intReturnValue = 0) Then
		sshCommand("rm -r -f " + strRemotepath + ";echo 1")
		End If

	End If
	End With

	Exit All
End Sub


''' @brief Configuration of CST_Linux_Scheduler
'''
''' Configures the scheduler system for example by checking for variables. Afterwards the functions for determining the setup in the cluster are triggered
Friend Sub Configure()
	With mobjClusterConf

		Dim Title As String
		Debug.Print "Configuring CST_Linux_Scheduler"

		mstrCstJobSubmit = .strCstInstallPath + "/cst_job_submit -b"

		.Log("Submit script is " + mstrCstJobSubmit)
		.Log("Asking cluster for queues and setting queue to default.")

		If CInt(.strDefaultQueue) > UBound(strAvailableQueues) Then
			.strDefaultQueue = "0"
		End If

		strQueue = strAvailableQueues(CInt(.strDefaultQueue)) ' Setting a queue also means registering all queue parameters!

	End With
End Sub


' ----- Help Functions
''' @brief Function determining info strInfotype (e.g. acceleration options) from a queue strQueue. If no queue is given the available queues will be fetched
'''
''' @returns String with answer from Cluster script
Private Function clusterInfo(strQueue As String, strInfotype As String) As String

	If (Len(strQueue) = 0) Then
		mobjClusterConf.Log("Getting clusterInfo " + strInfotype)
		clusterInfo = Trim$(Cstr(sshCommand(mstrCstJobSubmit + " --" + strInfotype)))
	Else
		mobjClusterConf.Log("Getting clusterInfo " + strInfotype + " from queue " + strQueue)
		clusterInfo = Trim$(CStr(sshCommand(mstrCstJobSubmit + " -q " + strQueue + " --" + strInfotype)))
	End If
End Function

''' @brief Function sending the command strCommand via ssh.

''' The command line is sent using plink (installed by CST STUDIO SUITE install path) and read by using an intermediate log file. Plink runs in batch mode to ensure that the macro does not get stuck because of a hanging command line. If an error occurs (output empty, too long or return value is not 0).
'''
''' @returns String returned from ssh command
Function sshCommand(strCommand As String, Optional visible As Integer = 0) As String

   With mobjClusterConf
	.Log("Executing ssh command: " + strCommand)


	Dim strOut() As String
	Dim strCommandLine As String
	Dim strCommandLineLog As String
	Dim intReturnValue As Integer
	Dim errornum As Integer
	Dim Errordescription As String
	Dim helpstring As String
	Dim helparray() As String

	' escape quotes in password
	Dim strPwd As string
	strPwd = Replace$(.strPassword, """", """""")
	strPwd = .quote(strPwd)

	strCommandLine  = .quote(GetInstallPath + "\plink.exe")
	strCommandLine += " " + "-no-antispoof"
	strCommandLine += " " + "-batch"
	strCommandLine += " " + "-pw" + " " + strPwd
	strCommandLine += " " + .strUsername + "@" + .strSchedulerNode
	strCommandLine += " " + strCommand

	strCommandLineLog  = .quote(GetInstallPath + "\plink.exe")
	strCommandLineLog += " " + "-no-antispoof"
	strCommandLineLog += " " + "-batch"
	strCommandLineLog += " " + "-pw" + " " + "<pw>"
	strCommandLineLog += " " + .strUsername + "@" + .strSchedulerNode
	strCommandLineLog += " " + strCommand

	.Log(strCommandLineLog)
	strOut = Split(RunHPCCommand(strCommandLine),vbCrLf)

	errornum = -3
	errornum = cint(strOut(0))


	if (strOut(1) = "Access denied") then errornum = 1
	if (UBound(strOut) <> 2) and not cbool(visible) then errornum = 1

	If (errornum = 1) Then
		strOut(0) = "Exist code: " + strOut(0)
		ReportErrorToWindow(strCommandLineLog + vbCrLf + Join(strOut, vbCrLf))
		strOut(0) = ""
		.Log("Error: " + Join(strOut))
		MsgBox("SSH command failed.", vbOkOnly + vbCritical)
		.ConfigGui
		Exit Function

	If errornum <> 0 Then

		.Log("Command returned with error code " + Cstr(errornum),3)
		strOut(0) = ""

		If errornum < 0 Then errornum = -errornum

		Select Case errornum
				Case 1: Errordescription =  "Can't create output file."
				'Case 2: same as else
				Case 3: Errordescription = "sshcommand terminated abnormally."
				Case Else: Errordescription = Join(strOut)
		End Select

		.Log("And value " + Errordescription)
		Err.Raise 70000 - errornum, "sshcommand", Errordescription

		Exit All

		End If

	end if

	if cbool(visible) then
		strout(0) = ""
		msgbox join(strout,vbCrLf)
	End If

   End With


   sshCommand = strOut(1)

End Function
