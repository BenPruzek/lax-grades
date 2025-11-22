VERSION 2019.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSTClusterConf"
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'CSTClusterConf.cls
'CSTClusterConf|New Module|Class Module
'Edit|Properties|Name=CSTClusterConf
'#Language "WWB-COM"
'#Uses "CST_Microsoft_HPC_Scheduler_2012.cls"
'#Uses "CST_Microsoft_HPC_Scheduler_2016.cls"
'#Uses "CST_Microsoft_HPC_Scheduler_2019.cls"
'#Uses "CST_Linux_Scheduler.cls"
Option Explicit

' The class CSTClusterConf handles Cluster specific options like user name, host names available GPUs. It creates the command line as well.
'
''' @version 20180705

''' Enumeration defining the possible Acceleration options. Needs to be defined before SolverConfig Enum!
Private Enum AccOptionType
	Single_Node' 0
	Distributed_computing' 1
	MPI_computing' 2
	Distributed_computing_And_MPI_computing' 3  = 1 + 2
	'Here we could define to make GPU Computing as number 4, would be more consistent, currently this is derived from the number of set GPUs
End Enum
Const mstrAccOptions As Variant = Array("Single_Node","Distributed_computing","MPI_computing","Distributed_computing And MPI_computing")

''' Enumeration representing The different available solver type, the list can be found in @link ../Online Help/mergedProjects/VBA_3D/special_vbasolver/special_vbasolver_solver_parameter_object.htm the help file @endlink of the SolverParameter CST VBA Class<br>
''' Enumeration is used in the Function #FillSolverTable as well!
Enum SolverType
	Transient_solver               ' HF: Time Domain (T)
	Frequency_domain_solver        ' HF: Frequency Domain (F)
	Eigenmode_solver               ' HF: Eigenmode (E)
	Integral_equation_solver       ' HF: Integral Equation (I)
	Asymptotic_solver              ' HF: Asymptotic (A)
	Multilayer_solver              ' HF: Multilayer (M)
	Partial_RLC_solver             ' LF: Partial RLC (RLC)
	Electrostatic_solver           ' LF: Electrostatic (Es)
	Magnetostatic_solver           ' LF: Magnetostatic (Ms)
	Stationary_current_solver      ' LF: Stationary Current (Js)
	LF_Frequency_domain_solver     ' LF: Frequency Domain (LF)
	LF_Time_domain_solver          ' LF: Time Domain (LF)
	Drift_Diffusion_solver         ' LF: Drift-Diffusion (DD)
	LF_Abaqus_CoSim                ' LF: Abaqus Co-Simulation
	PIC_solver                     ' Particles: Particle-in-Cell (Pic)
	Particle_tracking_solver       ' Particles: Particle Tracking (Trk)
	Electrostatic_PIC_solver       ' Particles: Electrostatic PIC (Es PIC)
	Wakefield_solver               ' Particles: Wakefield (Wak)
	Thermal_solver                 ' Thermal: Steady State (THs)
	Thermal_transient_solver       ' Thermal: Transient (THt)
	Conjugate_heat_transfer_solver ' Thermal: Conjugate Heat Transfer (CHT)
	Structural_mechanics_solver    ' Mechanics: Mechanical (ST)
	DS_run                         ' Simulation Tasks & SAM Workflows
	Optimizer                      ' Optimizer
	Parameter_sweep                ' Parameter Sweep
	Active_solver                  ' Active Solver
End Enum

''' Array with Strings corresponding to the Enum Names of the Solver Types in #SolverType
Const mstrSolverType As Variant = Array( _
	"Transient_solver", _
	"Frequency_domain_solver", _
	"Eigenmode_solver", _
	"Integral_equation_solver", _
	"Asymptotic_solver", _
	"Multilayer_solver", _
	"Partial_RLC_solver", _
	"Electrostatic_solver", _
	"Magnetostatic_solver", _
	"Stationary_current_solver", _
	"LF_Frequency_domain_solver", _
	"LF_Time_domain_solver", _
	"Drift_Diffusion_solver", _
	"LF_Abaqus_CoSim", _
	"PIC_solver", _
	"Particle_tracking_solver", _
	"Electrostatic_PIC_solver", _
	"Wakefield_solver", _
	"Thermal_solver", _
	"Thermal_transient_solver", _
	"Conjugate_heat_transfer_solver", _
	"Structural_mechanics_solver", _
	"DS_run", _
	"Optimizer", _
	"Parameter_sweep", _
	"Active_solver")

''' type defining properties for every solvertype
Public Type SolverConfig
	Solverkey As SolverType ' Solvertype with properties;
	SolverName As String' Clear name of Solverkey
	SolverType As String' Clear Name of SolverType
	IsSupported As Boolean' Boolean whether the solvertype is supported by the macro
	hasGPU As Boolean' Boolean whether the Solverkey is able to use GPU
	AccOption As AccOptionType' Which of Acceleration is possible?
	CommandLineOptions As String' string of Command Line Options to activate the Solverkey
	'could make handling of Type the same as enum
End Type
'' --- configuration of Table with solver Types, Names, Enums, Acceleration options and Command Line
Private aSolverTable() As SolverConfig 'Table with solver configurations
Private aSolverEnum() As SolverType ' SolverEnum defining an array of currently available solvers
Private aSolverChoicetoEnum() As Integer ' aSolverChoicetoEnum: reverse of SolverEnum resolving solver enum for currently available solvers

'' Enum representing supported Scheduler Types, default is set to Linux_Scheduler because Microsoft Scheduler might not be available

''' Scheduler definition

Public Enum SchedulerType
	Microsoft_HPC_Scheduler_2012
	Microsoft_HPC_Scheduler_2016
	Microsoft_HPC_Scheduler_2019
	Linux_Scheduler
End Enum
Const mstrSchedulerType = Array("Microsoft_HPC_Scheduler_2012", "Microsoft_HPC_Scheduler_2016", "Microsoft_HPC_Scheduler_2019", "Linux_Scheduler")

''' Selected Scheduler
Private mScheduler As SchedulerType ' Currently selected Scheduler Type ''redundant?
Private mobjScheduler As Object 'Scheduler child - by default this is #CST_Microsoft_HPC_Scheduler (set with #strScheduler Property);

''' Properties of Cluster
Public strSchedulerNode As String ' Node with installed Scheduler
Public strMainController As String ' DC Main Controller
Public strCstInstallPath As String ' Path of CST Installation on Cluster
Public strWorkDirectory As String ' Path of Temporary directory on Cluster

Public strUsername As String ' Job Owner / User name
Public strPassword As String ' User Password
Public strConsoleConnect as String
Public strConsoleConnects as Variant
Public strCopyToLocal as String
Public strCopyToLocals as Variant

Public strDefaultQueue As String
Public strAvailableQueues As Variant
Public strMachineFile As String ' Path to machine file

' -- Resources available on queue (set by Scheduler)
Public intQueueMaxGPUs As Integer ' Number of GPUs available
Public intQueueMaxNodes As Integer  ' Number of Nodes available
Private mQueueAccOption As AccOptionType ' Acceleration Type available in queue, handled by #fAccOption property; String value from #strAccOption property

' ----- Model specific settings
Private mstrModelFile As String ' String Path of current Model File
Private mInitSolver As SolverType ' Solver set Model
Private mSolver As SolverType ' Currently selected Solver


' -- Resources set by user in submit-dialog

Public strSetQueue As String ' Name of currently used queue
Public intSetGPUs As Integer' Number of GPUs used for Simulation
Public intSetNodes As Integer  ' Number of Nodes used for Simulation
Public intSetThreads As Integer  ' Complete Number of Threads used for Simulation


' ----- Set/Let and Get Functions of fixed Settings

''' ReadWrite Property
''' Kind of Scheduler System used by the cluster. <b>Needs to be one of #SchedulerType</b>
''' so currently CST_Microsoft_HPC_Scheduler or CST_Linux_Scheduler
''' @returns String with the currently used Scheduler

Property Let strScheduler(mstrScheduler As String)

	On Error GoTo notvalid
	Dim tmpObject As Object


	If Not(mobjScheduler Is Nothing) Then
		Set mobjScheduler = Nothing
	End If

	mScheduler = Eval(mstrScheduler)
	Set tmpObject = ModuleLoad(MacroDir + "\CST_" + mstrScheduler + ".cls",True)
	' alternative and maybe nicer: CallByName CSTCLusterConf "Scheduler" vbLet "strScheduler"

	If tmpObject Then
	   Set mobjScheduler = tmpObject
	   Set mobjScheduler.ClusterConf = Me
	   Log("Set Scheduler to " + mstrScheduler)
	Else
		UseLinuxScheduler()
	End If

ExitSub:
	Exit Property

notvalid:
	UseLinuxScheduler()
	Resume ExitSub

End Property

Property Get strScheduler() As String

	Select Case mScheduler
		Case Microsoft_HPC_Scheduler_2012
			strScheduler = "Microsoft_HPC_Scheduler_2012"
		Case Microsoft_HPC_Scheduler_2016
			strScheduler = "Microsoft_HPC_Scheduler_2016"
		Case Microsoft_HPC_Scheduler_2019
			strScheduler = "Microsoft_HPC_Scheduler_2019"
		Case Linux_Scheduler
			strScheduler = "Linux_Scheduler"
		Case Else ' fallback
			strScheduler = "Microsoft_HPC_Scheduler_2012"
	End Select
End Property

''' @brief Help function to set Linux Scheduler
Private Sub UseLinuxScheduler()

	Log("Defined Scheduler not valid using CST_Linux_Scheduler",1)
	mScheduler = Linux_Scheduler
	Set mobjScheduler = New CST_Linux_Scheduler
	Set mobjScheduler.ClusterConf = Me

	Log("Set Scheduler to Linux_Scheduler.")
End Sub

''' ReadWrite Property
''' Encrypted Password to login into cluster system (Name is innocent)
''' @returns String with Password on Cluster

Property Let strPassword1(Password As String)

	strPassword = Decrypt64$(Password,"cst")
	Log("Set password")
End Property
Property Get strPassword1() As String

	If Len(strPassword) = 0 Then
		strPassword1 = ""
	Else
		strPassword1 = Encrypt64$(Cstr(strPassword),"cst")
	End If
End Property

' ----- Let and Get Functions of variable settings - set by Scheduler

''' ReadWrite Property
''' Gets mQueueAccOption to a value of the Enum AccOptionType depending on if the queue has MPI_computing and/or Distributed_computing available by binary comparison; setting is done with setAccOption (needs to be a function, because Property allows only one input)
''' @returns Boolean whether MPI_computing and/or Distributed_computing is available for Queue and Solver

Property Get fAccOption(flag As AccOptionType) As Boolean

   Dim tmpAcc As Boolean

   tmpAcc = CBool(mQueueAccOption And aSolverTable(mSolver).AccOption And flag)
   Log(mstrAccOptions(flag) + IIf(tmpAcc," set and available"," not set/not available") + " for Cluster and Solver " + strCurrentSolver)
   fAccOption = tmpAcc

End Property

Friend Function setAccOption(flag As AccOptionType, setbit As Boolean)

   If setbit then
      mQueueAccOption =  mQueueAccOption Or flag
   Else
      mQueueAccOption =  mQueueAccOption And Not flag
   End If
   Log(IIf(setbit,"Activated ","Deactivated ") + mstrAccOptions(flag) + " on Cluster.")

End Function
''' ReadOnly Property
''' Currently set Acceleration Option for simulation, the value derived from the Array #mstrAccOptions with index derived from the Enum #AccOptionType
''' @returns String derived from Array mstrAccOptions set in the initialization Array index is the enum mQueueAccOption
Property Get strAccOption() As String

	strAccOption = mstrAccOptions(mQueueAccOption)
	Log("Current Acceleration Option " & strAccOption)
End Property


''' ReadOnly Property
''' If Distributed_computing or MPI_computing are chosen often a wrapper script initializing the Distributed_computing or MPI_computing subsystem is needed
''' @returns Boolean if Wrapper script is needed
Property Get fNeedsWrapper() As Boolean

   fNeedsWrapper =  fAccOption(Distributed_computing) Or fAccOption(MPI_computing)
End Property


' ----- (Let and) Get Functions of variable settings - set in Class Functions (e.g. using GUI)

''' ReadWrite Property
'''
''' Property with path to the current model file
''' @returns String with the Filename (full path) of the current CST-model
''' @sa strModelFile
Friend Property Let strModelFile(ModelFile As String)

	If BadFileName(ModelFile) then
		MsgBox "Filename contains invalid characters. Please don't use any of these characters: ~, ^, [, ], :, |, *, /, \, $, "", (, ), ;, comma or space in the filename."
		Exit All
	End If

	Log("Set Model file to " & ModelFile)

	mstrModelfile = ModelFile
End Property

Friend Property Get strModelFile() As String

	strModelfile = IIf(Len(mstrModelFile),mstrModelFile,Environ("Temp"))

End Property

''' ReadOnly Property
''' Current solver as string value; set by #DialogFunc
''' @returns String with the current Solver taken from mstrSolverType
''' @sa mSolver
Property Get strCurrentSolver() As String

	Dim mstrCurrentSolver as String

	mstrCurrentSolver = mstrSolverType(mSolver)
	Log("Current Solver is " & mstrCurrentSolver)

	strCurrentSolver = mstrCurrentSolver
End Property

''' ReadOnly Property
''' Current String Solver as Enum (or Integer) Value. Compare to SolverType for the corresponding values
''' @returns String  with the Enum Value of the current Solver
''' @sa mSolver
Property Get strCurrentSolverEn() As String
	Dim mstrCurrentSolverEn as String
	mstrCurrentSolverEn = Cstr(mSolver)
	Log("Current Solver En is " & mstrCurrentSolverEn)

	strCurrentSolverEn = mstrCurrentSolverEn
End Property

''' ReadOnly Property
''' determines if Cluster (or template or queue) and Current solver support GPUs
''' the number of available GPUs can be determined using the #NumGPU property
''' @returns Boolean if GPU can be used
''' @todo is this value on the user settings as well?
Property Get fWithGPU() As Boolean

	Dim tmpgpu As Boolean
	tmpgpu = (CBool(intQueueMaxGPUs) and aSolverTable(mSolver).hasGPU)
	Log("GPU" + IIf(tmpgpu,"set and available","not set/not available") + " for Cluster and Solver " + strCurrentSolver)

	fWithGPU = tmpgpu
End Property

''' ReadOnly Property
''' returns the command line in order to start cst, if a wrapper is needed this includes an acc switch
''' @returns String  with the cst Command Line
''' @todo determine if needs wrapper is really necessary
''' @todo implement ParameterSwitch
Friend Property Get CommandLine() As String

	If fNeedsWrapper Then
		CommandLine = quote(strCstInstallPath + "\CST DESIGN ENVIRONMENT.exe") + " " + CommandLineSwitches + " -defaultacc " + AccSwitch + GPUSwitch + ThreadSwitch + " -machinefile=" + strMachineFile + " """ + strModelFile  + ".cst""" 'full path of cst
	Else
		CommandLine = quote(strCstInstallPath + "\CST DESIGN ENVIRONMENT.exe") + " " + CommandLineSwitches + " -defaultacc " + GPUSwitch +  ThreadSwitch + """" + strModelFile  + ".cst"""
	End If
	Log("Build Command Line " + CommandLine,2)
End Property

''' ReadOnly Property
''' returns the command line switches depending on the current solver read from #aSolverTable
''' @returns Command Line options to activate solver
Friend Property Get CommandLineSwitches() As String

	If aSolverTable(mSolver).IsSupported Then
		CommandLineSwitches = aSolverTable(mSolver).CommandLineOptions
		Log("Command Line Switches are set to " + aSolverTable(mSolver).CommandLineOptions)
	Else
		Log("Solver not supported in Command Line",2)
		CommandLineSwitches  = ""
		Exit All
	End If

End Property

''' ReadOnly Property
''' returns the command line for GPU acceleration
''' @returns GPUwitch String with Command Line option activating GPU with a defined number of GPUs


Friend Property Get GPUSwitch() As String

	Dim strGPUSwitch As String
	If CBool(intSetGPUs) Then
		strGPUSwitch = "-withgpu=" + Cstr(intSetGPUs) + " "
	Else
		strGPUSwitch = ""
	End If
	Log("Set GPUSwitch to " + strGPUSwitch)
	GPUSwitch = strGPUSwitch
End Property

''' ReadOnly Property
''' returns the Command line for acceleration depending on #mQueueAccOption, sets strMainController if #Distributed_computing is selected
''' @returns String with Command Line option defining number of Threads
Friend Property Get ThreadSwitch() As String

	Dim strThreadSwitch As String

	If Not(Cbool(intSetThreads)) then intSetThreads = 1024

	strThreadSwitch = " --numthreads " + Cstr(intSetThreads) + " "

	Log("Set Thread Switch to " + strThreadSwitch)
	ThreadSwitch = strThreadSwitch
End Property


''' ReadOnly Property
''' returns the Command line for acceleration depending on #mQueueAccOption, sets strMainController if #Distributed_computing is selected
''' @returns String with Command Line option activating Acceleration Options
Friend Property Get AccSwitch() As String

	Dim strAccSwitch As String
	Select Case mQueueAccOption
		Case Distributed_computing : strAccSwitch = "-withdc=" + strMainController + " "
		Case MPI_computing : strAccSwitch = "-withmpi " 'Where do I set the number of nodes here?  I assume in the Machine file
		Case Else : strAccSwitch = ""
	End Select

	Log("Set Acceleration Switch to " + strAccSwitch)
	AccSwitch = strAccSwitch
End Property


''' ReadOnly Property
''' Function returning if Distributed_computing is activated either for a normal solver run or for parameter sweep and optimizer
''' @returns String with Command Line option activating Distributed_computing with a defined number of nodes
''' @todo where is the #MPI_computing command line returned?
Friend Property Get ParameterSwitch() As String

	Dim strParameterSwitch As String
	If (mQueueAccOption And Distributed_computing) Then
		If (mSolver = Parameter_sweep) Or (mSolver = Optimizer) Then
			strParameterSwitch = " -dcpar"
		Else
			strParameterSwitch = " -dc"
		End If
		strParameterSwitch = strParameterSwitch + CStr(intSetNodes) + " "
	Else
		strParameterSwitch = ""
	End If

	Log("Set Distributed Computing switch to " + strParameterSwitch)
	ParameterSwitch = strParameterSwitch
End Property

' ---- Class Initialization and Termination

''' Constructor of #CSTClusterConf
'''
''' Initialization does nothing specific so far
''' @sa Class_Terminate
Private Sub Class_Initialize()

	Debug.Print "Initializing CSTClusterConf."

End Sub

''' @brief Destructor of #CSTClusterConf
'''
''' frees the Memory of the Scheduler #mobjScheduler
''' \sa Class_Initialize
Private Sub Class_Terminate()

	If not mobjScheduler is nothing then

		Log("Terminating CSTClusterConf")
		Set mobjScheduler = Nothing

	End if
	If (strScheduler = "Microsoft_HPC_Scheduler_2012" or strScheduler = "Microsoft_HPC_Scheduler_2016" or strScheduler = "Microsoft_HPC_Scheduler_2019") Then
		quit
	End if



End Sub

Public Sub Configure()

	save

	strConsoleConnects = Array("Attach","Create new")
	strCopyToLocals = Array("False","True")


	Debug.Print "Checking for Cluster setup first."

	ReDim strAvailableQueues(1)

	If Len(strDefaultQueue) = 0 then
		strDefaultQueue ="0"
	End if

	CheckRequired


	MakeSolverTable
	mobjScheduler.Configure

End Sub
' ---- Settings GUI + Gui Function
''' Gui Sub for CSTCLusterConf
''' @toDo document in future
Sub ConfigGui()

	Const strNodeSelection = Array("Template") ', "Nodegroup")


	Begin Dialog ConfigDialog 560, 320, "Cluster Settings - Fill In All Fields", .ConfigDialogFunc
		GroupBox 20, 10, 520, 125, "Cluster settings"

		Text         35, 30, 300, 20, "Scheduler:", .SchedulerText
		DropListBox 270, 25, 250, 20, mstrSchedulerType(), .Scheduler

		Text     35, 55, 300, 20, "Hostname or IP of scheduler node:", .SchedulerNodeText
		TextBox 270, 50, 250, 20, .strSchedulerNode

		Text     35, 80, 300, 20, "CST installation path on cluster:", .CstInstallPathText
		TextBox 270, 75, 250, 20, .strCstInstallPath

		'Linux only
		Text     35, 105, 300, 20, "Working directory:", .WorkDirectoryText
		TextBox 270, 100, 250, 20, .strWorkDirectory

		GroupBox 20, 145, 520, 75, "User data"

		'Linux only
		Text     35, 165, 300, 20, "User name:", .UserNameText
		TextBox 270, 160, 250, 20, .strUsername

		'Linux only
		Text     35, 190, 300, 20, "Password:", .PasswordText
		TextBox 270, 185, 250, 20, .strPassword, -1

		'Windows only
		Text         35, 165, 300, 20, "Connection to compute node:", .ConsoleConnectText
		DropListBox 270, 160, 250, 20, strConsoleConnects(), .strConsoleConnect, 1

		'Windows only
		Text         35, 190, 300, 20, "Copy files locally:", .CopyToLocalText
		DropListBox 270, 185, 250, 20, strCopyToLocals(), .strCopyToLocal, 1

		'Test configuration
		GroupBox 20, 230, 520, 45, "Scheduler settings"

		Text         35, 250, 300, 20, "Available after pressing Test:", .defaultQueueText
		DropListBox 270, 245, 250, 20, strAvailableQueues(), .Queue

		OKButton      20, 285, 90, 21
		CancelButton 130, 285, 90, 21
		PushButton   450, 285, 90, 21, "Test", .GetQueues
	End Dialog

	Dim cdlg As ConfigDialog
	cdlg.Scheduler = mScheduler

	If CBool(Len(strDefaultQueue)) Then cdlg.Queue = Cint(strDefaultQueue)

	Dim mintdlg As Integer

	mintdlg = Dialog(cdlg)

	If mintdlg = 0 Then Exit All

	strScheduler = mstrSchedulerType(cdlg.Scheduler)
	WriteValuesToReg("Scheduler")

	Dim strSetting
	For Each strSetting In mobjScheduler.strRequiredSettings
		If Not(Len(strSetting) = 0) Then
			CallByName(Me,"str" + strSetting,vbLet,Eval("cdlg.str" + strSetting))
			WriteValuesToReg(strSetting)
		End If
	Next

	If Not (cdlg.Queue < 0) Then
		strDefaultQueue = Cstr(cdlg.Queue)
		WriteValuesToReg("DefaultQueue")
	End If


	If Not CheckRequired Then
		Exit Sub
	End If

	MakeSolverTable
	mobjScheduler.Configure

	If mintdlg = 1 Then ConfigGui 'TODO: If I press test, and CheckRequired finds something is missing it shows the ConfigGui (because CheckRequired is missing something) and then again because I pressed Test; I will leave this for now, because the configuration is only shown once anyway

End Sub

Private Function ConfigDialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Select Case Action%
	Case 1 ' Init of Dialog = default values

		If Not(mobjScheduler Is Nothing) Then

			SetSchedulerOptions

			Dim strSetting As String

			If (Not CBool(Len(strConsoleConnect))) Then strConsoleConnect = strConsoleConnects(1)
			If (Not CBool(Len(strCopyToLocal))) Then strCopyToLocal = strCopyToLocals(1)

			For Each strSetting In mobjScheduler.strRequiredSettings
				strSetting = "str" + strSetting
				DlgText strSetting,CallByName(Me,strSetting,vbGet)
			Next

			If (Not Len(strDefaultQueue) And Not IsEmpty(strAvailableQueues)) Then strDefaultQueue = "0"

			DlgText "defaultQueueText", IIf(Len(strDefaultQueue), "Select default " + LCase$(mobjScheduler.strSelectionLabel) + ":", "Fill in all fields")
			DlgVisible "Queue", CBool(Len(strDefaultQueue))
			DlgEnable "Queue", CBool(Len(strDefaultQueue))

			End If

	Case 2
		Select Case DlgItem$
		Case "Scheduler"
			strScheduler = mstrSchedulerType(SuppValue)
			SetSchedulerOptions

		End Select
	End Select
End Function

Private Sub SetSchedulerOptions()

	DlgValue "Scheduler", mScheduler
	Const OptionalItems = Array("Workdirectory","Username","Password","ConsoleConnect","CopyToLocal")',"NodeSelection")

	Dim OptionalItem As String

	For Each OptionalItem In OptionalItems
		DlgVisible "str" + OptionalItem,False
		DlgVisible OptionalItem + "Text",False
	Next

	For Each OptionalItem In mobjScheduler.strRequiredSettings

		DlgVisible "str" + OptionalItem,True
		DlgVisible OptionalItem + "Text",True
	Next

End Sub

' ---- Settings GUI + Gui Function


''' Function displaying Dialog in order to select the settings for the Simulation
''' Offered Solvers: If solver is supported by Macro and Scheduler, switch onChange
''' Other settings are:
''' - Distributed_computing/MPI_computing/Single Node
''' - Number of GPUS (if GPUs are available)
''' - Number of Nodes (for Distributed_computing and  MPI_computing)

Sub JobSettingsGui()


	Log("Constructing variable Job Settings Gui")
	Dim SolverChoice() As String
	Dim i As Integer, numSolverChoice As Integer, SolverIndex(100) As Integer

	' Fill Available Solver Vectors (Enums vector + Name vector)
	' Idea: omit first counting with
	'   ReDim Preserve astrItems(0 To lngIndex - 1)
	numSolverChoice = 0


	Log("Filling Array of supported Solvers from SolverTable")
	For i = LBound(aSolverTable) To UBound(aSolverTable)
		If aSolverTable(i).isSupported Then
			SolverIndex(numSolverChoice) = aSolverTable(i).SolverKey
			numSolverChoice += 1
		End If
	Next i


	ReDim SolverChoice(numSolverChoice-1)
	ReDim aSolverChoicetoEnum(numSolverChoice-1)
	ReDim aSolverEnum(UBound(aSolverTable))

	For i = 0 To numSolverChoice-1
		aSolverEnum(SolverIndex(i)) = i
		aSolverChoicetoEnum(i) = SolverIndex(i)
		SolverChoice(i) = aSolverTable(SolverIndex(i)).SolverName
	Next i

	Log("Defining user Dialog for Job Submission.")
	Begin Dialog UserDialog 440,270,"Interact With Scheduler",.JobDialogFunc ' %GRID:10,7,1,1
		GroupBox 10,10,420,75,"General"

		Text 25,30,130,20,"Choose " + LCase$(mobjScheduler.strSelectionLabel) + ":",.Queuelabel
		DropListBox 210,30,200,30,strAvailableQueues(),.QueueChoice

		Text 25,55,100,20,"Choose solver:"
		DropListBox 210,55,200,30,SolverChoice(),.SolverChoice

		' Acceleration may be done with Distributed_computing/MPI_computing and/or GPUs
		GroupBox 10,93,420,135,"Acceleration options"
		OptionGroup .AccOption
			OptionButton  25,107,100,30,"Single node"
			OptionButton 145,107,50,30,"DC",.DCChoice ' available if set in SolverTable
			OptionButton 215,107,50,30,"MPI",.MPIChoice ' available if set in SolverTable

		Text 25,147,200,20, "Number of GPUs per node:", .GPU
		Text 25,172,200,20, "Number of nodes:", .Nodes
		Text 25,197,200,20, "Number of threads:", .Threads

		TextBox 230,147,30,20, .numGPUs
		TextBox 230,172,30,20, .numNodes
		TextBox 230,197,30,20, .numThreads

		OKButton 10,240,90,21
		CancelButton 110,240,90,21
		PushButton 300,240,130,21,"Cluster Settings... ",.ClusterConfig

	End Dialog
	Dim dlg As UserDialog


	If (Len(strDefaultQueue) = 0) Then strDefaultQueue = "0"
	dlg.QueueChoice = CInt(strDefaultQueue)
	dlg.SolverChoice = aSolverEnum(mSolver)
	dlg.numNodes = "1"
	dlg.numThreads = "0"
	dlg.numGPUs = "0"

	Select Case Dialog(dlg)
		Case 0
			Exit All
		Case 1
			ConfigGui
			JobSettingsGui
			Exit Sub
	End Select

	'only one Acceleration option can be true, so either single, MPI_computing or Distributed_computing
	mQueueAccOption = dlg.AccOption

	If mQueueAccOption Then
		intSetNodes = min(CInt(dlg.numNodes),intQueueMaxNodes)
		Log("Defined " +mstrAccOptions(mQueueAccOption) + " with " + CStr(intSetNodes) + " Nodes in GUI")
	End If

	intSetThreads = CInt(dlg.numThreads)
	Log("Defined total " + CStr(intSetThreads) + " Threads in GUI")

	' check if GPUs set
	If (dlg.numGPUs <> "0") Then

		intSetGPUs =  min(CInt(dlg.numGPUs),intQueueMaxGPUs)
		Log("Defined " + Cstr(intSetGPUs) + " GPUs in Gui")

	End If

	Log("Defined " + aSolverTable(mSolver).SolverName & " in GUI")

	mobjScheduler.Submit

End Sub

''' @brief Function defining actions to be taken when something in dialog changes
'''
''' - if the solver Choice is varied
''' 	- all solver options are set to the new choice using #SetSolverOptions
''' - if the acceleration option is changed
''' 	- The Node Selection is changed to the boolean value of SuppValue
''' 	- 0 for #Single_Node = false
''' 	- 1 for #Distributed_computing or 2 for #MPI_computing = true
''' - in future there will be a sub menu for the cluster configuration which is opened here
''' @param DlgItem$ selected dialog item
''' @param Action% Action taken
''' @param SuppValue?	selected Value of DlgItem$
''' @returns true if dialog was successful
Private Function JobDialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

    Select Case Action%
	Case 1
		Log("Setting Dialog to default values")
		SetSolverOptions(mSolver)

	Case 2  'Value changing or button pressed
		Select Case DlgItem$
		Case "SolverChoice"
			Log("Varied Solver Choice.")
			SetSolverOptions(aSolverChoicetoEnum(SuppValue))
		Case "QueueChoice"
			Log("Varied Queue.")
			mobjScheduler.strQueue = strAvailableQueues(SuppValue)
			SetSolverOptions(mSolver)
		Case "AccOption"
			Log("Changed Acceleration Option")
			DlgEnable "Nodes", SuppValue
			DlgEnable "numNodes", SuppValue
			DlgText "numNodes", IIf(SuppValue,"2","1")
		End Select

	End Select

End Function
''' Routine enabling/disabling options after solver selection
''' - if a parameter Sweep or an Optimizer run is set the solver is temporarily set to the initial solver
''' - depending on the activated solver the functions #WithDC, #WithMPI and #WithGPU set the corresponding values
''' @param SolverKey selected Solver Key
''' @todo check if the node selection is available if DC or MPI are enabled
''' @todo check what must be done in case of parameter sweep or optimizer
Private Sub SetSolverOptions(ByVal SolverKey As SolverType)

	If (SolverKey = Parameter_sweep) Or (SolverKey = Optimizer) Then
		mSolver = mInitSolver
		Log("Solver Type is Parameter sweep or Optimizer, using initial Solver for acceleration options.")
	Else
		mSolver = SolverKey
	End If
	Log("Set Solver to " + CStr(aSolverTable(mSolver).SolverName))
	Log("ADJUST: Enabling Acceleration option " + mstrAccOptions(mQueueAccOption) + " in Dialog")

	DlgEnable "AccOption", (fAccOption(Distributed_computing) Or fAccOption(MPI_computing))

	' Default Acceleration option: Single PC + deactivate NodeSelection
	DlgValue "AccOption", 0
	DlgEnable "Nodes", False
	DlgEnable "numNodes", False
	intSetNodes = 1

	intSetThreads = 0

	' Select Options according to current settings
	DlgEnable "DCChoice", fAccOption(Distributed_computing)
	DlgEnable "MPIChoice",  fAccOption(MPI_computing)

	DlgEnable "GPU", fWithGPU
	DlgEnable "numGPUs", fWithGPU

	If (SolverKey = Parameter_sweep) Or (SolverKey = Optimizer) Then
		mSolver = SolverKey
	Log("Set Solver to" + CStr(aSolverTable(mSolver).SolverName))
	End If

End Sub

' ----- Class Functions



''' Public Routine to Submit a job<br>
''' Here The submit function of the current scheduler is called
Public Sub Run()

	Save
	strModelFile = GetProjectPath("Project")
    If InStr(strModelFile, "Untitled") <> 0 Then
    	MsgBox "Please assign a name to your model before submitting it."
    	Exit All
    End If


	mobjScheduler.Run

	If (GetApplicationName() = "DS") Then
		mInitSolver = DS_run
	Else
		mInitSolver = GetSolverKey(eval("GetSolverType"))
	End If

	mSolver = mInitSolver

	JobSettingsGui

	Exit Sub



End Sub

' ---- SolverConfiguration

''' @brief Subroutine Changing the Supported Solvers
'''
''' Sub functions can use this function to select only available solvers. This is the case if a queue is used for special solvers only. Insert the available solvers as string array.<br>
''' <b>The strings need to be equivalent to the SolverType Enum Values.</b>
''' @param astrSolver() String array with Enum Values of Supported solvers

Friend Sub ChangeSupported(astrSolver() As String)

	Log("Changing the Supported Solvers")
	dim i As integer

	For i = LBound(aSolverTable) To Ubound(aSolverTable)
		aSolverTable(i).isSupported = false
	Next i

	Log("These Solver are Supported")
	For i = LBound(astrSolver) To Ubound(astrSolver)
		aSolverTable(eval(astrSolver(i))).isSupported = true
		Log(astrSolver(i))
	Next i

End Sub

''' Subroutine Filling one element of the Solver definition Table
'''
''' @param SolverKey Enum Value of the solver type
''' @param StrSolverName String value of solver
''' @param strSolverType String Value of SolverType
''' @param fIsSupported Switch if Solver is supported by macro
''' @param fhasGPU Switch if Solver supports GPU
''' @param AccOption Switch if Solver supports DC or MPI
''' @param strCommand String with command line options to activate the solver

Private Sub FillSolverTable(SolverKey As SolverType, strSolverName As String, _
		strSolverType As String, fIsSupported As Boolean, fhasGPU As Boolean, _
		AccOption As AccOptionType, strCommand As String)

	aSolverTable(SolverKey).SolverKey = SolverKey
	aSolverTable(SolverKey).SolverName = strSolverName
	aSolverTable(SolverKey).SolverType = strSolverType
	aSolverTable(SolverKey).IsSupported = fIsSupported
	aSolverTable(SolverKey).hasGPU = fhasGPU
	aSolverTable(SolverKey).AccOption = AccOption
	aSolverTable(SolverKey).CommandLineOptions = strCommand

	Log("Added " + IIF(fIsSupported,"", "not ") + "supported " + strSolverName+ " with" + IIF(fhasGPU,"","out") + " GPU; " + _
	IIF(AccOption and Distributed_computing,"DC","no DC") + IIF(AccOption = Distributed_computing_And_MPI_computing, " and " , " but ") + _
	IIF(AccOption and MPI_computing,"MPI","no MPI") + " possible; Command Line: " + strCommand)

End Sub

''' Subroutine Filling the Solver definition Table<br>
''' The Routine fills the SolverTable in the following order:
''' - SolverKey As SolverType
''' - strSolverName As String
''' - strSolverType as String
''' - fIsSupported As Boolean
''' - fhasGPU As Boolean
''' - AccOption as AccOptionType: Single_Node, Distributed_computing, MPI_computing, Distributed_computing_And_MPI_computing
''' - strCommand As String
'''
''' @todo <b>This must be adjusted for changes in future versions of CST</b> Refer to Enums from "CST STUDIO SUITE | Visual Basic (VBA) Language | 2D/3D Simulation VBA | VBA Objects Overview | Solver | SolverParameter" (alternatives would be the VBA solvertypes from "CST DS VBA | SimulationProject Object" or "2D/3D Simulation VBA | VBA Objects Overview | Global | Project | ChangeSolverType") and Command line options from ../Online%20Help/advanced/commandlineoptions.htm Adjust #SolverKey as well. strSolverType is taken from "2D/3D Simulation VBA | VBA Objects Overview | Global | Project | ChangeSolverType"

Private Sub MakeSolverTable()

	Log("Filling Solver Table")
	Redim aSolverTable(Ubound(mstrSolverType)+1)

	FillSolverTable(Transient_solver, "HF: Time Domain", "HF Time Domain", true, true, Distributed_computing_And_MPI_computing, "-m -r")
	FillSolverTable(Frequency_domain_solver, "HF: Frequency Domain", "HF Frequency Domain", true, true, Distributed_computing_And_MPI_computing, "-m -f")
	FillSolverTable(Eigenmode_solver, "HF: Eigenmode", "HF Eigenmode", true, false, Distributed_computing, "-m -e")
	FillSolverTable(Integral_equation_solver, "HF: Integral Equation", "HF IntegralEq", true, true, Distributed_computing_And_MPI_computing, "-m -q")
	FillSolverTable(Asymptotic_solver, "HF: Asymptotic", "HF Asymptotic", true, true, Distributed_computing, "-m -a")
	FillSolverTable(Multilayer_solver, "HF: Multilayer", "HF Multilayer", true, true, Distributed_computing, "-m -ml")
	FillSolverTable(Partial_RLC_solver, "LF: Partial RLC", "PartialRLC", true, false, Distributed_computing, "-s -prlc")
	FillSolverTable(Electrostatic_solver, "LF: Electrostatic", "LF EStatic", true, false, Distributed_computing, "-s -se")
	FillSolverTable(Magnetostatic_solver, "LF: Magnetostatic", "LF MStatic", true, false, Distributed_computing, "-s -sh")
	FillSolverTable(Stationary_current_solver, "LF: Stationary Current", "LF Stationary Current", true, false, Distributed_computing, "-s -sj")
	FillSolverTable(LF_Frequency_domain_solver, "LF: Frequency Domain", "LF Frequency Domain", true, false, Distributed_computing, "-s -lf")
	FillSolverTable(LF_Time_domain_solver, "LF: Time Domain", "LF Time Domain (MQS)", true, false, Distributed_computing, "-s -ltmqs")
	FillSolverTable(Drift_Diffusion_solver, "LF: Drift-Diffusion", "Stationary Drift-Diffusion", true, false, Distributed_computing, "-s -sdd")
	FillSolverTable(LF_Abaqus_CoSim, "LF: Abaqus Co-Simulation", "", true, false, Distributed_computing, "-s -cse")
	FillSolverTable(PIC_solver, "Particles: Particle-in-Cell", "PT PIC", true, true, Distributed_computing, "-t -pic")
	FillSolverTable(Particle_tracking_solver, "Particles: Particle Tracking", "PT Tracking", true, false, Distributed_computing, "-t -tp")
	FillSolverTable(Electrostatic_PIC_solver, "Particles: Electrostatic PIC", "PT E-Static PIC", true, true, Distributed_computing, "-t -espic")
	FillSolverTable(Wakefield_solver, "Particles: Wakefield", "PT Wakefields", true, false, Distributed_computing_And_MPI_computing, "-t -tw")
	FillSolverTable(Thermal_solver, "Thermal: Steady State", "Thermal Steady State", true, false, Distributed_computing, "-mp -ht")
	FillSolverTable(Thermal_transient_solver, "Thermal: Transient", "Thermal Transient", true, false, Distributed_computing, "-mp -httd")
	FillSolverTable(Conjugate_heat_transfer_solver, "Thermal: Conjugate Heat Transfer", "Conjugate Heat Transfer", true, true, Distributed_computing, "-mp -cht")
	FillSolverTable(Structural_mechanics_solver, "Mechanics: Mechanical", "Mechanics", true, false, Distributed_computing, "-mp -mech")
	FillSolverTable(DS_run, "Simulation Tasks & SAM Workflows", "", true, true, Distributed_computing_And_MPI_computing, "-c -u")
	FillSolverTable(Optimizer, "Optimizer", "", true, true, Distributed_computing_And_MPI_computing, "-o")             ' settings depend on solver auto detection
	FillSolverTable(Parameter_sweep, "Parameter Sweep", "", true, true, Distributed_computing_And_MPI_computing, "-p") ' settings depend on solver auto detection
	FillSolverTable(Active_solver, "Active Solver", "", true, true, Distributed_computing_And_MPI_computing, "-as")    ' settings depend on solver auto detection

	Dim mstr
	Dim i as Integer
	For i = LBound(mstrSolverType) to Ubound(mstrSolverType)
		if Len(aSolverTable(i).SolverName) = 0 then
			Log("Missing Information for " + mstrSolverType(i) + ". Please check function FillSolverTable.",3)
		end if
	next i

	Log("Finished Filling Solver Table")
End Sub

Private Function GetSolverKey(strSolverType As String) As SolverType

	Dim Solver
	Debug.Print strSolverType
	For Each Solver In aSolverTable
		If Solver.SolverType = strSolverType Then GetSolverKey = Solver.SolverKey

	Next Solver

End Function

''' @brief logging:  depending on the loglevel dump to message, warning or debug log
Friend Sub Log(logtext As String, Optional loglevel As Integer = 3)
	Dim strCaller As String
	Dim debugprefix As String

	strCaller = CallersLine(-3)
	debugprefix = "[" & Cstr(Now) & "] " &  Mid$(strCaller , InStr(strCaller, "|" )+1, InStr(strCaller, "#") -InStr(strCaller, "|" )-1 ) & " : "

	' output to vba debug window
	Debug.Print debugprefix & logtext
	' send log info to modeler
	HPCReporting(debugprefix & logtext, loglevel)
End Sub

''' @brief Subroutine checking if necessary variables are set
'''
''' shows settings GUI if something is missing, required variables depend on selected scheduler

Friend Function min(value1 As Integer,value2 As Integer) As Integer

	Dim inttmp As Integer

	inttmp = IIf(value1<value2,value1,value2)

	Log("Minimum value of " + CStr(value1) + " and " + CStr(value2) + " is " + Cstr(inttmp),3)

	min = inttmp
End Function

''' Reads Settings from Registry and assigns variables
Friend Sub ReadValuesFromReg()
	Log("Reading Values from Registry.")

	Dim Settings As Variant
	Settings = GetAllSettings("CST STUDIO SUITE", "ClusterIntegration")
	If Not(IsEmpty(Settings)) Then
		Dim I As Integer
		For I = LBound(Settings) To UBound(Settings)

		Dim valName As String
		valName = Settings(I,0)
		'backward compatibility
		If valName = "Password" Then
			valName = "Password1"
		End If

		Dim valData As String
		valData = Settings(I,1)
		'backward compatibility
		If valData = "Microsoft_HPC_Scheduler" Then
			valData = "Microsoft_HPC_Scheduler_2012"
		End If

		CallByName(Me, "str" + valName, vbLet, valData)
		Next I
	End If
End Sub

Friend Sub WriteValuesToReg(RegKey As String)

	Log("Writing " + RegKey + " to Registry.")

	Dim RegValue As String

	If RegKey = "Password" Then RegKey="Password1"

	RegValue = CallByName(Me,"str" + RegKey,vbGet)
	If Len(RegValue) Then
		SaveSetting "CST STUDIO SUITE", "ClusterIntegration", RegKey, RegValue
	End If
End Sub

Friend Function CheckRequired() As Boolean

	ReadValuesFromReg

	If mobjScheduler Is Nothing Then
		strScheduler = "Microsoft_HPC_Scheduler_2012"
		ConfigGui
	End If

	Dim strRequiredSetting As String

	For Each strRequiredSetting In mobjScheduler.strRequiredSettings
		If Len(CallByName(Me,"str" + CStr(strRequiredSetting),vbGet)) = 0 Then ConfigGui
    Next

	CheckRequired = true

End Function
' ----- Help Functions

Friend Function quote(strVariable As String) As String

    quote = Chr$(34) + strVariable + Chr$(34)
End Function


''' @brief warns if filename has bad characters
Friend Function BadFileName(ByVal strFilename As String) As Boolean


	Dim fso As Object
	Set fso = CreateObject("Scripting.FileSystemObject")

	strFilename = fso.GetFileName(strFilename)
	BadFileName = False
	Const mstrForbiddens = Array("~","^","]","[",":","|","*","/","\","$","""","(",")",";"," ")

	Dim mstrForbidden As String

	For each mstrForbidden in mstrForbiddens
		If InStr(strFilename,mstrForbidden) Then
			BadFileName = True
			set fso = nothing
			Exit Function
		End If
	 Next
	set fso = nothing
End Function
