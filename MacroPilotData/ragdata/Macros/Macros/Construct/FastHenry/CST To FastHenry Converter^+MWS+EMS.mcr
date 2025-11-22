'#Language "WWB-COM"

' ================================================================================================
' Copyright 2008-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 29-Nov-2018 fsr: Removed BrowseForFolder and replaced it with GetFolder_Lib from vba_globals_all
' 12-Dec-2013 imu: Introduced configuration file, user can choose to always equivalence (or always not to equivalence) nodes of curves without corresponding bricks
' 20-Nov-2013 imu: Portunus v2.0 export release
' 20-Nov-2013 imu: Introduced checks for no bricks, empty curves, and zero number of segments, nodes or ports; dealing with no result from the FH run
' 15-Nov-2013 imu: Improved brick detection, works now also for nonstandard numbered brick nodes
' 11-Nov-2013 imu: Corrected bug with polygon sorting
' 08-May-2013 imu: Corrected bugs with "+parallel"
' 08-May-2013 imu: No Portunus calls, until released
' 21-Feb-2013 imu: Corrected pin numbers in function "determine_pins_nodes"
' 19-Feb-2013 imu: Increased number of nodes and curves
' 15-Oct-2012 imu: Replacement for all forbidden Portunus characters
' 25-Sep-2012 imu: Added frequency-independent value of R and L in the Portunus file
' 23-Jul-2012 imu: Corrected number of pins, nodes and cells in the Portunus output file
' 11-Jul-2012 imu: No ports for "equivs"
' 10-Jul-2012 imu: Added check for empty curve name; corrected bug with automatically equivalenced nodes
' 25-Jun-2012 imu: Corrected bugs in Portunus model generation
' 15-Apr-2012 imu: Generation of the Portunus network model; TODO: Determine node names in determine_pins_nodes ? Add user def. of frequencies?
'               TODO: BrowseForFolder? GetFolderName? Select one of the two existing
' 02-Jan-2012 imu: Added automatic call of FastHenry and choice of the result files
' 05-Dec-2011 imu: Introduced equivalence of nodes; Added alphabetic ordering of the curves
' 02-Aug-2011 fsr: replaced obsolete 'vba_globals.lib' with 'vba_globals_all.lib' and 'vba_globals_3d.lib'
' 25-Nov-2009 imu: Corrected bug related to the projected ports and traces on groundplanes
' 22-Dec-2008 imu: Added: ports of length 0; equivalence for "+parallel" curves;
'                 component names as comments in the FH file; check for traces scale factor;
'                 more correctness checks
' 26-Nov-2008 imu: First version

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

' ====== PUBLIC CONSTANTS
Public Const  MAX_CURVES = 99999
Public Const  MAX_NODES =  99999
Public Const  MAX_GND = 99

Public USER_EPS As Double

Public Const CHECK_PARALL_CURVES = False
Public Const ADD_EQUIVNODES = True

' ====== CONSTANTS FOR PORTUNUS OUTPUT FILE GENERATION
Const NMAX = 2
'Const PORTUNUS_Version = "1.0"
Const PORTUNUS_Version = "2.0"
Public nfrequencies As Long

' TO ELIMINATE
Public Const SORT_EDGES = True

' ====== NEW TYPES
Type point
	x As Double
	y As Double
	z As Double
End Type

Type point_ext	' "Extended" point type
	x As Double
	y As Double
	z As Double
	taken As Long	' How many times was the node taken in "series" type of curves?
	nname As String	' Name of the component to which this node corresponds
	hasport As Boolean	' Is this node the end of a port?
End Type

Type pnt
	coord(3) As Double
End Type

Type brick_mf_single	' Brick mid faces for a single brick and a single pair of opposite faces
	p1 As point
	p2 As point
	dist As Double
	wx As Double
	wy As Double
	wz As Double
	w As Double
	h As Double
End Type

Type brick_prop
	bname As String
	sigma As Double
End Type

Type curve_single	' Definition of curves made of several segments
	typ As String	' "Series" or "Parallel"
	cName As String	' Curve:polygon name
	nsegs As Long	' Number of segments
	point_ini() As Long	' Initial point of segment; indexing starts at 1
	point_fin() As Long	' Final point of segment; indexing starts at 1
	csName() As String	' Name of the segments
	leng() As Double		' Length of the curve segment
End Type

' File names
Public ResultDir As String
Public FH_inp_name As String, FH_result_name As String, Portunus_result_name As String
Public FH_options As String


Public FH_bricks() As brick_mf_single
' FH_bricks(solid, face_pair)   with solid starting at 0, face_pair starting at 1
Public FH_brick_properties() As brick_prop
Public nbricks As Long
Public bricks_found() As Long

Public comp_names() As String
Public solid_names() As Long

Public FH_curves() As curve_single
' FH_curves (curveno)   with curveno indexing starting at 0
Public ncurves As Long

Public curve_nodes(MAX_NODES) As point_ext
' curve_nodes(node) with node indexing starting at 1
Public nnodes As Long
Dim nsegs_tot As Long

Public nequivs As Long	' Indexing starts at 1
Public equiv_nodes(MAX_NODES, 2) As Long

' ======== FINAL DATA
Public nedges As Long
Public edges(MAX_NODES, 2) As Long	' Initial and final nodes of edges
Public sorted_edges(MAX_NODES) As Long	' Edge numbers, sorted according to the name of their component
Public edge_types(MAX_NODES) As String
Public edge_names(MAX_NODES) As String
Public edge_widths(MAX_NODES) As Double	'
Public edge_heights(MAX_NODES) As Double	'
Public edge_wx(MAX_NODES) As Double, edge_wy(MAX_NODES) As Double, edge_wz(MAX_NODES) As Double
Public edge_sigma(MAX_NODES) As Double

'Public edge_nhinc(MAX_NODES) As Integer, edge_nwinc(MAX_NODES) As Integer
'Public edge_rh(MAX_NODES) As Double, edge_rw(MAX_NODES) As Double
Public edge_nhinc As Integer, edge_nwinc As Integer
Public edge_rh As Double, edge_rw As Double

Public nports As Long
Public ports(MAX_NODES,2) As Long		' Contains numbers of nodes that are considered as ports
Public portnames(MAX_NODES) As String

Type gnd	' Groundplanes
	P(3) As point	' Points characterizing the groundplane
	normal As point	' Normal vector to the plane
	gname As String
	thick As Double
	sigma As Double
	typ As String	' "Uniform" or "Nonuniform"
	init_grid As String		' "None", "Uniform", "Meshed"
	init_grid_segs(2) As Integer	'Number of segments on x and y for the initial discretization
	refine_traces As Boolean
	trace_refinement_scale_factor As Double
	refine_ports As Boolean
	port_refinement_ratio As Double
End Type

Public ngrounds As Long
Public groundplanes(MAX_GND) As gnd	' Indexcing starts at 0

Public nprojtraces As Long
Public projtraces(MAX_NODES, 2) As point		' Nodes of projected traces
Public projtraces_w(MAX_NODES) As Double

Dim nprojports As Long
Public projports(MAX_NODES) As point		' Nodes = projected ports
Public projports_w(MAX_NODES) As Double
Public projportnames(MAX_NODES) As String

' For port user dialog
Public crtportDlg As Long
Dim portwidths_txt() As String
Dim portwidth_dlgtext() As String
Dim portlist() As String
Dim nptot As Long

Dim do_equivalence As Boolean, never_ask_equiv As Boolean	' = -1 if user was never asked


Sub Main
	Dim tmp As Double, nshapes As Double

	edge_nhinc = 1: edge_nwinc = 1
	edge_rh = 1.0: edge_rw = 1.0
	USER_EPS = 5e-3

	never_ask_equiv = False
	do_equivalence = False

	read_config_file()


	Begin Dialog UserDialog 460,301,"User definitions",.DialogFunc_userdef ' %GRID:10,7,1,1
		Text 20,14,160,14,"Geometric accuracy",.Text1
		TextBox 210,14,90,21,.acc
		GroupBox 10,35,430,231,"FastHenry settings",.GroupBox1
		CheckBox 20,56,130,14,"Run FastHenry",.RunFH
		CheckBox 180,56,240,14,"Generate Portunus network model",.GeneratePortunus
		Text 30,77,120,28,"FH command line options",.Text6
		TextBox 150,77,180,21,.FH_options
		CheckBox 20,245,200,14,"Remember filename settings",.RememberSettings
		Text 60,175,80,14,"FH file name",.Text2
		Text 20,119,120,14,"Results Directory",.Text5
		Text 20,196,120,14,"FH result file name",.Text3
		Text 20,217,120,14,"Portunus file name",.Text4
		TextBox 150,175,180,21,.FHInFile
		TextBox 150,112,180,49,.ResultDir,1
		TextBox 150,196,180,21,.FHOutFile
		TextBox 150,217,180,21,.PortunusOutFile
		PushButton 340,175,90,21,"Browse",.Browseinputfile
		PushButton 340,196,90,21,"Browse",.Browseinputfile2
		PushButton 340,217,90,21,"Browse",.Browseinputfile3
		PushButton 340,119,90,21,"Browse",.Browseinputfile4

		OKButton 10,273,90,21
		CancelButton 100,273,90,21
	End Dialog

	Dim dlg As UserDialog

	dlg.acc = CStr(USER_EPS)
	dlg.RunFH = 1
	dlg.GeneratePortunus = 1
	dlg.ResultDir = GetProjectPath("Model3D")	'Dirname(GetProjectbasename)
	dlg.FHInFile = "FastHenry.inp"
	dlg.FHOutFile = "Zc.mat"
	dlg.PortunusOutFile = "portunus_out.c2p"

	If (Dialog(dlg) = 0) Then Exit All

	'If dlg.RunFH = 0 Then dlg.GeneratePortunus = 0
tmp = Timer

	ResultDir	 	 	 = dlg.ResultDir
	If Right(ResultDir,1) <> "\" Then ResultDir = ResultDir + "\"
	FH_inp_name 	 	 = ResultDir + dlg.FHInFile
	FH_result_name 		 = ResultDir + dlg.FHOutFile
	Portunus_result_name = ResultDir + dlg.PortunusOutFile
	FH_options 			 = " " + dlg.FH_options
	' check_FH_options()

	ClearGlobalDataValues
	USER_EPS = Val(dlg.acc)

	' Frequencies
	' ===End user dialog

	nshapes = Solid.GetNumberOfShapes
	If nshapes = 0 Then
		MsgBox "No bricks found. Macro stops", vbCritical, "Error"
		Exit All
	End If
	Prepare_Step_File

	nports = 0

	read_solids()
	read_curves()
	determine_ports()
	generate_FH_data_structures()
	write_FH()

	If dlg.RunFH = 1 Then run_FH (dlg.GeneratePortunus)
		'run_FH ("0")
	If StrComp(Dir(FH_inp_name), "") = 0 Then  FileCopy GetProjectPath("Model3D") & dlg.FHInFIle,  FH_inp_name
	If (dlg.RunFH = 1 And StrComp(Dir(FH_result_name), "") = 0) Then FileCopy GetProjectPath("Model3D") & "Zc.mat", FH_result_name	' Also delete old Zc.mat ??
	If (dlg.GeneratePortunus = 1 And dlg.RunFH = 1 And StrComp(Dir(Portunus_result_name), "") = 0) Then FileCopy GetProjectPath("Model3D") + dlg.PortunusOutFile, Portunus_result_name

	Dim tmpstr As String

	tmpstr = "FastHenry file generation successful." + vbCrLf + vbCrLf + "FH input file: " + vbCrLf + FH_inp_name
tmp = Timer-tmp
'MsgBox ("Time: " + CSTr(tmp))

	If dlg.runFH = 1 Then
		tmpstr = tmpstr + vbCrLf + vbCrLf + "FH output file: " + vbCrLf + FH_result_name
		If  dlg.GeneratePortunus = 1 Then tmpstr = tmpstr + vbCrLf + vbCrLf + "Portunus output file: " + vbCrLf + Portunus_result_name
	Else
		If dlg.GeneratePortunus = 1 Then tmpstr = tmpstr +vbCrLf + vbCrLf + "Portunus file cannot be generated without running FH first"
	End If
	MsgBox tmpstr
														'GetProjectPath("Model3D") + "Fasthenry.inp"
End Sub

Sub read_config_file()
	Dim cfg_file As String, lib_dummy As String
	Dim string_item(30) As String
	Dim cst_nitems As Integer

	cfg_file = Dir$(GetMacroPath + "\Construct\FastHenry\CST2FH.cfg")
	If cfg_file = "" Then Exit Sub

	cfg_file = GetMacroPath + "\Construct\FastHenry\CST2FH.cfg"
	Open cfg_file For Input As #99
	' Parse the file
	While Not EOF(99)
		Line Input #99, lib_dummy
		cst_nitems = CSTSplit(lib_dummy, string_item)
		Select Case string_item(0)
			Case "NEVER_ASK_EQUIV"
				If string_item(1) = "True" Then	never_ask_equiv = True
			Case "DO_EQUIVALENCE"
				If string_item(1) = "True" Then	do_equivalence = True
		End Select
	Wend

	Close #99

End Sub


Sub check_FH_options()
' FH options:
' - (dash) Forces input to be read from the standard input.
' -s {ludecomp | iterative} - Specifies the matrix solution method used to solve the linear system
' -m {direct | multi} - Specifies the method to use to perform the matrix-vector product For the iterative algorithm
' -p {on | off | loc | posdef | cube | seg | diag | shells } - Specifies the method to precondition the matrix to accelerate iteration convergence
' -o n - Specifies n as the order of multipole expansions. Default is 2.
' -l {n | auto} - Specifies n as the number of partitioning levels for the multipole algorithm.
' -f {off | simple | refined | both | hierarchy } - Switches FastHenry to visualization mode only and specifies the type of FastCap generic file to make
' -g {on | off | thin | thickg - controls appearance of the ground plane when using the -f Option.
' -a {on | offg - on allows the multipole algorithm to automatically refine the structure As Is necessary To maintain accuracy In the approximation
' -i n - Specifies n as the level for initial refinement
' -d {on | off | mrl | mzmt | grids | meshes | pre | a | m | rl | ls}   - dump certain internal matrices To files. The Format of some of the files can be specified with the -k Option.
' -k {matlab | text | both} - Specifies type of file to dump with the -d option
' -t rtol, -b atol - Specifies the tolerance for iteration error.
' -c n - n = maximum number of iterations to perform for each solve (column). Overrides the Default of 200.
' -D {on | off} - Controls the printing of debugging information.
' -x portname - Specifies that only the column in the admittance matrix specified by portname should be computed. Multiple -x specifications can be used.
' -S suffix - This adds the string suffix to all filenames for this run. For instance,-S blah will produce the Output file Zc blah.mat.
' -r order - Specifies a reduced order model of the system as output. The size of the model will be order*number-of-ports.
' -M - If -r is specified with a nonzero order, then this option will cause FastHenry to Exit after generating a reduced order model.
' -R radius - If -p shells is specified, then this specifies the radius of the shells to use.
' -v - Regurgitate internal representation of the geometry to stdout in the input file Format
'

End Sub

Sub run_FH(GeneratePortunus As Boolean)
	Dim resultpath As String
	Dim fhobj As Object, couldRun As Long
	Set fhobj = CreateObject("FastHenry2.Document")

	resultpath = GetProjectPath("Model3D")
	resultpath = """"+resultpath + "fasthenry.inp"""
	resultpath = """"+FH_inp_name+FH_options+""""

  ' Run FastHenry
	couldRun = fhobj.Run(resultpath)
  Do While fhobj.IsRunning = True
    Wait 1
  Loop

  If Not couldRun Then
	MsgBox "Could not automatically start Fast Henry"
  End If

' Check if the FH run was successful
  Dim inductance() As Variant

  On Error GoTo NORESULT1

  inductance = fhobj.GetInductance()



If GeneratePortunus = True Then generate_Portunus_file(fhobj)


' Quit FastHenry2
fhobj.Quit
' Destroy FastHenry2 object
Set fhobj = Nothing
  Exit Sub

NORESULT1:
Dim tmpstr As String
  tmpstr ="No valid result obtained from the FastHenry run. "
  If GeneratePortunus = True Then tmpstr = tmpstr + "No Portunus file generated."
  tmpstr = tmpstr + vbCrLf + "Macro stops."
  MsgBox tmpstr, vbOkOnly+vbCritical, "Error"
  Exit All
End Sub


Sub write_FH()
	Dim filename As String, tmpstr As String, tmpsn As String
	Dim iii As Long, kkk As Long

	If nnodes = 0 Or nports = 0 Or nedges = 0 Then
		tmpstr = "The number of nodes, segments or ports cannot be zero." + vbCrLf + "Nodes: "+MyCStr(nnodes)+"  Segments: "+MyCStr(nedges) + "  Ports: "+MyCStr(nports) + vbCrLf
		tmpstr = tmpstr + vbCrLf + "Macro stops."
		MsgBox tmpstr, vbCritical, "Error"
		Exit All
	End If

'	filename = GetProjectPath("Model3D") + "\Fasthenry.inp"
	filename = FH_inp_name
	Open filename For Output As #99

	Print #99, "* FastHenry file generated from CST MWS"
	Print #99, "* CST MWS input file: " + GetProjectPath("Project") + vbCrLf

	' =========== Geometry units
	tmpstr = Units.GetUnit("Length")
	If (StrComp(tmpstr, "nm",vbBinaryCompare)=0 Or StrComp(tmpstr, "ft", vbBinaryCompare)=0) Then
		MsgBox ("Geometry unit " + tmpstr + " not implemented in the FH translator." + vbCrLf + _
				"Accepted units: m, cm, mm, um, in, mil. Please scale your model.")
		Exit All
	End If

	If (tmpstr = "mil") Then tmpstr = "mils"
	Print #99, ".units " + tmpstr + vbCrLf

	' ============ Nodes
	Print #99, "* Nodes"
	tmpsn = ""
	For iii = 1 To nnodes
		If curve_nodes(iii).nname <> tmpsn Then
			tmpsn = curve_nodes(iii).nname
			Print #99, vbCrLf + "* " + tmpsn
		End If
		Print #99, "N"+MyCStr(iii) + " " + _
			"x=" + MyCStr(curve_nodes(iii).x) + " " + _
			"y=" + MyCStr(curve_nodes(iii).y) + " " + _
			"z=" + MyCStr(curve_nodes(iii).z)
	Next iii
	Print #99, vbCrLf

	' ============ EQUIV's
	' Nodes that are double, for some reason (e.g. because two different curves share that node)
	' Nodes that are connected in parallel
	Print #99, "* Equivalenced Nodes"
	For iii = 1 To nequivs
		Print #99, ".equiv"+ " " + _
			"N" + MyCStr(equiv_nodes(iii,0)) + " " + _
			"N" + MyCStr(equiv_nodes(iii,1))
	Next iii
	Print #99, vbCrLf

	' ============ Segments
	Print #99, "* Segments connecting the nodes"

	' Fast Henry syntax:
	' Estr node1 node2 [w = value] [h = value] [sigma, rho = value]
	'	[wx = value wy = value wz = value]
	'	[nhinc = value] [nwinc = value] - discretization along h and w
	'	[rh = value] [rw = value] - ratio of adjacent filaments along h and w
	tmpsn = ""

	For iii = 1 To nedges
		' Print edge_names ??? First sort edges according to edge_names??
		kkk = sorted_edges(iii)
		If edge_names(kkk) <> tmpsn Then
			tmpsn = edge_names(kkk)
			Print #99, vbCrLf + "* Component " + tmpsn
		End If
		tmpstr = "E"+MyCStr(kkk) + " " + _
			"N" + MyCStr(edges(kkk,0)) + " " + _
			"N" + MyCStr(edges(kkk,1)) + " " + _
			" w=" + MyCStr(edge_widths(kkk)) + " h=" + MyCStr(edge_heights(kkk)) + " " + _
			" sigma=" + MyCStr(edge_sigma(kkk)) + " " + _
			" wx=" + MyCStr(edge_wx(kkk)) + " wy=" + MyCStr(edge_wy(kkk)) + " wz=" + MyCStr(edge_wz(kkk)) + " "
		If (edge_nhinc <> 1) Then 	tmpstr = tmpstr + " nhinc=" + MyCStr(edge_nhinc) + " "
		If (edge_nwinc <> 1) Then 	tmpstr = tmpstr + " nwinc=" + MyCStr(edge_nwinc) + " "
		If (edge_rh <> 1)    Then 	tmpstr = tmpstr + " rh=" + MyCStr(edge_rh) + " "
		If (edge_rw <> 1)    Then 	tmpstr = tmpstr + " rw=" + MyCStr(edge_rw) + " "
'			"nhinc=" + edge_nhinc(kkk) + " nwinc=" + edge_nwinc(kkk) + " " + _
'			"rh=" + edge_rh(kkk) + " rw=" + edge_rw(kkk)
		Print #99, tmpstr
	Next iii
	Print #99, vbCrLf


	' =========== Groundplanes
	Print #99, "* Groundplanes"
	For iii = 0 To ngrounds-1
		If groundplanes(iii).typ = "Nonuniform" Then
			print_nonunif_groundplane(iii)
		End If
	Next
	Print #99, vbCrLf

	' ============ Ports
	Print #99, "* Ports of the network"
	filename = GetProjectPath("Model3D") + "\Fasthenry_port_terminals.txt"
	Open filename For Output As #98

	For iii = 1 To nports
		Print #99, ".external " + "N"+MyCStr(ports(iii,0)) + " N"+MyCStr(ports(iii,1)) + " " + no_Portunus_forbidden_chars(portnames(iii))
		Print #98, MyCStr(curve_nodes(ports(iii,0)).x) + " " + MyCStr(curve_nodes(ports(iii,0)).y) + " " + MyCStr(curve_nodes(ports(iii,0)).z)
		Print #98, MyCStr(curve_nodes(ports(iii,1)).x) + " " + MyCStr(curve_nodes(ports(iii,1)).y) + " " + MyCStr(curve_nodes(ports(iii,1)).z)
	Next
	Print #99, vbCrLf
	Close #98

	' ============ Frequency range of interest
	Dim f_unit As Double

	Print #99, "* Frequency range of interest"
	f_unit = Units.GetFrequencyUnitToSI
	Print #99, ".freq fmin=" + MyCStr(Solver.GetFMin * f_unit) + " fmax=" + MyCStr(Solver.GetFMax * f_unit) ' + "ndec=" + MyCStr(ndecades)

	' .freq fmin=1e4 fmax=1e8 ndec=1
	Print #99, vbCrLf

	' Last line
	Print #99, ".end"

End Sub

Sub	print_nonunif_groundplane(g As Long)
	Dim iii As Long, tmpi As String, tmpstr As String

	Dim p1 As point, p2 As point

	' ======= Nonuniform groundplane
	Print #99, "G"+groundplanes(g).gname
	For iii = 0 To 2
		tmpi = MyCStr(iii+1) + "="
		Print #99, "+ x" + tmpi + MyCStr(groundplanes(g).P(iii).x) + " y" + tmpi+ MyCStr(groundplanes(g).P(iii).y) +" z"+tmpi+ MyCStr(groundplanes(g).P(iii).z) +" "
	Next iii
	Print #99, "+ thick=" + MyCStr(groundplanes(g).thick)
	Print #99, "+ file=NONE"
	Print #99, "+ sigma=" + MyCStr(groundplanes(g).sigma)

	' ====== Initial grid
	If StrComp(groundplanes(g).init_grid, "None") <> 0 Then
		Print #99, "*"
		Print #99, "* Define initial " + LCase(groundplanes(g).init_grid) + " grid"

		tmpstr = "+ contact initial"
		If StrComp(groundplanes(g).init_grid, "Meshed", vbBinaryCompare) = 0 Then tmpstr = tmpstr+"_meshed"
		tmpstr = tmpstr+"_grid (" + MyCStr(groundplanes(g).init_grid_segs(0)) + ", " + MyCStr(groundplanes(g).init_grid_segs(1)) + ")
		Print #99, tmpstr
	End If

	' ========= Refinement under / above ports
	If groundplanes(g).refine_ports Then
		Print #99, "*"
		Print #99, "* Refinement under / above external ports"
		determine_projected_ports(g)
		For iii = 0 To nprojports-1
			p1 = projports(iii)
			tmpstr ="+ contact connection " + "N_"+projportnames(iii) + "in ("
			tmpstr = tmpstr + MyCStr(p1.x) + ", " + MyCStr( p1.y) +", " +MyCStr(p1.z)  +", "
			tmpstr = tmpstr + MyCStr(projports_w(iii)) + ", " + MyCStr(projports_w(iii)) + ", "+ MyCStr(groundplanes(g).port_refinement_ratio) + ")"

			Print #99, tmpstr
		Next iii
	End If


	' ========= Refinement under / above traces
	If groundplanes(g).refine_traces Then
		Print #99, "*"
		Print #99, "* Refinement under / above signal lines"
		determine_projected_traces(g)
		For iii = 0 To nprojtraces-1
			p1 = projtraces(iii, 0)
			p2 = projtraces(iii, 1)
			tmpstr ="+ contact trace ("
			tmpstr = tmpstr + MyCStr(p1.x) + ", " + MyCStr( p1.y) +", " +MyCStr(p1.z)  +", "
			tmpstr = tmpstr + MyCStr(p2.x) + ", " + MyCStr( p2.y) +", " +MyCStr(p2.z)  +", "
			tmpstr = tmpstr + MyCStr(projtraces_w(iii)) + ", " + MyCStr(groundplanes(g).trace_refinement_scale_factor) + ")"

			Print #99, tmpstr
		Next iii
	End If


End Sub

Function MyCStr(x As Variant) As String
	MyCStr = cstr(x)
End Function

Sub determine_projected_traces(g As Long)
	Dim iii As Long, jjj As Long
	Dim xx(5) As Double, yy(5) As Double, zz(5) As Double
	Dim n1 As Long, n2 As Long
	Dim w As Double

	Dim true_proj_seg As Boolean

	nprojtraces = 0

	For iii = 1 To nedges
		n1 = edges(iii,0)
		n2 = edges(iii,1)
		w = edge_widths(iii)
		xx(1) = curve_nodes(n1).x: yy(1) = curve_nodes(n1).y: zz(1) = curve_nodes(n1).z
		xx(2) = curve_nodes(n2).x: yy(2) = curve_nodes(n2).y: zz(2) = curve_nodes(n2).z
		true_proj_seg = project_segment(xx, yy, zz, groundplanes(g).normal.x, groundplanes(g).normal.y, groundplanes(g).normal.z, _
						groundplanes(g).P(0).x, groundplanes(g).P(0).y, groundplanes(g).P(0).z)
		If true_proj_seg Then
		If tracenodes_inside_groundplane(g, xx, yy, zz) And trace_exists(nprojtraces, xx, yy, zz) = False Then
			For jjj = 0 To 1
				projtraces(nprojtraces, jjj).x = xx(jjj+3)
				projtraces(nprojtraces, jjj).y = yy(jjj+3)
				projtraces(nprojtraces, jjj).z = zz(jjj+3)
			Next jjj
			projtraces_w(nprojtraces) = w
			nprojtraces = nprojtraces + 1
		End If
		End If	' True projected segment

	Next iii

End Sub

Function tracenodes_inside_groundplane(g As Long, xx() As Double, yy() As Double, zz() As Double) As Boolean
	Dim xgmin As Double, xgmax As Double, ygmin As Double, ygmax As Double, zgmin As Double, zgmax As Double
	Dim iii As Integer

	xgmin = groundplanes(g).P(0).x: xgmax = xgmin
	ygmin = groundplanes(g).P(0).y: xgmax = xgmin
	zgmin = groundplanes(g).P(0).z: xgmax = xgmin
	For iii = 0 To 2
		If groundplanes(g).P(iii).x < xgmin Then xgmin = groundplanes(g).P(iii).x:
		If groundplanes(g).P(iii).x > xgmax Then xgmax = groundplanes(g).P(iii).x:
		If groundplanes(g).P(iii).y < ygmin Then ygmin = groundplanes(g).P(iii).y:
		If groundplanes(g).P(iii).y > ygmax Then ygmax = groundplanes(g).P(iii).y:
		If groundplanes(g).P(iii).z < zgmin Then zgmin = groundplanes(g).P(iii).z:
		If groundplanes(g).P(iii).z > zgmax Then zgmax = groundplanes(g).P(iii).z:
	Next iii

	Dim eps As Double
	eps = 1e-10

	If _
		xx(3) >= xgmin And xx(3) <= xgmax And xx(4) >= xgmin And xx(4) <= xgmax And _
		yy(3) >= ygmin And yy(3) <= ygmax And yy(4) >= ygmin And yy(4) <= ygmax And _
		zz(3) >= zgmin And zz(3) <= zgmax And zz(4) >= zgmin And zz(4) <= zgmax         Then
		tracenodes_inside_groundplane = True
'		Abs(xx(3)-xgmin)<eps And Abs(xx(3)-xgmax)<eps And Abs(xx(4) -xgmin)<eps And Abs(xx(4)-xgmax)<eps And _
'		Abs(yy(3)-ygmin)<eps And Abs(yy(3)-ygmax)<eps And Abs(yy(4) -ygmin)<eps And Abs(yy(4)-ygmax)<eps And _
'		Abs(zz(3)-zgmin)<eps And Abs(zz(3)-zgmax)<eps And Abs(zz(4) -zgmin)<eps And Abs(zz(4)-zgmax)<eps 		Then
	Else
		tracenodes_inside_groundplane = False
	End If

End Function

Function portnodes_inside_groundplane(g As Long, xx() As Double, yy() As Double, zz() As Double) As Boolean
	Dim xgmin As Double, xgmax As Double, ygmin As Double, ygmax As Double, zgmin As Double, zgmax As Double
	Dim iii As Integer

	xgmin = groundplanes(g).P(0).x: xgmax = xgmin
	ygmin = groundplanes(g).P(0).y: xgmax = xgmin
	zgmin = groundplanes(g).P(0).z: xgmax = xgmin
	For iii = 0 To 2
		If groundplanes(g).P(iii).x < xgmin Then xgmin = groundplanes(g).P(iii).x:
		If groundplanes(g).P(iii).x > xgmax Then xgmax = groundplanes(g).P(iii).x:
		If groundplanes(g).P(iii).y < ygmin Then ygmin = groundplanes(g).P(iii).y:
		If groundplanes(g).P(iii).y > ygmax Then ygmax = groundplanes(g).P(iii).y:
		If groundplanes(g).P(iii).z < zgmin Then zgmin = groundplanes(g).P(iii).z:
		If groundplanes(g).P(iii).z > zgmax Then zgmax = groundplanes(g).P(iii).z:
	Next iii

	Dim eps As Double
	eps = 1e-10

	If _
		xx(3) >= xgmin And xx(3) <= xgmax  And _
		yy(3) >= ygmin And yy(3) <= ygmax  And _
		zz(3) >= zgmin And zz(3) <= zgmax          Then
		portnodes_inside_groundplane = True
'		Abs(xx(3)-xgmin)<eps And Abs(xx(3)-xgmax)<eps And Abs(xx(4) -xgmin)<eps And Abs(xx(4)-xgmax)<eps And _
'		Abs(yy(3)-ygmin)<eps And Abs(yy(3)-ygmax)<eps And Abs(yy(4) -ygmin)<eps And Abs(yy(4)-ygmax)<eps And _
'		Abs(zz(3)-zgmin)<eps And Abs(zz(3)-zgmax)<eps And Abs(zz(4) -zgmin)<eps And Abs(zz(4)-zgmax)<eps 		Then
	Else
		portnodes_inside_groundplane = False
	End If

End Function

Function trace_exists(nprojtraces As Long, xx() As Double, yy() As Double, zz() As Double) As Boolean
	Dim iii As Long, jjj As Long
	Dim pi As point, pf As point
	Dim eq1 As Boolean, eq2 As Boolean

	If nprojtraces = 0 Then Exit Function

	For iii = 0 To nprojtraces-1
		pi = projtraces(iii, 0)
		pf = projtraces(iii, 1)
		If  Abs(xx(3)-pi.x) < USER_EPS And Abs(yy(3)-pi.y) < USER_EPS And Abs(zz(3)-pi.z) < USER_EPS And _
			Abs(xx(4)-pf.x) < USER_EPS And Abs(yy(4)-pf.y) < USER_EPS And Abs(zz(4)-pf.z) < USER_EPS Then eq1 = True
		If  Abs(xx(3)-pf.x) < USER_EPS And Abs(yy(3)-pf.y) < USER_EPS And Abs(zz(3)-pf.z) < USER_EPS And _
			Abs(xx(4)-pi.x) < USER_EPS And Abs(yy(4)-pi.y) < USER_EPS And Abs(zz(4)-pi.z) < USER_EPS Then eq2 = True
		If (eq1 Or eq2) Then
			trace_exists = True
			Exit Function
		End If
	Next iii
	trace_exists = False
End Function

Sub determine_projected_ports(g As Long)
	Dim iii As Long, jjj As Long, kkk As Long
	Dim xx(5) As Double, yy(5) As Double, zz(5) As Double
	Dim n1 As Long, n2 As Long
	Dim w As Double

	nprojports = 0

	For iii = 1 To nports
		n1 = ports(iii,0)
		n2 = ports(iii,1)
		w = 0	'	ATTENTION! TO DETERMINE FROM SOMETHING ELSE ... edge_widths(iii)
		xx(1) = curve_nodes(n1).x: yy(1) = curve_nodes(n1).y: zz(1) = curve_nodes(n1).z
		xx(2) = curve_nodes(n2).x: yy(2) = curve_nodes(n2).y: zz(2) = curve_nodes(n2).z
		For kkk = 1 To 2
			project_point(xx, yy, zz, kkk, groundplanes(g).normal.x, groundplanes(g).normal.y, groundplanes(g).normal.z, _
						groundplanes(g).P(0).x, groundplanes(g).P(0).y, groundplanes(g).P(0).z)
			If portnodes_inside_groundplane(g, xx, yy, zz) And projected_port_exists(nprojports, xx, yy, zz) = False Then
				projports(nprojports).x = xx(3)
				projports(nprojports).y = yy(3)
				projports(nprojports).z = zz(3)
				projports_w(nprojports) = w
				projportnames(nprojports) = portnames(iii)
				nprojports = nprojports + 1
			End If
		Next kkk
	Next iii

	ReDim portlist(nprojports+1)
	ReDim portwidth_dlgtext(nprojports+1)
	ReDim portwidths_txt(nprojports+1)

	For iii = 1 To nprojports
		portlist(iii) = projportnames(iii-1)
		portwidths_txt(iii) = "1"
	Next

	nptot = nprojports


	For iii=1 To nprojports
		portwidth_dlgtext(iii)= portlist(iii) + Space(30-Len(portlist(iii))) + vbTab + "1"
	Next iii

	Begin Dialog UserDialog 400,280,"Define width for the ports",.dialogfunc ' %GRID:10,7,1,1
		ListBox 30,28,120,56,portlist(),.Crtport
		Text 40,7,80,14,"Port mode",.Text1
		Text 180,7,100,14,"Width",.Text2
		TextBox 170,28,90,21,.Crtwidth

		PushButton 280,56,90,21,"Set All",.SetAll

		Text 30,112,300,14,"Port mode" +Space(30-9)+vbTab+"     Width",.Text4
		ListBox 30,133,300,105,portwidth_dlgtext(),.Finalwidths

		OKButton 30,245,100,21
		CancelButton 140,245,110,21
		Text 270,28,90,14,Units.GetUnit("Length"),.Text3

	End Dialog
	Dim dlg As UserDialog

	crtportDlg = 1
	dlg.Crtwidth = portwidths_txt(1)

	If Dialog(dlg) = 0 Then
	   Exit All
    End If

	For iii = 0 To nprojports-1
		projports_w(iii) = Val(portwidths_txt(iii+1))
	Next

End Sub

'-------------------------------------------------------------------------------------
Function DialogFunc_userdef(DlgItem$, Action%, SuppValue%) As Boolean

	Dim Extension As String, projectdir As String, filename As String
    Select Case Action%
    Case 1 ' Dialog box initialization
'    	DlgEnable "Browseinputfile2", False
'    	DlgEnable "Browseinputfile3", False
'    	DlgEnable "Text4", False
'    	DlgEnable "FHOutFile", False
'    	DlgEnable "PortunusOutFile", False
    	DlgEnable "RememberSettings", False
    Case 2 ' Value changing or button pressed
		projectdir = GetProjectPath("Model3D")	'Dirname(GetProjectbasename)
        Select Case DlgItem
        	Case "GeneratePortunus"
        		If DlgValue("GeneratePortunus") = 0 Then
			    	DlgEnable "PortunusOutFile", False
			    Else
			    	DlgEnable "PortunusOutFile", True
			    	DlgValue("RunFH"), 1
			    	DlgEnable "FHOutFile", True
			    End If
        	DialogFunc_userdef = True
        	Case "RunFH"
        		If DlgValue("RunFH") = 0 Then
			    	DlgEnable "FHOutFile", False
			    Else
			    	DlgEnable "FHOutFile", True
			    End If
        	DialogFunc_userdef = True

        	Case "Browseinputfile"
        		Extension = "inp"
                filename = GetFilePath(,Extension, projectdir, "Specify FH output file", 0)
                If (filename <> "") Then
                    DlgText "FHInfile", FullPath(filename, projectdir)
                    DlgText "FHInfile", ShortName(filename)
                End If
        	DialogFunc_userdef = True
        	Case "Browseinputfile2"
        		Extension = "mat"
                filename = GetFilePath(,Extension, projectdir, "Specify FH result file", 0)
                If (filename <> "") Then
                    DlgText "FHOutFile", FullPath(filename, projectdir)
                    DlgText "FHOutFile", ShortName(filename)
                End If
        	DialogFunc_userdef = True
        	Case "Browseinputfile3"
        		Extension = "c2p"
                filename = GetFilePath(,Extension, projectdir, "Specify Portunus result file", 0)
                If (filename <> "") Then
                    DlgText "PortunusOutFile", FullPath(filename, projectdir)
                    DlgText "PortunusOutFile", ShortName(filename)
                End If
        	DialogFunc_userdef = True
        	Case "Browseinputfile4"
        		Dim myprojectdir As Variant, myfilename As String
        		myprojectdir = projectdir
                myfilename = GetFolder_Lib(projectdir, False, False) ', "Specify result directory")	'GetFilePath(,Extension, projectdir, "Specify Portunus result file", 0)
                If (myfilename <> "") Then
                    DlgText "ResultDir", FullPath(myfilename, myprojectdir)
                End If
        	DialogFunc_userdef = True
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Private Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------


Dim crtp As Long

    Select Case Action%
    Case 1 ' Dialog box initialization
        DlgEnable "Finalwidths", False

    Case 2 ' Value changing or button pressed
    	If DlgItem$ = "Crtport" Then
    		crtportDlg = DlgValue("Crtport")+1
    		DlgText("Crtwidth", portwidths_txt(crtportDlg))
    	End If

    	If DlgItem$ = "Help" Then
			DialogFunc = True 'do not exit the dialog
    	End If

    	If DlgItem$ = "SetAll" Then
	Begin Dialog UserDialog 230,126,"Set all ports" ' %GRID:10,7,1,1
		Text 20,7,180,14,"Ports:    All ports",.Text1
		Text 20,35,90,14,"Width",.Text2
		OKButton 10,91,90,21
		CancelButton 130,91,90,21
		TextBox 120,28,90,21,.Amplall
	End Dialog
			Dim dlg1 As UserDialog
			dlg1.Amplall = "1"

			If ( Dialog(dlg1) = -1)  Then ' OK button was pressed
				Dim iii As Long
				For iii = 1 To nptot
					portwidths_txt(iii) = dlg1.Amplall
					portwidth_dlgtext(iii)= portlist(iii) + Space(30-Len(portlist(iii)))+ vbTab+portwidths_txt(iii) ' + vbTab + vbTab  + phshiftstxt(iii)
				Next iii
				DlgListBoxArray "Finalwidths", portwidth_dlgtext
				DlgText("crtwidth",  portwidths_txt(crtportDlg))
			End If
			DialogFunc = True 'do not exit the dialog
    	End If


    Case 3	' Textbox value changing
        If DlgItem$ = "Crtwidth" Then
        	crtp = crtportDlg  'DlgValue("Crtport")+1
			portwidths_txt(crtp) = DlgText("crtwidth")
			portwidth_dlgtext(crtp)= portlist(crtp) + Space(30-Len(portlist(crtp))) + vbTab+portwidths_txt(crtp) ' + vbTab + vbTab + phshiftstxt(crtp)
			DlgListBoxArray "Finalwidths", portwidth_dlgtext


            DialogFunc = True 'do not exit the dialog
        End If

    Case 4 ' Focus changed
    Case 6 ' Function key

    End Select

End Function


Function projected_port_exists(nprojports As Long, xx() As Double, yy() As Double, zz() As Double) As Boolean
	Dim iii As Long, jjj As Long
	Dim pi As point, pf As point
	Dim eq1 As Boolean, eq2 As Boolean

	If nprojports = 0 Then Exit Function

	For iii = 0 To nprojports-1
		pi = projports(iii)
		If  Abs(xx(3)-pi.x) < USER_EPS And Abs(yy(3)-pi.y) < USER_EPS And Abs(zz(3)-pi.z) < USER_EPS Then eq1 = True
		If (eq1 ) Then
			projected_port_exists = True
			Exit Function
		End If
	Next iii
	projected_port_exists = False
End Function

Sub determine_ports()
	Dim ncurve As Long

	Dim nc1 As Long, nc2 As Long
	Dim curvenodes(MAX_NODES) As Long
	Dim tmpstr As String, iii As Long

	For ncurve = 0 To ncurves-1
		tmpstr = FH_curves(ncurve).cname
		If (StrComp(LCase(Left(tmpstr, 6)), "ground", vbBinaryCompare) <>0) Then
		If InStr(tmpstr, "-ports") = 0 And InStr(LCase(tmpstr), "equivs") = 0 Then	' Name contains neither "no ports" nor "equivs": we should enter a port
		If FH_curves(ncurve).nsegs = 1 And FH_curves(ncurve).point_ini(1) = FH_curves(ncurve).point_fin(1) Then
			insert_one_port(ncurve)
		Else
			ports(nports+1,0) = FH_curves(ncurve).point_ini(1): curve_nodes(ports(nports+1,0)).taken = curve_nodes(ports(nports+1,0)).taken+1
			ports(nports+1,1) = FH_curves(ncurve).point_fin(FH_curves(ncurve).nsegs): curve_nodes(ports(nports+1,1)).taken= curve_nodes(ports(nports+1,1)).taken+1
			curve_nodes(ports(nports+1,0)).hasport = True
			curve_nodes(ports(nports+1,0)).hasport = True

			Select Case FH_curves(ncurve).typ
				Case "Series"
					tmpstr = Left( tmpstr,  InStr (tmpstr, ":")-1)
				Case "Parallel"
					tmpstr = Left( tmpstr,  InStr (tmpstr, ":")-1)
					tmpstr = Left( tmpstr,  InStr (tmpstr, "+")-1)
			End Select
			portnames(nports+1) = tmpstr 	' ATTENTION! Put something else here?
			nports = nports+1
		End If
		End If
		End If
	Next ncurve

End Sub

Sub insert_one_port(ncurve As Long)
	Dim iii As Long, jjj As Long
	Dim iiitmp1 As Long, iiitmp2 As Long, jjjtmp1 As Long, jjjtmp2 As Long, n0 As Long
	Dim no_ini As Long, no_fin As Long
	Dim tmpstr As String

	Dim n_ini As Long, n_fin As Long

	n0 = FH_curves(ncurve).point_ini(1)	' Curve has only one segment, both nodes identical
	no_ini = 0
	no_fin = 0

	' See if the port is in the middle of one curve
	For iii = 0 To ncurves - 1
	If iii <> ncurve Then
		For jjj = 1 To FH_curves(iii).nsegs
			If nodes_identical (FH_curves(iii).point_ini(jjj), n0, curve_nodes, USER_EPS) Then
				n_ini = FH_curves(iii).point_ini(jjj)
				no_ini = no_ini+1
				iiitmp1 = iii: jjjtmp1 = jjj
			End If
			If nodes_identical (FH_curves(iii).point_fin(jjj), n0, curve_nodes, USER_EPS) Then
				n_fin = FH_curves(iii).point_ini(jjj)
				no_fin = no_fin+1
				iiitmp2 = iii: jjjtmp2 = jjj
			End If
		Next jjj
	End If
	Next iii

	If no_ini > 1 Or no_fin > 1 Then
		MsgBox ("Found more than one possible positions for internal port " + FH_curves(iii).cname + ". Macro stops")
		Exit All
	End If
	If iiitmp1 <> iiitmp2 Then
		MsgBox "Internal port " + FH_curves(iii).cname + "does not seem to belong to the same curve. Macro stops"
		Exit All
	End If

	FH_curves(iiitmp1).point_ini(jjjtmp1) = n0
			ports(nports+1,0) = FH_curves(iiitmp1).point_fin(jjjtmp2): curve_nodes(ports(nports+1,0)).taken = curve_nodes(ports(nports+1,0)).taken+1
			ports(nports+1,1) = n0: curve_nodes(ports(nports+1,1)).taken= curve_nodes(ports(nports+1,1)).taken+1	'FH_curves(ncurve).point_ini(1)
			curve_nodes(ports(nports+1,0)).hasport = True
			curve_nodes(ports(nports+1,1)).hasport = True
			tmpstr = FH_curves(ncurve).cname

			Select Case FH_curves(ncurve).typ
				Case "Series"
					tmpstr = Left( tmpstr,  InStr (tmpstr, ":")-1)
				Case "Parallel"
					tmpstr = Left( tmpstr,  InStr (tmpstr, ":")-1)
					tmpstr = Left( tmpstr,  InStr (tmpstr, "+")-1)
			End Select
			portnames(nports+1) = tmpstr
			nports = nports+1


End Sub



Sub generate_FH_data_structures()
	Dim iii As Long, jjj As Long
	Dim w As Double, h As Double, wx As Double, wy As Double, wz As Double

	Dim tmpnedges As Long, brickno As Long, facepairno As Integer

	Dim found As Boolean

	Dim sName As String, compName As String
	Dim tmpstr As String, tmpbool As Boolean


	nedges = 0

	' ============ For the time being, deal only with "series" curves

	For iii = 0 To ncurves-1
		For jjj = 1 To FH_curves(iii).nsegs
		If FH_curves(iii).leng(jjj) <> 0 Then
			nedges = nedges+1
			edge_types(nedges) = "Series"
			found = find_brick(iii, jjj, w, h, wx, wy, wz, brickno, facepairno)
			If (found = False) Then
				If (StrComp(LCase(Left(FH_curves(iii).cname, 6)), "ground", vbBinaryCompare) <>0) Then
					If (StrComp(LCase(Left(FH_curves(iii).cname, 6)), "equivs", vbBinaryCompare) = 0) Then
						SelectTreeItem "Curves\"+ Replace(FH_curves(iii).csname(jjj), ":", "\")
						tmpbool = equiv_segment(iii, jjj)
						nedges = nedges-1
						GoTo nextjjj
					Else
						SelectTreeItem "Curves\"+ Replace(FH_curves(iii).csname(jjj), ":", "\")
						tmpbool = equiv_segment_ask(iii, jjj)
						nedges = nedges-1
						GoTo nextjjj
					End If
				Else	' Found a curve whose name starts with "ground"; used for possible future developments
					nedges = nedges-1
					Exit For
				End If
			End If
			edges(nedges, 0) = FH_curves(iii).point_ini(jjj)
			edges(nedges, 1) = FH_curves(iii).point_fin(jjj)

			edge_widths(nedges) = w
			edge_heights(nedges) = h
			edge_wx(nedges) = wx: edge_wy(nedges) = wy: edge_wz(nedges) = wz
			edge_sigma(nedges) = FH_brick_properties(brickno).sigma

			tmpstr = FH_brick_properties(brickno).bname: tmpstr = Left(tmpstr, InStr(tmpstr, ":")-1)
			edge_names(nedges) = tmpstr
					curve_nodes(edges(nedges, 0)).nname = tmpstr: curve_nodes(edges(nedges, 1)).nname = tmpstr
			'edge_nhinc(nedges) =
			'edge_nwinc(nedges) =
			'edge_rh(nedges) =
			'edge_rw(nedges) =
			End If
nextjjj:
		Next jjj

		If FH_curves(iii).typ = "Parallel" Then
			edge_types(nedges) = "Parallel"

			tmpnedges = nedges

			If FH_curves(iii).nsegs <> 1 Then
				MsgBox "Wrong number of segments for the parallel curve connection "+FH_curves(iii).cName + vbCrLf _
				+ "Expected: 1. Found: " + CStr(FH_curves(iii).nsegs), vbCritical, "Error"
				Exit All
			End If
			sName = FH_brick_properties(brickno).bname
			compName = Left(sName, InStr(sName, ":")-1)
			For jjj = 0 To nbricks-1
				tmpstr = FH_brick_properties(jjj).bname
				tmpstr = Left(tmpstr, InStr(tmpstr, ":")-1)
				If StrComp(tmpstr,compName,vbBinaryCompare)=0 And jjj <> brickno Then		' Found a brick that is connected in parallel
					' ATTENTION! It is assumed that the parallel edges correspond always to the
					' same facepair!

		'	found = find_brick(iii, jjj, w, h, wx, wy, wz, brickno, facepairno)
					facepairno = get_facepairno_parallel(iii, 1, jjj)

					' Add two nodes
					nnodes = nnodes+1
					curve_nodes(nnodes).x = FH_bricks(jjj, facepairno).P1.x
					curve_nodes(nnodes).y = FH_bricks(jjj, facepairno).P1.y
					curve_nodes(nnodes).z = FH_bricks(jjj, facepairno).P1.z
					curve_nodes(nnodes).nname = tmpstr

					nnodes = nnodes+1
					curve_nodes(nnodes).x = FH_bricks(jjj, facepairno).P2.x
					curve_nodes(nnodes).y = FH_bricks(jjj, facepairno).P2.y
					curve_nodes(nnodes).z = FH_bricks(jjj, facepairno).P2.z
					curve_nodes(nnodes).nname = tmpstr

					' Add an edge

					' First, verify which node should come first
					' Nodes of the first edge: edges(tmpnedges, 0), edges(tmpnedges, 1)
					' Nodes of the current edge: nnodes-1, nnodes

					Dim dist1 As Double, dist2 As Double
					dist1 = Sqr( _
						(curve_nodes(nnodes-1).x-curve_nodes(edges(tmpnedges, 0)).x)^2+ _
						(curve_nodes(nnodes-1).y-curve_nodes(edges(tmpnedges, 0)).y)^2+ _
						(curve_nodes(nnodes-1).z-curve_nodes(edges(tmpnedges, 0)).z)^2)
					dist2 = Sqr( _
						(curve_nodes(nnodes).x-curve_nodes(edges(tmpnedges, 0)).x)^2+ _
						(curve_nodes(nnodes).y-curve_nodes(edges(tmpnedges, 0)).y)^2+ _
						(curve_nodes(nnodes).z-curve_nodes(edges(tmpnedges, 0)).z)^2)

					nedges = nedges+1
					If dist1 < dist2 Then
						edges(nedges, 0) = nnodes-1: curve_nodes(nnodes-1).taken = 1
						edges(nedges, 1) = nnodes: curve_nodes(nnodes).taken = 1
					Else
						edges(nedges, 1) = nnodes-1: curve_nodes(nnodes-1).taken = 1
						edges(nedges, 0) = nnodes: curve_nodes(nnodes).taken = 1
					End If

					edge_widths(nedges) = FH_bricks(jjj, facepairno).w
					edge_heights(nedges) = FH_bricks(jjj, facepairno).h
					edge_wx(nedges) = FH_bricks(jjj, facepairno).wx
					edge_wy(nedges) = FH_bricks(jjj, facepairno).wy
					edge_wz(nedges) = FH_bricks(jjj, facepairno).wz
					edge_sigma(nedges) = FH_brick_properties(jjj).sigma
					edge_names(nedges) = tmpstr
					'edge_nhinc(nedges) =
					'edge_nwinc(nedges) =
					'edge_rh(nedges) =
					'edge_rw(nedges) =
					edge_types(nedges) = "Parallel"


					' Add equivalent nodes; may not be necessary, according to Mr. Linde, if bricks are electrically connected
					If ADD_EQUIVNODES = True Then
						nequivs = nequivs +1
						equiv_nodes(nequivs,0) = edges(tmpnedges, 0)
						equiv_nodes(nequivs,1) = edges(nedges, 0)

						nequivs = nequivs +1
						equiv_nodes(nequivs,0) = edges(tmpnedges, 1)
						equiv_nodes(nequivs,1) = edges(nedges, 1)
					End If


					If CHECK_PARALL_CURVES Then		' Construct the curve and the equiv's
With Polygon3D
     .Reset
     .Name "3dpolygon"+Cstr(jjj)
     .Curve compName+ "_parallelTest"
     .Point curve_nodes(nnodes-1).x, curve_nodes(nnodes-1).y, curve_nodes(nnodes-1).z
     .Point   curve_nodes(nnodes).x,   curve_nodes(nnodes).y,   curve_nodes(nnodes).z
     .Create
End With

With Polygon3D
     .Reset
     .Name "3dpolygon"+Cstr(jjj)
     .Curve compName+ "_cross1Test"
     .Point curve_nodes(nnodes-1).x, curve_nodes(nnodes-1).y, curve_nodes(nnodes-1).z
     .Point curve_nodes(edges(tmpnedges, 0)).x, curve_nodes(edges(tmpnedges, 0)).y, curve_nodes(edges(tmpnedges, 0)).z
     .Create
End With

With Polygon3D
     .Reset
     .Name "3dpolygon"+Cstr(jjj)
     .Curve compName+ "_cross2Test"
     .Point   curve_nodes(nnodes).x,   curve_nodes(nnodes).y,   curve_nodes(nnodes).z
     .Point curve_nodes(edges(tmpnedges, 1)).x, curve_nodes(edges(tmpnedges, 1)).y, curve_nodes(edges(tmpnedges, 1)).z
     .Create
End With

					End If
				End If
			Next

		End If
	Next iii

If SORT_EDGES Then
	' Sort edges according to their name
	Dim enames(MAX_NODES) As String, nenames As Long, etyp(MAX_NODES) As Long, tmpi As Long

	nenames = 1
	tmpstr = edge_names(1): enames(1) = tmpstr: etyp(1) = 1: sorted_edges(1) = 1
	For iii = 2 To nedges
		sorted_edges(iii) = iii
		For jjj = 1 To nenames
			If StrComp (edge_names(iii), tmpstr) = 0 Then
				etyp(iii) = jjj
			Else
				nenames = nenames+1
				tmpstr =edge_names(iii): enames(nenames) = tmpstr
				etyp (iii) = nenames
			End If
		Next jjj
	Next iii

	For iii = 1 To nedges-1
		For jjj = iii+1 To nedges
			If etyp(sorted_edges(jjj)) < etyp(sorted_edges(iii)) Then
				tmpi = sorted_edges(jjj)
				sorted_edges(jjj) = sorted_edges(iii)
				sorted_edges(iii) = tmpi
			End If
		Next
	Next
Else
	For iii = 1 To nedges
		sorted_edges(iii) = iii
	Next
End If	' SORT_EDGES

'	MsgBox "Segments" + display_segments_coords(nedges, edges, curve_nodes, 1)

	determine_nodes_taken()

'	display_nodes_taken()

	For iii = 1 To nnodes
			For jjj = iii+1 To nnodes
				If jjj <> iii Then
					If nodes_identical (iii, jjj, curve_nodes, USER_EPS) Then
						' equivalence nodes
						If is_a_port(iii, jjj) = False Then

						curve_nodes(iii).taken = curve_nodes(iii).taken + 1
						curve_nodes(jjj).taken = curve_nodes(jjj).taken + 1
						nequivs = nequivs +1
						equiv_nodes(nequivs,0) = iii
						equiv_nodes(nequivs,1) = jjj
						jjj = nnodes
						End If
					End If
				End If
			Next
	Next

'		display_nodes_taken()

	Pick.ClearAllPicks

	Dim nodes_unconnected As Boolean

	nodes_unconnected = False

	For iii = 1 To nnodes
		If curve_nodes(iii).taken < 2 Then
			nodes_unconnected = True
			Pick.PickPointFromCoordinates curve_nodes(iii).x,curve_nodes(iii).y,curve_nodes(iii).z
		End If
	Next

	If nodes_unconnected Then
		MsgBox "Nodes exist that are not connected to any other segment, equiv or port." + vbCrLf + "They are marked as picked points."
'		Exit All
	End If

End Sub

Function get_facepairno_parallel(ncurve As Long, nseg As Long, brickno As Integer) As Integer
	Dim jjj As Long, iii As Long
	Dim found As Boolean
	iii = brickno

			For jjj = 1 To 3	' Face pair
				If Abs(FH_bricks(iii, jjj).dist - FH_curves(ncurve).leng(nseg)) < USER_EPS Then	' Found possible candidate
	'				found = check_midpoints(iii, jjj, ncurve, nseg)
					found = check_ParallelCurves(iii, jjj, ncurve, nseg)
					If found = True Then
						get_facepairno_parallel = jjj
						Exit Function
					End If
				End If
			Next jjj

			MsgBox ("Some problem exists with the parallel connection of bricks." + vbCrLf + "Please check your structure.")

End Function

Function is_a_port(n1 As Long, n2 As Long) As Boolean
	Dim iii As Long

	For iii = 1 To nports
		If ports(iii,0) = n1 And ports(iii,1) = n2    Or   (ports(iii,0) = n2 And ports(iii,1) = n1) Then
			is_a_port = True
			Exit Function
		End If
	Next
	is_a_port = False
End Function


Sub	determine_nodes_taken()
	Dim iii As Long, jjj As Long
	Dim n1 As Long, n2 As Long

	For iii = 1 To nnodes
		curve_nodes(iii).taken = 0
	Next

	For iii = 1 To nedges
		n1 =edges(iii,0):n2 =edges(iii,1)
		curve_nodes(n1).taken = curve_nodes(n1).taken + 1
		curve_nodes(n2).taken = curve_nodes(n2).taken + 1
		If edge_types(iii) = "Parallel" Then
			curve_nodes(n1).taken = curve_nodes(n1).taken + 1
			curve_nodes(n2).taken = curve_nodes(n2).taken + 1
		End If
	Next

	For iii = 1 To nequivs
		n1 =equiv_nodes(iii,0):n2 =equiv_nodes(iii,1)
		curve_nodes(n1).taken = curve_nodes(n1).taken + 1
		curve_nodes(n2).taken = curve_nodes(n2).taken + 1
	Next

	For iii = 1 To nports
		n1 =ports(iii,0):n2 =ports(iii,1)
		curve_nodes(n1).taken = curve_nodes(n1).taken + 1
		curve_nodes(n2).taken = curve_nodes(n2).taken + 1
	Next

End Sub

Sub display_nodes_taken()
	Dim iii As Long, tmpstr As String

	tmpstr = "Nodes taken" + vbCrLf
	For iii = 1 To nnodes
		tmpstr = tmpstr + Cstr(iii) + " taken " + cstr(curve_nodes(iii).taken) + "      "
		If iii Mod 2 = 0 Then tmpstr = tmpstr + vbCrLf
		If curve_nodes(iii).taken < 2 Then Pick.PickPointFromCoordinates curve_nodes(iii).x,curve_nodes(iii).y,curve_nodes(iii).z

	Next

	MsgBox tmpstr

End Sub


Function equiv_segment(nc As Long, nseg As Integer) As Boolean

	Dim n1 As Long, n2 As Long

'	do_equivalence = True
'If do_equivalence = True Then
	nequivs = nequivs +1
	n1 =FH_curves(nc).point_ini(nseg)
	n2 = FH_curves(nc).point_fin(nseg)
	equiv_nodes(nequivs,0) = n1: curve_nodes(n1).taken = curve_nodes(n1).taken + 1
	equiv_nodes(nequivs,1) = n2: curve_nodes(n2).taken = curve_nodes(n2).taken + 1
	equiv_segment = True
		curve_nodes(n1).nname = "Equivalenced on curve  " +FH_curves(nc).csname(nseg)
		curve_nodes(n2).nname = "Equivalenced on curve  " +FH_curves(nc).csname(nseg)

	equiv_segment = False
'Else
'End If

End Function

Function equiv_segment_ask(nc As Long, nseg As Integer) As Boolean

	Dim n1 As Long, n2 As Long

If never_ask_equiv = False Then
	Begin Dialog UserDialog 480,182,"Curve segment without corresponding brick" ' %GRID:10,7,1,1
		Text 30,21,350,14,"Found a curve segment without corresponding brick.",.Text1
		Text 30,42,350,14,"Should its nodes be equivalenced?",.Text2
		CheckBox 110,105,250,14,"Do not ask again for this project",.dontask
		CheckBox 110,126,250,14,"Do not ask again for all projects",.dontask2
		OKButton 40,154,90,21
		CancelButton 160,154,90,21
		OptionGroup .equiv
			OptionButton 60,63,160,14,"Do not equivalence",.OptionButton1
			OptionButton 60,84,160,14,"Equivalence",.OptionButton2
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then	   Exit All

	If dlg.equiv = 0 Then
		do_equivalence = False
	Else
		do_equivalence = True
	End If

	If dlg.dontask = 1 Then never_ask_equiv = True
	If dlg.dontask2 = 1 Then
		never_ask_equiv = True
		Dim cfg_file As String

		cfg_file = Dir$(GetMacroPath + "\Construct\FastHenry\CST2FH.cfg")
		If cfg_file = "" Then
			cfg_file = GetMacroPath + "\Construct\FastHenry\CST2FH.cfg"
			Open cfg_file For Output As #99
		Else
			cfg_file = GetMacroPath + "\Construct\FastHenry\CST2FH.cfg"
			Open cfg_file For Append As #99
		End If
	Print #99, "NEVER_ASK_EQUIV True"
	If do_equivalence = True Then Print #99, "DO_EQUIVALENCE True"
	Close #99
	End If
End If

If do_equivalence = True Then
	nequivs = nequivs +1
	n1 =FH_curves(nc).point_ini(nseg)
	n2 = FH_curves(nc).point_fin(nseg)
	equiv_nodes(nequivs,0) = n1: curve_nodes(n1).taken = curve_nodes(n1).taken + 1
	equiv_nodes(nequivs,1) = n2: curve_nodes(n2).taken = curve_nodes(n2).taken + 1
	equiv_segment_ask = True
		curve_nodes(n1).nname = "Equivalenced on curve  " +FH_curves(nc).csname(nseg)
		curve_nodes(n2).nname = "Equivalenced on curve  " +FH_curves(nc).csname(nseg)

	equiv_segment_ask = False
Else
End If

End Function

Function find_brick(ncurve As Long, nseg As Long, w As Double, h As Double, wx As Double, wy As Double, wz As Double, brickno As Long, facepairno As Integer) As Boolean
	Dim iii As Long, jjj As Long
	Dim found As Boolean


	For iii = 0 To nbricks-1
		If bricks_found(iii) = 0 Then	' This brick has not yet been considered
			For jjj = 1 To 3	' Face pair
				If Abs(FH_bricks(iii, jjj).dist - FH_curves(ncurve).leng(nseg)) < USER_EPS Then	' Found possible candidate
					found = check_midpoints(iii, jjj, ncurve, nseg)
					If found = True Then
						bricks_found(iii) = 1
						w = FH_bricks(iii, jjj).w
						h = FH_bricks(iii, jjj).h
						wx = FH_bricks(iii, jjj).wx
						wy = FH_bricks(iii, jjj).wy
						wz = FH_bricks(iii, jjj).wz
						brickno = iii: facepairno = jjj
						find_brick = True
						Exit Function
					End If
				End If
			Next jjj
		End If	' bricks_found = 0
	Next iii
	find_brick = False

End Function

Function check_midpoints(nbrick As Long, facepair As Long, ncurve As Long, nseg As Long) As Boolean
	Dim nc1 As Long, nc2 As Long
	Dim xyz_c1(5) As Double, xyz_c2(5) As Double	' Curve segment coordinates x,y,z in array elements 1,2,3
	Dim xyz_b1(5) As Double, xyz_b2(5) As Double	' Brick midface coordinates x,y,z in array elements 1,2,3
	Dim c1_eq_b1 As Boolean, c1_eq_b2 As Boolean, c2_eq_b1 As Boolean, c2_eq_b2 As Boolean

	' Nodes of the curve segment
	nc1 = FH_curves(ncurve).point_ini(nseg)
	nc2 = FH_curves(ncurve).point_fin(nseg)
	' Coordinates of the curve segment points. xyz_c(1) contains the x, y, z coords for initial point
	xyz_c1(1) = curve_nodes(nc1).x:  xyz_c2(1) = curve_nodes(nc2).x
	xyz_c1(2) = curve_nodes(nc1).y:  xyz_c2(2) = curve_nodes(nc2).y
	xyz_c1(3) = curve_nodes(nc1).z:  xyz_c2(3) = curve_nodes(nc2).z

	' Nodes of the brick's opposite midfaces
	xyz_b1(1) = FH_bricks(nbrick, facepair).P1.x
	xyz_b1(2) = FH_bricks(nbrick, facepair).P1.y
	xyz_b1(3) = FH_bricks(nbrick, facepair).P1.z

	xyz_b2(1) = FH_bricks(nbrick, facepair).P2.x
	xyz_b2(2) = FH_bricks(nbrick, facepair).P2.y
	xyz_b2(3) = FH_bricks(nbrick, facepair).P2.z

	' Compare
	c1_eq_b1 = Compare_brick_seg(xyz_c1, xyz_b1)
	c1_eq_b2 = Compare_brick_seg(xyz_c1, xyz_b2)
	c2_eq_b1 = Compare_brick_seg(xyz_c2, xyz_b1)
	c2_eq_b2 = Compare_brick_seg(xyz_c2, xyz_b2)

	If ((c1_eq_b1 And c2_eq_b2) Or (c1_eq_b2 And c2_eq_b1)) = True Then
		check_midpoints = True
	Else
		check_midpoints = False
	End If

End Function

Function check_ParallelCurves(nbrick As Long, facepair As Long, ncurve As Long, nseg As Long) As Boolean
	Dim nc1 As Long, nc2 As Long
	Dim xyz_c1(5) As Double, xyz_c2(5) As Double	' Curve segment coordinates x,y,z in array elements 1,2,3
	Dim xyz_b1(5) As Double, xyz_b2(5) As Double	' Brick midface coordinates x,y,z in array elements 1,2,3
'	Dim c1_eq_b1 As Boolean, c1_eq_b2 As Boolean, c2_eq_b1 As Boolean, c2_eq_b2 As Boolean
	Dim scalprod As Double

	' Nodes of the curve segment
	nc1 = FH_curves(ncurve).point_ini(nseg)
	nc2 = FH_curves(ncurve).point_fin(nseg)
	' Coordinates of the curve segment points. xyz_c(1) contains the x, y, z coords for initial point
	xyz_c1(1) = curve_nodes(nc1).x:  xyz_c2(1) = curve_nodes(nc2).x
	xyz_c1(2) = curve_nodes(nc1).y:  xyz_c2(2) = curve_nodes(nc2).y
	xyz_c1(3) = curve_nodes(nc1).z:  xyz_c2(3) = curve_nodes(nc2).z

	' Nodes of the brick's opposite midfaces
	xyz_b1(1) = FH_bricks(nbrick, facepair).P1.x
	xyz_b1(2) = FH_bricks(nbrick, facepair).P1.y
	xyz_b1(3) = FH_bricks(nbrick, facepair).P1.z

	xyz_b2(1) = FH_bricks(nbrick, facepair).P2.x
	xyz_b2(2) = FH_bricks(nbrick, facepair).P2.y
	xyz_b2(3) = FH_bricks(nbrick, facepair).P2.z

	' Check if parallel by performing a scalar product; only on a single direction the scalar product is nonzero
	scalprod = (xyz_c2(1)-xyz_c1(1))*(xyz_b2(1)-xyz_b1(1)) + (xyz_c2(2)-xyz_c1(2))*(xyz_b2(2)-xyz_b1(2)) + (xyz_c2(3)-xyz_c1(3))*(xyz_b2(3)-xyz_b1(3))

	If scalprod = 0 Then
		check_ParallelCurves = False
	Else
		check_ParallelCurves = True
	End If

End Function

Function Compare_brick_seg(xyz1() As Double, xyz2() As Double) As Boolean
	Dim iii As Integer

	For iii = 1 To 3
		'If Abs(xyz1(iii) - xyz2(iii))/abs(xyz1(iii)) > USER_EPS Then	' Variant: relative error
		If Abs(xyz1(iii) - xyz2(iii)) > USER_EPS Then
			Compare_brick_seg = False
			Exit Function
		End If
	Next iii
	Compare_brick_seg = True
End Function


Sub read_solids()
	Dim nshapes As Long
	Dim sName As String
	Dim nv As Integer

	Dim idv As Integer
	Dim x As Double, y As Double, z As Double
	Dim xx(9) As Double, yy(9) As Double, zz(9) As Double

	Dim iii As Long, jjj As Long
	Dim compName As String

	nshapes = Solid.GetNumberOfShapes
	ReDim FH_bricks(nshapes, 3)
	ReDim FH_brick_properties(nshapes)

	ReDim bricks_found(nshapes+1)	' For later ...

	Dim tmpmat As String, kappax As Double, kappay As Double, kappaz As Double

	ngrounds = 0
	nbricks = 0

	compName = ""
	For iii = 0 To nshapes-1	' Fill in component information
		sName = Solid.GetNameOfShapeFromIndex (iii)
		compName = Left(sName, InStr(sName, ":")-1)
	Next iii

	For iii = 0 To nshapes-1
		sName = Solid.GetNameOfShapeFromIndex (iii)

		nv = Solid.GetNumberOfPoints(sName)

		If (nv = 8) Then	' Found a brick
			tmpmat = Solid.GetMaterialNameForShape(sName)
			Material.getKappa(tmpmat, kappax, kappay, kappaz)
			If kappax <= 0 Then
				MsgBox "Please define a nonzero, positive conductivity for material:  """ + tmpmat +"""."+vbCrLf + "Macro stops.", vbOkOnly+vbCritical, "Zero conductivity value"
				Exit All
			End If
			For idv=1 To nv
				If ( SolidPointCoordinatesSTEP(sName, Str(idv), x, y, z) ) Then
				'	MsgBox( sName + " " + Str(idv) + ": " + Str(x) + "," + Str(y) + "," + Str(z))
					xx(idv) = x: yy(idv)=y: zz(idv)=z
				Else
					MsgBox("Could not determine node coordinates for solid " + sName, vbCritical)
					Exit All
				End If
			Next idv
			If (StrComp(LCase(Left(sName, 6)), "ground", vbBinaryCompare) <>0) Then		' Brick is not a ground
				FillIn_brick_mf(nbricks, xx, yy, zz)
				FH_brick_properties(nbricks).bname = sName
				FH_brick_properties(nbricks).sigma = kappax * Units.GetGeometryUnitToSI
				nbricks = nbricks+1

			' if Groundplane then ...
			Else	' Fill in a part of the information about the groundplanes
				groundplanes(ngrounds).gname = sName
				generate_one_groundplane(ngrounds, xx, yy, zz)
				groundplanes(ngrounds).sigma = kappax * Units.GetGeometryUnitToSI
				ngrounds = ngrounds+1
			End If
'=========================================
		'show_brick_mf(sName, nbricks)
'=========================================
		End If	' Found a brick
	Next iii

	If nbricks = 0 Then
		MsgBox "No bricks found. Macro stops", vbCritical, "Error"
		Exit All
	End If

End Sub

Sub generate_one_groundplane(ng As Long, xx() As Double, yy() As Double, zz() As Double)
	Dim iii As Long, node As Long
	Dim edgelen(13) As Double, edgem(13) As point
	Dim xf1(5) As Double, yf1(5) As Double, zf1(5) As Double	' Node coordinates

	Dim tmpstr As String

	Dim g_edge
'	g_edge = Array(0,0,0,  0,7,6, 0,4,1, 0,3,2,   0,7,8, 0,6,5, 0,1,2,  0,7,4, 0,6,1, 0,5,2 )
	g_edge = Array(0,0,0,  0,3,2, 0,4,1, 0,7,6,   0,1,2, 0,6,5, 0,7,8,  0,7,4, 0,6,1, 0,5,2 ) ' No circular permutations ...

	Dim imin As Long, edgelenmin As Double

	imin = 0
	edgelenmin = 1e20

	For iii = 1 To 9	' Go through the "interesting" edges and determine their lengths
		For node = 1 To 2	' Nodes of the edge iii
			xf1(node) = xx(g_edge(iii*3 + node))
			yf1(node) = yy(g_edge(iii*3 + node))
			zf1(node) = zz(g_edge(iii*3 + node))
		Next node
		edgelen(iii) = edge_length(xf1, yf1, zf1)
		edgem(iii) = edge_midpoint(xf1, yf1, zf1)
		If edgelen(iii) < edgelenmin Then
			imin = iii
			edgelenmin = edgelen(iii)
		End If
	Next iii

	' ======== Determine the vector normal to the plane
	iii = imin
		For node = 1 To 2	' Nodes of the edge iii
			xf1(node) = xx(g_edge(iii*3 + node))
			yf1(node) = yy(g_edge(iii*3 + node))
			zf1(node) = zz(g_edge(iii*3 + node))
		Next node
		edge_vector_scaled(xf1, yf1, zf1)	' The vector is placed in component 3
		groundplanes(ng).normal.x = xf1(3)
		groundplanes(ng).normal.y = yf1(3)
		groundplanes(ng).normal.z = zf1(3)

	' ========= Consider the smallest dimension of the groundplane as the "z" dimension
	For iii = imin To imin+2	' The three points characterizing the groundplane
		groundplanes(ng).P(iii-imin) = edgem(iii)
	Next iii
	groundplanes(ng).thick = edgelenmin
	
	tmpstr = groundplanes(ng).gname
	tmpstr = Left( tmpstr, InStr (tmpstr, ":")-1)	'Right( tmpstr, Len(tmpstr) - InStr (tmpstr, ":"))

	Begin Dialog UserDialog 430,350,"Define ground: "+tmpstr,.DialogFunc_Ground ' %GRID:10,7,1,1
		GroupBox 20,84,400,140,"Initial groundplane discretization",.GroupBox2
		Text 30,126,210,14,"Initial discretization stepsize:",.Text1
		TextBox 210,133,90,21,.stepsize
		GroupBox 20,7,180,70,"Groundplane type",.gnd_typea
		OptionGroup .gnd_type
			OptionButton 30,28,110,14,"Nonuniform",.OptionButton1
			OptionButton 30,49,90,14,"Uniform",.OptionButton2
		GroupBox 20,231,400,77,"Refinement below/above traces",.GroupBox3
		CheckBox 30,252,160,14,"Refine below traces",.refine_traces
		CheckBox 240,252,160,14,"Refine below ports",.refine_ports
		Text 30,140,170,14,"(0 = no initial discretization)",.Text2
		OKButton 70,315,90,21
		CancelButton 180,315,90,21
		Text 310,140,70,14,Units.GetUnit("Length"),.Text3
		GroupBox 30,161,180,56,"Type of initial mesh grid",.GroupBox1
		OptionGroup .init_mesh_type
			OptionButton 50,175,130,14,"Uniform grid",.OptionButton3
			OptionButton 50,196,130,14,"Meshed grid",.OptionButton4
		CheckBox 40,105,240,14,"Use initial gnd discretization",.use_gnd_discr
		Text 30,273,90,14,"Scale factor",.Text4
		Text 240,273,90,14,"Ref. ratio",.Text6
		TextBox 130,273,90,21,.scale_factor
		TextBox 320,273,90,21,.scale_factor_port
		Text 30,287,90,14,"for traces (>1)",.Text5
		Text 240,287,70,14,"for ports",.Text7
	End Dialog
	Dim dlg As UserDialog

	dlg.gnd_type = 0
	dlg.refine_traces = 1
	dlg.refine_ports = 1
	dlg.stepsize = "1"
	dlg.use_gnd_discr = 1
	dlg.scale_factor = "1"
	dlg.scale_factor_port = "5"

	If Dialog(dlg) = 0 Then	   Exit All

	Dim stepsize As Double, typ As Integer
	Dim m1 As Long, m2 As Long, tmpint As Long

	m1 = imin+3: If m1 > 9 Then m1 = m1-9
	m2 =   m1+3: If m2 > 9 Then m2 = m2-9


	stepsize = Val(dlg.stepsize)
	typ = dlg.gnd_type

	If typ = 0 Then
		groundplanes(ng).typ = "Nonuniform"
	Else
		groundplanes(ng).typ = "Uniform"
	End If

	If dlg.use_gnd_discr = 0 Then
		groundplanes(ng).init_grid = "None"
	Else
		If dlg.init_mesh_type = 0 Then
			groundplanes(ng).init_grid = "Uniform"
		Else
			groundplanes(ng).init_grid = "Meshed"
		End If
		tmpint = Int( edgelen(m1) / stepsize):		If tmpint <> edgelen(m1)/stepsize Then tmpint = tmpint+1
		groundplanes(ng).init_grid_segs(0) = tmpint
		tmpint = Int( edgelen(m2) / stepsize): If tmpint <> edgelen(m2)/stepsize Then tmpint = tmpint+1
		groundplanes(ng).init_grid_segs(1) = tmpint
	End If
	groundplanes(ng).refine_traces = dlg.refine_traces
	groundplanes(ng).trace_refinement_scale_factor = Val(dlg.scale_factor)
	groundplanes(ng).refine_ports = dlg.refine_ports
	groundplanes(ng).port_refinement_ratio = Val(dlg.scale_factor_port)

End Sub

Function DialogFunc_Ground(DlgItem$, Action%, SuppValue?)   As Boolean

	Dim tmpscale As Double

    Select Case Action%
    Case 1 ' Dialog box initialization
    	DlgEnable "OptionButton2", 0
    	DlgEnable "OptionButton4", 0
    Case 2 ' Value changing or button pressed
    Case 3 ' TextBox or ComboBox text changed
    	If DlgItem$ = "scale_factor" Then
    		tmpscale = Val(DlgText("scale_factor"))
    		If tmpscale < 1 Then
    			MsgBox "Scale factor must be larger than 1"
	    		DlgText("scale_factor", "1")
    		End If
    	End If

    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function



Sub project_point(x() As Double, y() As Double, z() As Double, iii As Long, nx As Double, ny As Double, nz As Double, xp As Double, yp As Double, zp As Double)
	Dim lx As Double, ly As Double, lz As Double	'
	Dim distance As Double

	Dim px As Double, py As Double, pz As Double	' Projected vector

	' Projected vector - not needed
	lx = x(2)-x(1): ly = y(2)-y(1): lz=z(2)-z(1)
	px = (1-nx*nx) * lx +  (-nx*ny) * ly +  (-nx*nz) * lz
	py =  (-ny*nx) * lx + (1-ny*ny) * ly +  (-ny*nz) * lz
	pz =  (-nz*nx) * lx +  (-nz*ny) * ly + (1-nz*nz) * lz

'	iii = 1
	lx = x(iii)-xp
	ly = y(iii)-yp
	lz = z(iii)-zp
	distance = nx*lx + ny*ly + nz*lz

	x(iii+2) = x(iii)-distance*nx
	y(iii+2) = y(iii)-distance*ny
	z(iii+2) = z(iii)-distance*nz
End Sub

Function project_segment(x() As Double, y() As Double, z() As Double, nx As Double, ny As Double, nz As Double, xp As Double, yp As Double, zp As Double) As Boolean
	Dim lx As Double, ly As Double, lz As Double	'
	Dim distance As Double

	Dim px As Double, py As Double, pz As Double	' Projected vector
	Dim iii As Integer

	' Projected vector - not needed
	lx = x(2)-x(1): ly = y(2)-y(1): lz=z(2)-z(1)
	px = (1-nx*nx) * lx +  (-nx*ny) * ly +  (-nx*nz) * lz
	py =  (-ny*nx) * lx + (1-ny*ny) * ly +  (-ny*nz) * lz
	pz =  (-nz*nx) * lx +  (-nz*ny) * ly + (1-nz*nz) * lz

	For iii = 1 To 2
		lx = x(iii)-xp
		ly = y(iii)-yp
		lz = z(iii)-zp
		distance = nx*lx + ny*ly + nz*lz

		x(iii+2) = x(iii)-distance*nx
		y(iii+2) = y(iii)-distance*ny
		z(iii+2) = z(iii)-distance*nz
	Next iii

	Dim eps As Double
	eps = 1e-10

	If Abs(x(3) -x(4)) < eps And Abs(y(3)-y(4)) < eps And _
			Abs(z(3) - z(4)) < eps Then
		project_segment = False		' We projected a segment and got a point - no good!
	Else
		project_segment = True
	End If

End Function


Sub read_curves
	Dim cName As String, csName As String, treeName As String
	Dim nsegs As Integer, npolygons As Integer, nv As Integer
	Dim curveSegNames (MAX_CURVES) As String
	Dim tmpnode As Long

	Dim curve_is_open(MAX_CURVES) As Boolean

	Dim CurveNo As Long

	Dim idv As Integer
	Dim x As Double, y As Double, z As Double
	Dim xx(5) As Double, yy(5) As Double, zz(5) As Double

Dim iii As Long, tmpstr As String

Dim ngroundcurves As Long

	' Determine closed curves
	closed_curves(curve_is_open)

	ncurves = Curve.StartCurveNameIteration ("all")
	If ncurves = 0 Then
		MsgBox "No curves found. Please read the documentation for more information" + vbCrLf + vbCrLf +"Macro stops", vbCritical, "Error"
'		Exit All
	End If
	ReDim FH_curves(ncurves)

	nnodes = 0
	nsegs_tot = 0
	nports = 0
	ngroundcurves = 0

'treeName = "Curves\"+cName
'SelectTreeItem treeName

	For CurveNo = 0 To ncurves-1
		cName = Curve.GetNextCurveName
'		cName = Resulttree.GetFirstChildName (treeName)
		If (StrComp(LCase(Left(cName, 6)), "ground", vbBinaryCompare) =0) Then
			ngroundcurves = ngroundcurves+1
		Else


		treeName = "Curves\"+cName
		SelectTreeItem treeName

		' ====== Determine the number of curve polygons and their names
		' A curve may contain several polygons, each polygon several segments
		npolygons = 0
		nsegs = 0

		csName = Resulttree.GetFirstChildName (treeName)
		While csName <> ""
			npolygons = npolygons+1
			tmpstr = csName
			tmpstr = Replace(tmpstr, "\", ":")
			tmpstr = Right( tmpstr, Len(tmpstr) - InStr (tmpstr, ":"))
			curveSegNames(npolygons) = tmpstr
	    	nsegs = nsegs + Curve.GetNumberOfPoints(tmpstr)	' One segment less than number of points for open curves
	    	If curve_is_open(CurveNo) = True Then nsegs = nsegs-1
			csName = Resulttree.GetNextItemName(csName)
		Wend

		' ====== If nsegs is 0, then the curve is empty. Stop macro
		If nsegs = 0 Then
			MsgBox "Found one empty curve. Macro stops", vbCritical, "Error"
			Exit All
		End If

		' ====== Fill in some information in the curve data structure
		If InStr(cName, "+parallel") <> 0 Then
			FH_curves(CurveNo).typ = "Parallel"
		Else
			FH_curves(CurveNo).typ = "Series"
		End If
		FH_curves(CurveNo).nsegs = nsegs
		nsegs_tot = nsegs_tot + nsegs

		ReDim FH_curves(CurveNo).point_ini(nsegs+1)
		ReDim FH_curves(CurveNo).point_fin(nsegs+1)
		ReDim FH_curves(CurveNo).leng(nsegs+1)
		ReDim FH_curves(CurveNo).csName(nsegs+1)

		read_one_polygon(CurveNo, npolygons, curveSegNames, curve_is_open)

'=========================================
	'show_curve (cName, CurveNo, nsegs )
'=========================================
	End If	' Curve name does not start with "ground"; foreseen for possible future developments
	Next CurveNo

	ncurves = ncurves - ngroundcurves

'=========================================
	'show_nodes()
'=========================================

' Sort curves
sort_curves()

End Sub

Sub sort_curves()
	Dim iii As Long, jjj As Long
	For iii = 0 To ncurves-2
		For jjj = iii+1 To ncurves-1
			If StrComp(FH_curves(iii).cName, "") = 0 Or StrComp(FH_curves(jjj).cName, "") = 0 Then
				MsgBox "Empty curve name. Maybe a curve without any polygon? Macro stops"
				Exit All
			End If
			If StrComp(FH_curves(iii).cName, FH_curves(jjj).cname) = 1 Then 	' Name of curve jjj > Name of curve iii
				permute_curves(iii,jjj)
			End If
		Next
	Next
End Sub

Sub permute_curves(iii As Long, jjj As Long)
	Dim tmpcurve As curve_single

	tmpcurve = FH_curves(iii)
	FH_curves(iii) = FH_curves(jjj)
	FH_curves(jjj) = tmpcurve
End Sub


Sub read_one_polygon(CurveNo As Long, npolygons As Long, curveSegNames() As String, curve_is_open() As Boolean)

	Dim nsegs As Long
	Dim iii As Long
	Dim csName As String
	Dim nv As Long, idv As Long
	Dim x As Double, y As Double, z As Double
	Dim xx(5) As Double, yy(5) As Double, zz(5) As Double
	Dim tmpstr As String, tmpnode As Long

	Dim nn As Long, crtnodes(MAX_NODES) As point_ext, crtcurves(MAX_NODES,2) As Long, crtcurvenames(MAX_NODES) As String

	Dim curve_has_ports As Boolean

	tmpstr = curveSegNames(1)
	If (InStr(tmpstr, "-ports") <> 0) Then ' Curve has no ports
		curve_has_ports = False
	Else
		curve_has_ports = True
	End If

	nn = 0

		' ===== For each polygon, determine point coordinates
		nsegs = 0	' Reuse nsegs to count again the number of segments of a curve
		For iii = 1 To npolygons
			csName = curveSegNames(iii)
	    	nv = Curve.GetNumberOfPoints(csName)
			FH_curves(CurveNo).cname = csName
			For idv=1 To nv
				nn = nn+1
				nsegs = nsegs+1
				FH_curves(CurveNo).csname(nsegs) = csName
				If ( CurvePointCoordinatesSTEP(csName, Str(idv), x, y, z) ) Then
				'	MsgBox( sName + " " + Str(idv) + ": " + Str(x) + "," + Str(y) + "," + Str(z))
					crtnodes(nn).x = x
					crtnodes(nn).y = y
					crtnodes(nn).z = z
					If idv < nv Then	' I am not at the last segment
						' ===== Fill in initial node of the segment
						crtcurves(nsegs,0) = nn: crtnodes(nn).taken = crtnodes(nn).taken+1
						crtcurvenames(nsegs) = csName
						If (idv > 1) Then
							' ======= Fill in final node of the previous segment
							crtcurves(nsegs-1,1) = nn: crtnodes(nn).taken = crtnodes(nn).taken+1
							crtcurvenames(nsegs-1) = csName
						End If
					Else
						' ======= Fill in final node of the last segment
						If (curve_is_open(CurveNo) = True Or npolygons > 1) Then	' Curve is open
							If nsegs > 1 Then nsegs = nsegs-1
							crtcurves(nsegs,1) = nn: crtnodes(nn).taken = crtnodes(nn).taken+1
							crtcurvenames(nsegs) = csName
						Else
							crtcurves(nsegs-1,1) = nn: crtnodes(nn).taken = crtnodes(nn).taken+1
							crtcurves(nsegs,0) = nn: crtnodes(nn).taken = crtnodes(nn).taken+1
							crtcurvenames(nsegs) = csName: crtcurvenames(nsegs-1) = csName


							' Double the last node (why? always?)
							tmpnode = crtcurves(1,0)
							crtcurves(nsegs,1) = tmpnode
							crtcurvenames(nsegs) = csName
						End If
					End If
				Else
					MsgBox("Could not determine node coordinates for curve " + csName)
					Exit All
				End If	' Found the curve coordinates
			Next idv	' Nodes xof a polygon
		Next iii	' Polygons

		nn = sort_polygons(nsegs, nn, crtnodes, crtcurves, crtcurvenames, curve_has_ports, curve_is_open(CurveNo))

		For iii = 1 To nsegs
				FH_curves(CurveNo).point_ini(iii) = crtcurves(iii,0)+nnodes
				FH_curves(CurveNo).point_fin(iii) = crtcurves(iii,1)+nnodes
				FH_curves(CurveNo).csname(iii) = crtcurvenames(iii)
		Next iii

		For iii = 1 To nn
			curve_nodes(nnodes+iii).x = crtnodes(iii).x
			curve_nodes(nnodes+iii).y = crtnodes(iii).y
			curve_nodes(nnodes+iii).z = crtnodes(iii).z
			curve_nodes(nnodes+iii).taken = crtnodes(iii).taken
		Next iii

		nnodes = nnodes + nn

		For iii = 1 To nsegs
			xx(1) = curve_nodes(FH_curves(CurveNo).point_ini(iii)).x
			xx(2) = curve_nodes(FH_curves(CurveNo).point_fin(iii)).x
			yy(1) = curve_nodes(FH_curves(CurveNo).point_ini(iii)).y
			yy(2) = curve_nodes(FH_curves(CurveNo).point_fin(iii)).y
			zz(1) = curve_nodes(FH_curves(CurveNo).point_ini(iii)).z
			zz(2) = curve_nodes(FH_curves(CurveNo).point_fin(iii)).z
			FH_curves(CurveNo).leng(iii) = edge_length(xx, yy, zz)

		Next iii

		FH_curves(CurveNo).nsegs = nsegs

End Sub


Function sort_polygons(nsegs As Long, nn As Long, crtnodes() As point_ext, crtcurves() As Long, crtcurvenames() As String, curve_has_ports As Boolean, curve_is_open As Boolean) As Long
	Dim n1 As Long, n2 As Long
	Dim tmpcrtnodes(MAX_NODES) As point, tmpcrtcurves(MAX_NODES,2) As Long
	Dim n_sorted_nodes As Long
	Dim seg_taken(MAX_NODES) As Boolean

	Dim n2_crt As Long, tmpn1 As Long, tmpn2 As Long, tmpstr As String
	Dim iii As Long, jjj As Long, kkk As Long

	Dim sorted_nodes(MAX_NODES) As Long
	Dim seg_idx(MAX_NODES) As Long	' Determines the order in which segments are considered
	Dim node_multiplicity(MAX_NODES) As Long	' For each node, how many times it is taken


	sort_poly_nodes(nn,crtnodes,sorted_nodes, n_sorted_nodes, curve_is_open)
'MsgBox display_sorted_nodes(nn, sorted_nodes, 6)
	For iii = 1 To nsegs
		n1 = crtcurves(iii,0)
		n2 = crtcurves(iii,1)

		crtcurves(iii,0) = sorted_nodes(n1)
		crtcurves(iii,1) = sorted_nodes(n2)
	Next iii

'MsgBox display_segments(nsegs, crtcurves, 6)

	' Fill in the segment indexes
	If curve_is_open = False Or nsegs = 1 Then
		kkk = 1
	Else	' Curve is open
		' First, det. the node multiplicity
		For iii = 1 To nsegs
			n1 = crtcurves(iii,0)
			n2 = crtcurves(iii,1)
			For jjj = 1 To nsegs
				If iii <> jjj Then
					If crtcurves(jjj,0) = n1 Then node_multiplicity(n1) = node_multiplicity(n1)+1
					If crtcurves(jjj,1) = n1 Then node_multiplicity(n1) = node_multiplicity(n1)+1
					If crtcurves(jjj,0) = n2 Then node_multiplicity(n2) = node_multiplicity(n2)+1
					If crtcurves(jjj,1) = n2 Then node_multiplicity(n2) = node_multiplicity(n2)+1
				End If
			Next
		Next

		' Go through the segments again and see which one would be a good point to start with (one which has a node with multipl. = 1)
		kkk = 0		' Segment to start with
		For iii = 1 To nsegs-1
			n1 = crtcurves(iii,0)
			n2 = crtcurves(iii,1)
			If node_multiplicity(n1) = 0 Then
				kkk = iii
				Exit For
			End If

			If node_multiplicity(n2) = 0 Then
				' Final node has mult. 1, we resort the nodes of the segment first
				crtcurves(iii,0) = n2
				crtcurves(iii,1) = n1
				kkk = iii
				Exit For
			End If
		Next
	End If	' curve_is_open = False

	If kkk = 0 Then
		MsgBox "Wrong value in sub sort_polygons. Please contact support"
		Exit All
	End If

	If kkk <> 1 Then
		' Permute segment kkk with segment 1
		tmpn1 = crtcurves(kkk,0)
		tmpn2 = crtcurves(kkk,1)
		crtcurves(kkk,0) = crtcurves(1,0)
		crtcurves(kkk,1) = crtcurves(1,1)
		crtcurves(1,0) = tmpn1
		crtcurves(1,1) = tmpn2

		tmpstr = crtcurvenames(kkk)
		crtcurvenames(kkk) = crtcurvenames(1)
		crtcurvenames(1)=tmpstr
'		jjj = nsegs
	End If
'MsgBox display_segments(nsegs, crtcurves, 6)



	' Sort the segments
	For iii = 1 To nsegs
		n2_crt = crtcurves(iii,1)	' Final node of current segment
		For jjj =iii+1 To nsegs
			If iii <> jjj Then
				n1 = crtcurves(jjj,0)
				n2 = crtcurves(jjj,1)

				If n1 = n2_crt Then		' Found the next segment, with same orientation as the previous
					If jjj <> iii + 1 Then	' Need to permute segments iii+1 and jjj
						tmpn1 = crtcurves(iii+1,0)
						tmpn2 = crtcurves(iii+1,1)
						crtcurves(iii+1,0) = n1
						crtcurves(iii+1,1) = n2
						crtcurves(jjj,0) = tmpn1
						crtcurves(jjj,1) = tmpn2

						tmpstr = crtcurvenames(iii+1)
						crtcurvenames(iii+1) = crtcurvenames(jjj)
						crtcurvenames(jjj)=tmpstr
						jjj = nsegs
					End If
				Else
					If n2 = n2_crt Then ' Found the next segment, with opposite orientation as to the previous
						tmpn1 = crtcurves(iii+1,0)
						tmpn2 = crtcurves(iii+1,1)
						crtcurves(iii+1,0) = n2
						crtcurves(iii+1,1) = n1
						If jjj <> iii + 1 Then
							crtcurves(jjj,0) = tmpn1
							crtcurves(jjj,1) = tmpn2
						End If
						tmpstr = crtcurvenames(iii+1)
						crtcurvenames(iii+1) = crtcurvenames(jjj)
						crtcurvenames(jjj)=tmpstr
						jjj = nsegs
					End If
				End If
			End If
		Next jjj
	Next iii

'MsgBox display_segments(nsegs, crtcurves, 6)
	' Do something for closed curves with ports
	' Test if curve is closed
	If crtcurves(nsegs,1) = crtcurves(1,0) And nsegs > 1 Then
		curve_is_open = False
	Else
		curve_is_open = True
	End If

	If curve_has_ports=True And curve_is_open = False Then
		n_sorted_nodes = n_sorted_nodes +1
		crtcurves(nsegs,1) = n_sorted_nodes
		crtnodes(n_sorted_nodes) = crtnodes(1)
	End If
	If nsegs = 1 And crtcurves(1,1) = 0 Then crtcurves(1,1) = crtcurves(1,0)	' Deal with a port defined by using the same point twice

	sort_polygons = n_sorted_nodes

End Function

Function display_segments(nsegs As Long, crtcurves() As Long, items_per_line As Integer) As String

	Dim tmpstr As String, iii As Long, jjj As Long

	tmpstr = vbCrLf

	For iii = 1 To nsegs
		tmpstr = tmpstr + CStr(crtcurves(iii,0)) + " " + CStr(crtcurves(iii,1)) + ", "
		If iii Mod items_per_line = 0 Then tmpstr = tmpstr + vbCrLf
	Next

	display_segments = tmpstr

End Function

Function display_segments_coords(nsegs As Long, crtcurves() As Long, crtnodes() As point_ext, items_per_line As Integer) As String

	Dim tmpstr As String, iii As Long, jjj As Long

	tmpstr = vbCrLf

	For iii = 1 To nsegs
		tmpstr = tmpstr + CStr(crtcurves(iii,0)) + " "
		tmpstr = tmpstr + cstr(crtnodes(crtcurves(iii,0)).x) + " " + cstr(crtnodes(crtcurves(iii,0)).y) +" " + cstr(crtnodes(crtcurves(iii,0)).z) + ", "
		tmpstr = tmpstr + CStr(crtcurves(iii,1)) + " "
		tmpstr = tmpstr + cstr(crtnodes(crtcurves(iii,1)).x) + " " + cstr(crtnodes(crtcurves(iii,1)).y) +" " + cstr(crtnodes(crtcurves(iii,1)).z) + ", "
		If iii Mod items_per_line = 0 Then tmpstr = tmpstr + vbCrLf

'		If iii Mod items_per_line = 0 Then tmpstr = tmpstr + vbCrLf
	Next

	display_segments_coords = tmpstr

End Function

Sub sort_poly_nodes(nn As Long,crtnodes() As point_ext, sorted_nodes() As Long, n_sorted_nodes As Long, curve_is_open As Boolean)
	Dim node_used(MAX_NODES) As Boolean
	Dim iii As Long, jjj As Long, tmpn As Long, prev_node As Long, sorted_pnt(MAX_NODES) As point_ext, tmppnt As point_ext

	Dim eps As Double

	eps = USER_EPS	' ???

	n_sorted_nodes = nn

	sorted_nodes(1) = 1
	sorted_pnt(1) = crtnodes(1)

	' First, eliminate the identical nodes
	For iii = 2 To nn
		sorted_nodes(iii) = iii
		sorted_pnt(iii) = crtnodes(iii)
		For jjj = 1 To iii-1
			If nodes_identical(iii, jjj, crtnodes, eps) Then
				n_sorted_nodes = n_sorted_nodes-1
				sorted_nodes(iii) = jjj
				sorted_pnt(iii) = crtnodes(jjj)
				jjj = iii
			End If
		Next jjj
	Next iii

'MsgBox display_sorted_nodes(nn, sorted_nodes, 6)

	' Then, renumber the nodes so that they are numbered continously
	Dim tmp_sorted_nodes(MAX_NODES) As Long, poz(MAX_NODES) As Long

	' Initializations
	For iii = 1 To nn
		tmp_sorted_nodes(iii) = sorted_nodes(iii)
		poz(iii) = iii
	Next iii

	For iii = 1 To nn-1
		For jjj = iii+1 To nn
			If tmp_sorted_nodes(jjj) < tmp_sorted_nodes(iii) Then
				tmpn = tmp_sorted_nodes(iii)
				tmp_sorted_nodes(iii) = tmp_sorted_nodes(jjj)
				tmp_sorted_nodes(jjj) = tmpn

				tmpn = poz(iii)
				poz(iii) = poz(jjj)
				poz(jjj) = tmpn

				tmppnt = sorted_pnt(iii)
				sorted_pnt(iii) = sorted_pnt(jjj)
				sorted_pnt(jjj) = tmppnt
			End If
		Next jjj
	Next iii

	prev_node = 0
	tmpn = 0
	For iii = 1 To nn
		If tmp_sorted_nodes(iii) <> prev_node Then
			tmpn = tmp_sorted_nodes(iii) - tmp_sorted_nodes(iii-1)-1
			If tmpn > 0 Then
				For jjj = iii To nn
					tmp_sorted_nodes(jjj) = tmp_sorted_nodes(jjj)-tmpn
					sorted_nodes(poz(jjj)) = sorted_nodes(poz(jjj))-tmpn
				Next jjj
			End If
			prev_node = tmp_sorted_nodes(iii)
		End If
	Next iii

	jjj = 1
	crtnodes(1) = sorted_pnt(1)
		For iii = 2 To nn
			If tmp_sorted_nodes(iii) <> tmp_sorted_nodes(iii-1) Then
				jjj = jjj+1
				crtnodes(jjj) = sorted_pnt(iii)
			End If
		Next

'	MsgBox display_sorted_nodes(nn, sorted_nodes, 6)

	' If the curve is open, then the segments might need to be resorted
End Sub

Function display_sorted_nodes(nn As Long, sorted_nodes() As Long, items_per_line As Integer) As String
	Dim tmpstr As String, iii As Long, jjj As Long

	tmpstr = vbCrLf

	For iii = 1 To nn
		tmpstr = tmpstr + cstr(sorted_nodes(iii)) + " "
		If iii Mod items_per_line = 0 Then tmpstr = tmpstr + vbCrLf
	Next

	display_sorted_nodes = tmpstr

End Function

Function display_nodes(nn As Long, crtnodes() As point_ext, items_per_line As Integer) As String
	Dim tmpstr As String, iii As Long, jjj As Long

	tmpstr = vbCrLf

	For iii = 1 To nn
		tmpstr = tmpstr + cstr(crtnodes(iii).x) + " " + cstr(crtnodes(iii).y) +" " + cstr(crtnodes(iii).z) + ", "
		If iii Mod items_per_line = 0 Then tmpstr = tmpstr + vbCrLf
	Next

	display_nodes = tmpstr

End Function

Function nodes_identical(n1 As Long, n2 As Long, crtnodes() As point_ext, eps As Double) As Boolean
	If Abs(crtnodes(n1).x -crtnodes(n2).x) < eps And Abs(crtnodes(n1).y -crtnodes(n2).y) < eps And _
			Abs(crtnodes(n1).z -crtnodes(n2).z) < eps Then
		nodes_identical = True
	Else
		nodes_identical = False
	End If
End Function

Sub closed_curves(curve_is_open() As Boolean)

	Dim CurveNo As Long, iii As Long
	Dim nc As Long
	Dim nc_c As Long, nc_o As Long	' Number of closed, open curves
	Dim opencn(MAX_CURVES) As String, closedcn(MAX_CURVES) As String
	Dim tmpstr As String

	nc_c = Curve.StartCurveNameIteration ("closed")
	For CurveNo = 0 To nc_c-1
		closedcn(CurveNo) = Curve.GetNextCurveName
	Next CurveNo

	nc_o = Curve.StartCurveNameIteration ("open")
	For CurveNo = 0 To nc_o-1
		opencn(CurveNo) = Curve.GetNextCurveName
	Next CurveNo

	nc = Curve.StartCurveNameIteration ("all")
	For CurveNo = 0 To nc-1
		tmpstr = Curve.GetNextCurveName
		curve_is_open(CurveNo) = True
		For iii = 0 To nc_c-1
			If (StrComp(tmpstr, closedcn(iii),vbBinaryCompare)=0) Then	curve_is_open(CurveNo) = False
		Next
	Next CurveNo

End Sub

Sub Det_FaceNodes(xx() As Double, yy() As Double, zz() As Double, fnodes() As Long, crEdges() As Long, wEdge() As Long, hEdge() As Long)
	Dim iii As Long, jjj As Long, kkk As Long
	Dim n1 As Long, n2 As Long
	Dim tmpnodes(0 To 8) As Long

	Dim xmin As Double, ymin As Double, zmin As Double
	Dim xmax As Double, ymax As Double, zmax As Double

	Dim maxdist As Double, tmp As Double

	Dim facenodes
	facenodes= Array(0,0,0,0,0,  0,3,4,7,8,  0,1,6,5,2,  0,1,6,7,4,  0,2,5,8,3, 0,8,7,6,5, 0,1,2,3,4)

	Dim crossedges
	crossedges = Array(0,0,0,  0,7,6,  0,7,8,  0,7,4)

	Dim w_edge
	w_edge = Array(0,0,0, 0,7,8, 0,7,4, 0,7,6)

	Dim h_edge
	h_edge = Array(0,0,0, 0,7,4, 0,7,6, 0,7,8)


	' First find the longest distance from node 1 - this is the diagonal
	maxdist = 0
	tmpnodes(1) = 1
	tmpnodes(8) = maxdist_node(1, xx, yy, zz)


	' Find first face
	iii = 1: jjj = 2
	If tmpnodes(8) = jjj Then jjj = jjj+1
	n1 = 0: n2 = 0
	For n1 = 2 To 8
		If n1 <> tmpnodes(8) And n1 <> jjj Then
			For n2 = 2 To 8
				If n2<> tmpnodes(8) And n2 <> jjj And n2 <> n1 Then
					If coplanar(iii, jjj, n1, n2, xx, yy, zz ) Then
						kkk = 8
						GoTo NEXT_TMP
					End If
				End If
			Next n2
		End If
	Next

NEXT_TMP:
	tmpnodes(2) = jjj
	If dist(iii, n1, xx, yy, zz) > dist(iii, n2, xx, yy, zz) Then
		' Already correct order
		tmpnodes(3) = n1
		tmpnodes(4) = n2
	Else
		tmpnodes(3) = n2
		tmpnodes(4) = n1
	End If

	' Now find the other pair, in the standard case 2-3-5-8
	tmpnodes(5) = maxdist_node(tmpnodes(4), xx, yy, zz)
	tmpnodes(6) = maxdist_node(tmpnodes(3), xx, yy, zz)
	tmpnodes(7) = maxdist_node(tmpnodes(2), xx, yy, zz)

	' Fill in the face nodes
	For iii = 0 To 34
		fnodes(iii) = tmpnodes(facenodes(iii))
	Next

	' Fill in the crossedges
	For iii = 0 To 11
		crEdges(iii) = tmpnodes(crossedges(iii))
	Next

	' Fill in edge widths
	For iii = 0 To 11
		wEdge(iii) = tmpnodes(w_edge(iii))
	Next

	' Fill in the crossedges
	For iii = 0 To 11
		hEdge(iii) = tmpnodes(h_edge(iii))
	Next

End Sub

Function dist(iii As Long, jjj As Long, xx() As Double, yy() As Double, zz() As Double) As Double
	Dim tmp As Double
	tmp = (xx(iii)-xx(jjj))^2 + (yy(iii)-yy(jjj))^2 + (zz(iii)-zz(jjj))^2
	dist = Sqr(tmp)
End Function

Function maxdist_node(iii As Long, xx() As Double, yy() As Double, zz() As Double) As Long
	Dim jjj As Long, n1 As Long, maxdist As Double, tmp As Double

	maxdist = 0
	For jjj = 1 To 8
		tmp = dist(iii, jjj, xx, yy, zz)
		If tmp > maxdist Then
			n1 = jjj
			maxdist = tmp
		End If
	Next
	maxdist_node = n1
End Function



Function coplanar(i As Long, j As Long, k As Long, m As Long, xx() As Double, yy() As Double, zz() As Double) As Boolean
	Dim tmpx As Double, tmpy As Double, tmpz As Double, tmp As Double
	tmpz = (xx(j)-xx(i)) * (yy(k)-yy(m)) - (xx(k)-xx(m)) * (yy(j)-yy(i))
	tmpy = (xx(j)-xx(i)) * (zz(k)-zz(m)) - (xx(k)-xx(m)) * (zz(j)-zz(i))
		tmpy = -tmpy
	tmpx = (zz(j)-zz(i)) * (yy(k)-yy(m)) - (zz(k)-zz(m)) * (yy(j)-yy(i))

	tmp = (xx(m)-xx(i))*tmpx + (yy(m)-yy(i))*tmpy+(zz(m)-zz(i))*tmpz

	If Abs(tmp) < 1e-6 Then
		coplanar = True
	Else
		coplanar = False
	End If

End Function



Sub FillIn_brick_mf(idx As Long, xx() As Double, yy() As Double, zz() As Double)
	' indexing in xx, yy, zz starts at 1, not at 0

	Dim iii As Integer, node As Integer, Face As Integer

	Dim xf1(5) As Double, yf1(5) As Double, zf1(5) As Double	' Node coordinates for first face in pair
	Dim xf2(5) As Double, yf2(5) As Double, zf2(5) As Double	' Node coordinates for second face in pair
	Dim xm, ym, zm As Double	' Midpoint coordinates for a face
	Dim wtmp As Double, wxtmp As Double, wytmp As Double, wztmp As Double
	Dim htmp As Double, hxtmp As Double, hytmp As Double, hztmp As Double

	Dim facenodes(35) As Long, crossedges(12) As Long, w_edge(12) As Long, h_edge(12) As Long
	Det_FaceNodes(xx, yy, zz, facenodes, crossedges, w_edge, h_edge)

'	Dim facenodes

'	facenodes= Array(0,0,0,0,0,  0,3,4,7,8,  0,1,6,5,2,  0,1,6,7,4,  0,2,5,8,3, 0,8,7,6,5, 0,1,2,3,4)
	' It seems that it is not always the case, at least when the brick was obtained by cut by uv plane

'	Dim crossedges
'	crossedges = Array(0,0,0,  0,7,6,  0,7,8,  0,7,4)

'	Dim w_edge
'	w_edge = Array(0,0,0, 0,7,8, 0,7,4, 0,7,6)

'	Dim h_edge
'	h_edge = Array(0,0,0, 0,7,4, 0,7,6, 0,7,8)


	' Fill in the array brick_mf
	For iii = 1 To 3	' Face pairs
		' ======== FIRST CALCULATE MIDPOINTS OF FACE AND FILL IN P1 and P2 in array brick_mf
		Face = iii*2-1	' Current face number; first face in pair
		' Calculate midpoint of first face in pair
		For node = 1 To 4	' Nodes of the face
			xf1(node) = xx(facenodes(Face*5 + node))
			yf1(node) = yy(facenodes(Face*5 + node))
			zf1(node) = zz(facenodes(Face*5 + node))
		Next node
		face_midpoint (xf1, yf1, zf1, xm, ym, zm)
		FH_bricks(idx, iii).P1.x = xm
		FH_bricks(idx, iii).P1.y = ym
		FH_bricks(idx, iii).P1.z = zm

		Face = iii*2	' Current face number; second face in pair
		' Calculate midpoint of face
		For node = 1 To 4	' Nodes of the face
			xf2(node) = xx(facenodes(Face*5 + node))
			yf2(node) = yy(facenodes(Face*5 + node))
			zf2(node) = zz(facenodes(Face*5 + node))
		Next node
		face_midpoint (xf2, yf2, zf2, xm, ym, zm)
		FH_bricks(idx, iii).P2.x = xm
		FH_bricks(idx, iii).P2.y = ym
		FH_bricks(idx, iii).P2.z = zm

		' ======== THEN CALCULATE length of the edge between them
		For node = 1 To 2	' Midnodes of faces representing the cross-edge
			xf1(node) = xx(crossedges(iii*3 + node))
			yf1(node) = yy(crossedges(iii*3 + node))
			zf1(node) = zz(crossedges(iii*3 + node))
		Next node
		FH_bricks(idx, iii).dist = edge_length(xf1, yf1, zf1)

		' ======== CALCULATE vector w
		For node = 1 To 2	' Nodes of the edge representing the width
			xf1(node) = xx(w_edge(iii*3 + node))
			yf1(node) = yy(w_edge(iii*3 + node))
			zf1(node) = zz(w_edge(iii*3 + node))
		Next node
		edge_vector_scaled(xf1, yf1, zf1)	' Results will be put in the element 3 of the array
		wxtmp = xf1(3)
		wytmp = yf1(3)
		wztmp = zf1(3)

		' ======== CALCULATE length w (width)
		wtmp = edge_length(xf1, yf1, zf1)

		' ======== CALCULATE length h (height)
		For node = 1 To 2	' Nodes of the edge representing the width
			xf1(node) = xx(h_edge(iii*3 + node))
			yf1(node) = yy(h_edge(iii*3 + node))
			zf1(node) = zz(h_edge(iii*3 + node))
		Next node
		edge_vector_scaled(xf1, yf1, zf1)	' Results will be put in the element 3 of the array
		hxtmp = xf1(3)
		hytmp = yf1(3)
		hztmp = zf1(3)
		htmp = edge_length(xf1, yf1, zf1)

		' ========= RE-SELECT w as the largest between wtmp and htmp
		If wtmp > htmp Then
			FH_bricks(idx, iii).wx = wxtmp
			FH_bricks(idx, iii).wy = wytmp
			FH_bricks(idx, iii).wz = wztmp

			FH_bricks(idx, iii).w = wtmp
			FH_bricks(idx, iii).h = htmp
		Else
			FH_bricks(idx, iii).wx = hxtmp
			FH_bricks(idx, iii).wy = hytmp
			FH_bricks(idx, iii).wz = hztmp

			FH_bricks(idx, iii).w = htmp
			FH_bricks(idx, iii).h = wtmp
		End If

	Next iii	' Face pairs

End Sub

Sub face_midpoint (xf() As Double, yf() As Double, zf() As Double, xm As Double, ym As Double, zm As Double)
	xm = (xf(1) + xf(2) + xf(3) + xf(4)) /4
	ym = (yf(1) + yf(2) + yf(3) + yf(4)) /4
	zm = (zf(1) + zf(2) + zf(3) + zf(4)) /4
End Sub

Function edge_midpoint (xf() As Double, yf() As Double, zf() As Double) As point
	edge_midpoint.x = (xf(1) + xf(2)) /2
	edge_midpoint.y = (yf(1) + yf(2)) /2
	edge_midpoint.z = (zf(1) + zf(2)) /2
End Function


Function edge_length (xf() As Double, yf() As Double, zf() As Double) As Double
	edge_vector(xf, yf, zf)
	edge_length = Sqr(xf(3)^2 + yf(3)^2+zf(3)^2)
End Function

Sub	edge_vector(xf() As Double, yf() As Double, zf() As Double)
	' Input data are in xf(1) and xf(2)
	' Output data in xf(3)
	Dim length As Double

	xf(3) = xf(2) - xf(1)
	yf(3) = yf(2) - yf(1)
	zf(3) = zf(2) - zf(1)
End Sub

	Sub	edge_vector_scaled(xf() As Double, yf() As Double, zf() As Double)
	' Input data are in xf(1) and xf(2)
	' Output data in xf(3)
	Dim length As Double

	xf(3) = xf(2) - xf(1)
	yf(3) = yf(2) - yf(1)
	zf(3) = zf(2) - zf(1)
	length = Sqr(xf(3)^2 + yf(3)^2+zf(3)^2)
	xf(3) = xf(3) / length
	yf(3) = yf(3) / length
	zf(3) = zf(3) / length

End Sub

Sub show_brick_mf (sName As String, nbricks As Long)

	Dim kkk As Integer

For kkk = 1 To 3
			MsgBox( sName + "  Brick no: " + CStr(nbricks-1) + "  Node: " + CStr(kkk) + vbCrLf _
			+"Midpoint 1: "  + CStr(FH_bricks(nbricks-1,kkk).P1.x) + " " + CStr(FH_bricks(nbricks-1,kkk).P1.y) + " " +CStr(FH_bricks(nbricks-1,kkk).P1.z) + " " + vbCrLf _
			+"Midpoint 2: "  + CStr(FH_bricks(nbricks-1,kkk).P2.x) + " " + CStr(FH_bricks(nbricks-1,kkk).P2.y) + " " +CStr(FH_bricks(nbricks-1,kkk).P2.z) + " " + vbCrLf _
			+"Distance: " + CStr(FH_bricks(nbricks-1,kkk).dist) + " " +vbCrLf _
			+"W: " + CStr(FH_bricks(nbricks-1,kkk).w) + " " +vbCrLf _
			+"H: " + CStr(FH_bricks(nbricks-1,kkk).h) + " " +vbCrLf _
			+"w vector: "    + CStr(FH_bricks(nbricks-1,kkk).wx)   + " " + CStr(FH_bricks(nbricks-1,kkk).wy) +   " " +CStr(FH_bricks(nbricks-1,kkk).wz) )
Next kkk

End Sub

Sub show_curve (cName As String, CurveNo As Long, nsegs As Long)

	Dim iii As Integer
	Dim tmpstr As String

		tmpstr = ""
		For iii = 1 To nsegs
			tmpstr = tmpstr + "Segment " + Cstr(iii) + _
			" Nod ini: " + _
			CStr(FH_curves(CurveNo).point_ini(iii)) + " " + _
			" Nod fin: " + _
			CStr(FH_curves(CurveNo).point_fin(iii)) + " " + _
			" Length: " + _
			CStr(FH_curves(CurveNo).leng(iii)) + " " + _
			vbCrLf
		Next

		MsgBox ("Curve: " + cName + vbCrLf + _
		"Type: "  + FH_curves(CurveNo).typ + vbCrLf + _
		"Nsegs: " + cstr(FH_curves(CurveNo).nsegs ) + vbCrLf + _
		tmpstr)

End Sub

Sub show_nodes
Dim iii As Long, tmpstr As String

	tmpstr = ""
For iii = 0 To nnodes
	tmpstr = tmpstr  _
	+ "Node " + CStr(iii) + "   " _
	+ CStr(curve_nodes(iii).x) + " "  _
	+ CStr(curve_nodes(iii).y) + " "  _
	+ CStr(curve_nodes(iii).z) + " "  _
	+ vbCrLf

Next iii

MsgBox tmpstr

End Sub

Function GetFolderName(Optional OpenAt As String) As String
    Dim lCount As Long

    Dim objExcel As Object

    GetFolderName = "" 'vbNullString
Set objExcel = CreateObject("Excel.Application")

Dim msoFileDialogFolderPicker As Long
'msoFileDialogFolderPicker = 1		' Select a file
msoFileDialogFolderPicker = 4		' Select a directory

'For msoFileDialogFolderPicker = 0 To 10

    With objExcel.Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            GetFolderName = .SelectedItems(lCount)
        Next lCount
    End With
'Next
End Function


'================= PORTUNUS
Sub generate_Portunus_file(fhobj As Object)

  Dim inductance() As Variant, resistance() As Variant
  Dim frequency() As Variant
  Dim rowportnames() As Variant, colportnames() As Variant

  On Error GoTo NORESULT

  inductance = fhobj.GetInductance()
  resistance = fhobj.GetResistance()
  frequency = fhobj.GetFrequencies()
  rowportnames = fhobj.GetRowPortNames()
'  colportnames = fhobj.GetColPortNames()	' Return always null strings?

  Write_Portunus_File(resistance, inductance, frequency, rowportnames, Portunus_result_name)
  Exit Sub

NORESULT:
  MsgBox "No valid result obtained from the FastHenry run. No Portunus file generated." + vbCrLf + "Macro stops.", vbOkOnly+vbCritical, "Error"
  Exit All
End Sub

Sub determine_pins_nodes(rowportnames() As Variant, pinnames() As Variant, nodenames() As Variant, cellnames() As Variant)
	Dim n1 As Long, n2 As Long, iii As Long
	Dim jjj As Long 	' Counter for pins
	Dim kkk As Long		' Counter for cells

	' DOES NOT DETERMINE NODENAMES AT THIS MOMENT
	Dim tmpstr As String

	n1 = LBound(rowportnames, 1)
	n2 = UBound(rowportnames, 1)

	ReDim cellnames(n1 To n2)
'	ReDim pinnames(n1 to n2)
'	ReDim nodenames(n1 to n2)
	ReDim pinnames(0 To MAX_NODES)
	ReDim nodenames(0 To MAX_NODES)


	' Cells
	kkk = 0
	For iii = n1 To n2
		jjj = InStr(rowportnames(iii), ",")
		cellnames(kkk) = Right(rowportnames(iii), Len(rowportnames(iii)) - 1 - jjj)
		kkk = kkk+1
	Next
	ReDim Preserve cellnames(0 To kkk-1)

	' Pins
	Dim node1 As String, node2 As String
	Dim j1 As Long, j2 As Long
	Dim found As Boolean

	kkk = 0
	For iii = n1 To n2
		j1 = InStr(rowportnames(iii), "to")
		j2 = InStr(rowportnames(iii), ",")
		node1 = Left(rowportnames(iii), j1-2)
		node2 = Mid(rowportnames(iii), j1+3, j2-j1-3)
		' Search for node1 in pinnames
		found = False
		For jjj = 0 To kkk
			If StrComp(node1, pinnames(jjj)) = 0 Then found = True
		Next
		If Not found Then
			pinnames(kkk) = node1
			kkk = kkk+1
		End If

		found = False
		For jjj = 0 To kkk
			If StrComp(node2, pinnames(jjj)) = 0 Then found = True
		Next
		If Not found Then
			pinnames(kkk) = node2
			kkk = kkk+1
		End If
	Next

	ReDim Preserve pinnames(0 To kkk-1)

End Sub

Sub Write_Portunus_File(R() As Variant, L() As Variant, F() As Variant, rowportnames() As Variant, outfilen As String)
	Dim tmpstr As String, iii As Long, jjj As Long, kkk As Long, lll As Long

	Dim pinnames() As Variant, nodenames() As Variant, cellnames() As Variant
	determine_pins_nodes(rowportnames, pinnames, nodenames, cellnames)

	' ==========	TO CORRECT !
	Dim npins As Long, nnodes As Long, ncells As Long
'	npins = NMAX
'	ncells = NMAX
'	nnodes = NMAX

	npins = UBound(pinnames)+1
	ncells = UBound(cellnames)+1
	nnodes = 0
	' ==========


	Open outfilen For Output As #88

	' Main part
	' =========
	tmpstr = "[MAIN]" + vbCrLf + "Version = " + PORTUNUS_Version + vbCrLf

	tmpstr = tmpstr + "NumberOfPins = " + CStr(npins) + vbCrLf
	tmpstr = tmpstr + "NumberOfNodes = " + CStr(nnodes) + vbCrLf
	tmpstr = tmpstr + "NumberOfCells = " + CStr(ncells) + vbCrLf
	tmpstr = tmpstr + "Frequencies = "
	nfrequencies = UBound(F) + 1
	For iii = 0 To nfrequencies - 1
		tmpstr = tmpstr + F(iii) + "; "
	Next
	tmpstr = tmpstr + vbCrLf

	Print #88, tmpstr

	' Pins
	' =========
	Print #88, "[Pins]"
	For iii = 0 To UBound(pinnames, 1)
		Print #88, "NamePin_" + Cstr(iii+1) + " = " + pinnames(iii)	'no_Portunus_forbidden_chars(pinnames(iii))
	Next
	Print #88, ""


	' Nodes
	' =========


	' Cells
	' =========
	For iii = 0 To UBound(rowportnames,1)
		jjj = iii + 1
			tmpstr = "[Cell_" + Cstr(jjj) + "]" + vbCrLf
			tmpstr = tmpstr + "Name = " + cellnames(iii) + vbCrLf	'no_Portunus_forbidden_chars(cellnames(iii)) + vbCrLf
			kkk = InStr(rowportnames(iii), "to")
			lll = InStr(rowportnames(iii), ",")
			tmpstr = tmpstr + "Connection_1 = " + Left(rowportnames(iii), kkk-2) + vbCrLf
			tmpstr = tmpstr + "Connection_2 = " + Mid(rowportnames(iii), kkk+3, lll-kkk-3) + vbCrLf
			Print #88, tmpstr
	Next

	' R
	' =========
	Print #88, "[R]"
	For iii = 0 To UBound(rowportnames, 1)
		For jjj = 0 To iii
			tmpstr = "R/Cell_"+ Cstr(iii+1) + "/Cell_" + CStr(jjj+1) +" = "
			If UBound(F,1) = 0 Then		' Just one frequency
				tmpstr = tmpstr + Cstr(R(0, iii, jjj))
				tmpstr = tmpstr + " |"
			Else
				' First write the frequency-independent value
				tmpstr = tmpstr + Cstr(R(0, iii, jjj))
				tmpstr = tmpstr + " | "
				' Then write the frequency-dependent values
				For kkk = 0 To UBound(F, 1)	' Frequencies
					tmpstr = tmpstr + Cstr(R(kkk, iii, jjj)) + "; "
				Next kkk
			End If
			Print #88, tmpstr
		Next jjj
	Next	iii

	' L
	' =========
	Print #88, vbCrLf + "[L]"
	For iii = 0 To UBound(rowportnames, 1)
		For jjj = 0 To iii
			tmpstr = "L/Cell_"+ Cstr(iii+1) + "/Cell_" + CStr(jjj+1) +" = "
			If UBound(F,1) = 0 Then		' Just one frequency
				tmpstr = tmpstr + Cstr(L(0, iii, jjj))
				tmpstr = tmpstr + " |"
			Else
				' First write the frequency-independent value
				tmpstr = tmpstr + Cstr(L(0, iii, jjj))
				tmpstr = tmpstr + " | "
				For kkk = 0 To UBound(F, 1)	' Frequencies
					tmpstr = tmpstr + Cstr(L(kkk, iii, jjj)) + "; "
				Next kkk
			End If
			Print #88, tmpstr
		Next jjj
	Next	iii

	Close #88

End Sub

Function no_Portunus_forbidden_chars(str1 As String) As String
	' Forbidden Portunus characters as of 25.9.2012:    - . ; ,
	Dim str2 As String
	Dim forbidden As Variant
	Dim tmp As Long, iii As Long

	forbidden=Array(" ", ".", ",", ";", "-", ">", "<", "=", "&", "|", """", "\", ")" , "(", "[", "]")

	tmp = UBound(forbidden)

	str2 = Replace(str1, forbidden(0), "_", , -1)	' Replace all "-" by "_"

	For iii = 0 To tmp
		str2 = Replace(str2, forbidden(iii), "_", , -1)	' Replace all "-" by "_"
	Next

	' Here, we should check that there are no too long sequences of "_"


	no_Portunus_forbidden_chars = str2
End Function

