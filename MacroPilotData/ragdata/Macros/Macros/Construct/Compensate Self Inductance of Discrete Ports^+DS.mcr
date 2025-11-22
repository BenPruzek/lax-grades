' *Construct / Discrete Ports / Compute SelfInductance of Dis.Ports
' !!! Do not change the line above !!
' ================================================================================================
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 01-Sep-2021 iqa: Replaced deprecated 'ExternalPort.Number'-commands with the 'ExternalPort.Name'-command
' 02-Dec-2020 fcu: Remove the labels from the block's portnames (CST-64762).
' 13-Oct-2020 tsz: Improve robustness, if external ports existed before calling the macro.
' 11-Oct-2020 fhr: set capacitance/inductance values to ZERO in case of distributed face ports in !T (PBA), dialog popup improved
' 05-Oct-2020 fhr: added sufficient decimal places for C and L outputs to Schematic
' 22-Jun-2020 fcu: replace the link command with the net command
' 03-Jan-2020 ube:  add warning for distributed ports with FIT Solver (new default since v2019)
' 15-Nov-2019 fhr: Ports can now have labels
' 15-Oct-2015 gba:  if tuning caps are switched on, do not first create external ports and delete them again (improves routing and speed)
' 15-Oct-2015 gba:  do not assume that schematic block is named "MWSSCHEM1"
' 15-Oct-2015 gba:  speed up routing by calling port.position before port.create
' 15-Oct-2015 gba:  small speed up by calling block.setdoubleproperty before block.create
' 28-May-2014 gba:  calculate total number of block pins correctly
' 08-Jun-2012 gba:  speed up routing by calling block.rotate before block.create
' 13-May-2011 fhi:  layout improvements of the blocks
' 05-May-2011 fhi:  added multipin/multimode capabilty for waveguide-port pins in DS; indices of mutual port-coupling corrected
' 04-Apr-2011 fsr:  port indexes are now found using "StartPortNumberIteration/GetNextPortNumber"
' 24-Jan-2011 fhi:  compatible search String "Ports" for versions 2011, 2010 and earlier.
' 14-Jan-2011 fhi:  adjusted portlength equal to portlength if dis-port length is less than 1 mesh cell...
' 01-Oct-2010 fhi:  Portlengths of approx. zero ( < 1e-7 ) assigned with the correct capacitance
' 12-Jul-2010 fhi:  check angle and size of two parent ports, option dCaps instead of Ports
' 09-Jul-2010 fhi:  added neg. mutual couplings
' 25-May-2010 fhi:  Distinguish between !T and !F for Hex-Mesh: Hexmesh !F does not consider face-ports.
' 24-Feb-2010 fhi:  inactivated WG Port connections for the moment (multi-pin, multi-Mode not supported)
' 20-Jan-2010 fhi:  added Face-port capability for L and C computation, tet and hex mesh, (tet-face-port new formulation in 2010)
' 21-Oct-2009 fhi:  Port.GetFacePortSize( pnr,   b_dir1,   l_dir2 )
' 05-Mar-2009 fhi:  added links to Caps for differential/nondifferential Links
' 10-Sep-2007 fhi:  Caps on/off, mesh symmetry-planes considered , mesh-search based on min/max mesh-step
' 20-Aug-2007 fhi:  adapted to DE2008
' 03-Apr-2006 fhi:  Compute -L for arbitrarily orientated ports and adds them into DesignStudio
' 05-Apr-2006 fhi:  De2006: SP5 corrections: PR Tracker 2593, 2608 - 2610, 2614
' 07-Apr-2006 fhi:  Differential_ports and Blocks implemented, Dialog added, autom. deletion of existing ports/blocks
' 30-Oct-2006 fhi:  Correction of z_lower=Mesh.getzpos(Z_istop) in subroutine "search_for_zlower_z_upper"
' 31-Oct-2006 fhi:  Added equivalent capacitance of ports (port-wires rep. by two long cylinders)
' 11-Feb-2006 fhi:  Initial version


Option Explicit

Type Port_Info
	nr As String
	label As String
End Type


Public ListArray_mode() As String	'contains the nr of modes for a particular port-nr
Public ListPortArray() As String
Public PortArrayLength As Integer
'
Public Discrete_Port_List() As Port_Info, Total_Nr_of_Discrete_Ports As Integer, Total_Nr_of_WG_Ports As Integer, N_max_searches As Integer
Public SchematicBlockName As String

Private Sub GetSchematicBlockName()
	Dim nBlocks As Long
	Dim nIndex As Long
	Dim BlockNameArray() As String
	Dim BlockName As String
	Dim BlockType As String

	SchematicBlockName = "MWSSCHEM1"
	nBlocks = Block.StartBlockNameIteration
	ReDim BlockNameArray(nBlocks)

	For nIndex=0 To nBlocks-1
		BlockNameArray(nIndex) = Block.GetNextBlockName
	Next nIndex

	For nIndex=0 To nBlocks-1
		With Block
			.Reset
			.Name BlockNameArray(nIndex)
			BlockType = .GetTypeShortName
		End With
		If Right$(BlockType, 5) = "SCHEM" Then
			SchematicBlockName = BlockNameArray(nIndex)
		End If
	Next nIndex
End Sub

Private Sub ConnectComponents(Comps() As Variant)
' -------------------------------------------------------------------------------------------------
' ConnectComponents: This function connects the components.
' -------------------------------------------------------------------------------------------------
	Dim numberOfComponents As Integer
	numberOfComponents = UBound(Comps)-LBound(Comps)+1

	Dim subnet() As Variant
	ReDim subnet(numberOfComponents - 1,2)

	Dim subnetNames() As String
	ReDim subnetNames(numberOfComponents - 1)

	'Get current connectiontopology before using net command
	Net.Reset
	
	Dim Comp As Variant
	Dim i As Integer
	i = 0
	For Each Comp In Comps
		subnet(i,0) = Comp(0)
		subnet(i,1) = Comp(1)

		If (Comp(0) = "B" Or Comp(0) = "BLOCK") Then
			Block.Reset
			Block.Name Comp(1)
			subnet(i,0) = "BLOCK"
			subnet(i,2) = Block.GetPortIndex(Comp(2))
		Else
			subnet(i,2) = CInt(Comp(2))
		End If
		subnetNames(i) = Net.GetNetName(Array(subnet(i,0),subnet(i,1),subnet(i,2)))
		i = i + 1
	Next Comp

	Dim Dummy2() As Variant
	Dim netname As String
	Dim newSize As Long

	For Each netname In subnetNames
		If netname <> "" Then
			Dim funky() As Variant
			funky = Net.GetComponentPorts(netname, -1)
			subnet = AppendComps(subnet, funky)
		End If
	Next

	With Net
		.Reset
		.AddComponentPorts("", subnet, False)
		.Apply
	End With
End Sub

Private Function AppendComps(Comps() As Variant, CompsToAppend() As Variant) As Variant()
	Dim Dummy() As Variant
	Dummy = Comps

	Dim i,j As Integer

	Dim found As Boolean
	For i = LBound(CompsToAppend) To UBound(CompsToAppend)
		found = False
		For j = LBound(Comps) To UBound(Comps)
			If (CompsToAppend(i,0) = Comps(j,0) And CompsToAppend(i,1) = Comps(j,1) And CompsToAppend(i,2) = Comps(j,2)) Then
				found = True
			End If
		Next j

		If Not found Then
			Dummy = PushBack(Dummy, Array(CompsToAppend(i,0),CompsToAppend(i,1),CompsToAppend(i,2)))
		End If
	Next i

	AppendComps = Dummy
End Function

Private Function PushBack(Comps() As Variant, CompToAppend As Variant) As Variant()
	Dim Dummy() As Variant
	ReDim Dummy(UBound(Comps) - LBound(Comps) + 1, 2)

	Dim i,j As Long
	For i = LBound(Comps) To UBound(Comps)
		For j = 0 To 2
			Dummy(i,j) = Comps(i,j)
		Next j
	Next i

	For j = 0 To 2
		Dummy(UBound(Dummy) - LBound(Dummy), j) = CompToAppend(j)
	Next j

	PushBack = Dummy
End Function
Sub Main

Dim long_ref_dir As Long, long_ref_index As Long, port_nr As Integer, delta As Double
Dim x As Double, y As Double, z As Double, plength As Double, X_start_index As Long, X_istop As Long
Dim x_upper As Double, x_lower As Double, wire_radius As Double,  loop_index As Integer
Dim Y_start_index As Long, Y_istop As Long, y_upper As Double, y_lower As Double, L_i As Double
Dim L_a As Double, glength As Double, iii As Integer, port_name As String, index_p As Long, index_q As Long
Dim port_index As Long, z_upper As Double, z_lower As Double, output_string As String
Dim long_ref_index2 As Long, source_start_point As Double, source_end_point As Double
Dim source_length As Double, port_onearmlength As Double, C_port As Double, act_portnr As Integer
Dim face_port_dir1 As Double, face_port_dir2 As Double, face_port As Boolean, faceport_area As Double
Dim K_array() As Double, L_array () As Double, M As Double, k As Double, composed_port_number As String
Dim x0_p As Double, y0_p As Double, z0_p As Double, glength1 As Double,glength2 As Double
Dim x1_p As Double, y1_p As Double, z1_p As Double, port_distance As Double, L_p As Double, L_q As Double
Dim x0_q As Double, y0_q As Double, z0_q As Double, center_pointx1 As Double, center_pointx2 As Double
Dim x1_q As Double, y1_q As Double, z1_q As Double, center_pointy1 As Double, center_pointy2 As Double
Dim center_pointz1 As Double, center_pointz2 As Double, portlength_mean As Double, composed_block_number As String
Dim tuningC_flag As Boolean, slash_number As Integer, cst_iiia As Integer, cst_nr_of_modes As Integer, port_nr_offset As Integer, port_1000 As Integer
Dim Source As Variant

'DS
Dim mws_block_port_position_x As  Long, mws_block_port_position_y As  Long
Dim mws_block_center_position_x As Long, mws_block_center_position_y As  Long, make_abs As  String
Dim x_offset As Integer, y_offset As Integer, size_offset As Integer
Dim blockname_variable As String, rel_x_pos As Double, rel_y_pos As Double, rot_angle As Double, P_index As Integer
Dim differential_flag As Boolean, orientation As String, U_index As Integer , D_index As Integer , L_index As Integer , R_index As Integer

	Begin Dialog UserDialog 1050,203,"Compensate L and C for Discrete/Face Ports",.Dialogfunc ' %GRID:10,7,1,1
		CheckBox 170,98,170,14,"Include Capacitances",.add_caps
		CheckBox 20,98,140,14,"Create Link-Lines",.CreateLinks
		CheckBox 370,98,220,14,"Differential Ports and Blocks",.Diff_ports_blocks
		CheckBox 600,98,210,14,"Tuning Caps instead of Ports",.tuningC
		CheckBox 820,98,210,14,"Consider mutual Portcoupling",.port_coupling
		CheckBox 160,142,140,14,"Create Report",.report

		GroupBox 20,126,130,42,"Solver-Type",.Group1
		OptionGroup .solver_type
			OptionButton 30,140,50,21,"!T",.OptionButton1
			OptionButton 80,140,50,21,"!F",.OptionButton2
		OKButton 20,175,100,21
		CancelButton 130,175,90,21
		Picture 10,7,1020,84,GetInstallPath + "\Library\Macros\Construct\Discrete Ports\port_deemb.bmp",0,.Picture1
	End Dialog
 Dim dlg As UserDialog

 ' set dialog defaults
 dlg.Createlinks = 1 : dlg.diff_ports_blocks = 0 : dlg.report = 0 : dlg.add_caps = 1' 0 = no 1 = yes
 dlg.solver_type = 0  ' 0 = !T
 dlg.tuningC= 0
 If (Dialog(dlg) = 0) Then Exit All		'do the dialog

	If dlg.solver_type = 0 Then  ' T-FIT-Solver is used
		If	Solver.GetDiscreteItemEdgeUpdate = "Distributed" Or Solver.GetDiscreteItemFaceUpdate = "Distributed" Then
			If (MsgBox "Note: T-Solver is calculated with ""Distributed"" Discrete Ports, whereas Macro is designed for ""Gap"" Discrete Ports." + _
					vbCrLf + vbCrLf + "Please check setting in ""T-Solver -> Specials -> Solver -> Discrete Port Settings""" + vbCrLf + vbCrLf + _
					"No Compensation is required for distributed ports. When continuing, all L, C values will be set to zero." + vbCrLf + vbCrLf + _
					"Continue anyway?",vbExclamation+vbYesNo,"Compensate Self Inductance of Discrete Ports") = vbNo Then
				Exit All
			End If
		End If
	End If

 If dlg.diff_ports_blocks = 1 Then
  differential_flag= True
 Else
  differential_flag=False
 End If
 If dlg.tuningC = 1 Then
  tuningC_flag= True
 Else
  tuningC_flag=False
 End If

GetSchematicBlockName()
loop_index =0
U_index=0 : D_index=0 : L_index=0 : R_index=0 : P_index =0
 size_offset = 200	'position offset for drawing the elements
 port_nr_offset =10 :  port_1000 =1000	' Multiports are composed as 1000 + port_nr*10 +mode_nr

 FillPortNameArray		'all port numbers are stored in  portnamearray()
 Dim i As Integer

 For i= 0 To Port.StartPortNumberIteration-1
 	' get the modes of the first available port as default in the popup
 	setup_mode_list (ListArray_mode(), Total_Nr_of_Discrete_Ports  , Total_Nr_of_WG_Ports)

 '	MsgBox cstr(i)+": "+Discrete_Port_List(i)+" Nr Modes:  "+cstr(ListArray_mode(i))
 Next
'MsgBox "dis ports="+ cstr (Total_Nr_of_Discrete_Ports)+ "wg="+  cstr( Total_Nr_of_WG_Ports)

 'add text to a report-file
 output_string =  "Total Number of Discrete Ports : " + CStr(Total_Nr_of_Discrete_Ports)+ vbCrLf
 output_string =  output_string +"Total Number of Waveguide Ports : " + CStr(Total_Nr_of_WG_Ports)+ vbCrLf


	' redim the arrays for ind L and couplings K:
	ReDim K_array (0 To Total_Nr_of_Discrete_Ports-1, 0 To Total_Nr_of_Discrete_Ports-1)
	ReDim L_array (0 To Total_Nr_of_Discrete_Ports-1)


 'Loop over all Ports --------------------------------------------------------------------------

 For index_p = 0 To Port.StartPortNumberIteration-1 'Total_Nr_of_Discrete_Ports-1

  If CInt(Discrete_Port_List(index_p).nr) > 0 Then		'consider only discrete ports for deembedding- computation (+ sign)

	Port.GetFacePortSize(Discrete_Port_List(index_p).nr,  face_port_dir1,   face_port_dir2 )	'' width, length of face port
	faceport_area = Port.GetFacePortArea(Discrete_Port_List(index_p).nr )	'if no faceport : area = 0


	If faceport_area > 0 Then 	' FACE-Port

	'check if dimensions are zero
	  If face_port_dir1 = 0 Then	'width
		If face_port_dir2 = 0 Then
			face_port_dir2=DiscretePort.getlength (Discrete_Port_List(index_p).nr)	'take the length via dis-port vba command
			If face_port_dir2 = 0 Then ' portlength = 0 !
				MsgBox "Face-Port inconsistencies at Port Nr. "+ cstr(Discrete_Port_List(index_p).nr) + " , check length of port !"
				Exit All
			End If
			face_port_dir1 = faceport_area / face_port_dir2	'assume rectangular shaped faceport
		Else
			face_port_dir1 = faceport_area / face_port_dir2	' assuming rect faceports....A = w*h
		End If
	  ElseIf face_port_dir2 = 0 Then	'portlength = 0!
		If face_port_dir1 = 0 Then
			MsgBox "Face-Port inconsistencies at Port Nr. "+ cstr(Discrete_Port_List(index_p).nr) + " !"
			Exit All
		Else
			face_port_dir2 = faceport_area / face_port_dir1	' assuming rect faceports....A = w*h
		End If
	  End If

		face_port = True		' face port
	Else
		face_port = False		' ordinary discrete port
	End If


	'If !F Solver and Hex-Mesh is used , then switch to dis-port !!!!!!!!!!!!

	If dlg.solver_type = 1 And Mesh.GetMeshType = "PBA" Then ' solver_type= 1 = !F
		face_port = False ' discrete port
	End If




	Select Case Mesh.GetMeshType

	Case "PBA"	'hexa-mesh

	If face_port Then
	 output_string = output_string+ vbCrLf+ " Face Port "+CStr(Discrete_Port_List(index_p).nr)+ " (Hexa-Mesh)"+vbCrLf
	Else
		output_string = output_string+ vbCrLf+ " Edge Port "+CStr(Discrete_Port_List(index_p).nr)+ " (Hexa-Mesh)"+vbCrLf
	End If
	 output_string = output_string+ "Port Impedance = "+ CStr(Port.GetLineImpedance (Discrete_Port_List(index_p).nr,1) )+" Ohm  "
	 If face_port Then
	  plength = face_port_dir2
	 Else
	  plength=DiscretePort.getlength (Discrete_Port_List(index_p).nr)
	 End If
	 glength=  DiscretePort.getgridlength ( Discrete_Port_List(index_p).nr)
	 output_string = output_string+ "Gridlength= "+CStr( glength  )+ "  PortLength= "+CStr(plength)+ _
	               " Deviation= "+ CStr( Format( (100*(glength-plength)/plength),"###0.##") )+ " % "+vbCrLf
	 DiscretePort.GetElementDirIndex (Discrete_Port_List(index_p).nr,long_ref_dir, long_ref_index)
	 x= Mesh.getxpos(long_ref_index)
     y= Mesh.getypos(long_ref_index)
	 z= Mesh.getzpos(long_ref_index)
	 output_string = output_string+ "Source Center-location "+ CStr(x)+ " / "+CStr(y) + " / "+CStr(z)

	 '************ search -density for gridpoints
	 'delta = plength/100		'search increment to find next mesh index, 100 = 1% of Port-length
 	 delta =  Mesh.GetMinimumEdgeLength/2	'new method
	 N_max_searches = CInt(2*Mesh.GetMaximumEdgeLength/Mesh.GetMinimumEdgeLength)
	 '************ end of search-density

     If long_ref_dir = 2 Then '(orientation in z)
	  search_for_xlower_x_upper (delta, long_ref_index, x_upper, x_lower)
	  search_for_ylower_y_upper (delta, long_ref_index, y_upper, y_lower)
	  output_string = output_string+ "  Orientation = z"+vbCrLf
	  output_string = output_string+ "LocationMeshlines X: "+ CStr(x_lower) + "/ center "+ CStr(Mesh.getxpos(long_ref_index)) + " /"+ CStr(x_upper)+vbCrLf
	  output_string = output_string+ "LocationMeshlines Y: "+  CStr(y_lower) + "/ center "+ CStr(Mesh.getypos(long_ref_index)) + "/ "+ CStr(y_upper)

	 'effective radius
	  If face_port Then
		wire_radius = face_port_dir1/2	' half of width of faceport
	  Else
		wire_radius = (Abs(x_upper-x_lower)+ Abs(y_upper-y_lower))/(4*Exp(1)^2)	'mean value radius
	  End If
	  output_string = output_string+ vbCrLf+"Wire radius = "+CStr(wire_radius)

	 'Formula for L of a straight wire:
	 ' inner L
	  L_i = Units.GetGeometryUnitToSI()*plength*4*pi*1e-7/(8*pi)
	 'L_a = Units.GetGeometryUnitToSI()*plength*2*1e-7*(Log(2*plength/wire_radius)-1)
	  L_a = compute_L_a (wire_radius, plength)
	  output_string = output_string+ "   L_i= "+CStr (L_i*1e9)+"   L_a = "+CStr (L_a*1e9)+ " nH" + vbCrLf

	  '------ capacitance of discrete port z ------
	  If dlg.add_caps = 1 Then
	   source_start_point= Mesh.getzpos(long_ref_index)
	   DiscretePort.GetElement2ndIndex (Discrete_Port_List(index_p).nr, long_ref_index2 )
	   source_end_point  = Mesh.getzpos(long_ref_index2)
	   source_length = Abs(source_end_point-source_start_point)

	  If plength < source_length Then		' case when the port-length is smaller than one meshcell!!!
	    plength = source_length
	   End If

	   port_onearmlength = (plength-source_length)/2

	   C_port =Units.GetGeometryUnitToSI()*compute_C(2*wire_radius,port_onearmlength,source_length/2)
	   If face_port Then
	    C_port = C_port/correction_factor_face_ports(wire_radius,port_onearmlength,source_length/2)
	   Else
		C_port = C_port/correction_factor_dis_ports(wire_radius,port_onearmlength,source_length/2)
  	   End If

  	   ' !T Solver settings are on "distributed" : no capacitance occurs
  	 If dlg.solver_type = 0 Then
	 If	Solver.GetDiscreteItemEdgeUpdate = "Distributed" Or Solver.GetDiscreteItemFaceUpdate = "Distributed" Then
			C_port = 0.
			L_i = 0.  '....
			L_a = 0.
	 End If
	 End If

	   output_string = output_string+ " Port_Capacitance = "+CStr (C_port*1e12)+  " pF" + " Gap-length: " + cstr(source_length) +vbCrLf
	  End If
	'---------- end cap z ----------------------

    End If ' orientation z.....

    If long_ref_dir = 0 Then '(orientation in x)
	 search_for_zlower_z_upper (delta, long_ref_index, z_upper, z_lower)
	 search_for_ylower_y_upper (delta, long_ref_index, y_upper, y_lower)
	 output_string = output_string+ "  Orientation = x"+vbCrLf
	 output_string = output_string+ "LocationMeshlines Z: "+CStr(z_lower) + "/ center "+ CStr(Mesh.getzpos(long_ref_index)) + " /"+ CStr(z_upper)+vbCrLf
	 output_string = output_string+ "LocationMeshlines Y: "+CStr(y_lower) + "/  center  "+ CStr(Mesh.getypos(long_ref_index)) + "/ "+ CStr(y_upper)

	'effective radius
	 If face_port Then
		wire_radius = face_port_dir1/2	' half of width of faceport
	 Else
	    wire_radius = (Abs(z_upper-z_lower)+ Abs(y_upper-y_lower))/(4*Exp(1)^2)	'mean value radius
	 End If
	 output_string = output_string+ vbCrLf+"Wire radius = "+CStr(wire_radius)

	 'Formula for L of a straight wire:
	 ' inner L
	 L_i = Units.GetGeometryUnitToSI()*plength*4*pi*1e-7/(8*pi)
	 'L_a = Units.GetGeometryUnitToSI()*plength*2*1e-7*(Log(2*plength/wire_radius)-1)
	 L_a = compute_L_a (wire_radius, plength)
	 output_string = output_string+    "   L_i= "+CStr (L_i*1e9)+"    L_a = "+CStr (L_a*1e9)+ " nH"+vbCrLf

  	 '------ capacitance of discrete port x ------
	 If dlg.add_caps = 1 Then
	  source_start_point= Mesh.getxpos(long_ref_index)
	  DiscretePort.GetElement2ndIndex (Discrete_Port_List(index_p).nr, long_ref_index2 )
	  source_end_point  = Mesh.getxpos(long_ref_index2)
	  source_length = Abs(source_end_point-source_start_point)

  	If plength < source_length Then		' case when the port-length is smaller than one meshcell!!!
	    plength = source_length
	   End If

	  port_onearmlength = (plength-source_length)/2

	  C_port =Units.GetGeometryUnitToSI()*compute_C(2*wire_radius,port_onearmlength,source_length/2)
	  If face_port Then
		C_port = C_port/correction_factor_face_ports(wire_radius,port_onearmlength,source_length/2)
	  Else
		C_port = C_port/correction_factor_dis_ports(wire_radius,port_onearmlength,source_length/2)
	  End If

	  ' !T Solver settings are on "distributed" : no capacitance occurs
	  If dlg.solver_type = 0 Then
	 If	Solver.GetDiscreteItemEdgeUpdate = "Distributed" Or Solver.GetDiscreteItemFaceUpdate = "Distributed" Then
			C_port = 0.
			L_i = 0
			L_a = 0
	 End If
	End If

	  output_string = output_string+ " Port_Capacitance = "+CStr (C_port*1e12)+  " pF" + " Gap-length: " + cstr(source_length)+ vbCrLf
	 End If
	'---------- end cap x ----------------------

    End If ' orientation x.....

    If long_ref_dir = 1 Then '(orientation in y)
	 search_for_zlower_z_upper (delta, long_ref_index, z_upper, z_lower)
	 search_for_xlower_x_upper (delta, long_ref_index, x_upper, x_lower)
	 output_string = output_string+ "  Orientation = y"+vbCrLf
	 output_string = output_string+ "LocationMeshlines Z: "+CStr(z_lower) + "/  center "+ CStr(Mesh.getzpos(long_ref_index)) + " /"+ CStr(z_upper)+vbCrLf
	 output_string = output_string+ "LocationMeshlines X: "+CStr(x_lower) + "/   center "+ CStr(Mesh.getxpos(long_ref_index)) + "/ "+ CStr(x_upper)

	 'effective radius
	 If face_port Then
		wire_radius = face_port_dir1/2	' half of width of faceport
	 Else
	  wire_radius = (Abs(z_upper-z_lower)+ Abs(x_upper-x_lower))/(4*Exp(1)^2)	'mean value radius
	 End If
	 output_string = output_string+ vbCrLf+"Wire radius = "+CStr(wire_radius)

	 'Formula for L of a straight wire:
	 ' inner L
 	 L_i = Units.GetGeometryUnitToSI()*plength*4*pi*1e-7/(8*pi)
	 L_a = compute_L_a (wire_radius, plength)
	 output_string = output_string+ "   L_i= "+CStr (L_i*1e9)+"    L_a = "+CStr (L_a*1e9)+ " nH"+vbCrLf


    '------ capacitance of discrete port y ------
    If dlg.add_caps = 1 Then
	 source_start_point= Mesh.getypos(long_ref_index)
	 DiscretePort.GetElement2ndIndex (Discrete_Port_List(index_p).nr, long_ref_index2 )
	 source_end_point  = Mesh.getypos(long_ref_index2)
	 source_length = Abs(source_end_point-source_start_point)

	  If plength < source_length Then		' case when the port-length is smaller than one meshcell!!!
	    plength = source_length
	   End If

	 port_onearmlength = (plength-source_length)/2

	 C_port =Units.GetGeometryUnitToSI()*compute_C(2*wire_radius,port_onearmlength,source_length/2)
	 If face_port Then
		C_port = C_port/correction_factor_face_ports(wire_radius,port_onearmlength,source_length/2)
	 Else
		C_port = C_port/correction_factor_dis_ports(wire_radius,port_onearmlength,source_length/2)
	 End If

	 ' !T Solver settings are on "distributed" : no capacitance occurs
	 If dlg.solver_type = 0 Then
	 If	Solver.GetDiscreteItemEdgeUpdate = "Distributed" Or Solver.GetDiscreteItemFaceUpdate = "Distributed" Then
			C_port = 0.
			L_i = 0
			L_a = 0
	 End If
	 End If

	output_string = output_string+ " Port_Capacitance = "+CStr (C_port*1e12)+  " pF" + " Gap-length: " + cstr(source_length) +vbCrLf
	End If
	'---------- end cap y ----------------------

	End If ' orientation y.....


	Case "Tetrahedral" '  "tetmesh"
		If face_port Then
		output_string = output_string+ vbCrLf+ "Face-Port "+CStr(Discrete_Port_List(index_p).nr)+ " (Tetra-Mesh)" + vbCrLf
		Else
		output_string = output_string+ vbCrLf+ "Edge-Port "+CStr(Discrete_Port_List(index_p).nr)+ " (Tetra-Mesh)"+ vbCrLf
		End If

		output_string = output_string+ "Port Impedance = "+ CStr(Port.GetLineImpedance (Discrete_Port_List(index_p).nr,1) )+" Ohm  "
		If face_port Then
			plength = face_port_dir2 'length of port
		Else
			MsgBox "Discrete Ports using Tetra-Mesh are not supported at port # "+ CStr(Discrete_Port_List(index_p).nr)+"."+vbCrLf+ _
			"Please modify it to type ""Face-Port""",,"Warning:"
			'Exit All
		End If
		If face_port Then
		 'compute C_wires horizontally orientated, C effectively ZERO;
		 C_port=0
		 'set to a dummy-factor of 1e6 in length and radius for testing purposes
		 'C_port =Units.GetGeometryUnitToSI()*compute_C_tetmesh_new(face_port_dir1*0.000001,face_port_dir1/100000,plength/2) 'diameter=1%L, Length, heigth above ground
		Else
		 C_port =0
		End If
		output_string = output_string+ " Face-Port_Capacitance = "+CStr (C_port*1e12)+  " pF" + " Wire-length: " + cstr(face_port_dir2) +vbCrLf

		'Formula L of a straight wire:
		'
		If face_port Then
			wire_radius = face_port_dir1/2 	'assumed: radius = half of port-width
			L_i = Units.GetGeometryUnitToSI()*plength*4*pi*1e-7/(8*pi)
			L_a = compute_L_a (wire_radius, plength)
			L_i=0
			L_a=0
		Else
			L_i = 0 : L_a=0
		End If
		output_string = output_string+ "   L_i= "+CStr (L_i*1e9)+"    L_a = "+CStr (L_a*1e9)+ " nH"+vbCrLf

	Case Else
		MsgBox "Meshtype unkown,  not supported"
		Exit All
	End Select

  End If ' only discrete ports

'*==================================================================================================*
'*																									*
'* 					CST DesignStudio Schematics														*
'*																									*
'*==================================================================================================*
  'Start of CST-DesignStudio Implementation: draw neg. Ls + all external Ports + connections
  '--------------------------------------------------------------------------------------

  act_portnr = cInt(Discrete_Port_List(index_p).nr)


  ' get the port-position xy of the MWS Block
  With Block
 	.name SchematicBlockName
	.SetDifferentialPorts differential_flag
	mws_block_center_position_x = .GetPositionX
	mws_block_center_position_y = .GetPositionY
  End With
'----------------------------------------------
  mws_block_port_position_x  = getPortPos (Port.StartPortNumberIteration-1,index_p,0,"x")
  mws_block_port_position_y  = getPortPos (Port.StartPortNumberIteration-1,index_p,0,"y")
  '----------------------

 getPortOrientation (mws_block_center_position_x, mws_block_center_position_y, mws_block_port_position_x, mws_block_port_position_y, size_offset  ,   _
			 x_offset    ,y_offset  , rot_angle  , orientation   )


	Select Case orientation
    	Case "L"
    			L_index=L_index+1
    	Case "R"
				R_index=R_index+1
    	Case  "D"
				D_index=D_index+1
		Case "U"
				U_index=U_index+1
    End Select


  'Draw the neg. inductances (Block)
  If CInt(Discrete_Port_List(index_p).nr) > 0 Then 'only for discr. ports
   blockname_variable = "Port_L_"+CStr(Discrete_Port_List(index_p).nr)
'modify format (fhr3)
   If Units.GetInductanceSIToUnit() >= 1e6 Then
	DS.storeparameter (blockname_variable ,Format(-1.*(L_i+L_a)*Units.GetInductanceSIToUnit(),"###0.######"))', blockname_variable +" L in "+  CStr(Units.GetUnit("Inductance")))
   Else
 	DS.storeparameter (blockname_variable ,Format(-1.*(L_i+L_a)*Units.GetInductanceSIToUnit(),"0.000E+00"))', blockname_variable +" L in "+  CStr(Units.GetUnit("Inductance")))
   End If
   With Block ' check existance
		.name "Port__L_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name "Port__L_"+CStr(Discrete_Port_List(index_p).nr)

    'MsgBox "P("+cstr(index_p)+") ="+cstr(act_portnr)

    Select Case orientation
    	Case "L"
    		 .position (mws_block_port_position_x +x_offset+ L_index*x_offset/2 , mws_block_port_position_y+y_offset )	'<*
    	Case "R"
    		 .position (mws_block_port_position_x +x_offset+ R_index*x_offset/2 , mws_block_port_position_y+y_offset )	'<**
    	Case  "D"
			.position  (mws_block_port_position_x +x_offset ,                     mws_block_port_position_y + y_offset+ D_index*y_offset/2)  '<***
		Case "U"
			 .position (mws_block_port_position_x +x_offset ,                     mws_block_port_position_y + y_offset+ U_index*y_offset/2) '<****
    End Select
    .Rotate (rot_angle)
    .SetDoubleProperty ("Inductance",  blockname_variable  )
    .create
   End With
  End If

  'Draw the neg. capacitances (Block)
  If dlg.add_caps = 1 Then
  If CInt(Discrete_Port_List(index_p).nr) > 0 Then 'only for discr. ports
   blockname_variable = "Port_C_"+CStr(Discrete_Port_List(index_p).nr)
   If Units.GetCapacitanceSIToUnit() >= 1e6 Then
   	DS.storeparameter (blockname_variable ,Format(-1.*C_port*Units.GetCapacitanceSIToUnit(),"###0.######"))', blockname_variable +" C in "+  CStr(Units.GetUnit("Capacitance")))
   Else
	DS.storeparameter (blockname_variable ,Format(-1.*C_port*Units.GetCapacitanceSIToUnit(),"0.000E+00"))', blockname_variable +" C in "+  CStr(Units.GetUnit("Capacitance")))
   End If

   With Block ' check existance
		.name "Port__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name "Port__C_"+CStr(Discrete_Port_List(index_p).nr)

    Select Case orientation
    	Case "L"
    		If Not tuningC_flag Then
    	  	 .position (mws_block_port_position_x +1.5*x_offset+ L_index*x_offset/2 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.25 )) '<*
			Else
			 .position (mws_block_port_position_x +1.5*x_offset+ L_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.25 )) '<*
			End If
    		.Rotate (rot_angle-90)
    	Case "R"
    		If Not tuningC_flag Then
    	  	 .position (mws_block_port_position_x +1.5*x_offset+ R_index*x_offset/2 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.25 )) '<**
			Else
			 .position (mws_block_port_position_x +1.5*x_offset+ R_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.25 )) '<**
			End If
    		.Rotate (rot_angle-90)
    	Case  "D"
			If Not tuningC_flag Then
				 .position (mws_block_port_position_x +1.5*x_offset -Int(D_index*size_offset/1.25) , mws_block_port_position_y+1.5*y_offset +D_index*y_offset/2. ) '<***
	  		Else
	  			.position (mws_block_port_position_x +1.5*x_offset -Int(D_index*size_offset/1.25) , mws_block_port_position_y+1.5*y_offset +D_index*y_offset/1. ) '<***
	  		End If
   			.Rotate (rot_angle-90)
		Case "U"
		If Not tuningC_flag Then
				 .position (mws_block_port_position_x +1.5*x_offset -Int(U_index*size_offset/1.25) , mws_block_port_position_y+1.5*y_offset +U_index*y_offset/2. ) '<****
	  		Else
	  			.position (mws_block_port_position_x +1.5*x_offset -Int(U_index*size_offset/1.25) , mws_block_port_position_y+1.5*y_offset +U_index*y_offset/1. ) '<****
	  		End If
   			.Rotate (rot_angle-90)
    End Select
    .SetDoubleProperty ("Capacitance",  blockname_variable  )
		.create
   End With

	With Block ' check existance
		.name "GND__"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   		End With
   		If Not differential_flag Then
  			With Block
			.Reset
			.Type ("Ground")
			.name "GND__"+CStr(Discrete_Port_List(index_p).nr)

			Select Case orientation
    			Case "L"
					If Not tuningC_flag Then
    	  	 			.position (mws_block_port_position_x +1.5*x_offset+ L_index*x_offset/2 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.00 )) '<*
					Else
			 			.position (mws_block_port_position_x +1.5*x_offset+ L_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.00 )) '<*
					End If
    						.rotate (180)
    			Case "R"
    				If Not tuningC_flag Then
    	  	 			.position (mws_block_port_position_x +1.5*x_offset+ R_index*x_offset/2 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.00 )) '<**
					Else
			 			.position (mws_block_port_position_x +1.5*x_offset+ R_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.00 )) '<**
					End If
    						.rotate (180)
    			Case  "D"
    				If Not tuningC_flag Then
				 		.position (mws_block_port_position_x +1.5*x_offset -Int(D_index*size_offset/1.00) , mws_block_port_position_y+1.5*y_offset +D_index*y_offset/2. ) '<***
	  				Else
	  					.position (mws_block_port_position_x +1.5*x_offset -Int(D_index*size_offset/1.00) , mws_block_port_position_y+1.5*y_offset +D_index*y_offset/1. ) '<***
	  				End If
    					.rotate (90)
				Case "U"
					If Not tuningC_flag Then
				 		.position (mws_block_port_position_x +1.5*x_offset -Int(U_index*size_offset/1.00) , mws_block_port_position_y+1.5*y_offset +U_index*y_offset/2. ) '<****
	  				Else
	  					.position (mws_block_port_position_x +1.5*x_offset -Int(U_index*size_offset/1.00) , mws_block_port_position_y+1.5*y_offset +U_index*y_offset/1. ) '<****
	  				End If
    					.rotate (90)
   				End Select
			.Create
			End With
		End If
  	End If
  End If

  If dlg.add_caps = 0 Then
   With Block ' check existance
		.name "Port__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block ' check existance
		.name "GND__"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
  End If

  ' draw the EXTERNAL port ----------------------------
  'first delete existing DS-ports
  With ExternalPort
	make_abs =  CStr(Abs(CInt(Discrete_Port_List(index_p).nr)))
	.Name make_abs
	If  .doesexist Then
 		.delete
 	End If
  	'DS-ports connected to higher modes
    cst_nr_of_modes = Abs(cint(ListArray_mode(index_p)))
    For cst_iiia = 1 To cst_nr_of_modes
    	If cst_nr_of_modes > 1 Then
	     .Name  CStr(port_1000+cst_iiia+port_nr_offset*Abs(CInt(Discrete_Port_List(index_p).nr)))  'port_nr_offset As Integer, port_1000 As Integer
			If  .doesexist Then
 			.delete
 			End If
 		End If
 	Next
  End With

  With ExternalPort 'external DS-Ports representing deembedded MWS-dis.ports
   make_abs =  CStr(Abs(CInt(Discrete_Port_List(index_p).nr)))
   .Name make_abs 'Discrete_Port_List(index_p)
   If CInt(Discrete_Port_List(index_p).nr) > 0 Then 'only for discr. ports
   If Not tuningC_flag Then

	Select Case orientation
    	Case "L"
    		 .position (mws_block_port_position_x +2*x_offset+ L_index*x_offset/2 , mws_block_port_position_y+y_offset )	'<p
    	Case "R"
    		 .position (mws_block_port_position_x +2*x_offset+ R_index*x_offset/2 , mws_block_port_position_y+y_offset )	'<p
    	Case  "D"
			.position  (mws_block_port_position_x +x_offset ,                     mws_block_port_position_y + 2*y_offset+ D_index*y_offset/2)  '<p
		Case "U"
			 .position (mws_block_port_position_x +x_offset ,                     mws_block_port_position_y + 2*y_offset+ U_index*y_offset/2) '<p
    End Select

	.create
	.SetFixedImpedance( True)
	.SetImpedance CStr(Port.GetLineImpedance (Discrete_Port_List(index_p).nr,1) )

	.SetDifferential differential_flag
   End If

   Else	'external DS-ports (for MWS-WG-Ports, no deembedding required)

    ' check if higher modes are present
    cst_nr_of_modes = Abs(cint(ListArray_mode(index_p)))
    For cst_iiia = 1 To cst_nr_of_modes
    	If cst_nr_of_modes > 1 Then
	.Name  CStr(port_1000+cst_iiia+port_nr_offset*Abs(CInt(Discrete_Port_List(index_p).nr)))	'compose new port: numbers port_nr_offset As Integer, port_1000 As Integer

	 mws_block_port_position_x  = getPortPos (Port.StartPortNumberIteration-1,index_p,cst_iiia,"x")
  	 mws_block_port_position_y  = getPortPos (Port.StartPortNumberIteration-1,index_p,cst_iiia,"y")

  	 getPortOrientation (mws_block_center_position_x, mws_block_center_position_y, mws_block_port_position_x, mws_block_port_position_y,size_offset  , x_offset    ,   _
			y_offset  , rot_angle  , orientation   )

  	 Select Case orientation	'for multi mode
    	Case "L"
    			L_index=L_index+1
    	Case "R"
				R_index=R_index+1
    	Case  "D"
				D_index=D_index+1
		Case "U"
				U_index=U_index+1
    End Select

		Select Case orientation
    			Case "L"
    				.position (mws_block_port_position_x+ x_offset/8 -size_offset, mws_block_port_position_y+ y_offset/8 )
    			Case  "R"
    				.position (mws_block_port_position_x+ x_offset/8 +size_offset, mws_block_port_position_y+ y_offset/8 )
    			Case  "D"
					 .position (mws_block_port_position_x+ x_offset/4   ,mws_block_port_position_y+ Int(P_index*y_offset*.4) +size_offset )
				Case "U"
					.position (mws_block_port_position_x+ x_offset/4 ,  mws_block_port_position_y- Int(P_index*y_offset*.4)  -size_offset)
   		End Select
   		P_index = P_index + 1

	Else	'single mode

		Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x+x_offset/4 ,mws_block_port_position_y+y_offset/4  )	'<p
    			Case  "D"
					.position (mws_block_port_position_x+x_offset/4 ,mws_block_port_position_y+y_offset/4  )	'<p
				Case "U"
					.position (mws_block_port_position_x+x_offset/4 ,mws_block_port_position_y+y_offset/4 )		'<p
   		End Select

	End If

	.create
	'.SetFixedImpedance (True)
		.SetDifferential differential_flag
	Next cst_iiia

   End If
  End With
 ' End If
'---------------------------------------------------------------

'c instead of ports
If tuningC_flag Then
If CInt(Discrete_Port_List(index_p).nr) > 0 Then 'only for discr. ports
   blockname_variable = "dC_"+CStr(Discrete_Port_List(index_p).nr)
   If Units.GetCapacitanceSIToUnit() >= 1e6 Then
    DS.storeparameter (blockname_variable ,Format(0*Units.GetCapacitanceSIToUnit(),"###0.#######"))', blockname_variable +" C in "+  CStr(Units.GetUnit("Capacitance")))
   Else
	DS.storeparameter (blockname_variable ,Format(0*Units.GetCapacitanceSIToUnit(),"0.000E+00"))
   End If

   With Block ' check existance
		.name "d__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name "d__C_"+CStr(Discrete_Port_List(index_p).nr)
    Select Case orientation
    	Case "L"
    		.position (mws_block_port_position_x +1.5*x_offset+x_offset/2+ L_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.25 )) '<*
		Case "R"
			.position (mws_block_port_position_x +1.5*x_offset+x_offset/2+ R_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.25 )) '<**
    		Case  "D"
			.position (mws_block_port_position_x + x_offset - Int(D_index*size_offset/1.25 ), mws_block_port_position_y +1.5*y_offset+y_offset/2+ D_index*y_offset/1) '<***
		Case "U"
			.position (mws_block_port_position_x + x_offset - Int(U_index*size_offset/1.25 ), mws_block_port_position_y +1.5*y_offset+y_offset/2+ U_index*y_offset/1) '<****

    End Select
    Select Case orientation
    	Case "L"
    		 .Rotate (rot_angle+90)
    	Case "R"
             .Rotate (rot_angle-90)
    	Case  "D"
			 .Rotate (rot_angle+90)
		Case "U"
				.Rotate (rot_angle-90)
    End Select
    .SetDoubleProperty ("Capacitance",  blockname_variable  )
    .create
   End With

' gnd for instead
   With Block ' check existance
		.name "dGND__"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   	If Not differential_flag Then
  		With Block
			.Reset
			.Type ("Ground")
			.name "dGND__"+CStr(Discrete_Port_List(index_p).nr)

			Select Case orientation
    			Case "L"
					.position (mws_block_port_position_x +1.5*x_offset+x_offset/2+ L_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(L_index*size_offset/1.00 )) '<*
					.rotate (180)
				Case "R"
					.position (mws_block_port_position_x +1.5*x_offset+x_offset/2+ R_index*x_offset/1 , mws_block_port_position_y+y_offset - Int(R_index*size_offset/1.00 )) '<**
					.rotate (180)
    				Case  "D"
    					.position (mws_block_port_position_x + x_offset - Int(D_index*size_offset/1.00 ), mws_block_port_position_y +1.5*y_offset+y_offset/2+ D_index*y_offset/1) '<***
					.rotate (90)
				Case "U"
						.position (mws_block_port_position_x + x_offset - Int(U_index*size_offset/1.00 ), mws_block_port_position_y +1.5*y_offset+y_offset/2+ U_index*y_offset/1) '<****
					.rotate (90)
   				End Select
			.Create
		End With
     End If
 End If
 Else
 	With Block ' check existance
		.name "d__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block
			.name "dGND__"+CStr(Discrete_Port_List(index_p).nr)
	If  .doesexist Then
 		 .delete
 		End If
 		End With
End If
' end instead
'--------------------------

  'draw the connection lines
  If dlg.Createlinks = 1 Then

   If CInt(Discrete_Port_List(index_p).nr) > 0 Then 'only for discr. ports
   If Not tuningC_flag Then
	'Links between ext.port and L
	ConnectComponents(Array(Array("P", Discrete_Port_List(index_p).nr,"0"), Array("B", "Port__L_"+CStr(Discrete_Port_List(index_p).nr),"2")))
    End If
	'between L and MWS-ports
	ConnectComponents(Array(Array("B", "Port__L_"+CStr(Discrete_Port_List(index_p).nr),"1"), Array("B", SchematicBlockName, Discrete_Port_List (index_p+0).nr)))
	
    If dlg.add_caps = 1 Then
		'Links between C and L
		Select Case orientation
		Case "L", "D"
			ConnectComponents(Array(Array("B", "Port__L_"+CStr(Discrete_Port_List(index_p).nr),"2"), Array("B", "Port__C_"+CStr(Discrete_Port_List(index_p).nr),"2")))
		Case  "R", "U"
			ConnectComponents(Array(Array("B", "Port__L_"+CStr(Discrete_Port_List(index_p).nr),"2"), Array("B", "Port__C_"+CStr(Discrete_Port_List(index_p).nr),"1")))
		End Select
    End If

    If dlg.add_caps = 1 Then
	 	If Not differential_flag Then
    		'between C and ground
	 		Select Case orientation
    		Case "L","D"
				ConnectComponents(Array(Array("B", "GND__"   + Cstr(Discrete_Port_List(index_p).nr),"1"), Array("B", "Port__C_"+ CStr(Discrete_Port_List(index_p).nr),"1")))
			Case  "R","U"
				ConnectComponents(Array(Array("B", "GND__"   + Cstr(Discrete_Port_List(index_p).nr),"1"), Array("B", "Port__C_"+ CStr(Discrete_Port_List(index_p).nr),"2")))
    		End Select
		End If
	End If

	'instead Cs
	If tuningC_flag Then
		If dlg.add_caps = 1 Then
			If Not differential_flag Then
				'between C and ground
				 Select Case orientation
				Case "L","D"
					ConnectComponents(Array(Array("B", "dGND__"   + Cstr(Discrete_Port_List(index_p).nr),"1"), Array("B", "d__C_"+ CStr(Discrete_Port_List(index_p).nr),"2")))			 
				Case  "R","U"
					ConnectComponents(Array(Array("B", "dGND__"   + Cstr(Discrete_Port_List(index_p).nr),"1"), Array("B", "d__C_"+ CStr(Discrete_Port_List(index_p).nr),"2")))			 
				End Select
			End If
		End If

		'Links between L and insteadCs
		ConnectComponents(Array(Array("B", "Port__L_"+CStr(Discrete_Port_List(index_p).nr),"2"), Array("B", "d__C_"+CStr(Discrete_Port_List(index_p).nr),"1")))			 		
	End If



    If differential_flag Then
		If Not tuningC_flag Then
		'between ext.port and discrete MWS-port , only for the differential case
		ConnectComponents(Array(Array("P", Discrete_Port_List(index_p).nr,"1"), Array("B", SchematicBlockName,    Discrete_Port_List(   index_p+0).nr+ "'")))

	End If
	 
	If dlg.add_caps = 1 Then
		'between Caps and discrete MWS-port , only for the differential case
		Select Case orientation
		
		Case "L","D"
			Source = Array("B", "Port__C_"+ CStr(Discrete_Port_List(index_p).nr),"1")
		Case  "R","U"
			Source = Array("B", "Port__C_"+ CStr(Discrete_Port_List(index_p).nr),"2")
		End Select

		ConnectComponents(Array(Source, Array("B", SchematicBlockName,    Discrete_Port_List(   index_p).nr+ "'")))
		
		If tuningC_flag Then'	Join C And dC (where prev. was ground)
			' C and dC Grounds , only for the differential case
			ConnectComponents(Array(Array("B", "d__C_"+ CStr(Discrete_Port_List(index_p).nr),"2"), Array("B", SchematicBlockName,    Discrete_Port_List(   index_p).nr+ "'")))
		End If
	End If

    End If
   Else ' Waveguide links : direct link from port to MWS-block

	'between ext. port and MWS-port
	' check if higher modes are present
	 cst_nr_of_modes = Abs(cint(ListArray_mode(index_p)))
	 For cst_iiia = 1 To cst_nr_of_modes
		If cst_nr_of_modes > 1 Then
			'composed_port_number=  CStr(100+cst_iiia+Abs(CInt(Discrete_Port_List(index_p))))
			composed_port_number=  CStr(port_1000+cst_iiia+port_nr_offset*Abs(CInt(Discrete_Port_List(index_p).nr)))	'compose new port:   port_nr_offset As Integer, port_1000 As Integer
			composed_block_number = cstr(Abs(CInt(Discrete_Port_List(index_p).nr)))+"("+cstr(cst_iiia)+")"		' e.g. 10(3)
			ConnectComponents(Array(Array("P", composed_port_number, "0"), Array("B", SchematicBlockName,    composed_block_number)))
			If differential_flag Then
				ConnectComponents(Array(Array("P", composed_port_number, "1"), Array("B", SchematicBlockName,    composed_block_number + "'")))
			End If
		Else
			make_abs =  CStr(Abs(CInt(Discrete_Port_List(index_p).nr)))
			ConnectComponents(Array(Array("P", make_abs, "0"), Array("B", SchematicBlockName,    make_abs)))

			If differential_flag Then 'additional connection
			ConnectComponents(Array(Array("P", make_abs, "1"), Array("B", SchematicBlockName,    make_abs+ "'")))
			End If
		End If
	  Next


   End If	'dis or wg ports
  End If ' draw links
  '
  '

	'delete insteadCs+insteadGnds
   If Not tuningC_flag Then
   	With Block ' check existance
		.name "dGND__"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With
   With Block ' check existance
		.name "d__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
   End With

   End If

 loop_index = loop_index +1
 Next  ' loop over all ports


'--------------------------------
'--------------------------------
 ' coupling matrices:
 If dlg.port_coupling=1 Then

For index_p = 0 To  Total_Nr_of_WG_Ports+Total_Nr_of_Discrete_Ports-1
 For index_q = 0 To  Total_Nr_of_WG_Ports+Total_Nr_of_Discrete_Ports-1
  If index_q <> index_p  And index_q > index_p  Then
   If CInt(Discrete_Port_List(index_p).nr) > 0 And  CInt(Discrete_Port_List(index_q).nr) > 0 Then	'both partners are dis-ports

	'Get Port-Length, Start And End coordinates, L1 and L2
	glength1=  DiscretePort.getgridlength ( Discrete_Port_List(index_p).nr)
	glength2=  DiscretePort.getgridlength ( Discrete_Port_List(index_q).nr)

	portlength_mean = 	Units.GetGeometryUnitToSI()*(glength1+	glength2)/2

	DiscretePort.GetCoordinates ( Discrete_Port_List(index_p).nr,   x0_p,   y0_p,  z0_p,   x1_p,  y1_p,   z1_p  )
	DiscretePort.GetCoordinates ( Discrete_Port_List(index_q).nr,   x0_q,   y0_q,  z0_q,   x1_q,  y1_q,   z1_q  )
	center_pointx1  = (x0_p +x1_p)/2 :  center_pointy1  = (y0_p +y1_p)/2 :  center_pointz1  = (z0_p +z1_p)/2
	center_pointx2  = (x0_q +x1_q)/2 :  center_pointy2  = (y0_q +y1_q)/2 :  center_pointz2  = (z0_q +z1_q)/2

	port_distance = Units.GetGeometryUnitToSI()*Sqr( (center_pointx1-center_pointx2)^2 + (center_pointy1-center_pointy2)^2 + (center_pointz1-center_pointz2)^2 )

    blockname_variable = "Port_L_"+CStr(Discrete_Port_List(index_p).nr)
   	L_p = Units.GetInductanceUnitToSI()*DS.RestoredoubleParameter (blockname_variable)
   	blockname_variable = "Port_L_"+CStr(Discrete_Port_List(index_q).nr)
   	L_q = Units.GetInductanceUnitToSI()*DS.RestoredoubleParameter (  blockname_variable)

   	M = 2e-7*(portlength_mean*Log( (portlength_mean+Sqr(portlength_mean^2+port_distance^2))/port_distance) -  _
   				Sqr(portlength_mean^2+port_distance^2) +port_distance )
	If L_p <> 0 And L_q <> 0 Then
		k= M/Sqr(L_p*L_q)
	Else
		k=0
	End If



	'check angle between the two ports:
	If check_angle (x1_p-x0_p,y1_p-y0_p,z1_p-z0_p, x1_q-x0_q,y1_q-y0_q,z1_q-z0_q) > 45 Then ' greater than 45 deg : ignore coupling
		k=0
	End If
	'check if the length of the two ports is too different:  50% Limit
	If glength1/glength2 > 2 Or glength2/glength1 < 0.5 Then
		k=0
	End If
	If k<>0 Then
     DS.storeparameter ("k_"+cstr(Discrete_Port_List(index_p).nr)+"_"+cstr(Discrete_Port_List(index_q).nr) ,Format(-1.*(k),"###0.#######"))   ', blockname_variable +" L in "+  CStr(Units.GetUnit("Inductance")))
	End If
	With Block ' check existance
		.name "MUC_"+cstr( Discrete_Port_List(index_p).nr )+"_"+cstr( Discrete_Port_List(index_q).nr)   ' "Port__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
    End With
   If k <> 0 Then
    With Block
        .reset
        .name "MUC_"+cstr( Discrete_Port_List(index_p).nr )+"_"+cstr( Discrete_Port_List(index_q).nr)
        .type "CircuitBasic\Mutual Coupling"
        .position (mws_block_center_position_x-size_offset+ (size_offset/2)*index_p, mws_block_center_position_y+5*size_offset+ (size_offset/2)*index_q)
        .setdoubleproperty ("Coupling",  "k_"+cstr(Discrete_Port_List(index_p).nr)+"_"+cstr(Discrete_Port_List(index_q).nr)  )
        .SetStringProperty ("Inductor1", "Port__L_"+cstr( Discrete_Port_List(index_p).nr ))
        .SetStringProperty ("Inductor2", "Port__L_"+cstr( Discrete_Port_List(index_q).nr ))
        .create
    End With
   End If		'k<>0
   End If	'if index_p > 0
  End If	'symm p<>q
 Next 	'index_q
Next 	'index_p
'--- end coupling matrices
Else
	For index_p = 0 To Total_Nr_of_Discrete_Ports-1
 For index_q = 0 To Total_Nr_of_Discrete_Ports-1
  If index_q <> index_p  And index_q > index_p  Then
	With Block ' check existance
		.name "MUC_"+cstr( Discrete_Port_List(index_p).nr )+"_"+cstr( Discrete_Port_List(index_q).nr)   ' "Port__C_"+CStr(Discrete_Port_List(index_p).nr)
 		If  .doesexist Then
 		 .delete
 		End If
    End With
     DS.deleteparameter ("k_"+cstr(Discrete_Port_List(index_p).nr)+"_"+cstr(Discrete_Port_List(index_q).nr) )
    End If
    Next
    Next
End If
'-------------------------------------------
'----------------------------------------------


 Open (getprojectpath("Model3D"))+"Neg_L_C_Readme.txt" For Output As #33
 Print #33, "--------------------------------------------------"+ vbCrLf + CStr(Date) + " " + CStr(Time)  +   vbCrLf
 Print #33, "--------------------------------------------------"
 Print #33, output_string + vbCrLf
 Close #33
 'info Window (report)
 If dlg.report = 1 Then
   Start( (getprojectpath("Model3D"))+"Neg_L_C_Readme.txt")
 End If

 MsgBox "Deembedding macro completed"

End Sub

Sub search_for_xlower_x_upper (delta As Double, long_ref_index As Long, x_upper As Double, x_lower As Double)

Dim X_start_index As Long, X_istop_up As Long, X_istop_down As Long, x As Double, y As Double, z As Double, N_loops As Integer
    N_loops = 0
	y= Mesh.getypos(long_ref_index)
	z= Mesh.getzpos(long_ref_index)

'x: up
    X_start_index = long_ref_index
    x= Mesh.getxpos(long_ref_index)
    Do
    	x = x + delta : X_istop_up = Mesh.GetClosestPtIndex (  x,y,z ) : N_loops = N_loops +1
	Loop Until X_istop_up > X_start_index  Or N_loops > N_max_searches 	'found the next higher mesh index
	x_upper =  Mesh.getxpos(X_istop_up)	' and get the position
'x: down
	x= Mesh.getxpos(long_ref_index)	'reset position
	N_loops =0
	Do
    	x = x - delta : X_istop_down = Mesh.GetClosestPtIndex (  x,y,z ): N_loops = N_loops +1
	Loop Until  X_istop_down < X_start_index Or N_loops > N_max_searches
	x_lower = Mesh.getxpos(X_istop_down)
	x= Mesh.getxpos(long_ref_index)	'reset x position

	If (X_istop_up = X_start_index) Then		' no upper meshnode found, using mirrowed lower value
		x_upper = Mesh.getxpos(long_ref_index) + (Mesh.getxpos(long_ref_index) - x_lower)
	End If
	If (X_istop_down = X_start_index) Then		' no lower meshnode found, using mirrowed upper value
		x_lower = Mesh.getxpos(long_ref_index) + (Mesh.getxpos(long_ref_index) - x_upper)
	End If

End Sub

Sub search_for_ylower_y_upper (delta As Double, long_ref_index As Long, y_upper As Double, y_lower As Double)

Dim Y_start_index As Long, Y_istop_up As Long, Y_istop_down As Long, x As Double, y As Double, z As Double, N_loops As Integer
	N_loops = 0
	x= Mesh.getxpos(long_ref_index)
	z= Mesh.getzpos(long_ref_index)

'y: up
	Y_start_index = long_ref_index
	y= Mesh.getypos(long_ref_index)
    Do
    	y = y + delta : Y_istop_up = Mesh.GetClosestPtIndex (  x,y,z ): N_loops = N_loops +1
	Loop Until Y_istop_up > Y_start_index Or N_loops > N_max_searches
	y_upper =  Mesh.getypos(Y_istop_up)
'y: down
	y= Mesh.getypos(long_ref_index)
	N_loops =0
	Do
    	y = y - delta : Y_istop_down = Mesh.GetClosestPtIndex (  x,y,z ): N_loops = N_loops +1
	Loop Until  Y_istop_down < Y_start_index Or N_loops > N_max_searches
	y_lower=Mesh.getypos(Y_istop_down)
	y= Mesh.getypos(long_ref_index)

	If (Y_istop_up = Y_start_index) Then		' no upper meshnode found, using mirrowed lower value
		y_upper = Mesh.getypos(long_ref_index) + (Mesh.getypos(long_ref_index) + y_lower)
	End If
	If (Y_istop_down = Y_start_index) Then		' no lower meshnode found, using mirrowed upper value
		y_lower = Mesh.getypos(long_ref_index) + (Mesh.getypos(long_ref_index) - y_upper)
	End If

End Sub

Sub search_for_zlower_z_upper (delta As Double, long_ref_index As Long, z_upper As Double, z_lower As Double)

Dim Z_start_index As Long, Z_istop_up As Long, Z_istop_down As Long, x As Double, y As Double, z As Double, N_loops As Integer
	N_loops = 0
	x= Mesh.getxpos(long_ref_index)
	y= Mesh.getypos(long_ref_index)

'z: up
	Z_start_index = long_ref_index
	z= Mesh.getzpos(long_ref_index)
    Do
    	z = z + delta : Z_istop_up = Mesh.GetClosestPtIndex (  x,y,z ): N_loops = N_loops +1
	Loop Until Z_istop_up > Z_start_index Or N_loops > N_max_searches
	z_upper =  Mesh.getzpos(Z_istop_up)
'z: down
	z= Mesh.getzpos(long_ref_index)
	N_loops=0
	Do
    	z = z - delta : Z_istop_down = Mesh.GetClosestPtIndex (  x,y,z ): N_loops = N_loops +1
	Loop Until  Z_istop_down < Z_start_index Or N_loops > N_max_searches
	z_lower=Mesh.getzpos(Z_istop_down)
	z= Mesh.getzpos(long_ref_index)

	If (Z_istop_up = Z_start_index) Then		' no upper meshnode found, using mirrowed lower value
		z_upper = Mesh.getzpos(long_ref_index) + (Mesh.getzpos(long_ref_index) + z_lower)
	End If
	If (Z_istop_down = Z_start_index) Then		' no lower meshnode found, using mirrowed upper value
		z_lower = Mesh.getzpos(long_ref_index) + (Mesh.getzpos(long_ref_index) - z_upper)
	End If


End Sub

Function BaseName  (path As String) As String

        Dim dircount As Integer, extcount As Integer, filename As String

        dircount = InStrRev(path, "\")
        filename = Mid$(path, dircount+1)
        extcount = InStrRev(filename, ".")
        BaseName  = Left$(filename, IIf(extcount > 0, extcount-1, 999))

End Function

Sub Start (lib_filename As String)

        On Error GoTo Win95
        WINNT:
                Shell "cmd /c " + Quote(lib_filename)
                Exit Sub
        Win95:
                Shell "start " + Quote(lib_filename)
                Exit Sub

End Sub
Function Quote (lib_Text As String) As String

        Quote = Chr$(34) + lib_Text + Chr$(34)

End Function
Function compute_C (d As Double, xl As Double, h As Double) As Double
	' cyl.-rod above groundplane: h distance, d diameter, xL length
	If xl > 0 Then
	 	compute_C = (4*pi*8.856e-12*xl/(Log((4*h+xl+Sqr(d^2+(4*h+xl)^2))/(4*h+3*xl+Sqr(d^2+(4*h+3*xl)^2))) -Log((-xl+Sqr(d^2+xl^2))/(xl+Sqr(d^2+xl^2))) ) )
	Else
		compute_C=0
	End If
End Function

Function compute_C_tetmesh_new (d As Double, xl As Double, h As Double) As Double
	If xl > 0 Then
	' cyl.-rod above groundplane, parallel to ground: h distance, d diameter, xL length
		compute_C_tetmesh_new = 4*pi*8.856e-12*xl/(Log((4*xl^2*(-xl+Sqr(xl^2+16*h^2)))/(d^2*(xl+Sqr(xl^2+16*h^2)))))
	Else
		compute_C_tetmesh_new=0
	End If
End Function

Function compute_L_a (r As Double, xl As Double) As Double
	'compute_L_a = Units.GetGeometryUnitToSI()*xl*2*1e-7*(Log(2*xl/r)-1) ' simplified form
	compute_L_a = Units.GetGeometryUnitToSI()*2*1e-7*(xl*Log((xl+Sqr(xl^2+r^2))/(r))-(Sqr(xl^2+r^2))+ r) ' exact form
End Function



Function correction_factor_dis_ports (r As Double,L As Double, gap As Double) As Double
'
' Interpolation of the correction factor:
'r/L     Correction
' 0-0.025  	3-2.5
'0.025-.1	2.5-1.8
'0.1-0.5		1.8-1.2
' > 0.5 		1.2
'
' h/L  corr
' 0 1.8
' 0.1 1
' 0.2 3.0
'
	If L = 0 Or Abs(L) < 1e-7 Then
		correction_factor_dis_ports=1.
		Exit Function
	End If

If r/L >=0 And r/L <= 0.025 Then
	correction_factor_dis_ports= 2.5 +(.5/0.025)*(0.025-r/L)
	correction_factor_dis_ports=correction_factor_dis_ports*(2.5*(gap/L)+0.75)
	Exit Function
End If
If r/L >=0.025 And r/L <= 0.1 Then
	correction_factor_dis_ports= 1.8 +(.7/0.075)*(0.1-r/L)
	correction_factor_dis_ports=correction_factor_dis_ports*(2.5*(gap/L)+0.75)
	Exit Function
End If
If r/L > 0.1 And r/L <= 0.5 Then
	correction_factor_dis_ports= 1.2 +(0.6/0.4)*(0.5-r/L)
		    correction_factor_dis_ports=correction_factor_dis_ports*(2.5*(gap/L)+0.75)
	Exit Function
End If
If r/L > 0.5 Then
	correction_factor_dis_ports= 1.2
		    correction_factor_dis_ports=correction_factor_dis_ports*(2.5*(gap/L)+0.75)
	Exit Function
	'MsgBox "Ratio Radius/Length is too large :" + cstr(r/L)+ "; should be less than 0.5"
	'Exit All
End If

End Function

Function correction_factor_face_ports (r As Double,L As Double, gap As Double) As Double
'
' Interpolation of the correction factor:
'r/L     Correction
' 0-0.025  	2.6-3.1
'0.025-.2	3.1
'.2-0.5		3.1-4
' 0.5 1		4-7
'1 20       7-121
'
' h/L(gap/L)  corr
'0  3/4
'0.1 1
'0.2 5/4
'0.3 6/4
'
'MsgBox "faceport: r/L " +cstr(r/L)+ " gap/L : " + cstr(gap/L)

	If L = 0 Or Abs(L) < 1e-7 Then
		correction_factor_face_ports=1.
		Exit Function
	End If

If r/L >=0 And r/L <= 0.025 Then
	correction_factor_face_ports= 2.6 +(.5/.025)*(r/L)
	correction_factor_face_ports=correction_factor_face_ports*(2.5*(gap/L)+3/4)
	Exit Function
End If
If r/L >=0.025 And r/L <= 0.2 Then
	correction_factor_face_ports= 3.1
	correction_factor_face_ports=correction_factor_face_ports*(2.5*(gap/L)+3/4)
	Exit Function
End If
If r/L > 0.2 And r/L <= 0.5 Then ''''
	correction_factor_face_ports=  3*((r/L)-0.2)+ 3.1
	correction_factor_face_ports=correction_factor_face_ports*(2.5*(gap/L)+3/4)
	Exit Function
End If
If r/L > 0.5 And r/L <= 1 Then ''''
	correction_factor_face_ports=  6*((r/L)-0.5)+ 4
	correction_factor_face_ports=correction_factor_face_ports*(2.5*(gap/L)+3/4)
	Exit Function
End If
If r/L > 1 And r/L <= 20 Then ''''
	correction_factor_face_ports=  6*((r/L)-1)+ 7
	correction_factor_face_ports=correction_factor_face_ports*(2.5*(gap/L)+3/4)
	Exit Function
End If
If r/L > 20 Then
	MsgBox "Ratio Radius/Length for face-Ports is too large :" + cstr(r/L)+ "; should be less than 20."
	Exit All
End If

End Function

Function check_angle (x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double) As Double
	Dim cos_phi As Double
	cos_phi =  ( (x1*x2 + y1*y2 + z1*z2) /(Sqr(x1^2+y1^2+z1^2)*Sqr(x2^2+y2^2+z2^2)) )
	If cos_phi = 0 Then
		check_angle = 90.
	Else
		check_angle = (180/pi)*Atn(Sqr(1-cos_phi^2)/cos_phi)
	End If
End Function


Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Debug.Print "Action="; Action%
    Select Case Action%
    Case 1 ' Dialog box initialization
        If Mesh.GetMeshType= "Tetrahedral" Then
			DlgEnable "Group1",False
			DlgEnable "OptionButton1",False
			DlgEnable "OptionButton2",False
			DlgValue "solver_type", 1	'set to !F solver type
		Else
			DlgValue "solver_type", 0 '!T
        End If
        'Beep
    Case 2 ' Value changing or button pressed
        Select Case DlgItem$
        Case "Disable"
            DlgText DlgItem$,"&Enable"
            DlgEnable "Text",False
            DialogFunc = True 'do not exit the dialog
        Case "Enable"
            DlgText DlgItem$,"&Disable"
            DlgEnable "Text",True
            DialogFunc = True 'do not exit the dialog
        End Select
    End Select
End Function

Sub get_all_child_names (potentialname As String, potentialarray() As String, pot_array_index As Integer)

 Dim firstname As String, nextname As String, secondname As String
 Dim index_a As Integer, index_b As Integer
 pot_array_index = 0
 firstname =  Resulttree.GetFirstChildName (potentialname)
 Do
	 If InStr(firstname,potentialname) Then
		pot_array_index = pot_array_index +1
    	ReDim Preserve potentialarray(pot_array_index)
    	firstname = Mid(firstname,Len(potentialname)+1,Len(firstname))
    	potentialarray(pot_array_index-1)= firstname
	 End If
	firstname = Resulttree.GetNextItemName (potentialname+firstname)
 Loop Until firstname =""
End Sub

Sub setup_mode_list (modelist() As String, Total_Nr_of_Discrete_Ports As Integer, Total_Nr_of_WG_Ports As Integer)

 Dim cst_runindex As Integer, nr_of_modes_found As Integer
 Dim mode_type As String, maxmodenr As Integer, maxmodeatport As Integer, Number_of_Modes As Integer
 Dim len_of_Field As Integer, run_index As Integer, mode_impedance As Double

 ReDim  Preserve modelist (Port.StartPortNumberIteration-1)
 Total_Nr_of_Discrete_Ports =0 :  Total_Nr_of_WG_Ports =0

 For cst_runindex = 0 To Port.StartPortNumberIteration-1

 If Port.getType(Abs(cint(Discrete_Port_List(cst_runindex).nr))) = "Waveguide" Then
' Waveguide-Port Check
  modelist (cst_runindex) = CStr(-Port.GetNumberOfModes (Abs(cint(Discrete_Port_List(cst_runindex ).nr)))  )	'-neg.sign for wg - ports
  Total_Nr_of_WG_Ports= Total_Nr_of_WG_Ports + 1
Else
  ' Discrete-Port Check
	modelist (cst_runindex) = "1"'CStr(Port.GetNumberOfModes (portnamearray( cst_runindex )))
	Total_Nr_of_Discrete_Ports = Total_Nr_of_Discrete_Ports + 1
 ' end discrete-ports
 End If
 Next

End Sub

Private Sub FillPortNameArray()
' -------------------------------------------------------------------------------------------------
' FillPortNameArray: This function fills the global array of port names
' -------------------------------------------------------------------------------------------------

	Dim nIndex As Long, nCount As Long, nports As Integer

	' Determine the total number of ports first and reset the enumberation to the beginning of the
	' ports list.
	nports = Port.StartPortNumberIteration

	' Make the port name array large enough to hold all port names
	ReDim Discrete_Port_List(nports)

	' Now loop over all ports and add the ports to the port array
	nCount = 0
	Dim strPortName As Integer
	For nIndex = 0 To nports-1
		strPortName = Port.GetNextPortNumber
		Discrete_Port_List(nCount).nr = CStr(strPortName)
		Discrete_Port_List(nCount).label = Port.getlabel(strPortName)

	If Port.getType(Discrete_Port_List(nCount).nr) = "Waveguide" Then
		' Waveguide-Port Check
		Discrete_Port_List(nCount).nr= CStr(-strPortName) ' wg indicated as neg. sign
 	End If


		nCount = nCount + 1
	Next nIndex

	' Adjust the length of the port name array to the actual number
	If (nCount < nports) Then
		ReDim Preserve Discrete_Port_List(nCount)
	End If

End Sub

Function getPortPos (Nr_of_ports As Integer,look_for_index_p As Integer,look_for_mode_nr As Integer,x_or_y As String) As Double
	'look_for_mode_nr ...0 -> return posxy of the first available mode
	'look_for_mode_nr > 0  -> return posxy of the desired mode-nr

	Dim cst_index As Integer, port_nr As Integer, cst_nr_of_modes As Integer, 	composed_port_name As String, nr_of_real_pins As Integer

	'find out the number of all port-pins (included multipins/multimodes)
	nr_of_real_pins=0
	For cst_index = 0 To Nr_of_ports
		port_nr = cint(Discrete_Port_List(cst_index).nr)
		If port_nr < 0 Then
			cst_nr_of_modes = Abs(cint(ListArray_mode(cst_index)))
		Else
			cst_nr_of_modes = 1
		End If
		nr_of_real_pins = nr_of_real_pins + cst_nr_of_modes
	Next

	'compose the proper pin-name for the desired port-index
		port_nr = cint(Discrete_Port_List(look_for_index_p).nr)
		If port_nr < 0 Then
			cst_nr_of_modes = Abs(cint(ListArray_mode(look_for_index_p)))
		Else
			cst_nr_of_modes = 1
		End If
		If cst_nr_of_modes = 1 Then
			composed_port_name= cstr(Abs(port_nr))
		Else
			If look_for_mode_nr = 0 Then
				composed_port_name= cstr(Abs(port_nr))+"(1)"
			Else
				composed_port_name= cstr(Abs(port_nr))+"("+cstr(look_for_mode_nr)+")"
			End If
		End If


		For cst_index = 0 To nr_of_real_pins-1

			With Block
				.name SchematicBlockName
				If cstr(.GetPortName (cst_index))= composed_port_name Then
					If x_or_y = "x" Or x_or_y ="X" Then
						getPortPos = 	.GetPortPositionX (cst_index)
					Else
						getPortPos = 	.GetPortPositionY (cst_index)
					End If
				End If
			 End With
		Next cst_index

End Function


Sub getPortOrientation (XPosCenter As Long, YPosCenter As Long, XPosPort As Long, YPosPort As Long ,size_offset As Integer, _
						  x_offset As Integer  ,   y_offset As Integer ,   rot_angle As Double ,   orientation As String   )
'compute the orientation of the lumped elements; either left / right / up /down
'input XY center, XYPort, size_offset
'output: xy offset, angle, orientation

	Dim rel_x_pos As Double, rel_y_pos As Double

  rel_x_pos =  XPosCenter - XPosPort
  rel_y_pos =  YPosCenter - YPosPort

  If Abs(rel_x_pos) >= Abs(rel_y_pos) Then 'horizontal orientated
   If rel_x_pos <0 Then
  	x_offset =  size_offset : y_offset =  0 : 	rot_angle = 0
  	 orientation="R" 'to the right
   Else
	x_offset = - size_offset : y_offset =  0 : 	rot_angle = 180
	 orientation="L" 'to the left
   End If
  End If
  If Abs(rel_x_pos) < Abs(rel_y_pos) Then ' vertical orientated
   If rel_y_pos <0 Then
  	x_offset =   0 : y_offset =   size_offset : 	rot_angle = 90
   orientation="D" 	'   down
   Else
	x_offset =   0 : y_offset = - size_offset : 	rot_angle = -90
	 orientation="U" ' up
   End If
  End If

End Sub
