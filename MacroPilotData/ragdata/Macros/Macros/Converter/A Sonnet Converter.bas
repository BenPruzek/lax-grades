' Sonnet Converter.BAS
'
' (tested for Sonnet versions 8.x - 11.x) 
'
' read in by menu.cfg (File \ Import \ ...)
'
' ----------------------------------------------------------------------------------------------------------------------
' 31-Aug-2015 mac: allow import For Sonnet 14.x versions; corrected reading of normal polygons (ignore string lines)
' 13-Sep-2011 mac: allow import for Sonnet 13.x versions; corrected reading of box (prevent reading of sbox)
' 15-Nov-2010 ube: allow import for Sonnet 12.x versions
' 08-Sep-2010 fhi: corrected tolerance digits for xy-coordinates, added "SONNET" comment in header
' 08-Jan-2009 fhi: added "TOLEVEL TOP" for vias
' 20-Dec-2007 fhi: removed "staircase" option for 2D sheets (obsolete), "AddToHistory" replaced 2nd reopening of project
' 06-Dec-2007 fhi: upgraded to V2008 (Save-Dir, .cst file extensions and getprojectPath() adapted
' 17-Jan-2006 ube: save satement included, so that it is working properly with new Openfile command
' 07-Jun-2005 fhi: extended to Sonnet Version 10.x (syntax compatible to 9.x at least for basic commands)
' 24-nov-2004 fhi: command '@ Optimize Mesh for planar structures" skipped for MWS 5.0.x
' 24-Nov-2004 fhi: For EdgeVias the "TOLEVEL TOP" command correctly interpreted;
' 23-Nov-2004 fhi: significant digits after comma for polynomial points (round up) in userdialog; reordered the boolean insert
' 11-Nov-2004 ube: MinimumLineNumber replaced by MinimumStepNumber
' 22-Oct-2004 ube: Const max_nr_of_vertices increased from 300 to 3000
' 02-Jun-2004 fhi: met-thickness = 0 allowed; metalization-layers are sheets (staircase-mesh)
' 01-Jun-2004 fhi: No TST at Ports; Mirrored Y-components corrected; cylinders can be created rather than polygonal shaped vias, add arbitrary MeshFixpoints in Z
' 13-May-2004 fhi: WG-portshielding on; Dual-Proc.=2, Warnings displayed in separate MsgBOXes
' 11-May-2004 fhi: smallest mesh-distance (rather than mesh-ratio); substrate higher mesh-priority (1); transparency on/off for Substrate-Layers
' 10-May-2004 fhi: correction of met-layer index/diel-index; correct conversion ohms/square->conductivity;check double-points on polygons
' 15-Jan-2004 fhi: correction of position TMET and BMET MetalPlane
' 12-Jan-2004 fhi: implement. of "TOLEVEL GND"; correction of hight-extensions
' 12-Jan-2004 ube: vias with more than 5 corner points are not considered in automesh (instead a wire is created in the center for fix mesh point)
' 12-Jan-2004 ube: included planar mesh settings + energy based
' 08-Jan-2004 fhi: added Diel-Boxes and normal Vias for bool-intersect, read all types for metal, added individual thicknesses
' 07-Jan-2004 fhi: updated syntax for MWS v5
' 07-oct-2003 fhi: correct edit-filename directory; correction of some of the history syntaxes for displaying "properties"
' 06-Oct-2003 ube: small cosmetics, bas-file
' 05-Oct-2003 fhi: correction of extrude-operation defined by the sign of the polygon area
' 01-Oct-2003 fhi: boolean operations
' 25-Sep-2003 fhi: Initial version
'
'
' todo:
' 2-Jun-2004: fhi: Setup for the parameters; define autom. Ports; create at local CS;
'
'
'known problems: Extrusion of polygons: Not a single path !!!!!
'                          +--+
' a)                       |  |
' +---------+---+------+---+--+-+
' |         |   |      |
' +---------+   +------+
'
' b)
'crossing edges (8)
'
'  +------+
'  \       /
'   \    /
'     \/
'      +
'    /  \
'  /     \
'  +-----+
'
' c)
' zero area
'
'  1                   2
'   +-----------------+
'   |                 |
'   +----+----+-------+ 3
'   6    4    5
'
'--------------------------------------------------------------------------------------------
Option Explicit
Const MacroName = "Import~Sonnet"
Const max_nr_of_vertices = 3000 ' in a polyline
'--------------------------------------------------------------------------------------------
Public layer_position() As Double
Public diel_layer_position() As Double
Public cst_level_thickness As Double
Public Polygon_array () As Double			'contains Polygon's XY Data
Public Polygon_index () As Integer			'contains Polygon's Nr.,#of vertices
Public Evias_nr_array () As Integer			' contains EdgeVias info: Polygon-Nr, Startindex nd to_level
Public Evias_initvertex_array () As Integer
Public Evias_level_array () As Integer
Public diel_properties() As String
Public diel_properties_Nr() As String
Public metal_properties() As String
Public metal_properties_Nr() As String
Public metal_thickness() As Double

Public cst_sonnet_dir_filename As String
Public cst_export_dir_filename As String

Public material_inconsistent As Boolean
Public dlg_substrate_wireframe As Integer
Public duplicate_warning_text As String
Public not_created_polygon_warning_text As String
Public NEG_sign As Double
Public Number_of_Mesh_Layers As Integer
Public tolerance_digits As Integer


Sub Main ()

Dim cst_filestr_modexport As String
Dim nextline As String, nr_of_polygons_at_layer As Integer, Entries_all_Polygons () As String
Dim cst_complete_sonnet_filestr As String, nr_of_entries_polygons As Integer, Entries_Polygons_at_Layer() As String
Dim Entries_Via_Layer() As String,nr_of_via_layers As Integer
Dim Entries_Metal_Layer() As String,nr_of_Metal_entries As Integer
Dim All_Entries() As String,nr_of_All_entries As Integer, nr_of_found_metal_layers As Integer
Dim Entries_Normal_Via_Layer() As String,nr_of_Normal_via_layers As Integer, nr_of_found_normal_vias As Integer
Dim Entries_Normal_Vias_at_Layer()As String,nr_of_Normal_vias_at_layer As Integer
Dim Entries_Diel_blocks() As String, nr_of_found_diel_blocks As Integer
Dim cst_check_vias As Integer,  run_thru_diel_index As Integer, smallest_diel_thickness As Double
Dim cst_via_thickness As Double, Entries_Vias_at_Layer() As String, nr_of_vias_at_layer As Integer
Dim Entries_diel_blocks_at_Layer() As String,nr_of_diel_blocks_at_layer As Integer
Dim Entries_substrates() As String, nr_of_found_substrates As Integer
Dim Entries_substrates_at_Layer() As String, nr_of_substrates_at_layer As Integer

Dim cst_option_coord As Integer
Dim shapelist() As Integer
Dim cst_II As Integer, cst_KK As Integer, i As Integer, jj As Integer
Dim cst_JJ As Integer, skip_flag As Boolean
Dim cst_lossy As Integer, unit_scale As Double
Dim cst_kappa As String
Dim x_gauss(max_nr_of_vertices) As Double, y_gauss (max_nr_of_vertices) As Double, filename As String
Dim units_dim As String
Dim units_frq As String, cst_master_name As String, cst_slave_name As String, cst_master As Integer, cst_slave As Integer
Dim frq_factor As Double, dim_factor As Double
Dim fmin As Double, metal_conductivity As Double
Dim fmax As Double, T_metal_conductivity As Double, B_metal_conductivity As Double
Dim Nr_of_diel_layers  As Integer, cst_polygon_index As Integer
Dim X_size As Double, poly_debug_Nr As Integer, index_for_fixpoints As Integer
Dim Y_size As Double, poly_level As Integer, poly_met_type As Integer, to_poly_level As Integer
Dim	Nr_of_layer_positions As Integer, Nr_of_polygons As Integer

Dim Nr_of_met_layers As Double, bottom_GND As Double
Dim poly_check_type As String, cst_vertices_index As Integer
Dim diel_thickness() As Double, 	material_index As Integer
Dim diel_eps_r() As Double, evia_initial_vertex As Double
Dim diel_eps_i() As Double
Dim diel_mu_r() As Double
Dim diel_mu_i() As Double
Dim diel_tan_d() As Double
Dim diel_name() As String	, poly_nr_of_vertices As Integer, actual_layer_thickness_fixpoint As Double
	 
Dim  X_poly() As Double, Y_poly() As Double
Dim nr_of_diff_metals As Integer, metal_type_MET As String, metal_name_MET As String, MET_thickness As Double
Dim metal_name_TMET As String, metal_name_BMET As String, metal_type_TMET As String, metal_type_BMET As String
Dim boundary_zmax As String,  TMET_exists As Boolean, TMET_thickness As Double, BMET_thickness As Double
Dim boundary_zmin As String, BMET_exists As Boolean ,EVIA_exists As Boolean ,VIA_exists As Boolean
Dim counter_Normal_polygon As Integer, Nr_of_evias As Integer, Diel_Brick_exists  As Boolean
Dim nr_of_diff_diel As Integer, diel_name_MET As String
Dim merged_metal_layers As Integer, merged_via_layers As Integer, merged_evia_layers As Integer

Dim ex1 As Double, ey1 As Double, ez1 As Double
Dim ex2 As Double, ey2 As Double, ez2 As Double
Dim ex3 As Double, ey3 As Double, ez3 As Double
Dim ex4 As Double, ey4 As Double, ez4 As Double

Dim Info_text As String, merge_via As String, merge_evia As String, merge_metal As String
Dim warning_text As String

	Begin Dialog UserDialog 680,329,"Import Sonnet Projects 7.x - 14.x",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,7,640,56,"SONNET Import-File",.GroupBox1
		TextBox 40,28,480,21,.edit_filename
		PushButton 540,28,100,21,"Browse...",.push_browse
		GroupBox 20,70,640,56,"MWStudio Export File",.GroupBox2
		TextBox 40,91,480,21,.export_filename
		PushButton 540,91,100,21,"Browse...",.push_export
		GroupBox 20,133,640,98,"Default Options",.GroupBox3
		TextBox 190,154,120,21,.edit_thickness
		Text 30,161,160,14,"Metallization Thickness",.Text1
		OKButton 20,301,90,21
		CancelButton 120,301,90,21
		CheckBox 460,154,190,21,"Skip Boolean Insertions",.skip_bool
		CheckBox 460,182,190,14,"Substrates in Wireframe",.substrate_wireframe
		Text 320,161,110,14,"(in Sonnet-Units)",.Text2
		CheckBox 460,210,190,14,"Cylindrical Vias",.cyl_vias
		GroupBox 20,238,640,49,"Mesh-Settings",.GroupBox4
		Text 40,259,260,21,"Number of Meshcells/Substrate_Layer",.Text3
		TextBox 300,252,40,21,.Number_meshlayers
		Text 30,182,380,14,"[Thickness of Zero creates Sheet-Objects]",.Text4
		Text 30,210,140,14,"PointTolerance within ",.Text5
		TextBox 190,203,40,21,.tolerance
		Text 250,210,160,14,"digits ",.Text6
	End Dialog

Dim dlg As UserDialog

'set default values
With dlg
 .skip_bool=0
 .cyl_vias = 1 '0=polygons, 1=cylinders
 .substrate_wireframe = 1
' .export_filename = BaseName(getprojectbasename)+"_Sonnet_8_9.mod" 'later -> txt
 .export_filename =  getprojectpath("Project")+"_Sonnet_export_.mod" 'keep it as .mod !
 .edit_thickness="0."
 .tolerance = "3"
 .Number_meshlayers="1"
End With
dlg_substrate_wireframe = dlg.substrate_wireframe
warning_text=""
duplicate_warning_text=""
not_created_polygon_warning_text=""
NEG_sign = -1.			'"mirrors the y-component of polygons, etc.

'call dialog
If (Dialog(dlg) = 0) Then Exit All		'do the dialog
screenupdating False

If Convert2Double(dlg.edit_thickness) < 0 Then
 MsgBox " Metalization-Thickness must be greater/equal 0!"
 Exit All
End If

If CInt(dlg.Number_meshlayers) <= 0 Or CInt(dlg.Number_meshlayers) > 10 Then
	MsgBox " Select a Number between 1 and 10!"
	Exit All
End If

cst_complete_sonnet_filestr=cst_sonnet_dir_filename+"\"+dlg.edit_filename
cst_filestr_modexport=cst_export_dir_filename+"\"+dlg.export_filename	'complete export mod-file
If InStr(cst_filestr_modexport,".mod")=0 Then 'Or   InStr(cst_filestr_modexport,".cst")=0 Then
  cst_filestr_modexport=cst_filestr_modexport+".mod"
End If

'****** File import - export **************************************
'MsgBox "Sonnet-importfile "+cst_complete_sonnet_filestr
'MsgBox "Export-Modfile "+  cst_filestr_modexport
'****** File import - export **************************************

On Error Resume Next
Kill cst_filestr_modexport		'erase mod-file, if already exists
On Error GoTo 0

' open sonnet project file
On Error GoTo nofile
Open cst_complete_sonnet_filestr For Input As #1
On Error GoTo nofile_export
Open cst_filestr_modexport For Output  As #2
On Error GoTo 0


tolerance_digits=CInt(dlg.tolerance)
Number_of_Mesh_Layers=CInt(dlg.Number_meshlayers)
cst_level_thickness=Convert2Double(dlg.edit_thickness)
ReDim metal_thickness(1) 'define array
metal_thickness(0) =cst_level_thickness	'default
counter_Normal_polygon=0

'!!!!! if sheet objects , then ignore boolean operations!!!!!!!!!!!!!
'If cst_level_thickness = 0 Then
'	dlg.skip_bool=1
'End If

' check if it is a sonnet project and version number
Line Input #1, nextline				'1.st line
Line Input #1, nextline				'2nd line contains version info
Info_text = "Sonnet "+nextline
If InStr(nextline, "8.")=0 And InStr(nextline, "9.")=0 And InStr(nextline, "10.")=0 And InStr(nextline, "11.")=0  And InStr(nextline, "12.")=0 And InStr(nextline, "13.")=0 And InStr(nextline, "14.")=0 Then
 MsgBox "Not a Sonnet 8.x, 9.x, 10.x, 11.x, 12.x, 13.x or 14.x Project", "Error"
 Close #1
 Exit Sub
End	If

Info_text = Info_text +	vbCrLf+ "Sonnet-Import: "+cst_complete_sonnet_filestr+ vbCrLf+"Export: "+cst_filestr_modexport

frq_factor=1			'frequency scaling factor	
dim_factor=1			' dimension scaling factor

' read units 
Seek #1, 1				'rewind file
'serach for HEADER
skip_flag=False
While nextline<>"HEADER"
 If Not EOF(1) Then
  Line Input #1, nextline
 Else
   'MsgBox("No Header specified", "Warning")
   warning_text = warning_text + vbCrLf + "No Header specified"
   skip_flag = True
  Exit While
 End If
Wend	
If Not skip_flag Then
While nextline<>"END HEADER"
 Line Input #1, nextline	
 Print #2, "'@ "+ nextline
Wend 	
End If
'
Print #2, "'@ "
Print #2, "'@ "+ "SONNET conversion into CST-MWS"
Print #2, "'@ "
'
'
'search for DIMENSIONS
skip_flag=False
While nextline<>"DIM"
 If Not EOF(1) Then
  Line Input #1, nextline
 Else
   'MsgBox("No Dimensions specified", "Warning")
   warning_text = warning_text + vbCrLf + "No Dimensions specified"
   skip_flag= True
  Exit While
 End If
Wend
If Not skip_flag Then
 Do Until 	nextline="END DIM"
 Line Input #1, nextline
  If InStr(nextline,"FREQ") > 0 Then
  'GetString(nextline$)
   units_frq=GetSubString(nextline,2," ")
   'units_frq=GetString(nextline$)
   Select Case units_frq
	Case "THZ"
	 units_frq="GHZ"
	 frq_factor=1000
	Case "PHZ"
	 units_frq="GHZ"
	 frq_factor=1000000
	End Select
  End If
  
  If InStr(nextline,"LNG") > 0 Then
  'GetString(nextline$)
   units_dim=GetSubString(nextline,2," ")
   ' for compatibility reasons
   If units_dim="mils" Then
	 units_dim="mil"
   End If

  End If
  
 Loop 
End If		 
' read frequency range settings (if existing at all)
'search for FREQ
skip_flag=False
While nextline<>"FREQ"
 If Not EOF(1) Then
  Line Input #1, nextline
  If nextline="END CONTROL" Then 	'stop immediately prior to the GEO-Line
    'MsgBox("No Frequency-Range specified", "Warning")
    warning_text = warning_text + vbCrLf + "No frequency range specified"
   skip_flag=True
   Exit While
  End If
 Else
   'MsgBox("No Frequency-Range specified", "Warning")
   warning_text = warning_text + vbCrLf + "No Frequency range specified"
   skip_flag= True
  Exit While
 End If
Wend
If Not skip_flag Then
 Do Until 	nextline="END FREQ"
 Line Input #1, nextline
  If (InStr(nextline,"SIMPLE") > 0) Or (InStr(nextline,"SWEEP") > 0)   _
      Or (InStr(nextline,"ESWEEP") > 0)  Or (InStr(nextline,"ABS") > 0)Then
   
   fmin=Convert2Double( GetSubString(nextline,2," ")   )
   fmax=Convert2Double( GetSubString(nextline,3," ")   )

   ' set frequency settings
   If fmin > fmax Then
   Print #2, "'@ define frequency range"
   Print #2, "Solver.FrequencyRange " + evaluate(fmax*frq_factor) + " , "+ evaluate(fmin*frq_factor)
   Else
   Print #2, "'@ define frequency range"
   Print #2, "Solver.FrequencyRange " + evaluate(fmin*frq_factor) + " , "+ evaluate(fmax*frq_factor)
   End If
  End If
  If InStr(nextline,"STEP") > 0 Then
   fmin=0
   fmax=Convert2Double(GetSubString(nextline,2," ") ) 
   ' set frequency settings
   Print #2, "'@ define frequency range"
   Print #2, "Solver.FrequencyRange 0. , "+ evaluate(fmax*frq_factor) 			
  End If
 Loop 
End If

Select Case units_dim	'required to compute Ohms/square: R=L/(sigma.w.t) = 1/(sigma*t)*L/w= OhmS_Square *L/w
 Case "M"
   unit_scale = 1.
   Case "DM"
   unit_scale = 0.1
   Case "CM"
   unit_scale = 0.01
   Case "MM"
   unit_scale = 1.e-3
   Case "UM"
   unit_scale = 1.e-6
   Case "NM"
   unit_scale = 1.e-9
   Case "FT"
   unit_scale = 1./3.2808399
   Case "IN"
   unit_scale = 0.0254
   Case "MIL"
   unit_scale = 0.0254e-3
 Case Else
 unit_scale = 1.
End Select
		
'search for GEO Section
BMET_exists = False
TMET_exists = False
EVIA_exists = False
VIA_exists = False
Diel_Brick_exists= False
material_inconsistent=False
Nr_of_evias = 0
merged_metal_layers=0
merged_via_layers=0
merged_evia_layers=0
merge_metal=""
merge_via =""
merge_evia=""

'set the default Lossless-Material:
nr_of_diff_metals = 1
    ReDim  Preserve metal_properties(nr_of_diff_metals)
    ReDim  Preserve metal_properties_Nr(nr_of_diff_metals)
    metal_properties(nr_of_diff_metals-1)= "Lossless"
    metal_properties_Nr(nr_of_diff_metals-1)= "0"
	Print #2, "'@ define material: "+ "Lossless"
     Define_layer_PEC  "Lossless"

'set the default Lossless-Dielectricum(Vacuum):
	nr_of_diff_diel = 1
    ReDim  Preserve diel_properties(nr_of_diff_diel)
    ReDim  Preserve diel_properties_Nr(nr_of_diff_diel)
    diel_properties(nr_of_diff_diel-1)= "Vacuum_"
    diel_properties_Nr(nr_of_diff_diel-1)= "0"
	Print #2, "'@ define material: "+ "Vacuum_"
    Define_layer "Vacuum_",  1., 1.,0., 0., 0.,0.,0.,0.

' global WCS
      Print #2, "'@ activate global coordinates"
		Print #2, "WCS.ActivateWCS "+ Chr$(34)+"Global"+Chr$(34)
'align with xy-plane
	  Print #2, "'@ align wcs with global plane"
      Print #2, "WCS.SetNormal "+Chr$(34)+"0"+Chr$(34)+","+ Chr$(34)+"0"+Chr$(34)+","+ Chr$(34)+"1"+Chr$(34)  ' "0", "0", "1"
      Print #2, "WCS.SetOrigin "+Chr$(34)+"0"+Chr$(34)+","+ Chr$(34)+"0"+Chr$(34)+","+ Chr$(34)+"0"+Chr$(34)  ' 0", "0", "0"
      Print #2, "WCS.SetUVector "+Chr$(34)+"1"+Chr$(34)+","+ Chr$(34)+"0"+Chr$(34)+"," +Chr$(34)+"0"+Chr$(34)   '1", "0", "0"
      Print #2, "WCS.ActivateWCS " + Chr$(34)+"local"+Chr$(34)

skip_flag=False
While nextline<>"GEO"
 If Not EOF(1) Then
  Line Input #1, nextline
 Else
   MsgBox("No Geometry specified", "Warning")
   skip_flag= True
  Exit While
 End If
Wend
If Not skip_flag Then
 Do Until 	nextline="END GEO"  ' loop inside the Geo - section --------------------
 Line Input #1, nextline
 
  If (InStr(nextline,"TMET") > 0) Then
    nextline=  eliminate_blanks_in_strings (nextline$)
    metal_name_TMET = GetSubString(nextline,2," ")
     metal_type_TMET = GetSubString(nextline,4," ")
    Select Case metal_type_TMET
    Case "WGLOAD"
    Case "FREESPACE"
    	boundary_zmax="open"
    Case "NOR"
		T_metal_conductivity = Convert2Double(GetSubString(nextline,5," "))
		If T_metal_conductivity = 0 Then
		 'MsgBox "TMET: NOR Metal-conductivity for "+metal_name_TMET+ " = "+CStr(T_metal_conductivity)+vbCrLf+"set to 5.8e7 S/m", vbCritical
		 warning_text = warning_text + vbCrLf + "TMET: NOR Metal-conductivity for "+metal_name_TMET+ " = "+CStr(T_metal_conductivity)+vbCrLf+"set to 5.8e7 S/m"
		 T_metal_conductivity = 5.8e7
		End If
		TMET_thickness = Convert2Double(GetSubString(nextline,7," "))
		If TMET_thickness = 0 Then
		' MsgBox "Metal-thickness for "+metal_name_TMET+ " = "+CStr(TMET_thickness)+vbCrLf+ "set to "+ CStr(cst_level_thickness), vbCritical
		warning_text = warning_text + vbCrLf + "Metal-thickness for "+metal_name_TMET+ " = "+CStr(TMET_thickness)+vbCrLf+ "set to "+ CStr(cst_level_thickness)
		 TMET_thickness = cst_level_thickness
		End If
     	Print #2, "'@ define material: box_top_metal"
    	Define_layer_surf_imp  "box_top_"+replace_forbidden_characters(metal_name_TMET), T_metal_conductivity
    	TMET_exists = True	' a "True" creats a solid top-brick
    Case "RES","NAT","SUP","SEN"
		T_metal_conductivity = Convert2Double(GetSubString(nextline,5," "))	' reads resistance ...
		If T_metal_conductivity = 0 Then
		 'MsgBox "TMET: RES Metal-resistance for "+metal_name_TMET+ " = "+CStr(T_metal_conductivity)+vbCrLf+"set to PEC!", vbCritical
		 warning_text = warning_text + vbCrLf + "TMET: RES Metal-resistance for "+metal_name_TMET+ " = "+CStr(T_metal_conductivity)+vbCrLf+"set to PEC!"
		 boundary_zmax="electric"
		Else
		 TMET_thickness = cst_level_thickness
		 T_metal_conductivity= 1./(TMET_thickness*unit_scale*T_metal_conductivity)	'conductivity =
		boundary_zmax="conducting wall"
 		End If

    Case Else     
     MsgBox "No Appropriate Metal found for Top Cover Metal TMET", vbCritical
    End Select


  End If

If (InStr(nextline,"BMET") > 0) Then
    nextline=  eliminate_blanks_in_strings (nextline$)
    metal_name_BMET = GetSubString(nextline,2," ")
     metal_type_BMET = GetSubString(nextline,4," ")
    Select Case metal_type_BMET
    Case "WGLOAD"
    Case "FREESPACE"
    	boundary_zmin="open"
    Case "NOR"
		B_metal_conductivity = Convert2Double(GetSubString(nextline,5," "))
		If B_metal_conductivity = 0 Then
		 'MsgBox "BMET: NOR Metal-conductivity for "+metal_name_BMET+ " = "+CStr(B_metal_conductivity)+vbCrLf+"set to 5.8e7 S/m", vbCritical
			warning_text = warning_text + vbCrLf + "BMET: NOR Metal-conductivity for "+metal_name_BMET+ " = "+CStr(B_metal_conductivity)+vbCrLf+"set to 5.8e7 S/m"
		 B_metal_conductivity = 5.8e7
		End If
		BMET_thickness = Convert2Double(GetSubString(nextline,7," "))
		If BMET_thickness = 0 Then
		 MsgBox "Metal-thickness for "+metal_name_BMET+ " = "+CStr(BMET_thickness)+vbCrLf+ "set to "+ CStr(cst_level_thickness), vbCritical
		 warning_text = warning_text + vbCrLf + "Metal-thickness for "+metal_name_BMET+ " = "+CStr(BMET_thickness)+vbCrLf+ "set to "+ CStr(cst_level_thickness)
		 BMET_thickness = cst_level_thickness
		End If
     	Print #2, "'@ define material: box_bottom_metal"
    	Define_layer_surf_imp  "box_bottom_"+replace_forbidden_characters(metal_name_BMET), B_metal_conductivity
    	BMET_exists = True	' a "True" creats a solid top-brick
    Case "RES","NAT","SUP","SEN"
		B_metal_conductivity = Convert2Double(GetSubString(nextline,5," "))	' reads resistance ...
		If B_metal_conductivity = 0 Then
		 'MsgBox "BMET: RES Metal-resistance for "+metal_name_BMET+ " = "+CStr(B_metal_conductivity)+vbCrLf+"set to PEC!", vbCritical
			warning_text = warning_text + vbCrLf + "BMET: RES Metal-resistance for "+metal_name_BMET+ " = "+CStr(B_metal_conductivity)+vbCrLf+"set to PEC!"
		 boundary_zmin="electric"
		Else
		 BMET_thickness = cst_level_thickness
		 B_metal_conductivity= 1./(BMET_thickness*unit_scale*B_metal_conductivity)	'conductivity
		 boundary_zmin="conducting wall"
		End If

    Case Else
     MsgBox "No Appropriate Metal found for Bottom Cover Metal BMET", vbCritical
    End Select
  End If
'----
 If InStr(nextline,"MET") > 0 Then
   	nextline=  eliminate_blanks_in_strings (nextline)
   	If Not (GetSubString(nextline,1," ") = "BMET" Or  _
   	        GetSubString(nextline,1," ") = "TMET") Then 
    metal_name_MET = GetSubString(nextline,2," ")
	metal_type_MET = GetSubString(nextline,4," ")

	Select Case metal_type_MET
    Case "NOR","TMM"
        If GetSubString(nextline,5," ") = "INF" Then
			Print #2, "'@ define material: " + replace_forbidden_characters(metal_name_MET)
			Define_layer_PEC  replace_forbidden_characters(metal_name_MET)
        Else
		 metal_conductivity = Convert2Double(GetSubString(nextline,5," "))
		 If metal_conductivity = 0 Then
		  'MsgBox "MET: NOR Metal-conductivity for "+metal_name_MET+ " = "+CStr(metal_conductivity)+vbCrLf+"set to 5.8e7 S/m", vbCritical
		  warning_text = warning_text + vbCrLf + "MET: NOR Metal-conductivity for "+metal_name_MET+ " = "+CStr(metal_conductivity)+vbCrLf+"set to 5.8e7 S/m"
		  metal_conductivity = 5.8e7
		 End If

     	 Print #2, "'@ define material: " + replace_forbidden_characters(metal_name_MET)
    	 Define_layer_surf_imp  replace_forbidden_characters(metal_name_MET), metal_conductivity
		End If
    	MET_thickness = Convert2Double(GetSubString(nextline,7," "))
		If MET_thickness = 0 Then
		 'MsgBox "Metal-thickness for "+metal_name_MET+ " = "+CStr(MET_thickness)+vbCrLf+"set to "+ CStr(cst_level_thickness), vbCritical
		 warning_text = warning_text + vbCrLf + "Metal-thickness for "+metal_name_MET+ " = "+CStr(MET_thickness)+vbCrLf+"set to "+ CStr(cst_level_thickness)
		 MET_thickness = cst_level_thickness
		End If
    Case "RES","NAT","SUP","SEN"
		metal_conductivity = Convert2Double(GetSubString(nextline,5," "))	' reads resistance Ohms/square...
		If metal_conductivity = 0 Then
		 'MsgBox "MET: RES Metal-resistance for "+metal_name_MET+ " = "+CStr(metal_conductivity)+vbCrLf+"set to PEC!", vbCritical
			 warning_text = warning_text + vbCrLf + "MET: RES Metal-resistance for "+metal_name_MET+ " = "+CStr(metal_conductivity)+vbCrLf+"set to PEC!"
		 Print #2, "'@ define material: " + replace_forbidden_characters(metal_name_MET)
			Define_layer_PEC  replace_forbidden_characters(metal_name_MET)
		Else
		  If cst_level_thickness = 0 Then
			Print #2, "'@ define material: " + replace_forbidden_characters(metal_name_MET)
			Define_layer_PEC  replace_forbidden_characters(metal_name_MET)
		  Else
		   metal_conductivity= 1./(unit_scale*cst_level_thickness*metal_conductivity)	'conductivity= Ohms_square*thickness
		   Print #2, "'@ define material: " + replace_forbidden_characters(metal_name_MET)
		   Define_layer_surf_imp  replace_forbidden_characters(metal_name_MET), metal_conductivity
		  End If
		End If
		MET_thickness = cst_level_thickness
    Case Else
     'MsgBox "No Appropriate Metal found for Metal MET", vbCritical
		warning_text = warning_text + vbCrLf + "No Appropriate Metal found for Metal MET"
    End Select

    nr_of_diff_metals =nr_of_diff_metals + 1

    ReDim  Preserve metal_properties(nr_of_diff_metals)
    ReDim  Preserve metal_properties_Nr(nr_of_diff_metals)
    ReDim  Preserve metal_thickness(nr_of_diff_metals)
    metal_properties(nr_of_diff_metals-1)= metal_name_MET
    metal_properties_Nr(nr_of_diff_metals-1)= GetSubString(nextline,3," ")
	metal_thickness(nr_of_diff_metals-1) = MET_thickness

  End If
 End If
 '-------------------------
'----diel bricks
 If InStr(nextline,"BRI") > 0 Then
   	nextline=  eliminate_blanks_in_strings (nextline)
   	If Not (GetSubString(nextline,1," ")) = "BRI POL"  Then
    diel_name_MET = GetSubString(nextline,2," ")

	Print #2, "'@ define material: " + replace_forbidden_characters(diel_name_MET)
	Define_layer   replace_forbidden_characters(diel_name_MET), Convert2Double(GetSubString(nextline,4," ")),1.,Convert2Double(GetSubString(nextline,6," ")),  _
	                      Convert2Double(GetSubString(nextline,5," ")),0.,0.,0.,0.

    nr_of_diff_diel =nr_of_diff_diel + 1

    ReDim  Preserve diel_properties(nr_of_diff_diel)
    ReDim  Preserve diel_properties_Nr(nr_of_diff_diel)
    diel_properties(nr_of_diff_diel-1)= diel_name_MET
    diel_properties_Nr(nr_of_diff_diel-1)= GetSubString(nextline,3," ")	'material Nr.

  End If
 End If
 '-------------------------



 If (InStr(nextline,"BOX") > 0 And InStr(nextline, "SBOX") = 0) Then
    Nr_of_met_layers = Convert2Double(GetSubString(nextline,2," ")) +1  '(  + 0 level)
    Nr_of_diel_layers = Nr_of_met_layers '+ 1
    Nr_of_layer_positions = Nr_of_diel_layers + 1
    X_size = Convert2Double(GetSubString(nextline,3," "))
    Y_size = Convert2Double(GetSubString(nextline,4," ")) * NEG_sign
    smallest_diel_thickness = Abs(Y_size)  ' set the overall y-direction as largest mesh-distance
    ReDim diel_thickness(Nr_of_diel_layers) As Double
    ReDim diel_eps_r(Nr_of_diel_layers) As Double
    ReDim diel_eps_i(Nr_of_diel_layers) As Double
    ReDim diel_mu_r(Nr_of_diel_layers) As Double
    ReDim diel_mu_i(Nr_of_diel_layers) As Double
    ReDim diel_tan_d(Nr_of_diel_layers) As Double
    ReDim diel_name(Nr_of_diel_layers) As String
    ReDim layer_position(Nr_of_layer_positions) As Double
    ReDim diel_layer_position(Nr_of_layer_positions) As Double

    'layer_position(Nr_of_diel_layers) = 0 ' top position
    diel_layer_position(0) = 0 ' top position
    
   For run_thru_diel_index =  0 To Nr_of_diel_layers-1 'To 0 STEP -1
     Line Input #1, nextline
     nextline=  eliminate_blanks_in_strings (nextline)
    diel_thickness(run_thru_diel_index) = Convert2Double(GetSubString(nextline,1," "))
    'find the smallest substrate thick of layers
    If smallest_diel_thickness > diel_thickness(run_thru_diel_index) Then
		smallest_diel_thickness = diel_thickness(run_thru_diel_index)
    End If
    diel_eps_r(run_thru_diel_index) = Convert2Double(GetSubString(nextline,2," ")) 
    diel_eps_i(run_thru_diel_index) = Convert2Double(GetSubString(nextline,6," ")) 
    diel_mu_r(run_thru_diel_index) = Convert2Double(GetSubString(nextline,3," "))
    diel_mu_i(run_thru_diel_index) = Convert2Double(GetSubString(nextline,5," "))
    diel_tan_d(run_thru_diel_index) = Convert2Double(GetSubString(nextline,4," "))
    diel_name(run_thru_diel_index) = GetSubString(nextline,8," ")+CStr(run_thru_diel_index) 'append a number, in case nonunique diel-names
    
    'layer_position(run_thru_diel_index) = layer_position(run_thru_diel_index+1) - diel_thickness(run_thru_diel_index) '+
    diel_layer_position(run_thru_diel_index+1) = diel_layer_position(run_thru_diel_index) - diel_thickness(run_thru_diel_index)

    Print #2, "'@ new component: " + "Substrate_Layer_"+CStr(run_thru_diel_index)
    Print #2, "Component.New "+Chr$(34)+"Substrate_Layer_"+CStr(run_thru_diel_index)+Chr$(34)

    Print #2, "'@ new component: " + "Metal_Layer_"+CStr(run_thru_diel_index)
    Print #2, "Component.New "+Chr$(34)+"Metal_Layer_"+CStr(run_thru_diel_index)+Chr$(34)

    'Solid.MergeMaterialsOfComponent "Metal_Layer_2"
	'@ merge materials of component: VIA_Layer_0

    merge_metal= merge_metal+"'@ merge materials of component: " + "Metal_Layer_"+CStr(run_thru_diel_index)+ vbCrLf
    merge_metal=merge_metal+"Solid.MergeMaterialsofComponent "+Chr$(34)+"Metal_Layer_"+CStr(run_thru_diel_index)+Chr$(34)+vbCrLf

    Print #2, "'@ new component: " + "VIA_Layer_"+CStr(run_thru_diel_index)
    Print #2, "Component.New "+Chr$(34)+"VIA_Layer_"+CStr(run_thru_diel_index)+Chr$(34)

    merge_via=merge_via+"'@ merge materials of component: " + "VIA_Layer_"+CStr(run_thru_diel_index)+vbCrLf
	merge_via=merge_via+"Solid.MergeMaterialsofComponent "+Chr$(34)+"VIA_Layer_"+CStr(run_thru_diel_index)+Chr$(34)+vbCrLf
     
     Print #2, "'@ define material: " + replace_forbidden_characters(diel_name(run_thru_diel_index))
     Define_layer  replace_forbidden_characters(diel_name(run_thru_diel_index)), diel_eps_r(run_thru_diel_index),diel_mu_r(run_thru_diel_index),0, _
               diel_tan_d(run_thru_diel_index),(fmin+fmax)/2,diel_mu_i(run_thru_diel_index),0,0
     
     'create diele-boxes
     
     Print #2, "'@ define brick: " + "Substrate_Layer_" +CStr(run_thru_diel_index)+ ":Substrate_Layer_"+CStr(run_thru_diel_index) '?
     Print #2, "With Brick"
     Print #2, ".Reset" 
     Print #2, ".Name " +Chr$(34)+  "Substrate_Layer_"+CStr(run_thru_diel_index) +Chr$(34)
     Print #2, ".component " +Chr$(34)+  "Substrate_Layer_"+CStr(run_thru_diel_index) +Chr$(34)
     Print #2, ".material "  +Chr$(34)+ diel_name(run_thru_diel_index) +Chr$(34)
     Print #2, ".Xrange    0 " + " , " + evaluate(X_size)
     Print #2, ".Yrange    0 " + " , " + evaluate(Y_size)
     Print #2, ".Zrange "  + evaluate(diel_layer_position(run_thru_diel_index+1)) + " , " + evaluate(diel_layer_position(run_thru_diel_index))
     Print #2, ".Create"
	 Print #2, "End With"

	' set meshproperty higher than 0 (e.g. 2) so that each face gets a mesh-line
	' define automesh for: Substrate_Layer_3:Substarte_Layer_3
	' Solid.SetAutomeshParameters "Substrate_Layer_3:Substrate_Layer_3", "4711", "True"

		Print #2, "'@ define automesh for: "+ "Substrate_Layer_" +CStr(run_thru_diel_index)+ ":Substrate_Layer_"+CStr(run_thru_diel_index)
   		Print #2, "Solid.SetAutomeshParameters "+Chr$(34)+ "Substrate_Layer_" +CStr(run_thru_diel_index)+ ":Substrate_Layer_"+CStr(run_thru_diel_index)+  _
						Chr$(34)+ ", "+ Chr$(34)+"2"+ Chr$(34)+", "+ Chr$(34)+"True"+ Chr$(34)

        If Number_of_Mesh_Layers > 1 Then		'add fixpoints
 		actual_layer_thickness_fixpoint = Abs(diel_layer_position(run_thru_diel_index+1) -diel_layer_position(run_thru_diel_index))/ Number_of_Mesh_Layers
 			For index_for_fixpoints = 1 To Number_of_Mesh_Layers-1
				Print #2, "'@ new automesh fixpoint"
				Print #2, "Mesh.AddAutomeshFixpoint "+ Chr$(34)+"0"+Chr$(34)+","+ Chr$(34)+"0"+ Chr$(34)+ ","+ Chr$(34)+"1"+ Chr$(34)+ ","+  _
				              Chr$(34)+"0"+ Chr$(34)+ ","+Chr$(34)+"0"+ Chr$(34)+ ","+ Chr$(34)+   _
				                evaluate(diel_layer_position(run_thru_diel_index+1)+actual_layer_thickness_fixpoint*index_for_fixpoints)+Chr$(34)
    	    Next index_for_fixpoints

        End If

   Next run_thru_diel_index


	'index of met-layer starts with index 0 at diel_index =1; thus shift by one
	For run_thru_diel_index =  0 To Nr_of_diel_layers-1
		layer_position(run_thru_diel_index) =diel_layer_position(run_thru_diel_index+1)
	Next

	bottom_GND = (layer_position(Nr_of_diel_layers-1)) ' GND.... lowest dimenson in -Z at substrate(0)
    'MsgBox CStr(bottom_GND)
  End If
 
 If InStr(nextline,"NUM") > 0 Then									' start reading polygons
   	nextline=  eliminate_blanks_in_strings (nextline)
    Nr_of_polygons = Convert2Double(GetSubString(nextline,2," "))	' NUM #polygons 
    ReDim Polygon_array(Nr_of_polygons+1,max_nr_of_vertices,2)			' define Array to store all polygon's XY
    ReDim Polygon_index(Nr_of_polygons+1,2)			' define Array to store all polygon's index and #of vertices
    For cst_polygon_index = 1 To Nr_of_polygons	'+1					' loop thru all polygons fhi?
     Line Input #1, nextline
     nextline=  eliminate_blanks_in_strings (nextline)	
     poly_check_type = GetSubString(nextline,1," ")
     'MsgBox poly_check_type
     Select Case poly_check_type

      Case "BRI"	'dielectric brick
        Line Input #1, nextline	
      	poly_level = CInt(GetSubString(nextline,1," "))				' read level, #vertices, debug_nr, met-type
     	poly_nr_of_vertices = CInt(GetSubString(nextline,2," "))
     	poly_debug_Nr= CInt(GetSubString(nextline,5," "))
     	poly_met_type = CInt(GetSubString(nextline,3," "))				' 
     	ReDim X_poly(poly_nr_of_vertices)								' redimensioning XY-Points
     	ReDim Y_poly(poly_nr_of_vertices)
     	For cst_vertices_index = 1 To poly_nr_of_vertices				' loop thru individual polygons
      		Line Input #1, nextline
      		X_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,1," "))
      		Y_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,2," "))*NEG_sign
     	Next cst_vertices_index
     	Line Input #1, nextline		'read "END" line
     	' draw polygon+ extrude
     	check_duplicate_points X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr

     	If check_polygon_area X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr Then

			draw_polygon X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index,	poly_debug_Nr,"BRI_Polygon", 0
     		Print #2, "'@ new component: " + "Dielectric_Blocks"
        	Print #2, "Component.New "+Chr$(34)+ "Dielectric_Blocks" +Chr$(34)
        	create_extrude_diel X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index, poly_debug_Nr,"BRI_Polygon", _
                       poly_level+1,"Dielectric_Blocks"
        	Diel_Brick_exists= True

     	End If
      
      Case "VIA"	'via polygon
        Line Input #1, nextline	
      	poly_level = CInt(GetSubString(nextline,1," "))				' read level, #vertices, debug_nr, met-type
     	poly_nr_of_vertices = CInt(GetSubString(nextline,2," "))
     	poly_debug_Nr= CInt(GetSubString(nextline,5," "))
     	poly_met_type = CInt(GetSubString(nextline,3," "))				' 
     	ReDim X_poly(poly_nr_of_vertices)								' redimensioning XY-Points
     	ReDim Y_poly(poly_nr_of_vertices)
     	Line Input #1, nextline								'ToLevel-Line
     	'
		'jan 2009
		Select Case GetSubString(nextline,2," ")
		Case "GND"
     	'If	GetSubString(nextline,2," ") = "GND" Then	'TOLEVEL GND
     		to_poly_level = Nr_of_diel_layers-1 '  bottom_GND
     		'MsgBox CStr(poly_level) + " to " + CStr(to_poly_level)
     		'MsgBox "GND"+CStr(layer_position(poly_level)) + " to " + CStr(layer_position(to_poly_level))
     	Case "TOP"
     		to_poly_level = Nr_of_diel_layers '  TOLEVEL TOP
     		'MsgBox "TOP"+CStr(layer_position(poly_level)) + " to " + CStr(layer_position(to_poly_level))
     	Case Else										' normal layer number (0- ...etc
     		to_poly_level = CInt(GetSubString(nextline,2," "))
		End Select





     	For cst_vertices_index = 1 To poly_nr_of_vertices				' loop thru individual polygons
      		Line Input #1, nextline
      		X_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,1," "))
      		Y_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,2," ")) * NEG_sign
     	Next cst_vertices_index
     	Line Input #1, nextline		'read "END" line
     	' draw polygon+ extrude
     	check_duplicate_points X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr

     	If check_polygon_area X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr Then

     		draw_polygon X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index, _
     	             poly_debug_Nr,"VIA_Polygon", 0. '-cst_level_thickness/2 ' -<
	     	If dlg.cyl_vias =  0 Then ' create polygonal Vias
	         create_extrude X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index, poly_debug_Nr,"VIA_Polygon", _
	                       to_poly_level, "VIA_Layer_"
	        Else	' create cylindrical shaped Vias
			 create_extrude_cyl X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index, poly_debug_Nr,"VIA_Polygon", _
	                       to_poly_level, "VIA_Layer_"
	        End If
			VIA_exists = True

     	End If

 	  Case Else		'normal polygon
 	    
        poly_level = CInt(GetSubString(nextline,1," "))				' read level, #vertices, debug_nr, met-type
     	poly_nr_of_vertices = CInt(GetSubString(nextline,2," "))
     	poly_debug_Nr= CInt(GetSubString(nextline,5," "))
     	Polygon_index (counter_Normal_polygon,0) = poly_debug_Nr
     	Polygon_index (counter_Normal_polygon,1) = poly_level '
     	poly_met_type = CInt(GetSubString(nextline,3," "))				' 
     	ReDim X_poly(poly_nr_of_vertices)								' redimensioning XY-Points
     	ReDim Y_poly(poly_nr_of_vertices)
     	For cst_vertices_index = 1 To poly_nr_of_vertices				' loop thru individual polygons
      		Line Input #1, nextline

      		If IsNumeric(GetSubString(nextline,1," ")) Then
				X_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,1," "))
      			Y_poly(cst_vertices_index-1) = Convert2Double(GetSubString(nextline,2," ")) *NEG_sign

      			'round up numbers to a certain digit
				X_poly(cst_vertices_index-1) = (Int(X_poly(cst_vertices_index-1)*10^CInt(dlg.tolerance)+.5)/10^CInt(dlg.tolerance))
				Y_poly(cst_vertices_index-1) = (Int(Y_poly(cst_vertices_index-1)*10^CInt(dlg.tolerance)+.5)/10^CInt(dlg.tolerance))

      			Polygon_array(counter_Normal_polygon,cst_vertices_index, 0) = X_poly(cst_vertices_index-1)
      			Polygon_array(counter_Normal_polygon,cst_vertices_index, 1) = Y_poly(cst_vertices_index-1)
			Else
				cst_vertices_index = cst_vertices_index - 1
      		End If
      	Next cst_vertices_index

     	Line Input #1, nextline		'read "END" line
     	' draw polygon

     	material_index =0
			While CInt(metal_properties_Nr(material_index)) <> (poly_met_type+1)
            material_index = material_index+1
			If material_index >= UBound(metal_properties_Nr) Then
			    material_inconsistent=True
				'MsgBox "Polygon: Material for " + CStr(poly_met_type) + "-Type not found, PEC assumed!",vbCritical
				material_index=0
				Exit While
			End If
            Wend

    	check_duplicate_points X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr

    	If check_polygon_area X_poly(),Y_poly(),poly_nr_of_vertices,poly_debug_Nr Then
			draw_polygon X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index, _
     	             poly_debug_Nr,"Polygon", -metal_thickness(material_index)/2
	        create_extrude X_poly(),Y_poly(),poly_nr_of_vertices,poly_level,poly_met_type,cst_polygon_index,	poly_debug_Nr,"Polygon", _
	                        poly_level, "Metal_Layer_"
	        counter_Normal_polygon = counter_Normal_polygon+1
    	End If

      End Select
 
    Next cst_polygon_index
    
  
  End If
 
 ' extract edge vias
 
 If InStr(nextline,"EVIA1") > 0 Then
 	Line Input #1, nextline
   	nextline=  eliminate_blanks_in_strings (nextline)
	If GetSubString(nextline,1," ") = "POLY" Then
   	 poly_debug_Nr = CInt (GetSubString(nextline,2," ") ) 'POLY Nr
   	Else
   	 MsgBox " error in EVIA1 "+ nextline
   	End If
   	Line Input #1, nextline
   	evia_initial_vertex =	CInt(GetSubString(nextline,1," "))		'Start Vertex at polygon level
   	Line Input #1, nextline
   	If (GetSubString(nextline,1," ") = "TOLEVEL" ) Then
		If	GetSubString(nextline,2," ") = "TOP" Then
		 	poly_level =-1											'get the polygon target top level
        Else
  		   	poly_level = CInt(GetSubString(nextline,2," "))		'get the polygon target bottom level
		End If
    Else
     MsgBox " error in EVIA1 "+ nextline
    End If
    
    ReDim Preserve Evias_nr_array(Nr_of_evias)
    Evias_nr_array(Nr_of_evias)= poly_debug_Nr
    ReDim Preserve Evias_initvertex_array(Nr_of_evias)
    Evias_initvertex_array(Nr_of_evias)= evia_initial_vertex
    ReDim Preserve Evias_level_array(Nr_of_evias)
    Evias_level_array(Nr_of_evias)= poly_level
    
    Nr_of_evias = Nr_of_evias + 1
 	EVIA_exists =True
 End If
 
 
 'end loop
 
  Loop  'geo section ---------------------------------------------------------------------------
End If	'skip-flag
 	

' final beautifying and final measures
    

	 For i = 0 To nr_of_diff_metals-1 
	 Print #2, "'@ define material colour: " + metal_properties(i)
	 Print #2, "With material"
     Print #2, ".Name " + Chr$(34)   + metal_properties(i)+ Chr$(34)
     Print #2, ".Colour " + Chr$(34)  + evaluate(i/nr_of_diff_metals) +   _
                            Chr$(34) +" ,"" 0.6 "", "+  Chr$(34)+ evaluate(1-i/nr_of_diff_metals)+Chr$(34)

     Print #2, ".SetMaterialUnit " + Chr$(34)+ "GHz"+ Chr$(34) + ", "  + Chr$(34)+ "mm"+ Chr$(34)
     Print #2, ".Wireframe " + Chr$(34) +"False"+ Chr$(34) 
     Print #2, ".Transparency " + Chr$(34)+ "0"+ Chr$(34)
     Print #2, ".ChangeColour" 
	 Print #2," End With" 

	Next i
	
'--optimize mesh settings for planar structures
'        change mesh adaption scheme to energy 
' 		(planar structures tend to store high energy 
'     	 locally at edges rather than globally in volume)

If  InStr(getapplicationversion,"5.0.")  Then
'skip this command for version 5.0
Else	'MWS 5.1
		Print #2, "'@ Optimize Mesh for planar structures"
   		Print #2, "With Mesh"
   		Print #2, "     .MergeThinPECLayerFixpoints ""True"""
   		Print #2, "     .RatioLimit ""10"""
   		Print #2, "     .AutomeshRefineAtPecLines ""True"", ""2"""
   		Print #2, "     .UseRatioLimit ""False"""
   		Print #2, "     .LinesPerWavelength ""10"""
   		Print #2, "     .MinimumStepNumber ""10"""
   		Print #2, "     .AutoMesh ""True"""
   		Print #2, "     .SmallestMeshStep " + Chr$(34) + evaluate(smallest_diel_thickness)+"/"+evaluate(Number_of_Mesh_Layers) + Chr$(34) ' thickness of dielectric used;
   		Print #2, "End With"
   		Print #2, "MeshAdaption3D.SetAdaptionStrategy ""Energy"""
 End If

   		Print #2, "'@ define special solver parameters"
   		Print #2, "With Solver"
   		Print #2, "    .SetPortShielding  ""True"""
   		Print #2, "    .MaximumNumberOfProcessors ""2"""
   		Print #2, "    .UseTSTAtPort  ""False"""
   		Print #2, "End With"

 '--units and boundaries
	
		Print #2, "'@ define units"									'	'store in Variable und zum schluss Units 
   		Print #2, "Units.Geometry """+LCase(units_dim)+"""	
		Print #2, "Units.Frequency """+LCase(units_frq)+"""	

	  Print #2, "'@ define boundaries"  
	  If BMET_exists Then
	   Print #2, "boundary.zmin ""magnetic"
	  Else
	   Print #2, "boundary.zmin "+  Chr$(34)+boundary_zmin+Chr$(34)
	   If boundary_zmin = "conducting wall" Then
       	Print #2, "boundary.WallConductivity "+ Chr$(34)+ evaluate(B_metal_conductivity)  +Chr$(34)
       End If
	  End If
	  If TMET_exists Then
	   Print #2, "boundary.zmax ""magnetic"
	  Else
       Print #2, "boundary.zmax "+  Chr$(34)+boundary_zmax+Chr$(34)
       If boundary_zmax = "conducting wall" Then
       	Print #2, "boundary.WallConductivity "+ Chr$(34)+ evaluate(T_metal_conductivity)  +Chr$(34)
       End If
 	  End If
 		Print #2, "boundary.xmin ""open"
 		Print #2, "boundary.xmax ""open"
 		Print #2, "boundary.ymin ""open"
 		Print #2, "boundary.ymax ""open"


 		'background materials:
 		Print #2, "'@ define background "
        Print #2,"Background.Type "+Chr$(34)+"Normal"+Chr$(34)
		Print #2,"Background.Epsilon "+Chr$(34)+"1.0"+Chr$(34)
		Print #2,"Background.Mu "+Chr$(34)+"1.0"+Chr$(34)
		Print #2,"Background.XminSpace "+Chr$(34)+"0.0"+Chr$(34)
		Print #2,"Background.YminSpace "+Chr$(34)+"0.0"+Chr$(34)
		Print #2,"Background.ZminSpace "+Chr$(34)+"0.0"+Chr$(34)
		Print #2,"Background.XmaxSpace "+Chr$(34)+"0.0"+Chr$(34)
		Print #2,"Background.YmaxSpace "+Chr$(34)+"0.0"+Chr$(34)
		Print #2,"Background.ZmaxSpace "+Chr$(34)+"0.0"+Chr$(34)


 		
	'create bottom metalization
      If BMET_exists Then

        Print #2, "'@ activate global coordinates"
		Print #2, "WCS.ActivateWCS "+ Chr$(34)+"Global"+Chr$(34)

       	Print #2, "'@ new component: Bottom_metalization"
      	Print #2, "Component.New "+Chr$(34)+"Bottom_metalization"+Chr$(34)
     	Print #2, "'@ define brick: " +   "box_bottom_"+metal_name_BMET + " : "+ "Bottom_plate"
     	Print #2, "With Brick"
     	Print #2, ".Reset" 
     	Print #2, ".Name " + Chr$(34)+"Bottom_plate" +Chr$(34)'+ "box_bottom_" +metal_name_BMET +Chr$(34) 
     	Print #2, ".material "  + Chr$(34)+"box_bottom_" + metal_name_BMET +Chr$(34)
     	Print #2, ".Component "+Chr$(34)+"Bottom_metalization"+Chr$(34)
     	Print #2, ".Xrange    0 " + " , " + evaluate(X_size)
     	Print #2, ".Yrange    0 " + " , " + evaluate(Y_size)
    	'Print #2, ".Zrange "  + evaluate(layer_position(0) -BMET_thickness) + " , " + evaluate(layer_position(0))
		Print #2, ".Zrange "  + evaluate(bottom_GND -BMET_thickness) + " , " + evaluate(bottom_GND)

    	Print #2, ".Create"
	 	Print #2, "End With"
     End If

    'create top metalization
      If TMET_exists Then

      	Print #2, "'@ activate global coordinates"
		Print #2, "WCS.ActivateWCS "+ Chr$(34)+"Global"+Chr$(34)

      	Print #2, "'@ new component: Top_metalization"
      	Print #2, "Component.New "+Chr$(34)+"Top_metalization"+Chr$(34)
     	Print #2, "'@ define brick: " +   "box_top_"+metal_name_TMET + " : "+ "Top_plate"
     	Print #2, "With Brick"
     	Print #2, ".Reset" 
     	Print #2, ".Name " + Chr$(34)+"Top_plate" +Chr$(34)'+ "box_top_" +metal_name_TMET +Chr$(34) 
     	Print #2, ".material "   + Chr$(34)+"box_top_" + metal_name_TMET +Chr$(34)
     	Print #2, ".Component "+Chr$(34)+"Top_metalization"+Chr$(34)
     	Print #2, ".Xrange    0 " + " , " + evaluate(X_size)
     	Print #2, ".Yrange    0 " + " , " + evaluate(Y_size)
    	'Print #2, ".Zrange "  + evaluate(layer_position(Nr_of_layer_positions-1)) + " , " + _
    	'                        evaluate(layer_position(Nr_of_layer_positions-1)+TMET_thickness)
    	Print #2, ".Zrange "  + evaluate(layer_position(Nr_of_diel_layers-0)) + " , " + _
    	                        evaluate(layer_position(Nr_of_diel_layers-0)+TMET_thickness)
    	Print #2, ".Create"
	 	Print #2, "End With"
     End If
'-----     
'create edge-vias 
      If EVIA_exists Then

    Print #2, "'@ new component: EdgeVia"
    Print #2, "Component.New "+Chr$(34)+"EdgeVia"+Chr$(34)

	'@ merge materials of component: EdgeVia
	'Solid.MergeMaterialsOfComponent "EdgeVia"

	merge_evia = merge_evia + "'@ merge materials of component: EdgeVia"+ vbCrLf
	merge_evia = merge_evia +  "Solid.MergeMaterialsofComponent "+Chr$(34)+"EdgeVia"+Chr$(34)

        Print #2, "'@ define material: EdgeVia"
      	Define_layer_PEC "EdgeVia"
      
		'@ set wcs properties
      		Print #2,"'@ set wcs properties"
      		Print #2, "WCS.SetUVector " +"1,0,0"						'.SetUVector "1", "0", "0" 
			  Print #2, "WCS.SetNormal " +"0,0,1"				'.setnormal "0","0","1"
			  Print #2, "WCS.SetOrigin " +"0,0,0" 
			  
			  Print #2,"'@ activate local coordinates"
			  Print #2, "WCS.ActivateWCS " +Chr$(34)+"local"+ Chr$(34)'	.ActivateWCS "local"
			  

     ' create a curve-rectangle  in MWS
    For i = 0 To Nr_of_evias-1
    	 
   	 	For jj = 0 To Nr_of_polygons+1 '?
   	 	 If Polygon_index(jj,0) = Evias_nr_array(i) Then
   	 	 
   	 	 ex1= Polygon_array(jj,Evias_initvertex_array(i)+1,0)
   	 	 ey1= Polygon_array(jj,Evias_initvertex_array(i)+1,1)	
   	 	 ez1= layer_position(Polygon_index(jj,1))
   	 	 ex2= Polygon_array(jj,Evias_initvertex_array(i)+2,0)
   	 	 ey2= Polygon_array(jj,Evias_initvertex_array(i)+2,1)	
   	 	 ez2= layer_position(Polygon_index(jj,1))
   	 	 ex3= Polygon_array(jj,Evias_initvertex_array(i)+2,0)
   	 	 ey3= Polygon_array(jj,Evias_initvertex_array(i)+2,1)
   	 	 If Evias_level_array(i) < 0 Then		' "To LEVEL TOP"
           ez3 = diel_layer_position(0)  ' = 0
         Else
   	 	  ez3= layer_position(Evias_level_array(i))
		 End If
   	 	 ex4= Polygon_array(jj,Evias_initvertex_array(i)+1,0)
   	 	 ey4= Polygon_array(jj,Evias_initvertex_array(i)+1,1)
		 If Evias_level_array(i) < 0 Then	 '	"To LEVEL TOP"
           ez4 = diel_layer_position(0)	' = 0
         Else
   	 	   ez4= layer_position(Evias_level_array(i))
   	 	 End If
   	 	 
   	 	  '@ new curve: curve1
   	 	  Print #2, "'@ new curve: "+  "curve_"+CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))
   	 	  Print #2,"Curve.NewCurve "+ Chr$(34)+"curve_"+CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+Chr$(34)
   	 	 
   	 	 '@ define curve 3dpolygon: curve1:3dpolygon1
         Print #2,"'@ define curve 3dpolygon: "+"curve_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+":"+ _
                                      "3dpolygon"+ CStr( Evias_nr_array(i) )

     	Print #2,"Polygon3d.Reset" 
    	Print #2,"Polygon3d.Name " +Chr$(34)+"3dpolygon"+ CStr( Evias_nr_array(i) )+Chr$(34)
    	Print #2,"Polygon3D.Curve "+ Chr$(34)+"curve_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+ Chr$(34)
    	Print #2,"Polygon3D.Point "+evaluate(ex1) +", "+ evaluate(ey1)+", " +evaluate(ez1)  
    	Print #2,"Polygon3D.Point "+evaluate(ex2) +", "+ evaluate(ey2)+", " +evaluate(ez2)  
    	Print #2,"Polygon3D.Point "+evaluate(ex3) +", "+ evaluate(ey3)+", " +evaluate(ez3)  
    	Print #2,"Polygon3D.Point "+evaluate(ex4) +", "+ evaluate(ey4)+", " +evaluate(ez4)  
    	Print #2,"Polygon3D.Point "+evaluate(ex1) +", "+ evaluate(ey1)+", " +evaluate(ez1)  
   		Print #2,"Polygon3D.Create" 
		 
		'@ define coverprofile: Vacuum:solid3
         Print #2,"'@ define coverprofile: "+"EdgeVia:"+"E_Via_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))
         
 
   	 	 
     	Print #2,"CoverCurve.Reset" 
    	Print #2,"CoverCurve.Name "+Chr$(34)+ "E_Via_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+Chr$(34)
    	Print #2,"CoverCurve.Material "+Chr$(34)+ "EdgeVia"+ Chr$(34)
    	Print #2,"CoverCurve.Component "+Chr$(34)+ "EdgeVia"+ Chr$(34)
    	Print #2,"CoverCurve.Curve " +Chr$(34)+"curve_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+":"+ _
                                      "3dpolygon"+ CStr( Evias_nr_array(i) )+Chr$(34)
    	Print #2,"covercurve.Create"

		'delete curve
			Print #2,"'@ delete curve: "+"curve_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))'+Chr$(34)
	    Print #2,"Curve.deleteCurve " +Chr$(34)+"curve_"+ CStr( Evias_nr_array(i) )+"_"+CStr(Evias_initvertex_array(i))+Chr$(34)


   	 	 Else

   	 	 	End If
   	 	Next jj
   	 	'
    Next i
   End If     


'-----BOOL !!!!!!!!!!!!!!!!!!
'start Boolean operations to avoid overlappings....

Print #2,merge_metal
Print #2,merge_via
Print #2,merge_evia


'----close the mod-file ----------------------------------------------
Close #2

On Error GoTo 0


If material_inconsistent Then
	Info_text = Info_text + vbCrLf +"Material-Inconsistencies detected when creating Polygons, PEC assumed!"
End If
Info_text=Info_text+vbCrLf+"Nr.of Diel-Blocks: "+CStr(nr_of_diff_diel)
Info_text=Info_text+vbCrLf+"Nr.of Met.Layers : "+CStr(Nr_of_met_layers)
Info_text=Info_text+vbCrLf+"Nr.of Diel.Layers: "+CStr(Nr_of_diel_layers)
Info_text=Info_text+vbCrLf+"Nr.of Layer-Positions: "+CStr(Nr_of_layer_positions)
Info_text=Info_text+vbCrLf+"Nr.of Polygons: "+CStr(counter_Normal_polygon)
'Info_text=Info_text+vbCrLf+"Nr.of merged metal-layers: "+CStr(merged_metal_layers)
'Info_text=Info_text+vbCrLf+"Nr.of merged Via-layers: "+CStr(merged_via_layers)
'Info_text=Info_text+vbCrLf+"Nr.of merged EdgeVia-layers: "+CStr(merged_evia_layers)

If Len(duplicate_warning_text) > 0 Then
	MsgBox duplicate_warning_text,vbInformation,"Duplicate Points"
End If
If Len(not_created_polygon_warning_text) > 0 Then
	MsgBox not_created_polygon_warning_text,vbInformation,"Not Created Polygons"
End If
If Len(warning_text) > 0 Then
	MsgBox warning_text,vbInformation,"Information"
End If
MsgBox Info_text,vbInformation,"Summary"

Openfile cst_filestr_modexport	'open in MWS, it will probably ask for a backup
'MsgBox CStr(dlg.skip_bool)
If dlg.skip_bool=1 Then
   MsgBox "Sonnet Import completed "+vbCrLf+"(Costly Boolean operations skipped)",vbInformation
   Exit All
End If

screenupdating False

'-------open the mod-file for editing again, using AddToHistory instead
'obsolete Open cst_filestr_modexport For Append As #2 'output

'bool insert for each EdgeVia-Layer with all Metal_layers
nr_of_entries_polygons=0
nr_of_found_metal_layers=0
nr_of_found_normal_vias=0
nr_of_found_diel_blocks=0 '
nr_of_found_substrates=0	'

ReDim Entries_Metal_Layer(1) 'erase contents
get_all_child_names "Components\", All_Entries(), nr_of_All_entries		' get ALL entries
For run_thru_diel_index = 0 To nr_of_All_entries -1
	If InStr (All_Entries(run_thru_diel_index), "Metal_Layer_") > 0 Then		'find all beginning with "Metal_Layer_..."
		nr_of_found_metal_layers=nr_of_found_metal_layers+1
		ReDim Preserve 	Entries_Metal_Layer(nr_of_found_metal_layers)
		Entries_Metal_Layer(nr_of_found_metal_layers-1) =	All_Entries(run_thru_diel_index)
	End If
	'select Normal_Vias
	If InStr (All_Entries(run_thru_diel_index), "VIA_Layer_") > 0 Then		'find all beginning with "VIA_Layer_..."
		nr_of_found_normal_vias=nr_of_found_normal_vias+1
		ReDim Preserve Entries_Normal_Via_Layer(nr_of_found_normal_vias)
		Entries_Normal_Via_Layer(nr_of_found_normal_vias-1) =	All_Entries(run_thru_diel_index)
	End If
	'Diel Bricks:
	If InStr (All_Entries(run_thru_diel_index), "Dielectric_Blocks") > 0 Then		'find all beginning with "Diele_..."
		nr_of_found_diel_blocks=nr_of_found_diel_blocks+1
		ReDim Preserve Entries_Diel_blocks(nr_of_found_diel_blocks)
		Entries_Diel_blocks(nr_of_found_diel_blocks-1) =	All_Entries(run_thru_diel_index)
	End If
	'Substrates:
	If InStr (All_Entries(run_thru_diel_index), "Substrate_Layer_") > 0 Then		'find all beginning with "Substrate_Layer_..."
		nr_of_found_substrates=nr_of_found_substrates+1
		ReDim Preserve Entries_substrates(nr_of_found_substrates)
		Entries_substrates(nr_of_found_substrates-1) =	All_Entries(run_thru_diel_index)
	End If
Next


'-----*

For run_thru_diel_index = 0 To nr_of_found_metal_layers -1	'loop start ; loop thru all met layers
 For cst_II = 0 To nr_of_polygons_at_layer-1		'loop thru all polygon-entries at met.layer
  ReDim Preserve Entries_all_Polygons(nr_of_entries_polygons)
  Entries_all_Polygons(nr_of_entries_polygons)= Entries_Polygons_at_Layer(cst_II)	'add to a list containing all polygons
	Entries_all_Polygons(nr_of_entries_polygons) =  Entries_Metal_Layer(run_thru_diel_index) +":"+ Entries_all_Polygons(nr_of_entries_polygons)
	nr_of_entries_polygons=nr_of_entries_polygons+1
 Next cst_II
Next run_thru_diel_index

'-----*

get_all_child_names "Components\"+"EdgeVia\",Entries_Vias_at_Layer(),nr_of_vias_at_layer		'edge vias

For cst_JJ = 0 To nr_of_vias_at_layer-1			'loop thru all enries at the edgeVia-component
 Entries_Vias_at_Layer(cst_JJ) = "EdgeVia:"+ Entries_Vias_at_Layer(cst_JJ)

 'Perform all inserts of EdgeVias with all polygons
 For cst_II = 0 To  nr_of_entries_polygons-1		'loop thru all polygons at all metal-layers
  AddToHistory ("'@ boolean insert shapes: "+Entries_Vias_at_Layer(cst_JJ)+", "+ Entries_all_Polygons(cst_II) ,  _
              "Solid.Insert "+Chr$(34)+ Entries_Vias_at_Layer(cst_JJ)+Chr$(34)+", "+ Chr$(34)+  Entries_all_Polygons(cst_II)+Chr$(34) )
 Next cst_II
Next cst_JJ

For run_thru_diel_index = 0 To nr_of_found_metal_layers -1	'loop thru all met layers
  get_all_child_names "Components\"+ Entries_Metal_Layer(run_thru_diel_index)+"\",Entries_Polygons_at_Layer(),nr_of_polygons_at_layer
  'perform boolean inserts at individual layer-level:
   For cst_master=0 To nr_of_polygons_at_layer-1
  	cst_master_name = Entries_Metal_Layer(run_thru_diel_index) +":"+ Entries_Polygons_at_Layer(cst_master)
  	For cst_slave=cst_master+1 To nr_of_polygons_at_layer-1
		cst_slave_name = Entries_Metal_Layer(run_thru_diel_index) +":"+ Entries_Polygons_at_Layer(cst_slave)
		AddToHistory ("'@ boolean insert shapes: "+ cst_master_name+", "+ cst_slave_name	, _
						"Solid.Insert "+Chr$(34)+  cst_master_name+Chr$(34)+" ,  "+ Chr$(34)+  cst_slave_name +Chr$(34) )
  	Next cst_slave
  Next cst_master
  'end individual
Next run_thru_diel_index		'loop end

'----------------------------
'If Diel_Brick_exists:  Entries_Diel_blocks() ,nr_of_found_diel_blocks;  For Substrates:Entries_substrates(),nr_of_found_substrates
If Diel_Brick_exists Then
 For run_thru_diel_index = 0 To nr_of_found_diel_blocks-1
  get_all_child_names "Components\"+ Entries_Diel_blocks(run_thru_diel_index)+"\",Entries_diel_blocks_at_Layer(),nr_of_diel_blocks_at_layer
  For cst_II = 0 To nr_of_diel_blocks_at_layer-1
	'Entries_diel_blocks_at_Layer(cst_II) '... name of diel
	 For cst_JJ = 0 To nr_of_found_substrates-1
  		get_all_child_names "Components\"+ Entries_substrates(cst_JJ)+"\",Entries_substrates_at_Layer(),nr_of_substrates_at_layer
        For cst_KK = 0 To nr_of_substrates_at_layer-1
		'Entries_substrates_at_Layer(cst_kk)  '...name of substrate
		AddToHistory ("'@ boolean insert shapes: "+ Entries_substrates(cst_JJ)+":"+Entries_substrates_at_Layer(cst_KK)	+", "+ _
							 Entries_Diel_blocks(run_thru_diel_index)+":"+Entries_diel_blocks_at_Layer(cst_II), _
						"Solid.Insert "+Chr$(34)+ Entries_substrates(cst_JJ)+":" +Entries_substrates_at_Layer(cst_KK)+Chr$(34) +", "+ _
						Chr$(34)+Entries_Diel_blocks(run_thru_diel_index)+":"+ Entries_diel_blocks_at_Layer(cst_II)+Chr$(34)  )
        Next cst_KK
 	Next cst_JJ
   Next cst_II
 Next run_thru_diel_index
End If
'---...

'VIA_exists
If VIA_exists Then
 For run_thru_diel_index = 0 To nr_of_found_normal_vias -1	'loop thru all vias
  get_all_child_names "Components\"+ Entries_Normal_Via_Layer(run_thru_diel_index)+"\",Entries_Normal_Vias_at_Layer(),nr_of_Normal_vias_at_layer
   For cst_JJ = 0 To nr_of_Normal_vias_at_layer-1
	For cst_II = 0 To  nr_of_entries_polygons-1		'loop thru all polygons at all metal-layers
     AddToHistory ("'@ boolean insert shapes: "+Entries_Normal_Via_Layer(run_thru_diel_index)+":"+Entries_Normal_Vias_at_Layer(cst_JJ)+", "+ Entries_all_Polygons(cst_II) , _
  		"Solid.Insert "+Chr$(34)+Entries_Normal_Via_Layer(run_thru_diel_index)+":"+ Entries_Normal_Vias_at_Layer(cst_JJ)+Chr$(34)+", "+ Chr$(34)+  Entries_all_Polygons(cst_II)+Chr$(34) )
 	Next cst_II
   Next cst_JJ
 Next run_thru_diel_index
End If

' end Boolean operations
'------------------------

nofile:
nofile_export:

On Error GoTo 0

MsgBox "Sonnet Import completed",vbInformation
screenupdating True

End Sub	' of main()


 Sub check_duplicate_points (X_poly() As Double,Y_poly() As Double, poly_nr_of_vertices As Integer,poly_debug_Nr As Integer)
	Dim cst_vertices_index As Integer, polyvertex_temp_counter As Integer, X_temp() As Double, Y_temp() As Double
	Dim poly_nr_tolerance As Double,  cst_counter_equal As Integer, X_temp1 As Double, Y_temp1 As Double
	Dim display_error_message_once As Boolean
	display_error_message_once = True

	poly_nr_tolerance = 1/10^CInt(tolerance_digits)	'was fixed to 1.e-3 in prev. version !!!

	ReDim X_temp(poly_nr_of_vertices)
	ReDim Y_temp(poly_nr_of_vertices)

	X_temp1 = X_poly(0)
	Y_temp1 = Y_poly(0)
	X_temp(0) = X_poly(0)	'save always 1st point into array;
	Y_temp(0) = Y_poly(0)
	polyvertex_temp_counter = 0
	For cst_vertices_index = 1 To poly_nr_of_vertices-1				' loop thru xy points
      	If Abs(Abs(X_poly(cst_vertices_index))-Abs(X_temp1)) > poly_nr_tolerance Or _
		   Abs(Abs(Y_poly(cst_vertices_index))-Abs(Y_temp1)) > poly_nr_tolerance Then	'Points unequal
				polyvertex_temp_counter = polyvertex_temp_counter +1
      			X_temp(polyvertex_temp_counter)=X_poly(cst_vertices_index)
				Y_temp(polyvertex_temp_counter)=Y_poly(cst_vertices_index)
				X_temp1 = X_poly(cst_vertices_index)
				Y_temp1 = Y_poly(cst_vertices_index)
        Else	'equal points found, ignore this one !! is a tol.-problem !!!error message ->
        		If display_error_message_once Then
				'MsgBox "Polygon "+ CStr(poly_debug_Nr)+ " shows tolerance problems !"
				duplicate_warning_text= duplicate_warning_text +vbCrLf+ "Polygon "+ CStr(poly_debug_Nr)+ " shows duplicate points (within tolerance of "+ _
				                                                        CStr(poly_nr_tolerance)+")  !" + " or not a single path!"
				display_error_message_once = False
				End If
        End If
    Next
	'check last found point and 1st point of the closed polygon
    If Abs(Abs(X_poly(poly_nr_of_vertices-1))-Abs(X_poly(0))) > poly_nr_tolerance Or _
	   Abs(Abs(Y_poly(poly_nr_of_vertices-1))-Abs(Y_poly(0))) > poly_nr_tolerance Then	'last two Points unequal
       'Do Nothing, but is a serious error -> polygon not closed!!!!!!!!!
    	MsgBox "Polygon "+ CStr(poly_debug_Nr)+ " is not closed ! Conversion aborted"
       Exit All
    Else ' Save the first XY-point rather than the second last !
		X_temp(polyvertex_temp_counter) = X_temp(0)
		Y_temp(polyvertex_temp_counter) = Y_temp(0)
    End If

    For cst_counter_equal = 0 To polyvertex_temp_counter+1		'save back remaining unique points
        X_poly(cst_counter_equal) = X_temp(cst_counter_equal)
		Y_poly(cst_counter_equal) = Y_temp(cst_counter_equal)
    Next cst_counter_equal

    poly_nr_of_vertices =polyvertex_temp_counter+1


End Sub

Function check_polygon_area (X_poly() As Double,Y_poly() As Double ,poly_nr_of_vertices As Integer,poly_debug_Nr As Integer) As Boolean

	check_polygon_area = True

	If Abs(area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)) > 0 Then
		check_polygon_area = True
	Else
		check_polygon_area = False
		not_created_polygon_warning_text= not_created_polygon_warning_text +vbCrLf+ "Polygon "+ CStr(poly_debug_Nr)+ " could not be created because it has zero area!"
 	End If

End Function

Function replace_forbidden_characters (name_to_check As String) As String

	replace_forbidden_characters = name_to_check

	replace_forbidden_characters = Replace(replace_forbidden_characters, "\\", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "*", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, ":", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "|", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "/", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "$", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "[", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "]", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "~", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, ",", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "<", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, ">", "_")
	replace_forbidden_characters = Replace(replace_forbidden_characters, "?", "_")

End Function


 Sub draw_polygon (X_poly() As Double,Y_poly() As Double ,poly_nr_of_vertices As Integer, poly_level As Integer, poly_met_type As Integer, _
                     cst_polygon_index As Integer, cst_polygon_Nr As Integer,polygon_name As String, thickness_offset As Double )

 Dim jj As Integer, cst_area As Double, material_index As Integer
	'define curve first for later extrusion !!!!!!!!!!!!!!!!
			'With WCS
			 'Print #2, "WCS.SetNormal " +"0,0,1"' +Chr$(34)  			'.SetNormal "0", "0", "1"
    		 'Print #2, "WCS.SetOrigin " +"0,0,0"  						'.SetOrigin "0", "0", "0"
    		 'Print #2, "WCS.SetUVector " +"1,0,0"						'.SetUVector "1", "0", "0"
    		 Print #2, "'@ activate local coordinates"
     		 Print #2, "WCS.ActivateWCS " +Chr$(34)+"local"+ Chr$(34)'	.ActivateWCS "local" 
			'End With

			'material_index =0
			'While CInt(metal_properties_Nr(material_index)) <> (poly_met_type+1)
            'material_index = material_index+1
			'If material_index >= UBound(metal_properties_Nr) Then
			 '   material_inconsistent=True
			'	'MsgBox "Polygon: Material for " + CStr(poly_met_type) + "-Type not found, PEC assumed!",vbCritical
			'	material_index=0
			'	Exit While
			'End If
            'Wend
            'new:::
            'thickness_offset = (-metal_thickness(material_index)/2)

			'With WCS
			Print #2, "'@ set wcs properties"
			  Print #2, "WCS.SetNormal " +"0,0,1"				'.setnormal "0","0","1"
			  Print #2, "WCS.SetOrigin " +"0,0,"+ evaluate(layer_position(poly_level) +thickness_offset)   '?
			  Print #2, "WCS.SetUVector " +"1,0,0"				'.setuvector "1","0","0"
			'End With
			
			Print #2, "'@ new curve: " + polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))	 
			Print #2,"Curve.newcurve "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)	
			'With Polygon
			 Print #2, "'@ define curve polygon: "+ polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+":"+ polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
			 Print #2,"polygon.Reset"
			 Print #2,"polygon.Name "+ Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			 Print #2,"polygon.curve "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			 		'first point				 
				Print #2,"polygon.Point "+ evaluate(X_poly(0))+","+evaluate(Y_poly(0))  
				' connections
				For jj=1 To poly_nr_of_vertices-1	 
				 
				Print #2,"polygon.LineTo "+ evaluate(X_poly(jj) )+","+evaluate( Y_poly(jj))  
				Next
				Print #2,"polygon.Create"		'OK
			 	If Abs(area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)) > 0 Then
                  cst_area = 	area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)			

				Else
				Print #10, "Could not create Via because area = 0!"
				For jj=0 To poly_nr_of_vertices-1
				 Print #10, CStr(X_poly(jj))+ " "+CStr(Y_poly(jj))
				Next
				Print #10," "
				'GoTo jump_over
				End If


	End Sub
Sub create_extrude (X_poly() As Double,Y_poly() As Double ,poly_nr_of_vertices As Integer, poly_level As Integer, poly_met_type As Integer, _
                     cst_polygon_index As Integer, cst_polygon_Nr As Integer,polygon_name As String, to_poly_level As Integer, layer_name As String)
			'---- create extrude, read geometry directly into extrude definittion
			Dim cst_area As Double, jj As Integer, material_index As Integer
			'layer_name = "Metal_Layer_"
			cst_area = 	area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)
			Print #2,"'@ define extrudeprofile: "+ layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
			Print #2,"ExtrudeCurve.Reset"
			Print #2,"ExtrudeCurve.Name "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			Print #2,"ExtrudeCurve.component "+Chr$(34)+layer_name+CStr(poly_level)+Chr$(34)
			'search for metal-type:
			material_index =0
			While CInt(metal_properties_Nr(material_index)) <> (poly_met_type+1)
            material_index = material_index+1
			If material_index >= UBound(metal_properties_Nr) Then
			'	MsgBox "Material for " + CStr(poly_met_type) + "-Type not found, PEC assumed!",vbCritical
				material_index=0
				Exit While
			End If
            Wend
			Print #2,"ExtrudeCurve.material "+Chr$(34)+metal_properties(material_index)+Chr$(34) '"sonnet_level_"+LTrim(CStr(poly_level))
			If (layer_position(to_poly_level)- layer_position(poly_level)) <> 0 Then' extrusion extends over more levels
			'If (diel_layer_position(to_poly_level+1)- diel_layer_position(poly_level+1)) <> 0 Then' extrusion extends over more levels
			  Print #2,"ExtrudeCurve.thickness "+Replace(evaluate( Sgn(cst_area)* (layer_position(to_poly_level+1)- layer_position(poly_level+1))   ),",",".")
			Else	'only 1 level
				Print #2,"ExtrudeCurve.thickness "+Replace(evaluate(Sgn(cst_area)*metal_thickness(material_index)) ,",",".")
			End If
			Print #2,"ExtrudeCurve.twistangle 0" 
			Print #2,"ExtrudeCurve.taperangle 0" ' 1, 0, 0
			Print #2,"ExtrudeCurve.curve "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			Print #2,"ExtrudeCurve.create"

			If metal_thickness(material_index) = 0 Then			'set Polygon to staircase mesh if thickness = 0:
 			'	Print #2,"'@ define automesh For: "+  layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
 			'	Print #2,"Solid.SetMeshProperties "+ Chr$(34)+  layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr)) +  _
 			'	                       Chr$(34)+", "+Chr$(34)+ "Staircase"+ Chr$(34)+ ", "+Chr$(34)+"False"+ Chr$(34)
			End If


' start ube 12.Jan
			If (Left(layer_name,3)="VIA") And (poly_nr_of_vertices > 5) Then

				' switch-off from automesh and create a wire in the center for the fix-points

				Print #2,"'@ switch off from automesh: "+ layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
				Print #2,"Solid.SetAutomeshParameters """+ layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+""", ""0"", ""False""

				' define a 3d-wire in the center to create the fixpoints

				Dim x_mean_tmp As Double, y_mean_tmp As Double
				x_mean_tmp = 0.0
				y_mean_tmp = 0.0

				' start loop at 1 (not at 0), otherwise start and end point is considered twice !!!
				For jj=1 To poly_nr_of_vertices-1
					x_mean_tmp = x_mean_tmp + X_poly(jj)
					y_mean_tmp = y_mean_tmp + Y_poly(jj)
				Next

				If (poly_nr_of_vertices > 1) Then
					x_mean_tmp = x_mean_tmp / (poly_nr_of_vertices-1)
					y_mean_tmp = y_mean_tmp / (poly_nr_of_vertices-1)
				End If

				Print #2,"'@ activate global coordinates"
				Print #2,"      WCS.ActivateWCS ""Global"""
				Print #2,""
				Print #2,"'@ define curve for via-wire"
				Print #2,"With Polygon3D"
				Print #2,"     .Reset"
				Print #2,"     .Name ""3dpolygon_via"""
				Print #2,"     .Curve """+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+"""
				Print #2,"     .Point """+ evaluate(x_mean_tmp)+ """, """+ evaluate(y_mean_tmp)+ """, """ + evaluate(layer_position(poly_level))+"""
				Print #2,"     .Point """+ evaluate(x_mean_tmp)+ """, """+ evaluate(y_mean_tmp)+ """, """ + evaluate(layer_position(to_poly_level))+"""
				Print #2,"     .Create"
				Print #2,"End With"
				Print #2,""
				Print #2,"'@ define curvewire: Wire_"+ LTrim(CStr(cst_polygon_Nr))
				Print #2,""
				Print #2,"With Wire"
				Print #2,"     .Reset"
				Print #2,"     .Name ""Wire_" + LTrim(CStr(cst_polygon_Nr)) +"""
				Print #2,"     .Radius ""0.0"""
				Print #2,"     .Type ""CurveWire"""
				Print #2,"     .Curve """ + polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+":3dpolygon_via"""
				Print #2,"     .Add"
				Print #2,"End With"
				Print #2,""

			End If
' end ube 12.Jan

			Print #2,"'@ delete curve: "+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
	    	Print #2,"Curve.deleteCurve " +Chr$(34)+  polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))  +Chr$(34)

			GoTo 	jump_over
		    cantdoit:	
		    On Error GoTo 0
			Print #10, "Could not create Extrusion because of self-intersection!"
			For jj=0 To poly_nr_of_vertices-1
				 Print #10, CStr(X_poly(jj))+ " "+CStr(Y_poly(jj))
				Next
				Print #10," ---------------------- "
				Print #10," "
				
			    jump_over:
			
			End Sub

Sub create_extrude_cyl (X_poly() As Double,Y_poly() As Double ,poly_nr_of_vertices As Integer, poly_level As Integer, poly_met_type As Integer, _
                     cst_polygon_index As Integer, cst_polygon_Nr As Integer,polygon_name As String, to_poly_level As Integer, layer_name As String)
			'---- create cylinders!, read geometry Points and dreive center and radius of cylinder
			Dim cst_area As Double, jj As Integer, material_index As Integer
			'layer_name = "Metal_Layer_"
			cst_area = 	area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)
			'search for metal-type:
			material_index =0
			While CInt(metal_properties_Nr(material_index)) <> (poly_met_type+1)
            material_index = material_index+1
			If material_index >= UBound(metal_properties_Nr) Then
			'	MsgBox "Material for " + CStr(poly_met_type) + "-Type not found, PEC assumed!",vbCritical
				material_index=0
				Exit While
			End If
            Wend

			If (Left(layer_name,3)="VIA") And (poly_nr_of_vertices > 3) Then

				Dim x_mean_tmp As Double, y_mean_tmp As Double, cyl_radius As Double
				x_mean_tmp = 0.0
				y_mean_tmp = 0.0

				' start loop at 1 (not at 0), otherwise start and end point is considered twice !!!
				For jj=1 To poly_nr_of_vertices-1
					x_mean_tmp = x_mean_tmp + X_poly(jj)
					y_mean_tmp = y_mean_tmp + Y_poly(jj)
				Next

				If (poly_nr_of_vertices > 1) Then
					x_mean_tmp = x_mean_tmp / (poly_nr_of_vertices-1)
					y_mean_tmp = y_mean_tmp / (poly_nr_of_vertices-1)
				End If
				' radius = distance (vertex to center) multiplied by 1/2 of cos(angle_to_next_vertex): inner_circle of polygon
				cyl_radius = Sqr(Abs(x_mean_tmp -X_poly(0))^2+Abs(y_mean_tmp -Y_poly(0))^2)*Cos(pi/(poly_nr_of_vertices-1))

				Print #2,"'@ activate global coordinates"
				Print #2,"      WCS.ActivateWCS ""Global"""
				Print #2,""
				Print #2,"'@ define cylinder: "+layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
				Print #2,"With Cylinder"
				Print #2,"   .Reset"
				Print #2,"   .Name "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
				Print #2,"   .component "+Chr$(34)+layer_name+CStr(poly_level)+Chr$(34)
				Print #2,"   .material "+Chr$(34)+metal_properties(material_index)+Chr$(34)
				Print #2,"     .InnerRadius ""0.0"""
				Print #2,"     .Axis ""z"""
				Print #2,"     .Segments ""0"""
				Print #2,"     .XCenter """+ evaluate(x_mean_tmp)+ """
				Print #2,"     .YCenter """+ evaluate(y_mean_tmp)+ """
			 If (layer_position(to_poly_level)- layer_position(poly_level)) <> 0 Then' extrusion extends over more levels
			    If layer_position(poly_level)  <  layer_position(to_poly_level) Then 'direction of cyl
			  	Print #2,"     .ZRange "+Chr$(34)+Replace(evaluate( layer_position(poly_level)),",",".")    +Chr$(34) +", "+  _
									     Chr$(34)+Replace(evaluate( layer_position(to_poly_level)),",",".") +Chr$(34)
				Else
					Print #2,"     .ZRange "+Chr$(34)+Replace(evaluate( layer_position(to_poly_level)),",",".")    +Chr$(34) +", "+  _
									     Chr$(34)+Replace(evaluate( layer_position(poly_level)),",",".") +Chr$(34)
				End If

		   	 Else	'only 1 level
				If layer_position(poly_level)  <  layer_position(to_poly_level) Then	'direction of cyl
				Print #2,"     .ZRange "+Chr$(34)+Replace(evaluate(layer_position(poly_level)),",",".") +Chr$(34)+", "+  _
									     Chr$(34)+Replace(evaluate(layer_position(poly_level)+metal_thickness(material_index)),",",".") +Chr$(34)
				Else
				Print #2,"     .ZRange "+Chr$(34)+Replace(evaluate(layer_position(poly_level)+metal_thickness(material_index)),",",".") +Chr$(34)+", "+  _
									     Chr$(34)+Replace(evaluate(layer_position(poly_level)),",",".") +Chr$(34)
				End If
			 End If
				Print #2,"     .OuterRadius """+ evaluate(cyl_radius)+ """
				Print #2,"     .Create"
				Print #2,"End With"
				Print #2,""
				' set priority to 1, moves meshline to center of cylinder
				Print #2,"'@ define automesh For: "+  layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
 				Print #2,"Solid.SetAutomeshParameters "+ Chr$(34)+  layer_name+CStr(poly_level) +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr)) +  _
 				                       Chr$(34)+", "+Chr$(34)+ "1"+ Chr$(34)+ ", "+Chr$(34)+"True"+ Chr$(34)
			End If

			Print #2,"'@ delete curve: "+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
	    	Print #2,"Curve.deleteCurve " +Chr$(34)+  polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))  +Chr$(34)

			GoTo 	jump_over
		    cantdoit:
		    On Error GoTo 0
			Print #10, "Could not create Extrusion because of self-intersection!"
			For jj=0 To poly_nr_of_vertices-1
				 Print #10, CStr(X_poly(jj))+ " "+CStr(Y_poly(jj))
				Next
				Print #10," ---------------------- "
				Print #10," "

			    jump_over:

			End Sub


Sub create_extrude_diel (X_poly() As Double,Y_poly() As Double ,poly_nr_of_vertices As Integer, poly_level As Integer, poly_met_type As Integer, _
                     cst_polygon_index As Integer, cst_polygon_Nr As Integer,polygon_name As String, to_poly_level As Integer, layer_name As String)
			'---- create extrude, read geometry directly into extrude definittion
			Dim cst_area As Double, jj As Integer, material_index As Integer
			'layer_name = "Metal_Layer_"
			cst_area = 	area_gauss(X_poly(),Y_poly(),poly_nr_of_vertices)
			Print #2,"'@ define extrudeprofile: "+ layer_name +":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
			Print #2,"ExtrudeCurve.Reset"
			Print #2,"ExtrudeCurve.Name "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			Print #2,"ExtrudeCurve.component "+Chr$(34)+layer_name+Chr$(34)
			'search for metal-type:
			material_index =0
			While CInt(diel_properties_Nr(material_index)) <> (poly_met_type)' indices equal....
            material_index = material_index+1
			If material_index >= UBound(diel_properties_Nr) Then
				MsgBox "Material for " + CStr(poly_met_type) + "-Type not found, Vacuum assumed!",vbCritical
				material_index=0
				Exit While
			End If
            Wend
			Print #2,"ExtrudeCurve.material "+Chr$(34)+diel_properties(material_index)+Chr$(34) '"sonnet_level_"+LTrim(CStr(poly_level))
			If (layer_position(to_poly_level)- layer_position(poly_level)) <> 0 Then' extrusion extends over more levels
			  Print #2,"ExtrudeCurve.thickness "+Replace(evaluate( Sgn(cst_area)* (layer_position(to_poly_level)- layer_position(poly_level))   ),",",".")
			Else
				Print #2,"ExtrudeCurve.thickness "+Replace(evaluate(Sgn(cst_area)*cst_level_thickness) ,",",".") 	' only met-thickness (+ diection) hier sign(area!)checken
			End If
			Print #2,"ExtrudeCurve.twistangle 0"
			Print #2,"ExtrudeCurve.taperangle 0" ' 1, 0, 0
			Print #2,"ExtrudeCurve.curve "+Chr$(34)+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+":"+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))+Chr$(34)
			Print #2,"ExtrudeCurve.create"
			Print #2,"'@ delete curve: "+polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))
	    	Print #2,"Curve.deleteCurve " +Chr$(34)+  polygon_name+"_"+LTrim(CStr(cst_polygon_Nr))  +Chr$(34)

			GoTo 	jump_over
		    cantdoit:
		    On Error GoTo 0
			Print #10, "Could not create Extrusion because of self-intersection!"
			For jj=0 To poly_nr_of_vertices-1
				 Print #10, CStr(X_poly(jj))+ " "+CStr(Y_poly(jj))
				Next
				Print #10," ---------------------- "
				Print #10," "

			    jump_over:

			End Sub





'---------------------------------------------------------------------------------

Function GetString(lin As String) As String

	Dim index As Integer
	Dim svalue As String
	
	lin=LTrim(lin)
	index=InStr(lin, " ")
	
	If index=0 Then
		GetString=lin
	Else 
		GetString=Left(lin,index)
	End If
	
	lin=Mid(lin, index+1)
	
End Function





'------------------------------------------------------------------

Function AngleSum(P() As Double, pointlist() As Double, npoints As Integer) As Double

	Dim x As Double
	Dim y As Double
	Dim phisum As Double
	Dim x1, x2, y1, y2, phi1, phi2 As Double
	Dim II As Integer
	
	x=P(0)
	y=P(1)
	
	phisum=0
	
	For II=1 To npoints-1
		x1=pointlist(II-1, 0)-x
		x2=pointlist(II, 0)-x
		y1=pointlist(II-1, 1)-y
		y2=pointlist(II, 1)-y
		
		phi1=MyAtn(RealValx(x1), RealValx(y1))
		phi2=MyAtn(RealValx(x2), RealValx(y2))
		
		phisum=phisum+phi2-phi1
	
	Next
	
	AngleSum=phisum
	
End	Function

'----------------------------------------------------------------------------
	
Function MyAtn(x As Double, y As Double) As Double

Dim phi As Double
	
	If x=0 Then
		
		phi=Pi*Sgn(y)
		
	Else
		phi=Atn(y/x)
		If x<0 Then
			phi=Pi+phi
		End If
	End If
	
	If phi<0 Then
		phi=2*PI+phi
	End If
	
	MyAtn=phi
	
End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function RealValx(lib_Text As Variant) As Double
       RealValx  =    CDbl(evaluate(lib_Text))'   evaluate(Format(lib_Text,"#.00")) ...falsch ! 2 Stellen werden abgeschnitten
End Function


Function evaluate(xvalue As Variant) As String
	Dim index As Long, evaluate1 As String
	index = InStrRev(xvalue, "+")
	If (index = 0) Then index = InStrRev(xvalue, "-")
	If (index > 1) Then
		If (Mid(xvalue, index-1, 1) <> "E" And Mid(xvalue, index-1, 1) <> "e") Then
			xvalue = Left(xvalue, index-1) + "E" + Mid(xvalue, index)
		End If
	End If

	Dim cst_separator    As String								'<<<<<<<<<<
	cst_separator    = Mid(CStr(0.5), 2, 1)					'<<<<<<<<<<<
	evaluate1        = (Replace(xvalue, ",", cst_separator,1,1))	'<<<<<<<<<<<
	evaluate        = (Replace(evaluate1, ".", cst_separator,1,1))

End Function
'----------------

Function area_gauss (x() As Double, y() As Double, n As Integer) As Double
Dim i As Integer, j As Integer
 area_gauss = 0.
 For i = 0 To n-1
  If i=n-1 Then 
   j=0
  Else
   j=i+1
  End If
  area_gauss = area_gauss +0.5*(x(i)-x(j))*(y(i)+y(j))
 Next
End Function
Function BaseName (path As String) As String

        Dim dircount As Integer, extcount As Integer, filename As String

        dircount = InStrRev(path, "\")
        filename = Mid$(path, dircount+1)
        extcount = InStrRev(filename, ".")
        BaseName = Left$(filename, IIf(extcount > 0, extcount-1, 999))

End Function
Function ShortName (lib_path As String) As String

        Dim lib_dircount As Integer

        lib_dircount  = InStrRev(lib_path, "\")
        ShortName = Mid$(lib_path, lib_dircount+1, 999)

End Function

Function ExtName (filename As String) As String

	Dim extcount As Integer

	extcount = InStrRev(filename, ".")
	ExtName = IIf(extcount > 0, Mid$(filename, extcount+1), "")

End Function



 Function GetSubString(ByVal sline As String, str_GrepIndex As Integer, delimiter As String) As String
' sline ... input string   		example : "ABC 345 ""un_known"" END 8.88"
' str_GrepIndex ... item-nr of input-string		example : 3
' delimiter  ... character used as a delimiter  example " "
'  result: GetSubString -> example : "un_known"

	Dim index As Integer
	Dim svalue As String
	Dim i As Integer
	Dim lib_count As Integer, loop_counter As Integer
	 sline=LTrim(sline)
	lib_count=0
    loop_counter =  0
    While lib_count < str_grepindex-1
     loop_counter = loop_counter+1
	 If (loop_counter > str_GrepIndex) Then			
	  Exit Function
	 End If	
	 sline=LTrim(sline)
	 If InStr(sline,delimiter) Then
	  lib_count = lib_count + 1
	  index = InStr(sline,delimiter)
	  sline=LTrim(Mid(sline,index))
	 End If
	Wend
	index = InStr(sline,delimiter)
	If index > 0 Then
	 GetSubString=LTrim(Left(sline,index))
	 GetSubString=RTrim(Left(GetSubString,index))
	Else
	 GetSubString = sline
	End If	
End Function

Function eliminate_blanks_in_strings (ByVal sline As String) As String
' replaces blanks by underscore
' example : ABC 345 "un known" END 8.88  -> ABC 345 un_known END 8.88

 Dim index As Integer, substring As String, substring_repl As String
 index = InStr(sline,Chr$(34))
 If index Then
  substring=LTrim(Mid(sline,index+1))
  index = InStr(substring,Chr$(34))
  substring=LTrim(Left(substring,index-1))
  substring_repl= Replace (substring," ","_")
  eliminate_blanks_in_strings=Replace(sline,substring,substring_repl)
  eliminate_blanks_in_strings=Replace(eliminate_blanks_in_strings,Chr$(34)," ")
 Else
  eliminate_blanks_in_strings = sline
 End If
End Function

Function Convert2Double(ByVal sline As String) As Double
 Dim significant_digits As String
 Dim zero_number As Double
 significant_digits = "0.00E+00"  
 zero_number = 0.
 If Len(sline)>0 Then
  'Convert2Double  =  evaluate(Format(evaluate(sline),significant_digits))
   Convert2Double = CDbl(evaluate(sline))
 Else
  Convert2Double = CDbl(evaluate(Format(evaluate(zero_number),significant_digits)))
 End If
End Function
Sub Define_layer _
 (L_Name As String, epsilon As Double, mu As Double,kappa As Double,tand As Double,tandfreq As Double,   _
            kappam As Double, tandm As Double, tandmfreq As Double)

Print #2, "With Material"
  Print #2, "   .Reset "
    Print #2, " .Name "+ Chr$(34)+L_Name+Chr$(34)
     Print #2, ".FrqType ""hf"" 
    Print #2, " .Type ""Normal""
     Print #2, ".Epsilon "+ Chr$(34)+evaluate(epsilon)+Chr$(34)
    Print #2, " .Mu "+ Chr$(34)+evaluate(mu)+Chr$(34)
     Print #2, ".Kappa "+ Chr$(34)+evaluate(kappa)+Chr$(34)
     Print #2, ".SetMaterialUnit " + Chr$(34)+ "GHz"+ Chr$(34) + ", "  + Chr$(34)+ "mm"+ Chr$(34)
    Print #2, " .TanD " + Chr$(34)+evaluate(tand)+Chr$(34) 
    Print #2, " .TanDFreq "+Chr$(34)+evaluate(tandfreq)+Chr$(34) 
     Print #2, ".TanDGiven ""True"" 
    Print #2, " .TanDModel ""ConstTanD"" 
     Print #2, ".KappaM "+Chr$(34)+evaluate(kappam)+Chr$(34)  
     Print #2, ".TanDM "+Chr$(34)+evaluate(tandm)+Chr$(34) 
     Print #2, ".TanDMFreq "+Chr$(34)+evaluate(tandmfreq)+Chr$(34)  
     Print #2, ".TanDMGiven ""False"" 
     Print #2, ".DispModelEps ""None"" 
     Print #2, ".DispModelMu ""None"" 
     Print #2, ".Rho ""0.0"" 
     Print #2, " .Colour  1,1,0.001 'yellow
     If dlg_substrate_wireframe = 1 Then	'wireframe-mode
		Print #2, ".Wireframe ""True""
     Else
      Print #2, ".Wireframe ""False"" 
      Print #2, ".Transparency ""80""	'transparent mode
     End If
      Print #2, ".Create"
 Print #2, "End With" 
End Sub 
 

Sub Define_layer_PEC(L_Name As String)

Print #2, "With Material"
  Print #2, "   .Reset "
    Print #2, " .Name "+ Chr$(34)+L_Name+Chr$(34)
     Print #2, ".FrqType ""hf"" 
    Print #2, " .Type ""Pec""
     Print #2, ".Rho ""0.0"" 
     Print #2, ".SetMaterialUnit " + Chr$(34)+ "GHz"+ Chr$(34) + ", "  + Chr$(34)+ "mm"+ Chr$(34)
    Print #2, " .Colour  0.75,.751,.751  	'gray
      Print #2, ".Wireframe ""False"" 
      Print #2, ".Transparency ""0"" 
      Print #2, ".Create"
 Print #2, "End With" 
End Sub
Sub Define_layer_surf_imp(L_Name As String,kappa As Double)

Print #2, "With Material"
  Print #2, "   .Reset "
    Print #2, " .Name "+ Chr$(34)+L_Name+Chr$(34)
     Print #2, ".FrqType ""hf"" 
    Print #2, " .Type ""Lossy metal""
    Print #2, " .Mu ""1.0"" 
     Print #2, ".Kappa "+ Chr$(34)+evaluate(kappa)+Chr$(34) 
     Print #2, ".SetMaterialUnit " + Chr$(34)+ "GHz"+ Chr$(34) + ", "  + Chr$(34)+ "mm"+ Chr$(34)
     Print #2, ".Rho ""0.0"" 
    Print #2, " .Colour  0.751,0.751,.750  	'gray
      Print #2, ".Wireframe ""False"" 
      Print #2, ".Transparency ""0"" 		'transparency
      Print #2, ".Create"
 Print #2, "End With" 
End Sub
'
' Find all child entry-names
' Parameters:  String of mother's name, Array containing all Child names, Number of entries found
' Nov 18_2002 fhi: initial version
'
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
Function DirName (lib_path As String) As String

        Dim lib_dircount As Integer

        lib_dircount = InStrRev(lib_path, "\")
        DirName  = Left$(lib_path, IIf(lib_dircount > 1, lib_dircount - 1, 0))

End Function
'--------------------------------------------------------------------------------------------
' Implement Dialog Callbacks

Function DialogFunc%(DlgItem$, Action%, SuppValue%)

Dim sonnet_file As String, export_file As String

Select Case Action%
Case 1 ' Dialog box initialization

Case 2 ' Value changing or button pressed
    Select Case DlgItem$
    Case "push_browse"
      'sonnet_file=GetFilePath("", "son",DirName(getprojectbasename),,0)
      sonnet_file=GetFilePath("", "son", getprojectpath("Root"),,0)

	  'cst_export_dir_filename = DirName(getprojectbasename)	' sets the output dir in case that the export dialog is not openend
	cst_export_dir_filename =  getprojectpath("Root")

	 If sonnet_file <>"" Then
	    cst_sonnet_dir_filename = DirName(sonnet_file)
		DlgText "edit_filename", ShortName(sonnet_file)
		'DlgText "export_filename", BaseName(getprojectbasename)+"_"+BaseName(ShortName(sonnet_file))+".mod"
		'DlgText "export_filename",  getprojectpath("Root")+"\"+BaseName(ShortName(sonnet_file))+".cst"  'long version
		DlgText "export_filename",  ".\"+BaseName(ShortName(sonnet_file))+"_Sonnet.mod"   'short version, keep extension .mod!
	 End If
	DialogFunc%=True		' do not exit the dialog
    Case "push_export"
      'export_file=GetFilePath("", "mod",DirName(getprojectbasename),,1)
      export_file=GetFilePath("", "mod",DirName(getprojectpath("Root")),,1)

	 If export_file<>"" Then
	  cst_export_dir_filename = DirName(export_file)
		DlgText "export_filename", ShortName(export_file)
	 End If
	DialogFunc%=True		' do not exit the dialog

	End Select
Case 4 ' Focus changed

End Select
End Function
