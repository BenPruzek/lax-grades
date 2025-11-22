'#Language "WWB-COM"
' Read a FastHenry file and generate an MWS structure

'--------------------------------------------------------------------------------------------------------------------
' 11-nov-2013 imu: AddToHistory, Ports as discrete ports and not as curves
' 30-jul-2009 ube: Split replaced by CSTSplit, since otherwise compeating with standard VBA-Split function
' 15-jan-2009 imu: first version
'--------------------------------------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Public Const  MAX_NODES =  99999
' ================================================================================================
' Copyright 2019-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 17-Feb-2019 ube: First version
' ================================================================================================

Public ADD_TO_HISTORY As Boolean
Public PORTS_AS_CURVES As Boolean

Type node
	nname As String
	x As Double
	y As Double
	z As Double
End Type

Type pnt
	coord(3) As Double
End Type

Type bricktype
	bname As String
	compname As String
	nname(2) As String
	n(2) As Long
	sigma As Double
	w As Double
	h As Double
	wx As Double
	wy As Double
	wz As Double
End Type

Type porttype
	portname As String
	nodes(2) As node
End Type


Dim nnodes As Long
Dim nbricks As Long
Dim nports As Long, nequivs As Long
Dim nodes(MAX_NODES) As node
Dim bricks(MAX_NODES) As bricktype
Dim FH_ports(MAX_NODES) As porttype
Dim FH_equivs(MAX_NODES) As porttype

Dim FH_unit_name As String
Dim fmin As Double, fmax As Double

Dim draw_curves As Boolean

Public  cst_outfile As String, cst_infile_terminals As String

Sub Main
	read_FH()

'	Save
'	FileNew	' New MWS file

	MWS_settings()
	generate_bricks()
	generate_ports()
	generate_equivs()
'	DrawXYZPickPoints_imu cst_infile_terminals, 1000
End Sub

Sub MWS_settings()
	Dim tmpstr As String

	If ADD_TO_HISTORY = False Then
	Units.SetUnit("Length", LCase(FH_unit_name))
	Solver.FrequencyRange CStr(fmin), CStr(fmax)
	Else
		AddToHistory ( "FastHenry2CST define units", "Units.SetUnit(""Length"",""" + LCase(FH_unit_name) + """)" )
		AddToHistory ( "FastHenry2CST define frequency range", "Solver.FrequencyRange "+ CStr(fmin) + "," + CStr(fmax) )
	End If

End Sub


Sub generate_bricks()
	Dim tmpstr As String, tmpstr1 As String

	Dim iii As Long, jjj As Long
	Dim n1 As node, n2 As node
	Dim nd As node, length As Double

	Dim nmat As Long, sigma As Double
	Dim matnames(MAX_NODES) As String, matsigma(MAX_NODES) As Double

	nmat = 0	' Materials

	tmpstr = ""

If ADD_TO_HISTORY = False Then

	For iii = 0 To nbricks-1
	'	AddToHistory ("activate local coordinates", "WCS.ActivateWCS ""local""")
		WCS.ActivateWCS "local"
		find_nodes(bricks(iii).nname(0), bricks(iii).nname(1), n1, n2)
		node_diff(n1, n2, nd, length)

		With WCS
			.SetNormal nd.x, nd.y, nd.z
			.SetOrigin n1.x, n1.y, n1.z
			.SetUVector bricks(iii).wx, bricks(iii).wy, bricks(iii).wz
		End With
'		AddToHistory ("set WCS properties", tmpstr)

		sigma =bricks(iii).sigma
		With Brick
			.Reset
			.Name "MyBrick"+Cstr(iii)
			.Component "Z_"+bricks(iii).compname
			.Material "PEC"
			.Xrange - bricks(iii).w/2,  + bricks(iii).w/2
			.Yrange - bricks(iii).h/2,  + bricks(iii).h/2
			.Zrange 0, length
			.Create
		End With

		If draw_curves = True Then
'			AddToHistory ("activate global coordinates", "WCS.ActivateWCS ""global""")
			WCS.ActivateWCS "global"

			With Polygon3D
				.Reset
				.Name bricks(iii).compname+"_"+Cstr(iii)
				.Curve "Z_curves"
				.Point n1.x, n1.y, n1.z
				.Point n2.x, n2.y, n2.z
				.Create
			End With
		End If

	Next


Else	' ADD_TO_HISTORY = True
	tmpstr = tmpstr + "'### BRICKS ### " + vbCrLf
	tmpstr1 = tmpstr1 + "'### CURVES ### " + vbCrLf

	For iii = 0 To nbricks-1
	'	AddToHistory ("activate local coordinates", "WCS.ActivateWCS ""local""")
				tmpstr = tmpstr + "WCS.ActivateWCS ""local""" + vbCrLf
		find_nodes(bricks(iii).nname(0), bricks(iii).nname(1), n1, n2)
		node_diff(n1, n2, nd, length)

		tmpstr = tmpstr + "With WCS" + vbCrLf
		tmpstr = tmpstr + "	.SetNormal " + cstr(nd.x) + ", " + CStr(nd.y)+ ", " + CStr( nd.z) + vbCrLf
		tmpstr = tmpstr + "	.SetOrigin "  + cstr(n1.x) + ", " + CStr(n1.y)+ ", " + CStr( n1.z) + vbCrLf
		tmpstr = tmpstr + "	.SetUVector " + cstr(bricks(iii).wx) + ", " + CStr(bricks(iii).wy) + ", " + CStr( bricks(iii).wz) + vbCrLf
		tmpstr = tmpstr + "End With"+ vbCrLf+ vbCrLf
'		AddToHistory ("set WCS properties", tmpstr)

		sigma =bricks(iii).sigma
		tmpstr = tmpstr + "With Brick" + vbCrLf
		tmpstr = tmpstr + "	.Reset" + vbCrLf
		tmpstr = tmpstr + "	.Name ""MyBrick"+Cstr(iii) + """"+vbCrLf
		tmpstr = tmpstr + "	.Component ""Z_"+bricks(iii).compname + """" + vbCrLf
		tmpstr = tmpstr + "	.Material ""PEC""" + vbCrLf
		tmpstr = tmpstr + "	.Xrange " + cstr(- bricks(iii).w/2) + ",  "+ cstr(bricks(iii).w/2) + vbCrLf
		tmpstr = tmpstr + "	.Yrange " + cstr(- bricks(iii).h/2) + ",  "+ cstr(bricks(iii).h/2) + vbCrLf
		tmpstr = tmpstr + "	.Zrange ""0"", " + cstr(length) + vbCrLf
		tmpstr = tmpstr + "	.Create" + vbCrLf
		tmpstr = tmpstr + "End With" + vbCrLf

		If draw_curves = True Then
'			AddToHistory ("activate global coordinates", "WCS.ActivateWCS ""global""")
			tmpstr1 = tmpstr1 + "'### CURVES ### " + vbCrLf
			tmpstr1 = tmpstr1 + "WCS.ActivateWCS ""global"""  + vbCrLf

			tmpstr1 = tmpstr1 + "With Polygon3D" + vbCrLf
			tmpstr1 = tmpstr1 + "	.Reset" + vbCrLf
			tmpstr1 = tmpstr1 + "	.Name """+bricks(iii).compname+"_"+Cstr(iii) +""""+ vbCrLf
			tmpstr1 = tmpstr1 + "	.Curve ""Z_curves""" + vbCrLf
			tmpstr1 = tmpstr1 + "	.Point " + cstr(n1.x) + ", " + cstr(n1.y) + ", " + cstr(n1.z) + vbCrLf
			tmpstr1 = tmpstr1 + "	.Point " + cstr(n2.x) + ", " + cstr(n2.y) + ", " + cstr(n2.z) + vbCrLf
			tmpstr1 = tmpstr1 + "	.Create" + vbCrLf
			tmpstr1 = tmpstr1 + "End With" + vbCrLf
		End If
	Next

	AddToHistory ( "FastHenry2CST bricks ", tmpstr)
	AddToHistory ( "FastHenry2CST curves ", tmpstr1)

End If
	'AddToHistory ( String header, String contents
End Sub

Sub generate_ports()
	Dim tmpstr As String

	Dim iii As Long, jjj As Long
	Dim n1 As node, n2 As node
	Dim nd As node, length As Double

'	Dim nmat As Long, sigma As Double
'	Dim matnames(MAX_NODES) As String, matsigma(MAX_NODES) As Double

'	nmat = 0	' Materials

If ADD_TO_HISTORY = False Then
'		AddToHistory ("activate global coordinates", "WCS.ActivateWCS ""global""")
		WCS.ActivateWCS "global"


	For iii = 0 To nports-1
	'	AddToHistory ("activate local coordinates", "WCS.ActivateWCS ""local""")
	'	WCS.ActivateWCS "local"
		find_nodes(FH_ports(iii).nodes(0).nname, FH_ports(iii).nodes(1).nname, n1, n2)
'		node_diff(n1, n2, nd, length)

'		With WCS
'			.SetNormal nd.x, nd.y, nd.z
'			.SetOrigin n1.x, n1.y, n1.z
'			.SetUVector bricks(iii).wx, bricks(iii).wy, bricks(iii).wz
'		End With
'		AddToHistory ("set WCS properties", tmpstr)

'		sigma =bricks(iii).sigma

		If n1.x = n2.x And n1.y = n2.y And n1.z = n2.z Then
			n2.x = n1.x + 0.001
			n2.y = n1.y + 0.001
			n2.z = n1.z + 0.001
		End If


If PORTS_AS_CURVES = True Then
		With Polygon3D
			.Reset
			.Name FH_ports(iii).portname
			.Curve "Z_ports"
			.Point n1.x, n1.y, n1.z
			.Point n2.x, n2.y, n2.z
			.Create
		End With
Pick.PickCurveEndpointFromId "Z_ports:"+FH_ports(iii).portname, "1"
End If

If PORTS_AS_CURVES = False Then
		With DiscretePort
	     .Reset
	     .PortNumber iii+1
	     .Type "SParameter"
	     .Label FH_ports(iii).portname
    	 .Impedance "50.0"
'	     .VoltagePortImpedance "0.0"
'	     .Voltage "1.0"
'	     .Current "1.0"
	     .SetP1 "False", n1.x, n1.y, n1.z
	     .SetP2 "False", n2.x, n2.y, n2.z
	     .InvertDirection "False"
	     .LocalCoordinates "False"
	     .Monitor "True"
	     .Radius "0.0"
	     .Wire ""
	     .Position "end1"
	     .Create
		End With
End If

	Next
Else
	tmpstr = ""
	tmpstr = tmpstr + "'### PORTS ### " + vbCrLf
	tmpstr = tmpstr+	"WCS.ActivateWCS ""global""" + vbCrLf


	For iii = 0 To nports-1
		find_nodes(FH_ports(iii).nodes(0).nname, FH_ports(iii).nodes(1).nname, n1, n2)

		If n1.x = n2.x And n1.y = n2.y And n1.z = n2.z Then
			n2.x = n1.x + 0.001
			n2.y = n1.y + 0.001
			n2.z = n1.z + 0.001
		End If

If 1 = 1 Then
		tmpstr = tmpstr + "With Polygon3D" + vbCrLf
		tmpstr = tmpstr + "	.Reset" + vbCrLf
		tmpstr = tmpstr + "	.Name """ + FH_ports(iii).portname + """" + vbCrLf
		tmpstr = tmpstr + "	.Curve ""Z_ports""" + vbCrLf
		tmpstr = tmpstr + "	.Point " + cstr(n1.x) + ", " + CStr(n1.y) + ", " + CStr(n1.z)  + vbCrLf
		tmpstr = tmpstr + "	.Point " + cstr(n2.x) + ", " + CStr( n2.y) + ", " + CStr(n2.z)  + vbCrLf
		tmpstr = tmpstr + "	.Create" + vbCrLf
		tmpstr = tmpstr + "End With" + vbCrLf
		Pick.PickCurveEndpointFromId "Z_ports:"+FH_ports(iii).portname, "1"
End If

If 1 = 0 Then
		tmpstr = tmpstr + "With DiscretePort" + vbCrLf
	    tmpstr = tmpstr + " .Reset" + vbCrLf
	    tmpstr = tmpstr + " .PortNumber " + cstr(iii+1) + vbCrLf
	    tmpstr = tmpstr + " .Type ""SParameter""" + vbCrLf
	    tmpstr = tmpstr + " .Label """ + FH_ports(iii).portname + """" + vbCrLf
    	tmpstr = tmpstr + " .Impedance ""50.0""" + vbCrLf
'	     .VoltagePortImpedance "0.0"
'	     .Voltage "1.0"
'	     .Current "1.0"
	    tmpstr = tmpstr + " .SetP1 ""False"", " + cstr(n1.x) + ", "+ CSTr(n1.y) + ", "+ CSTr(n1.z) + vbCrLf
	    tmpstr = tmpstr + " .SetP2 ""False"", " + cstr(n2.x) + ", "+ CSTr(n2.y) + ", "+ CSTr(n2.z) + vbCrLf
	    tmpstr = tmpstr + " .InvertDirection ""False""" + vbCrLf
	    tmpstr = tmpstr + " .LocalCoordinates ""False""" + vbCrLf
	    tmpstr = tmpstr + " .Monitor ""True""" + vbCrLf
	    tmpstr = tmpstr + " .Radius ""0.0""" + vbCrLf
	    tmpstr = tmpstr + " .Wire """"" + vbCrLf
	    tmpstr = tmpstr + " .Position ""end1""" + vbCrLf
	    tmpstr = tmpstr + " .Create" + vbCrLf
		tmpstr = tmpstr + "End With" + vbCrLf
End If

	Next
	AddToHistory ( "FastHenry2CST ports", tmpstr)
End If 	' ADD TO HISTORY
	'AddToHistory ( String header, String contents
End Sub

Sub generate_equivs()
	Dim tmpstr As String

	Dim iii As Long, jjj As Long
	Dim n1 As node, n2 As node
	Dim nd As node, length As Double


If ADD_TO_HISTORY = False Then

'		AddToHistory ("activate global coordinates", "WCS.ActivateWCS ""global""")
		WCS.ActivateWCS "global"


	For iii = 0 To nequivs-1
		find_nodes(FH_equivs(iii).nodes(0).nname, FH_equivs(iii).nodes(1).nname, n1, n2)

		If n1.x = n2.x And n1.y = n2.y And n1.z = n2.z Then
			n2.x = n1.x + 0.1
			n2.y = n1.y + 0.1
			n2.z = n1.z + 0.1
		End If


		With Polygon3D
			.Reset
			.Name CStr(iii)+"_"+FH_equivs(iii).portname
			.Curve "Z_equivs"
			.Point n1.x, n1.y, n1.z
			.Point n2.x, n2.y, n2.z
			.Create
		End With

	Next
Else
	tmpstr = ""
	tmpstr = tmpstr + "'### EQUIVS ### " + vbCrLf

	'		AddToHistory ("activate global coordinates", "WCS.ActivateWCS ""global""")
	tmpstr = tmpstr + "	WCS.ActivateWCS ""global""" + vbCrLf


	For iii = 0 To nequivs-1
		find_nodes(FH_equivs(iii).nodes(0).nname, FH_equivs(iii).nodes(1).nname, n1, n2)

		If n1.x = n2.x And n1.y = n2.y And n1.z = n2.z Then
			n2.x = n1.x + 0.1
			n2.y = n1.y + 0.1
			n2.z = n1.z + 0.1
		End If


		tmpstr = tmpstr + "With Polygon3D" + vbCrLf
		tmpstr = tmpstr + "	.Reset" + vbCrLf
		tmpstr = tmpstr + "	.Name """ +  CStr(iii)+"_"+FH_equivs(iii).portname + """" + vbCrLf
		tmpstr = tmpstr + "	.Curve ""Z_equivs""" + vbCrLf
		tmpstr = tmpstr + "	.Point " + cstr(n1.x) + ", " + CStr(n1.y) + ", " + CStr(n1.z)  + vbCrLf
		tmpstr = tmpstr + "	.Point " + cstr(n2.x) + ", " + CStr( n2.y) + ", " + CStr(n2.z)  + vbCrLf
		tmpstr = tmpstr + "	.Create" + vbCrLf
		tmpstr = tmpstr + "End With" + vbCrLf

	Next
	AddToHistory ( "FastHenry2CST equivs", tmpstr)
End If 	' ADD TO HISTORY
	'AddToHistory ( String header, String contents
End Sub

Sub node_diff(n1 As node, n2 As node, nd As node, length As Double)
	nd.x = n2.x - n1.x
	nd.y = n2.y - n1.y
	nd.z = n2.z - n1.z
	length = Sqr(nd.x^2 + nd.y^2 + nd.z^2)
End Sub


Function find_nodes(n1_in As String, n2_in As String, n1 As node, n2 As node) As Boolean
	Dim iii As Long
	Dim found_n1 As Boolean, found_n2 As Boolean

	For iii = 0 To nnodes-1
		If StrComp (LCase(n1_in), LCase(nodes(iii).nname)) = 0 Then
			n1 = nodes(iii)
			found_n1 = True
		End If
		If StrComp (LCase(n2_in), LCase(nodes(iii).nname)) = 0 Then
			n2 = nodes(iii)
			found_n2 = True
		End If
	Next

	If found_n1 = False Or found_n2 = False Then
		MsgBox "Could not find one of the nodes " + n1_in + ", " + n2_in
		Exit All
	End If
End Function




Sub read_FH()
	Dim iii As Long

	'Dim cst_infile As String, cst_outfile As String
	Dim cst_infile As String, cstfile As String
	Dim cst_inline As String, cst_inline_bak As String
	Dim cst_line_array(30) As String
	Dim cst_macro_via_outer_radius As Double, cst_macro_via_inner_radius As Double
	Dim cst_noc As Integer, cst_via_nr As Long
	Dim cst_mod_file As String, cst_sat_file As String, cst_backup_sat_file As String
	Dim cst_tmp As Double, cst_z1 As Double, cst_z2 As Double

	Dim tmpstr As String, tmpstr1 As String

	Dim sigma_def As Double, wx_def As Double, wy_def As Double, wz_def As Double, w_def As Double, h_def As Double
	Dim compon As String

	compon = "Z_bricks"
	draw_curves=False


	Begin Dialog UserDialog 600,245,"Read FastHenry .inp file",.DialogFunc ' %GRID:10,7,1,1
		Text 20,28,150,14,"FastHenry input file",.Text1
'		Text 30,63,150,14,"Output file",.Text2
		TextBox 180,28,310,21,.Infile
'		TextBox 180,56,310,21,.Outfile
		PushButton 500,28,90,21,"Browse",.Browseinputfile
'		PushButton 500,56,90,21,"Browse",.Browseoutputfile
		OKButton 310,217,90,21
		CancelButton 420,217,90,21
		GroupBox 30,56,240,70,"Draw bricks / curves",.GroupBox1
		OptionGroup .drawcurves
			OptionButton 70,77,180,14,"Draw bricks and curves",.OptionButton2
			OptionButton 70,98,130,14,"Draw only bricks",.OptionButton1
		Text 20,7,570,14,"Please perform a history update when you are finished analyzing the results of this macro",.Text3
		CheckBox 320,175,140,14,"Write to history",.write_to_hist
		GroupBox 40,140,230,63,"Ports",.GroupBox2
		OptionGroup .Ports_as_curves
			OptionButton 60,154,90,14,"As curves",.OptionButton3
			OptionButton 60,182,170,14,"As MWS discrete ports",.OptionButton4
	End Dialog
'		Text 30,63,150,14,"Output file",.Text2
'		TextBox 180,56,310,21,.Outfile
'		PushButton 500,56,90,21,"Browse",.Browseoutputfile
	Dim dlg As UserDialog
	dlg.drawcurves = 0
	dlg.Infile = "Fasthenry.inp"
	If (Dialog(dlg) = 0) Then Exit All

	cst_infile = dlg.Infile
'	cst_outfile = dlg.Outfile
	cst_infile_terminals = Left(cst_infile, InStr(cst_infile, ".inp")-1) + "_port_terminals.txt"
	If dlg.drawcurves = 0 Then draw_curves = True
	If dlg.write_to_hist = 1 Then
		ADD_TO_HISTORY = True
	Else
		ADD_TO_HISTORY = False
	End If

	If dlg.Ports_as_curves = 0 Then
		PORTS_AS_CURVES = True
	Else
		PORTS_AS_CURVES = False
	End If

	If StrComp(cst_infile, "Fasthenry.inp") = 0 Then cst_infile = GetProjectPath("Model3D") + "Fasthenry.inp"

	Open cst_infile For  Input As #1
'	cst_outfile = cst_mod_file
'	Open cst_outfile For Output As #2

	wx_def = 1: wy_def = 0: wz_def = 0

	While Not EOF(1)

		'--- read data from file
		Line Input #1,cst_inline
		cst_inline_bak = cst_inline

'		If IsNumeric(Left(LTrim(cst_inline),1)) Or IsNumeric(Left(LTrim(cst_inline),2)) Then
			cst_noc = CSTSplit(cst_inline, cst_line_array())
'			cst_z1 = RealVal(cst_line_array(2))
'			cst_z2 = RealVal(cst_line_array(3))
'			If cst_z2 < cst_z1 Then
'				cst_tmp = cst_z2
'				cst_z2 = cst_z1
'				cst_z1 = cst_tmp
'			End If
'		End If

		Select Case UCase(  Left(cst_inline_bak,1))
			Case "*"
				' Comment
				If StrComp(LCase(cst_line_array(1)), "component", vbBinaryCompare) = 0 Then
					compon = cst_line_array(2)
				End If
			Case "N"
				nodes(nnodes).nname = LCase(cst_line_array(0))
				nodes(nnodes).x = RealVal(Right(cst_line_array(1),Len(cst_line_array(1))-2))
				nodes(nnodes).y = RealVal(Right(cst_line_array(2),Len(cst_line_array(2))-2))
				nodes(nnodes).z = RealVal(Right(cst_line_array(3),Len(cst_line_array(3))-2))
				nnodes = nnodes+1
			Case "E"
				' Initialize default values; might be overwritten later
				bricks(nbricks).w = w_def
				bricks(nbricks).h = h_def
				bricks(nbricks).sigma = sigma_def
				bricks(nbricks).wx = wx_def
				bricks(nbricks).wy = wy_def
				bricks(nbricks).wz = wz_def

				bricks(nbricks).bname = cst_line_array(0)
				bricks(nbricks).nname(0) = LCase(cst_line_array(1))
				bricks(nbricks).nname(1) = LCase(cst_line_array(2))
				For iii = 3 To cst_noc-1
					tmpstr = cst_line_array(iii)
					tmpstr1 = LCase(Left( tmpstr, InStr(tmpstr, "=")))
					tmpstr = Right(tmpstr, Len(tmpstr)-InStr(tmpstr, "="))
					Select Case tmpstr1
						Case "w="
							bricks(nbricks).w = RealVal(tmpstr)
						Case "h="
							bricks(nbricks).h = RealVal(tmpstr)
						Case "wx="
							bricks(nbricks).wx = RealVal(tmpstr)
						Case "wy="
							bricks(nbricks).wy = RealVal(tmpstr)
						Case "wz="
							bricks(nbricks).wz = RealVal(tmpstr)
						Case "sigma="
							bricks(nbricks).sigma = RealVal(tmpstr)
					End Select
				Next
				bricks(nbricks).compname = compon
				check_brick(nbricks)	' Check if no dimension is zero ...
				nbricks = nbricks+1

			Case "."
				tmpstr = cst_line_array(0)
				Select Case LCase(Right(tmpstr, Len(tmpstr)-1))
					Case "units"
						FH_unit_name = cst_line_array(1)
					Case "default"
						For iii = 1 To cst_noc-1
							tmpstr = cst_line_array(iii)
							tmpstr1 = LCase(Left( tmpstr, InStr(tmpstr, "=")))
							tmpstr = Right(tmpstr, Len(tmpstr)-InStr(tmpstr, "="))
							Select Case tmpstr1
								Case "sigma="
									sigma_def = RealVal(tmpstr)
								Case "w="
									w_def = RealVal(tmpstr)
								Case "h="
									h_def = RealVal(tmpstr)
								Case "wx="
									wx_def = RealVal(tmpstr)
								Case "wy="
									wy_def = RealVal(tmpstr)
								Case "wz="
									wz_def = RealVal(tmpstr)
							End Select

						Next iii
					Case "external"
						FH_ports(nports).nodes(0).nname = cst_line_array(1)
						FH_ports(nports).nodes(1).nname = cst_line_array(2)
						FH_ports(nports).portname = cst_line_array(3)
						nports = nports+1
					Case "equiv"
						For iii = 2 To cst_noc-1
							FH_equivs(nequivs).nodes(0).nname = cst_line_array(1)
							FH_equivs(nequivs).nodes(1).nname = cst_line_array(iii)
							FH_equivs(nequivs).portname = cst_line_array(1)+"_"+cst_line_array(iii)
							nequivs = nequivs+1
						Next iii
					Case "freq"
						For iii = 1 To cst_noc-1
							tmpstr = cst_line_array(iii)
							tmpstr1 = LCase(Left( tmpstr, InStr(tmpstr, "=")))
							tmpstr = Right(tmpstr, Len(tmpstr)-InStr(tmpstr, "="))
							Select Case tmpstr1
								Case "fmin="
									fmin = RealVal(tmpstr)
								Case "fmax="
									fmax = RealVal(tmpstr)
							End Select
						Next iii
						If (fmax = fmin) Then fmax = fmin+1
				End Select

		'--- split line into parts and export data; check if numeric line before split

		End Select
	Wend

	Close #1

	Save

'	cstfile = GetProjectPath("Project") + ".cst"
'	cst_mod_file = GetProjectPath("Model3D")+"Model.mod"
'	cst_sat_file = GetProjectPath("ModelCache")+"Model.sat"
'	cst_backup_sat_file = GetProjectPath("ModelCache")+"Model_backup.sat"

'	FileNew


End Sub

Sub check_brick(nb As Long)

End Sub

'-------------------------------------------------------------------------------------
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean

	Dim Extension As String, projectdir As String, filename As String
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        Select Case DlgItem
        	Case "Browseinputfile"
        		Extension = "inp;txt;dat"
                projectdir = GetProjectPath("Model3D")	'Dirname(GetProjectbasename)
                filename = GetFilePath(,Extension, projectdir, "Specify input file", 0)
                If (filename <> "") Then
                    DlgText "Infile", FullPath(filename, projectdir)
                End If
        		dialogfunc = True
        	Case "Browseoutputfile"
        		Extension = "cst"
                projectdir = Dirname(GetProjectbasename)
                filename = GetFilePath(,Extension, projectdir, "Specify CST output file", 1)
                If (filename <> "") Then
                    DlgText "Outfile", FullPath(filename, projectdir)
                End If
        	dialogfunc = True
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Sub DrawXYZPickPoints_imu (sFileCoordxyz As String, nMaxPicks As Long)

	If Dir$(sFileCoordxyz)="" Then Exit Sub

	SelectTreeItem "Components"
	Plot.Wireframe True
	Pick.ClearAllPicks

	Dim iLineCounter As Long, iLastLine As Long, dDiffID As Double, dNextID As Double
	Dim sLine_lib As String, string_item(30) As String, cst_nitems As Integer

	iLineCounter = 0

    Open sFileCoordxyz For Input As #11
		While Not EOF(11)
			Line Input #11, sLine_lib
			iLineCounter = iLineCounter + 1
		Wend
    Close #11

    iLastLine = iLineCounter
    '
    ' --- now pick up to nMaxPicks points, always take first and last point
    '
	dDiffID = CDbl(iLastLine)/nMaxPicks
	dNextID = 0.0

    Open sFileCoordxyz For Input As #11
		While Not EOF(11)
			Line Input #11, sLine_lib
			iLineCounter = iLineCounter + 1
			If (iLineCounter = iLastLine Or iLineCounter > dNextID) Then
				' pick this point
				cst_nitems = CSTSplit(sLine_lib, string_item)
				Pick.PickPointFromCoordinates string_item(0),string_item(1),string_item(2)

				dNextID = dNextID + dDiffID
			End If
		Wend
    Close #11

	Plot.Update

End Sub
