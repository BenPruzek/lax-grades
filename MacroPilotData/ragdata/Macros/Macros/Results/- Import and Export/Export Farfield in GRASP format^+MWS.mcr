' *Farfield / GRASP feed file export
' !!! Do not change the line above !!!
' macro.802
'
'--------------------------------------------------------------------------------------------
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 04-Oct-2017 hcg: Excluded farfield cuts from selection list
' 13 Jun-2013 dta: corrected error when calculating total radiated power
' 05-Feb-2013 dta: Fixed problem when exporting FF with a Theta Step < 0.5°. Function DRound to keep 1 decimal digit
' 12-Dec-2012 dta: Bug fix for export using ICOMP=3 L3ECoCx (Horizontal-Vertical) and for LHCP/RHCP. It now exports first RHCP and then LHCP components.
' 24-Jul-2009 ube: CST GmbH\CST MicroWave Studio replaced by CST STUDIO SUITE
' 19-Jun-2007 ube: adapted to 2008
' 21-Oct-2005 ube: Included into Online Help
' reh: 25-08-2005  ffapprox=on, bugfix
' reh: 02-08-2005  Phase fix for rhcp/lhcp,l3cocx & ffapprox on/off functionality added
' reh: 10-06-2005  Redesign for Version 2006: Additional Options, All FF/Single FF, New FFList commands for speed up
' lsa: 14-04-2005  Define Free Origin for FarField
' reh: 28-11-2003  Andrew Special Line inserted
' reh: 26-11-2003  Bug fix for theta angles smaller then 0
' reh: 10-02-2003  Initial version
'--------------------------------------------------------------------------------------------
Option Explicit
Public cst_title_ini As String
Public cst_freq() As Double

'#include "vba_globals_all.lib"

Sub Main

	'=============================================================================
	Dim cst_tree() As String, cst_tmpstr As String
	Dim cst_iloop As Long, cst_iloop2 As Long
	Dim cst_nff As Long, cst_nom As Long
	Dim cst_ffq(3) As String

	Dim cst_theta_start As Double, cst_theta_step As Double, cst_theta_stop As Double, cst_ntheta As Long
	Dim cst_phi_start As Double, cst_phi_step As Double, cst_phi_stop As Double, cst_nphi As Long
	Dim cst_phi As Double, cst_theta As Double, cst_theta_calc As Double, cst_phi_calc As Double
	Dim cst_origin_type As String
	Dim cst_origin_x As Double, cst_origin_y As Double, cst_origin_z As Double
	Dim cst_dt As Double, cst_dp As Double

	Dim cst_icomp As Integer, cst_ncomp As Integer
	Dim cst_floop_start As Long, cst_floop_end As Long, cst_ifloop As Long

	Dim cst_title As String, cst_title_ini As String, cst_ffname As String, cst_filename As String
	Dim cst_radpow As Double, cst_ffam_renorm As Double
	Dim CST_FN As Long, cst_iff As Long
	Dim cst_theta_ph_renorm As Double, cst_phi_ph_renorm As Double, cst_phase_renorm As Double
	Dim cst_theta_re As Double, cst_theta_im As Double, cst_theta_am As Double, cst_theta_ph As Double
	Dim cst_phi_re As Double, cst_phi_im As Double, cst_phi_am As Double, cst_phi_ph As Double
	Dim cst_LHCP_re As Double, cst_LHCP_im As Double, cst_RHCP_re As Double, cst_RHCP_im As Double
	Dim cst_ffpol(3) As String
	Dim cst_fforigin As String
	Dim cst_xval As Double, cst_yval As Double, cst_zval As Double
	Dim cst_rad As Double, cst_rad_re As Double, cst_rad_im As Double
	Dim cst_time As Double
	'=============================================================================

	cst_time = Timer
	' On Error GoTo ERROR_NO_FARFIELDS
	ReDim cst_tree(1)
	cst_tree(0) = Resulttree.GetFirstChildName ("Farfields")
	cst_iloop = 0
	If cst_tree(0)="Farfields\Farfield Cuts" Then
	   cst_tree(0) = Resulttree.GetNextItemName ("Farfields\Farfield Cuts")
	End If
	If cst_tree(0)="" Then
		GoTo ERROR_NO_FARFIELDS
	End If
	Do
		cst_tmpstr = Resulttree.GetNextItemName(cst_tree(cst_iloop))
		If cst_tmpstr <> "" Then
			cst_iloop = cst_iloop+1
			ReDim Preserve cst_tree(cst_iloop)
			cst_tree(cst_iloop)=cst_tmpstr
		End If
	Loop Until cst_tmpstr = ""
	cst_nff = cst_iloop+1

	For cst_iloop = 1 To cst_nff
		cst_tmpstr = Replace(cst_tree(cst_iloop-1),"Farfields\","")
		cst_tree(cst_iloop-1)=cst_tmpstr
	Next cst_iloop

	ReDim cst_freq(cst_nff)
	cst_nom = Monitor.GetNumberOfMonitors
	For cst_iloop = 1 To cst_nom
		If Monitor.GetMonitorTypeFromIndex(cst_iloop-1) = "Farfield" Then
			For cst_iloop2 = 1 To cst_nff
				If InStr(cst_tree(cst_iloop2-1),Monitor.GetMonitorNameFromIndex(cst_iloop-1)) > 0 Then
					cst_freq(cst_iloop2-1) = Monitor.GetMonitorFrequencyFromIndex(cst_iloop-1)
				End If
			Next cst_iloop2
		End If
	Next cst_iloop

	cst_ffq(0) = "Spherical: theta,phi"
	cst_ffq(1) = "RHCP, LHCP"
	cst_ffq(2) = "Ludwig 3: Horizontal, Vertical"

	cst_ffpol(0) = "theta-phi"
	cst_ffpol(1) = "RHCP-LHCP"
	cst_ffpol(2) = "L3-CoCx"

	Begin Dialog UserDialog 850,284,"Farfield Export to GRASP",.DialogFunc ' %GRID:10,4,1,1
		OKButton 30,260,90,20
		CancelButton 130,260,90,20
		GroupBox 30,88,390,92,"Angular variation",.GroupBox1
		GroupBox 440,12,390,108,"Farfield Origin",.GroupCenterBox1
		TextBox 100,104,110,20,.thlow
		TextBox 290,104,110,20,.phlow
		GroupBox 440,124,390,60,"Select Polarization Type",.GroupBox4
		TextBox 290,128,110,20,.phhigh
		DropListBox 450,140,360,192,cst_ffq(),.ffq
		TextBox 290,152,110,20,.phstep
		TextBox 100,128,110,20,.thhigh
		TextBox 100,152,110,20,.thstep
		Text 40,108,50,12,"thlow",.Text1
		Text 230,108,50,12,"phlow",.Text4
		Text 230,132,50,12,"phhigh",.Text5
		Text 230,156,50,12,"phstep",.Text6
		Text 40,132,50,12,"thhigh",.Text2
		Text 40,156,50,12,"thstep",.Text3
		OptionGroup .GroupCenterBox
			OptionButton 460,32,290,16,"Center of bounding box",.OptionButton1
			OptionButton 460,52,290,16,"Origin of coordinate system",.OptionButton2
			OptionButton 460,72,290,16,"Free",.OptionButton3
		TextBox 490,92,50,20,.xcenter
		TextBox 580,92,50,20,.ycenter
		TextBox 670,92,50,20,.zcenter
		Text 470,96,10,16,"X",.Text7
		Text 560,96,10,16,"Y",.Text8
		Text 650,96,10,16,"Z",.Text9
		GroupBox 30,12,390,64,"Select Farfield Monitor",.GroupBox2
		DropListBox 40,48,370,192,cst_tree(),.FarfieldSelect
		CheckBox 40,28,150,16,"Export all Farfields",.FFall
		GroupBox 440,188,390,64,"Polarisation Vector (y'-axis)",.GroupBox3
		CheckBox 450,204,320,16,"Main Lobe / Polarization Vector Alignment",.CheckPolAlignment
		TextBox 490,224,50,20,.xpol
		TextBox 580,224,50,20,.ypol
		TextBox 670,224,50,20,.zpol
		Text 470,228,10,16,"X",.Text10
		Text 560,228,10,16,"Y",.Text11
		Text 650,228,10,16,"Z",.Text12
		GroupBox 30,188,390,64,"Nearfield / Farfield",.GroupBox5
		CheckBox 50,204,220,16,"Use farfield approximation",.FarfNearf
		TextBox 190,224,50,20,.radius
		Text 50,228,130,16,"Reference distance:",.Text13
		Text 250,228,20,16,"m",.Text14
		PushButton 230,260,90,20,"Help",.Help
	End Dialog
	Dim dlg As UserDialog

	'--- get registry settings
	dlg.thlow  = GetString("CST STUDIO SUITE", "GraspExport", "theta_low", "0.0")
	dlg.thhigh = GetString("CST STUDIO SUITE", "GraspExport", "theta_high", "180.0")
	dlg.thstep = GetString("CST STUDIO SUITE", "GraspExport", "theta_step", "45.0")
	dlg.phlow  = GetString("CST STUDIO SUITE", "GraspExport", "phi_low", "0.0")
	dlg.phhigh = GetString("CST STUDIO SUITE", "GraspExport", "phi_high", "360.0")
	dlg.phstep = GetString("CST STUDIO SUITE", "GraspExport", "phi_step", "90.0")

	cst_title_ini = "CST MWS Results: " + ShortName(GetProjectPath("Project"))

	If (Dialog(dlg) = 0) Then Exit All


	'--- write back registry settings
	SaveString  "CST STUDIO SUITE", "GraspExport", "theta_low", dlg.thlow
	SaveString  "CST STUDIO SUITE", "GraspExport", "theta_high", dlg.thhigh
	SaveString  "CST STUDIO SUITE", "GraspExport", "theta_step", dlg.thstep
	SaveString  "CST STUDIO SUITE", "GraspExport", "phi_low", dlg.phlow
	SaveString  "CST STUDIO SUITE", "GraspExport", "phi_high", dlg.phhigh
	SaveString  "CST STUDIO SUITE", "GraspExport", "phi_step", dlg.phstep

	'--- Theta start, Theta step, number of Theta steps, phicut, polarisation=1(theta,phi), polarisation=3(Ludwig3)
	cst_theta_start = RealVal(dlg.thlow)
	cst_theta_step  = RealVal(dlg.thstep)
	cst_theta_stop = RealVal(dlg.thhigh)
	cst_ntheta = IIf(cst_theta_step=0,0,Abs(cst_theta_stop-cst_theta_start)/cst_theta_step)
	cst_phi_start = RealVal(dlg.phlow)
	cst_phi_step = RealVal(dlg.phstep)
	cst_phi_stop = RealVal(dlg.phhigh)
	cst_nphi = IIf(cst_phi_step=0,0,Abs(cst_phi_stop-cst_phi_start)/cst_phi_step)
	'--- FarField Origin Free X, Y, Z
	cst_origin_type = CInt(dlg.GroupCenterBox)
	cst_origin_x = RealVal(dlg.xcenter)
	cst_origin_y = RealVal(dlg.ycenter)
	cst_origin_z = RealVal(dlg.zcenter)

	cst_rad = RealVal(dlg.radius)

	'--- check theta and phi settings
	cst_dt = cst_theta_stop - cst_theta_start
	cst_dp = cst_phi_stop - cst_phi_start

	If Not (((cst_dt <= 180.0) And (cst_dp <= 360.0)) Or ((cst_dt <= 360.0) And (cst_dp <= 180.0))) Then
		MsgBox "Please check angular settings: Some areas covered twice !",vbOkOnly+vbCritical,"Execution stopped"
		Exit Sub
	End If

	If cst_theta_start > cst_theta_stop Or cst_phi_start > cst_phi_stop Then
		MsgBox "Please check angular settings: lower values bigger than upper values !",vbOkOnly+vbCritical,"Execution stopped"
		Exit Sub
	End If

	'---
	' 	icomp=1:	E_theta, E_phi
	'	icomp=2:	E_rhcp, E_lhcp
	'	icomp=3:	Ludwig3 E_Co / E_Cx (co means horizontal (x-axis) and cross means vertical)
	cst_icomp  = CInt(dlg.ffq) + 1

	'--- number of field components: =2: for pure farfield data; =3 including near field
	If dlg.FarfNearf = 1 Then
		cst_ncomp = 2
	Else
		cst_ncomp = 3
	End If

	'--- set farfield evaluation points
	With FarfieldPlot
		.Reset
		.SetPlotMode "efield"
		.SetScaleLinear "True"
		Select Case cst_origin_type
			Case 0
				.Origin "bbox"
				cst_fforigin = "center of bounding box"
			Case 1
				.Origin "zero"
				cst_fforigin = "(x=0, y=0, z=0)"
			Case 2
				.Origin "free"
				'.Freeorigin cst_origin_x, cst_origin_y, cst_origin_z
				.Userorigin cst_origin_x, cst_origin_y, cst_origin_z
				cst_fforigin = "(x=" + Replace(Format(cst_origin_x,"Scientific"),",",".") + ", y=" + Replace(Format(cst_origin_y,"Scientific"),",",".") + ", z=" + Replace(Format(cst_origin_z,"Scientific"),",",".") + ")"
		End Select

		'--- Polarisation Vector
		If dlg.CheckPolAlignment = 1 Then
			.AlignToMainLobe "True"
	    	.PolarizationVector RealVal(dlg.xpol), RealVal(dlg.ypol), RealVal(dlg.zpol)
		Else
			.AlignToMainLobe "False"
		End If
		.UseFarfieldApproximation "True"
		If dlg.FarfNearf = 1 Then
    		.UseFarfieldApproximation "True"
    	Else
			.UseFarfieldApproximation "False"
    	End If

	End With

	'--- write list of farfield/nearfield points
	If dlg.FarfNearf = 1 Then
		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
			For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
				cst_F_tpcalc (cst_theta,cst_phi,cst_theta_calc,cst_phi_calc)
				FarfieldPlot.AddListItem(cst_theta_calc, cst_phi_calc, 0)
			cst_theta = DRound (cst_theta,2)
			Next cst_theta
		cst_phi = DRound (cst_phi,2)
		Next cst_phi
	Else
		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
			For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
				cst_F_tpcalc cst_theta,cst_phi,cst_theta_calc,cst_phi_calc
				FarfieldPlot.AddListItem(cst_theta_calc, cst_phi_calc, cst_rad)
			cst_theta = DRound (cst_theta,2)
			Next cst_theta
		cst_phi = DRound (cst_phi,2)
		Next cst_phi
	End If

	'--- All Farfields or just one
	If dlg.FFall = 0 Then
		cst_floop_start = CInt(dlg.FarfieldSelect)
		cst_floop_end = CInt(dlg.FarfieldSelect)
	Else
		cst_floop_start = 0
		cst_floop_end = cst_nff - 1
	End If

	For cst_ifloop = cst_floop_start To cst_floop_end

		cst_title = cst_title_ini + ", " + cst_tree(cst_ifloop) + ", Polarization: " + cst_ffpol(dlg.ffq) +", FF Origin: " + cst_fforigin
		'--- farfield plot init
		SelectTreeItem "Farfields\"+cst_tree(cst_ifloop)

			'dta 13 Jun 2013 corrected code to get total radiated power
			cst_radpow = FarfieldPlot.GetTRP*2

			If cst_radpow = 0 Then cst_radpow = 1

		cst_ffam_renorm = Sqr(4.0*Pi/cst_radpow)/Sqr(Sqr(4*Pi*1e-7/8.8542e-12))

		FarfieldPlot.CalculateList(cst_tree(cst_ifloop))

		cst_ffname = cst_tree(cst_ifloop)
		If dlg.FarfNearf = 1 Then
			cst_filename = GetProjectPath("Result") + "CSTGraspFeed_FFapprox_on_" + "(" + cst_ffname + "_" + cst_ffpol(dlg.ffq) +").cut"
		Else
			cst_filename = GetProjectPath("Result") + "CSTGraspFeed_FFapprox_off[r=" + CStr(cst_rad) + "m]_" + "(" + cst_ffname + "_" + cst_ffpol(dlg.ffq) +").cut"
		End If

		CST_FN = FreeFile
		Open cst_filename For Output As #CST_FN

		'--- calculate renorm angle at theta = 0 degree and for arbitrary phi
		'cst_theta_ph_renorm = -FarfieldPlot.CalculatePoint(0.0, 90.0, "Th_Phase", cst_tree(cst_ifloop))
		'cst_phi_ph_renorm = -FarfieldPlot.CalculatePoint(0.0, 90.0, "Ph_Phase", cst_tree(cst_ifloop))
		'cst_phase_renorm = cst_theta_ph_renorm

		cst_iff = 0

		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step

			Print #CST_FN,	cst_title
			Print #CST_FN, 	" " + _
						cstreh_Pretty(cst_theta_start) + _
						cstreh_Pretty(cst_theta_step)  + " " + _
						Format(cst_ntheta+1,"###")     + " " + _
						cstreh_Pretty(cst_phi)         + " " + _
						Format(cst_icomp,"###") + " 1 " + _
						Format(cst_ncomp,"###")

			For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step

				'--- radial components in case farf. approx. switched off
				If dlg.FarfNearf = 0 Then
					cst_rad_re   = FarfieldPlot.GetListItem(cst_iff,"radial re")
					cst_rad_im   = FarfieldPlot.GetListItem(cst_iff,"radial im")
				End If

				'--- transversal components
				If cst_icomp = 1 Then
					'--- spherical coordinates: theta, phi
					'Dim cst_t As String, cst_p As String
					'cst_t = FarfieldPlot.GetListItem(cst_iff,"Point_T")
					'cst_p = FarfieldPlot.GetListItem(cst_iff,"Point_P")

					cst_theta_re = FarfieldPlot.GetListItem(cst_iff,"th_re")
					cst_theta_im = FarfieldPlot.GetListItem(cst_iff,"th_im")
					cst_phi_re   = FarfieldPlot.GetListItem(cst_iff,"ph_re")
					cst_phi_im   = FarfieldPlot.GetListItem(cst_iff,"ph_im")

					If cst_theta < 0 Then
						cst_theta_re = -cst_theta_re
						cst_theta_im = -cst_theta_im
						cst_phi_re = -cst_phi_re
						cst_phi_im = -cst_phi_im
					End If

				Else
					'--- RHCP,LHCP, Ludwig 3 (Theta-> Horizontal; Phi-> Vertical)
					cst_theta_am = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 horizontal")
					cst_theta_ph = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 hor. phase")
					cst_phi_am   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 vertical")
					cst_phi_ph   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 ver. phase")

					cst_theta_re = cst_theta_am * CosD(cst_theta_ph)
					cst_theta_im = cst_theta_am * SinD(cst_theta_ph)
					cst_phi_re = cst_phi_am * CosD(cst_phi_ph)
					cst_phi_im = cst_phi_am * SinD(cst_phi_ph)



					If cst_icomp = 2 Then
						'--- RHCP, LHCP.

						L3CoCx_To_RhcpLhcp(cst_theta_re, cst_theta_im, cst_phi_re, cst_phi_im, cst_LHCP_re, cst_LHCP_im, cst_RHCP_re, cst_RHCP_im)

						'first export RHCP component, then LHCP component
						cst_theta_re=cst_RHCP_re
						cst_theta_im=cst_RHCP_im
						cst_phi_re=cst_LHCP_re
						cst_phi_im=cst_LHCP_im

					End If

				End If


				'--- renorm amplitude and phase for the grasp input file (Radiated Peak Power: 4 Pi Watts)
				cst_theta_re = cst_theta_re * cst_ffam_renorm
				cst_theta_im = cst_theta_im * cst_ffam_renorm
				cst_phi_re = cst_phi_re * cst_ffam_renorm
				cst_phi_im = cst_phi_im * cst_ffam_renorm
				cst_rad_re = cst_rad_re * cst_ffam_renorm
				cst_rad_im = cst_rad_im * cst_ffam_renorm

				'cst_theta_ph = cst_theta_ph + cst_phase_renorm
				'cst_phi_ph = cst_phi_ph + cst_phase_renorm

				'cst_phi_am = Sqr(cst_phi_re^2 + cst_phi_im^2)
				'cst_phi_ph = ATn2D(cst_phi_im,cst_phi_re)

				'--- write out farfield patterns for theta variation
				If dlg.FarfNearf = 1 Then
					Print #CST_FN, " " + cstreh_Pretty(cst_theta_re) + cstreh_Pretty(cst_theta_im) + cstreh_Pretty(cst_phi_re) + cstreh_Pretty(cst_phi_im)
				Else
					Print #CST_FN, " " + cstreh_Pretty(cst_theta_re) + cstreh_Pretty(cst_theta_im) + cstreh_Pretty(cst_phi_re) + cstreh_Pretty(cst_phi_im) + cstreh_Pretty(cst_rad_re) + cstreh_Pretty(cst_rad_im)
				End If

				cst_iff = cst_iff + 1

			'used Dround to overcome numerical error
			cst_theta = DRound (cst_theta,2)
			Next cst_theta

			cst_phi = DRound (cst_phi,2)
			Next cst_phi

		Close #CST_FN

	Next cst_ifloop

	cst_time = Timer - cst_time
	'MsgBox "Time/s = " + CStr(cst_time),vbOkOnly,"Time Check"
	Exit All

ERROR_NO_FARFIELDS:
	MsgBox("No farfield results available !!",vbOkOnly+vbCritical,"Macro Execution Stopped")
	Exit Sub

End Sub
Sub cst_F_tpcalc(ByVal theta As Double, ByVal phi As Double, ByRef tout As Double, ByRef pout As Double)

	While theta < 0
		theta = theta + 360
	Wend
	While theta > 360
		theta = theta - 360
	Wend

	If theta > 180 Then
		theta = 360 - theta
		phi = phi + 180
	End If

	While phi < 0
		phi = phi + 360
	Wend
	While phi > 360
		phi = phi - 360
	Wend

	tout = theta
	pout = phi

End Sub
Sub L3CoCx_To_RhcpLhcp(ByRef L3Eco_re As Double, ByRef L3Eco_im As Double, ByRef L3Ecx_re As Double, ByRef L3Ecx_im As Double, ByRef ELHCP_re As Double, ByRef ELHCP_im As Double, ByRef ERHCP_re As Double, ByRef ERHCP_im As Double)

	' Conversion of Ludwig 3 Co- and Crosspolar components into Left- and Right Hand Circular Polarisation Components

	ELHCP_re = 1/Sqr(2) * (L3Eco_re + L3Ecx_im)
	ELHCP_im = 1/Sqr(2) * (L3Eco_im - L3Ecx_re)

	ERHCP_re = 1/Sqr(2) * (L3Eco_re - L3Ecx_im)
	ERHCP_im = 1/Sqr(2) * (L3Eco_im + L3Ecx_re)

End Sub
Function cstreh_Pretty(x As Variant) As String

        cstreh_Pretty = Replace(Left$(IIf(Left$(CStr(x), 1) = "-", "", " ") + Format(x,"0.0000000000E+00") + String$(16, " "), 18), ",", ".")

End Function
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean

    Select Case Action%
	    Case 1 ' Dialog box initialization
				DlgEnable "xcenter",False
				DlgEnable "ycenter",False
				DlgEnable "zcenter",False
				DlgText  "xcenter","0"
				DlgText  "ycenter","0"
				DlgText  "zcenter","0"
				DlgValue  "FFall",1
				DlgEnable "FarfieldSelect",False
				DlgValue "CheckPolAlignment",0
				DlgText "xpol","0"
				DlgText "ypol","1"
				DlgText "zpol","0"
				DlgEnable "xpol",False
				DlgEnable "ypol",False
				DlgEnable "zpol",False
				DlgValue "FarfNearf",1
				DlgText "radius","1"
				DlgEnable "radius",False
	    Case 2 ' Value changing or button pressed
	    	Select Case DlgItem$
				Case "Help"
					StartHelp "common_preloadedmacro_farfield_grasp_feed_file_export"
					dialogfunc = True
		    	Case "GroupCenterBox"
		    		If SuppValue = 2 Then
						DlgEnable "xcenter",True
						DlgEnable "ycenter",True
						DlgEnable "zcenter",True
						dialogfunc = True
		    		Else
						DlgEnable "xcenter",False
						DlgEnable "ycenter",False
						DlgEnable "zcenter",False
						dialogfunc = True
					End If
				Case "FFall"
					If SuppValue = 1 Then
						DlgValue  "FFall",1
						DlgEnable "FarfieldSelect",False
						dialogfunc = True
					Else
						DlgValue  "FFall",0
						DlgEnable "FarfieldSelect",True
						dialogfunc = True
					End If
				Case "CheckPolAlignment"
					If SuppValue = 1 Then
						DlgValue  "CheckPolAlignment",1
						DlgEnable "xpol",True
						DlgEnable "ypol",True
						DlgEnable "zpol",True
						dialogfunc = True
					Else
						DlgValue  "CheckPolAlignment",0
						DlgEnable "xpol",False
						DlgEnable "ypol",False
						DlgEnable "zpol",False
						dialogfunc = True
					End If
				Case "FarfNearf"
					If SuppValue = 1 Then
						DlgValue "FarfNearf",1
						DlgEnable "radius",False
					Else
						DlgValue "FarfNearf",0
						DlgEnable "radius",True
					End If
	    	End Select
	    Case 4 ' Focus changed
	    Case 5 ' Idle
	    Case 6 ' Function key
    End Select

End Function
