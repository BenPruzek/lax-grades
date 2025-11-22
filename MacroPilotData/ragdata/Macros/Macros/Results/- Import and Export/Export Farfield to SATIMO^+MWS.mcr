'----------------------------------------------------------------------------------------------------
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'----------------------------------------------------------------------------------------------------
' 19-Apr-2021 dta: fixed issue with non regular frequency step (LIST)
' 26-Mar-2021 dta: fixed issue arising for models having more than 1 port when selecting export Broadband
' 04-Oct-2017 hcg: Excluded farfield cuts from selection list
' 15-Dec-2014 dta: added option to normalize exported field to directivity
' 11-Jul-2013 dta: added option to export multifrequency TRXV5 file using excitation string as input
' 13-Jun-2013 dta: corrected problem on how to get correct total radiated power
' 25-Jan-2013 ube: changed scaling factor cst_ffam_renorm to 1
' 10-Jun-2011 ube: some fixes, tested by Satimo, put into official 2011-release-SP4
' 12-Apr-2011 ube: Initial version
'----------------------------------------------------------------------------------------------------
Option Explicit
Public cst_title_ini As String
Public cst_freq() As Double

'#include "vba_globals_all.lib"

Sub Main

	'=============================================================================
	Dim cst_tree() As String, cst_tmpstr As String, cst_excitation_names() As String, cst_temp_excit_names() As String, cst_ff_monitor_names() As String
	Dim nff_per_excit () As Long, index_ff_excit_list () As Long, excit_counter As Long, ff_index As Long

	Dim ExcitAlreadyExist As Boolean
	Dim cst_iloop As Long, cst_iloop2 As Long, index As Long
	Dim cst_nff As Long, cst_nff_monitors As Long, cst_n_of_excitations As Long, cst_temp_n_of_excitations
	Dim cst_ffq(3) As String

	Dim cst_theta_start As Double, cst_theta_step As Double, cst_theta_stop As Double, cst_ntheta As Long
	Dim cst_phi_start As Double, cst_phi_step As Double, cst_phi_stop As Double, cst_nphi As Long
	Dim cst_phi As Double, cst_theta As Double, cst_theta_calc As Double, cst_phi_calc As Double
	Dim cst_origin_type As String
	Dim cst_origin_x As Double, cst_origin_y As Double, cst_origin_z As Double
	Dim cst_dt As Double, cst_dp As Double

	Dim cst_icomp As Integer, cst_ncomp As Integer
	Dim cst_floop_start As Long, cst_floop_end As Long, cst_ifloop As Long, cst_excit_loop_start As Long, cst_excit_loop_end As Long, cst_i_excit_loop As Long

	Dim satimo_header As String, cst_ffname As String, cst_filename As String
	Dim cst_radpow As Double, cst_ffam_renorm As Double
	Dim CST_FN As Long, cst_iff As Long
	Dim cst_theta_ph_renorm As Double, cst_phi_ph_renorm As Double, cst_phase_renorm As Double
	Dim cst_theta_re As Double, cst_theta_im As Double, cst_theta_am As Double, cst_theta_ph As Double
	Dim cst_phi_re As Double, cst_phi_im As Double, cst_phi_am As Double, cst_phi_ph As Double
	Dim cst_th_re As Double, cst_th_im As Double, cst_ph_re As Double, cst_ph_im As Double
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

	'=======calculate number of different excitations=====================


		ReDim cst_temp_excit_names(cst_nff)

			For cst_iloop2 = 0 To cst_nff-1    ' loop over farfield results to get the excitation strings

				index=InStrRev(cst_tree (cst_iloop2), " ")
				cst_temp_excit_names(cst_iloop2)= Mid$(cst_tree(cst_iloop2),index+1)

			Next cst_iloop2



		ReDim cst_excit_names(1)

		cst_excit_names(0)=cst_temp_excit_names(0)			'assign automatically the first excitation
		cst_n_of_excitations=1

		For cst_iloop2 = 0 To cst_nff-1

			ExcitAlreadyExist=False

				For cst_iloop=0 To cst_n_of_excitations-1    ' loop over current array of excitation

					If cst_temp_excit_names(cst_iloop2)=cst_excit_names(cst_iloop) Then
						ExcitAlreadyExist=True
					End If
				Next cst_iloop

					If ExcitAlreadyExist=False Then
						cst_n_of_excitations=cst_n_of_excitations+1
						ReDim Preserve cst_excit_names(cst_n_of_excitations)
						cst_excit_names(cst_n_of_excitations-1)=cst_temp_excit_names(cst_iloop2)
					End If

		Next cst_iloop2

	'=============================================================================
	ReDim nff_per_excit(cst_n_of_excitations-1)    	 'Writes in this array how many farfield results are available in the Tree for each excitation

	ReDim index_ff_excit_list(cst_nff-1)				 'Writes in this array the farfield result indexes for each excitation  e.g. [1st excit. 0,3,7...2nd excit 1,4,8, ecc]

	'=======load farfield result indexes for each excitation==============================

	ff_index=0

	For cst_iloop = 0 To cst_n_of_excitations-1
		excit_counter=0

		For cst_iloop2=0 To cst_nff-1
			If cst_temp_excit_names(cst_iloop2)=cst_excit_names(cst_iloop) Then
				index_ff_excit_list(ff_index)=cst_iloop2
				ff_index=ff_index+1
				excit_counter=excit_counter+1
			End If
		Next cst_iloop2

		If cst_iloop=0 Then
		nff_per_excit(cst_iloop)=excit_counter  'assign n° of farfield results for the 1st excitation
		Else
		nff_per_excit(cst_iloop)=excit_counter+nff_per_excit(cst_iloop-1)    'assign for the following excitation the sum with previous excitation  [e.g. 3,6,9,10,11] 2nd excitation starts at index given by 1st excitation and ends at index given by actual excitation
		End If


	Next cst_iloop

	'=============================================================================

	cst_ffq(0) = "Spherical: theta,phi"
	cst_ffq(1) = "LHCP, RHCP"
	cst_ffq(2) = "Ludwig 3: Vertical, Horizontal"

	cst_ffpol(0) = "theta-phi"
	cst_ffpol(1) = "LHCP-RHCP"
	cst_ffpol(2) = "L3-CoCx"

	Begin Dialog UserDialog 850,312,"Farfield Export to SATIMO (TRXV5 format)",.DialogFunc ' %GRID:10,4,1,1
		OKButton 30,288,90,20
		CancelButton 130,288,90,20
		GroupBox 30,105,470,104,"Angular variation",.GroupBox1
		GroupBox 510,8,320,112,"Farfield Origin",.GroupCenterBox1
		TextBox 100,134,110,20,.thlow
		TextBox 290,134,110,20,.phlow
		GroupBox 510,124,320,60,"Select Polarization Type",.GroupBox4
		TextBox 290,158,110,20,.phhigh
		DropListBox 520,144,300,120,cst_ffq(),.ffq
		TextBox 290,182,110,20,.phstep
		TextBox 100,158,110,20,.thhigh
		TextBox 100,182,110,20,.thstep
		Text 40,138,50,12,"thlow",.Text1
		Text 230,138,50,12,"phlow",.Text4
		Text 230,162,50,12,"phhigh",.Text5
		Text 230,186,50,12,"phstep",.Text6
		Text 40,162,50,12,"thhigh",.Text2
		Text 40,186,50,12,"thstep",.Text3
		OptionGroup .GroupCenterBox
			OptionButton 530,32,290,16,"Center of bounding box",.OptionButton1
			OptionButton 530,52,290,16,"Origin of coordinate system",.OptionButton2
			OptionButton 530,72,290,16,"Free",.OptionButton3
		TextBox 560,92,50,20,.xcenter
		TextBox 650,92,50,20,.ycenter
		TextBox 740,92,50,20,.zcenter
		Text 540,96,10,16,"X",.Text7
		Text 630,96,10,16,"Y",.Text8
		Text 720,96,10,16,"Z",.Text9
		GroupBox 30,8,470,92,"Select type of export",.GroupBox2
		DropListBox 270,70,220,120,cst_tree(),.FarfieldSelect
		DropListBox 40,70,220,120,cst_excit_names(),.ExcitationSelect
		CheckBox 270,48,140,16,"Export all farfields",.FFall
		CheckBox 45,48,170,16,"Export all excitations",.ExcitAll
		GroupBox 510,188,320,64,"Polarisation Vector (y'-axis)",.GroupBox3
		CheckBox 520,204,320,16,"Main Lobe / Polarization Vector Alignment",.CheckPolAlignment
		TextBox 560,224,50,20,.xpol
		TextBox 650,224,50,20,.ypol
		TextBox 740,224,50,20,.zpol
		Text 540,228,10,16,"X",.Text10
		Text 630,228,10,16,"Y",.Text11
		Text 720,228,10,16,"Z",.Text12
		GroupBox 30,215,250,65,"Nearfield / Farfield",.GroupBox5
		GroupBox 290,215,210,65,"Farfield Normalization",.GroupBox6
		CheckBox 50,234,220,16,"Use farfield approximation",.FarfNearf
		TextBox 190,254,50,20,.radius
		Text 50,258,130,16,"Reference distance:",.Text13
		Text 250,258,20,16,"m",.Text14
		OptionGroup .GroupSelect
			OptionButton 45,28,170,16,"Broadband",.OptionButton4
			OptionButton 270,29,170,16,"Single Frequency",.OptionButton5
		CheckBox 300,244,170,24,"Normalize to Directivity",.Norm_Dir
	End Dialog
	Dim dlg As UserDialog




	'--- get registry settings
	dlg.thlow  = GetString("CST STUDIO SUITE", "SatimoExport", "theta_low", "-180.0")
	dlg.thhigh = GetString("CST STUDIO SUITE", "SatimoExport", "theta_high", "180.0")
	dlg.thstep = GetString("CST STUDIO SUITE", "SatimoExport", "theta_step", "45.0")
	dlg.phlow  = GetString("CST STUDIO SUITE", "SatimoExport", "phi_low", "0.0")
	dlg.phhigh = GetString("CST STUDIO SUITE", "SatimoExport", "phi_high", "180.0")
	dlg.phstep = GetString("CST STUDIO SUITE", "SatimoExport", "phi_step", "30.0")

	If (Dialog(dlg) = 0) Then Exit All


	'--- write back registry settings
	SaveString  "CST STUDIO SUITE", "SatimoExport", "theta_low", dlg.thlow
	SaveString  "CST STUDIO SUITE", "SatimoExport", "theta_high", dlg.thhigh
	SaveString  "CST STUDIO SUITE", "SatimoExport", "theta_step", dlg.thstep
	SaveString  "CST STUDIO SUITE", "SatimoExport", "phi_low", dlg.phlow
	SaveString  "CST STUDIO SUITE", "SatimoExport", "phi_high", dlg.phhigh
	SaveString  "CST STUDIO SUITE", "SatimoExport", "phi_step", dlg.phstep

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
	'	icomp=3:	Ludwig3 E_Co / E_Cx (co means vertical (y-axis) and cross means horizontal)
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
		.SetPlotMode "epattern"
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

	cst_phi_stop = cst_phi_stop - cst_phi_step
	'--- write list of farfield/nearfield points
	If dlg.FarfNearf = 1 Then
		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
			For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
				cst_F_tpcalc cst_theta,cst_phi,cst_theta_calc,cst_phi_calc
				FarfieldPlot.AddListItem(cst_theta_calc, cst_phi_calc, 0)
			Next cst_theta
		Next cst_phi
	Else
		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
			For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
				cst_F_tpcalc cst_theta,cst_phi,cst_theta_calc,cst_phi_calc
				FarfieldPlot.AddListItem(cst_theta_calc, cst_phi_calc, cst_rad)
			Next cst_theta
		Next cst_phi
	End If

'============================================================================================
' Loop setting inizialization
'============================================================================================

Dim dfrq_orig As Double,dfrq_end As Double,dfrq_current As Double, dfrq_LIST As String, sname_without_frq_orig As String


	'--- All Excitations or just one
    If dlg.GroupSelect=0 Then


		If dlg.ExcitAll = 0 Then
			cst_excit_loop_start = CInt(dlg.ExcitationSelect)
			cst_excit_loop_end = CInt(dlg.ExcitationSelect)

		Else
			cst_excit_loop_start = 0
			cst_excit_loop_end = cst_n_of_excitations-1
		End If

'loop over the selected excitations
'-----------------------------------------------------------------------------------

		For cst_i_excit_loop = cst_excit_loop_start To cst_excit_loop_end

			If cst_i_excit_loop=0 Then
				cst_floop_start=0
				cst_floop_end =nff_per_excit(cst_excit_loop_start)-1
			Else
				cst_floop_start=nff_per_excit(cst_i_excit_loop-1)
				cst_floop_end =nff_per_excit(cst_i_excit_loop)-1
			End If

		'SATIMO Header written once
		'-----------------------------------------------------------------------------------

			satimo_header = "TRXV5"  + vbCrLf
		If dlg.FarfNearf = 1 Then
			satimo_header = satimo_header + "3"  + vbCrLf
			satimo_header = satimo_header + "4"  + vbCrLf
			satimo_header = satimo_header + "14"  + vbCrLf
			satimo_header = satimo_header + "0"  + vbCrLf
		Else
			satimo_header = satimo_header + "4"  + vbCrLf
			satimo_header = satimo_header + "6"  + vbCrLf
			satimo_header = satimo_header + "10"  + vbCrLf
			satimo_header = satimo_header + "0"  + vbCrLf
		End If


		If Not bSeparateFarfieldFrq(cst_tree(index_ff_excit_list(cst_floop_start)), dfrq_orig, sname_without_frq_orig) Then
				ReportError "Farfield Result Name does not contain frequency term: (f=...)."
				Exit All
		End If

		If Not bSeparateFarfieldFrq(cst_tree(index_ff_excit_list(cst_floop_end)), dfrq_end, sname_without_frq_orig) Then
				ReportError "Farfield Result Name does not contain frequency term: (f=...)."
				Exit All
		End If

		dfrq_LIST=""

		For cst_ifloop = cst_floop_start To cst_floop_end
			If Not bSeparateFarfieldFrq(cst_tree(index_ff_excit_list(cst_ifloop)), dfrq_current, sname_without_frq_orig) Then
				ReportError "Farfield Result Name does not contain frequency term: (f=...)."
				Exit All
			End If
			dfrq_LIST=dfrq_LIST+"	"+ Cstr(dfrq_current*Units.GetFrequencyUnitToSI)
		Next cst_ifloop

		satimo_header = satimo_header + "Frequency"  + vbCrLf

		'satimo_header = satimo_header + "5     "+CStr(cst_floop_end-cst_floop_start+1)+"      " + Cstr(dfrq_orig*Units.GetFrequencyUnitToSI)+"      " + Cstr(dfrq_end*Units.GetFrequencyUnitToSI) + "  LIN" +  vbCrLf

		'write always LIST of frequencies. Single frequecny, multiple frequencies with linear step or non equidistant step
		satimo_header = satimo_header + "5     "+CStr(cst_floop_end-cst_floop_start+1)+"      " + Cstr(dfrq_orig*Units.GetFrequencyUnitToSI)+"      " + Cstr(dfrq_end*Units.GetFrequencyUnitToSI) + "  LIST" +"      " +dfrq_LIST + vbCrLf



		If dlg.FarfNearf = 1 Then
			' farfield
			cst_filename = GetProjectPath("Result") + "CST2Satimo_FFapprox_on_(farfield (f="+CStr(dfrq_orig)+" To "+CStr(dfrq_end)+" ("+CStr(cst_floop_end-cst_floop_start+1)+")) " + cst_excit_names(cst_i_excit_loop) + "_" + cst_ffpol(dlg.ffq) +")).trx"
			satimo_header = satimo_header + "Phi"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_nphi) + "     " + CStr(cst_phi_start*Pi/180) + "     " + CStr(cst_phi_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Theta"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_ntheta+1) + "     " + CStr(cst_theta_start*Pi/180) +  "     " + CStr(cst_theta_stop*Pi/180) + "  LIN" +  vbCrLf
		Else
			' nearfield
			cst_filename = GetProjectPath("Result") + "CST2Satimo_FFapprox_off[r=" + CStr(cst_rad) + "m]_" + "(" + cst_ffname + "_" + cst_ffpol(dlg.ffq) +").trx"
			satimo_header = satimo_header + "R"  + vbCrLf
			satimo_header = satimo_header + "1     1      " + CStr(cst_rad) + "     " + CStr(cst_rad) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Phi"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_nphi) + "     " + CStr(cst_phi_start*Pi/180) + "     " + CStr(cst_phi_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Theta"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_ntheta+1) + "     " + CStr(cst_theta_start*Pi/180) +  "     " + CStr(cst_theta_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "E(R). Real part"  + vbCrLf
			satimo_header = satimo_header + "13"  + vbCrLf
			satimo_header = satimo_header + "E(R). Imaginary part"  + vbCrLf
			satimo_header = satimo_header + "14"  + vbCrLf
		End If

		satimo_header = satimo_header + "E(Phi). Real part"  + vbCrLf
		satimo_header = satimo_header + "13"  + vbCrLf
		satimo_header = satimo_header + "E(Phi). Imaginary part"  + vbCrLf
		satimo_header = satimo_header + "14"  + vbCrLf
		satimo_header = satimo_header + "E(Theta). Real part"  + vbCrLf
		satimo_header = satimo_header + "13"  + vbCrLf
		satimo_header = satimo_header + "E(Theta). Imaginary part"  + vbCrLf
		satimo_header = satimo_header + "14" +  vbCrLf

		If dlg.FarfNearf = 1 Then
			' farfield
			satimo_header = satimo_header + "ETotal . dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000010	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "ETotal . Lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000011	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000013	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000016	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000011	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000017	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000015	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000018	0	1	2	3"
		Else
			' nearfield
			satimo_header = satimo_header + "ETotal . dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC000002F	0	1	2	3	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002 0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002 2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	4	5"
		End If

		CST_FN = FreeFile
		Open cst_filename For Output As #CST_FN

		'ube MsgBox satimo_header
		Print #CST_FN,	satimo_header

	'-------------------------------------------------------------------------------------
	'-------------------------------------------------------------------------------------


		'-----------------------------------------------------------------------------------
		'loop over the the farfields for each excitation
		'-----------------------------------------------------------------------------------
		For cst_ifloop = cst_floop_start To cst_floop_end

			'--- farfield plot init
		SelectTreeItem "Farfields\"+cst_tree(index_ff_excit_list(cst_ifloop))

		' ---get total radiated power. *2 to get peak value----------------
		cst_radpow = FarfieldPlot.GetTRP*2

		If cst_radpow = 0 Then cst_radpow = 1

		' ube 25-jan-2013  changed to 1
		'cst_ffam_renorm = Sqr(4.0*Pi/cst_radpow)/Sqr(Sqr(4*Pi*1e-7/8.8542e-12))
		'cst_ffam_renorm = 1.0

		' dta 13-jun-2013   E(norm in directivity)= Efar/(Sqr(30*radiated power))

		If dlg.Norm_Dir=1 Then    'Field are normalized in Directivity
			cst_ffam_renorm=1/Sqr(Sqr(4*Pi*1e-7/8.8542e-12)*cst_radpow/(4*Pi))
		Else
			cst_ffam_renorm=1
		End If

		FarfieldPlot.CalculateList(cst_tree(index_ff_excit_list(cst_ifloop)))




		cst_iff = 0

		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step

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

					cst_theta_re = -FarfieldPlot.GetListItem(cst_iff,"th_re")
					cst_theta_im = -FarfieldPlot.GetListItem(cst_iff,"th_im")
					cst_phi_re   = FarfieldPlot.GetListItem(cst_iff,"ph_re")
					cst_phi_im   = FarfieldPlot.GetListItem(cst_iff,"ph_im")

					If cst_theta <= 0 Then
						cst_theta_re = -cst_theta_re
						cst_theta_im = -cst_theta_im
						cst_phi_re = -cst_phi_re
						cst_phi_im = -cst_phi_im
					End If

				Else
					'--- RHCP,LHCP, Ludwig 3
					cst_theta_am = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 vertical")
					cst_theta_ph = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 ver. phase")
					cst_phi_am   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 horizontal")
					cst_phi_ph   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 hor. phase")

					cst_theta_re = cst_theta_am * CosD(cst_theta_ph)
					cst_theta_im = cst_theta_am * SinD(cst_theta_ph)
					cst_phi_re = cst_phi_am * CosD(cst_phi_ph)
					cst_phi_im = cst_phi_am * SinD(cst_phi_ph)

					If cst_icomp = 2 Then
						'--- RHCP, LHCP
						cst_th_re = cst_theta_re
						cst_th_im = cst_theta_im
						cst_ph_re = cst_phi_re
						cst_ph_im = cst_phi_im
						L3CoCx_To_RhcpLhcp(cst_th_re, cst_th_im, cst_ph_re, cst_ph_im, cst_theta_re, cst_theta_im, cst_phi_re, cst_phi_im)
					End If

				End If

				'--- renorm amplitude and phase for the satimo input file (Radiated Peak Power: 4 Pi Watts)

				cst_theta_re = cst_theta_re * cst_ffam_renorm
				cst_theta_im = cst_theta_im * cst_ffam_renorm
				cst_phi_re = cst_phi_re * cst_ffam_renorm
				cst_phi_im = cst_phi_im * cst_ffam_renorm
				cst_rad_re = cst_rad_re * cst_ffam_renorm
				cst_rad_im = cst_rad_im * cst_ffam_renorm

				'--- write out farfield patterns for theta variation
				If dlg.FarfNearf = 1 Then
					Print #CST_FN, PPretty(cst_phi_re) + PPretty(cst_phi_im) + PPretty(cst_theta_re) + PPretty(cst_theta_im)
				Else
					Print #CST_FN, PPretty(cst_rad_re) + PPretty(cst_rad_im) + PPretty(cst_phi_re) + PPretty(cst_phi_im) + PPretty(cst_theta_re) + PPretty(cst_theta_im)
				End If

				cst_iff = cst_iff + 1

			Next cst_theta

		Next cst_phi




			Next cst_ifloop

		Close #CST_FN

		Next cst_i_excit_loop

'=============================================================================================
'=============================================================================================

	Else

	'--- All Farfields or just one
	If dlg.FFall = 0 Then
		cst_floop_start = CInt(dlg.FarfieldSelect)
		cst_floop_end = CInt(dlg.FarfieldSelect)
	Else
		cst_floop_start = 0
		cst_floop_end = cst_nff - 1
	End If

	For cst_ifloop = cst_floop_start To cst_floop_end

		satimo_header = "TRXV5"  + vbCrLf
		If dlg.FarfNearf = 1 Then
			satimo_header = satimo_header + "3"  + vbCrLf
			satimo_header = satimo_header + "4"  + vbCrLf
			satimo_header = satimo_header + "14"  + vbCrLf
			satimo_header = satimo_header + "0"  + vbCrLf
		Else
			satimo_header = satimo_header + "4"  + vbCrLf
			satimo_header = satimo_header + "6"  + vbCrLf
			satimo_header = satimo_header + "10"  + vbCrLf
			satimo_header = satimo_header + "0"  + vbCrLf
		End If

		'Dim dfrq_orig As Double, sname_without_frq_orig As String

		If Not bSeparateFarfieldFrq(cst_tree(cst_ifloop), dfrq_orig, sname_without_frq_orig) Then
				ReportError "Farfield Result Name does not contain frequency term: (f=...)."
				Exit All
		End If

		satimo_header = satimo_header + "Frequency"  + vbCrLf
		satimo_header = satimo_header + "5     1      " + Cstr(dfrq_orig*Units.GetFrequencyUnitToSI)+"      " + Cstr(dfrq_orig*Units.GetFrequencyUnitToSI) + "  LIN" +  vbCrLf


		'--- farfield plot init
		SelectTreeItem "Farfields\"+cst_tree(cst_ifloop)

		' ---get total radiated power. *2 to get peak value----------------
		cst_radpow = FarfieldPlot.GetTRP*2

		If cst_radpow = 0 Then cst_radpow = 1

		' ube 25-jan-2013  changed to 1
		'cst_ffam_renorm = Sqr(4.0*Pi/cst_radpow)/Sqr(Sqr(4*Pi*1e-7/8.8542e-12))
		'cst_ffam_renorm = 1.0

		' dta 13-jun-2013   E(norm in directivity)= Efar/(Sqr(30*radiated power))

		If dlg.Norm_Dir=1 Then    'Field are normalized in Directivity
			cst_ffam_renorm=1/Sqr(Sqr(4*Pi*1e-7/8.8542e-12)*cst_radpow/(4*Pi))
		Else
			cst_ffam_renorm=1
		End If

		FarfieldPlot.CalculateList(cst_tree(cst_ifloop))

		cst_ffname = cst_tree(cst_ifloop)

		If dlg.FarfNearf = 1 Then
			' farfield
			cst_filename = GetProjectPath("Result") + "CST2Satimo_FFapprox_on_" + "(" + cst_ffname + "_" + cst_ffpol(dlg.ffq) +").trx"
			satimo_header = satimo_header + "Phi"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_nphi) + "     " + CStr(cst_phi_start*Pi/180) + "     " + CStr(cst_phi_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Theta"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_ntheta+1) + "     " + CStr(cst_theta_start*Pi/180) +  "     " + CStr(cst_theta_stop*Pi/180) + "  LIN" +  vbCrLf
		Else
			' nearfield
			cst_filename = GetProjectPath("Result") + "CST2Satimo_FFapprox_off[r=" + CStr(cst_rad) + "m]_" + "(" + cst_ffname + "_" + cst_ffpol(dlg.ffq) +").trx"
			satimo_header = satimo_header + "R"  + vbCrLf
			satimo_header = satimo_header + "1     1      " + CStr(cst_rad) + "     " + CStr(cst_rad) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Phi"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_nphi) + "     " + CStr(cst_phi_start*Pi/180) + "     " + CStr(cst_phi_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "Theta"  + vbCrLf
			satimo_header = satimo_header + "7     " + CStr(cst_ntheta+1) + "     " + CStr(cst_theta_start*Pi/180) +  "     " + CStr(cst_theta_stop*Pi/180) + "  LIN" +  vbCrLf
			satimo_header = satimo_header + "E(R). Real part"  + vbCrLf
			satimo_header = satimo_header + "13"  + vbCrLf
			satimo_header = satimo_header + "E(R). Imaginary part"  + vbCrLf
			satimo_header = satimo_header + "14"  + vbCrLf
		End If

		satimo_header = satimo_header + "E(Phi). Real part"  + vbCrLf
		satimo_header = satimo_header + "13"  + vbCrLf
		satimo_header = satimo_header + "E(Phi). Imaginary part"  + vbCrLf
		satimo_header = satimo_header + "14"  + vbCrLf
		satimo_header = satimo_header + "E(Theta). Real part"  + vbCrLf
		satimo_header = satimo_header + "13"  + vbCrLf
		satimo_header = satimo_header + "E(Theta). Imaginary part"  + vbCrLf
		satimo_header = satimo_header + "14" +  vbCrLf

		If dlg.FarfNearf = 1 Then
			' farfield
			satimo_header = satimo_header + "ETotal . dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000010	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "ETotal . Lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000011	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000013	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000016	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000011	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000017	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar LC . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000015	0	1	2	3"  + vbCrLf
			satimo_header = satimo_header + "Polar RC . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000018	0	1	2	3"
		Else
			' nearfield
			satimo_header = satimo_header + "ETotal . dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC000002F	0	1	2	3	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp lin"  + vbCrLf
			satimo_header = satimo_header + "0	0xC0000001	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002 0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002 2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Amp dB"  + vbCrLf
			satimo_header = satimo_header + "12	0xC0000002	4	5"  + vbCrLf
			satimo_header = satimo_header + "E(R) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	0	1"  + vbCrLf
			satimo_header = satimo_header + "E(Phi) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	2	3"  + vbCrLf
			satimo_header = satimo_header + "E(Theta) . Phase"  + vbCrLf
			satimo_header = satimo_header + "7	0xC0000003	4	5"
		End If

		CST_FN = FreeFile
		Open cst_filename For Output As #CST_FN

		'ube MsgBox satimo_header
		Print #CST_FN,	satimo_header

		cst_iff = 0

		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step

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

					cst_theta_re = -FarfieldPlot.GetListItem(cst_iff,"th_re")
					cst_theta_im = -FarfieldPlot.GetListItem(cst_iff,"th_im")
					cst_phi_re   = FarfieldPlot.GetListItem(cst_iff,"ph_re")
					cst_phi_im   = FarfieldPlot.GetListItem(cst_iff,"ph_im")

					If cst_theta <= 0 Then
						cst_theta_re = -cst_theta_re
						cst_theta_im = -cst_theta_im
						cst_phi_re = -cst_phi_re
						cst_phi_im = -cst_phi_im
					End If

				Else
					'--- RHCP,LHCP, Ludwig 3
					cst_theta_am = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 vertical")
					cst_theta_ph = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 ver. phase")
					cst_phi_am   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 horizontal")
					cst_phi_ph   = FarfieldPlot.GetListItem(cst_iff,"ludwig 3 hor. phase")

					cst_theta_re = cst_theta_am * CosD(cst_theta_ph)
					cst_theta_im = cst_theta_am * SinD(cst_theta_ph)
					cst_phi_re = cst_phi_am * CosD(cst_phi_ph)
					cst_phi_im = cst_phi_am * SinD(cst_phi_ph)

					If cst_icomp = 2 Then
						'--- RHCP, LHCP
						cst_th_re = cst_theta_re
						cst_th_im = cst_theta_im
						cst_ph_re = cst_phi_re
						cst_ph_im = cst_phi_im
						L3CoCx_To_RhcpLhcp(cst_th_re, cst_th_im, cst_ph_re, cst_ph_im, cst_theta_re, cst_theta_im, cst_phi_re, cst_phi_im)
					End If

				End If

				'--- renorm amplitude and phase for the satimo input file (Radiated Peak Power: 4 Pi Watts)

				cst_theta_re = cst_theta_re * cst_ffam_renorm
				cst_theta_im = cst_theta_im * cst_ffam_renorm
				cst_phi_re = cst_phi_re * cst_ffam_renorm
				cst_phi_im = cst_phi_im * cst_ffam_renorm
				cst_rad_re = cst_rad_re * cst_ffam_renorm
				cst_rad_im = cst_rad_im * cst_ffam_renorm

				'--- write out farfield patterns for theta variation
				If dlg.FarfNearf = 1 Then
					Print #CST_FN, PPretty(cst_phi_re) + PPretty(cst_phi_im) + PPretty(cst_theta_re) + PPretty(cst_theta_im)
				Else
					Print #CST_FN, PPretty(cst_rad_re) + PPretty(cst_rad_im) + PPretty(cst_phi_re) + PPretty(cst_phi_im) + PPretty(cst_theta_re) + PPretty(cst_theta_im)
				End If

				cst_iff = cst_iff + 1

			Next cst_theta

		Next cst_phi

		Close #CST_FN

	Next cst_ifloop

	End If

	cst_time = Timer - cst_time
	'MsgBox "Time/s = " + CStr(cst_time),vbOkOnly,"Time Check"
	Exit All

ERROR_NO_FARFIELDS:
	MsgBox("No farfield results available !!",vbOkOnly+vbCritical,"Macro Execution Stopped")
	Exit Sub

End Sub
Sub cst_F_tpcalc(ByVal theta As Double, ByVal phi As Double, ByRef tout As Double, ByRef pout As Double)

	While theta < -180
		theta = theta + 360
	Wend
	While theta > 180
		theta = theta - 360
	Wend

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
Function PPretty(x As Variant) As String

        PPretty = Replace(Left$(IIf(Left$(CStr(x), 1) = "-", "", " ") + Format(x,"0.0000000000E+00") + String$(16, " "), 18), ",", ".")

End Function
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean

    Select Case Action%
	    Case 1 ' Dialog box initialization
				DlgEnable "ffq",False
				DlgEnable "xcenter",False
				DlgEnable "ycenter",False
				DlgEnable "zcenter",False
				DlgText  "xcenter","0"
				DlgText  "ycenter","0"
				DlgText  "zcenter","0"
				DlgValue  "FFall",1
				DlgEnable "FFall",False
				DlgValue  "ExcitAll",1
				DlgEnable "FarfieldSelect",False
				DlgEnable "ExcitationSelect",False
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
				'ube DlgEnable "ffq",False
	    Case 2 ' Value changing or button pressed
	    	Select Case DlgItem$
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
				Case "GroupSelect"
					If SuppValue = 1 Then
					DlgEnable "FFall",True
					DlgValue "FFall",1
					DlgEnable "ExcitAll",False
					DlgEnable "ExcitationSelect", False
					dialogfunc = True
					Else
					DlgEnable "FFall",False
					DlgEnable "FarfieldSelect",False
					DlgEnable "ExcitAll",True
					DlgValue "ExcitAll",1
					DlgEnable "ExcitationSelect", False
					dialogfunc = True
					End If
				Case "ExcitAll"
					If SuppValue = 1 Then
						DlgValue  "ExcitAll",1
						DlgEnable "ExcitationSelect",False
						dialogfunc = True
					Else
						DlgValue  "ExcitAll",0
						DlgEnable "ExcitationSelect",True
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
Function bSeparateFarfieldFrq(ffname As String, dfrq As Double, sname_without_frq As String) As Boolean
	' This function checks if a meaningful expression "(f=x.xxx)" can be found in ffname, where x.xxx represents a frequency
	' Input: ffname
	' Output: dfrg: the frequency entry found, sname_without_frq: ffname with the frequency entry removed
	' Returns "True" if a frequency has been identified, "False" otherwise
	Dim i2, i3, i4 As Integer

	bSeparateFarfieldFrq = False
	sname_without_frq = ""

	i2 = InStr(ffname,"(f=")
	If i2>0 Then
		' frequency start found
		i3 = InStr(Mid(ffname,i2+3,),")")

		If i3>1 Then
			' meaningful frequency end found
			bSeparateFarfieldFrq = True
			dfrq = Evaluate(Mid(ffname,i2+3,i3-1))
			sname_without_frq = Left(ffname,i2+2) + Mid(ffname,i2+i3+2)
		End If

	End If

End Function
