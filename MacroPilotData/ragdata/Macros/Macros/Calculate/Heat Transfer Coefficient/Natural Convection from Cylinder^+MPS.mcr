Option Explicit
'#include "vba_globals_all.lib"
' Heat Flux Calculator

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' 09-Jan-2019 mtn: Changed to allow matching desired power output of source, adjusted dialog box to fit new function.
' 07-Feb-2018 ube,rsh: Improved formatting of the help text and did some adjustments on the layout of the graphical user interface.

Sub Main ()
	' Plate parameters
	Dim cst_l As Double
	Dim cst_w As Double
	Dim cst_Ar As Double
	Dim cst_Le As Double
	Dim cst_Tp As Double
	Dim cst_E As Double
	Dim cst_boltz As Double
	' Flow Parameters
	Dim cst_B As Double
	Dim cst_Tamb As Double
	Dim cst_Tf As Double
	Dim cst_g As Double
	' Fluid Parameters
	Dim cst_mu As Double
	Dim cst_Cp As Double
	Dim cst_k As Double
	Dim cst_rhoRef As Double
	Dim cst_a As Double
	Dim cst_nu As Double
	Dim cst_rhoFilm As Double
	' Non-dimensional parameters
	Dim cst_Pr As Double
	Dim cst_Gr As Double
	Dim cst_Ra As Double
	Dim cst_Nus As Double
	Dim cst_RaT As Double
	Dim cst_X As Double
	Dim cst_phase As Integer
	Dim cst_power As Integer
	Dim cst_config As Integer
	Dim Low As Double
	Dim High As Double
	Dim m As Double
	' Results
	Dim cst_h As Double
	Dim cst_q As Double
	Dim cst_q_con As Double
	Dim cst_q_rad As Double
	Dim cst_q_in As Double

	'
	Dim cst_pi As Double
	Dim cst_conv As Double
	Dim cst_len_units As String
	' initialize plate
	cst_l = 0.5
	cst_w = 0.1
	cst_Tp = 30
	cst_E = 0.85
	cst_boltz = 5.67*10^(-8)
	cst_q_in = 0
	' initialize flow
	cst_k = 0.026
	cst_Tamb = 20
	cst_g = 9.81
	' initialize fluid
	cst_mu = 1.84*(10^(-5))
	cst_Cp = 1005
	cst_rhoRef = 1.204
	cst_phase = 0
	cst_power = 0
	'
	cst_pi = 4*Atn(1)
	cst_conv = Units.GetGeometryUnitToSI
	cst_len_units = Units.GetUnit("Length")
	While True


		cst_Tf = (cst_Tp+cst_Tamb)/2
		If cst_phase = 0 Then
			cst_B = 1/(cst_Tf+273.15) ' make an isobaric approximation of COTE
		End If
		cst_rhoFilm = cst_rhoRef*(1-cst_B*(cst_Tf-cst_Tamb))
		cst_nu = cst_mu/cst_rhoFilm ' fluid kinematic viscosity
		cst_a = cst_k/(cst_rhoFilm*cst_Cp) ' fluid thermal diffusivity
		cst_Ar = cst_l*cst_w*cst_pi ' plate area
		cst_Le = cst_l
		cst_Pr = cst_Cp*cst_mu/cst_k ' fluid prandtl number
		cst_Gr = cst_g*cst_B*(cst_Le^3)*Abs(cst_Tp - cst_Tamb)/(cst_nu^2)
		cst_Ra = cst_Gr*cst_Pr ' calculate Rayleigh Number

		'Mills, A.F., "Heat and Mass Transfer," Eqs.4.87-88, ISBN 0-256-11443-9, 1995.
		'If cst_Ra > 10^9 Then
			'cst_X = (1 + (0.559/cst_Pr)^(9/16))^(16/9)
			'cst_Nus = (0.6+0.387*(cst_Ra/cst_X)^(1/6))^2
		'Else
		    'cst_X = (1 + (0.559/cst_Pr)^(9/16))^(4/9)
			'cst_Nus = 0.36 + (.518*cst_Ra^(1/4))/cst_X
		'End If

		'Incropera, De Witt., "Fundamentals of Heat and Mass Transfer," 3rd ed., John Wiley & Sons, p551, eq 9.34, 1990.
		'Churchill, S. W., and H. H. S. Chu, "Correlating Equations for Laminar and Turbulent Free Convection from a Horizontal Cylinder," Int. J. Heat Mass Transfer, 18, 1049, 1974.
		cst_X = (1 + (0.559/cst_Pr)^(9/16))^(16/9)
		cst_Nus = (0.6+0.387*(cst_Ra/cst_X)^(1/6))^2

		cst_h = cst_Nus*cst_k/cst_Le
		cst_q_con = cst_h*cst_Ar*(cst_Tp-cst_Tamb)
		cst_q_rad = cst_E*cst_boltz*cst_Ar*((cst_Tp+273.15)^4 - (cst_Tamb+273.15)^4)
		cst_q = cst_q_con + cst_q_rad
		cst_l = cst_l/cst_conv
		cst_w = cst_w/cst_conv

		If cst_power = 0 Then

			If cst_q >cst_q_in Then
				Low = -273
				High = cst_Tp
				Else
				Low = cst_Tp
				High = 10000
			End If

		While Low <= High

			m = Round((Low+High)/2,3)
			cst_Tf = (m+cst_Tamb)/2
			If cst_phase = 0 Then
				cst_B = 1/(cst_Tf+273.15) ' make an isobaric approximation of COTE
			End If
			cst_rhoFilm = cst_rhoRef*(1-cst_B*(cst_Tf-cst_Tamb))
			cst_nu = cst_mu/cst_rhoFilm ' fluid kinematic viscosity
			'cst_a = cst_k/(cst_rhoFilm*cst_Cp) ' fluid thermal diffusivity
			'cst_Ar = cst_l*cst_w*cst_pi ' plate area
			'cst_Le = cst_l
			'cst_Pr = cst_Cp*cst_mu/cst_k ' fluid prandtl number
			cst_Gr = cst_g*cst_B*(cst_Le^3)*Abs(m - cst_Tamb)/(cst_nu^2)
			cst_Ra = cst_Gr*cst_Pr ' calculate Rayleigh Number

			'Mills, A.F., "Heat and Mass Transfer," Eqs.4.87-88, ISBN 0-256-11443-9, 1995.
			'If cst_Ra > 10^9 Then
				'cst_X = (1 + (0.559/cst_Pr)^(9/16))^(16/9)
				'cst_Nus = (0.6+0.387*(cst_Ra/cst_X)^(1/6))^2
			'Else
				'cst_X = (1 + (0.559/cst_Pr)^(9/16))^(4/9)
				'cst_Nus = 0.36 + (.518*cst_Ra^(1/4))/cst_X
			'End If

			'Incropera, De Witt., "Fundamentals of Heat and Mass Transfer," 3rd ed., John Wiley & Sons, p551, eq 9.34, 1990.
			'Churchill, S. W., and H. H. S. Chu, "Correlating Equations for Laminar and Turbulent Free Convection from a Horizontal Cylinder," Int. J. Heat Mass Transfer, 18, 1049, 1974.
			'cst_X = (1 + (0.559/cst_Pr)^(9/16))^(16/9)
			cst_Nus = (0.6+0.387*(cst_Ra/cst_X)^(1/6))^2

			cst_h = cst_Nus*cst_k/cst_Le
			cst_q_con = cst_h*cst_Ar*(m-cst_Tamb)
			cst_q_rad = cst_E*cst_boltz*cst_Ar*((m+273.15)^4 - (cst_Tamb+273.15)^4)
			cst_q = cst_q_con + cst_q_rad
			'cst_l = cst_l/cst_conv
			'cst_w = cst_w/cst_conv
			If  cst_q < cst_q_in  Then
				Low = m+0.001
				Else
					If cst_q > cst_q_in  Then
				High =m-0.001
				Else
					Exit While
				End If
			End If
		Wend

		cst_Tp = m
		Else 
			cst_q_in = round(cst_q,3)
		End If


	Begin Dialog UserDialog 700,483,"Natural Convection From Cylinder",.dlgfunc ' %GRID:10,7,1,1
		GroupBox 330,322,360,105,"Results",.GroupBox1
		GroupBox 20,186,260,112,"Fluid Properties (@ Filim Temperature)",.GroupBox2
		GroupBox 20,70,260,91,"Cylinder Specifications",.GroupBox3
		GroupBox 20,14,500,49,"Flow Specification",.GroupBox4
		GroupBox 20,322,300,105,"Specify",.GroupBox5
		OptionGroup .phase
			OptionButton 300,33,60,21,"gas",.gas
			OptionButton 380,33,80,21,"liquid",.liq
		OptionGroup .power_it
			OptionButton 30,349,160,20,"Power [W]:",.iterate
			OptionButton 30,385,160,21,"Temperature [C]:",.single_calc
		TextBox 195,133,70,21,.emis
		Text 360,385,310,14,"Radiative Heat Transfer: " +Format(cst_q_rad,"Standard") +  " W",.Text8
		Text 360,364,310,14,"Convective Heat Transfer: " +Format(cst_q_con,"Standard") +  " W",.Text7
		Text 30,270,150,14,"Specific heat [J/Kg-K]:",.FlowDensitykgm6
		TextBox 195,249,70,21,.fluidk
		TextBox 195,270,70,21,.fluidCp
		Text 30,249,140,14,"Conductivity [W/m-K]:",.FlowDensitykgm5
		TextBox 195,207,70,21,.fluidrho
		Text 30,207,140,14,"Fluid Density [kg/m3]:",.FlowDensitykgm3
		Text 30,228,140,14,"Viscosity [kg/m-s]:",.FlowDensitykgm4
		TextBox 195,91,70,21,.PlateL
		TextBox 195,112,70,21,.plateW
		Text 40,91,148,14,"Diameter (d) ["+cst_len_units+"]:",.PlateLength
		Text 40,112,130,14,"Length (L) ["+cst_len_units+"]:",.PlateWidth
		'Text 60,400,150,14,"Rod Temperature [C]:",.plateTemp
		TextBox 200,386,70,21,.plateT
		Picture 290,77,390,224,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\CylFlowNatConvStreamlines.bmp",0,.Picture2
		Text 40,35,140,14,"Flow temperature [C]: ",.Text4
		TextBox 195,33,70,21,.flowtemp
		TextBox 195,228,70,21,.fluidmu
		Text 360,343,310,14,"Heat Transfer Coeffecient: " + Format(cst_h,"Standard") +  " W/m2-K",.Text1
		Text 360,406,310,14,"Total Heat Transfer: " + Format(cst_q,"Standard") +  " W",.Text6
		'Text 60,364,150,14,"Match Power [W]:",.platePower
		TextBox 200,349,70,21,.plateQ

		Text 40,133,140,14,"Surface Emissivity:",.Text3
		Text 30,301,170,14,"Thermal Expansion Coeff:",.Text2
		TextBox 195,301,70,21,.beta
		PushButton 20,455,90,21,"Calculate",.PushButton1
		CancelButton 120,455,90,21
		PushButton 590,455,90,21,"Help",.helpButton
	End Dialog
		Dim dlg As UserDialog
		dlg.fluidk = CStr(cst_k)
		dlg.fluidCp = CStr(cst_Cp)
		dlg.fluidrho = CStr(cst_rhoRef)
		dlg.fluidmu = CStr(cst_mu)

		dlg.plateL = CStr(cst_l)
		dlg.plateW = CStr(cst_w)
		dlg.plateT = CStr(cst_Tp)
		dlg.emis = CStr(cst_E)
		dlg.flowtemp = CStr(cst_Tamb)
		dlg.phase = cst_phase
		dlg.beta = cstr(cst_B)
		dlg.plateQ = cstr(cst_q_in)
		dlg.power_it = cstr(cst_power)

		If (Dialog(dlg) = 0) Then Exit All

		cst_k = Evaluate(dlg.fluidk)
		cst_Cp = Evaluate(dlg.fluidCp)
		cst_rhoRef = Evaluate(dlg.fluidrho)
		cst_mu = Evaluate(dlg.fluidmu)

		cst_l = Evaluate(dlg.plateL)*cst_conv
		cst_w = Evaluate(dlg.plateW)*cst_conv
		cst_Tp = Evaluate(dlg.plateT)
		cst_E = Evaluate(dlg.emis)
		cst_phase = evaluate(dlg.phase)
		cst_B = evaluate(dlg.beta)
		cst_Tamb = Evaluate(dlg.flowtemp)
		cst_q_in = Evaluate(dlg.plateQ)
		cst_power = Evaluate(dlg.power_it)
	Wend
End Sub

Function dlgfunc(item$, act%, supp?) As Integer
	Dim phase_f As Variant
	Dim dlgItem As String
	Select Case act
	Case 1
		If evaluate(DlgValue("phase")) = 0 Then
			DlgVisible "beta",False
			DlgVisible "Text2", False
		End If
	Case 2
		phase_f = evaluate (DlgValue ("phase"))
			If phase_f = 0 Then
				DlgVisible "beta", False
				DlgVisible "Text2",False
			Else
				DlgVisible "beta", True
				DlgVisible "Text2", True
			End If

			If item = "helpButton" Then
				StartHelp("common_preloadedmacro_calculate_Heat_Transfer_Coefficient")
				dlgfunc = True
			End If
	End Select
End Function
