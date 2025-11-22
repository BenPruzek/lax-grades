Option Explicit
'#include "vba_globals_all.lib"
' Heat Flux Calculator

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' 20-Jan-2021 mha: In case of cst_power = 0 and cst_q < cst_q_in, changed hardcoded upper limit for temperature to cst_Tp + 100; also re-arranged order of textboxes.
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
	Dim cst_Tfilm As Double
	Dim cst_g As Double
	' Fluid Parameters
	Dim cst_mu As Double
	Dim cst_Cp As Double
	Dim cst_k As Double
	Dim cst_rho As Double
	Dim cst_a As Double
	Dim cst_nu As Double
	Dim cst_rhoF As Double
	' Non-dimensional parameters
	Dim cst_Pr As Double
	Dim cst_Gr As Double
	Dim cst_Ra As Double
	Dim cst_Nus As Double
	Dim cst_RaT As Double
	Dim cst_X As Double
	Dim cst_phase As Integer
	Dim cst_power As Integer
	Dim Low As Double
	Dim High As Double
	Dim m As Double
	' Results
	Dim cst_h As Double
	Dim cst_q As Double
	Dim cst_q_rad As Double
	Dim cst_q_con As Double
	Dim cst_q_in As Double
	'
	Dim cst_conv As Double
	Dim cst_len_units As String
	' initialize plate
	cst_l     = 0.5
	cst_w     = 0.1
	cst_Tp    = 30
	cst_E     = 0.85
	cst_boltz = 5.67*10^(-8)
	cst_q_in  = 0
	' initialize flow
	cst_k     = 0.026
	cst_Tamb  = 20
	cst_g     = 9.81
	' initialize fluid
	cst_phase = 0
	cst_mu    = 1.84*(10^(-5))
	cst_Cp    = 1005
	cst_rho   = 1.204
	cst_power = 0

	cst_conv      = Units.GetGeometryUnitToSI
	cst_len_units = Units.GetUnit("Length")
	While True
		cst_Tfilm = (cst_Tp+cst_Tamb)/2
		If cst_phase = 0 Then
			cst_B = 1/(cst_Tfilm+273.15) ' make an isobaric approximation of COTE
		End If
		cst_rhoF = cst_rho  ' input density is supposed to be at film temperature
		cst_nu = cst_mu/cst_rhoF ' fluid kinematic viscosity
		cst_a = cst_k/(cst_rhoF*cst_Cp) ' fluid thermal diffusivity
		cst_Ar = cst_l*cst_w ' plate area
		cst_Le = cst_l
		cst_Pr = cst_nu/cst_a ' fluid prandtl number
		cst_Gr = cst_g*cst_B*(cst_Le^3)*Abs(cst_Tp - cst_Tamb)/(cst_nu^2)
		cst_Ra = cst_Gr*cst_Pr ' calculate Rayleigh Number

		If cst_Ra < 10^9 Then
			cst_Nus = 0.68 + 0.670 * cst_Ra^0.25/(1+(0.492/cst_Pr)^(9/16))^(4/9)
		Else
			cst_Nus = (0.825 + (0.387 * cst_Ra^(1/6))/(1+(0.492/cst_Pr)^(9/16))^(8/27))^2
		End If

		cst_h     = cst_Nus*cst_k/cst_Le
		cst_q_con = cst_h*cst_Ar*(cst_Tp-cst_Tamb)
		cst_q_rad = cst_E*cst_boltz*cst_Ar*((cst_Tp+273.15)^4 - (cst_Tamb+273.15)^4)
		cst_q     = cst_q_con + cst_q_rad
		'cst_l = cst_l/cst_conv
		'cst_w = cst_w/cst_conv

		If cst_power = 0 Then
			If cst_q > cst_q_in Then
				Low  = -273
				High = cst_Tp
			Else
				Low  = cst_Tp
				High = cst_Tp + 100
			End If

			While Low <= High
				m = Round((Low+High)/2,3)
				cst_Tfilm = (m+cst_Tamb)/2
				If cst_phase = 0 Then
					cst_B = 1/(cst_Tfilm+273.15) ' make an isobaric approximation of COTE
				End If
				cst_rhoF = cst_rho  ' input density is supposed to be at film temperature
				cst_nu   = cst_mu/cst_rhoF ' fluid kinematic viscosity
				cst_a    = cst_k/(cst_rhoF*cst_Cp) ' fluid thermal diffusivity
				cst_Ar   = cst_l*cst_w ' plate area
				cst_Le   = cst_l
				cst_Pr   = cst_nu/cst_a ' fluid prandtl number
				cst_Gr   = cst_g*cst_B*(cst_Le^3)*Abs(m - cst_Tamb)/(cst_nu^2)
				cst_Ra   = cst_Gr*cst_Pr ' calculate Rayleigh Number

				If cst_Ra < 10^9 Then
					cst_Nus = 0.68 + 0.670 * cst_Ra^0.25/(1+(0.492/cst_Pr)^(9/16))^(4/9)
				Else
					cst_Nus = (0.825 + (0.387 * cst_Ra^(1/6))/(1+(0.492/cst_Pr)^(9/16))^(8/27))^2
				End If

				cst_h     = cst_Nus*cst_k/cst_Le
				cst_q_con = cst_h*cst_Ar*(m-cst_Tamb)
				cst_q_rad = cst_E*cst_boltz*cst_Ar*((m+273.15)^4 - (cst_Tamb+273.15)^4)
				cst_q     = cst_q_con + cst_q_rad

				'cst_l = cst_l/cst_conv
				'cst_w = cst_w/cst_conv

				If cst_q < cst_q_in Then
					Low = m + 0.001
				Else
					If cst_q > cst_q_in Then
						High = m - 0.001
					Else
						Exit While
					End If
				End If
			Wend

			cst_Tp = m
			cst_l  = cst_l/cst_conv
			cst_w  = cst_w/cst_conv

		Else
			cst_q_in = Round(cst_q,3)
			cst_l = cst_l/cst_conv
			cst_w = cst_w/cst_conv
		End If


		Begin Dialog UserDialog 600,491,"Natural Convection From a Vertical Plate",.dlgfunc ' %GRID:10,7,1,1
			GroupBox  10,   7, 270,  95, "Plate Specifications",.GroupBox1
			GroupBox  10, 117, 270, 161, "Fluid Specifications (@film temperature)",.GroupBox2
			GroupBox  10, 290, 270,  49, "Flow Specifications",.GroupBox3
			GroupBox 290, 351, 290, 100, "Results",.GroupBox4
			GroupBox  10, 351, 270, 100, "Specify",.GroupBox5

			Picture 290,56,280,252,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\vertPlateGlyphs.bmp",0,.Picture2

			OptionGroup .phase
				OptionButton  40, 257, 60, 14, "gas",.gas
				OptionButton 140, 257, 60, 14, "liquid",.liq
			OptionGroup .power_it
				OptionButton 30, 370, 160, 14, "Power [W]:",.iterate
				OptionButton 30, 397, 170, 14, "Temperature [C]:",.single_calc

			Text     30,  33, 135, 14,"Height (h) ["+cst_len_units+"]:",.PlateLength
			TextBox 180,  28,  70, 21,.PlateL
			Text     30,  54, 135, 14," Width (w) ["+cst_len_units+"]:",.PlateWidth
			TextBox 180,  49,  70, 21,.plateW
			Text     30,  75, 130, 14,"Surface Emissivity:",.plateTemp2
			'Text    200, 417, 150, 14,"Plate Temperature [C]:",.plateTemp
			TextBox 180,  70,  70, 21,.emis
			Text     30, 138, 140, 14,"Fluid Density [kg/m3]:",.FlowDensitykgm3
			TextBox 180, 138,  70, 21,.fluidrho
			Text     30, 159, 140, 14,"Viscosity [kg/m-s]:",.FlowDensitykgm4
			TextBox 180, 159,  70, 21,.fluidmu
			Text     30, 180, 140, 14,"Conductivity [W/m-K]:",.FlowDensitykgm5
			TextBox 180, 180,  70, 21,.fluidk
			Text     30, 201, 150, 14,"Specific heat [J/Kg-K]:",.FlowDensitykgm6
			TextBox 180, 201,  70, 21,.fluidCp
			Text     30, 222, 140, 28,"Coefficient of Thermal Expansion:",.Text2
			TextBox 180, 222,  70, 21,.beta
			Text     30, 311, 140, 14,"Flow temperature [C]: ",.Text4
			TextBox 180, 311,  70, 21,.flowtemp
			Text    300, 371, 270, 14,"Heat Transfer Coeffecient: " + Format(cst_h,"Standard") +  " W/m2-K",.Text1
			Text    300, 388, 270, 14,"Convective Heat Transfer: " + Format(cst_q_con,"Standard") +  " W",.Text7
			Text    300, 406, 270, 14,"Radiative Heat Transfer: " + Format(cst_q_rad,"Standard") +  " W",.Text3
			Text    300, 427, 270, 14,"Total Heat Transfer: " + Format(cst_q,"Standard") +  " W",.Text6
			TextBox 200, 370,  70, 21,.plateQ
			TextBox 200, 399,  70, 21,.plateT

			PushButton    10, 458, 90, 21,"Calculate",.PushButton1
			CancelButton 110, 458, 90, 21
			PushButton   480, 458, 90, 21,"Help",.helpButton
		End Dialog

		Dim dlg As UserDialog
		dlg.fluidk   = CStr(cst_k)
		dlg.fluidCp  = CStr(cst_Cp)
		dlg.fluidrho = CStr(cst_rho)
		dlg.fluidmu  = CStr(cst_mu)
		dlg.plateL   = CStr(cst_l)
		dlg.plateW   = CStr(cst_w)
		dlg.plateT   = CStr(cst_Tp)
		dlg.emis     = CStr(cst_E)
		dlg.beta     = CStr(cst_B)
		dlg.phase    = cst_phase
		dlg.plateQ   = cstr(cst_q_in)
		dlg.power_it = cstr(cst_power)

		dlg.flowtemp = CStr(cst_Tamb)

		If (Dialog(dlg) = 0) Then Exit All

		cst_k     = Evaluate(dlg.fluidk)
		cst_Cp    = Evaluate(dlg.fluidCp)
		cst_rho   = Evaluate(dlg.fluidrho)
		cst_mu    = Evaluate(dlg.fluidmu)
		cst_l     = Evaluate(dlg.plateL)*cst_conv
		cst_w     = Evaluate(dlg.plateW)*cst_conv
		cst_Tp    = Evaluate(dlg.plateT)
		cst_E     = Evaluate(dlg.emis)
		cst_B     = Evaluate(dlg.beta)
		cst_phase = dlg.phase
		cst_q_in  = Evaluate(dlg.plateQ)
		cst_power = Evaluate(dlg.power_it)
		cst_Tamb  = Evaluate(dlg.flowtemp)
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
