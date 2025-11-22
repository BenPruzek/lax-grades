Option Explicit
'#include "vba_globals_all.lib"
' Heat Flux Calculator

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' 20-Jan-2021 mha: In case of cst_power = 0 and cst_q < cst_q_in, changed hardcoded upper limit for temperature to cst_Tp + 50; also re-arranged order of textboxes.
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
	Dim cst_rho As Double
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
	Dim cst_config As Integer
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
	Dim cst_conv As Double
	Dim cst_len_unit As String
	Dim cst_q_in As Double
	' initialize plate
	cst_l = 0.5
	cst_w = 0.1
	cst_Tp = 50
	cst_E = 0.85
	cst_boltz = 5.6703*10^(-8)
	cst_q_in = 10
	' initialize flow
	cst_k = 0.026
	cst_Tamb = 20
	cst_g = 9.81
	' initialize fluid
	cst_mu = 1.84*(10^(-5))
	cst_Cp = 1005
	cst_rho = 1.204
	cst_phase = 0
	cst_power = 0
	' Set initial Configuration
	cst_config = 0
	cst_len_unit = Units.GetUnit("Length")
	cst_conv = Units.GetGeometryUnitToSI
	While True
		cst_Tf = (cst_Tp+cst_Tamb)/2
		If cst_phase = 0 Then
			cst_B = 1/(cst_Tf+273.15) ' make an isobaric approximation of COTE
		End If
		cst_rhoFilm = cst_rho  ' input density is supposed to be at film temperature
		cst_nu      = cst_mu/cst_rhoFilm ' fluid kinematic viscosity
		cst_a       = cst_k/(cst_rhoFilm*cst_Cp) ' fluid thermal diffusivity
		cst_Ar      = cst_l*cst_w ' plate area
		cst_Le      = cst_Ar/(2*cst_w + 2*cst_l)
		cst_Pr      = cst_nu/cst_a ' fluid prandtl number
		cst_Gr      = cst_g*cst_B*(cst_Le^3)*Abs(cst_Tp - cst_Tamb)/(cst_nu^2)
		cst_Ra      = cst_Gr*cst_Pr ' calculate Rayleigh Number
		If cst_config = 0 Then
			' calculate the nusselt number for the up plate configurations
			cst_Nus = 0.54*(cst_Ra^(0.25))		' valid for 2*10^4 < cst_Ra < 10^6
			If cst_Ra >= 10^6 And cst_Ra < 10^11 Then
				cst_Nus = 0.15*(cst_Ra)^(1/3)
			End If
			If cst_Ra < 2*10^4  Or cst_Ra > 10^11 Then
				MsgBox("Rayleigh number is out of empirical correlation range")
			End If
		ElseIf cst_config = 1 Then
			' calculate the nusselt number for the down plate configuration
			cst_Nus = 0.27*(cst_Ra^(0.25))
			If cst_Ra < 10^5 Or cst_Ra > 10^11 Then
				MsgBox("Rayleigh number is out of empirical correlation range")
			End If
		End If

		cst_h = cst_Nus*cst_k/cst_Le
		cst_q_con = cst_h*cst_Ar*Abs(cst_Tp-cst_Tamb)
		cst_q_rad = cst_E*cst_boltz*cst_Ar*Abs((cst_Tp+273.15)^4 - (cst_Tamb+273.15)^4)
		cst_q = cst_q_con + cst_q_rad
		If cst_Tp<cst_Tamb Then
			cst_q = - cst_q
			cst_q_con = -cst_q_con
			cst_q_rad = -cst_q_rad
		End If

		If cst_power = 0 Then			' In dialog options "Power" was chosen
			If cst_q > cst_q_in Then	' cst_q -> convective + radiative power; cst_q_in -> power entered by user.
				Low  = -273
				High = cst_Tp			' cst_Tp -> plate temperature in dialog box
			Else
				Low  = cst_Tp
				High = cst_Tp + 50
			End If

			While Low <= High
				m = Round((Low+High)/2,3)
				cst_Tf = (m+cst_Tamb)/2
				If cst_phase = 0 Then
					cst_B = 1/(cst_Tf+273.15) ' make an isobaric approximation of COTE
				End If
				cst_rhoFilm = cst_rho  ' input density is supposed to be at film temperature
				cst_nu      = cst_mu/cst_rhoFilm ' fluid kinematic viscosity
				cst_a       = cst_k/(cst_rhoFilm*cst_Cp) ' fluid thermal diffusivity
				cst_Ar      = cst_l*cst_w ' plate area
				cst_Le      = cst_Ar/(2*cst_w + 2*cst_l)
				cst_Pr      = cst_nu/cst_a ' fluid prandtl number
				cst_Gr      = cst_g*cst_B*(cst_Le^3)*Abs(m - cst_Tamb)/(cst_nu^2)
				cst_Ra      = cst_Gr*cst_Pr ' calculate Rayleigh Number

				If cst_config = 0 Then
					' calculate the nusselt number for the up plate configurations
					cst_Nus = 0.54*(cst_Ra^(0.25))		' valid for 2*10^4 < cst_Ra < 10^6
					If cst_Ra >= 10^6 And cst_Ra < 10^11 Then
						cst_Nus = 0.15*(cst_Ra)^(1/3)
					End If
				ElseIf cst_config = 1 Then
					' calculate the nusselt number for the down plate configuration
					cst_Nus = 0.27*(cst_Ra^(0.25))
				End If

				cst_h     = cst_Nus*cst_k/cst_Le
				cst_q_con = cst_h*cst_Ar*Abs(m-cst_Tamb)
				cst_q_rad = cst_E*cst_boltz*cst_Ar*Abs((m+273.15)^4 - (cst_Tamb+273.15)^4)
				cst_q     = cst_q_con + cst_q_rad

				If m<cst_Tamb Then
					cst_q     = - cst_q
					cst_q_con = -cst_q_con
					cst_q_rad = -cst_q_rad
				End If
				'cst_l = cst_l/cst_conv
				'cst_w = cst_w/cst_conv

				If cst_q <cst_q_in Then
					Low = m + 0.001
				Else
					If cst_q >cst_q_in Then
						High = m - 0.001
					Else
						Exit While
					End If
				End If
			Wend

			cst_Tp = m
			'cst_l = cst_l/cst_conv
			'cst_w = cst_w/cst_conv
		Else
			cst_q_in = Round(cst_q,3)
		End If

		cst_l = cst_l/cst_conv
		cst_w = cst_w/cst_conv

		If cst_config = 0 Then
			If cst_Ra < 2*10^4  Or cst_Ra > 10^11 Then
				MsgBox("Rayleigh number is out of empirical correlation range")
				Exit While
			End If
		ElseIf cst_config = 1 Then
			' calculate the nusselt number for the down plate configuration
			If cst_Ra < 10^5 Or cst_Ra > 10^11 Then
				MsgBox("Rayleigh number is out of empirical correlation range")
			End If
		End If

		Begin Dialog UserDialog 590,567,"Natural Convection From a Horizontal Plate",.dlgfunc ' %GRID:10,7,1,1
			GroupBox 300,427,270,105,"Results",.GroupBox1
			GroupBox 20,252,270,168,"Fluid Properties (@Film Temperature)",.GroupBox2
			GroupBox 310,252,260,168,"Plate Specifications",.GroupBox3
			GroupBox 20,14,550,231,"Flow Specfications",.GroupBox4
			GroupBox 20,427,270,105,"Specify",.GroupBox5

			Picture 40,63,210,154,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\hup-isot.bmp",0,.Picture1
			Picture 320,63,210,154,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\hdp-isot.bmp",0,.Picture2

			OptionGroup .Group1
				OptionButton 50,224,130,14,"Up plate",.up
				OptionButton 320,224,130,14,"Down plate",.Down
			OptionGroup .power_it
				OptionButton 30,448,160,21,"Power [W]:",.iterate
				OptionButton 30,476,160,21,"Temperature [C]:",.single_calc
			OptionGroup .phase
				OptionButton 60,392,50,14,"gas",.gas
				OptionButton 160,392,70,14,"liquid",.liq

			Text     50,  35, 140, 21,"Flow temperature [C]: ",.Text4
			TextBox 190,  32,  70, 21,.flowtemp
			Text     40, 273, 140, 14,"Fluid Density [kg/m3]:",.FlowDensitykgm3
			TextBox 190, 273,  70, 21,.fluidrho
			Text     40, 294, 140, 14,"Viscosity [kg/m-s]:",.FlowDensitykgm4
			TextBox 190, 294,  70, 21,.fluidmu
			Text     40, 315, 140, 14,"Conductivity [W/m-K]:",.FlowDensitykgm5
			TextBox 190, 315,  70, 21,.fluidk
			Text     40, 336, 150, 14,"Specific heat [J/Kg-K]:",.FlowDensitykgm6
			TextBox 190, 336,  70, 21,.fluidCp
			Text     40, 357, 140, 28,"Thermal Expansion Coefficient:",.Text2
			TextBox 190, 357,  70, 21,.beta
			Text    330, 273, 130, 14,"Length ["+cst_len_unit+"]:",.PlateLength
			TextBox 480, 273,  70, 21,.PlateL
			Text    330, 294, 120, 14,"Width ["+cst_len_unit+"]:",.PlateWidth
			TextBox 480, 294,  70, 21,.plateW
			Text    330, 322, 140, 14,"Surface Emissivity:",.Text3
			TextBox 480, 315,  70, 21,.emis
			TextBox 200, 448,  70, 21,.PlateQ
			TextBox 200, 476,  70, 21,.plateT

			Text    320, 448, 220,14,"Heat Transfer Coeffecient: " + Format(cst_h,"Standard") +  " W/m2-K",.Text1
			Text    320, 469, 220,14,"Convective Heat Transfer: " +Format(cst_q_con,"Standard") +  " W",.Text7
			Text    320, 490, 220,14,"Radiative Heat Transfer: " +Format(cst_q_rad,"Standard") +  " W",.Text8
			Text    320, 511, 220,14,"Total Heat Transfer: " + Format(cst_q,"Standard") +  " W",.Text6

			'Text 320,315,150,14,"Plate Temperature [C]:",.plateTemp
			PushButton    20, 539, 90, 21,"Calculate",.PushButton1
			CancelButton 120, 539, 90, 21
			PushButton   460, 539, 90, 21,"Help",.helpButton
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
		dlg.PlateQ   = cstr(cst_q_in)
		dlg.power_it = cstr(cst_power)

		dlg.flowtemp = CStr(cst_Tamb)
		dlg.Group1   = cst_config
		dlg.phase    = cst_phase
		dlg.beta     = cstr(cst_B)

		If (Dialog(dlg) = 0) Then Exit All

		cst_k   = Evaluate(dlg.fluidk)
		cst_Cp  = Evaluate(dlg.fluidCp)
		cst_rho = Evaluate(dlg.fluidrho)
		cst_mu  = Evaluate(dlg.fluidmu)

		cst_l  = Evaluate(dlg.plateL)*cst_conv
		cst_w  = Evaluate(dlg.plateW)*cst_conv
		cst_Tp = Evaluate(dlg.plateT)
		cst_E  = Evaluate(dlg.emis)

		cst_Tamb   = Evaluate(dlg.flowtemp)
		cst_config = dlg.Group1
		cst_phase  = dlg.phase
		cst_B      = evaluate(dlg.beta)
		cst_q_in   = Evaluate(dlg.PlateQ)
		cst_power  = Evaluate(dlg.power_it)
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
