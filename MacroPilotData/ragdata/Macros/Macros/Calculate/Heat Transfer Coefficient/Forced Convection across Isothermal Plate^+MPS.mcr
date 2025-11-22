Option Explicit
'#include "vba_globals_all.lib"
' Heat Flux Calculator

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' 30-Sep-2021 mha: Fixed a typo in dialog.
' 07-Feb-2018 ube,rsh: Added help and did some adjustments on the layout of the graphical user interface.

Sub Main ()
	' Plate parameters
	Dim cst_l As Double
	Dim cst_w As Double
	Dim cst_A As Double
	Dim cst_Tp As Double
	' Flow Parameters
	Dim cst_U As Double
	Dim cst_Tamb As Double
	Dim cst_Tf As Double
	' Fluid Parameters
	Dim cst_mu As Double
	Dim cst_Cp As Double
	Dim cst_k As Double
	Dim cst_rho As Double
	Dim cst_rhoFilm As Double
	Dim cst_TFilm As Double
	' Non-dimensional parameters
	Dim cst_Pr As Double
	Dim cst_Re As Double
	Dim cst_Nu As Double
	Dim cst_ReL As Double
	' Results
	Dim cst_h As Double
	Dim cst_q As Double
	'
	Dim cst_len_units As String
	Dim cst_conv As Double

	' initialize plate
	cst_l = 0.1
	cst_w = 0.05
	cst_Tp = 30
	' initialize flow
	cst_k = 0.024
	cst_U = 1
	cst_Tamb = 20
	cst_Tf = (cst_Tp+cst_Tamb)/2
	' initialize fluid
	cst_mu = 1.84*(10^(-5))
	cst_Cp = 1007
	cst_rho = 1.204
	' set turbulent transition
	cst_ReL = 5*10^5
	'
	cst_len_units = Units.GetUnit("Length")
	cst_conv = Units.GetGeometryUnitToSI
	While True
		cst_TFilm = (cst_Tamb + cst_Tp)/2
		cst_rhoFilm = cst_rho*(1-(1/(273.15+cst_TFilm))*(cst_TFilm-cst_Tamb))
		cst_A = cst_l*cst_w
		cst_Re = cst_rhoFilm*cst_U*cst_l/cst_mu
		cst_Pr = (cst_mu/cst_rho)/(cst_k/(cst_Cp*cst_rho))
		If cst_Re < cst_ReL Then
			cst_Nu = 0.664*(cst_Re^.5)*(cst_Pr^(1/3))
		End If
		If cst_Re > cst_ReL Then
			cst_Nu = (cst_Pr^(1/3))*(0.664*cst_Re^0.5 + 0.37*((cst_ReL^0.8) - cst_Re^.8)*(cst_Pr)^(1/3))
		End If
		cst_h = cst_Nu*cst_k/cst_l
		cst_q = cst_h*cst_A*(cst_Tp-cst_Tamb)
		cst_l = cst_l/cst_conv
		cst_w = cst_w/cst_conv
	Begin Dialog UserDialog 640,504,"Forced Convection across Isothermal Flat Plate",.dialogfunc ' %GRID:10,7,1,1
		GroupBox 10,385,610,84,"Results",.GroupBox1
		GroupBox 10,14,300,105,"Plate Specifications",.GroupBox2
		GroupBox 330,7,300,112,"Fluid Properties (@Film Temperature)",.GroupBox3
		GroupBox 10,126,620,63,"Flow Specifications",.GroupBox4
		Text 350,91,150,14,"Specific heat [J/kg-K]:",.FlowDensitykgm6
		TextBox 510,70,80,21,.fluidk
		TextBox 510,91,80,21,.fluidCp
		Text 350,70,140,14,"Conductivity [W/m-K]:",.FlowDensitykgm5
		TextBox 510,28,80,21,.fluidrho
		Text 350,28,140,14,"Fluid Density [kg/m3]:",.FlowDensitykgm3
		Text 350,49,120,14,"Viscosity [kg/m-s]:",.FlowDensitykgm4
		TextBox 200,35,80,21,.PlateL
		TextBox 200,56,80,21,.plateW
		Text 30,35,150,14,"Plate Length (L) ["+cst_len_units+"]:",.PlateLength
		Text 30,56,150,14,"Plate Width (W) ["+cst_len_units+"]:",.PlateWidth
		Text 30,77,150,14,"Plate Temperature [C]:",.plateTemp
		TextBox 200,77,80,21,.plateT
		TextBox 200,147,80,21,.flow_U
		Text 30,150,160,14,"Flow Velocity (U) [m/s]:",.Text3
		Text 350,150,140,14,"Flow temperature [C]: ",.Text4
		TextBox 510,147,80,21,.flowtemp
		TextBox 510,49,80,21,.fluidmu
		Picture 50,196,540,182,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\PlateFlowForceConvGlyphs.bmp",0,.Picture1
		Text 30,413,480,14,"Heat Transfer Coeffecient: " + Format(cst_h,"Standard") +  " W/m2-K",.Text1
		Text 30,441,470,14,"Total Heat Transfer: " + Format(cst_q,"Standard") +  " W",.Text6
		PushButton 10,476,90,21,"Calculate",.PushButton1
		CancelButton 110,476,90,21
		PushButton 530,476,90,21,"Help",.helpButton
	End Dialog
		Dim dlg As UserDialog
		dlg.fluidk = CStr(cst_k)
		dlg.fluidCp = CStr(cst_Cp)
		dlg.fluidrho = CStr(cst_rho)
		dlg.fluidmu = CStr(cst_mu)

		dlg.plateL = CStr(cst_l)
		dlg.plateW = CStr(cst_w)
		dlg.plateT = CStr(cst_Tp)

		dlg.flow_U = CStr(cst_U)
		dlg.flowtemp = CStr(cst_Tamb)

		If (Dialog(dlg) = 0) Then Exit All

		cst_k = Evaluate(dlg.fluidk)
		cst_Cp = Evaluate(dlg.fluidCp)
		cst_rho = Evaluate(dlg.fluidrho)
		cst_mu = Evaluate(dlg.fluidmu)

		cst_l = Evaluate(dlg.plateL)*cst_conv
		cst_w = Evaluate(dlg.plateW)*cst_conv
		cst_Tp = Evaluate(dlg.plateT)

		cst_U = Evaluate(dlg.flow_U)
		cst_Tamb = Evaluate(dlg.flowtemp)

	Wend
End Sub

Function dialogfunc(DlgItem$, Action%, SuppValue&) As Boolean 'for the help button
    Select Case Action
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
    	Select Case DlgItem
    	Case "helpButton"
   			StartHelp("common_preloadedmacro_calculate_Heat_Transfer_Coefficient")
        	dialogfunc = True
    	End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
