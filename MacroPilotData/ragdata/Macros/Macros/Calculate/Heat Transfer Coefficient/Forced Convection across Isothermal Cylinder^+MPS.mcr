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
	'
	' Coefficient and constants for empirical correlations
	Dim cst_coef As Double
	Dim cst_exp As Double


	' initialize plate
	cst_l = 0.01
	cst_w = 0.005
	cst_Tp = 30
	' initialize flow
	cst_k = 0.024
	cst_U = 1
	cst_Tamb = 20
	cst_Tf = (cst_Tp+cst_Tamb)/2
	' initialize fluid
	cst_mu = 1.84*(10^(-5))
	cst_Cp = 1007
	cst_rho = 1.2
	' set turbulent transition
	cst_ReL = 5*10^5
	'
	cst_len_units = Units.GetUnit("Length")
	cst_conv = Units.GetGeometryUnitToSI
	While True
		cst_A = cst_l*cst_w*pi
		cst_Re = cst_rho*cst_U*cst_l/cst_mu
		cst_Pr = (cst_mu/cst_rho)/(cst_k/(cst_Cp*cst_rho))
		
		' initial values below are valid only for 0.04 <= cst_Re < 4
		cst_coef = 0.989
		cst_exp = 0.330
		If cst_Re >= 4 And cst_Re < 40 Then
		   cst_coef = 0.911
		   cst_exp = 0.385
		ElseIf cst_Re >= 40 And cst_Re < 4000 Then
		   cst_coef = 0.683
		   cst_exp = 0.466
		ElseIf cst_Re >= 4000 And cst_Re < 40000 Then
		   cst_coef = 0.193
		   cst_exp = 0.618
		ElseIf cst_Re >= 40000 And cst_Re < 400000 Then
		   cst_coef = 0.0266
		   cst_exp = 0.805
		End If
		cst_Nu = cst_coef * cst_Re^cst_exp * cst_Pr^(1/3)

		If cst_Re > 4*10^5 Or cst_Re < 0.4 Then
			MsgBox("Reynold's Number is out of empirical correlation range")
		End If

		cst_h = cst_Nu*cst_k/cst_l
		cst_q = cst_h*cst_A*(cst_Tp-cst_Tamb)
		cst_l = cst_l/cst_conv
		cst_w = cst_w/cst_conv
	Begin Dialog UserDialog 580,462,"Forced Convection across Isothermal Cylinder",.dialogfunc ' %GRID:10,7,1,1
		GroupBox 20,336,550,77,"Results",.GroupBox1
		GroupBox 20,203,270,119,"Fluid Properties (@Film Temperature)",.GroupBox2
		GroupBox 20,91,270,98,"Rod Properties",.GroupBox3
		GroupBox 20,7,270,77,"Flow Properties",.GroupBox4
		Text 40,287,150,14,"Specific heat [J/kg-K]:",.FlowDensitykgm6
		TextBox 195,266,80,21,.fluidk
		TextBox 195,287,80,21,.fluidCp
		Text 40,266,140,14,"Conductivity [W/m-K]:",.FlowDensitykgm5
		TextBox 195,224,80,21,.fluidrho
		Text 40,224,140,14,"Fluid Density [kg/m3]:",.FlowDensitykgm3
		Text 40,245,140,14,"Viscosity [kg/m-s]:",.FlowDensitykgm4
		TextBox 195,112,80,21,.PlateL
		TextBox 195,133,80,21,.plateW
		Text 40,112,150,14,"Rod Diameter (d) ["+cst_len_units+"]:",.PlateLength
		Text 40,133,120,14,"Rod length (L) ["+cst_len_units+"]:",.PlateWidth
		Text 40,154,150,14,"Rod Temperature [C]:",.plateTemp
		TextBox 195,154,80,21,.plateT
		TextBox 195,25,80,21,.flow_U
		Text 40,28,150,14,"Flow Velocity (U) [m/s]:",.Text3
		Text 40,49,140,14,"Flow temperature [C]: ",.Text4
		TextBox 195,46,80,21,.flowtemp
		TextBox 195,245,80,21,.fluidmu
		Picture 300,98,260,175,GetInstallPath + "\Library\Macros\Calculate\Heat Transfer Coefficient\CylFlowStreamlines.bmp",0,.Picture1
		Text 60,357,480,21,"Heat Transfer Coeffecient: " + Format(cst_h,"Standard") +  " W/m2-K",.Text1
		Text 60,385,470,21,"Total Heat Transfer: " + Format(cst_q,"Standard") +  " W",.Text6
		PushButton 20,427,90,21,"Calculate",.PushButton1
		CancelButton 120,427,90,21
		PushButton 480,427,90,21,"Help",.helpButton1
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
    	Case "helpButton1"
   			StartHelp("common_preloadedmacro_calculate_Heat_Transfer_Coefficient")
        	dialogfunc = True
    	End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
