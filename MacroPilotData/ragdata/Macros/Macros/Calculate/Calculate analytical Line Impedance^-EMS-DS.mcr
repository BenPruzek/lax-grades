'----------------------------------------------------------------------------------------------------------------------------------------------
'The macro calculates the line impedamce for different line types
'----------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 24-Jan-2018 rsh: Fixed problem of jumping values when changing a value ("Line length" or "eps_r" or "Phase shift") and selecting another textbox
' 18-Sep-2013 fsr: Fixed a problem with line length/phase calculation for cases where eps<>eps_eff
' 01-Mar-2013 ube: corrected picture name for asymmetric stripline
' 30-Jan-2013 fsr: Streamlined code further, implemented different formula for asymmetric stripline
' 28-Sep-2012 fsr: Added asymmetric strip line; switched from radio buttons to drop down list; straightened up code
' 08-Jun-2012 fsr: Initial values for parameters were not displayed correctly; added option to build 3D model of calculated line
' 17-Mar-2011 fsr: Made macro DS compatible
' 27-Apr-2010 fde: Added warning message for Thick Stripline if parameter out of range.
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (pervisouly only first macropath was search)
' 09-Sep-2007 fde: included inverted suspended MStrip
' 08-Sep-2007 fde: included suspended MStrip
' 07-Sep-2007 fde: included Differential Lines
' 04-Jan-2007 fde: included phase shift and delay line length
' 21-May-2006 ube: picture passes adjusted to 2006b
' 18-Oct-2005 ube: Included into Online Help
' 30-Sep-2005 ube: bugfix for thin microstrip (one space in string too much)
' 28-Sep-2005 fde: thick stripline included
' 23-Feb-2004 twi: bugfix for thick microstrip
' 12-Mar-2003 fde: changes for latest version SAX basic
' 21-Feb-2002 fde: include dispersion For some cases
' 14-Feb-2002 fde:  beta
'----------------------------------------------------------------------------------------------------------------------------------------------
'Option Explicit <-- This option needs to be OFF for DS compatibility

Public macropath As String
Public cst_Linetype As String
Public cst_dlg_linelength As String
Public cst_dlg_phase As String
Public cst_dlg_freq As String
Public cst_filename As String
Public cst_impedancefile  As String
Public cst_d1 As Double
Public cst_eps As Double
Public cst_eps_eff As Double
Public cst_frequency As Double
Public cst_linelength As Double
Public cst_phase As Double
Public cst_d2 As Double
Public cst_w As Double
Public cst_h As Double
Public cst_g As Double
Public cst_t As Double
Public cst_s As Double
Public cst_a As Double
Public cst_b As Double
Public cst_Dispersion_on As Boolean
Public cst_Dispersion_plot_on As Boolean
Public cst_LineImpedance As Double
Public cst_lauf_index As Integer
Public resultdir As String

Dim dFmin As Double, dFmax As Double

Public Const AvailableLineTypes = Array("Coax", _
											"Stripline", _
											"Thick Stripline", _
											"Asymmetric Thick Stripline", _
											"Differential Stripline", _
											"Thin Microstrip", _
											"Thick Microstrip", _
											"Suspended Microstrip", _
											"Inverted Suspended Microstrip", _
											"Thick Coplanar Waveguide", _
											"Coplanar Waveguide", _
											"Coplanar Waveguide with Ground")

Sub Main

	Dim sAvailableLineTypes() As String, i As Long

	ReDim sAvailableLineTypes(UBound(AvailableLineTypes))
	For i = 0 To UBound(AvailableLineTypes)
		sAvailableLineTypes(i) = AvailableLineTypes(i)
	Next

	cst_Dispersion_on = False

	resultdir     = GetProjectPath ("Result")

 	macropath = GetInstallPath + "\Library\Macros"
	Begin Dialog UserDialog 700,392,"Impedance Calculation",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,680,294,"Setup",.GroupBox1
		DropListBox 30,28,360,192,sAvailableLineTypes(),.LineTypeSelectionDLB
		Picture 30,56,360,231,"Picture1",0,.Picture1
		Text 410,35,80,14,"Length unit:",.Text3
		Text 560,35,40,14,Units.GetUnit("Length"),.Text4
		Text 410,63,90,14,"Frequency:",.Text6
		TextBox 560,56,50,21,.freq_textbox
		Text 620,63,30,14,Units.GetUnit("Frequency"),.Text5
		GroupBox 400,98,270,112,"Geometry Data",.GroupBox2
		Text 420,126,20,14,"l1",.TextParameter1
		Text 540,126,20,14,"l2",.TextParameter2
		Text 420,154,20,14,"l3",.TextParameter3
		Text 540,154,20,14,"l4",.TextParameter4
		TextBox 440,119,90,21,.Parameter1
		TextBox 560,119,90,21,.Parameter2
		TextBox 440,147,90,21,.Parameter3
		TextBox 560,147,90,21,.Parameter4
		Text 450,182,80,14,"Line length:",.Text7
		TextBox 560,175,90,21,.linelength_textbox
		GroupBox 400,210,270,77,"Permittivity",.GroupBox3
		Text 420,238,50,14,"eps_r =",.TextEpsilonRT
		TextBox 480,231,50,21,.EpsilonRT
		CheckBox 420,266,140,14,"Include Dispersion",.Dispersion_on

		GroupBox 10,301,680,56,"Impedance",.GroupBox4
		Text 30,329,40,21,"Z_0 =",.Text1
		TextBox 80,322,90,21,.Lineimpedance
		Text 180,329,40,14,"Ohm",.Text2
		Text 260,329,60,14,"eps_eff =",.Text9
		TextBox 330,322,90,21,.eps_eff_textbox
		Text 470,329,90,14,"Phase shift =",.Text8
		TextBox 560,322,90,21,.phase_textbox

		PushButton 330,364,90,21,"Calculate",.Calculate
		PushButton 420,364,90,21,"Build 3D",.Build3DPB
		PushButton 600,364,90,21,"Help",.Help
		PushButton 510,364,90,21,"Exit",.ExitPB

	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All
End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
    Select Case Action%
    Case 1 ' Dialog box initialization
    	If GetApplicationName = "DS" Then
    		dFmin = 0
    		dFmax = 1
    		DlgEnable("Build3DPB", False)
    	Else
			dFmin = CallByName(Solver, "GetFMin", vbGet) ' equivalent to Solver.GetFMin but will not trigger preprocessor warning
			dFmax = CallByName(Solver, "GetFMax", vbGet) ' equivalent to Solver.GetFMax but will not trigger preprocessor warning
			dFmax = IIf(dFmax = dFmin, dFmin + 1, dFmax) ' Make sure fmin and fmax are not identical to prevent division by zero later
    		DlgEnable("Build3DPB", True)
		End If
			' Coax is default
			cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_coax.bmp"
			DlgSetPicture "Picture1",cst_filename,0
			DlgEnable "Parameter1",True
			DlgEnable "Parameter2",True
			DlgEnable "Parameter3",False
			DlgEnable "Parameter4",False
			DlgEnable "EpsilonRT",True
			DlgText "TextParameter1","d:"
			DlgText "TextParameter2","D:"
			DlgText "TextParameter3",""
			DlgText "TextParameter4",""
			DlgText "Parameter1" ,"1"
			DlgText "Parameter2" ,"2"
			DlgText "Parameter3" ,"0.1"
			DlgText "Parameter4" ,"4"
			DlgText "EpsilonRT" ,"1"
			DlgText "freq_textbox", "5"
    		DlgEnable "Dispersion_on",False
            DlgText("linelength_textbox","1.0000e+00")
			cst_frequency = 5
			cst_Linetype = "Coax"
			' Calculate impedance once initially to have proper dialog values
			cst_LineImpedance = calculate_coax(DlgText("Parameter1"),DlgText("Parameter2"),DlgText("EpsilonRT"))
			DlgText "Lineimpedance" , Format (cst_LineImpedance,"Fixed")
			DlgText "eps_eff_textbox" , Format (cst_eps_eff,"Fixed")
			cst_linelength = Abs(RealVal(CStr(DlgText("linelength_textbox"))))
			cst_phase=360*cst_linelength*Units.GetGeometryUnitToSI*cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)/CLight
			cst_dlg_phase = CStr(cst_phase)
			DlgText "phase_textbox" , Format(cst_phase, "0.0000e+00")

    Case 2 ' Value changing or button pressed
        Select Case DlgItem$

        '-------------------Case for calculation types----------------

        Case "Help"
			StartHelp "common_preloadedmacro_calculate_calculate_analytical_line_impedance"
			DialogFunc = True

        Case "LineTypeSelectionDLB"

        	' Default: all parameter boxes are active, dispersion is "off"
			DlgEnable("Parameter1",True)
			DlgEnable("Parameter2",True)
			DlgEnable("Parameter3",True)
			DlgEnable("Parameter4",True)
			DlgEnable("Dispersion_on", False)
		    cst_Dispersion_on = False

			Select Case DlgText("LineTypeSelectionDLB")
				Case "Coax"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_coax.bmp"
		            DlgEnable "Parameter3",False
		            DlgEnable "Parameter4",False
		            DlgText "Lineimpedance",""
		            DlgText "TextParameter1","d"
					DlgText "TextParameter2","D"
					DlgText "TextParameter3",""
					DlgText "TextParameter4",""
		    		cst_Linetype = "Coax"

				Case "Stripline"
		            cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_strip.bmp"
		            DlgEnable "Parameter3",False
		            DlgEnable "Parameter4",False
		            DlgText "Lineimpedance",""
					DlgText "TextParameter1","h"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3",""
					DlgText "TextParameter4",""
		    		cst_Linetype = "Stripline"

				Case "Thick Stripline"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_thick_stripline.bmp"
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","h"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3","t"
					DlgText "TextParameter4",""
					cst_Linetype = "Thick Stripline"

				Case "Asymmetric Thick Stripline"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_thick_asymmetric_stripline.bmp"
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","h1"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3","t"
					DlgText "TextParameter4","h2"
					cst_Linetype = "Asymmetric Thick Stripline"

				Case "Thin Microstrip"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_mStrip.bmp"
					DlgEnable "Parameter3",False
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","h"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3",""
					DlgText "TextParameter4",""
					DlgEnable "Dispersion_on",True
					cst_Linetype = "Thin Microstrip"

				Case "Thick Microstrip"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_thick_strip.bmp"
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","h"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3","t"
					DlgEnable "Dispersion_on",True
					cst_Linetype = "Thick Microstrip"

				Case "Coplanar Waveguide with Ground"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_cpw_mg.bmp"
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","h"
					DlgText "TextParameter2","W"
					DlgText "TextParameter3","g"
					DlgText "TextParameter4",""
					DlgEnable "Dispersion_on",True
					cst_Linetype = "cpw_mg"

				Case "Coplanar Waveguide"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_cpw.bmp"
					DlgEnable "Parameter3",False
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","W"
					DlgText "TextParameter2","g"
					DlgText "TextParameter3",""
					DlgText "TextParameter4",""
					cst_Linetype = "cpw"

				Case "Thick Coplanar Waveguide"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_thick_cpw.bmp"
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","W"
					DlgText "TextParameter2","g"
					DlgText "TextParameter3","t"
					DlgText "TextParameter4","h"
					DlgEnable "Dispersion_on",True
					cst_Linetype = "thick_cpw"

				Case "Differential Stripline"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_coupled_strip.bmp"
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","W"
					DlgText "TextParameter2","h"
					DlgText "TextParameter3","t"
					DlgText "TextParameter4","s"
					cst_Linetype = "diff_strip"

				Case "Suspended Microstrip"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_suspended_strip.bmp"
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","W"
					DlgText "TextParameter2","a"
					DlgText "TextParameter3","b"
					DlgText "TextParameter4",""
					cst_Linetype = "suspended_strip"

				Case "Inverted Suspended Microstrip"
					cst_filename = macropath+"\Calculate\Calculate analytical Line Impedance_inverse_suspended_strip.bmp"
					DlgEnable "Parameter4",False
					DlgText "Lineimpedance",""
					DlgText "TextParameter1","W"
					DlgText "TextParameter2","a"
					DlgText "TextParameter3","b"
					DlgText "TextParameter4",""
					cst_Linetype = "inverse_suspended_strip"

				Case Else
					MsgBox("Unknown linetype. Please contact support.")

			End Select
			DlgSetPicture "Picture1",cst_filename,0
            DialogFunc% = True 'do not exit the dialog

        '---- Calculate button has been pressed-----------------------------------------
		Case "Calculate"
			DialogFunc% = True 'do not exit the dialog
			cst_LineImpedance = -1 ' default value in case of error
			cst_eps= CDbl(DlgText("EpsilonRT")) ' eps is needed for all types

			Select Case cst_Linetype
				Case "Coax"
					cst_d1= CDbl(DlgText("Parameter1"))
            		cst_d2= CDbl(DlgText("Parameter2"))
					If (cst_eps>=1 And cst_d2>cst_d1 And cst_d1>0) Then
						cst_LineImpedance = calculate_coax(cst_d1,cst_d2,cst_eps)
						DlgText "GroupBox4" , "Impedance static"
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "Stripline", "Thin Stripline"
					cst_h= CDbl(DlgText("Parameter1"))
					cst_w= CDbl(DlgText("Parameter2"))
					If ((cst_eps>=1) And (cst_h>0) And (cst_w>0) )Then
						cst_LineImpedance = calculate_stripline(cst_h,cst_w,cst_eps)
						DlgText "GroupBox4" , "Impedance static"
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "Thick Stripline"
					cst_h= CDbl(DlgText("Parameter1"))
					cst_w= CDbl(DlgText("Parameter2"))
					cst_t= CDbl(DlgText("Parameter3"))
					If ((cst_eps>=1) And (cst_h>0) And (cst_w>0) )Then
						cst_LineImpedance = calculate_thick_stripline(cst_h,cst_w,cst_t,cst_eps)
						DlgText "GroupBox4" , "Impedance static"
					Else
 						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "Asymmetric Thick Stripline"
					cst_h1= CDbl(DlgText("Parameter1"))
					cst_w= CDbl(DlgText("Parameter2"))
					cst_t = CDbl(DlgText("Parameter3"))
					cst_h2= CDbl(DlgText("Parameter4"))
					If ((cst_eps>=1) And (cst_h1>0) And (cst_w>0) And (cst_h2>0) And (cst_t>0)) Then
						cst_LineImpedance = calculate_AsymmetricThickStripline(cst_h1,cst_h2,cst_w,cst_t,cst_eps)
						DlgText "GroupBox4" , "Impedance static"
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "Thin Microstrip", "Thick Microstrip"
					cst_h= CDbl(DlgText("Parameter1"))
					cst_w= CDbl(DlgText("Parameter2"))
					If (cst_Linetype="Thin Microstrip") Then
						cst_t = 0.00001*cst_w
					Else
						cst_t = CDbl(DlgText("Parameter3"))
					End If
					If ((cst_eps>=1) And (cst_h>0) And (cst_w>0) And (cst_t>0)) Then
						cst_LineImpedance = calculate_ThickMicrostripline(cst_h,cst_w,cst_t,cst_eps)
						If cst_Dispersion_on Then
							DlgText "GroupBox4" , "Impedance at f_center"
						Else
							DlgText "GroupBox4" , "Impedance static"
						End If
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "cpw_mg"
					cst_h= CDbl(DlgText("Parameter1"))
					cst_w= CDbl(DlgText("Parameter2"))
					cst_g= CDbl(DlgText("Parameter3"))
					If (cst_eps>=1) Then
						cst_LineImpedance = calculate_cpw_mg(cst_h,cst_w,cst_g,cst_eps)
						If cst_Dispersion_on Then
							DlgText  "GroupBox4", "Impedance at f_center"
						Else
							DlgText  "GroupBox4", "Impedance static"
						End If
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "cpw"
					cst_w= CDbl(DlgText("Parameter1"))
					cst_g= CDbl(DlgText("Parameter2"))
					cst_h= 20*cst_w
					cst_t=0.001*cst_g
					If (cst_eps>=1) Then
						cst_LineImpedance = calculate_cpw(cst_w,cst_g,cst_t,cst_h,cst_eps)
						If cst_Dispersion_on Then
							DlgText  "GroupBox4", "Impedance at f_center"
						Else
							DlgText  "GroupBox4", "Impedance static"
						End If
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "thick_cpw"
					cst_w= CDbl(DlgText("Parameter1"))
					cst_g= CDbl(DlgText("Parameter2"))
					cst_t= CDbl(DlgText("Parameter3"))
					cst_h= CDbl(DlgText("Parameter4"))
					If (cst_t >= 0.1 * cst_g) Then
						MsgBox "Model not valid for t/g > 0.1"
					End If
					If (cst_eps>=1 And cst_t <= 0.1 * cst_g ) Then
						cst_LineImpedance = calculate_cpw(cst_w,cst_g,cst_t,cst_h,cst_eps)
						If cst_Dispersion_on Then
							DlgText  "GroupBox4", "Impedance at f_center"
						Else
							DlgText  "GroupBox4", "Impedance static"
						End If
					Else
						MsgBox "Error in Input, Please check numerical values"
					End If

				Case "diff_strip"
					cst_w= CDbl(DlgText("Parameter1"))
					cst_h= CDbl(DlgText("Parameter2"))
					cst_t= CDbl(DlgText("Parameter3"))
					cst_s= CDbl(DlgText("Parameter4"))
					cst_eps_eff = cst_eps
					cst_LineImpedance = calculate_diff_strip(cst_w,cst_h,cst_t,cst_s,cst_eps)
					DlgText  "GroupBox4", "Impedance static"

				Case "suspended_strip"
					cst_w= CDbl(DlgText("Parameter1"))
					cst_a= CDbl(DlgText("Parameter2"))
					cst_b= CDbl(DlgText("Parameter3"))
					cst_LineImpedance = calculate_suspended_strip(cst_w,cst_a,cst_b,cst_eps)
					DlgText  "GroupBox4", "Impedance static"

				Case "inverse_suspended_strip"
					cst_w= CDbl(DlgText("Parameter1"))
					cst_a= CDbl(DlgText("Parameter2"))
					cst_b= CDbl(DlgText("Parameter3"))
					cst_LineImpedance = calculate_inverse_suspended_strip(cst_w,cst_a,cst_b,cst_eps)
					DlgText  "GroupBox4", "Impedance static"

			End Select
			DlgText "Lineimpedance" , Format (cst_LineImpedance,"Fixed")
			DlgText "eps_eff_textbox" , Format (cst_eps_eff,"Fixed")

			cst_linelength = Abs(RealVal(CStr(DlgText("linelength_textbox"))))
			cst_phase=360*cst_linelength*Units.GetGeometryUnitToSI*cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)/CLight
			cst_dlg_phase = CStr(cst_phase)
			DlgText "phase_textbox" , Format(cst_dlg_phase, "0.0000e+00")
			If (cst_frequency <> 0) Then
				cst_linelength = cst_phase*CLight/(cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)*360)*Units.GetGeometrySIToUnit
				cst_dlg_linelength = CStr(cst_linelength)
				DlgText("linelength_textbox" ,Format(cst_dlg_linelength, "0.0000e+00"))
			Else
				DlgText "linelength_textbox" , "N/A"
			End If

 		Case "ExitPB"
 			Exit All

		Case "Build3DPB"
			Build3DTransmissionLine(cst_Linetype, _
									InputBox("Physical length ["+Units.GetUnit("Length")+"]", "Physical Length", DlgText("linelength_textbox")), _
									DlgText("Parameter1"), _
									DlgText("Parameter2"), _
									DlgText("Parameter3"), _
									DlgText("Parameter4"), _
									DlgText("EpsilonRT"))
			DialogFunc% = False

        Case "Dispersion_on"
            cst_Dispersion_on = SuppValue
	        DialogFunc% = True 'do not exit the dialog

        End Select

	Case 3   ' Text box changed
		Select Case DlgItem
			Case "linelength_textbox"
				cst_linelength = Abs(RealVal(CStr(DlgText("linelength_textbox"))))
				cst_dlg_linelength = CStr(cst_linelength )
				DlgText("linelength_textbox" ,Format(cst_linelength, "0.0000e+00"))
				cst_phase=360*cst_linelength*Units.GetGeometryUnitToSI*cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)/CLight
				cst_dlg_phase = CStr(cst_phase)
				DlgText "phase_textbox" , Format(cst_phase, "0.0000e+00")

	       	 Case "phase_textbox"
				cst_phase = Abs(RealVal(CStr(DlgText("phase_textbox"))))
				cst_dlg_phase = CStr(cst_phase)
				DlgText "phase_textbox" , Format(cst_phase, "0.0000e+00")
				If (cst_frequency <> 0) Then
		        	cst_linelength = cst_phase*CLight/(cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)*360)*Units.GetGeometrySIToUnit
	            	cst_dlg_linelength = CStr(cst_linelength)
					DlgText("linelength_textbox" ,Format(cst_linelength, "0.0000e+00"))
				Else
					DlgText "linelength_textbox" , ""
				End If

	       	 Case "freq_textbox"
				cst_frequency = Abs(RealVal(CStr(DlgText("freq_textbox"))))
				cst_dlg_freq = CStr(cst_frequency)
				DlgText "freq_textbox" ,CStr(cst_dlg_freq)
	    		cst_phase=360*cst_linelength*Units.GetGeometryUnitToSI*cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)/CLight
	    		cst_dlg_phase = CStr(cst_phase)
				DlgText "phase_textbox" , Format(cst_phase, "0.0000e+00")
				If (cst_frequency <> 0) Then
		        	cst_linelength = cst_phase*CLight/(cst_frequency*Units.GetFrequencyUnitToSI*Sqr(cst_eps_eff)*360)*Units.GetGeometrySIToUnit
	            	cst_dlg_linelength = CStr(cst_linelength)
		        	DlgText("linelength_textbox" ,Format(cst_linelength, "0.0000e+00"))
	           	Else
					DlgText "linelength_textbox" , ""
				End If

	        End Select

    End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------------
Function RealVal(lib_Text As Variant) As Double
	If Len(lib_Text) > 0 Then
        If (CDbl("0.5") > 1) Then
                RealVal = CDbl(Replace(CStr(lib_Text), ".", ","))
        Else
                RealVal = CDbl(CStr(lib_Text))
        End If
	End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------
Function calculate_coax(cst_d1f As Double, cst_d2f As Double, cst_epsf As Double) As Double
      'Formula: Standard Formula
      cst_eps_eff =  cst_epsf
      calculate_coax = 376.734*Sqr(1/cst_epsf)*(Log(cst_d2f/cst_d1f)/(2*3.1415))
End Function

'-----------------------------------------------------------------------------------------------------------------------------
Function calculate_stripline(cst_hf As Double, cst_wf As Double, cst_epsf As Double) As Double 'Formula Microwave and Engineering (Pozar) pp.156

	Dim cst_corr As Double, cst_We As Double
	Dim cst_k As Double, cst_k_prime As Double

	'Formula: Pozar, Microwave Engineering, 2nd Edition. pp. 156
	' --- Calculate corr factor
	'If ((cst_wf/cst_hf)< (35/100)) Then
	' cst_corr= (35/100-cst_wf/cst_hf)^2
	'Else
	' cst_corr = 0
	' End If
	'cst_We=cst_hf*(cst_wf/cst_hf-cst_corr)
	cst_eps_eff =  cst_epsf
	cst_k = sech(Pi*cst_wf/(2*cst_hf))
	cst_k_prime= tanh(Pi * cst_wf/(2 * cst_hf))
	calculate_stripline= 30*Pi/Sqr(cst_epsf )*K_int(cst_k)/K_int(cst_k_prime)

	'calculate_stripline = 30*Pi/Sqr(cst_epsf)*(cst_hf/(cst_We+0.441*cst_hf))

End Function

Function calculate_thick_stripline(cst_hf As Double, cst_wf As Double,cst_tf As Double, cst_epsf As Double) As Double

	Dim cst_corr As Double
	Dim cst_We As Double
	Dim cfprime As Double
	Dim cst_p1 As Double
	Dim cst_p2 As Double

	'Zinke Brunswig pp 161/ for w/(h-t) > 0.35Cohn, Transmission Lines
	' --- Calculate corr factor
	cst_eps_eff =  cst_epsf
	If cst_tf/cst_hf <= 0.25 Then
		If ((cst_wf/(cst_hf-cst_tf))< (35/100)) Then
	  		calculate_thick_stripline =60/Sqr(cst_epsf)*Log((8*cst_hf)/(Pi*cst_wf)/(1+cst_tf/(pi*cst_wf)*(1+Log(4*pi*cst_wf/cst_tf))+0.51*(cst_tf/cst_wf)^2))
	 	Else
	 		'calculate_thick_stripline = 94.25*(1-cst_t/cst_hf)/Sqr(cst_epsf)/(cst_wf/cst_hf+2/pi*Log((2-cst_t/cst_hf)/(1-cst_t/cst_hf))-cst_t/(pi*cst_hf)*Log((cst_t*(2-cst_t/cst_hf))/(cst_hf*(1-cst_t/cst_hf)^2)))
	  		cst_p1 = 2/(1-cst_tf/cst_hf)*Log(1/(1-cst_tf/cst_hf)+1)
	  		cst_p2 = (1/(1-cst_tf/cst_hf)-1)*Log(1/(1-cst_tf/cst_hf)^2 - 1)
	  		cfprime =0.0885*cst_epsf/pi*(cst_p1-cst_p2)
	  		calculate_thick_stripline = 94.15/(Sqr(cst_epsf)*(cst_wf/cst_hf/(1-cst_tf/cst_hf)+cfprime/(0.0885*cst_epsf)))
	 	End If
	Else
		calculate_thick_stripline = 0
		MsgBox "t/h too large. Analytical equations do not cover this parameter range."
	End If

End Function

Function calculate_ThickMicrostripline(cst_hf As Double, cst_wf As Double, cst_tf As Double, cst_epsf As Double) As Double 'Formula Microwave and Engineering (Pozar) pp.162
      Dim cst_eps_effl As Double, cst_Wf_eff As Double, cst_Fcorr As Double, cst_C_corr As Double, cst_Z0 As Double, cst_Z0_ff As Double, cst_ff As Double, cst_f_50 As Double,cst_f_k_TM0 As Double, cst_eps_eff_ff As Double
      Dim cst_m_c As Double, cst_m_0 As Double, cst_m As Double, cst_hf_si As Double, cst_wf_si As Double,cst_ff_run As Double
      Dim cst_sqr_wf_hf As Double, cst_ff_step As Double

	'Formulas: Collin, Foundation of Microwave Engineering, pp. 150

	' --- Calculate eps_eff
	cst_C_corr = (cst_epsf-1)/4.6*(cst_tf/cst_hf)/Sqr(cst_wf/cst_hf)

	'Formula: Pozar, Microwave Engineering, 2nd Edition. pp. 162
	If (cst_wf/cst_hf<1) Then
		cst_Fcorr=1/Sqr(1+(12*cst_hf/cst_wf))+0.04*(1-cst_wf/cst_hf)^2
	Else
		cst_Fcorr=1/Sqr(1+(12*cst_hf/cst_wf))
	End If
	cst_eps_effl =((cst_epsf+1)/2)+((cst_epsf-1)/2)*cst_Fcorr
	cst_eps_eff =  cst_eps_effl

	'Formulas: Collin, Foundation of Microwave Engineering, pp. 150
	If (cst_wf/cst_hf<1/(2*Pi)) Then
		cst_Wf_eff= cst_wf + 0.398*cst_tf*(1+Log(4*Pi*cst_wf/cst_tf))
	Else
		cst_Wf_eff= cst_wf + 0.398*cst_tf*(1+Log(2*cst_hf/cst_tf))
	End If

	'Formula: Pozar, Microwave Engineering, 2nd Edition. pp. 162
	If (cst_wf/cst_hf < 1 ) Then
       	cst_Z0=(60/Sqr(cst_eps_effl))*Log((8*cst_hf/cst_Wf_eff)+(cst_Wf_eff/(4*cst_hf)))
	Else
       	cst_Z0=(120*Pi)/(Sqr(cst_eps_effl)*((cst_Wf_eff/cst_hf)+1.393+0.667*Log((cst_Wf_eff/cst_hf)+1.444)))
	End If

	If (cst_Dispersion_on) Then    ' now it is going really nasty !!! from Microstriplines and Slotlines... K. C. Gupta u.a. (pp. 29)
		cst_impedancefile = resultdir + "\" + "eps=" + CStr(cst_epsf)+ "_imp.sig"
		Open cst_impedancefile For Output As #1
		If cst_epsf = 1 Then
			cst_epsf = 1.01 ' to avoid problems in case of no Substrat
        End If
		cst_ff = Units.GetFrequencyUnitToSI*(dFmax-dFmin)/2   ' get center freq.
        cst_ff_run = Units.GetFrequencyUnitToSI*dFmin
	    cst_ff_step=Units.GetFrequencyUnitToSI*(dFmax-dFmin)/101
	    cst_hf_si= cst_hf*Units.GetGeometryUnitToSI
	    cst_wf_si= cst_wf*Units.GetGeometryUnitToSI
	    cst_sqr_wf_hf = Sqr(cst_wf/cst_hf)
	    cst_f_k_TM0 = CLight*Atn(cst_epsf*Sqr((cst_eps_eff-1)/(cst_epsf-cst_eps_eff)))/(2*Pi*cst_hf_si*Sqr(cst_epsf-cst_eps_eff))  ' cutoff for higher modes
		'calculate correction_factors
		cst_m_0= 1+(1/(1+cst_sqr_wf_hf))+0.32*(1/(1+cst_sqr_wf_hf))^3
		cst_f_50=cst_f_k_TM0/(0.75+(0.75-(0.332/cst_epsf^1.73))*cst_wf_si/cst_hf_si)
		cst_m = cst_m_0*cst_m_c
		For cst_lauf_index = 0 To 100
			cst_ff_run = cst_ff_run+cst_ff_step
			If cst_wf/cst_hf<0.7 Then
				cst_m_c= 1+ (1.4/(1+cst_wf/cst_hf))*(0.15-0.235*Exp(-0.45*cst_ff_run/cst_f_k_TM0))
			Else
				cst_m_c= 1
			End If
			cst_m = cst_m_0*cst_m_c
			cst_eps_eff_ff = cst_epsf-((cst_epsf-cst_eps_effl)/(1+(cst_ff_run/cst_f_50)^cst_m))
			cst_Z0_ff = cst_Z0*((cst_eps_eff_ff-1)/(cst_eps_effl-1))*Sqr(cst_eps_effl/cst_eps_eff_ff)
			Print #1, cst_ff_run*Units.GetFrequencySIToUnit;"		";cst_Z0_ff
		Next cst_lauf_index
		Close #1

		Dim myResult1D As Object
		Set myResult1D = Result1D(cst_impedancefile)
		'myResult1D.Type "XYSignal"
		myResult1D.Xlabel "freq"
		myResult1D.Title "Impedance"
		If(Left(GetApplicationName,2) = "DS") Then
			myResult1D.AddToTree("Design\Results\Impedance\Impedance eps = " + CStr(cst_epsf))
			SelectTreeItem ("Design\Results\Impedance\Impedance eps = " + CStr(cst_epsf))
		Else
			myResult1D.AddToTree("1D Results\Impedance\Impedance eps = " + CStr(cst_epsf))
			SelectTreeItem ("1D Results\Impedance\Impedance eps = " + CStr(cst_epsf))
		End If

		' and once agin for center freq
		If cst_wf/cst_hf<0.7 Then
			cst_m_c= 1+ (1.4/(1+cst_wf/cst_hf))*(0.15-0.235*Exp(-0.45*cst_ff/cst_f_k_TM0))
		Else
			cst_m_c= 1
		End If
		cst_m = cst_m_0*cst_m_c
		cst_eps_eff_ff = cst_epsf-((cst_epsf-cst_eps_eff)/(1+(cst_ff/cst_f_50)^cst_m))
		cst_Z0_ff = cst_Z0*((cst_eps_eff_ff-1)/(cst_eps_eff-1))*Sqr(cst_eps_eff/cst_eps_eff_ff)
		calculate_ThickMicrostripline = cst_Z0_ff
	Else
		calculate_ThickMicrostripline = cst_Z0
	End If

End Function

Function calculate_AsymmetricThickStripline(cst_h1f As Double, cst_h2f As Double, cst_wf As Double, cst_tf As Double, cst_epsf As Double) As Double

	cst_eps_eff = cst_epsf
	calculate_AsymmetricThickStripline = -1
	If (((cst_wf/cst_h1f)<0.1) Or ((cst_wf/cst_h1f)<0.1)) Then
		MsgBox("Cannot calculate impedance. Formula only valid for 0.1<=w/h<=2.")
	ElseIf (cst_tf/cst_h1f >= 0.25) Then
		MsgBox("Cannot calculate impedance. Formula only valid for t/h<0.25")
	Else
		' Formulas from Seymour Cohn, "Problems in Strip Transmission Lines", Mtt-3, No. 2, March 1955, pp. 199-126
		calculate_AsymmetricThickStripline = (2*calculate_thick_stripline(2*cst_h1f+cst_tf, cst_wf, cst_tf, cst_epsf)*calculate_thick_stripline(2*cst_h2f+cst_tf, cst_wf, cst_tf, cst_epsf))/(calculate_thick_stripline(2*cst_h1f+cst_tf, cst_wf, cst_tf, cst_epsf)+calculate_thick_stripline(2*cst_h2f+cst_tf, cst_wf, cst_tf, cst_epsf))
	End If

End Function

' Deprecated, always use "thick microstrip"
'Function calculate_ThinMicrostripline(cst_hf As Double, cst_wf As Double, cst_epsf As Double) As Double 'Formula Microwave and Engineering (Pozar) pp.162
'        Dim cst_eps_effl As Double, cst_We As Double, cst_Fcorr As Double
'         ' --- Calculate eps_eff
'         If (cst_wf/cst_hf<1) Then
'         cst_Fcorr=1/Sqr(1+(12*cst_hf/cst_wf))+0.04*(1-cst_wf/cst_hf)^2
'         Else
'         cst_Fcorr=1/Sqr(1+(12*cst_hf/cst_wf))
'         End If
'         cst_eps_effl=((cst_epsf+1)/2) + ((cst_epsf-1)/2)*cst_Fcorr
'
'        cst_eps_eff =  cst_eps_effl
'       '-------------------
'
'       If (cst_wf/cst_hf < 1 ) Then
'         calculate_ThinMicrostripline=(60/Sqr(cst_eps_eff))*Log((8*cst_hf/cst_wf)+(cst_wf/(4*cst_hf)))
'         Else
'         calculate_ThinMicrostripline=(120*Pi)/(Sqr(cst_eps_eff)*((cst_wf/cst_hf)+1.393+0.667*Log((cst_wf/cst_hf)+1.444)))
'         End If
'End Function


Function calculate_cpw_mg(cst_hf As Double, cst_wf As Double, cst_gf As Double, cst_epsf As Double) As Double 'Formula Microwave and Engineering (Pozar) pp.162

	Dim cst_epsf_eff As Double, cst_k As Double, cst_k_prime As Double, cst_k1 As Double, cst_k1_prime As Double, cst_a As Double, cst_b As Double, cst_k_product As Double

	cst_a=cst_wf
	cst_b=cst_wf+2*cst_gf
    cst_k=cst_a/cst_b
    cst_k_prime=Sqr(1-cst_k^2)
    'a_t=a+1.25*tf/Pi*(1+Log(4*Pi*a/t))
    'b_t=b-1.25*tf/Pi*(1+Log(4*Pi*a/t))
    'k_t= a_t/b_t
    'k_t_prime=Sqr(1-k_t^2)
    'k1=tanh((Pi*a_t)/(4*hf))/tanh((Pi*b_t)/(4*hf))
    cst_k1=tanh((Pi*cst_a)/(4*cst_hf))/tanh((Pi*cst_b)/(4*cst_hf))
    cst_k1_prime = Sqr(1-cst_k1^2)
    cst_k_product=((K_int(cst_k_prime)*K_int(cst_k1))/(K_int(cst_k)*K_int(cst_k1_prime)))
    cst_epsf_eff= (1+cst_epsf*cst_k_product)/(1+cst_k_product)
    cst_eps_eff =  cst_epsf_eff

	'espf_eff_t=espf_eff-((espf_eff-1)/((b-a)*K_int(k)/(0.7*tf*2*K_int(k_prime))+1))
	calculate_cpw_mg=60*Pi/Sqr(cst_epsf_eff)*(1/((K_int(cst_k)/K_int(cst_k_prime))+(K_int(cst_k1)/K_int(cst_k1_prime))))

End Function

'------------------------ This function is used to calculate thin and thick CPW with finit and infinit thick substrate

Function calculate_cpw(cst_wf As Double, cst_gf As Double, cst_tf As Double , cst_hf As Double, cst_epsf As Double) As Double '
      Dim cst_epsf_eff_t As Double
      Dim cst_epsf_eff As Double
      Dim cst_eps_eff_ff As Double
      Dim cst_k As Double
      Dim cst_k_prime As Double
      Dim cst_k_t As Double
      Dim cst_k_t_prime As Double
      Dim cst_k1 As Double
      Dim cst_k1_prime As Double
      Dim cst_a As Double
      Dim cst_a_t As Double
      Dim cst_b As Double
      Dim cst_b_t As Double
      Dim cst_u As Double, cst_p As Double,  cst_v As Double, cst_gf_si As Double, cst_hf_si As Double
      Dim cst_wf_si As Double, cst_ff_run As Double
      Dim cst_ff As Double, cst_ff_step As Double
      Dim cst_f_te As Double, cst_Z0_ff As Double, cst_Z0 As Double


      cst_a=cst_wf
      cst_b=cst_wf+2*cst_gf
      cst_k=cst_a/cst_b
      cst_k_prime=Sqr(1-cst_k^2)
      cst_a_t=cst_a+1.25*cst_tf/Pi*(1+Log(4*Pi*cst_a/cst_tf))
      cst_b_t=cst_b-1.25*cst_tf/Pi*(1+Log(4*Pi*cst_a/cst_tf))
      cst_k_t= cst_a_t/cst_b_t
      cst_k_t_prime=Sqr(1-cst_k_t^2)
      cst_k1=sinh((Pi*cst_a_t)/(4*cst_hf))/sinh((Pi*cst_b_t)/(4*cst_hf))
      cst_k1_prime = Sqr(1-cst_k1^2)
      cst_epsf_eff= 1+ _
           ((cst_epsf-1)/2) _
            *((K_int(cst_k_prime)*K_int(cst_k1))/(K_int(cst_k)*K_int(cst_k1_prime)))

      cst_epsf_eff_t=cst_epsf_eff-((cst_epsf_eff-1)/((cst_b-cst_a)*K_int(cst_k)/(0.7*cst_tf*2*K_int(cst_k_prime))+1))
      cst_eps_eff = cst_epsf_eff
If (cst_Dispersion_on) Then    ' now it is going really nasty !!! from Microstriplines and Slotlines... K. C. Gupta u.a. (pp. 29)
          cst_impedancefile = resultdir + "\" + "eps=" + CStr(cst_epsf)+ "_imp.sig"
          Open cst_impedancefile For Output As #1
          cst_ff = Units.GetFrequencyUnitToSI*(dFmax-dFmin)/2   ' get center freq.
          cst_ff_run = Units.GetFrequencyUnitToSI*dFmin
          cst_ff_step=Units.GetFrequencyUnitToSI*(dFmax-dFmin)/101
          cst_hf_si= cst_hf*Units.GetGeometryUnitToSI
          cst_wf_si= cst_wf*Units.GetGeometryUnitToSI
          cst_gf_si= cst_gf*Units.GetGeometryUnitToSI
          cst_f_te= CLight/(4*cst_hf_si*Sqr(cst_epsf-1))
          cst_p = Log(2*cst_gf_si/cst_hf_si)
          cst_v = 0.43-0.86*cst_p+0.54*cst_p^2
          cst_u =0.54 -0.64*cst_p+0.015*cst_p^2

          cst_eps_eff_ff=1
          For cst_lauf_index = 0 To 100
         cst_ff_run = cst_ff_run+cst_ff_step
	     cst_Z0_ff = cst_Z0*((cst_eps_eff_ff-1)/(cst_epsf_eff-1))*Sqr(cst_epsf_eff/cst_eps_eff_ff)
	     Print #1, cst_ff_run*Units.GetFrequencySIToUnit;"		";cst_Z0_ff

	     Next cst_lauf_index

          Close #1
		Dim myResult1D As Object
		Set myResult1D = Result1D(cst_impedancefile)
		'myResult1D.Type "XYSignal"
		myResult1D.Xlabel "freq"
		myResult1D.Title "Impedance"
		If(Left(GetApplicationName,2) = "DS") Then
			myResult1D.AddToTree("Design\Results\Impedance\Impedance eps = " + CStr(cst_epsf))
    	    SelectTreeItem ("Design\Results\Impedance\Impedance eps = " + CStr(cst_epsf))
		Else
			myResult1D.AddToTree("1D Results\Impedance\Impedance eps = " + CStr(cst_epsf))
	        SelectTreeItem ("1D Results\Impedance\Impedance eps = " + CStr(cst_epsf))
		End If
  ' and once agin for center freq


         End If


      calculate_cpw=30*Pi/Sqr(cst_epsf_eff_t)*K_int(cst_k_t_prime)/K_int(cst_k_t)

      End Function

' calculate diff Strip Impedance
Function calculate_diff_strip(cst_wf As Double, cst_hf As Double, cst_tf As Double , cst_sf As Double, cst_epsf As Double)
      Dim cst_ko As Double
      Dim cst_ko_prime As Double
      Dim cst_ke As Double
      Dim cst_ke_prime As Double

      Dim cst_tw As Double
      Dim cst_cf_tw As Double
	  Dim cst_cf As Double
	  Dim cst_Zo As Double  'zero tickness odd mode impedace coupled line
      Dim cst_Ze As Double  'zero tickness even mode impedace coupled line

	  Dim cst_Zof As Double  'finite tickness isolated impedacen
	  Dim cst_Zoo As Double   ' zero thciness isolated impedance
      'help
      Dim cst_m As Double
      Dim dw1 As Double, dw2 As Double, Dw3 As Double, dw As Double
      Dim cst_wp As Double
      Dim z01 As Double, z02 As Double, z03 As Double, z04 As Double

 cst_tw = cst_tf/cst_hf

cst_eps_eff = cst_epsf
' stray capp
 Dim cst_aaa As Double
 cst_aaa =  1/(1-cst_tw)

 cst_cf_tw = 0.0885*cst_epsf/Pi*(2*cst_aaa * Log(cst_aaa + 1) - (cst_aaa - 1) * Log(cst_aaa^2 - 1))
 cst_cf = 0.0885*cst_epsf/Pi*2*Log(2)

'finite tickness isolated impedance from www.ideaconsulting.com/dstrip.htm

'cst_m = 6 / (3 + (2 * cst_t / (cst_hf - cst_t)))
'dw1 = 1 / (2 * (cst_hf - cst_t)/cst_t + 1)
'dw2 = (1 / (4 * Pi)) /(cst_wf/cst_t + 1.1)
'Dw3 = dw1^2 + dw2^cst_m
'dw = (cst_t/Pi) * ( 1 - 0.5 * Log(Dw3))
'cst_wp = cst_w + dw
'z01 = 30 / Sqr(cst_epsf)
'z02 = (4*(cst_hf-cst_t)) /(Pi * cst_wp)
'z03 = z02 * 2
'z04 = Sqr(z03^2+ 6.27)
'cst_Zof =  z01 * Log( 1 + z02 * (z03 + z04))

'      Thick isolated stripline
cst_Zof = calculate_thick_stripline(cst_hf, cst_wf,cst_tf, cst_epsf)
'      Thin isolated Stripline
cst_Zoo = calculate_stripline(cst_hf, cst_wf, cst_epsf)

'     calculate zero-thickness odd impedance of edge-coupled stripline
 cst_ko = tanh((Pi*cst_wf)/(2*cst_hf))*coth((Pi*(cst_wf+cst_sf))/(2*cst_hf))
 cst_ko_prime=Sqr(1-cst_ko^2)
 cst_Zo = 30*Pi/Sqr(cst_epsf )*K_int(cst_ko_prime)/K_int(cst_ko) 'Im missing a factor of two somewere. Just added it here
'   calculate zero-thickness even impedance of edge-coupled stripline
 cst_ke = tanh((Pi*cst_wf)/(2*cst_hf))*tanh((Pi*(cst_wf+cst_sf))/(2*cst_hf))
 cst_ke_prime=Sqr(1-cst_ke^2)
 cst_Ze = 30*Pi/Sqr(cst_epsf )*K_int(cst_ko_prime)/K_int(cst_ko)


     If (cst_sf/cst_tf > 5) Then
     Dim cst_A1 As Double, cst_A2 As Double, CST_A3_even As Double
     Dim cst_A3_odd As Double, cst_A4_even As Double, cst_a4_odd As Double


     cst_A1 = 1/cst_Zof
	 cst_A2 = cst_cf_tw /cst_cf
	 CST_A3_even = 1 /cst_Zoo
	 cst_A3_odd = 1 /cst_Zo
	 cst_A4_even = 1 /cst_Ze
	 cst_a4_odd = 1 / cst_Zoo
	 calculate_diff_strip = 2 / (cst_A1 + cst_A2 * (cst_A3_odd - cst_a4_odd))


      Else
      Dim cst_p1 As Double
      Dim cst_p2 As Double
      Dim cst_p3 As Double

      cst_p1 = 1/cst_Zo
      cst_p2 = 1/cst_Zof-1/cst_Zoo
      cst_p3 = 2/377*(cst_cf_tw/cst_epsf-cst_cf/cst_epsf)
      calculate_diff_strip = 2/(cst_p1+cst_p2-cst_p3+2*cst_t/(377*cst_sf))

      End If

      '30*Pi/Sqr(cst_epsf )*K_int(cst_ko_prime)/K_int(cst_ko)*2
      'Im missing a factor of two somewere. Just added it here)
      End Function
' calculate suspended Mstrip

Function calculate_suspended_strip(cst_wss As Double, cst_ass As Double, cst_bss As Double, cst_epss As Double) As Double
      'Tomar et.al.,Formula from "New Quasi-Static Models for the COmputer aided Design of suspended Microstrips
      'IEEE Transactions on Mocrowave Theory and Techniques, vol 35, No. 4
      Dim cst_f As Double
      Dim cst_d00 As Double, cst_d01 As Double, cst_d02 As Double, cst_d03 As Double
	  Dim cst_d10 As Double, cst_d11 As Double, cst_d12 As Double, cst_d13 As Double
      Dim cst_d20 As Double, cst_d21 As Double, cst_d22 As Double, cst_d23 As Double
      Dim cst_d30 As Double, cst_d31 As Double, cst_d32 As Double, cst_d33 As Double
      Dim cst_c0 As Double, cst_c1 As Double, cst_c2 As Double, cst_c3 As Double
      Dim cst_fu As Double, cst_u As Double, cst_Z0 As Double
      Dim cst_f1 As Double, cst_f2 As Double, cst_sqrteps As Double
      Dim cst_en As Double

      cst_en = 2.71829

    cst_f = Log(cst_epss)
	cst_d00 = (176.2576-43.1240*cst_f+13.4094*cst_f^2-1.7010*cst_f^3)*1e-2
	cst_d01 = (4665.2320-1790.4*cst_f+291.5858*cst_f^2-8.0888*cst_f^3)*1e-4
	cst_d02 = (-3025.5070-141.9368*cst_f-3099.47*cst_f^2+777.6151*cst_f^3)*1e-6
	cst_d03 = (2491.569+143.3860*cst_f+10095.55*cst_f^2-2599.132*cst_f^3)*1e-8
	cst_d10 = (-1410.2050+149.9293*cst_f+198.2892*cst_f^2-32.1679*cst_f^3)*1e-4
	cst_d11 = (2548.791+1531.9310*cst_f-1027.5200*cst_f^2+138.4192*cst_f^3)*1e-4
	cst_d12 = (999.3135-4036.7910*cst_f+1762.4120*cst_f^2-298.0241*cst_f^3)*1e-6
	cst_d13 = (-1983.7890+8523.9290*cst_f-5235.4600*cst_f^2+1145.7880*cst_f^3)*1e-8
	cst_d20 = (1954.072+333.3873*cst_f-700.7473*cst_f^2+121.3212*cst_f^3)*1e-5
	cst_d21 = (-3931.09-1890.719*cst_f+1912.266*cst_f^2-319.6794*cst_f^3)*1e-5
	cst_d22 = (-532.1326+7274.7210*cst_f-4955.738*cst_f^2+941.4134*cst_f^3)*1e-7
	cst_d23 = (138.2037-1412.427*cst_f+1184.27*cst_f^2-270.0047*cst_f^3)*1e-8
	cst_d30 = (-983.4028-255.1229*cst_f+455.8729*cst_f^2-83.9468*cst_f^3)*1e-6
	cst_d31 = (1956.3170+779.9975*cst_f-995.9494*cst_f^2+183.1957*cst_f^3)*1e-6
	cst_d32 = (62.855-3462.5*cst_f+2909.923*cst_f^2-614.7068*cst_f^3)*1e-8
	cst_d33 = (-35.2531+601.0291*cst_f-643.0814*cst_f^2+161.2689*cst_f^3)*1e-9
	cst_f1 = 1-Sqr(1/cst_epss)
	cst_c0 = cst_d00+cst_d01*(cst_bss/cst_ass)+cst_d02*(cst_bss/cst_ass)^2+cst_d03*(cst_bss/cst_ass)^3
	cst_c1 = cst_d10+cst_d11*(cst_bss/cst_ass)+cst_d12*(cst_bss/cst_ass)^2+cst_d13*(cst_bss/cst_ass)^3
	cst_c2 = cst_d20+cst_d21*(cst_bss/cst_ass)+cst_d22*(cst_bss/cst_ass)^2+cst_d23*(cst_bss/cst_ass)^3
	cst_c3 = cst_d30+cst_d31*(cst_bss/cst_ass)+cst_d32*(cst_bss/cst_ass)^2+cst_d33*(cst_bss/cst_ass)^3
	cst_f2 = 1/(cst_c0+cst_c1*(cst_wss/cst_bss)+cst_c2*(cst_wss/cst_bss)^2+cst_c3*(cst_wss/cst_bss)^3)
	cst_sqrteps = 1/(1-cst_f1*cst_f2)
	cst_u = (cst_wss/cst_bss)/(1+cst_ass/cst_bss)
	cst_fu =6+(2*Pi-6)*cst_en^-((30.666/cst_u)^0.7528)
    cst_Z0= 60*Log(cst_fu/cst_u+Sqr(1+4/cst_u^2))
    cst_eps_eff =  cst_sqrteps^2
    calculate_suspended_strip = cst_Z0/cst_sqrteps

End Function

' calculate inverse suspended Mstrip

Function calculate_inverse_suspended_strip(cst_wss As Double, cst_ass As Double, cst_bss As Double, cst_epss As Double) As Double
      'Tomar et.al.,Formula from "New Quasi-Static Models for the COmputer aided Design of suspended Microstrips
      'IEEE Transactions on Mocrowave Theory and Techniques, vol 35, No. 4
      Dim cst_f As Double
      Dim cst_d00 As Double, cst_d01 As Double, cst_d02 As Double, cst_d03 As Double
	  Dim cst_d10 As Double, cst_d11 As Double, cst_d12 As Double, cst_d13 As Double
      Dim cst_d20 As Double, cst_d21 As Double, cst_d22 As Double, cst_d23 As Double
      Dim cst_d30 As Double, cst_d31 As Double, cst_d32 As Double, cst_d33 As Double
      Dim cst_c0 As Double, cst_c1 As Double, cst_c2 As Double, cst_c3 As Double
      Dim cst_fu As Double, cst_u As Double, cst_Z0 As Double
      Dim cst_f1 As Double, cst_f2 As Double, cst_sqrteps As Double
      Dim cst_en As Double

      cst_en = 2.71829

    cst_f = Log(cst_epss)
	cst_d00 = (2359.4010-97.1644*cst_f-5.7706*cst_f^2+11.4112*cst_f^3)*1e-3
	cst_d01 = (4855.9472-3408.5207*cst_f+15296.73*cst_f^2-2418.1785*cst_f^3)*1e-5
	cst_d02 = (1763.34+961.0481*cst_f-2089.28*cst_f^2+375.8805*cst_f^3)*1e-5
	cst_d03 = (-556.0909-268.6165*cst_f+623.7094*cst_f^2-119.1402*cst_f^3)*1e-6
	cst_d10 = (219.0660-253.0864*cst_f+208.7469*cst_f^2-27.3285*cst_f^3)*1e-3
	cst_d11 = (915.5589+338.4033*cst_f-253.2933*cst_f^2+40.4745*cst_f^3)*1e-3
	cst_d12 = (-1957.3790-1170.9360*cst_f+1480.8570*cst_f^2-347.6403*cst_f^3)*1e-5
	cst_d13 = (486.7425+279.8323*cst_f-431.3625*cst_f^2+108.824*cst_f^3)*1e-6
	cst_d20 = (5602.7670+4403.3560*cst_f-4517.034*cst_f^2+743.2717*cst_f^3)*1e-5
	cst_d21 = (-2823.481-1562.782*cst_f+3646.15*cst_f^2-823.4223*cst_f^3)*1e-5
	cst_d22 = (253.893+158.5529*cst_f-3235.485*cst_f^2+919.3661*cst_f^3)*1e-6
	cst_d23 = (-147.0235+62.4343*cst_f+887.5211*cst_f^2-270.7555*cst_f^3)*1e-7
	cst_d30 = (-3170.21-1931.852*cst_f+2715.327*cst_f^2-519.342*cst_f^3)*1e-6
	cst_d31 = (596.3251+188.1409*cst_f-1741.477*cst_f^2+465.6756*cst_f^3)*1e-6
	cst_d32 = (124.9655+577.5381*cst_f+1366.453*cst_f^2-481.13*cst_f^3)*1e-7
	cst_d33 = (-530.2099-2666.3520*cst_f-3220.0960*cst_f^2+1324.499*cst_f^3)*1e-9

	cst_f1 = Sqr(cst_epss)-1
	cst_c0 = cst_d00+cst_d01*(cst_bss/cst_ass)+cst_d02*(cst_bss/cst_ass)^2+cst_d03*(cst_bss/cst_ass)^3
	cst_c1 = cst_d10+cst_d11*(cst_bss/cst_ass)+cst_d12*(cst_bss/cst_ass)^2+cst_d13*(cst_bss/cst_ass)^3
	cst_c2 = cst_d20+cst_d21*(cst_bss/cst_ass)+cst_d22*(cst_bss/cst_ass)^2+cst_d23*(cst_bss/cst_ass)^3
	cst_c3 = cst_d30+cst_d31*(cst_bss/cst_ass)+cst_d32*(cst_bss/cst_ass)^2+cst_d33*(cst_bss/cst_ass)^3
	cst_f2 = 1/(cst_c0+cst_c1*(cst_wss/cst_bss)+cst_c2*(cst_wss/cst_bss)^2+cst_c3*(cst_wss/cst_bss)^3)
	cst_sqrteps = 1+cst_f1*cst_f2
	cst_u = cst_wss/cst_bss
	cst_fu =6+(2*Pi-6)*cst_en^-((30.666/cst_u)^0.7528)
    cst_Z0= 60*Log(cst_fu/cst_u+Sqr(1+4/cst_u^2))
    cst_eps_eff =  cst_sqrteps^2
    calculate_inverse_suspended_strip = cst_Z0/cst_sqrteps

End Function

' ------------------------------------------------------------------------

Function K_int(cst_k As Double) 'complet elliptical inegral (iteration) from Transmission Line design handbock from Brian C. Wadell
	 Dim cst_a As Double, cst_b As Double, cst_c As Double, cst_e As Double, cst_f As Double, cst_i As Integer
	cst_a = 1
	cst_b = Sqr(1-cst_k^2)
	cst_c = cst_k
	While cst_c>0.000000000001 ' Stop when  c= 0
	cst_e=cst_a
	cst_f=cst_b
	cst_a=(cst_e+cst_f)/2
	cst_b=Sqr(cst_e*cst_f)
	cst_c=(cst_a-cst_b)/2
	Wend
   K_int = Pi/(2*cst_a)
End Function
' ------------------------------------------------------------------------

Function sinh(cst_a As Double) As Double
sinh=1/2*(Exp(cst_a)-Exp(-cst_a))
End Function

Function tanh(cst_a As Double) As Double
tanh=(Exp(cst_a)-Exp(-cst_a))/(Exp(cst_a)+Exp(-cst_a))
End Function
Function coth(cst_a As Double) As Double
coth=(Exp(cst_a)+Exp(-cst_a))/(Exp(cst_a)-Exp(-cst_a))
End Function
Function sech(cst_a As Double) As Double
sech=2/(Exp(cst_a)+Exp(-cst_a))
End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function Build3DTransmissionLine(sLineType As String, sPhysicalLength As String, sParameter1 As String, sParameter2 As String, sParameter3 As String, sParameter4 As String, sEpsilonRT As String)

	Dim sCommand As String

	StoreParameter("Epsilon_r", sEpsilonRT)
	StoreParameter("Length", sPhysicalLength)

	sCommand = ""
	sCommand = sCommand + "With Material" + vbNewLine
    sCommand = sCommand + ".Reset" + vbNewLine
    sCommand = sCommand + ".Name " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
    sCommand = sCommand + ".Folder " + Chr(34) + Chr(34) + vbNewLine
    sCommand = sCommand + ".Type " + Chr(34) + "Normal" + Chr(34) + vbNewLine
    sCommand = sCommand + ".Epsilon " + Chr(34 )+ "Epsilon_r" + Chr(34) + vbNewLine
    sCommand = sCommand + ".Create" + vbNewLine
    sCommand = sCommand + "End With" + vbNewLine
	AddToHistory("define material: Dielectric", sCommand)

	sCommand = ""
    sCommand = sCommand + "Component.New " + Chr(34) + sLineType + Chr(34) + vbNewLine
	AddToHistory("new component: "+sLineType, sCommand)

	Select Case sLineType
		Case "Coax"
        	StoreParameter("InnerDiameter", sParameter1)
        	StoreParameter("OuterDiameter", sParameter2)
        	StoreParameter("OuterConductorThickness", 0.1*Evaluate(sParameter2))

			sCommand = ""
			sCommand = sCommand + "With Cylinder" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "InnerConductor" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".OuterRadius " + Chr(34) + "InnerDiameter/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".InnerRadius " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Axis " + Chr(34) + "z" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xcenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Ycenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Segments " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define cylinder: "+sLineType+":InnerConductor", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Cylinder" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".OuterRadius " + Chr(34) + "OuterDiameter/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".InnerRadius " + Chr(34) + "InnerDiameter/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Axis " + Chr(34) + "z" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xcenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Ycenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Segments " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define cylinder: "+sLineType+":Dielectric", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Cylinder" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "OuterConductor" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".OuterRadius " + Chr(34) + "OuterDiameter/2+OuterConductorThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".InnerRadius " + Chr(34) + "OuterDiameter/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Axis " + Chr(34) + "z" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xcenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Ycenter " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Segments " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define cylinder: "+sLineType+":OuterConductor", sCommand)

		Case "Stripline", "Thin Stripline", "Thick Stripline", "Asymmetric Thick Stripline"
        	StoreParameter("TraceThickness", IIf(InStr(sLineType, "Thick Stripline")>0, sParameter3, "0"))
        	If (sLineType = "Asymmetric Thick Stripline") Then
				StoreParameter("SubstrateHeightTop", sParameter1)
				StoreParameter("SubstrateHeightBottom", sParameter4)
				StoreParameter("SubstrateHeight", "SubstrateHeightTop + TraceThickness + SubstrateHeightBottom")
			Else
	        	StoreParameter("SubstrateHeight", sParameter1)
	        End If
        	StoreParameter("TraceWidth", sParameter2)
        	StoreParameter("GroundThickness", IIf(InStr(sLineType, "Thick Stripline")>0, sParameter3, "0"))
        	StoreParameter("SubstrateWidth", 4*(Evaluate(sParameter1)+Evaluate(sParameter2)))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth/2" + Chr(34) + ", " + Chr(34) + "TraceWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-TraceThickness/2" + Chr(34) + ", " + Chr(34) + "TraceThickness/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace", sCommand)

			' Move trace in case of asymmetric case
        	If (sLineType = "Asymmetric Thick Stripline") Then
				sCommand = ""
				sCommand = sCommand + "With Transform" + vbNewLine
				sCommand = sCommand + ".Reset" + vbNewLine
				sCommand = sCommand + ".Name " + Chr(34) + sLineType+":Trace" + Chr(34) + vbNewLine
				sCommand = sCommand + ".Vector " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "-TraceThickness/2+SubstrateHeightBottom/2-SubstrateHeightTop/2" + Chr(34) + ", " + Chr(34) + "0" + Chr(34) + vbNewLine
				sCommand = sCommand + ".UsePickedPoints " + Chr(34) + "False" + Chr(34) + vbNewLine
				sCommand = sCommand + ".InvertPickedPoints " + Chr(34) + "False" + Chr(34) + vbNewLine
				sCommand = sCommand + ".MultipleObjects " + Chr(34) + "False" + Chr(34) + vbNewLine
				sCommand = sCommand + ".GroupObjects " + Chr(34) + "False" + Chr(34) + vbNewLine
				sCommand = sCommand + ".Repetitions " + Chr(34) + "1" + Chr(34) + vbNewLine
				sCommand = sCommand + ".MultipleSelection " + Chr(34) + "False" + Chr(34) + vbNewLine
				sCommand = sCommand + ".Transform " + Chr(34) + "Shape" + Chr(34) + ", " + Chr(34) + "Translate" + Chr(34) + vbNewLine
				sCommand = sCommand + "End With" + vbNewLine
				AddToHistory("transform: translate"+sLineType+":Trace", sCommand)
			End If

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "SubstrateHeight/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "SubstrateHeight/2+GroundThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopGround", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "BottomGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-GroundThickness-SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":BottomGround", sCommand)

        Case "Thin Microstrip", "Thick Microstrip"
   			StoreParameter("SubstrateHeight", sParameter1)
        	StoreParameter("TraceWidth", sParameter2)
        	StoreParameter("TraceThickness", IIf(sLineType = "Thick Microstrip", sParameter3, "0"))
        	StoreParameter("GroundThickness", IIf(sLineType = "Thick Microstrip", sParameter3, "0"))
        	StoreParameter("SubstrateWidth", 4*(Evaluate(sParameter1)+Evaluate(sParameter2)))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth/2" + Chr(34) + ", " + Chr(34) + "TraceWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Ground" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-GroundThickness-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Ground", sCommand)

		Case "cpw_mg"
   			StoreParameter("SubstrateHeight", sParameter1)
        	StoreParameter("TraceWidth", sParameter2)
        	StoreParameter("TraceThickness", "0")
        	StoreParameter("GapWidth", sParameter3)
        	StoreParameter("GroundThickness", "0")
        	StoreParameter("SubstrateWidth", 4*(Evaluate(sParameter1)+Evaluate(sParameter2)))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth/2" + Chr(34) + ", " + Chr(34) + "TraceWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopLeftGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "-TraceWidth/2-GapWidth" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopLeftGround", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopRightGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "TraceWidth/2+GapWidth" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopRightGround", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "BottomGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-GroundThickness-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":BottomGround", sCommand)

		Case "cpw", "thick_cpw"
        	StoreParameter("TraceWidth", sParameter1)
        	StoreParameter("GapWidth", sParameter2)
        	StoreParameter("TraceThickness", IIf(sLineType="cpw", "0", sParameter3))
   			StoreParameter("SubstrateHeight", IIf(sLineType="cpw", Evaluate(20*Evaluate(sParameter1)), sParameter4))
        	StoreParameter("GroundThickness", "0")
        	StoreParameter("SubstrateWidth", 20*Evaluate(sParameter1))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth/2" + Chr(34) + ", " + Chr(34) + "TraceWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopLeftGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "-TraceWidth/2-GapWidth" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopLeftGround", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopRightGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "TraceWidth/2+GapWidth" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopRightGround", sCommand)

        Case "diff_strip"
        	StoreParameter("TraceWidth", sParameter1)
        	StoreParameter("SubstrateHeight", sParameter2)
        	StoreParameter("TraceThickness", sParameter3)
        	StoreParameter("GapWidth", sParameter4)
        	StoreParameter("GroundThickness", sParameter3)
        	StoreParameter("SubstrateWidth", 4*(2*Evaluate(sParameter1)+Evaluate(sParameter3)+Evaluate(sParameter2)))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace1" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth-GapWidth/2" + Chr(34) + ", " + Chr(34) + "-GapWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-TraceThickness/2" + Chr(34) + ", " + Chr(34) + "TraceThickness/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace1", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "GapWidth/2" + Chr(34) + ", " + Chr(34) + "GapWidth/2+TraceWidth" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-TraceThickness/2" + Chr(34) + ", " + Chr(34) + "TraceThickness/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace2", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "SubstrateHeight/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "TopGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "SubstrateHeight/2+GroundThickness" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":TopGround", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "BottomGround" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-GroundThickness-SubstrateHeight/2" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":BottomGround", sCommand)

        Case "suspended_strip", "inverse_suspended_strip"
        	StoreParameter("TraceWidth", sParameter1)
			StoreParameter("SubstrateHeight", sParameter2)
			StoreParameter("SuspensionHeight", sParameter3)
        	StoreParameter("TraceThickness", "0")
        	StoreParameter("GroundThickness", "0")
        	StoreParameter("SubstrateWidth", 4*(Evaluate(sParameter1)+Evaluate(sParameter2)))

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Trace" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-TraceWidth/2" + Chr(34) + ", " + Chr(34) + "TraceWidth/2" + Chr(34) + vbNewLine
			If sLineType = "suspended_strip" Then
				sCommand = sCommand + ".Yrange " + Chr(34) + "0" + Chr(34) + ", " + Chr(34) + "TraceThickness" + Chr(34) + vbNewLine
			Else
				sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight-TraceThickness" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight" + Chr(34) + vbNewLine
			End If
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Trace", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Substrate" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Dielectric" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight" + Chr(34) + ", " + Chr(34) + "0" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Substrate", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "SuspensionGap" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "Vacuum" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-SubstrateHeight-SuspensionHeight" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":SuspensionGap", sCommand)

			sCommand = ""
			sCommand = sCommand + "With Brick" + vbNewLine
			sCommand = sCommand + ".Reset" + vbNewLine
			sCommand = sCommand + ".Name " + Chr(34) + "Ground" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Component " + Chr(34) + sLineType + Chr(34) + vbNewLine
			sCommand = sCommand + ".Material " + Chr(34) + "PEC" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Xrange " + Chr(34) + "-SubstrateWidth/2" + Chr(34) + ", " + Chr(34) + "SubstrateWidth/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Yrange " + Chr(34) + "-GroundThickness-SubstrateHeight-SuspensionHeight" + Chr(34) + ", " + Chr(34) + "-SubstrateHeight-SuspensionHeight" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Zrange " + Chr(34) + "-Length/2" + Chr(34) + ", " + Chr(34) + "Length/2" + Chr(34) + vbNewLine
			sCommand = sCommand + ".Create" + vbNewLine
			sCommand = sCommand + "End With" + vbNewLine
			AddToHistory("define brick: "+sLineType+":Ground", sCommand)

		Case Else
			MsgBox("Unknown line type.")

	End Select

End Function
