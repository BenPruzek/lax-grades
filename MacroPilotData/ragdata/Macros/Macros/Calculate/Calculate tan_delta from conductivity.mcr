' *Calculate / Calculate tan_delta from conductivity
' !!! Do not change the line above !!!

' ================================================================================================
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 17-Nov-2010 ube: USformat
' 05-May-2010 fsr: added digits to result, changed to project units, added reverse calculation
' macro.939

Option Explicit
'#include "vba_globals_all.lib"

Sub Main () 

	Dim cst_frq As Double
	Dim cst_frq_unit As Double
	Dim cst_frq_unitS As String
	Dim cst_epsr As Double
	Dim cst_input As Double
	Dim cst_tand As Double
	Dim cst_output As Double
	Dim cst_output_Text As String
	Dim cst_input_List() As String
	Dim cst_mode As Integer

	cst_frq_unit = Units.GetFrequencyUnitToSI
	cst_frq_unitS = Units.GetUnit("Frequency")

	cst_frq=1.0
	cst_epsr=1.0
	cst_input=0.01
	cst_tand=0

	cst_mode = 0
	ReDim cst_input_List(1)
	cst_input_List(0) = "conductivity [S/m]"
	cst_input_List(1) = "tan-delta"

	While True

		Select Case cst_mode
			Case 0 ' kappa->tand
				' tand = kappa / omega eps
				cst_output = cst_input / ( 2*Pi * cst_frq * cst_frq_unit * cst_epsr * Eps0)
				'		cst_tand = cst_kappa / ( 2*Pi * cst_frq * cst_epsr)
				cst_output_Text = "tan-delta = "
			Case 1 ' tand->kappa
				cst_output = cst_input * ( 2*Pi * cst_frq * cst_frq_unit * cst_epsr * Eps0)
				cst_output_Text = "kappa [S/m] = "
		End Select


	Begin Dialog UserDialog 300,175,"Loss Angle tan-delta" ' %GRID:10,7,1,1
		Text 30,21,120,14,"frequency  ["+cst_frq_unitS+"]",.Text1
		Text 30,77,120,14,"eps_relative",.Text3
		TextBox 190,21,90,21,.frq
		DropListBox 25,49,150,192,cst_input_List(),.DropListBox1
		TextBox 190,49,90,21,.cstinput
		TextBox 190,77,90,21,.epsr
		OKButton 50,140,90,21
		CancelButton 160,140,90,21
		Text 30,112,260,14,cst_output_Text + USFormat(cst_output,"0.000E+00"),.Text4
	End Dialog
		Dim dlg As UserDialog
		dlg.frq=CStr(cst_frq)
		dlg.epsr=CStr(cst_epsr)
		dlg.cstinput=CStr(cst_input)
		
		If (Dialog(dlg) = 0) Then Exit All
		
		cst_frq = Evaluate(dlg.frq)
		cst_epsr = Evaluate(dlg.epsr)
		cst_input = Evaluate(dlg.cstinput)
		cst_mode = dlg.DropListBox1
		
	Wend	


End Sub
