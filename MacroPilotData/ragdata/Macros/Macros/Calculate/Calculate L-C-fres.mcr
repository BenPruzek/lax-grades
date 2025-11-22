' *Calculate / Calculate L-C-fres
' !!! Do not change the line above !!!

' macro.548
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 13-Nov-2003 ube: option explicit added
'------------------------------------------------------------------------------------

Option Explicit

Sub Main () 

	Dim cst_frq As Double
	Dim cst_L As Double
	Dim cst_C As Double, i_result As Integer

	cst_L=100.0
	cst_C=100.0
	i_result = 2

	While True
	
	Begin Dialog UserDialog 340,154,"Calculate L - C - res.frq" ' %GRID:10,7,1,1
		Text 60,28,50,14,"L  (nH)",.Text1
		Text 60,56,60,14,"C  (pF)",.Text2
		Text 60,84,80,14,"f_res (MHz)",.Text3
		TextBox 150,28,170,21,.L
		TextBox 150,56,170,21,.C
		TextBox 150,84,170,21,.f
		OKButton 30,119,90,21
		CancelButton 130,119,90,21
		OptionGroup .Group1
			OptionButton 30,28,20,14,"OptionButton1",.OptionButton1
			OptionButton 30,56,20,14,"OptionButton2",.OptionButton2
			OptionButton 30,84,20,14,"OptionButton3",.OptionButton3
		Text 20,7,120,14,"check result value",.Text4
	End Dialog
		Dim dlg As UserDialog
		dlg.Group1 = i_result

		Select Case i_result
			Case 0
				dlg.C=CStr(cst_C)
				dlg.f=CStr(cst_frq)

				cst_L =  1.0e9 / (cst_C*1e-12) / (2 * Pi * cst_frq*1e6)^2
				dlg.L=CStr(cst_L)
			Case 1
				dlg.L=CStr(cst_L)
				dlg.f=CStr(cst_frq)

				cst_C =  1.0e12 / (cst_L*1e-9) / (2 * Pi * cst_frq*1e6)^2
				dlg.C=CStr(cst_C)
			Case 2
				dlg.L=CStr(cst_L)
				dlg.C=CStr(cst_C)

				cst_frq =  1.0e-6 / Sqr(cst_L*1e-9 * cst_C*1e-12) / (2 * Pi)
				dlg.f=CStr(cst_frq)
		End Select

		If (Dialog(dlg) = 0) Then Exit All
		
		cst_frq = Evaluate(dlg.f)
		cst_L = Evaluate(dlg.L)
		cst_C = Evaluate(dlg.C)
		i_result = dlg.Group1
		
	Wend	


End Sub

