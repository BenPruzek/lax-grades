' *Calculate / Calculate Wavelength
' !!! Do not change the line above !!!

' macro.546
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 12-Jun-2010 ahfr,dit: modified to give 1/4 and 1/2 wavelentgh (instead of 1/10)
' 02-Apr-2010 fsr: modified to use project units
' 13-Nov-2003 ube: option axplicit added
'------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Sub Main () 

	Dim cst_frq As Double
	Dim cst_mu_r As Double
	Dim cst_epsr As Double
	Dim cst_wavelength As Double

	cst_frq=5.0
	cst_mu_r=1.0
	cst_epsr=1.0

	While True
	
		' wavelength = clight / frq
		cst_wavelength =  CLight / Sqr(cst_mu_r*cst_epsr) / (cst_frq * Units.GetFrequencyUnitToSI)

		If (Units.GetUnit("Length") = "m") Then
			Begin Dialog UserDialog 340,242,"Calculate Wavelength" ' %GRID:10,7,1,1
				Text 30,21,120,14,"frequency  ["+Units.GetUnit("Frequency")+"]",.Text1
				Text 30,49,120,14,"eps_relative",.Text2
				Text 30,77,120,14,"mu_relative",.Text3
				TextBox 195,21,120,21,.frq
				TextBox 195,49,120,21,.epsr
				TextBox 195,77,120,21,.mu_r
				OKButton 65,214,90,21
				CancelButton 175,214,90,21
				Text 30,110,290,14,"Wavelength = " + Format(cst_wavelength,"Scientific") + " m",.Text4
				Text 30,135,290,14,"Half-Wavelength = " + Format(cst_wavelength/2,"Scientific") + " m",.Text5
				Text 30,160,290,14,"Quarter-Wavelength = " + Format(cst_wavelength/4,"Scientific") + " m",.Text6
				Text 30,185,290,14,"speed of light = " + Format(1/ Sqr(cst_mu_r*cst_epsr),"Standard") + " * c0",.Text7
			End Dialog
			Dim dlg1 As UserDialog
			dlg1.frq=CStr(cst_frq)
			dlg1.epsr=CStr(cst_epsr)
			dlg1.mu_r=CStr(cst_mu_r)

			If (Dialog(dlg1) = 0) Then Exit All

			cst_frq = Evaluate(dlg1.frq)
			cst_mu_r = Evaluate(dlg1.mu_r)
			cst_epsr = Evaluate(dlg1.epsr)
		Else
			Begin Dialog UserDialog 340,242,"Calculate Wavelength" ' %GRID:10,7,1,1
				Text 30,21,120,14,"frequency  ["+Units.GetUnit("Frequency")+"]",.Text1
				Text 30,49,120,14,"eps_relative",.Text2
				Text 30,77,120,14,"mu_relative",.Text3
				TextBox 195,21,120,21,.frq
				TextBox 195,49,120,21,.epsr
				TextBox 195,77,120,21,.mu_r
				OKButton 65,214,90,21
				CancelButton 175,214,90,21
				Text 30,110,290,14,"Wavelength = " + Format(cst_wavelength/Units.GetGeometryUnitToSI,"Standard") + " " + Units.GetUnit("Length"),.Text4
				Text 30,135,290,14,"Half Wavelength = " + Format(cst_wavelength/Units.GetGeometryUnitToSI/2,"Standard") + " " + Units.GetUnit("Length"),.Text5
				Text 30,160,290,14,"Quarter Wavelength = " + Format(cst_wavelength/Units.GetGeometryUnitToSI/4,"Standard") + " " + Units.GetUnit("Length"),.Text6
				Text 30,185,290,14,"speed of light = " + Format(1/ Sqr(cst_mu_r*cst_epsr),"Standard") + " * c0",.Text7
			End Dialog
			Dim dlg2 As UserDialog
			dlg2.frq=CStr(cst_frq)
			dlg2.epsr=CStr(cst_epsr)
			dlg2.mu_r=CStr(cst_mu_r)

			If (Dialog(dlg2) = 0) Then Exit All

			cst_frq = Evaluate(dlg2.frq)
			cst_mu_r = Evaluate(dlg2.mu_r)
			cst_epsr = Evaluate(dlg2.epsr)
		End If

	Wend	


End Sub

