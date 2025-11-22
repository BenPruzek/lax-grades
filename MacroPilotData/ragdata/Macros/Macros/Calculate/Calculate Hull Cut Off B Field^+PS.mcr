' Calculate Hull Cut Off B Field

'#Language "WWB-COM"

' ------------------------------------------------------------------------------------------
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 04-Jan-2016 mbk: first version
' ------------------------------------------------------------------------------------------

Option Explicit

'#include "vba_globals_all.lib"

Sub Main

	Begin Dialog UserDialog 310,133,"Calculate Hull Cut Off" ' %GRID:10,3,1,1
		TextBox 150,14,140,21,.ra
		TextBox 150,42,140,21,.rc
		TextBox 150,70,140,21,.Voltage
		Text 20,18,120,15,"ra [mm] :",.Text1
		Text 20,46,120,15,"rc [mm] :",.Text2s
		Text 20,74,110,15,"Voltage [kV] :",.Text3
		OKButton 20,105,90,21
		CancelButton 130,105,90,21
	End Dialog
	Dim dlg As UserDialog
	dlg.ra = "0"
	dlg.rc = "0"
	dlg.Voltage = "0"

	If (Dialog(dlg) = 0) Then Exit All

	Dim cst_ra As Double
	Dim cst_rc As Double
	Dim cst_V As Double

	cst_ra = Evaluate(dlg.ra)
	cst_rc = Evaluate(dlg.rc)
	cst_V  = Evaluate(dlg.Voltage)

	Dim cst_gamma As Double
	Dim de_S As Double, de_P As Double 'S= Schamiloglu P = Palevsky in Palevsky's Diss there's most likely a mistake in de
	Dim B_HC_S As Double, B_HC_P As Double

	cst_gamma = 1+cst_V/511

	de_S = (cst_ra^2 - cst_rc^2)/(2*cst_ra)
	de_P = (cst_ra^2 - cst_rc^2)/(2*cst_rc)


	B_HC_S = 511/3*Sqr(cst_gamma^2-1) *1e-2 /de_S
	B_HC_P = 511/3*Sqr(cst_gamma^2-1) *1e-2 /de_P

	MsgBox "Hull Cut Off Field according to Schamiloglu's Book" + vbCrLf  + vbCrLf  + "B* / T = " + Format(B_HC_S,"0.0000")  '+ vbCrLf + vbCrLf + "Hull Cut Off Field according To Palevsky" + vbCrLf  + "B* / T =" + Cstr (B_HC_P)


End Sub
