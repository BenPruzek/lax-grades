' Calculate Buneman Hartree Field

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

	Begin Dialog UserDialog 520,126,"Calculate Buneman Hartree (PI Mode)" ' %GRID:10,3,1,1
		TextBox 120,14,140,21,.ra
		TextBox 120,42,140,21,.rc
		TextBox 120,70,140,21,.Voltage
		Text 20,18,60,15,"ra [mm] :",.Text1
		Text 20,46,60,15,"rc [mm] :",.Text2
		Text 20,74,90,15,"Voltage [kV] :",.Text3
		OKButton 320,99,90,21
		CancelButton 420,99,90,21
		Text 280,18,70,15,"f [GHz] :",.Text4
		TextBox 370,14,140,21,.freq
		Text 280,46,70,15,"Nvanes : ",.Text5
		TextBox 370,42,140,21,.N
	End Dialog

	Dim dlg As UserDialog
	dlg.ra = "0"
	dlg.rc = "0"
	dlg.Voltage = "0"
	dlg.freq    = "0"
	dlg.N       = "0"

	If (Dialog(dlg) = 0) Then Exit All

	Dim cst_ra As Double
	Dim cst_rc As Double
	Dim cst_V As Double
	Dim cst_N As Integer, nmod As Integer
	Dim freq As Double, wn As Double

	Dim cst_gamma As Double
	Dim de_S As Double, de_P As Double 'S= Schamiloglu P = Palevsky

	cst_ra = Evaluate(dlg.ra)
	cst_rc = Evaluate(dlg.rc)
	cst_V  = Evaluate(dlg.Voltage)
	cst_N  = Evaluate(dlg.N)
	freq   = Evaluate(dlg.freq)

	nmod  = cst_N/2 'Assumption PI Mode
	wn    = 2*Pi*freq*1e9 ' Conversion to Si

	cst_gamma = 1+cst_V/511

	de_S = (cst_ra^2 - cst_rc^2)/(2*cst_ra)
	de_P = (cst_ra^2 - cst_rc^2)/(2*cst_rc)

	Dim K1 As Double, K2_S As Double, K2_P As Double 'Hilfskonstanten siehe Schamiloglu Buch aufloesen nach Bz

	K1 = cst_gamma - Sqr( 1 - (cst_ra*1e-3*wn/(3e8*nmod))^2)

	K2_S = cst_ra * de_S * 1e-6* wn/(511*1e3*nmod)
	K2_P = cst_ra * de_P * 1e-6* wn/(511*1e3*nmod)


	Dim B_BH_S As Double, B_BH_P As Double

	B_BH_S = K1/K2_S
	B_BH_P = K1/K2_P

	MsgBox "Buneman Hartree Threshold according to Schamiloglu's Book" + vbCrLf  + "B_BH / T =" + Format(B_BH_S,"0.0000") ' + vbCrLf + vbCrLf + "Buneman Hartree Threshold according to Palevsky" + vbCrLf  + "B* / T =" + Cstr (B_BH_P)


End Sub
