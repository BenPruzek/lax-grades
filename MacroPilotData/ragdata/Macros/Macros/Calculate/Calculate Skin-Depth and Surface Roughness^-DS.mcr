' *Calculate / Calculate Skin-Depth and Surface Roughness
' !!! Do not change the line above !!!
'
' macro.547 (formerly 937, but also needed for EMS)
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'----------------------------------------------------------------------------------------------------------------
' 08-Oct-2010 jwa: change the equation of corrected conductivity for roughness
' 03-Nov-2008 msc: added info about project unit
' 29-Oct-2008 ube: 2 more digits so that 2um are visible in meter as well
' 28-May-2008 msc: Skin depth in local geometry units now (metric and imperial)
' 30-Jul-2003 ube: EMS gets other defaults (Hz + mm) than MWS (GHz + um)
' 27-Jun-2003 jsw: add effective conductivity for roughness
'----------------------------------------------------------------------------------------------------------------

Option Explicit

'#include "vba_globals_all.lib"

Sub Main ()

	Dim cst_frq As Double
	Dim cst_mu_r As Double
	Dim cst_kappa As Double
	Dim cst_rough As Double
	Dim skindepth As Double
	Dim effecond As Double

	Dim sFrqUnit As String
	Dim dFrqFactor As Double

	Dim sSkinUnit As String
	Dim dSkinFactor As Double
	Dim cra As Double

	GetLenghtUnitAndFactor (dSkinFactor , sSkinUnit)

	If bMWS Then
		cst_frq=1.0
		sFrqUnit = "[GHz]"
		dFrqFactor = 1.0e9
	Else
		cst_frq=50.0
		sFrqUnit = "[Hz]"
		dFrqFactor = 1.0
	End If

	cst_mu_r=1.0
	cst_kappa=5.8e7
	cst_rough=0.0

    effecond=cst_kappa

	While True
		' skindepth = sqr ( 2 / (2 pi f * mue * kappa) )
		skindepth = 1 / ( 2*pi * Sqr (cst_frq * dFrqFactor * 1.0e-7 * cst_mu_r * cst_kappa))
		' MsgBox "Skin-Depth = " + CStr(Format(skindepth*1.0e6,"Standard")) + " micrometer"

		If (cst_rough=0) Then
			effecond=cst_kappa
		Else
			'effecond=cst_kappa/(1+Exp(-skindepth/cst_rough*1e6)^1.6)^2
			cra=1+2/pi*Atn(1.4*(cst_rough/skindepth*1e-6)^2)
            effecond=(cst_kappa/cra)^2/cst_kappa
		End If

	Begin Dialog UserDialog 320,245,"Skindepth and Surface Roughness" ' %GRID:10,7,1,1
		Text 30,21,120,14,"Frequency  " + sFrqUnit,.Text1
		Text 30,49,120,14,"Conductivity [S/m]",.Text2
		Text 30,77,120,14,"mu_relative",.Text3
		TextBox 170,21,110,21,.frq
		TextBox 170,49,110,21,.kappa
		TextBox 170,77,110,21,.mu_r
		TextBox 170,105,110,21,.rough
		OKButton 60,217,90,21
		CancelButton 160,217,90,21
		Text 30,140,260,14,"Skin-Depth = " + Format(skindepth*dSkinFactor,"0.000000") + sSkinUnit + " (project unit)",.Text4
		Text 30,189,270,14,"eff. cond.   = " + Format(effecond,"Standard") + " S/m",.Text6
		Text 30,105,120,14,"Roughness [um]",.Text7
		Text 30,168,270,14,"Effective Conductivity for Rough Surface:",.Text5
	End Dialog
		Dim dlg As UserDialog
		dlg.frq=CStr(cst_frq)
		dlg.mu_r=CStr(cst_mu_r)
		dlg.kappa=CStr(cst_kappa)
		dlg.rough=CStr(cst_rough)

		If (Dialog(dlg) = 0) Then Exit All
		
		cst_frq = Evaluate(dlg.frq)
		cst_mu_r = Evaluate(dlg.mu_r)
		cst_kappa = Evaluate(dlg.kappa)
        cst_rough = Evaluate(dlg.rough)
	Wend


End Sub

Function GetLenghtUnitAndFactor (dFac As Double, sLabel As String)

	sLabel = Units.GetUnit("Length")

    Select Case sLabel
    Case "m"
        dFac = 1.0
    Case "cm"
        dFac = 0.01
    Case "mm"
        dFac = 1e-3
    Case "um"
        dFac = 1e-6
    Case "nm"
        dFac = 1e-9
    Case "ft"
        dFac = 0.3048
    Case "in"
        dFac = 0.0254
    Case "mil"
        dFac = 2.54e-5
    End Select

    dFac = 1/dFac

End Function
