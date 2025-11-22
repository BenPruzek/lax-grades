' Estimate Numerical Cerenkov Noise Frequency

Option Explicit

'#include "vba_globals_all.lib"

' ================================================================================================
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 10-Dec-2015 ube: added Help button and online help
' 25-Aug-2015 mbk: initial version
'------------------------------------------------------------------------------------
Const HelpFileName = "common_preloadedmacro_Estimate_Numerical_Cerenkov_Noise_Frequency"

Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses Any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------

	If (Action%=1 Or Action%=2) Then
			' Action%=1: The dialog box is initialized
			' Action%=2: The user changes a value or presses a button

		If (DlgItem = "Help") Then
			StartHelp HelpFileName
			DialogFunction = True
		End If

		If (DlgItem = "OK") Then

		    ' The user pressed the Ok button. Check the settings and display an error message if some required
		    ' fields have been left blank.

		End If

	End If
End Function

Sub Main ()

	Begin Dialog UserDialog 400,98,"Estimate Numerical Cerenkov Noise Frequency",.DialogFunction ' %GRID:10,7,1,1
		Text 10,14,260,14,"Average Longitudinal Mesh Step / [mm]:",.Text1
		TextBox 290,14,90,21,.dx
		Text 10,42,160,14,"Beam Energy/ [keV]:",.Text3
		TextBox 290,42,90,21,.Vbeam
		OKButton 10,70,90,21
		CancelButton 110,70,90,21
		PushButton 290,70,90,21,"Help",.Help
	End Dialog
	Dim dlg As UserDialog

	dlg.dx = "0"
	dlg.Vbeam = "0"

	If (Dialog(dlg) = 0) Then Exit All

	Dim dx As Double
	Dim Vbeam As Double
	Dim fmax  As Double

	'Conversion to SI
	dx = Evaluate(dlg.dx)*1e-3
	Vbeam = Evaluate(dlg.Vbeam)*1e3

	fmax  = clight/(Pi*dx)

	Dim rgamma As Double, rbeta As Double, rvelo As Double
	rgamma= 1+Abs(Vbeam)/511e3
	rbeta = Sqr(1-1/(rgamma*rgamma))
	rvelo = clight*rbeta

	'Begin Draw Beam Line

	Dim stmpfile As String
	stmpfile = "Test1D_tmp.txt"

	Dim r1d As Object
	Set r1d = Result1D("")

	r1d.Initialize(2)

	r1d.SetXYDouble 0, fmax/1000/Units.GetFrequencyUnitToSI, rvelo
	r1d.SetXYDouble 1, fmax/Units.GetFrequencyUnitToSI, rvelo

	r1d.Save stmpfile
	r1d.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
	r1d.SetYLabelAndUnit "" , "m/s"
	r1d.Title "Grid Dispersion Relation"
	r1d.AddToTree "1D Results\Grid Dispersion Relation\Particle Velocity"

	'Draw Grid Dispersion

	Dim stmpfile2 As String
	stmpfile2 = "PhaseVelocity.txt"

	Dim r2d As Object
	Set r2d = Result1D("")

	r2d.Initialize(1000)

	Dim ip As Integer, ftemp As Double, xx As Double
	For ip = 1 To 1000
		ftemp = ip*fmax/1000
		xx = 2*pi*ftemp*dx/(2*CLight)
		r2d.SetXYDouble ip-1, ftemp/Units.GetFrequencyUnitToSI, (2*pi*ftemp*dx/2)/ asin((2*pi*ftemp*dx/(2*CLight)))
	Next ip

	r2d.Save stmpfile2
	r2d.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
	r2d.SetYLabelAndUnit "" , "m/s"
	r2d.Title "Grid Dispersion Relation"
	r2d.AddToTree "1D Results\Grid Dispersion Relation\Numerical Phase Velocity"

	SelectTreeItem("1D Results\Grid Dispersion Relation")
		
End Sub
