'#Language "WWB-COM"

Option Explicit

' This macro generates and imports a farfield source from an analytical equation

' Version history:
' ================================================================================================
' Copyright 2012-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' --------------------------------------------------------------------------------------------------------------------------------------------------------------
' 01-Feb-2024 ube: Added "Wide beam horiz-pol", useful as Approximate Automotive Radar Farfield
' 12-Oct-2022 mha: Changed parentheses for combinations of "CStr" and "Evaluate" - should now correctly evalute regardless of "." or ","
' 04-Jan-2022 hch,ube: added isotropic ff sources, changed default to "Isotropic - E-theta lin-pol"
' 14-Sep-2017 fsr: Linked to online help; turned back time by 800 years
' 28-Aug-2017 fsr: Introduced additional validity check for input parameters; small code improvements
' 21-Aug-2017 jpl: Added sample values for preloaded field parameters, improved Circular Horn equations, and added a check to ensure input parameters are valid
' 30-Aug-2013 yta: First version of Gaussian beam
' 03-Aug-2012 fsr: Added name check, some preloaded field expressions
' 26-Jan-2012 fsr: Initial version
' --------------------------------------------------------------------------------------------------------------------------------------------------------------

'#include "vba_globals_all.lib"

Private Const HelpFileName = "common_preloadedmacro_Analytical_Farfield_Source"

' Format of constant: Name, EThRe, EThIm, EPhRe, EPhIm, Parameter list
Public Const PreloadedFields = Array(	Array("Preloaded Fields...", "SinD(TH)", "0", "0", "0", ""), _
										Array("Half-Wave Dipole", "0", "Mue0*f*I0*LDipole*sinD(TH)/2", "0", "0", "I0|LDipole"), _
										Array("Small Circular Loop", "0", "0", "Mue0*rLoop^2*I0*4*pi^2*f^2*sinD(TH)/(4*CLight)", "0", "I0|rLoop"), _
										Array("Circular Horn", "1/300*-2*pi*f/2*Mue0*(1+Sqr((2*pi*f/CLight)^2-(1.8412/((B-SinD(TH))*SinD(Alpha)))^2)/(2*pi*f/CLight)*CosD(TH))*cyl_bessel_j(1,1.8412/((B-SinD(TH))*SinD(Alpha))*a1)*cyl_bessel_j(1,2*pi*f/CLight*SinD(TH)*a1)/SinD(TH+1e-15)*SinD(PH)*Cos(-2*pi*f/CLight)", "1/300*-2*pi*f/2*Mue0*(1+Sqr((2*pi*f/CLight)^2-(1.8412/((B-SinD(TH))*SinD(Alpha)))^2)/(2*pi*f/CLight)*CosD(TH))*cyl_bessel_j(1,1.8412/((B-SinD(TH))*SinD(Alpha))*a1)*cyl_bessel_j(1,2*pi*f/CLight*a1*SinD(TH))/SinD(TH+1e-15)*SinD(PH)*Sin(-2*pi*f/CLight)", "1/300*-2*pi*f/CLight*a1*2*pi*f*Mue0/2*(Sqr((2*pi*f/CLight)^2-(1.8412/((B-SinD(TH))*SinD(Alpha)))^2)/(2*pi*f/CLight)+CosD(TH))*cyl_bessel_j(1,1.8412/((B-SinD(TH))*SinD(Alpha)))*1/2*(cyl_bessel_j(0,2*pi*f/CLight*a1*SinD(TH))-cyl_bessel_j(2,2*pi*f/CLight*a1*SinD(TH)))/(1-(2*pi*f/CLight*SinD(TH)/(1.8412/((B-SinD(TH))*SinD(Alpha))))^2)*CosD(PH)*Cos(-2*pi*f/CLight)", "1/300*-2*pi*f/CLight*a1*2*pi*f*Mue0/2*(Sqr((2*pi*f/CLight)^2-(1.8412/((B-SinD(TH))*SinD(Alpha)))^2)/(2*pi*f/CLight)+CosD(TH))*cyl_bessel_j(1,1.8412/((B-SinD(TH))*SinD(Alpha)))*1/2*(cyl_bessel_j(0,2*pi*f/CLight*a1*SinD(TH))-cyl_bessel_j(2,2*pi*f/CLight*a1*SinD(TH)))/(1-(2*pi*f/CLight*SinD(TH)/(1.8412/((B-SinD(TH))*SinD(Alpha))))^2)*CosD(PH)*Sin(-2*pi*f/CLight)", "a1|Alpha|B"), _
										Array("Gaussian Beam", "Nf*CLight*Exp(2*Pi*f*b*cosD(TH)/CLight)*(1+cosD(TH))*cosD(PH)/(2*Pi*f)", "0", "-Nf*CLight*Exp(2*Pi*f*b*cosD(TH)/CLight)*(1+cosD(TH))*sinD(PH)/(2*Pi*f)", "0", "b|Nf"), _
										Array("Isotropic - E-theta lin-pol", "sqr(377/(4*Pi))", "0", "0", "0", ""), _
										Array("Isotropic - E-phi lin-pol", "0", "0", "sqr(377/(4*Pi))", "0", ""), _
										Array("Isotropic - LHCP", "sqr(377/(2*4*Pi))", "0", "0",  "sqr(377/(2*4*Pi))", ""), _
										Array("Isotropic - RHCP", "sqr(377/(2*4*Pi))", "0", "0", "-sqr(377/(2*4*Pi))", ""), _
										Array("Wide beam horiz-pol", "0", "0", "21.8*(CosD(PH/2)^4)*(SinD(TH)^14)", "0", ""))

' Circular horn equation based on paper: P. Hajach, "A Conical Dielectric-Loaded Horn with Symetrical Radiation Pattern," Radioengineering, vol. 2, no.3, Nov. 1993. and scaled by 1/300 to match simulation results
' small circular loop EPhRe: 0.01*Pi*I0*Sqr(Mue0/Eps0)*f/CLight*cyl_bessel_j(1,2*Pi*f/CLight*rLoop*SinD(TH))
Sub Main
	Dim i As Long, sPreloadedFieldsList() As String

	ReDim sPreloadedFieldsList(UBound(PreloadedFields))
	For i = 0 To UBound(PreloadedFields)
		sPreloadedFieldsList(i) = PreloadedFields(i)(0)
	Next

	Begin Dialog UserDialog 610,275,"Generate Analytical Farfield Source",.DialogFunction ' %GRID:10,5,1,1
		Text         420, 225, 110,  15, "txt3",.txt3
		Text          80, 225,  60,  15, "txt1",.txt1
		Text         250, 225,  90,  15, "txt2",.txt2
		Text          20,  15, 140,  15, "Farfield source name:",.Text1
		TextBox      190,  10, 390,  20, .FFSourceNameT
		Text          20,  40, 140,  15, "Angular resolution:",.Text2
		Text         190,  40,  50,  15, "Theta:",.Text5
		Text         300,  40,  30,  15, "Phi:",.Text10
		TextBox      240,  35,  50,  20, .ResolutionThetaT
		TextBox      340,  35,  50,  20, .ResolutionPhiT
		Text         410,  40,  60,  15, "[degrees]",.Text6
		Text          20,  65, 140,  15, "Frequency range:",.Text7
		Text         190,  65,  40,  15, "From",.Text8
		TextBox      240,  60,  50,  20, .FminT
		Text         310,  65,  20,  15, "to",.Text9
		TextBox      340,  60,  50,  20, .FmaxT
		Text         400,  65,  70,  15, ", resolution",.Text11
		TextBox      480,  60,  50,  20, .DeltaFT
		Text         540,  65,  40,  15, "["+Units.GetUnit("Frequency")+"]",.Text12
		Text          20,  95, 230,  15, "Expressions for E-field components:",.Text3
		DropListBox  260,  90, 320,  20, sPreloadedFieldsList(),.PreloadedFieldsDLB
		Text          30, 120, 150,  15, "E_theta_re(TH,PH,f) =",.exprv3
		TextBox      190, 115, 390,  20, .EThetaRealT
		Text          30, 145, 150,  15, "E_theta_im(TH,PH,f) =",.exprv2
		TextBox      190, 140, 390,  20, .EThetaImagT
		Text          30, 175, 150,  15, "E_phi_re(TH,PH,f)    =",.exprv
		TextBox      190, 170, 390,  20, .EPhiRealT
		Text          30, 200, 150,  15, "E_phi_im(TH,PH,f)    =",.exprv4
		TextBox      190, 195, 390,  20, .EPhiImagT
		TextBox      150, 220,  60,  20, .para1
		TextBox      340, 220,  60,  20, .para2
		TextBox      520, 220,  60,  20, .para3

		PushButton   290, 245,  90,  20, "Generate",.GeneratePB
		PushButton   390, 245,  90,  20, "Close",.ClosePB
		PushButton   490, 245,  90,  20, "Help",.HelpPB
		CancelButton 520,   0,  90,  20 ' needed for red "X" to work

	End Dialog

	Dim dlg As UserDialog

	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean

	Dim dFmin As Double, dFMax As Double
	Dim dGeometryUnit As Double, dFrequencyUnit As Double, dElectricalCurrentUnit As Double

	dGeometryUnit = Units.GetGeometryUnitToSI
	dFrequencyUnit = Units.GetFrequencyUnitToSI
	dElectricalCurrentUnit = Units.GetCurrentUnitToSI

	dFmin = Solver.GetFMin
	dFMax = Solver.GetFMax

	If ((dFmin = 0) And (dFMax = 0)) Then
		MsgBox("Please define a non-zero frequency range first.", "Check solver settings")
		Exit All
	End If

	DlgVisible("Cancel", False) ' hide it, but it needs to be there for the red "X" to work
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgValue("PreloadedFieldsDLB",5)   '  default  is "Isotropic - E-theta lin-pol"
		DlgText("FFSourceNameT", "Isotropic-E-thetalin-pol")
		DlgText("ResolutionThetaT", "5")
		DlgText("ResolutionPhiT", "5")
		DlgText("FminT", CStr((dFMax+dFmin)/2))
		DlgText("FmaxT", CSTr((dFMax+dFmin)/2))
		DlgText("DeltaFT", "0")
		DlgText("EThetaRealT", "sqr(377/(4*Pi))")   '  default  is "Isotropic - E-theta lin-pol"
		DlgText("EThetaImagT", "0")
		DlgText("EPhiRealT", "0")
		DlgText("EPhiImagT", "0")
			DlgVisible "txt1", 0
			DlgVisible "txt2", 0
			DlgVisible "txt3", 0
			DlgVisible "para1", 0
			DlgVisible "para2", 0
			DlgVisible "para3", 0

	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "Cancel"
				DialogFunction=False
				Exit All
			Case "HelpPB"
				DialogFunction = True
				StartHelp HelpFileName
			Case "GeneratePB"
				DialogFunction = True
				DlgEnable("GeneratePB", False)
				DlgEnable("ClosePB", False)
				If (Resulttree.GetFirstChildName("Farfield Sources\"+DlgText("FFSourceNameT"))<>"") Then
					If MsgBox("A source with this name already exists, do you want to overwrite it?", vbYesNo, "Name Check") = vbNo Then
						DlgEnable("GeneratePB", True)
						DlgEnable("ClosePB", True)
						Exit Function
					End If
				End If

				If DlgValue("PreloadedFieldsDLB") = 1 Then
					If ((Not IsNumeric(DlgText("para1"))) Or (Not IsNumeric(DlgText("para2"))) Or (DlgText("para2")<=0) Or (DlgText("para1") = 0)) Then
						MsgBox("Invalid value entered for input parameter","Error")
						'Checks to see that input parameters make logical sense
						Exit All
					End If
					DlgText("EThetaImagT", Replace(Replace(PreloadedFields(1)(2), "I0", Cstr(Evaluate(DlgText("para1"))*dElectricalCurrentUnit)), "LDipole", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)))
				End If

				If DlgValue("PreloadedFieldsDLB") = 2 Then
					If ((Not IsNumeric(DlgText("para1"))) Or (Not IsNumeric(DlgText("para2"))) Or (DlgText("para2")<=0)) Then
						MsgBox("Invalid value entered for input parameter","Error")
						Exit All
					End If
					DlgText("EPhiRealT", Replace(Replace(PreloadedFields(2)(3), "I0", Cstr(Evaluate(DlgText("para1"))*dElectricalCurrentUnit)), "rLoop", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)))
				End If

				If DlgValue("PreloadedFieldsDLB") = 3 Then
					If ((Not IsNumeric(DlgText("para1"))) Or (Not IsNumeric(DlgText("para2"))) Or (Not IsNumeric(DlgText("para3"))) Or (DlgText("para2")<=0) Or (DlgText("para1") <=0) Or (DlgText("para2")<=0)) Then
						MsgBox("Invalid value entered for input parameter","Error")
						Exit All
					End If
					DlgText("EThetaRealT", Replace(Replace(Replace(PreloadedFields(3)(1), "a1", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "B", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)),"Alpha", DlgText("para3")))
					DlgText("EThetaImagT", Replace(Replace(Replace(PreloadedFields(3)(2), "a1", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "B", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)),"Alpha", DlgText("para3")))
					DlgText("EPhiRealT", Replace(Replace(Replace(PreloadedFields(3)(3), "a1", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "B", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)), "Alpha", DlgText("para3")))
					DlgText("EPhiImagT", Replace(Replace(Replace(PreloadedFields(3)(4), "a1", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "B", Cstr(Evaluate(DlgText("para2"))*dGeometryUnit)), "Alpha", DlgText("para3")))
				End If

				If DlgValue("PreloadedFieldsDLB") = 4 Then
					If ((Not IsNumeric(DlgText("para1"))) Or (Not IsNumeric(DlgText("para2")))) Then
						MsgBox("Invalid value entered for input parameter","Error")
						Exit All
					End If
					DlgText("EThetaRealT", Replace(Replace(PreloadedFields(4)(1), "b", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "Nf", DlgText("para2")))
					DlgText("EPhiRealT", Replace(Replace(PreloadedFields(4)(3), "b", Cstr(Evaluate(DlgText("para1"))*dGeometryUnit)), "Nf", DlgText("para2")))
				End If
				GenerateAnalyticalFFSource(DlgText("FFSourceNameT"), _
											DlgText("ResolutionThetaT"), _
											DlgText("ResolutionPhiT"), _
											CStr(Evaluate(DlgText("FminT"))*dFrequencyUnit), _
											CStr(Evaluate(DlgText("FmaxT"))*dFrequencyUnit), _
											CStr(Evaluate(DlgText("DeltaFT"))*dFrequencyUnit), _
											DlgText("EThetaRealT"), _
											DlgText("EThetaImagT"), _
											DlgText("EPhiRealT"), _
											DlgText("EPhiImagT"))
				DlgEnable("GeneratePB", True)
				DlgEnable("ClosePB", True)
			Case "PreloadedFieldsDLB"
				DialogFunction = True
				DlgText("FFSourceNameT", Replace(PreloadedFields(DlgValue("PreloadedFieldsDLB"))(0), " ", ""))
				DlgText("EThetaRealT", PreloadedFields(DlgValue("PreloadedFieldsDLB"))(1))
				DlgText("EThetaImagT", PreloadedFields(DlgValue("PreloadedFieldsDLB"))(2))
				DlgText("EPhiRealT", PreloadedFields(DlgValue("PreloadedFieldsDLB"))(3))
				DlgText("EPhiImagT", PreloadedFields(DlgValue("PreloadedFieldsDLB"))(4))


				Dim n As Integer
				n = DlgValue("PreloadedFieldsDLB")

				Select Case n

					Case 0			'select predefined field
						DlgVisible "txt1", 0
						DlgVisible "txt2", 0
						DlgVisible "txt3", 0
						DlgVisible "para1", 0
						DlgVisible "para2", 0
						DlgVisible "para3", 0

					Case 1			'half-wavelength dipole
						DlgVisible "txt1", 1
						DlgVisible "txt2", 1
						DlgVisible "txt3", 0
						DlgVisible "para1",1
						DlgVisible "para2",1
						DlgVisible "para3",0
						DlgEnable "txt1", 1
						DlgEnable "txt2", 1
						DlgEnable "para1", 1
						DlgEnable "para2", 1
						DlgVisible "para3", 0
						DlgText("txt1", "I0 ["+Units.GetUnit("Current")+"]")
						DlgText("txt2", "LDipole ["+Units.GetUnit("Length")+"]")
						DlgText("para1","1")		'I0 defaults to 1
						DlgText("para2",(cLight/((dFmin+dFMax)*dFrequencyUnit)/dGeometryUnit))	'LDipole defaults to a half wavelength defined by midpoint of Fmin and Fmax

					Case 2			'small circular loop
						DlgVisible "txt1", 1
						DlgVisible "txt2", 1
						DlgVisible "txt3", 0
						DlgVisible "para1",1
						DlgVisible "para2",1
						DlgVisible "para3",0
						DlgEnable "para1", 1
						DlgEnable "para2", 1
						DlgVisible "para3", 0
						DlgText("txt1", "I0 ["+Units.GetUnit("Current")+"]")
						DlgText("txt2", "rLoop ["+Units.GetUnit("Length")+"]")
						DlgText("para1","1")
						DlgText("para2",(cLight/((dFmin+dFMax)*dFrequencyUnit)/dGeometryUnit*2*.03))		'rLoop defaults to a wavelength *.03

					Case 3			'circular horn
						DlgVisible "txt1", 1
						DlgVisible "txt2", 1
						DlgVisible "txt3", 1
						DlgVisible "para1", 1
						DlgVisible "para2", 1
						DlgVisible "para3", 1
						DlgEnable "txt1", 1
						DlgEnable "txt2", 1
						DlgEnable "txt3", 1
						DlgEnable "para1", 1
						DlgEnable "para2", 1
						DlgEnable "para3", 1
						DlgText("txt1", "a1 ["+Units.GetUnit("Length")+"]")
						DlgText("txt2", "B ["+Units.GetUnit("Length")+"]")
						DlgText("txt3", "Alpha [degree]")
						DlgText("para1",(cLight/((dFmin+dFMax)*dFrequencyUnit)/dGeometryUnit*2*4))	'Optimum radius is 2*wavelength for gain of 25dB
						DlgText("para2",(cLight/((dFmin+dFMax)*dFrequencyUnit)/dGeometryUnit*2*20))	'Optimum axial length is 20*wavelength for gain of 25dB
						DlgText("para3","20")	'Defaults to 20 degrees

					Case 4			'gaussian beam
						DlgVisible "txt1", 1
						DlgVisible "txt2", 1
						DlgVisible "txt3", 0
						DlgVisible "para1", 1
						DlgVisible "para2", 1
						DlgVisible "para3", 0
						DlgEnable "txt1", 1
						DlgEnable "txt2", 1
						DlgEnable "para1", 1
						DlgEnable "para2", 1
						DlgText("txt1", "b ["+Units.GetUnit("Length")+"]")
						DlgText("txt2", "Nf:")
						DlgText("para1","1")	'Defaults to 1
						DlgText("para2","1")	'Defaults to 1

					Case 5,6,7,8			'all isotropic sources
						DlgVisible "txt1", 0
						DlgVisible "txt2", 0
						DlgVisible "txt3", 0
						DlgVisible "para1", 0
						DlgVisible "para2", 0
						DlgVisible "para3", 0

					End Select

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select



End Function

Function GenerateAnalyticalFFSource(FFSourceName As String, _
									ThetaResolution As String, _
									PhiResolution As String, _
									Fmin As String, _
									Fmax As String, _
									DeltaF As String, _
									EThetaReal As String, _
									EThetaImag As String, _
									EPhiReal As String, _
									EPhiImag As String) As Integer


	' This function creates a farfield source (ffs) file from analytical field expressions
	' Inputs:
	'	FFSourceName 	As String: 	The name of the farfield source
	'	ThetaResolution As String: 	Angular resolution for theta
	'	PhiResolution 	As String:	Angular resolution for phi
	'	Fmin			As String:	Minimum frequency
	'	Fmax			As String:	Maximum frequency
	'	DeltaF			As String:	Frequency step width
	'	EThetaReal 		As String:	Analytical expression for the real part of the ETheta field component. Angle variables are 'TH' and 'PH'; Frequency variable is 'f'
	'	EThetaImag 		As String:	Analytical expression for the imaginary part of the ETheta field component. Angle variables are 'TH' and 'PH'; Frequency variable is 'f'
	'	EPhiReal 		As String:	Analytical expression for the real part of the EPhi field component. Angle variables are 'TH' and 'PH'; Frequency variable is 'f'
	'	EPhiImag 		As String:	Analytical expression for the imaginary part of the EPhi field component. Angle variables are 'TH' and 'PH'; Frequency variable is 'f'

	Dim dTheta As Double, dPhi As Double, dThetaMin As Double, dThetaMax As Double, dPhiMin As Double, dPhiMax As Double, dThetaStep As Double, dPhiStep As Double
	Dim dFreq As Double, dFmin As Double, dFMax As Double, dDeltaF As Double, nFreqSamples As Long
	Dim nPhiSamples As Integer, nThetaSamples As Integer
	Dim sOutputFileName As String, iOutputFile As Integer, sFarfieldName As String
	Dim i As Long
	Dim sHistoryEntry As String

	sOutputFileName = GetProjectPath("Model3D")+""+FFSourceName+".ffs"
	sHistoryEntry = ""

	Dim dEThetaReal As Double, dEThetaImag As Double, dEPhiReal As Double, dEPhiImag As Double

	' Set up parameters
	dThetaMin = 0
	dThetaMax = 180
	dThetaStep = Evaluate(ThetaResolution)
	dPhiMin = 0
	dPhiMax = 360
	dPhiStep = Evaluate(PhiResolution)
	dFmin = Evaluate(Fmin)

	If (dFmin = 0) Then
		MsgBox("Please enter a minimum frequency larger than zero.", "Error")
		Exit Function
	End If
	dFMax = Evaluate(Fmax)
	dDeltaF = Evaluate(DeltaF)
	If (dDeltaF = 0) Then
		dFMax = dFmin
		dDeltaF = 1
		nFreqSamples = 1
	ElseIf dDeltaF > 0 Then
		nFreqSamples = Fix((dFMax-dFmin)/dDeltaF+1e-14) + 1 ' fsr: Added 1e-15 to avoid some problems with "fix"
	Else
		MsgBox("Please enter a frequency step larger than or equal to zero.", "Error")
		Exit Function
	End If
	nThetaSamples = Fix((dThetaMax-dThetaMin)/dThetaStep) + 1
	nPhiSamples = Fix((dPhiMax-dPhiMin)/dPhiStep) + 1

	EThetaReal = CST_ReplaceString(EThetaReal, "PH")
	EThetaReal = CST_ReplaceString(EThetaReal, "TH")
	EThetaReal = CST_ReplaceString(EThetaReal, "f")
	EThetaImag = CST_ReplaceString(EThetaImag, "PH")
	EThetaImag = CST_ReplaceString(EThetaImag, "TH")
	EThetaImag = CST_ReplaceString(EThetaImag, "f")

	EPhiReal = CST_ReplaceString(EPhiReal, "PH")
	EPhiReal = CST_ReplaceString(EPhiReal, "TH")
	EPhiReal = CST_ReplaceString(EPhiReal, "f")
	EPhiImag = CST_ReplaceString(EPhiImag, "PH")
	EPhiImag = CST_ReplaceString(EPhiImag, "TH")
	EPhiImag = CST_ReplaceString(EPhiImag, "f")

	iOutputFile = FreeFile()

	' Open farfield source file
	Open sOutputFileName For Output As iOutputFile
	Print #iOutputFile, "// CST Farfield Source File"
	Print #iOutputFile, "// Version:"
	Print #iOutputFile, "3.0"
	Print #iOutputFile, "// Data Type"
	Print #iOutputFile, "Farfield"
	Print #iOutputFile, "// Number of Frequency Samples"
	Print #iOutputFile, CStr(nFreqSamples)
	Print #iOutputFile, "// Position"
	Print #iOutputFile, "0 0 0"
	Print #iOutputFile, "// z-Axis"
	Print #iOutputFile, "0 0 1"
	Print #iOutputFile, "// x-Axis"
	Print #iOutputFile, "1 0 0"+vbNewLine
	Print #iOutputFile, "// Radiated Power/Accepted Power/Stimulated Power/Frequency [Hz], one line per entry, one block per frequency sample"
	For dFreq = dFmin To (dFMax+1e-14) STEP dDeltaF ' add 1e-14 to avoid issues with floating point accuracy
		Print #iOutputFile, "-1"
		Print #iOutputFile, "-1"
		Print #iOutputFile, "-1"
		Print #iOutputFile, CSTr(dFreq)+vbNewLine
		If dFreq = dFMax Then Exit For ' Exit point to prevent endless loop for single freq sample
	Next dFreq

	For dFreq = dFmin To (dFMax+1e-14) STEP dDeltaF ' add 1e-14 to avoid issues with floating point accuracy
		Print #iOutputFile, "// Total number of phi And theta samples"
		Print #iOutputFile, CStr(nPhiSamples)+" "+CStr(nThetaSamples)
		Print #iOutputFile, "// phi, theta, Re(Etheta), Im(Etheta), Re(Ephi), Im(Ephi)"

		For dPhi = dPhiMin To (dPhiMax+1e-14) STEP dPhiStep  ' add 1e-14 to avoid issues with floating point accuracy
			For dTheta = dThetaMin To (dThetaMax+1e-14) STEP dThetaStep ' add 1e-14 to avoid issues with floating point accuracy
	     		dEThetaReal = Evaluate(Replace(Replace(Replace(EThetaReal, "f_cst_tmp", "("+CStr(dFreq)+")"), "PH_cst_tmp", "("+CStr(dPhi)+")"), "TH_cst_tmp", "("+CStr(dTheta)+")"))
	     		dEThetaImag = Evaluate(Replace(Replace(Replace(EThetaImag, "f_cst_tmp", "("+CStr(dFreq)+")"), "PH_cst_tmp", "("+CStr(dPhi)+")"), "TH_cst_tmp", "("+CStr(dTheta)+")"))
				dEPhiReal = Evaluate(Replace(Replace(Replace(EPhiReal, "f_cst_tmp", "("+CStr(dFreq)+")"), "PH_cst_tmp", "("+CStr(dPhi)+")"), "TH_cst_tmp", "("+CStr(dTheta)+")"))
				dEPhiImag = Evaluate(Replace(Replace(Replace(EPhiImag, "f_cst_tmp", "("+CStr(dFreq)+")"), "PH_cst_tmp", "("+CStr(dPhi)+")"), "TH_cst_tmp", "("+CStr(dTheta)+")"))
				Print #iOutputFile, CStr(dPhi)+" "+CStr(dTheta)+" "+CStr(dEThetaReal)+" "+CStr(dEThetaImag)+" "+CStr(dEPhiReal)+" "+CStr(dEPhiImag)
			Next dTheta
		Next dPhi
		If dFreq = dFMax Then Exit For ' Exit point to prevent endless loop for single freq sample
	Next dFreq

	Close iOutputFile
	' Shell "notepad " & sOutputFileName, 3

	' Add to project
	sHistoryEntry = sHistoryEntry + "With FARFIELDSOURCE"+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".Reset"+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".Name "+Chr(34)+FFSourceName+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".Id "+Chr(34)+CStr(FARFIELDSOURCE.GetNextID)+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".UseCopyOnly "+Chr(34)+"True"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetPosition "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"0"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetTheta0XYZ "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"1"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetPhi0XYZ "+Chr(34)+"1"+Chr(34)+", "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"0"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".Import "+Chr(34)+sOutputFileName+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".UseMultipoleFFS "+Chr(34)+"True"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetAlignmentType "+Chr(34)+"user"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetMultipoleDegree "+Chr(34)+"1"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".SetMultipoleCalcMode "+Chr(34)+"automatic"+Chr(34)+vbNewLine
	sHistoryEntry = sHistoryEntry + vbTab + ".Store"+vbNewLine
	sHistoryEntry = sHistoryEntry + "End With"
	AddToHistory("define farfield source: " + FFSourceName, sHistoryEntry)

	Resulttree.UpdateTree
	MsgBox("Farfield source successfully created.","Success")

End Function


Function fact(x As Double) As Double

	Dim n As Long

		If x < 0 Then
			MsgBox("Argument must be greater than or equal to 0", vbInformation)
			Exit Function
		ElseIf x = 0 Then
			fact = 1
			Exit Function
		Else
		fact = 1
		For n = 1 To x
			fact = fact*n
		Next n
		End If
End Function

'Fresnel Sine integral
Function S(x As Double) As Double
Dim n As Double
Dim fs As Double
	fs = 0
	For	 n = 0 To 10
		fs = fs + (2^(-2*n-1)*pi^(2*n+1)*((-x^4)^n)/((4*n+3)*fact(2*n+1)))
	Next n
	S = x^3*fs
End Function

'Fresnel Cosine integral
Function C(x As Double) As Double
Dim n As Double
Dim fc As Double
	fc = 0
	For	 n = 0 To 10
		fc = fc + (2^(-2*n)*pi^(2*n)*((-x^4)^n)/((4*n+1)*fact(2*n)))
	Next n
	C = x*fc
End Function

Function Lommel(n As Double, w As Double, z As Double) As Double
	Dim m As Integer, sum As Double
	sum = 0
	For m = 0 To 50
		sum = sum + (-1)^m*(w/z)^(n+z*m)*cyl_bessel_j(n+2*m, z)
	Next m
	Lommel = sum

End Function

