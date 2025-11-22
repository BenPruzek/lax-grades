'#Language "WWB-COM"

'#include "vba_globals_all.lib"
'#include "exports.lib"

'-------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 10-Aug-2023 fsr: Fixed power normalization for 2D cases
' 26-May-2017 ksr,ube: fixed bug with supported unit types and variable i and j overflow in generateFaceFieldData() function
' 07-Mar-2017 fsr: Using buffered file write for improved performance (~4x)
' 26-Oct-2016 fsr: Source size was 0.5*deltaX/Y/Z too small - fixed
' 18-Jan-2016 fsr: Performance improvements; prepared for frequency-dependent beam radius
' 16-Nov-2015 fsr: Automatically define and set up Gaussian sine excitation signal; fixed 2D source for new NFS format
' 22-Sep-2015 teu: Fixed phase reference to ensure causal data
' 16-Jul-2015 jfl: Removed frequency resampling, new method can handle arbitrary samples
' 19-Jun-2015 jfl: changed to NFS format. Fixed transforms when WCS is active
' 30-Oct-2014 fsr: Fixed a bug that led to division by zero when using "Min. beam radius" option
' 25-Jul-2014 yta: added 0 and 90 degree polarization in 2D option
' 27-May-2014 fsr: Added option to enter field normalization in Watts; added 2D option
' 26-Jun-2013 hsn: proper handling of field source id's
' 20-Apr-2013 fsr: Revised limits for w0 and zR: Warning for numerical limit, abort for physical limit;
' 					removed unit check: customers are also using macro for non-optical applications; small GUI improvements and bugfixes
' 01-Aug-2012 fsr: Major overhaul: Orientation and polarization now parameterized, multiple frequencies or wavelengths,
'					fields inside box for scattering analysis
' 27-Oct-2011 fsr: Added option for circular polarization
' 21-Feb-2011 ube: included into online help
' 31-Aug-2010 fsr: save beam parameters in source and info file
' 03-Aug-2010 ube,fsr: function USFormat moved into vba_globals_all.lib
' 25-Jul-2010 ube: commenting out help-button was a bad idea from me...
' 15-Jul-2010 fsr: included fields on walls parallel to propagation direction, adjusted output for non-US locales
' 07-Jul-2010 fsr: more explicit field calculations to avoid numerical issues
' 30-Jun-2010 fsr: added some sanity checks, improved unit handling, added history entry
' 21-Jun-2010 fsr: polarization angle
' 17-Jun-2010 fsr: arbitrary beam direction
' 14-Jun-2010 fsr: initial version with GUI
'-------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit


Const HelpFileName = "common_preloadedmacro_solver_gaussian_beam"

Public k As Double		' propagation constant
Public b As Double		' normalized variable
Public w0 As Double 	' Minimum beam radius at lambdamax
Public w0Lambda As Double ' wavelength dependent w0
Public omega As Double
Public zR As Double		' Rayleigh length

' The following variables are used across multiple functions; declare global to prevent the need to re-calculate them in nested for loops
Public PhiXYZ As Double
Public AlphaXYZ As Double
Public COSDPH As Double
Public SINDPH As Double
Public ExReal As Double
Public ExImag As Double
Public EyReal As Double
Public EyImag As Double
Public EzReal As Double
Public EzImag As Double
Public HxReal As Double
Public HxImag As Double
Public HyReal As Double
Public HyImag As Double
Public HzReal As Double
Public HzImag As Double

Public SFormat As String
Public SxmlFormat As String
Public lUnit As Double
Public lUnitS As String
Public fUnit As Double
Public fUnitS As String

Public Const PolarizationTypes = Array("Linear", "RHEP", "LHEP")

Sub Main


	lUnit = Units.GetGeometryUnitToSI
	lUnitS = Units.GetUnit("Length")
	fUnit = Units.GetFrequencyUnitToSI
	fUnitS = Units.GetUnit("Frequency")

	Dim RayleighOrWidth(1) As String
	RayleighOrWidth(0) = "Rayleigh length:"
	RayleighOrWidth(1) = "Min. beam radius:"
	Dim WavelengthOrFrequency(1) As String
	WavelengthOrFrequency(0) = "Wavelength:"
	WavelengthOrFrequency(1) = "Frequency:"
	Dim PowerOrField(1) As String
	PowerOrField(0) = "Beam Power (W):"
	PowerOrField(1) = "E-field amplitudes (V/m):"

	Begin Dialog UserDialog 660,413,"Generate Gaussian Beam Source",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,630,112,"Beam parameters",.BeamParasGB
		DropListBox 30,28,140,192,WavelengthOrFrequency(),.WavelengthOrFrequencyDLB
		TextBox 180,28,90,21,.FMinT
		Text 280,35,20,14,"...",.Text2
		TextBox 300,28,90,21,.FmaxT
		Text 400,35,30,14,"nm",.FreqUnitsL
		Text 430,35,170,14,", using                sample(s).",.Text1
		TextBox 480,28,50,21,.FreqSamplesT
		Text 30,63,110,14,"Focus distance:",.Text3
		TextBox 180,56,90,21,.z0T
		Text 280,63,40,14,lUnitS,.Text4
		DropListBox 30,84,140,192,RayleighOrWidth(),.RayleighDLB
		TextBox 180,84,90,21,.RayleighLengthT
		Text 280,91,270,14,lUnitS+" (at lambda = lambda_max)",.Text6

		GroupBox 10,126,330,231,"Orientation and polarization",.OrAPoGB
		Text 30,154,180,14,"Propagation vector (x/y/z):",.Text9
		TextBox 210,147,30,21,.propXT
		Text 240,154,10,14,"/",.Text13
		TextBox 250,147,30,21,.propYT
		Text 280,154,10,14,"/",.Text14
		TextBox 290,147,30,21,.propZT

		Text 30,182,140,14,"Polarization type:",.Text5
		DropListBox 210,175,110,192,PolarizationTypes(),.PolarizationTypeDLB
		DropListBox 30,203,180,121,PowerOrField(),.PowerOrFieldDLB
		Text 40,238,30,14,"P1 =",.PorELabel1
		TextBox 80,231,60,21,.E1AmpT
		Text 170,238,30,14,"P2 =",.PorELabel2
		TextBox 210,231,60,21,.E2AmpT
		Text 40,266,170,14,"E1-field polarization angle:",.Text8
		TextBox 210,259,60,21,.polAlphaT
		Text 280,266,30,14,"deg.",.Text12
		CheckBox 50,301,130,14,"2D source width:",.TwoDSourceCB
		TextBox 210,294,60,21,.TwoDSliceWidthT
		Text 280,301,30,14,lUnitS,.Text17
		Text 30,329,150,14,"E1-field orientation - 2D",.Text16
		OptionGroup .EpolarOB
			OptionButton 190,329,60,14,"0 deg.",.d0
			OptionButton 260,329,70,14,"90 deg.",.d90

		GroupBox 350,126,290,231,"Miscellaneous",.MiscGB
		Text 370,154,150,14,"Lines per wavelength:",.Text11
		TextBox 550,147,60,21,.SamplesT
		CheckBox 370,182,250,14,"Excite fields inside source box only",.InsideFieldsCB
		Text 390,210,170,14,"Source box size (x/y/z) in",.Text7
		Text 560,210,30,14,lUnitS+":",.BoxSizeUnitLabel
		Text 470,238,10,14,"/",.Text19
		Text 540,238,10,14,"/",.Text20
		TextBox 410,231,60,21,.BoxSizeXT
		TextBox 480,231,60,21,.BoxSizeYT
		TextBox 550,231,60,21,.BoxSizeZT
		Text 390,266,110,14,"Truncation error:",.Text10
		TextBox 550,259,60,21,.truncErrT
		Text 370,294,120,14,"Add field monitors:",.Text15
		CheckBox 490,294,30,14,"E",.EMonitorsCB
		CheckBox 530,294,40,14,"H",.HMonitorsCB
		CheckBox 570,294,40,14,"FF",.FFMonitorsCB

		Text 20,364,620,14,"",.outputT
		CheckBox 30,378,60,14,"Abort",.AbortCB

		OKButton 350,385,90,21
		PushButton 450,385,90,21,"Exit",.ExitPB
		CancelButton 250,385,90,21
		PushButton 550,385,90,21,"Help",.HelpB


	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg)=0) Then ' user pressed cancel
		Exit All
	End If


End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim E1Amp As Double, E2Amp As Double, PhaseShiftDeg As Double
	Dim dFMin As Double, dFMax As Double, FreqSamples As Long, dDeltaFreq As Double, iDeltaFreqModulo As Long
	Dim dLambdaMin As Double, dLambdaMax As Double
	Dim sTmpString As String, dTmpDouble As Double, sFrequencyOrWavelength As String

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgVisible("Cancel", False) ' only included to enable X in top right corner
		DlgText("FMinT", "633")
		DlgText("FmaxT", "633")
		DlgText("FreqSamplesT", "1")
		DlgText("z0T", Format(2*633/lUnit*1e-9,"Scientific"))
		DlgText("RayleighLengthT", Format(4*633/Pi/lUnit*1e-9,"Scientific"))
		DlgText("propXT", "0")
		DlgText("propYT", "0")
		DlgText("propZT", "1")
		DlgText("polAlphaT", "0")
		DlgText("E1AmpT", "1")
		DlgEnable("E2AmpT", False)
		DlgText("E2AmpT", "0")
		DlgText("truncErrT", "1e-5")
		DlgEnable("truncErrT", Not CBool(DlgValue("InsideFieldsCB")))
		DlgEnable("BoxSizeXT", CBool(DlgValue("InsideFieldsCB")))
		DlgEnable("BoxSizeYT",  CBool(DlgValue("InsideFieldsCB")) And Not CBool(DlgValue("TwoDSourceCB")))
		DlgEnable("BoxSizeZT", CBool(DlgValue("InsideFieldsCB")))
		DlgText("BoxSizeXT", "1500")
		DlgText("BoxSizeYT", "1500")
		DlgText("BoxSizeZT", "1500")
		DlgText("SamplesT", "10")
		DlgValue("EMonitorsCB", True)
		DlgEnable("TwoDSliceWidthT", False)
		DlgText("TwoDSliceWidthT", "1")
		DlgValue("EpolarOB", 0)
		DlgEnable("EpolarOB", False)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "HelpB"
				StartHelp HelpFileName
				DialogFunc = True
			Case "ExitPB"
 				Exit All
			Case "PowerOrFieldDLB"
				If DlgValue("PowerOrFieldDLB") = 0 Then
					DlgText("PorELabel1", "P1 =")
					DlgText("PorELabel2", "P2 =")
				Else
					DlgText("PorELabel1", "E1 =")
					DlgText("PorELabel2", "E2 =")
				End If
			Case "PolarizationTypeDLB"
				If DlgText("PolarizationTypeDLB")<>"Linear" Then
					DlgEnable("E2AmpT", True)
					DlgText("E2AmpT", DlgText("E1AmpT"))
				Else
					DlgEnable("E2AmpT", False)
				End If
			Case "WavelengthOrFrequencyDLB"
				sTmpString = DlgText("FminT") ' needed for swapping data between FminT and FmaxT
				If (DlgText("WavelengthOrFrequencyDLB") = "Wavelength:") Then
					DlgText("FreqUnitsL", "nm")
					DlgText("FminT", Format(CLight/(Evaluate(DlgText("FmaxT"))*fUnit)/1e-9, "Scientific"))
					DlgText("FmaxT", Format(CLight/(Evaluate(sTmpString)*fUnit)/1e-9, "Scientific"))
				ElseIf (DlgText("WavelengthOrFrequencyDLB") = "Frequency:") Then
					DlgText("FreqUnitsL", fUnitS)
					DlgText("FminT", Format(CLight/(Evaluate(DlgText("FmaxT"))*1e-9)/fUnit, "Scientific"))
					DlgText("FmaxT", Format(CLight/(Evaluate(sTmpString)*1e-9)/fUnit, "Scientific"))
				Else
					ReportError("Wavelength or Frequency DLB: Unknown option.") ' This should not happen
				End If
			Case "RayleighDLB"
				dTmpDouble = IIf(DlgText("WavelengthOrFrequencyDLB")="Wavelength:", Evaluate(DlgText("FmaxT"))*1e-9, CLight/(Evaluate(DlgText("FminT"))*fUnit)) ' lambda_max in m
				If (DlgText("RayleighDLB") = "Rayleigh length:") Then
					DlgText("RayleighLengthT", Format(Evaluate(DlgText("RayleighLengthT"))^2*lUnit*Pi/dTmpDouble, "Scientific"))
				ElseIf (DlgText("RayleighDLB") = "Min. beam radius:") Then
					DlgText("RayleighLengthT", Format(Sqr(Evaluate(DlgText("RayleighLengthT"))/lUnit/Pi*dTmpDouble), "Scientific"))
				Else
					ReportError("Wavelength or Frequency DLB: Unknown option.") ' This should not happen
				End If
			Case "InsideFieldsCB"
				DlgEnable("truncErrT", Not CBool(DlgValue("InsideFieldsCB")))
				DlgEnable("BoxSizeXT", CBool(DlgValue("InsideFieldsCB")))
				DlgEnable("BoxSizeYT", CBool(DlgValue("InsideFieldsCB")) And Not CBool(DlgValue("TwoDSourceCB")))
				DlgEnable("BoxSizeZT", CBool(DlgValue("InsideFieldsCB")))
			Case "TwoDSourceCB"
				DlgEnable("TwoDSliceWidthT", CBool(DlgValue("TwoDSourceCB")))
				DlgEnable("BoxSizeYT", CBool(DlgValue("InsideFieldsCB")) And Not CBool(DlgValue("TwoDSourceCB")))
				DlgText("polAlphaT", "0")
				DlgEnable("polAlphaT", Not CBool (DlgValue("TwoDSourceCB")))
				DlgEnable("EpolarOB", CBool(DlgValue("TwoDSourceCB")))
			Case "EpolarOB"
				DlgText("polAlphaT", IIf(DlgValue("EpolarOB")=0, "0", "90"))
			Case "OK"
				DlgEnable "OK", False
				DlgEnable "ExitPB", False
				DlgEnable "HelpB", False
				Select Case DlgText("WavelengthOrFrequencyDLB")
					Case "Wavelength:"
						sFrequencyOrWavelength = "Wavelength"
						dFMin = CLight/Evaluate(DlgText("FMaxT"))/1e-9
						dFMax = CLight/Evaluate(DlgText("FminT"))/1e-9
					Case "Frequency:"
						sFrequencyOrWavelength = "Frequency"
						dFMin = Evaluate(DlgText("FMinT"))*fUnit
						dFMax = Evaluate(DlgText("FmaxT"))*fUnit
				End Select
				FreqSamples = Evaluate(DlgText("FreqSamplesT"))
				If (FreqSamples <= 0) Then
					MsgBox("At least 1 frequency sample is required.", "Check sample setting")
					DialogFunc = True
					DlgEnable "OK", True
					DlgEnable "ExitPB", True
					DlgEnable "HelpB", True
					Exit Function
				ElseIf ((FreqSamples = 1 And Abs(Evaluate(DlgText("FmaxT"))/Evaluate(DlgText("FminT"))-1)>1e-12) _
						Or (FreqSamples > 1 And Abs(Evaluate(DlgText("FmaxT"))/Evaluate(DlgText("FminT"))-1)<1e-12)) Then
						MsgBox("Fmin and Fmax settings not compatible with number of samples.", "Check sample setting")
						DialogFunc = True
						DlgEnable "OK", True
						DlgEnable "ExitPB", True
						DlgEnable "HelpB", True
						Exit Function
				End If

				dLambdaMin = CLight/dFMax
				dLambdaMax = CLight/dFMin

				If (DlgValue("RayleighDLB")=0) Then ' Rayleigh length
					zR = Evaluate(DlgText("RayleighLengthT"))*lUnit
				ElseIf (DlgValue("RayleighDLB")=1) Then
					zR = Pi*(Evaluate(DlgText("RayleighLengthT"))*lUnit)^2/dLambdaMax
				End If

				If Solver.GetFMax = 0 Then AddToHistory("define frequency range", "Solver.FrequencyRange "+Chr(34)+"0"+Chr(34)+","+Chr(34)+CStr(Fix(1.1*dFMax/fUnit))+Chr(34))
				If (DlgValue("RayleighDLB")=0) Then ' Rayleigh length
					w0 = Sqr(dLambdaMax/Pi*Evaluate(DlgText("RayleighLengthT"))*lUnit) ' calculate min beam radius for the case lambda=lambda_max
				ElseIf (DlgValue("RayleighDLB")=1) Then ' Min beam radius
					w0 = Evaluate(DlgText("RayleighLengthT"))*lUnit
				End If

				If (w0<=2*dLambdaMax/Pi^2) Then
					MsgBox("Rayleigh length z_R (min. beam width w_0) is too small. Please ensure that z_R > 4*lambda_max/pi^3 (or w_0 > 2*lambda_max/pi^2).","Parameter check")
					DialogFunc = True
					DlgEnable "OK", True
					DlgEnable "ExitPB", True
					DlgEnable "HelpB", True
				Else
					If (w0<2*dLambdaMax/Pi) Then ' Siegmann: Divergence < 30 degrees, or Pi/6
						If (MsgBox("Rayleigh length z_R (min. beam width w_0) is very small and results might be inaccurate."+ vbNewLine _
									+ "The paraxial approximation used to calculate the beam assumes the following:" + vbNewLine _
									+ "z_R > 4*lambda_max/pi, w_0 > 2*lambda_max/pi" + vbNewLine + vbNewLine  _
									+ "Do you wish to continue?", vbYesNo, "Parameter check") = vbNo) Then
							DialogFunc = True
							DlgEnable "OK", True
							DlgEnable "ExitPB", True
							DlgEnable "HelpB", True
							Exit Function
						End If
					End If

					E1Amp = Evaluate(DlgText("E1AmpT"))

					Select Case DlgText("PolarizationTypeDLB")
						Case "Linear"
							E2Amp = 0
							PhaseShiftDeg = 0
						Case "RHEP"
							E2Amp = Evaluate(DlgText("E2AmpT"))
							PhaseShiftDeg = +90
						Case "LHEP"
							E2Amp = Evaluate(DlgText("E2AmpT"))
							PhaseShiftDeg = -90
					End Select
					Dim dStartTime As Double
					dStartTime = Timer
					If (CreateSourceFile(0, 0, 0, _
											Evaluate(DlgText("propXT")), Evaluate(DlgText("propYT")), Evaluate(DlgText("propZT")), _
											Evaluate(DlgText("z0T"))*lUnit, Evaluate(DlgText("polAlphaT")), _
											Evaluate(DlgText("truncErrT")), Evaluate(DlgText("SamplesT")), _
											Evaluate(DlgText("BoxSizeXT")), Evaluate(DlgText("BoxSizeYT")), Evaluate(DlgText("BoxSizeZT")), _
											E1Amp, E2Amp, CBool(DlgValue("PowerOrFieldDLB")=0), PhaseShiftDeg, _
											dFMin, dFMax, FreqSamples, _
											CBool(DlgValue("InsideFieldsCB")), _
											CBool(DlgValue("TwoDSourceCB")), Evaluate(DlgText("TwoDSliceWidthT")), DlgValue("EpolarOB")) = 0) Then
						 ' All went well, close Dialog
						DialogFunc = False
						AddFieldMonitors(dFMin, dFMax, FreqSamples, sFrequencyOrWavelength, _
											CBool(DlgValue("EMonitorsCB")), CBool(DlgValue("HMonitorsCB")), CBool(DlgValue("FFMonitorsCB")))
					Else
						DialogFunc = True
						DlgEnable "OK", True
						DlgEnable "ExitPB", True
						DlgEnable "HelpB", True
					End If
					'ReportInformationToWindow(Timer-dStartTime)
				End If

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function CreateSourceFile(Ox As Double, Oy As Double, Oz As Double, _
					 		propX As Double, propY As Double, propZ As Double, _
							focusDistance As Double, polAlpha As Double, _
							truncErr As Double, NSamples As Integer, _
							BoxSizeX As Double, BoxSizeY As Double, BoxSizeZ As Double, _
							dInputAmp1 As Double, dInputAmp2 As Double, bAmplitudesRepresentPower As Boolean, PhaseShiftDeg As Double, _
							fMin As Double, fMax As Double, FreqSamples As Long, _
							bFieldsInSource As Boolean, b2DSource As Boolean, d2DSliceWidth As Double, EpolarOB As Integer) As Integer

	Dim fileName As String
	Dim fsName As String
	Dim streamNum As Long
	Dim historyString As String
	Dim sFreqFormat As String
	Dim dFrequency As Double, dLambda As Double
	Dim dApparentWL As Double, dTruncRadius As Double

	Dim Amp1 As Double, Amp2 As Double
	Dim HHxRe As Double, HHxIm As Double, HHyRe As Double, HHyIm As Double, HHzRe As Double, HHzIm As Double
	Dim EExRe As Double, EExIm As Double, EEyRe As Double, EEyIm As Double, EEzRe As Double, EEzIm As Double

	' Beam direction and angle
	Dim dirBeam(2) As Double	' normalized beam propagation vector
	Dim dirBeamAbs As Double	' length of normal vector
	' Box size per dimension
	Dim Lx As Double, Ly As Double, Lz As Double
	' Number of samples per dimension
	Dim Nx As Double, Ny As Double, Nz As Double

	Dim i As Long, j As Long, iFreqSample As Long
	Dim x As Double, y As Double, z As Double		' global coordinates

	historyString = ""

	CreateSourceFile = -1
	SFormat = " 0.0000000000000000E+00;-0.0000000000000000E+00"
	SxmlFormat = "0.0000000000000000E+00;-0.0000000000000000E+00"
	sFreqFormat = "0.00000000"

	dirBeamAbs = Sqr(propX^2+propY^2+propZ^2)
	If (dirBeamAbs=0) Then
		MsgBox("Please enter a non-zero propagation vector.")
		Exit Function
	End If
	' Store normalized propagation vector
	dirBeam(0)=propX/dirBeamAbs
	dirBeam(1)=propY/dirBeamAbs
	dirBeam(2)=propZ/dirBeamAbs
	dirBeamAbs = 1

	' box size width = 2*distance at which field has decayed to truncErr
	' plus some scaling according to beam direction
	' First calculation, assuming no box
	k = -2*Pi/CLight*fMin
	b = w0^2*k/2 ' b = zR = Rayleigh range
	dTruncRadius = Sqr(-2*(focusDistance^2+b^2)/(k*b)*Log(truncErr))
	If Not bFieldsInSource Then
		If b2DSource Then
			If EpolarOB = 0 Then ' 0 degree case
				Lx = 2*dTruncRadius
				Ly = d2DSliceWidth*lUnit
				Lz = 0
			ElseIf EpolarOB = 1 Then '90 degree case
				Ly = 2*dTruncRadius
				Lx = d2DSliceWidth*lUnit
				Lz = 0
			End If
		Else
			Ly = 2*dTruncRadius
			Lx =  Ly
			Lz =  0
		End If
	Else
		If b2DSource Then
			If EpolarOB = 0 Then
				Lx = d2DSliceWidth*lUnit
				Ly = BoxSizeX*lUnit
				Lz = BoxSizeZ*lUnit
			ElseIf EpolarOB = 1 Then
				Ly = d2DSliceWidth*lUnit
				Lx = BoxSizeY*lUnit
				Lz = BoxSizeZ*lUnit
			End If
		Else
			Ly = BoxSizeX*lUnit
			Lx = Ly
			Lz = BoxSizeZ*lUnit
		End If
	End If

	' wz = w0*Sqr(1+(Lz/2/zR)^2)	'spot size(beam radius) at Rayleigh range

	' number of samples; should be odd
	Nx = 3+Round(Lx/CLight*fMax*NSamples)
	If Nx Mod 2 = 0 Then Nx = Nx +1
	Ny = 3+Round(Ly/CLight*fMax*NSamples)
	If Ny Mod 2 = 0 Then Ny = Ny +1
	Nz = 3+Round(Lz/CLight*fMax*NSamples)
	If Nz Mod 2 = 0 Then Nz = Nz +1

	If (Nx*Ny*Nz > 100000) Then
		If (MsgBox("The number of samples is very high, do you want to continue?", vbYesNo, "Check Settings") = vbNo) Then
			Exit Function
		End If
	End If

	streamNum = FreeFile
	Open GetProjectPath("Model3D")+"GaussianBeamInfo.txt" For Output As #streamNum
		Print #streamNum, "Gaussian beam parameters:"
		Print #streamNum, "Wavelength: "+CStr(CLight/fMax/1e-9) + " ... " +CStr(CLight/fMin/1e-9) +" nm in " + DlgText("FreqSamplesT") + " samples."
		Print #streamNum, "Propagation vector: ("+cstr(propX)+"/"+cstr(propY)+"/"+cstr(propZ)+")"
		Print #streamNum, "Beam waist at lambda_max: "+USFormat(w0,"Scientific")+" m"
'		Print #streamNum, "Rayleigh length: "+USFormat(w0^2*Pi/lambda,"Scientific")+" m"
		Print #streamNum, "Source center: ("+USFormat(Ox,"Scientific")+"/"+USFormat(Oy,"Scientific")+"/"+USFormat(Oz,"Scientific")+") m"
		Print #streamNum, "Focus point distance from source center: "+USFormat(focusDistance,"Scientific")+" m"
		If PhaseShiftDeg = 0 Then
			Print #streamNum, "Polarization: Linear"
			If bAmplitudesRepresentPower Then
				Print #streamNum, "Beam power: P1=" + USFormat(dInputAmp1, "Scientific") + " W"
			Else
				Print #streamNum, "Field amplitude: E1=" + USFormat(dInputAmp1, "Scientific") + " V/m"
			End If
		Else
			Print #streamNum, "Polarization: " + IIf(PhaseShiftDeg = +90, "RHEP", "LHEP")
			If bAmplitudesRepresentPower Then
				Print #streamNum, "Beam powers: P1=" + USFormat(dInputAmp1, "Scientific") + " W, P2=" + USFormat(dInputAmp2, "Scientific") + " W"
			Else
				Print #streamNum, "Field amplitudes: E1=" + USFormat(dInputAmp1, "Scientific") + " V/m, E2=" + USFormat(dInputAmp2, "Scientific") + " V/m"
			End If
		End If
		Print #streamNum, "E1 Polarization angle: "+cstr(polAlpha)+" degrees"
		Print #streamNum, "Truncation error: "+USFormat(truncErr,"Scientific")
		Print #streamNum, "Resolution: "+cstr(NSamples)+" lines per wavelength"
	Close #streamNum

	DlgText "outputT", "Generating file, Nx="+cstr(Nx)+", Ny="+cstr(Ny)+", Nz="+cstr(Nz)

	'jfl 6/18/2015: export now done in nfs format. The file output has been generalized to a single function

	Dim x0 As Double, y0 As Double, z0 As Double, x1 As Double ,y1 As Double ,z1 As Double

	dLambda = -2*Pi/k

	x0 = -Lx/2+Ox*lUnit
	y0 = -Ly/2+Oy*lUnit
	z0 = -Lz/2+Oz*lUnit
	x1 = Lx/2+Ox*lUnit
	y1 = Ly/2+Oy*lUnit
	z1 = Lz/2+Oz*lUnit

	If (bFieldsInSource) Then
		If (generateFaceFieldData(x0, y0, z0, x0, y1, z1, _
						1, Ny, Nz, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Min", b2DSource, EpolarOB) <> 0) Then Exit Function
		If (generateFaceFieldData(x1, y0, z0, x1, y1, z1, _
						1, Ny, Nz, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Max", b2DSource, EpolarOB) <> 0) Then Exit Function
		If (generateFaceFieldData(x0, y0, z0, x1, y0, z1, _
						Nx, 1, Nz, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Min", b2DSource, EpolarOB) <> 0) Then Exit Function
		If (generateFaceFieldData(x0, y1, z0, x1, y1, z1, _
						Nx, 1, Nz, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Max", b2DSource, EpolarOB) <> 0) Then Exit Function
		If (generateFaceFieldData(x0, y0, z0, x1, y1, z0, _
						Nx, Ny, 1, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Min", b2DSource, EpolarOB) <> 0) Then Exit Function
		If (generateFaceFieldData(x0, y0, z1, x1, y1, z1, _
						Nx, Ny, 1, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Max", b2DSource, EpolarOB) <> 0) Then Exit Function
	Else
		If (generateFaceFieldData(x0, y0, 0, x1, y1, 0, _
						Nx, Ny, 1, fMin, fMax, FreqSamples, bAmplitudesRepresentPower, dInputAmp1, _
						dInputAmp2, PhaseShiftDeg, focusDistance, "Max", b2DSource, EpolarOB) <> 0) Then Exit Function
	End If

'=====================================================================================================================================================

	' Set names
	Dim nFiles As Integer
	Dim wcsActive As Boolean
	Dim fsFileName As String
	Dim exportPath As String
	Dim fsID As String

	nFiles = IIf(bFieldsInSource, 24, 4)
	ReDim fsFileNames(0 To nFiles-1) As String
    fsName="fsGBMacro_"+USFormat(dLambda*1e9,"000.000")+"nm"
	exportPath = GetProjectPath("Root")+"\GBMacro_Export^"+Split(GetProjectPath("Project"),"\")(UBound(Split(GetProjectPath("Project"),"\")))+"\"

    wcsActive = WCS.IsWCSActive() = "local"
	If Len(Dir(exportPath, vbDirectory)) = 0 Then MkDir exportPath

	' Copy file into root project directory and import field source
	For i=0 To nFiles-1
		Dim sFace As String, d1 As String, d2 As String
		If((i\8) Mod 3=0) Then
			sFace = "z"
			d1 = "x"
			d2 = "y"
		ElseIf((i\8) Mod 3=1) Then
			sFace = "y"
			d1 = "x"
			d2 = "z"
		Else
			sFace = "x"
			d1 = "y"
			d2 = "z"
		End If
		fsFileNames(i) = IIf(i Mod 2=0,"E","H")+IIf((i\2) Mod 2=0,d1,d2)+"_"+sFace+IIf((i\4) Mod 2=0,"Max","Min")
		On Error GoTo IOError
		FileCopy(GetProjectPath("TempDS")+"GB_"+fsFileNames(i)+".dat",exportPath+"GB_"+fsFileNames(i)+".dat")
		FileCopy(GetProjectPath("TempDS")+"GBMacro_"+fsFileNames(i)+".xml",exportPath+"GBMacro_"+fsFileNames(i)+".xml")
	Next i
GoTo IOSuccess
IOError:
	MsgBox("Could not export files. This can sometimes happen when the project is stored in a temp directory. Please save the project and try again.")
	Exit All
IOSuccess:
	' Get ID
	With FieldSource
		.Reset
		.FileName "GBMacro_"+fsFileNames(0)+".xml"
		fsID = .GetNextId
		.Reset
	End With

	' Disable WCS
	If (wcsActive) Then historyString = historyString + "WCS.ActivateWCS("+Chr(34)+"global"+Chr(34)+")"+vbNewLine

    ' Construction of field source import history command
	historyString = historyString + "With FieldSource"+vbNewLine
	historyString = historyString + " .Delete "+Chr(34)+fsName+Chr(34)+vbNewLine
    historyString = historyString + " .Reset"+vbNewLine
    historyString = historyString + " .Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	For i=0 To nFiles-1
    	historyString = historyString + " .FileName "+Chr(34)+exportPath+"GBMacro_"+fsFileNames(i)+".xml"+Chr(34)+vbNewLine
    Next i
    For i=0 To nFiles-1
    	historyString = historyString + " .Id "+Chr(34)+cstr(cint(fsID)+i)+Chr(34)+vbNewLine
    Next i
    historyString = historyString + " .Read"+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Shell "notepad " & fileName, 3

	StoreParameterWithDescription ("alpha", DlgText("polAlphaT"), "Polarization angle (not for '2D' option)")
	StoreParameterWithDescription ("theta", ACosD(dirBeam(2)), "Beam angle Theta")
	StoreParameterWithDescription ("phi", Atn2D(dirBeam(1), dirBeam(0)), "Beam angle Phi")
	StoreParameterWithDescription ("BeamSourceOriginX", "0", "X-Coordinate of source center")
	StoreParameterWithDescription ("BeamSourceOriginY", "0", "Y-Coordinate of source center")
	StoreParameterWithDescription ("BeamSourceOriginZ", "0", "Z-Coordinate of source center")

	' Rotate source about z axis for polarization
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Origin "+Chr(34)+"Free"+Chr(34)+vbNewLine
	historyString = historyString + "	.Center "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+vbNewLine
	historyString = historyString + "	.Angle "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"alpha"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Rotate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Rotate source to adjust propagation vector
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Origin "+Chr(34)+"Free"+Chr(34)+vbNewLine
	historyString = historyString + "	.Center "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+vbNewLine
	historyString = historyString + "	.Angle "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"theta"+Chr(34)+","+Chr(34)+"0"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Rotate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Origin "+Chr(34)+"Free"+Chr(34)+vbNewLine
	historyString = historyString + "	.Center "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+vbNewLine
	historyString = historyString + "	.Angle "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"phi"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Rotate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Rotate around final beam axis again to restore polarization
	historyString = historyString + "With WCS" + vbNewLine
	historyString = historyString + "	.Store " + Chr(34) + "GB_WCS" + Chr(34) + vbNewLine
	historyString = historyString + "	.ActivateWCS " + Chr(34) + "local" + Chr(34) + vbNewLine
	historyString = historyString + "	.SetNormal "+Chr(34)+"SinD(theta)*CosD(phi)"+Chr(34)+","+Chr(34)+"SinD(theta)*SinD(phi)"+Chr(34)+","+Chr(34)+"CosD(theta)"+Chr(34)+vbNewLine
	historyString = historyString + "End With" + vbNewLine
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Origin "+Chr(34)+"Free"+Chr(34)+vbNewLine
	historyString = historyString + "	.Center "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+vbNewLine
	historyString = historyString + "	.Angle "+Chr(34)+"0"+Chr(34)+","+Chr(34)+"0"+Chr(34)+","+Chr(34)+"-phi"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Rotate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine
	historyString = historyString + "WCS.ActivateWCS " + Chr(34) + "global" + Chr(34) + vbNewLine

	' Apply translation of origin (0 by default)
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Vector "+Chr(34)+"BeamSourceOriginX"+Chr(34)+","+Chr(34)+"BeamSourceOriginY"+Chr(34)+","+Chr(34)+"BeamSourceOriginZ"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Translate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Reenable WCS
	historyString = historyString + "With WCS" + vbNewLine
	historyString = historyString + "	.Restore(" + Chr(34) + "GB_WCS" + Chr(34) + ")" + vbNewLine
	If (wcsActive) Then
		historyString = historyString + "	.ActivateWCS(" + Chr(34) + "local" + Chr(34) + ")" + vbNewLine
	Else
		historyString = historyString + "	.ActivateWCS(" + Chr(34) + "global" + Chr(34) + ")" + vbNewLine
	End If
	historyString = historyString + "End With" + vbNewLine

	If (FreqSamples > 1) Then

		' Define Gaussian sine signal, set as reference
		historyString = historyString + "With TimeSignal" + vbNewLine
	    historyString = historyString + " 	.Reset" + vbNewLine
	    historyString = historyString + " 	.Name " + Chr(34) + "GB Gaussian Sine" + Chr(34) + vbNewLine
	    historyString = historyString + " 	.SignalType " + Chr(34) + "Gaussian sine" + Chr(34) + vbNewLine
	    historyString = historyString + " 	.ProblemType " + Chr(34) + "High Frequency" + Chr(34) + vbNewLine
	    historyString = historyString + " 	.Fmin " + Chr(34) + CStr(fMin/fUnit) + Chr(34) + vbNewLine
	    historyString = historyString + " 	.Fmax " + Chr(34) + CStr(fMax/fUnit) + Chr(34) + vbNewLine
	    historyString = historyString + " 	.Create" + vbNewLine
		historyString = historyString + " 	.ExcitationSignalAsReference " + Chr(34) + "GB Gaussian Sine" + Chr(34) + ",  " + Chr(34) + "High Frequency" + Chr(34) + vbNewLine
		historyString = historyString + "End With" + vbNewLine

		' Assign Gaussian sine as excitation, normalize to reference
		historyString = historyString + "With Solver" + vbNewLine
		historyString = historyString + "	.NormalizeToReferenceSignal " + Chr(34) + "True" + Chr(34) + vbNewLine
	    historyString = historyString + " 	.ResetExcitationModes" + vbNewLine
	    historyString = historyString + " 	.ExcitationFieldSource " + Chr(34) + fsName + Chr(34) + ", " + Chr(34) + "1.0" + Chr(34) + ", " + Chr(34) + "0.0" + Chr(34) + ", " + Chr(34) + "GB Gaussian Sine" + Chr(34) + ", " + Chr(34) + "True" + Chr(34) + vbNewLine
		historyString = historyString + "End With" + vbNewLine

	End If

	AddToHistory("Define Gaussian Beam (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", historyString)

	CreateSourceFile = 0

End Function

Function AddFieldMonitors(dFMin As Double, dFMax As Double, iFreqSamples As Long, sFrequencyOrWavelength As String, _
							bAddEMonitors As Boolean, bAddHMonitors As Boolean, bAddFFMonitors As Boolean) As Integer

	Dim sHistoryString As String
	Dim iFreqSample As Long, dFrequency As Double, dLambda As Double

	' Define e and h field monitors at proper frequency
	If bAddEMonitors Then
		For iFreqSample = 0 To iFreqSamples-1

			If iFreqSamples > 1 Then
				dFrequency = dFMax - iFreqSample*(dFMax-dFMin)/(iFreqSamples-1)
			Else
				dFrequency = dFMin
			End If
			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"e-field (lambda="+cstr(dLambda*1e9)+" nm)"+Chr(34)+vbNewLine
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"e-field (f="+CStr(CLight/dLambda/fUnit)+")"+Chr(34)+vbNewLine
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If
			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Efield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				AddToHistory("define monitor: e-field (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				AddToHistory("define monitor: e-field (f="+CStr(CLight/dLambda/fUnit)+")", sHistoryString)
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If
		Next iFreqSample
	End If

	If bAddHMonitors Then
		For iFreqSample = 0 To iFreqSamples-1

			If iFreqSamples > 1 Then
				dFrequency = dFMax - iFreqSample*(dFMax-dFMin)/(iFreqSamples-1)
			Else
				dFrequency = dFMin
			End If
			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"h-field (lambda="+cstr(dLambda*1e9)+" nm)"+Chr(34)+vbNewLine
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"h-field (f="+CStr(CLight/dLambda/fUnit)+")"+Chr(34)+vbNewLine
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If

			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Hfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				AddToHistory("define monitor: h-field (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				AddToHistory("define monitor: h-field (f="+CStr(CLight/dLambda/fUnit)+")", sHistoryString)
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If
		Next iFreqSample
	End If

	If bAddFFMonitors Then
		sHistoryString = ""
		If iFreqSamples > 1 Then
			dFrequency = dFMax - iFreqSample*(dFMax-dFMin)/(iFreqSamples-1)
		Else
			dFrequency = dFMin
		End If

		If (iFreqSamples = 1) Then
			dLambda = CLight/dFrequency
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (lambda="+cstr(dLambda*1e9)+" nm)"+Chr(34)+vbNewLine
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (f="+CStr(CLight/dLambda/fUnit)+")"+Chr(34)+vbNewLine
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			If (sFrequencyOrWavelength = "Wavelength") Then
				AddToHistory("define farfield monitor: farfield (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
			ElseIf (sFrequencyOrWavelength = "Frequency") Then
				AddToHistory("define farfield monitor: farfield (f="+CStr(CLight/dLambda/fUnit)+")", sHistoryString)
			Else ' This should not happen
				ReportError("AddFieldMonitors: Unknown option")
			End If
		ElseIf (Solver.GetFMax > 0) Then
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (broadband)"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Time"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Accuracy "+Chr(34)+"1e-3"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FrequencySamples "+Chr(34)+CStr(2*iFreqSamples+1)+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr((dFMax+dFMin)/2/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .TransientFarfield "+Chr(34)+"False"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define farfield monitor: farfield (broadband)", sHistoryString)
		Else
			ReportWarningToWindow("Could not create broadband farfield monitor due to Fmax = 0, please increase Fmax.")
		End If
	End If

End Function

Function ExRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	ExRe = AlphaXYZ*(z*Cos(PhiXYZ)-b*Sin(PhiXYZ))
End Function

Function ExIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	ExIm = AlphaXYZ*(z*Sin(PhiXYZ)+b*Cos(PhiXYZ))
End Function

Function EyRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	EyRe = 0
End Function

Function EyIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	EyIm = 0
End Function

Function EzRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	EzRe = -AlphaXYZ*x/(z^2+b^2)*((z^2-b^2)*Cos(PhiXYZ)-2*z*b*Sin(PhiXYZ))
End Function

Function EzIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	EzIm = -AlphaXYZ*x/(z^2+b^2)*((z^2-b^2)*Sin(PhiXYZ)+2*z*b*Cos(PhiXYZ))
End Function

Function HxRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HxRe = 0
End Function

Function HxIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HxIm = 0
End Function

Function HyRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HyRe = -((z*Sin(PhiXYZ)+b*Cos(PhiXYZ))*dAlphadz(x,y,z)+ExRe(AlphaXYZ, PhiXYZ, x,y,z)*dPhidz(x,y,z)+AlphaXYZ*Sin(PhiXYZ)+x/(z^2+b^2)*((z^2-b^2)*Sin(PhiXYZ)+2*z*b*Cos(PhiXYZ))*dAlphadx(x,y,z)+AlphaXYZ*((z^2-b^2)*Sin(PhiXYZ)+2*z*b*Cos(PhiXYZ))/(z^2+b^2)-EzRe(AlphaXYZ, PhiXYZ, x,y,z)*dPhidx(x,y,z))/omega/4e-7/Pi
End Function

Function HyIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HyIm = (-(z*Cos(PhiXYZ)-b*Sin(PhiXYZ))*dAlphadz(x,y,z)-ExIm(AlphaXYZ, PhiXYZ, x,y,z)*dPhidz(x,y,z)-AlphaXYZ*Cos(PhiXYZ)-x/(z^2+b^2)*((z^2-b^2)*Cos(PhiXYZ)-2*z*b*Sin(PhiXYZ))*dAlphadx(x,y,z)-AlphaXYZ*((z^2-b^2)*Cos(PhiXYZ)-2*z*b*Sin(PhiXYZ))/(z^2+b^2)-EzIm(AlphaXYZ, PhiXYZ, x,y,z)*dPhidx(x,y,z))/omega/4e-7/Pi
End Function

Function HzRe(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HzRe = 0
End Function

Function HzIm(AlphaXYZ As Double, PhiXYZ As Double, x As Double, y As Double, z As Double) As Double
	HzIm = 0
End Function

Function alpha(x As Double, y As Double, z As Double) As Double
	'Alpha = Sqr(k*b/Pi)*omega/(z^2+b^2)*Exp(-b*k*(x^2+y^2)/2/(z^2+b^2))	' rescale to avoid overly large field values
	' Renormalize such that Abs(Ex) = 1 V/m at focus point
	alpha = b/(z^2+b^2)*Exp(-b*k*(x^2+y^2)/2/(z^2+b^2))
End Function

Function dAlphadx(x As Double, y As Double, z As Double) As Double
	'dAlphadx = -k*b*x/(z^2+b^2)*Alpha(x,y,z)
	dAlphadx = -k*b^2*x/(z^2+b^2)^2*Exp(-b*k*(x^2+y^2)/2/(z^2+b^2)) ' calculate explicitely instead of calling Alpha(x,y,z) for better performance
End Function

Function dAlphady(x As Double, y As Double, z As Double) As Double
	'dAlphady = -k*b*y/(z^2+b^2)*Alpha(x,y,z)
	dAlphady = -k*b^2*y/(z^2+b^2)^2*Exp(-b*k*(x^2+y^2)/2/(z^2+b^2)) ' calculate explicitely instead of calling Alpha(x,y,z) for better performance
End Function

Function dAlphadz(x As Double, y As Double, z As Double) As Double
	'dAlphadz = Alpha(x,y,z)*z/(z^2+b^2)*(k*b*(x^2+y^2)/(z^2+b^2)-2)
	dAlphadz = z/(z^2+b^2)^2*(k*b*(x^2+y^2)/(z^2+b^2)-2)*b*Exp(-b*k*(x^2+y^2)/2/(z^2+b^2)) ' calculate explicitely instead of calling Alpha(x,y,z) for better performance
End Function

Function phi(x As Double, y As Double, z As Double) As Double
	phi = k*z*((x^2+y^2)/2/(z^2+b^2))
End Function

Function dPhidx(x As Double, y As Double, z As Double) As Double
	dPhidx = k*z*x/(z^2+b^2)
End Function

Function dPhidy(x As Double, y As Double, z As Double) As Double
	dPhidy = k*z*y/(z^2+b^2)
End Function

Function dPhidz(x As Double, y As Double, z As Double) As Double
	dPhidz = k*(1+(x^2+y^2)/2/(z^2+b^2)-z^2*(x^2+y^2)/(z^2+b^2)^2)
End Function

Sub calculateFieldsOnX(x As Double, y As Double, z As Double, _
						ByRef EEyRe As Double, ByRef EEyIm As Double, ByRef EEzRe As Double, ByRef EEzIm As Double, _
						ByRef HHyRe As Double, ByRef HHyIm As Double, ByRef HHzRe As Double, ByRef HHzIm As Double, _
						Amp1 As Double, Amp2 As Double, PhaseShiftDeg As Double)

	' Calculates the fields for a Gaussian beam in z direction at coordinates (x,y,z)
	' Ampl1, Amp2, and PhaseShiftDeg are used for circular polarization
	' Circular polarization is addition of two beams, one of which is rotated by 90 degrees about the z axis

	If (PhaseShiftDeg = 0) Then
		EEyRe = 0
		EEyIm = 0
		EEzRe = Amp1*EzReal
		EEzIm = Amp1*EzImag
		HHyRe = Amp1*HyReal
		HHyIm = Amp1*HyImag
		HHzRe = 0
		HHzIm = 0
	Else
		EEyRe = -Amp2*(COSDPH*ExReal-SINDPH*ExImag)
		EEyIm = -Amp2*(COSDPH*ExImag+SINDPH*ExReal)
		EEzRe = Amp1*EzReal+Amp2*(COSDPH*EzReal-SINDPH*EzImag)
		EEzIm = Amp1*EzImag+Amp2*(COSDPH*EzImag+SINDPH*EzReal)
		HHyRe = Amp1*HyReal
		HHyIm = Amp1*HyImag
		HHzRe = 0
		HHzIm = 0
	End If

End Sub

Sub calculateFieldsOnY(x As Double, y As Double, z As Double, _
						ByRef EExRe As Double, ByRef EExIm As Double, ByRef EEzRe As Double, ByRef EEzIm As Double, _
						ByRef HHxRe As Double, ByRef HHxIm As Double, ByRef HHzRe As Double, ByRef HHzIm As Double, _
						Amp1 As Double, Amp2 As Double, PhaseShiftDeg As Double)

	' Calculates the fields for a Gaussian beam in z direction at coordinates (x,y,z)
	' Ampl1, Amp2, and PhaseShiftDeg are used for circular polarization
	' Circular polarization is addition of two beams, one of which is rotated by 90 degrees about the z axis

	If (PhaseShiftDeg = 0) Then
		EExRe = Amp1*ExReal
		EExIm = Amp1*ExImag
		EEzRe = Amp1*EzReal
		EEzIm = Amp1*EzImag
		HHxRe = 0
		HHxIm = 0
		HHzRe = 0
		HHzIm = 0
	Else
		EExRe = Amp1*ExReal
		EExIm = Amp1*ExImag
		EEzRe = Amp1*EzReal+Amp2*(COSDPH*EzReal-SINDPH*EzImag)
		EEzIm = Amp1*EzImag+Amp2*(COSDPH*EzImag+SINDPH*EzReal)
		HHxRe = Amp2*(COSDPH*HyReal-SINDPH*HyImag)
		HHxIm = Amp2*(COSDPH*HyImag+SINDPH*HyReal)
		HHzRe = 0
		HHzIm = 0
	End If

End Sub

Sub calculateFieldsOnZ(x As Double, y As Double, z As Double, _
						ByRef EExRe As Double, ByRef EExIm As Double, ByRef EEyRe As Double, ByRef EEyIm As Double, _
						ByRef HHxRe As Double, ByRef HHxIm As Double, ByRef HHyRe As Double, ByRef HHyIm As Double, _
						Amp1 As Double, Amp2 As Double, PhaseShiftDeg As Double)

	' Calculates the fields for a Gaussian beam in z direction at coordinates (x,y,z)
	' Ampl1, Amp2, and PhaseShiftDeg are used for circular polarization
	' Circular polarization is addition of two beams, one of which is rotated by 90 degrees about the z axis

	If (PhaseShiftDeg = 0) Then
		EExRe = Amp1*ExReal
		EExIm = Amp1*ExImag
		EEyRe = 0
		EEyIm = 0
		HHxRe = 0
		HHxIm = 0
		HHyRe = Amp1*HyReal
		HHyIm = Amp1*HyImag
	Else
		EExRe = Amp1*ExReal
		EExIm = Amp1*ExImag
		EEyRe = -Amp2*(COSDPH*ExReal-SINDPH*ExImag)
		EEyIm = -Amp2*(COSDPH*ExImag+SINDPH*ExReal)
		HHxRe = Amp2*(COSDPH*HyReal-SINDPH*HyImag)
		HHxIm = Amp2*(COSDPH*HyImag+SINDPH*HyReal)
		HHyRe = Amp1*HyReal
		HHyIm = Amp1*HyImag
	End If

End Sub

Sub printFields(streamNum As Long, _
				E1Re As Double, E1Im As Double, E2Re As Double, E2Im As Double, _
				H1Re As Double, H1Im As Double, H2Re As Double, H2Im As Double)

	Print #streamNum, USFormat(E1Re,SFormat)+" "+USFormat(E1Im,SFormat)+" "+USFormat(E2Re,SFormat)+" "+USFormat(E2Im,SFormat)+" "+USFormat(H1Re,SFormat)+" "+USFormat(H1Im,SFormat)+" "+USFormat(H2Re,SFormat)+" "+USFormat(H2Im,SFormat)

End Sub

Function mycstr(aText As Double) As String
	mycstr=USFormat(aText, "0.##")
End Function

Function generateFaceFieldData(ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
						Nx As Integer, Ny As Integer, Nz As Integer, fMin As Double, fMax As Double, fSamp As Double, _
						bAmplitudesRepresentPower As Boolean, dInputAmp1 As Double, dInputAmp2 As Double, PhaseShiftDeg As Double, focusDistance As Double, minMax As String, b2DSource As Boolean, EpolarOB As Integer) As Integer

	' Outputs field data in to separate files for each face and field component according to the NFS file format specifications.
	' Exactly one of Nx, Ny, Nz must equal 1, the other two must be positive

	Dim Ere1 As Double, Ere2 As Double, Hre1 As Double, Hre2 As Double, Eim1 As Double, Eim2 As Double, Him1 As Double, Him2 As Double
	Dim Amp1 As Double, Amp2 As Double, wz As Double

	Dim Estream1 As Long, Estream2 As Long, Hstream1 As Long, Hstream2 As Long
	Dim sBuffer1 As String, sBuffer2 As String, sBuffer3 As String, sBuffer4 As String

	Dim x As Double, y As Double, z As Double, dFrequency As Double, dx As Double, dy As Double, dz As Double
	Dim i As Long, j As Long, kSigh As Integer, f As Integer
	Dim dTemp As Double

	Dim iCurrentLocale As Integer

	' Store current locale and switch to US temporarily. Then 'Format' can be used for output, which is faster than 'USFormat'
	iCurrentLocale = GetLocale
	SetLocale(&H409)
	On Error GoTo ExitFieldExport

	If x0 > x1 Then
		dTemp = x0
		x0 = x1
		x1 = dTemp
	End If
	If y0 > y1 Then
		dTemp = y0
		y0 = y1
		y1 = dTemp
	End If
	If z0 > z1 Then
		dTemp = z0
		z0 = z1
		z1 = dTemp
	End If

	' Global variables to be used in nested for loop, calculate outside of loop for performance
	COSDPH = CosD(PhaseShiftDeg)
	SINDPH = SinD(PhaseShiftDeg)

	wz = w0*Sqr(1+((z1-z0)/2/zR)^2)	'spot size(beam radius) at Lz

	If (Nx = 1) Then
		dx = 0
	Else
		dx = (x1 - x0)/(Nx-1)
	End If
	If (Ny = 1) Then
		dy = 0
	Else
		dy = (y1 - y0)/(Ny-1)
	End If
	If (Nz = 1) Then
		dz = 0
	Else
		dz = (z1 - z0)/(Nz-1)
	End If

	generateXmlFile(x0, y0, z0, x1, y1, z1, dx, dy, dz, fMin, fMax, fSamp, _
					IIf(Nx > 1,IIf(Ny > 1,"z","y"),"x"), _
					IIf(Nx > 1,"Ex","Ey"), minMax)
	generateXmlFile(x0, y0, z0, x1, y1, z1, dx, dy, dz, fMin, fMax, fSamp, _
					IIf(Nx > 1,IIf(Ny > 1,"z","y"),"x"), _
					IIf(Nz > 1,"Ez","Ey"), minMax)
	generateXmlFile(x0, y0, z0, x1, y1, z1, dx, dy, dz, fMin, fMax, fSamp, _
					IIf(Nx > 1,IIf(Ny > 1,"z","y"),"x"), _
					IIf(Nx > 1,"Hx","Hy"), minMax)
	generateXmlFile(x0, y0, z0, x1, y1, z1, dx, dy, dz, fMin, fMax, fSamp, _
					IIf(Nx > 1,IIf(Ny > 1,"z","y"),"x"), _
					IIf(Nz > 1,"Hz","Hy"), minMax)

	Estream1 = FreeFile
	Open GetProjectPath("TempDS")+"GB_"+IIf(Nx > 1,IIf(Ny > 1,"Ex_z","Ex_y"),"Ey_x")+minMax+".dat" For Output As #Estream1
	Estream2 = FreeFile
	Open GetProjectPath("TempDS")+"GB_"+IIf(Nx > 1,IIf(Ny > 1,"Ey_z","Ez_y"),"Ez_x")+minMax+".dat" For Output As #Estream2
	Hstream1 = FreeFile
	Open GetProjectPath("TempDS")+"GB_"+IIf(Nx > 1,IIf(Ny > 1,"Hx_z","Hx_y"),"Hy_x")+minMax+".dat" For Output As #Hstream1
	Hstream2 = FreeFile
	Open GetProjectPath("TempDS")+"GB_"+IIf(Nx > 1,IIf(Ny > 1,"Hy_z","Hz_y"),"Hz_x")+minMax+".dat" For Output As #Hstream2
		' The nested loops should iterate across the active face, as one and only one sample count should be 1

		' Create frequency-dependent but spatially independent values only once and store in array
		Dim dFrequencyList() As Double, kList() As Double, w0LambdaList() As Double, bList() As Double, omegaList() As Double, Amp1List() As Double, Amp2List() As Double

		ReDim dFrequencyList(fSamp-1)
		ReDim kList(fSamp-1)
		ReDim w0LambdaList(fSamp-1)
		ReDim bList(fSamp-1)
		ReDim omegaList(fSamp-1)
		ReDim Amp1List(fSamp-1)
		ReDim Amp2List(fSamp-1)

		For f = 0 To fSamp - 1
			If fSamp > 1 Then
				dFrequencyList(f) = fMax - f*(fMax-fMin)/(fSamp-1)
			Else
				dFrequencyList(f) = fMin
			End If
			kList(f) = -2*Pi*(dFrequencyList(f)/CLight)				' negative k for propagation in +z direction
			' b = w0^2*k/2
			w0LambdaList(f) = w0 ' * fMin / dFrequencyList(f) ' uncomment the last part to activate frequency-dependent beam radius
			bList(f) = w0LambdaList(f)^2*kList(f)/2
			omegaList(f) = 2*Pi*dFrequencyList(f)

			'=====================================================================================================================================================
			'Power and E-field orientation
			If bAmplitudesRepresentPower Then	'dInputAmp1 read in as power
				If b2DSource Then		  '2D strip option enabled
					If 	EpolarOB = 0 Then '2D 0 degree case
						Amp1List(f) = Sqr(2*dInputAmp1 * Sqr(Mue0/Eps0) * 1/(Sqr(pi)*w0LambdaList(f)*Sqr(2)*(y1-y0)))
						Amp2List(f) = Sqr(2*dInputAmp1 * Sqr(Mue0/Eps0) * 1/(Sqr(pi)*w0LambdaList(f)*Sqr(2)*(y1-y0)))
					ElseIf EpolarOB = 1 Then '2D 90 degree case
						Amp1List(f) = Sqr(2*dInputAmp1 * Sqr(Mue0/Eps0) * 1/(Sqr(pi)*w0LambdaList(f)*Sqr(2)*(x1-x0)))
						Amp2List(f) = Sqr(2*dInputAmp1 * Sqr(Mue0/Eps0) * 1/(Sqr(pi)*w0LambdaList(f)*Sqr(2)*(x1-x0)))
					End If

				Else					  '3D case
					Amp1List(f) = Sqr(2*dInputAmp1/Pi/w0LambdaList(f)^2*Sqr(Mue0/Eps0))
					Amp2List(f) = Sqr(2*dInputAmp2/Pi/w0LambdaList(f)^2*Sqr(Mue0/Eps0))
				End If
			Else						  'dInputAmp1 and 2 read in as E-field strength
					Amp1List(f) = dInputAmp1
					Amp2List(f) = dInputAmp2
			End If

			'=====================================================================================================================================================
		Next f

		For kSigh = 0 To Nz - 1
			z = IIf(Nz > 1, z0 + dz*kSigh, z0) - focusDistance
			For j = 0 To Ny - 1
				y = IIf(Ny > 1, y0 + dy*j, y0)
				If (b2DSource And EpolarOB = 0) Then y = 0 ' 0 degree rotated 2D beam
				For i = 0 To Nx - 1
					If((IIf(Nz > 1, kSigh, j) Mod 5=0) Or (IIf(Nz > 1, kSigh, j)=IIf(Nz > 1, Nz - 1, Ny - 1))) Then DlgText "outputT", "Writing data for "+IIf(Nx > 1,IIf(Ny > 1,"z","y"), "x")+minMax+": "+Format(IIf(Nz > 1, (kSigh+1)/Nz, (j+1)/Ny)*100,"00.00")+"%"
					If (DlgValue("AbortCB") = 1) Then
						DlgText("outputT", "Aborted.")
						DlgValue("AbortCB", 0)
						Close #Estream1
						Close #Estream2
						Close #Hstream1
						Close #Hstream2
						generateFaceFieldData = -2
						Exit Function
					End If
					x = IIf(Nx > 1, x0 + dx*i, x0)
					If (b2DSource And EpolarOB = 1) Then x = 0 ' 90 degree rotated 2D beam
					For f = 0 To fSamp - 1
						' Calculate some global variables to be used in CalculateFieldsOnX/Y/Z
						dFrequency 	= dFrequencyList(f)
						k 			= kList(f)
						w0Lambda 	= w0LambdaList(f)
						b 			= bList(f)
						omega 		= omegaList(f)
						Amp1 		= Amp1List(f)
						Amp2		= Amp2List(f)

						PhiXYZ = phi(x,y,z)
						AlphaXYZ = alpha(x,y,z)

						ExReal = ExRe(AlphaXYZ, PhiXYZ, x,y,z)
						ExImag = ExIm(AlphaXYZ, PhiXYZ, x,y,z)
						EyReal = EyRe(AlphaXYZ, PhiXYZ, x,y,z)
						EyImag = EyIm(AlphaXYZ, PhiXYZ, x,y,z)
						EzReal = EzRe(AlphaXYZ, PhiXYZ, x,y,z)
						EzImag = EzIm(AlphaXYZ, PhiXYZ, x,y,z)
						HxReal = HxRe(AlphaXYZ, PhiXYZ, x,y,z)
						HxImag = HxIm(AlphaXYZ, PhiXYZ, x,y,z)
						HyReal = HyRe(AlphaXYZ, PhiXYZ, x,y,z)
						HyImag = HyIm(AlphaXYZ, PhiXYZ, x,y,z)
						HzReal = HzRe(AlphaXYZ, PhiXYZ, x,y,z)
						HzImag = HzIm(AlphaXYZ, PhiXYZ, x,y,z)

						If (Nx <= 1) Then
							calculateFieldsOnX(x, y, z, Ere1, Eim1, Ere2, Eim2, Hre1, Him1, Hre2, Him2, _
							IIf(minMax="Min",-1,1)*Amp1, IIf(minMax="Min",-1,1)*Amp2, PhaseShiftDeg)
						ElseIf (Ny <= 1) Then
							calculateFieldsOnY(x, y, z, Ere1, Eim1, Ere2, Eim2, Hre1, Him1, Hre2, Him2, _
							IIf(minMax="Min",-1,1)*Amp1, IIf(minMax="Min",-1,1)*Amp2, PhaseShiftDeg)
						ElseIf (Nz <= 1) Then
							calculateFieldsOnZ(x, y, z, Ere1, Eim1, Ere2, Eim2, Hre1, Him1, Hre2, Him2, _
							IIf(minMax="Min",-1,1)*Amp1, IIf(minMax="Min",-1,1)*Amp2, PhaseShiftDeg)
						Else
							ReportError("Tried to output a 3D source.")
							generateFaceFieldData = -1
							Close #Estream1
							Close #Estream2
							Close #Hstream1
							Close #Hstream2
							Exit Function
						End If

						sBuffer1 = sBuffer1 & Format(Ere1, SxmlFormat)+" "+Format(Eim1, SxmlFormat)+IIf(f < fSamp-1, " ", "")
						sBuffer2 = sBuffer2 & Format(Ere2, SxmlFormat)+" "+Format(Eim2, SxmlFormat)+IIf(f < fSamp-1, " ", "")
						sBuffer3 = sBuffer3 & Format(Hre1, SxmlFormat)+" "+Format(Him1, SxmlFormat)+IIf(f < fSamp-1, " ", "")
						sBuffer4 = sBuffer4 & Format(Hre2, SxmlFormat)+" "+Format(Him2, SxmlFormat)+IIf(f < fSamp-1, " ", "")

						If (i*j Mod 1000 = 0) Then
							Print #Estream1, sBuffer1;
							Print #Estream2, sBuffer2;
							Print #Hstream1, sBuffer3;
							Print #Hstream2, sBuffer4;
							sBuffer1 = ""
							sBuffer2 = ""
							sBuffer3 = ""
							sBuffer4 = ""
						End If
						' FSC performance note 1/18/2016: Time spent about 1/3 on calculation and 2/3 on writing to file
					Next f
					' New lines
					sBuffer1 = sBuffer1 & vbNewLine
					sBuffer2 = sBuffer2 & vbNewLine
					sBuffer3 = sBuffer3 & vbNewLine
					sBuffer4 = sBuffer4 & vbNewLine
				Next i
			Next j
			' flush buffer
			Print #Estream1, sBuffer1;
			Print #Estream2, sBuffer2;
			Print #Hstream1, sBuffer3;
			Print #Hstream2, sBuffer4;
			sBuffer1 = ""
			sBuffer2 = ""
			sBuffer3 = ""
			sBuffer4 = ""
		Next kSigh

	Close #Estream1
	Close #Estream2
	Close #Hstream1
	Close #Hstream2
	generateFaceFieldData = 0

	ExitFieldExport:
	SetLocale(iCurrentLocale)

End Function

Sub generateXmlFile(x0 As Double, y0 As Double, z0 As Double, x1 As Double, y1 As Double, z1 As Double, _
					dx As Double, dy As Double, dz As Double, fMin As Double, fMax As Double, fSamp As Double, _
					sFace As String, sField As String, minMax As String)
	Dim f As Integer
	Dim streamNum As Long

	streamNum = OpenBufferedFile_LIB(GetProjectPath("TempDS")+"GBMacro_"+sField+"_"+sFace+minMax+".xml", "Output")
		BufferedFileWriteLine_LIB(streamNum,"<?xml version="+Chr(34)+"1.0"+Chr(34)+" encoding="+Chr(34)+"UTF-8"+Chr(34)+"?>")
		BufferedFileWriteLine_LIB(streamNum,"<EmissionScan>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"<Nfs_ver>1.0</Nfs_ver>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"<Filename>GBMacro_"+sField+"_"+sFace+minMax+".xml</Filename>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"<File_ver>1</File_ver>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"<Probe>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Field>"+sField+"</Field>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"</Probe>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"<Data>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Coordinates>none</Coordinates>")

		' XML file coordinate data will be in meters
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<X0>"+USFormat(x0, SxmlFormat)+"m</X0>")
		If (dx > 0) Then
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Xstep>"+USFormat(dx, SxmlFormat)+"m</Xstep>")
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Xmax>"+USFormat(x1, SxmlFormat)+"m</Xmax>")
		End If
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Y0>"+USFormat(y0, SxmlFormat)+"m</Y0>")
		If (dy > 0) Then
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Ystep>"+USFormat(dy, SxmlFormat)+"m</Ystep>")
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Ymax>"+USFormat(y1, SxmlFormat)+"m</Ymax>")
		End If
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Z0>"+USFormat(z0, SxmlFormat)+"m</Z0>")
		If (dz > 0) Then
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Zstep>"+USFormat(dz, SxmlFormat)+"m</Zstep>")
			BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Zmax>"+USFormat(z1, SxmlFormat)+"m</Zmax>")
		End If
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Frequencies>")
		BufferedFileWrite_LIB(streamNum,vbTab+vbTab+vbTab+"<List>")
		For f = 0 To fSamp - 1
			If fSamp > 1 Then
				BufferedFileWrite_LIB(streamNum,USFormat(fMax - f*(fMax-fMin)/(fSamp-1), SxmlFormat)+IIf(f < fSamp-1, " ", ""))
			Else
				BufferedFileWrite_LIB(streamNum,USFormat(fMin, SxmlFormat))
			End If
		Next f
		BufferedFileWriteLine_LIB(streamNum,"</List>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"</Frequencies>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"<Measurement>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+vbTab+"<Unit>V/m</Unit>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+vbTab+"<Format>ri</Format>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+vbTab+"<Data_files>"+"GB_"+sField+"_"+sFace+minMax+".dat</Data_files>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+vbTab+"</Measurement>")
		BufferedFileWriteLine_LIB(streamNum,vbTab+"</Data>")
		BufferedFileWriteLine_LIB(streamNum,"</EmissionScan>")
	CloseBufferedFile_LIB(streamNum)
End Sub

	
