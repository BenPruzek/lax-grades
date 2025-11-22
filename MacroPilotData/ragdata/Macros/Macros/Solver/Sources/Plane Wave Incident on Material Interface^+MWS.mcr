'#include "vba_globals_all.lib"

' This macro enable the user to set up a plane wave incidence on a material interface, for example to investigate scattering from an object on the interface.
'
' ================================================================================================
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------------------------------------------------------------
' 01-Jul-2013 fsr: Tweaking; TODO: Move COSR... into parametric function
' 26-Jun-2013 fsr: Added source ID, fixes a problem with multiple frequencies, frequency unit string now updated correctly, added some checks; bugfixes
' 01-Aug-2012 fsr: Multiple frequencies or wavelengths, parametric angle
' 05-Jul-2012 fsr: Initial version
' ------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Public Const sFormat = " 0.000000E+00;-0.000000E+00" ' 1e-7 as accuracy
Public Const sFreqFormat = "0.00000000"
Public lUnit As Double, fUnit As Double, lUnitS As String, fUnitS As String

Sub Main

	lUnit = Units.GetGeometryUnitToSI
	lUnitS = Units.GetUnit("Length")
	fUnit = Units.GetFrequencyUnitToSI
	fUnitS = Units.GetUnit("Frequency")

	If (fUnit < 1e9) Then
		If(MsgBox("The frequency unit is unusually small for optical applications, would you like to continue? Units have to be set up before macro execution.",vbYesNo,"Unit check")=vbNo) Then Exit All
	End If
	If (lUnit > 1e-3) Then
		If(MsgBox("The geometry unit is unusually large for optical applications, would you like to continue? Units have to be set up before macro execution.",vbYesNo,"Unit check")=vbNo) Then Exit All
	End If

	Dim WavelengthOrFrequency(1) As String
	WavelengthOrFrequency(0) = "Wavelength:"
	WavelengthOrFrequency(1) = "Frequency:"

	Begin Dialog UserDialog 570,308,"Generate Plane Wave Incident on Interface",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 30,14,140,121,WavelengthOrFrequency(),.WavelengthOrFrequencyDLB
		TextBox 200,14,60,21,.FMinT
		Text 270,21,20,14,"...",.Text2
		TextBox 290,14,60,21,.FmaxT
		Text 360,21,30,14,"nm",.FreqUnitsL
		Text 390,21,170,14,", using                sample(s).",.Text1
		TextBox 440,14,50,21,.FreqSamplesT
		Text 30,49,110,14,"Polarization type:",.Text13
		OptionGroup .PolarizationOG
			OptionButton 200,49,40,14,"s",.PolarizationSOB
			OptionButton 250,49,40,14,"p",.PolarizationPOB
		Text 30,140,130,14,"Angle of incidence:",.Text3
		TextBox 200,133,60,21,.ThetaT
		Text 270,140,30,14,"deg",.Text4
		CheckBox 40,161,160,14,"Parameterized angle",.ParameterizeThetaCB
		Picture 310,49,240,126,"Picture1",0,.Picture1
		Text 30,77,160,14,"Refractive index for z<0:",.Text5
		TextBox 200,70,60,21,.n1T
		Text 30,112,160,14,"Refractive index for z>0:",.Text6
		TextBox 200,105,60,21,.n2T
		Text 30,196,160,14,"Source size (x/y/z) in "+lUnitS+":",.Text7
		TextBox 200,189,60,21,.LxT
		TextBox 280,189,60,21,.LyT
		TextBox 360,189,60,21,.LzT
		Text 30,224,150,14,"Lines per wavelength:",.Text11
		TextBox 200,217,60,21,.SamplesT
		Text 30,252,120,14,"Add field monitors:",.Text15
		CheckBox 200,252,30,14,"E",.EMonitorsCB
		CheckBox 250,252,40,14,"H",.HMonitorsCB
		CheckBox 290,252,40,14,"FF",.FFMonitorsCB

		OKButton 360,280,90,21
		PushButton 460,280,90,21,"Exit",.ExitPB

	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim dFMin As Double, dFMax As Double, nFreqSamples As Long, dDeltaFreq As Double, iDeltaFreqModulo As Long
	Dim lambdaMin As Double, lambdaMax As Double

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("FMinT", "1000")
		DlgText("FMaxT", "1000")
		DlgText("FreqSamplesT", "1")
		DlgText("ThetaT", "0")
		DlgValue("PolarizationOG", 0)
		DlgText("n1T", "1")
		DlgText("n2T", "2")
		DlgText("LxT", Evaluate(3*Evaluate(DlgText("FMaxT"))*1e-9/lUnit))
		DlgText("LyT", Evaluate(3*Evaluate(DlgText("FMaxT"))*1e-9/lUnit))
		DlgText("LzT", Evaluate(3*Evaluate(DlgText("FMaxT"))*1e-9/lUnit))
		DlgText("SamplesT", "30")
		DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Solver\Sources\s-Fresnel.bmp",0)
		DlgValue("EMonitorsCB", 1)
		DlgValue("HMonitorsCB", 1)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "PolarizationOG"
				DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Solver\Sources\"+IIf(DlgValue("PolarizationOG")=0,"s-Fresnel.bmp","p-Fresnel.bmp"),0)
			Case "Cancel"
				Exit All
			Case "WavelengthOrFrequencyDLB"
				DlgText("FreqUnitsL", IIf(InStr(DlgText("WavelengthOrFrequencyDLB"), "Frequency"), fUnitS, "nm"))
			Case "OK"
				DlgEnable("OK", False)
				DlgEnable("ExitPB", False)
				Select Case DlgText("WavelengthOrFrequencyDLB")
					Case "Wavelength:"
						dFMin = CLight/Evaluate(DlgText("FMaxT"))/1e-9
						dFMax = CLight/Evaluate(DlgText("FminT"))/1e-9
					Case "Frequency:"
						dFMin = Evaluate(DlgText("FMinT"))*fUnit
						dFMax = Evaluate(DlgText("FmaxT"))*fUnit
				End Select
				nFreqSamples = Evaluate(DlgText("FreqSamplesT"))
				' Fmin must be an integer multiple of (Fmax-Fmin)/(NSamples-1) for broadband source -> adjust Fmax and thus FMin accordingly
				If (nFreqSamples > 1) Then
					' Make sure that Fmax>Fmin
					If (dFMax<=dFMin) Then
						MsgBox("The number of frequency samples is larger than 1, but Fmax is not larger than Fmin. Please check your settings.", "Frequency Check")
						DialogFunc = True
						DlgEnable("OK", True)
						DlgEnable("ExitPB", True)
						Exit Function
					End If
					dDeltaFreq = (dFMax-dFMin)/(nFreqSamples-1)
					iDeltaFreqModulo = Fix(dFMin/dDeltaFreq)
					dDeltaFreq = dFMin/iDeltaFreqModulo
					If (dFMax <> dFMin + (nFreqSamples-1) * dDeltaFreq) Then
						ReportInformationToWindow("Fresnel Source: The highest frequency (lowest wavelength) was increased (decreased) automatically to meet the nearfield source sampling requirements.")
						dFMax = dFMin + (nFreqSamples-1) * dDeltaFreq
					End If
				End If

				lambdaMin = CLight/dFMax
				lambdaMax = CLight/dFMin

				' Check if bounding box is of reasonable size
				If (Evaluate(DlgText("LxT"))*Evaluate(DlgText("LyT"))*Evaluate(DlgText("LzT"))*lUnit^3/lambdaMax^3 > 100) Then
					If (MsgBox("The source volume is larger than 100 cubic wavelengths, it may take a very long time to create the source. Do you wish to continue?", vbYesNo, "Source size check") = vbNo) Then
						DialogFunc = True
						DlgEnable("OK", True)
						DlgEnable("ExitPB", True)
						Exit Function
					End If
				End If

				If Solver.GetFMax = 0 Then AddToHistory("define frequency range", "Solver.FrequencyRange "+Chr(34)+"0"+Chr(34)+","+Chr(34)+CStr(Fix(1.1*dFMax/fUnit))+Chr(34))
				If CBool(DlgValue("ParameterizeThetaCB")) Then
					CreateSourceFileParameterized(dFMin, dFMax, nFreqSamples, _
													Evaluate(DlgText("ThetaT"))/180*Pi, IIf(DlgValue("PolarizationOG")=0, "s", "p"), Evaluate(DlgText("n1T")), Evaluate(DlgText("n2T")), _
													Evaluate(DlgText("LxT"))*lUnit, Evaluate(DlgText("LyT"))*lUnit, Evaluate(DlgText("LzT"))*lUnit, Evaluate(DlgText("SamplesT")))
				Else
					CreateSourceFile(dFMin, dFMax, nFreqSamples, _
										Evaluate(DlgText("ThetaT"))/180*Pi, IIf(DlgValue("PolarizationOG")=0, "s", "p"), Evaluate(DlgText("n1T")), Evaluate(DlgText("n2T")), _
										Evaluate(DlgText("LxT"))*lUnit, Evaluate(DlgText("LyT"))*lUnit, Evaluate(DlgText("LzT"))*lUnit, Evaluate(DlgText("SamplesT")))
				End If
				AddFieldMonitors(dFMin, dFMax, nFreqSamples, _
								CBool(DlgValue("EMonitorsCB")), CBool(DlgValue("HMonitorsCB")), CBool(DlgValue("FFMonitorsCB")))
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function CreateSourceFileParameterized(dFMin As Double, dFMax As Double, nFreqSamples As Long, _
										Theta_i As Double, sPolarizationType As String, n1 As Double, n2 As Double, _
										Lx As Double, Ly As Double, Lz As Double, dLinesPerWL As Double) As Integer

	Dim fsName As String, fileName As String
	Dim Nx As Long, Ny As Long, Nz As Long
	Dim sHistoryString As String

	sHistoryString = ""

	sHistoryString = sHistoryString + "Dim i As Long, j As Long, iFreq As Long, sFormat As String, sFreqFormat As String"+vbNewLine

	sHistoryString = sHistoryString + "Dim dFrequency As Double, dLambda As Double, dLambda1 As Double, dLambda2 As Double"+vbNewLine
	sHistoryString = sHistoryString + "' incident, reflected, transmitted angles; assumption: propagation in x-z-plane, theta_i = 0 is boresight on interface"+vbNewLine
	sHistoryString = sHistoryString + "Dim Theta_i As Double, Theta_r As Double, Theta_t As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim dFMin As Double, dFMax As Double, nFreqSamples As Long"+vbNewLine
	sHistoryString = sHistoryString + "Dim sPolarizationType As String, n1 As Double, n2 As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim Lx As Double, Ly As Double, Lz As Double"+vbNewLine
	sHistoryString = sHistoryString + "' incident field"+vbNewLine
	sHistoryString = sHistoryString + "Dim Eix As Double, Eiy As Double, Eiz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim Hix As Double, Hiy As Double, Hiz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim kix As Double, kiy As Double, kiz As Double"+vbNewLine
	sHistoryString = sHistoryString + "' reflected field"+vbNewLine
	sHistoryString = sHistoryString + "Dim Erx As Double, Ery As Double, Erz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim Hrx As Double, Hry As Double, Hrz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim krx As Double, kry As Double, krz As Double"+vbNewLine
	sHistoryString = sHistoryString + "' transmitted field"+vbNewLine
	sHistoryString = sHistoryString + "Dim Etx As Double, Ety As Double, Etz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim Htx As Double, Hty As Double, Htz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim ktx As Double, kty As Double, ktz As Double"+vbNewLine
	sHistoryString = sHistoryString + "' reflection and transmission coefficients"+vbNewLine
	sHistoryString = sHistoryString + "Dim r As Double, t As Double"+vbNewLine

	sHistoryString = sHistoryString + "Dim fileName As String, fsName As String"+vbNewLine
	sHistoryString = sHistoryString + "Dim streamnum As Long"+vbNewLine
	sHistoryString = sHistoryString + "Dim Nx As Long, Ny As Long, Nz As Long"+vbNewLine
	sHistoryString = sHistoryString + "Dim dx As Double, dy As Double, dz As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim x As Double, y As Double, z As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim EExRe As Double, EExIm As Double, EEyRe As Double, EEyIm As Double, EEzRe As Double, EEzIm As Double"+vbNewLine
	sHistoryString = sHistoryString + "Dim HHxRe As Double, HHxIm As Double, HHyRe As Double, HHyIm As Double, HHzRe As Double, HHzIm As Double"+vbNewLine

	' Write values for function parameters and other variables into history
	StoreParameter("theta", CStr(Theta_i))
	sHistoryString = sHistoryString + " Theta_i = Evaluate("+Chr(34)+"theta"+Chr(34)+")/180*Pi"+vbNewLine
	sHistoryString = sHistoryString + " dFMin = "+CStr(dFMin)+vbNewLine
	sHistoryString = sHistoryString + " dFMax = "+CStr(dFMax)+vbNewLine
	sHistoryString = sHistoryString + " nFreqSamples = "+CStr(nFreqSamples)+vbNewLine
	sHistoryString = sHistoryString + " sPolarizationType = "+Chr(34)+sPolarizationType+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + " n1 = "+CStr(n1)+vbNewLine
	sHistoryString = sHistoryString + " n2 = "+CStr(n2)+vbNewLine
	sHistoryString = sHistoryString + " Lx = "+CStr(Lx)+vbNewLine
	sHistoryString = sHistoryString + " Ly = "+CStr(Ly)+vbNewLine
	sHistoryString = sHistoryString + " Lz = "+CStr(Lz)+vbNewLine
	sHistoryString = sHistoryString + " sFormat = "+Chr(34)+sFormat+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + " sFreqFormat = "+Chr(34)+sFreqFormat+Chr(34)+vbNewLine

	sHistoryString = sHistoryString + " Theta_r = Theta_i"+vbNewLine
	sHistoryString = sHistoryString + " If Abs(Sin(Theta_i)/(n2/n1) < 1) Then"+vbNewLine
	sHistoryString = sHistoryString + " 	Theta_t = ASin(Sin(Theta_i)/(n2/n1))"+vbNewLine
	sHistoryString = sHistoryString + " Else"+vbNewLine
	sHistoryString = sHistoryString + " 	Theta_t = 90"+vbNewLine
	sHistoryString = sHistoryString + " End If"+vbNewLine

	Nx = Fix(dLinesPerWL*Lx/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1
	If Nx < 5 Then Nx = 5 ' Use at least 5 samples
	Ny = 5'Fix(dLinesPerWL*Ly/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1 ' source is homogeneous in y direction, 3 samples would suffice, use 5 samples for better interpolation
	Nz = Fix(dLinesPerWL*Lz/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1
	If Nz Mod 2>0 Then Nz = Nz+1 ' Nz should always be even
	If Nz < 5 Then Nz = 4 ' Use at least 4 samples
	sHistoryString = sHistoryString + "	Nx = "+CStr(Nx)+vbNewLine
	sHistoryString = sHistoryString + "	Ny = "+CStr(Ny)+vbNewLine
	sHistoryString = sHistoryString + "	Nz = "+CStr(Nz)+vbNewLine

	sHistoryString = sHistoryString + "	dx = "+CStr(Lx/Nx)+vbNewLine
	sHistoryString = sHistoryString + "	dy = "+CStr(Ly/Ny)+vbNewLine
	sHistoryString = sHistoryString + "	dz = "+CStr(Lz/Nz)+vbNewLine

	sHistoryString = sHistoryString + "	streamnum = FreeFile"+vbNewLine
	fileName = GetProjectPath("TempDS")+"FresnelInterface.nfd"
	sHistoryString = sHistoryString + "	fileName = "+Chr(34)+fileName+Chr(34)+vbNewLine

	sHistoryString = sHistoryString + "	If fileName="+Chr(34)+Chr(34)+" Then"+vbNewLine
	sHistoryString = sHistoryString + "		ReportError("+Chr(34)+"Invalid file Name!"+Chr(34)+")"+vbNewLine
	sHistoryString = sHistoryString + "	End If"+vbNewLine

	sHistoryString = sHistoryString + "' Temporarily switch to US locale to ensure period is used as separator"+vbNewLine
	sHistoryString = sHistoryString + "Dim iCurrentLocale As Long"+vbNewLine
	sHistoryString = sHistoryString + "iCurrentLocale = GetLocale"+vbNewLine
	sHistoryString = sHistoryString + "SetLocale &H409"+vbNewLine

	sHistoryString = sHistoryString + "	' Ouput everything in SI units"+vbNewLine
	sHistoryString = sHistoryString + "	Open fileName For Output As #streamnum"+vbNewLine
	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"cell_number"+Chr(34)+"+Str(Nx)+Str(Ny)+Str(Nz)"+vbNewLine
	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"cell_size "+Chr(34)+"+Format(dx,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(dy,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(dz,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"box_min "+Chr(34)+"+Format(-Lx/2,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(-Ly/2,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(-Lz/2,sFormat)"+vbNewLine

	sHistoryString = sHistoryString + " For iFreq = 0 To nFreqSamples-1"+vbNewLine

	sHistoryString = sHistoryString + "	If nFreqSamples > 1 Then"+vbNewLine
	sHistoryString = sHistoryString + "		dFrequency = dFMax - iFreq*(dFMax-dFmin)/(nFreqSamples-1)"+vbNewLine
	sHistoryString = sHistoryString + "	Else"+vbNewLine
	sHistoryString = sHistoryString + "		dFrequency = dFmin"+vbNewLine
	sHistoryString = sHistoryString + "	End If"+vbNewLine
	sHistoryString = sHistoryString + "	dLambda = CLight/dFrequency ' free space wave length for reference"+vbNewLine
	sHistoryString = sHistoryString + "	dLambda1 = CLight/n1/dFrequency"+vbNewLine
	sHistoryString = sHistoryString + "	dLambda2 = CLight/n2/dFrequency"+vbNewLine

	sHistoryString = sHistoryString + "	kix = Sin(Theta_i)*2*Pi/dLambda1"+vbNewLine
	sHistoryString = sHistoryString + "	kiy = 0"+vbNewLine
	sHistoryString = sHistoryString + "	kiz = Cos(Theta_i)*2*Pi/dLambda1"+vbNewLine

	sHistoryString = sHistoryString + "	krx = Sin(Theta_r)*2*Pi/dLambda1"+vbNewLine
	sHistoryString = sHistoryString + "	kry = 0"+vbNewLine
	sHistoryString = sHistoryString + "	krz = -Cos(Theta_r)*2*Pi/dLambda1"+vbNewLine

	sHistoryString = sHistoryString + "	ktx = Sin(Theta_t)*2*Pi/dLambda2"+vbNewLine
	sHistoryString = sHistoryString + "	kty = 0"+vbNewLine
	sHistoryString = sHistoryString + "	ktz = Cos(Theta_t)*2*Pi/dLambda2"+vbNewLine

	sHistoryString = sHistoryString + "	If (sPolarizationType = "+Chr(34)+"s"+Chr(34)+") Then"+vbNewLine
	sHistoryString = sHistoryString + "		Eix = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Eiy = 1"+vbNewLine
	sHistoryString = sHistoryString + "		Eiz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Hix = -kiz/2/Pi/dFrequency/Mue0"+vbNewLine
	sHistoryString = sHistoryString + "		Hiy = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Hiz = kix/2/Pi/dFrequency/Mue0"+vbNewLine

	sHistoryString = sHistoryString + "		r = (n1*Cos(Theta_i)-n2*Cos(Theta_t))/(n1*Cos(Theta_i)+n2*Cos(Theta_t))"+vbNewLine
	sHistoryString = sHistoryString + "		t = IIf(Theta_t = 90, 0, 2*n1*Cos(Theta_i)/(n1*Cos(Theta_i)+n2*Cos(Theta_t)))"+vbNewLine

	sHistoryString = sHistoryString + "		Erx = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Ery = r*Eiy"+vbNewLine
	sHistoryString = sHistoryString + "		Erz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Hrx = r*Hix"+vbNewLine
	sHistoryString = sHistoryString + "		Hry = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Hrz = -r*Hiz"+vbNewLine

	sHistoryString = sHistoryString + "		Etx = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Ety = t*Eiy"+vbNewLine
	sHistoryString = sHistoryString + "		Etz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Htx = t*Hix/Cos(Theta_i)*Cos(Theta_t)"+vbNewLine
	sHistoryString = sHistoryString + "		Hty = 0"+vbNewLine
	sHistoryString = sHistoryString + "		If (Theta_i=0) Then"+vbNewLine
	sHistoryString = sHistoryString + "			Htz = t*Hiz"+vbNewLine
	sHistoryString = sHistoryString + "		Else"+vbNewLine
	sHistoryString = sHistoryString + "			Htz = t*Hiz/Sin(Theta_i)*Sin(Theta_t)"+vbNewLine
	sHistoryString = sHistoryString + "		End If"+vbNewLine

	sHistoryString = sHistoryString + "	ElseIf (sPolarizationType = "+Chr(34)+"p"+Chr(34)+") Then"+vbNewLine
	sHistoryString = sHistoryString + "		Hix = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Hiy = Eps0*n1*CLight"+vbNewLine
	sHistoryString = sHistoryString + "		Hiz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Eix = kiz*Hiy/2/Pi/dFrequency/Eps0/n1^2"+vbNewLine
	sHistoryString = sHistoryString + "		Eiy = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Eiz = -kix*Hiy/2/Pi/dFrequency/Eps0/n1^2"+vbNewLine

	sHistoryString = sHistoryString + "		r = (n2*Cos(Theta_i)-n1*Cos(Theta_t))/(n1*Cos(Theta_t)+n2*Cos(Theta_i))"+vbNewLine
	sHistoryString = sHistoryString + "		t = 2*n1*Cos(Theta_i)/(n1*Cos(Theta_t)+n2*Cos(Theta_i))"+vbNewLine

	sHistoryString = sHistoryString + "		Hrx = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Hry = r*Hiy"+vbNewLine
	sHistoryString = sHistoryString + "		Hrz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Erx = r*Eix"+vbNewLine
	sHistoryString = sHistoryString + "		Ery = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Erz = -r*Eiz"+vbNewLine

	sHistoryString = sHistoryString + "		Htx = 0"+vbNewLine
	sHistoryString = sHistoryString + "		Hty = t*Hiy"+vbNewLine
	sHistoryString = sHistoryString + "		Htz = 0"+vbNewLine

	sHistoryString = sHistoryString + "		Etx = t*Eix/Cos(Theta_i)*Cos(Theta_t)"+vbNewLine
	sHistoryString = sHistoryString + "		Ety = 0"+vbNewLine
	sHistoryString = sHistoryString + "		If (Theta_i=0) Then"+vbNewLine
	sHistoryString = sHistoryString + "			Etz = t*Eiz"+vbNewLine
	sHistoryString = sHistoryString + "		Else"+vbNewLine
	sHistoryString = sHistoryString + "			Etz = t*Eiz/Sin(Theta_i)*Sin(Theta_t)"+vbNewLine
	sHistoryString = sHistoryString + "		End If"+vbNewLine
	sHistoryString = sHistoryString + "	Else"+vbNewLine
	sHistoryString = sHistoryString + "		ReportError("+Chr(34)+"Undefined polarization Type."+Chr(34)+")"+vbNewLine
	sHistoryString = sHistoryString + "	End If"+vbNewLine

	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"data"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"frequency "+Chr(34)+"+Format(dFrequency, sFreqFormat)"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Xlower"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			x = -Lx/2"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Ny"+vbNewLine
	sHistoryString = sHistoryString + "				y = -Ly/2 - dy/2 + i*dy"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nz/2 ' in medium 1"+vbNewLine
	sHistoryString = sHistoryString + "					z = -Lz/2 - dz/2 + j*dz"+vbNewLine
	sHistoryString = sHistoryString + "						' E = Ei + Er, H = Hi + Hr"+vbNewLine
	sHistoryString = sHistoryString + "						EEyRe = Eiy*Cos(-(kix*x+kiy*y+kiz*z))+Ery*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEyIm = Eiy*Sin(-(kix*x+kiy*y+kiz*z))+Ery*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzRe = Eiz*Cos(-(kix*x+kiy*y+kiz*z))+Erz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzIm = Eiz*Sin(-(kix*x+kiy*y+kiz*z))+Erz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyRe = Hiy*Cos(-(kix*x+kiy*y+kiz*z))+Hry*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyIm = Hiy*Sin(-(kix*x+kiy*y+kiz*z))+Hry*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzRe = Hiz*Cos(-(kix*x+kiy*y+kiz*z))+Hrz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzIm = Hiz*Sin(-(kix*x+kiy*y+kiz*z))+Hrz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						Print #streamnum, Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "					For j = Nz/2 + 1 To Nz ' in medium 2"+vbNewLine
	sHistoryString = sHistoryString + "					z = -Lz/2 - dz/2 + j*dz"+vbNewLine
	sHistoryString = sHistoryString + "						' E = Et, H = Ht"+vbNewLine
	sHistoryString = sHistoryString + "						EEyRe = Ety*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEyIm = Ety*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzRe = Etz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzIm = Etz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyRe = Hty*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyIm = Hty*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzRe = Htz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzIm = Htz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						Print #streamnum, Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Xupper"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			x = Lx/2"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Ny"+vbNewLine
	sHistoryString = sHistoryString + "				y = -Ly/2 - dy/2 + i*dy"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nz/2 ' in medium 1"+vbNewLine
	sHistoryString = sHistoryString + "					z = -Lz/2 - dz/2 + j*dz"+vbNewLine
	sHistoryString = sHistoryString + "						' E = Ei + Er, H = Hi + Hr"+vbNewLine
	sHistoryString = sHistoryString + "						EEyRe = Eiy*Cos(-(kix*x+kiy*y+kiz*z))+Ery*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEyIm = Eiy*Sin(-(kix*x+kiy*y+kiz*z))+Ery*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzRe = Eiz*Cos(-(kix*x+kiy*y+kiz*z))+Erz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzIm = Eiz*Sin(-(kix*x+kiy*y+kiz*z))+Erz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyRe = Hiy*Cos(-(kix*x+kiy*y+kiz*z))+Hry*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyIm = Hiy*Sin(-(kix*x+kiy*y+kiz*z))+Hry*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzRe = Hiz*Cos(-(kix*x+kiy*y+kiz*z))+Hrz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzIm = Hiz*Sin(-(kix*x+kiy*y+kiz*z))+Hrz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						Print #streamnum, Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "				For j = Nz/2 + 1 To Nz ' in medium 2"+vbNewLine
	sHistoryString = sHistoryString + "					z = -Lz/2 - dz/2 + j*dz"+vbNewLine
	sHistoryString = sHistoryString + "						' E = Et, H = Ht"+vbNewLine
	sHistoryString = sHistoryString + "						EEyRe = Ety*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEyIm = Ety*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzRe = Etz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						EEzIm = Etz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyRe = Hty*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHyIm = Hty*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzRe = Htz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						HHzIm = Htz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "						Print #streamnum, Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)+" _
																+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Ylower"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			y = -Ly/2"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Nz/2"+vbNewLine
	sHistoryString = sHistoryString + "				z = -Lz/2 - dz/2 + i*dz"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "					x = -Lx/2 - dx/2 + j*dx"+vbNewLine
	sHistoryString = sHistoryString + "					EEzRe = Eiz*Cos(-(kix*x+kiy*y+kiz*z))+Erz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEzIm = Eiz*Sin(-(kix*x+kiy*y+kiz*z))+Erz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Eix*Cos(-(kix*x+kiy*y+kiz*z))+Erx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Eix*Sin(-(kix*x+kiy*y+kiz*z))+Erx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzRe = Hiz*Cos(-(kix*x+kiy*y+kiz*z))+Hrz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzIm = Hiz*Sin(-(kix*x+kiy*y+kiz*z))+Hrz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Hix*Cos(-(kix*x+kiy*y+kiz*z))+Hrx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Hix*Sin(-(kix*x+kiy*y+kiz*z))+Hrx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			For i = Nz/2+1 To Nz"+vbNewLine
	sHistoryString = sHistoryString + "				z = -Lz/2 - dz/2 + i*dz"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "					x = -Lx/2 - dx/2 + j*dx"+vbNewLine
	sHistoryString = sHistoryString + "					EEzRe = Etz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEzIm = Etz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Etx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Etx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzRe = Htz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzIm = Htz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Htx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Htx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Yupper"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			y = Ly/2"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Nz/2"+vbNewLine
	sHistoryString = sHistoryString + "				z = -Lz/2 - dz/2 + i*dz"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "					x = -Lx/2 - dx/2 + j*dx"+vbNewLine
	sHistoryString = sHistoryString + "					EEzRe = Eiz*Cos(-(kix*x+kiy*y+kiz*z))+Erz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEzIm = Eiz*Sin(-(kix*x+kiy*y+kiz*z))+Erz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Eix*Cos(-(kix*x+kiy*y+kiz*z))+Erx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Eix*Sin(-(kix*x+kiy*y+kiz*z))+Erx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzRe = Hiz*Cos(-(kix*x+kiy*y+kiz*z))+Hrz*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzIm = Hiz*Sin(-(kix*x+kiy*y+kiz*z))+Hrz*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Hix*Cos(-(kix*x+kiy*y+kiz*z))+Hrx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Hix*Sin(-(kix*x+kiy*y+kiz*z))+Hrx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			For i = Nz/2+1 To Nz"+vbNewLine
	sHistoryString = sHistoryString + "				z = -Lz/2 - dz/2 + i*dz"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "					x = -Lx/2 - dx/2 + j*dx"+vbNewLine
	sHistoryString = sHistoryString + "					EEzRe = Etz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEzIm = Etz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Etx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Etx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzRe = Htz*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHzIm = Htz*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Htx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Htx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EEzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHzRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHzIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Zlower"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			z = -Lz/2 ' all in medium 1"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "				x = -Lx/2 - dx/2 + i*dx"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Ny"+vbNewLine
	sHistoryString = sHistoryString + "					y = -Ly/2 - dy/2 + j*dy"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Eix*Cos(-(kix*x+kiy*y+kiz*z))+Erx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Eix*Sin(-(kix*x+kiy*y+kiz*z))+Erx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEyRe = Eiy*Cos(-(kix*x+kiy*y+kiz*z))+Ery*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEyIm = Eiy*Sin(-(kix*x+kiy*y+kiz*z))+Ery*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Hix*Cos(-(kix*x+kiy*y+kiz*z))+Hrx*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Hix*Sin(-(kix*x+kiy*y+kiz*z))+Hrx*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHyRe = Hiy*Cos(-(kix*x+kiy*y+kiz*z))+Hry*Cos(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHyIm = Hiy*Sin(-(kix*x+kiy*y+kiz*z))+Hry*Sin(-(krx*x+kry*y+krz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"Zupper"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"{"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "			z = Lz/2 ' all in medium 2"+vbNewLine
	sHistoryString = sHistoryString + "			For i = 1 To Nx"+vbNewLine
	sHistoryString = sHistoryString + "				x = -Lx/2 - dx/2 + i*dx"+vbNewLine
	sHistoryString = sHistoryString + "				For j = 1 To Ny"+vbNewLine
	sHistoryString = sHistoryString + "					y = -Ly/2 - dy/2 + j*dy"+vbNewLine
	sHistoryString = sHistoryString + "					EExRe = Etx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EExIm = Etx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEyRe = Ety*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					EEyIm = Ety*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxRe = Htx*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHxIm = Htx*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHyRe = Hty*Cos(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					HHyIm = Hty*Sin(-(ktx*x+kty*y+ktz*z))"+vbNewLine
	sHistoryString = sHistoryString + "					Print #streamnum, Format(EExRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EExIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(EEyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(EEyIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHxRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHxIm,sFormat)+" _
															+Chr(34)+" "+Chr(34)+"+Format(HHyRe,sFormat)+"+Chr(34)+" "+Chr(34)+"+Format(HHyIm,sFormat)"+vbNewLine
	sHistoryString = sHistoryString + "				Next j"+vbNewLine
	sHistoryString = sHistoryString + "			Next i"+vbNewLine
	sHistoryString = sHistoryString + "			Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine
	sHistoryString = sHistoryString + "		Print #streamnum, "+Chr(34)+"}"+Chr(34)+vbNewLine

	sHistoryString = sHistoryString + "	Next iFreq"+vbNewLine

	sHistoryString = sHistoryString + "Close #streamnum"+vbNewLine
	'sHistoryString = sHistoryString + "Wait 5"+vbNewLine
	'sHistoryString = sHistoryString + "Shell "+Chr(34)+"notepad "+Chr(34)+" & fileName, 3"

	sHistoryString = sHistoryString + "' Copy file into project 3D folder and import field source"+vbNewLine
	fsName="fsFresnelMacro"
	sHistoryString = sHistoryString + "FileCopy(fileName, GetProjectPath("+Chr(34)+"Model3D"+Chr(34)+")+"+Chr(34)+"FresnelSource^"+fsName+".nfd"+Chr(34)+")"+vbNewLine
	'sHistoryString = sHistoryString + "Wait 5"+vbNewLine

	sHistoryString = sHistoryString + "With FieldSource"+vbNewLine
	sHistoryString = sHistoryString + " .Delete "+Chr(34)+fsName+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .Reset"+vbNewLine
    sHistoryString = sHistoryString + " .Name "+Chr(34)+fsName+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .FileName "+Chr(34)+"*FresnelSource^"+fsName+".nfd"+Chr(34)+vbNewLine
    ' Initialize FieldSource object to get next valid ID
    FieldSource.FileName("*FresnelSource^"+fsName+".nfd")
    sHistoryString = sHistoryString + " .ID "+Chr(34)+Cstr(FieldSource.GetNextID)+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .Read"+vbNewLine
	sHistoryString = sHistoryString + "End With"+vbNewLine

	sHistoryString = sHistoryString + "' Change back to default locale"+vbNewLine
	sHistoryString = sHistoryString + "SetLocale iCurrentLocale"+vbNewLine

	AddToHistory("Define Fresnel Interface", sHistoryString)

End Function

Function CreateSourceFile(dFMin As Double, dFMax As Double, nFreqSamples As Long, _
							Theta_i As Double, sPolarizationType As String, n1 As Double, n2 As Double, _
							Lx As Double, Ly As Double, Lz As Double, dLinesPerWL As Double) As Integer

	Dim i As Long, j As Long, iFreq As Long

	Dim dFrequency As Double, dLambda As Double, dLambda1 As Double, dLambda2 As Double
	' incident, reflected, transmitted angles; assumption: propagation in x-z-plane, theta_i = 0 is boresight on interface
	Dim Theta_r As Double, Theta_t As Double
	' incident field
	Dim Eix As Double, Eiy As Double, Eiz As Double
	Dim Hix As Double, Hiy As Double, Hiz As Double
	Dim kix As Double, kiy As Double, kiz As Double
	' reflected field
	Dim Erx As Double, Ery As Double, Erz As Double
	Dim Hrx As Double, Hry As Double, Hrz As Double
	Dim krx As Double, kry As Double, krz As Double
	' transmitted field
	Dim Etx As Double, Ety As Double, Etz As Double
	Dim Htx As Double, Hty As Double, Htz As Double
	Dim ktx As Double, kty As Double, ktz As Double
	' reflection and transmission coefficients
	Dim r As Double, t As Double

	Dim fileName As String, fsName As String
	Dim streamnum As Long
	Dim Nx As Long, Ny As Long, Nz As Long
	Dim dx As Double, dy As Double, dz As Double
	Dim x As Double, y As Double, z As Double
	Dim EExRe As Double, EExIm As Double, EEyRe As Double, EEyIm As Double, EEzRe As Double, EEzIm As Double
	Dim HHxRe As Double, HHxIm As Double, HHyRe As Double, HHyIm As Double, HHzRe As Double, HHzIm As Double
	Dim sHistoryString As String
	Dim COSI As Double, SINI As Double, COSR As Double, SINR As Double, COST As Double, SINT As Double

	Theta_r = Theta_i
	If Abs(Sin(Theta_i)/(n2/n1) < 1) Then
		Theta_t = ASin(Sin(Theta_i)/(n2/n1))
	Else
		Theta_t = 90
	End If

	Nx = Fix(dLinesPerWL*Lx/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1
	If Nx < 5 Then Nx = 5 ' Use at least 5 samples
	Ny = 5'Fix(dLinesPerWL*Ly/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1 ' source is homogeneous in y direction, 3 samples would suffice, use 5 samples for better interpolation
	Nz = Fix(dLinesPerWL*Lz/IIf(n2>n1, CLight/n2/dFMax, CLight/n1/dFMax))+1
	If Nz Mod 2>0 Then Nz = Nz+1 ' Nz should always be even
	If Nz < 5 Then Nz = 4 ' Use at least 4 samples

	dx = Lx/Nx
	dy = Ly/Ny
	dz = Lz/Nz

	streamnum = FreeFile
	fileName = GetProjectPath("TempDS")+"FresnelInterface.nfd"
	If fileName="" Then
		MsgBox("Invalid file name!")
		Exit All
	End If

	' Ouput everything in SI units
	Open fileName For Output As #streamnum
		Print #streamnum, "cell_number"+Str(Nx)+Str(Ny)+Str(Nz)
		Print #streamnum, "cell_size "+USFormat(dx,sFormat)+" "+USFormat(dy,sFormat)+" "+USFormat(dz,sFormat)
		Print #streamnum, "box_min "+USFormat(-Lx/2,sFormat)+" "+USFormat(-Ly/2,sFormat)+" "+USFormat(-Lz/2,sFormat)

		For iFreq = 0 To nFreqSamples-1

			If nFreqSamples > 1 Then
				dFrequency = dFMax - iFreq*(dFMax-dFMin)/(nFreqSamples-1)
			Else
				dFrequency = dFMin
			End If

			'ReportInformationToWindow("Fresnel Source: Calculating fields for frequency sample "+CStr(iFreq+1)+"/"+CStr(nFreqSamples))

			dLambda = CLight/dFrequency ' free space wave length for reference
			dLambda1 = CLight/n1/dFrequency
			dLambda2 = CLight/n2/dFrequency

			kix = Sin(Theta_i)*2*Pi/dLambda1
			kiy = 0
			kiz = Cos(Theta_i)*2*Pi/dLambda1

			krx = Sin(Theta_r)*2*Pi/dLambda1
			kry = 0
			krz = -Cos(Theta_r)*2*Pi/dLambda1

			ktx = Sin(Theta_t)*2*Pi/dLambda2
			kty = 0
			ktz = Cos(Theta_t)*2*Pi/dLambda2

			If (sPolarizationType = "s") Then
				Eix = 0
				Eiy = 1
				Eiz = 0

				Hix = -kiz/2/Pi/dFrequency/Mue0*Eiy
				Hiy = 0
				Hiz = kix/2/Pi/dFrequency/Mue0*Eiy

				' reflection and transmission factors for the ELECTRIC field
				r = (n1*Cos(Theta_i)-n2*Cos(Theta_t))/(n1*Cos(Theta_i)+n2*Cos(Theta_t))
				t = IIf(Theta_t = 90, 0, 2*n1*Cos(Theta_i)/(n1*Cos(Theta_i)+n2*Cos(Theta_t)))

				Erx = 0
				Ery = r*Eiy
				Erz = 0

				Hrx = -krz/2/Pi/dFrequency/Mue0*Ery
				Hry = 0
				Hrz = krx/2/Pi/dFrequency/Mue0*Ery

				Etx = 0
				Ety = t*Eiy
				Etz = 0

				Htx = -ktz/2/Pi/dFrequency/Mue0*Ety
				Hty = 0
				Htz = ktx/2/Pi/dFrequency/Mue0*Ety

			ElseIf (sPolarizationType = "p") Then
				Hix = 0
				Hiy = 2*Pi*dFrequency/CLight^2*n1^2/Mue0/Sqr(kix^2+kiz^2)
				Hiz = 0

				Eix = CLight^2*Mue0/2/Pi/dFrequency/n1^2*kiz*Hiy
				Eiy = 0
				Eiz = -CLight^2*Mue0/2/Pi/dFrequency/n1^2*kix*Hiy

				' reflection and transmission factors for the ELECTRIC field
				'r = (n2*Cos(Theta_i)-n1*Cos(Theta_t))/(n1*Cos(Theta_t)+n2*Cos(Theta_i))
				't = 2*n1*Cos(Theta_i)/(n1*Cos(Theta_t)+n2*Cos(Theta_i))
				' reflection and transmission factors for the MAGNETIC field
				r = (n2^2*Cos(Theta_i)-n1^2*Cos(Theta_t))/(n2^2*Cos(Theta_r)+n1^2*Cos(Theta_t))
				t = (2*Cos(Theta_i))/(n1^2/n2^2*Cos(Theta_t)+Cos(Theta_r))

				Hrx = 0
				Hry = r*Hiy
				Hrz = 0

				Erx = CLight^2*Mue0/2/Pi/dFrequency/n1^2*krz*Hry
				Ery = 0
				Erz = -CLight^2*Mue0/2/Pi/dFrequency/n1^2*krx*Hry

				Htx = 0
				Hty = t*Hiy
				Htz = 0

				Etx = CLight^2*Mue0/2/Pi/dFrequency/n2^2*ktz*Hty
				Ety = 0
				Etz = -CLight^2*Mue0/2/Pi/dFrequency/n2^2*ktx*Hty

			Else
				MsgBox("Undefined polarization type.")
				Exit All
			End If
			Print #streamnum, "data"
			Print #streamnum, "{"
				Print #streamnum, "frequency "+USFormat(dFrequency, sFreqFormat)
				Print #streamnum, "Xlower"
				Print #streamnum, "{"
					x = -Lx/2
					For i = 1 To Ny
						y = -Ly/2 - dy/2 + i*dy
						For j = 1 To Nz/2 ' in medium 1
							z = -Lz/2 - dz/2 + j*dz
								' E = Ei + Er, H = Hi + Hr
								COSI = Cos(-(kix*x+kiy*y+kiz*z))
								SINI = Sin(-(kix*x+kiy*y+kiz*z))
								COSR = Cos(-(krx*x+kry*y+krz*z))
								SINR = Sin(-(krx*x+kry*y+krz*z))
								EEyRe = Eiy*COSI+Ery*COSR
								EEyIm = Eiy*SINI+Ery*SINR
								EEzRe = Eiz*COSI+Erz*COSR
								EEzIm = Eiz*SINI+Erz*SINR
								HHyRe = Hiy*COSI+Hry*COSR
								HHyIm = Hiy*SINI+Hry*SINR
								HHzRe = Hiz*COSI+Hrz*COSR
								HHzIm = Hiz*SINI+Hrz*SINR
								printFields(streamnum, EEyRe, EEyIm, EEzRe, EEzIm, HHyRe, HHyIm, HHzRe, HHzIm)
						Next j
						For j = Nz/2 + 1 To Nz ' in medium 2
							z = -Lz/2 - dz/2 + j*dz
								' E = Et, H = Ht
								COST = Cos(-(ktx*x+kty*y+ktz*z))
								SINT = Sin(-(ktx*x+kty*y+ktz*z))
								EEyRe = Ety*COST
								EEyIm = Ety*SINT
								EEzRe = Etz*COST
								EEzIm = Etz*SINT
								HHyRe = Hty*COST
								HHyIm = Hty*SINT
								HHzRe = Htz*COST
								HHzIm = Htz*SINT
								printFields(streamnum, EEyRe, EEyIm, EEzRe, EEzIm, HHyRe, HHyIm, HHzRe, HHzIm)
						Next j
					Next i
					Print #streamnum, "}"
					Print #streamnum, "Xupper"
					Print #streamnum, "{"
					x = Lx/2
					For i = 1 To Ny
						y = -Ly/2 - dy/2 + i*dy
						For j = 1 To Nz/2 ' in medium 1
							z = -Lz/2 - dz/2 + j*dz
							' E = Ei + Er, H = Hi + Hr
							COSI = Cos(-(kix*x+kiy*y+kiz*z))
							SINI = Sin(-(kix*x+kiy*y+kiz*z))
							COSR = Cos(-(krx*x+kry*y+krz*z))
							SINR = Sin(-(krx*x+kry*y+krz*z))
							EEyRe = Eiy*COSI+Ery*COSR
							EEyIm = Eiy*SINI+Ery*SINR
							EEzRe = Eiz*COSI+Erz*COSR
							EEzIm = Eiz*SINI+Erz*SINR
							HHyRe = Hiy*COSI+Hry*COSR
							HHyIm = Hiy*SINI+Hry*SINR
							HHzRe = Hiz*COSI+Hrz*COSR
							HHzIm = Hiz*SINI+Hrz*SINR
							printFields(streamnum, EEyRe, EEyIm, EEzRe, EEzIm, HHyRe, HHyIm, HHzRe, HHzIm)
						Next j
						For j = Nz/2 + 1 To Nz ' in medium 2
							z = -Lz/2 - dz/2 + j*dz
							' E = Et, H = Ht
							COST = Cos(-(ktx*x+kty*y+ktz*z))
							SINT = Sin(-(ktx*x+kty*y+ktz*z))
							EEyRe = Ety*COST
							EEyIm = Ety*SINT
							EEzRe = Etz*COST
							EEzIm = Etz*SINT
							HHyRe = Hty*COST
							HHyIm = Hty*SINT
							HHzRe = Htz*COST
							HHzIm = Htz*SINT
							printFields(streamnum, EEyRe, EEyIm, EEzRe, EEzIm, HHyRe, HHyIm, HHzRe, HHzIm)
						Next j
					Next i
					Print #streamnum, "}"
					Print #streamnum, "Ylower"
					Print #streamnum, "{"
					y = -Ly/2
					For i = 1 To Nz/2
						z = -Lz/2 - dz/2 + i*dz
						For j = 1 To Nx
							x = -Lx/2 - dx/2 + j*dx
							COSI = Cos(-(kix*x+kiy*y+kiz*z))
							SINI = Sin(-(kix*x+kiy*y+kiz*z))
							COSR = Cos(-(krx*x+kry*y+krz*z))
							SINR = Sin(-(krx*x+kry*y+krz*z))
							EEzRe = Eiz*COSI+Erz*COSR
							EEzIm = Eiz*SINI+Erz*SINR
							EExRe = Eix*COSI+Erx*COSR
							EExIm = Eix*SINI+Erx*SINR
							HHzRe = Hiz*COSI+Hrz*COSR
							HHzIm = Hiz*SINI+Hrz*SINR
							HHxRe = Hix*COSI+Hrx*COSR
							HHxIm = Hix*SINI+Hrx*SINR
							printFields(streamnum, EEzRe, EEzIm, EExRe, EExIm, HHzRe, HHzIm, HHxRe, HHxIm)
						Next j
					Next i
					For i = Nz/2+1 To Nz
						z = -Lz/2 - dz/2 + i*dz
						For j = 1 To Nx
							x = -Lx/2 - dx/2 + j*dx
							COST = Cos(-(ktx*x+kty*y+ktz*z))
							SINT = Sin(-(ktx*x+kty*y+ktz*z))
							EEzRe = Etz*COST
							EEzIm = Etz*SINT
							EExRe = Etx*COST
							EExIm = Etx*SINT
							HHzRe = Htz*COST
							HHzIm = Htz*SINT
							HHxRe = Htx*COST
							HHxIm = Htx*SINT
							printFields(streamnum, EEzRe, EEzIm, EExRe, EExIm, HHzRe, HHzIm, HHxRe, HHxIm)
						Next j
					Next i
					Print #streamnum, "}"
					Print #streamnum, "Yupper"
					Print #streamnum, "{"
					y = Ly/2
					For i = 1 To Nz/2
						z = -Lz/2 - dz/2 + i*dz
						For j = 1 To Nx
							x = -Lx/2 - dx/2 + j*dx
							COSI = Cos(-(kix*x+kiy*y+kiz*z))
							SINI = Sin(-(kix*x+kiy*y+kiz*z))
							COSR = Cos(-(krx*x+kry*y+krz*z))
							SINR = Sin(-(krx*x+kry*y+krz*z))
							EEzRe = Eiz*COSI+Erz*COSR
							EEzIm = Eiz*SINI+Erz*SINR
							EExRe = Eix*COSI+Erx*COSR
							EExIm = Eix*SINI+Erx*SINR
							HHzRe = Hiz*COSI+Hrz*COSR
							HHzIm = Hiz*SINI+Hrz*SINR
							HHxRe = Hix*COSI+Hrx*COSR
							HHxIm = Hix*SINI+Hrx*SINR
							printFields(streamnum, EEzRe, EEzIm, EExRe, EExIm, HHzRe, HHzIm, HHxRe, HHxIm)
						Next j
					Next i
					For i = Nz/2+1 To Nz
						z = -Lz/2 - dz/2 + i*dz
						For j = 1 To Nx
							x = -Lx/2 - dx/2 + j*dx
							COST = Cos(-(ktx*x+kty*y+ktz*z))
							SINT = Sin(-(ktx*x+kty*y+ktz*z))
							EEzRe = Etz*COST
							EEzIm = Etz*SINT
							EExRe = Etx*COST
							EExIm = Etx*SINT
							HHzRe = Htz*COST
							HHzIm = Htz*SINT
							HHxRe = Htx*COST
							HHxIm = Htx*SINT
							printFields(streamnum, EEzRe, EEzIm, EExRe, EExIm, HHzRe, HHzIm, HHxRe, HHxIm)
						Next j
					Next i
					Print #streamnum, "}"
					Print #streamnum, "Zlower"
					Print #streamnum, "{"
					z = -Lz/2 ' all in medium 1
					For i = 1 To Nx
						x = -Lx/2 - dx/2 + i*dx
						For j = 1 To Ny
							y = -Ly/2 - dy/2 + j*dy
							COSI = Cos(-(kix*x+kiy*y+kiz*z))
							SINI = Sin(-(kix*x+kiy*y+kiz*z))
							COSR = Cos(-(krx*x+kry*y+krz*z))
							SINR = Sin(-(krx*x+kry*y+krz*z))
							EExRe = Eix*COSI+Erx*COSR
							EExIm = Eix*SINI+Erx*SINR
							EEyRe = Eiy*COSI+Ery*COSR
							EEyIm = Eiy*SINI+Ery*SINR
							HHxRe = Hix*COSI+Hrx*COSR
							HHxIm = Hix*SINI+Hrx*SINR
							HHyRe = Hiy*COSI+Hry*COSR
							HHyIm = Hiy*SINI+Hry*SINR
							printFields(streamnum, EExRe, EExIm, EEyRe, EEyIm, HHxRe, HHxIm, HHyRe, HHyIm)
						Next j
					Next i
					Print #streamnum, "}"
					Print #streamnum, "Zupper"
					Print #streamnum, "{"
					z = Lz/2 ' all in medium 2
					For i = 1 To Nx
						x = -Lx/2 - dx/2 + i*dx
						For j = 1 To Ny
							y = -Ly/2 - dy/2 + j*dy
							COST = Cos(-(ktx*x+kty*y+ktz*z))
							SINT = Sin(-(ktx*x+kty*y+ktz*z))
							EExRe = Etx*COST
							EExIm = Etx*SINT
							EEyRe = Ety*COST
							EEyIm = Ety*SINT
							HHxRe = Htx*COST
							HHxIm = Htx*SINT
							HHyRe = Hty*COST
							HHyIm = Hty*SINT
							printFields(streamnum, EExRe, EExIm, EEyRe, EEyIm, HHxRe, HHxIm, HHyRe, HHyIm)
						Next j
					Next i
				Print #streamnum, "}"
			Print #streamnum, "}"

		Next iFreq

	Close #streamnum

	'Shell "notepad " & fileName, 3

	sHistoryString = ""

	' Copy file into project 3D folder and import field source
	fsName="fsFresnelMacro_"+USFormat(dLambda*1e9,"000.000")+"nm"
	FileCopy(fileName, GetProjectPath("Model3D")+"FresnelSource^"+fsName+".nfd")
	sHistoryString = sHistoryString + "With FieldSource"+vbNewLine
	sHistoryString = sHistoryString + " .Delete "+Chr(34)+fsName+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .Reset"+vbNewLine
    sHistoryString = sHistoryString + " .Name "+Chr(34)+fsName+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .FileName "+Chr(34)+"*FresnelSource^"+fsName+".nfd"+Chr(34)+vbNewLine
    ' Initialize FieldSource object to get next valid ID
    FieldSource.FileName("*FresnelSource^"+fsName+".nfd")
    sHistoryString = sHistoryString + " .ID "+Chr(34)+Cstr(FieldSource.GetNextID)+Chr(34)+vbNewLine
    sHistoryString = sHistoryString + " .Read"+vbNewLine
	sHistoryString = sHistoryString + "End With"+vbNewLine

	AddToHistory("Define Fresnel Interface (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)

End Function

Function AddFieldMonitors(dFMin As Double, dFMax As Double, nFreqSamples As Long, _
							bAddEMonitors As Boolean, bAddHMonitors As Boolean, bAddFFMonitors As Boolean) As Integer

	Dim sHistoryString As String
	Dim iFreq As Long, dFrequency As Double, dLambda As Double

	' Define e and h field monitors at proper frequency
	If bAddEMonitors Then
		For iFreq = 0 To nFreqSamples-1

			If nFreqSamples > 1 Then
				dFrequency = dFMax - iFreq*(dFMax-dFMin)/(nFreqSamples-1)
			Else
				dFrequency = dFMin
			End If
			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"e-field (lambda="+cstr(dLambda*1e9)+"nm)"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Efield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define monitor: e-field (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
		Next iFreq
	End If

	If bAddHMonitors Then
		For iFreq = 0 To nFreqSamples-1

			If nFreqSamples > 1 Then
				dFrequency = dFMax - iFreq*(dFMax-dFMin)/(nFreqSamples-1)
			Else
				dFrequency = dFMin
			End If
			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"h-field (lambda="+cstr(dLambda*1e9)+"nm)"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Hfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define monitor: h-field (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
		Next iFreq
	End If

	If bAddFFMonitors Then
		sHistoryString = ""
		If (nFreqSamples = 1) Then
			dLambda = CLight/dFrequency
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (lambda="+cstr(dLambda*1e9)+"nm)"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define farfield monitor: farfield (lambda="+USFormat(dLambda*1e9,"000.000")+" nm)", sHistoryString)
		ElseIf (Solver.GetFMax > 0) Then
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (broadband)"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Time"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Accuracy "+Chr(34)+"1e-3"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FrequencySamples "+Chr(34)+CStr(2*nFreqSamples+1)+Chr(34)+vbNewLine
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

Sub printFields(streamnum As Long, _
				E1Re As Double, E1Im As Double, E2Re As Double, E2Im As Double, _
				H1Re As Double, H1Im As Double, H2Re As Double, H2Im As Double)

	Print #streamnum, USFormat(E1Re,sFormat)+" "+USFormat(E1Im,sFormat)+" "+USFormat(E2Re,sFormat)+" "+USFormat(E2Im,sFormat)+" "+USFormat(H1Re,sFormat)+" "+USFormat(H1Im,sFormat)+" "+USFormat(H2Re,sFormat)+" "+USFormat(H2Im,sFormat)

End Sub
