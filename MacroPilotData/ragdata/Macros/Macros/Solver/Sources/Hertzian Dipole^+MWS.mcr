'#Language "WWB-COM"

'#include "vba_globals_all.lib"

'------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 08-Jan-2018 ksr,ubr: Dipole source rotation is enabled.
' 03-Jan-2018 ksr,ubr: XML file coordinates data unit type will be in meters and converted to project units when imported.
' 17-Mar-2017 ksr,ubr: Added additional Inputs, E-Field at 1m for Electric Dipole, H-Field at 1m for Magnetic Dipole and Dipole moment for both dipoles.
' 07-Mar-2017 fsr: Minor GUI improvements; set default orientation to 0/0/1; set default excitation signal to Gaussian sine
' 19-Dec-2016 wch: sheet electricl and magnetic current source for broadband frequencies and background materials; disable transient broadband ffs option
' 02-Dec-2016 wch: Use sheet electric/magnetic current as field source (single frequency; vacuum)
' 02-Nov-2016 wch: Added an ideal magnetic dipole option and made settings general for all frequencies; added the reference
' 19-Jun-2015 jfl: Changed to NFS format. Fixed transforms when WCS is active. Optimized calculations.
' 26-Jun-2014 yta: option to normalize to 1 W or of 1 V/m
' 27-Sep-2013 fsr: Minor tweaks and some bugfixes
' 30-Aug-2013 fde: Fixed automatic calculation of dFmax from dFmin and number of samples for broadband source
' 26-Jun-2013 hsn: proper handling of field source id's
' 01-Aug-2012 fsr: initial version with GUI
'------------------------------------------------------------------------------------------------------------------------------------
' Reference: Antenna Theory: Analysis and design 2nd Edition (Constantine A. Balanis) Wiley
' Electric dipole radiation field/power formula P. 135-137, Magnetic dipole radiation field/power formula P. 207-209
' Use farfield approximation to calculate Power and use it for field renormalization
'------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Const HelpFileName = "common_preloadedmacro_solver_hertzian_dipole"

'Public k As Double		' propagation constant

Public SFormat As String
Public lUnit As Double
Public lUnitS As String
Public fUnit As Double
Public fUnitS As String
Public Const SxmlFormat = "0.0000000000000000E+00;-0.0000000000000000E+00"


Sub Main

	lUnit = Units.GetGeometryUnitToSI
	lUnitS = Units.GetUnit("Length")
	fUnit = Units.GetFrequencyUnitToSI
	fUnitS = Units.GetUnit("Frequency")

	Dim WavelengthOrFrequency(1) As String
	WavelengthOrFrequency(0) = "Wavelength:"
	WavelengthOrFrequency(1) = "Frequency:"

	Dim DipoleType(1) As String
	DipoleType(0) = "Electric Dipole"
	DipoleType(1) = "Magnetic Dipole"

	Dim PowerOrEFieldOrMoment(3) As String
	PowerOrEFieldOrMoment(0) = "Power:"
	PowerOrEFieldOrMoment(1) = "E-Field at 1 meter:"
	PowerOrEFieldOrMoment(2) = "Dipole moment:"
	PowerOrEFieldOrMoment(3) = "H-Field at 1 meter:"

	Begin Dialog UserDialog 440,273,"Generate Hertzian Dipole Source",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,420,231,"Settings",.SettingsGB
		Text 30,35,90,14,"Dipole Type:",.Text7
		DropListBox 210,28,200,21,DipoleType(),.DipoleTypeDLB
		DropListBox 30,63,160,49,WavelengthOrFrequency(),.WavelengthOrFrequencyDLB
		TextBox 210,63,60,21,.dFminT
		Text 270,63,10,14,"...",.Text2
		TextBox 280,63,60,21,.dFmaxT
		Text 350,70,30,14,"freq project unit",.FreqUnitsL
		Text 250,98,30,14,"in",.Text6
		TextBox 280,91,60,21,.FreqSamplesT
		Text 350,98,60,14,"samples.",.Text1
		Text 30,126,160,14,"Dipole orientation (x/y/z):",.Text9
		TextBox 210,119,60,21,.orientationXT
		Text 270,126,10,14,"/",.Text13
		TextBox 280,119,60,21,.orientationYT
		Text 340,126,10,14,"/",.Text14
		TextBox 350,119,60,21,.orientationZT
		DropListBox 30,147,160,49,PowerOrEFieldOrMoment(),.PowerOrEFieldOrMomentDLB
		Text 280,151,30,14,"W",.PwrUnitsL
		TextBox 210,147,60,21,.AmpT
		Text 30,182,110,14,"Background:",.Text3
		Text 150,182,50,14,"eps_r =",.Text4
		TextBox 210,175,60,21,.eps_rT
		Text 290,182,60,14,", mu_r =",.Text5
		TextBox 360,175,50,21,.Mu_rT
		Text 30,210,120,14,"Add field monitors:",.Text15
		CheckBox 210,210,30,14,"E",.EMonitorsCB
		CheckBox 260,210,40,14,"H",.HMonitorsCB
		CheckBox 310,210,40,14,"FF",.FFMonitorsCB
		Text 10,238,90,28,"",.outputT
		CheckBox 30,245,60,14,"Abort",.AbortCB
		PushButton 160,245,90,21,"OK",.OkPB
		PushButton 250,245,90,21,"Exit",.ExitPB
		CancelButton 250,245,90,21 ' needed to activate "X" button in top right corner
		PushButton 340,245,90,21,"Help",.HelpB
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
		Exit All
	End If

End Sub


Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Dim dFmin As Double, dFmax As Double, FreqSamples As Long, dDeltaFreq As Double, iDeltaFreqModulo As Long, tmp As Double
	Dim dLambdaMin As Double, dLambdaMax As Double

	'Check frequency range settings
	If (Solver.GetFmax = 0 And Solver.GetFmin = 0) Then
		MsgBox("Please set the frequency range before using the macro.")
		Exit All
	End If

	Dim SLambdaMid As String
	SLambdaMid = Format(CLight/((Solver.GetFmin+Solver.GetFmax)/2*fUnit)/lUnit,"###0.00")

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgVisible("Cancel", False) ' only included to enable X in top right corner
		DlgText("dFminT", SLambdaMid)
		DlgText("dFmaxT", SLambdaMid)
		DlgText("FreqSamplesT", "1")
		DlgText("orientationXT", "0")
		DlgText("orientationYT", "0")
		DlgText("orientationZT", "1")
		DlgText("AmpT", Format(1, "##0.00"))
		DlgText("eps_rT", "1")
		DlgText("Mu_rT", "1")
		DlgText("FreqUnitsL", lUnitS)

	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "HelpB"
				StartHelp HelpFileName
				DialogFunc = True

			Case "ExitPB"
				Exit All

			Case "WavelengthOrFrequencyDLB"
				Dim dpMinValue As Double, dpMaxValue As Double
				If (DlgText("WavelengthOrFrequencyDLB") = "Frequency:") Then
					'convert value to freq
					DlgText("FreqUnitsL", fUnitS)
					dpMinValue = CLight/Evaluate(DlgText("dFmaxT"))/lUnit/fUnit
					dpMaxValue = CLight/Evaluate(DlgText("dFminT"))/lUnit/fUnit
					DlgText("dFminT", Format(dpMinValue,"##0.00"))
					DlgText("dFmaxT", Format(dpMaxValue,"##0.00"))
				ElseIf (DlgText("WavelengthOrFrequencyDLB") = "Wavelength:") Then
					DlgText("FreqUnitsL", lUnitS)
					dpMinValue = CLight/Evaluate(DlgText("dFmaxT"))/lUnit/fUnit
					dpMaxValue = CLight/Evaluate(DlgText("dFminT"))/lUnit/fUnit
					DlgText("dFminT", Format(dpMinValue,"##0.00"))
					DlgText("dFmaxT", Format(dpMaxValue,"##0.00"))
				End If

			Case "PowerOrEFieldOrMomentDLB", "DipoleTypeDLB"
				If (DlgText("PowerOrEFieldOrMomentDLB") = "Power:") Then
					DlgText("PwrUnitsL", "W")
				ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "E-Field at 1 meter:") Then
					DlgText("PwrUnitsL", "V/m")
				ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "H-Field at 1 meter:") Then
					DlgText("PwrUnitsL", "A/m")
				ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "Dipole moment:") Then
					If (DlgText("DipoleTypeDLB") = "Electric Dipole") Then
						DlgText("PwrUnitsL", "Am")
					Else
						DlgText("PwrUnitsL", "Vm")
					End If
				End If

			Case "OkPB"
				DlgEnable "OkPB", False
				DlgEnable "ExitPB", False
				DlgEnable "HelpB", False

				Select Case DlgText("WavelengthOrFrequencyDLB")
					Case "Wavelength:"
						dFmin = CLight/Evaluate(DlgText("dFmaxT"))/lUnit
						dFmax = CLight/Evaluate(DlgText("dFminT"))/lUnit
					Case "Frequency:"
						dFmin = Evaluate(DlgText("dFminT"))*fUnit
						dFmax = Evaluate(DlgText("dFmaxT"))*fUnit
				End Select

				If (dFmin > dFmax) Then
					tmp = dFmin
					dFmin = dFmax
					dFmax = tmp
				End If

				FreqSamples = Evaluate(DlgText("FreqSamplesT"))

				dLambdaMin = CLight/dFmax
				dLambdaMax = CLight/dFmin

				If Solver.GetFmax = 0 Then AddToHistory("define frequency range", "Solver.FrequencyRange "+Chr(34)+"0"+Chr(34)+","+Chr(34)+CStr(Fix(1.1*dFmax/fUnit))+Chr(34)+Chr(34))

				DialogFunc = True ' by default, leave dialog open
				If (DlgText("DipoleTypeDLB") = "Magnetic Dipole") Then
					If (DlgText("PowerOrEFieldOrMomentDLB") = "E-Field at 1 meter:") Then
					MsgBox("For Magnetic Dipole please specify input 'H-field at 1 meter!'")
					DialogFunc = True
					DlgEnable "OkPB", True
					DlgEnable "ExitPB", True
					DlgEnable "HelpB", True
					Exit Function
					End If
				End If

				If (DlgText("DipoleTypeDLB") = "Electric Dipole") Then
					If (DlgText("PowerOrEFieldOrMomentDLB") = "H-Field at 1 meter:") Then
					MsgBox("For Electric Dipole please specify input 'E-field at 1 meter!'")
					DialogFunc = True
					DlgEnable "OkPB", True
					DlgEnable "ExitPB", True
					DlgEnable "HelpB", True
					Exit Function
					End If
				End If

				If dFmax/fUnit > Solver.GetFmax Then
					MsgBox("Source frequency range must be within solver frequency range. Please check your settings.", "Input Error")
				ElseIf ((FreqSamples > 1) And (dFmax <= dFmin)) Then
					MsgBox("Multiple samples requested but frequency range is zero. Please check your settings.", "Input Error")
				ElseIf (CreateSourceFile(Evaluate(DlgText("orientationXT")), Evaluate(DlgText("orientationYT")), Evaluate(DlgText("orientationZT")), _
								    Evaluate(DlgText("AmpT")), Evaluate(DlgText("eps_rT")), Evaluate(DlgText("Mu_rT")), _
									dFmin, dFmax, FreqSamples)=0) Then
					 ' All went well, close Dialog
					DialogFunc = False

					' Reportinformationtowindow(FreqSamples)
					AddFieldMonitors(dFmin, dFmax, FreqSamples, _
										CBool(DlgValue("EMonitorsCB")), CBool(DlgValue("HMonitorsCB")), CBool(DlgValue("FFMonitorsCB")))
				Else
					MsgBox("Error creating source file", "Error")
				End If
				DlgEnable "OkPB", True
				DlgEnable "ExitPB", True
				DlgEnable "HelpB", True

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function CreateSourceFile( orientationX As Double, orientationY As Double, orientationZ As Double, _
							AmpT As Double, EpsR As Double, MuR As Double, dFmin As Double, dFmax As Double, nFreqSamples As Double) As Integer

	Dim fileName As String
	Dim fsName As String
	Dim streamNum As Long
	Dim historyString As String
	Dim sFreqFormat As String
	Dim dFrequency As Double

	Dim i As Long, j As Long, iFreqSample As Long
	Dim x As Double, y As Double, z As Double		' global coordinates

	Dim DipoleType(1) As String
	DipoleType(0) = "Electric_Dipole"
	DipoleType(1) = "Magnetic_Dipole"

	historyString = ""

	CreateSourceFile = -1
	SFormat = " 0.0000000000000000E+00;-0.0000000000000000E+00"
	sFreqFormat = "0.00000000"

	streamNum = FreeFile
	Open GetProjectPath("Model3D")+"HertzianDipoleInfo.txt" For Output As #streamNum
		Print #streamNum, "Hertzian dipole parameters:"
		If DlgValue("DipoleTypeDLB")  = 0 Then
			Print #streamNum, "Hertzian dipole type      : ideal electric dipole"
		Else
			Print #streamNum, "Hertzian dipole type      : ideal magnetic dipole"
		End If
		Print #streamNum, "Frequency                 : " + USFormat(dFmin/fUnit,"###0.000") + " ... "  +USFormat(dFmax/fUnit,"###0.000")  + " " + fUnitS + " in " + DlgText("FreqSamplesT") + " samples."
		Print #streamNum, "Wavelength                : " + USFormat(CLight/dFmax/lUnit,"###0.000") + " ... " +USFormat(CLight/dFmin/lUnit,"###0.000") + " " + lUnitS + " in " + DlgText("FreqSamplesT") + " samples."
        Print #streamNum, "Background (EpsR,MuR)     : (" + DlgText("eps_rT")+"/"+DlgText("Mu_rT")+")"
        ' FSR: Do not show orientation and center here as they are parameterized and may be changed by user
		' Print #streamNum, "Dipole orientation (x/y/z): (" + cstr(orientationX)+"/"+cstr(orientationY)+"/"+cstr(orientationZ)+")"
		' Print #streamNum, "Source center             : (0/0/0) "+lUnitS
		If (DlgText("PowerOrEFieldOrMomentDLB") = "Power:") Then
			Print #streamNum, "Average radiated power	  : " + Cstr(Evaluate(DlgText("AmpT"))/2) + " Watts"
		ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "E-Field at 1 meter:") Then
			Print #streamNum, "E-Field at 1 meter    	  : " + Cstr(Evaluate(DlgText("AmpT"))) + " Volt per meter"
		ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "H-Field at 1 meter:") Then
			Print #streamNum, "H-Field at 1 meter    	  : " + Cstr(Evaluate(DlgText("AmpT"))) + " Ampere per meter"
		ElseIf (DlgText("PowerOrEFieldOrMomentDLB") = "Dipole moment:") Then
			If (DlgText("DipoleTypeDLB") = "Electric Dipole") Then
				Print #streamNum, "Electric Dipole Moment	  : " + Cstr(Evaluate(DlgText("AmpT"))) + " Ampere meter"
			Else
				Print #streamNum, "Magnetic Dipole Moment     : " + Cstr(Evaluate(DlgText("AmpT"))) + " Volt meter"
			End If
		End If
		Close #streamNum

	streamNum = FreeFile
	fileName = GetProjectPath("TempDS")+"HertzianDipole.nfd"
	If fileName="" Then
		MsgBox("Invalid file name!")
		CreateSourceFile=1
		Exit Function
	End If


	' Set names
	Dim nFiles As Integer
	Dim wcsActive As Boolean
	Dim fsFileName As String
	Dim exportPath As String
	Dim fsID As String


	'Name field source based on either frequency or wavelength
	Dim sfsName As String
	If DlgValue("WavelengthOrFrequencyDLB") = 0 Then
		sfsName = USFormat(CLight/dFmin/lUnit,"###0.000")+lUnitS '
	Else
		sfsName = CStr(dFmin/fUnit)+fUnitS
	End If


	If DlgValue("DipoleTypeDLB")  = 0 Then
	    fsName="fsElectricDipole_" + sfsName
	Else
		fsName="fsMagneticDipole_" + sfsName
	End If

    exportPath = GetProjectPath("Root")+"\HDMacro_Export^"+Split(GetProjectPath("Project"),"\")(UBound(Split(GetProjectPath("Project"),"\")))+"\"

    wcsActive = WCS.IsWCSActive() = "local"
    If Len(Dir(exportPath, vbDirectory)) = 0 Then MkDir exportPath

   	generateHDSCurrent_FieldData(0, 0, 0, dFmin, dFmax, nFreqSamples, AmpT, EpsR, MuR, Evaluate(DlgValue("DipoleTypeDLB")) )

	FileCopy(GetProjectPath("TempDS")+"HD_"+DipoleType(DlgValue("DipoleTypeDLB"))+"_FieldData.dat", exportPath+"HD_"+DipoleType(DlgValue("DipoleTypeDLB"))+"_FieldData.dat")
	FileCopy(GetProjectPath("TempDS")+"HDMacro_"+DipoleType(DlgValue("DipoleTypeDLB"))+".xml", exportPath+"HDMacro_"+DipoleType(DlgValue("DipoleTypeDLB"))+".xml")

	On Error GoTo IOError

GoTo IOSuccess
IOError:
	MsgBox("Could not export files. This can sometimes happen when the project is stored in a temp directory. Please save the project and try again.")
	Exit All
IOSuccess:
	' Get ID
	With FieldSource
		.Reset
		.FileName "HDMacro_"+DipoleType(DlgValue("DipoleTypeDLB"))+".xml"
		fsID = .GetNextId
		.Reset
	End With

	' Disable WCS
	If (wcsActive) Then historyString = historyString + "WCS.ActivateWCS("+Chr(34)+"global"+Chr(34)+")"+vbNewLine

    ' ' Construction of field source import history command
	historyString = historyString + "With FieldSource"+vbNewLine
	historyString = historyString + " .Delete "+Chr(34)+fsName+Chr(34)+vbNewLine
    historyString = historyString + " .Reset"+vbNewLine
    historyString = historyString + " .Name "+Chr(34)+fsName+Chr(34)+vbNewLine
    historyString = historyString + " .FileName "+Chr(34)+exportPath+"HDMacro_"+DipoleType(DlgValue("DipoleTypeDLB"))+".xml"+Chr(34)+vbNewLine
	historyString = historyString + " .Id "+Chr(34)+cstr(cint(fsID))+Chr(34)+vbNewLine
    historyString = historyString + " .Read"+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Shell "notepad " & fileName, 3

	StoreParameter("theta", ACosD(orientationZ/Sqr(orientationX^2+orientationY^2+orientationZ^2)))
	StoreParameter("phi", Atn2D(orientationY, orientationX))
	StoreParameter("SourceOriginX", "0")
	StoreParameter("SourceOriginY", "0")
	StoreParameter("SourceOriginZ", "0")

	' ' Rotate source to adjust orientation vector
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

	' ' Apply translation of origin (0 by default)
	historyString = historyString + "With Transform"+vbNewLine
	historyString = historyString + "	.Reset"+vbNewLine
    historyString = historyString + "	.Name "+Chr(34)+fsName+Chr(34)+vbNewLine
	historyString = historyString + "	.Vector "+Chr(34)+"SourceOriginX"+Chr(34)+","+Chr(34)+"SourceOriginY"+Chr(34)+","+Chr(34)+"SourceOriginZ"+Chr(34)+vbNewLine
	historyString = historyString + "	.Transform "+Chr(34)+"CurrentDistribution"+Chr(34)+","+Chr(34)+"Translate"+Chr(34)+vbNewLine
	historyString = historyString + "End With"+vbNewLine

	' Reenable WCS
	If (wcsActive) Then historyString = historyString + "WCS.ActivateWCS(" + Chr(34) + "local" + Chr(34) + ")" + vbNewLine

	AddToHistory("define Hertzian Dipole: " + fsName, historyString)

	If dFmin > 0 Then
		' Set excitation to Gaussian sine for broadband excitation with fmin > 0
		historyString = ""
		AppendHistoryLine_LIB(historyString, "With TimeSignal")
	    AppendHistoryLine_LIB(historyString, " .Reset")
	    AppendHistoryLine_LIB(historyString, " .Name", "default")
	    AppendHistoryLine_LIB(historyString, " .SignalType", "Gaussian sine")
	    AppendHistoryLine_LIB(historyString, " .ProblemType", "High Frequency")
	    AppendHistoryLine_LIB(historyString, " .Fmin", IIf(nFreqSamples > 1, dFmin/fUnit, Solver.GetFMin))
	    AppendHistoryLine_LIB(historyString, " .Fmax", IIf(nFreqSamples > 1, dFmax/fUnit, Solver.GetFMax))
	    AppendHistoryLine_LIB(historyString, " .Create")
		AppendHistoryLine_LIB(historyString, "End With")
		AddToHistory("define excitation signal: default", historyString)
	End If

	CreateSourceFile = 0

End Function

Function AddFieldMonitors(dFmin As Double, dFmax As Double, iFreqSamples As Long, _
							bAddEMonitors As Boolean, bAddHMonitors As Boolean, bAddFFMonitors As Boolean) As Integer

	Dim sHistoryString As String
	Dim iFreqSample As Long, dFrequency As Double, dLambda As Double

	'Name field source based on either frequency or wavelength
	Dim sfsName As String

	' Define e and h field monitors at proper frequency
	If bAddEMonitors Then
		For iFreqSample = 0 To iFreqSamples-1

			If iFreqSamples > 1 Then
				dFrequency = dFmax - iFreqSample*(dFmax-dFmin)/(iFreqSamples-1)
			Else
				dFrequency = dFmin
			End If

			If DlgValue("WavelengthOrFrequencyDLB") = 0 Then
				sfsName = "(wl="+USFormat(CLight/dFrequency/lUnit,"###0.000")+lUnitS+")"
			Else
				sfsName = "(f="+USFormat(CStr(dFrequency/fUnit),"###0.000")+fUnitS+")"
			End If

			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"e-field "+sfsName+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Efield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define monitor: e-field " + sfsName, sHistoryString)
		Next iFreqSample
	End If

	If bAddHMonitors Then
		For iFreqSample = 0 To iFreqSamples-1

			If iFreqSamples > 1 Then
				dFrequency = dFmax - iFreqSample*(dFmax-dFmin)/(iFreqSamples-1)
			Else
				dFrequency = dFmin
			End If

			If DlgValue("WavelengthOrFrequencyDLB") = 0 Then
				sfsName = "(wl="+USFormat(CLight/dFrequency/lUnit,"###0.000")+lUnitS+")"
			Else
				sfsName = "(f="+USFormat(CStr(dFrequency/fUnit),"###0.000")+fUnitS+")"
			End If

			dLambda = CLight/dFrequency
			sHistoryString = ""
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"h-field "+sfsName+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Dimension "+Chr(34)+"Volume"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Hfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define monitor: h-field " + sfsName, sHistoryString)
		Next iFreqSample
	End If

	If bAddFFMonitors Then
		For iFreqSample = 0 To iFreqSamples-1

			If iFreqSamples > 1 Then
				dFrequency = dFmax - iFreqSample*(dFmax-dFmin)/(iFreqSamples-1)
			Else
				dFrequency = dFmin
			End If

			If DlgValue("WavelengthOrFrequencyDLB") = 0 Then
				sfsName = "(wl="+USFormat(CLight/dFrequency/lUnit,"###0.000")+lUnitS+")"
			Else
				sfsName = "(f="+USFormat(CStr(dFrequency/fUnit),"###0.000")+fUnitS+")"
			End If

			dLambda = CLight/dFrequency
			sHistoryString = sHistoryString + "With Monitor"+vbNewLine
			sHistoryString = sHistoryString + "     .Reset"+vbNewLine
			sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield "+sfsName+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
			sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
			sHistoryString = sHistoryString + "     .Create"+vbNewLine
			sHistoryString = sHistoryString + "End With"+vbNewLine
			AddToHistory("define farfield monitor: farfield " + sfsName, sHistoryString)
		Next iFreqSample
	End If

	' FSC/WCH 03/2017: Do not use broadband FF monitor
	'If bAddFFMonitors Then
	'	sHistoryString = ""
	'	If (iFreqSamples = 1) Then
	'		dFrequency = dFmin
	'		dLambda = CLight/dFrequency
	'		sHistoryString = sHistoryString + "With Monitor"+vbNewLine
	'		sHistoryString = sHistoryString + "     .Reset"+vbNewLine
	'		sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield "+sfsName+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Frequency"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Frequency "+cstr(dFrequency/fUnit)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Create"+vbNewLine
	'		sHistoryString = sHistoryString + "End With"+vbNewLine
	'		AddToHistory("define farfield monitor: farfield " + sfsName, sHistoryString)
	'	ElseIf (Solver.GetFmax > 0) Then
	'		sHistoryString = sHistoryString + "With Monitor"+vbNewLine
	'		sHistoryString = sHistoryString + "     .Reset"+vbNewLine
	'		sHistoryString = sHistoryString + "     .Name "+Chr(34)+"farfield (broadband)"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Domain "+Chr(34)+"Time"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .FieldType "+Chr(34)+"Farfield"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Accuracy "+Chr(34)+"1e-3"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .FrequencySamples "+Chr(34)+CStr(2*iFreqSamples+1)+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Frequency "+cstr((dFmax+dFmin)/2/fUnit)+vbNewLine
	'		sHistoryString = sHistoryString + "     .TransientFarfield "+Chr(34)+"False"+Chr(34)+vbNewLine
	'		sHistoryString = sHistoryString + "     .Create"+vbNewLine
	'		sHistoryString = sHistoryString + "End With"+vbNewLine
	'		AddToHistory("define farfield monitor: farfield (broadband)", sHistoryString)
	'	Else
	'		ReportWarningToWindow("Could not create broadband farfield monitor due to dFmax = 0, please increase dFmax.")
	'	End If
	'End If
End Function

Function mycstr(aText As Double) As String
	mycstr=USFormat(aText, "0.##")
End Function

Function generateHDSCurrent_FieldData(ByVal x0 As Double, ByVal y0 As Double, ByVal z0 As Double, _
						fMin As Double, fMax As Double, fSamp As Long, _
						AmpT As Double, EpsR As Double, MuR As Double, _
	                    DipoleType As Integer) As Integer

	Dim Freq As Double, dFreq As Double
	If (fSamp = 1) Then
		dFreq = fMax
	Else
		dFreq = (fMax - fMin)/(fSamp-1)
	End If

	Dim sDipoleType(1) As String
		sDipoleType(0) = "Electric_Dipole"
		sDipoleType(1) = "Magnetic_Dipole"

	Dim dl As Double, dw As Double
	'Assuming the dipole length d is lamdbda_min/10 for broadband source, current width is also lambda_min/100
	dl  = CLight/fMax/10 'dipole length
	dw  = CLight/fMax/100 'sheet current width - field is aligned in this direction

	generateXmlFile(0,0,0,fMin,fMax, fSamp, dl, dw, DipoleType)

	Dim streamNum As Long, fn As Long, ii As Long

	streamNum = FreeFile
	Open GetProjectPath("TempDS")+"HD_"+sDipoleType(DipoleType)+"_FieldData.dat" For Output As #streamNum

    Dim real_Field As Double, imag_Field As Double


	For ii = 0 To 3 'calculate field source at 4 points
		For fn = 0 To fSamp-1
			If dFreq = 0 Then
				Freq = fMax
			Else
				Freq = fMax - fn*dFreq
			End If
			real_Field = FieldValueforDipoleSource(Freq, AmpT, EpsR, MuR, dl, dw, DipoleType)
    		imag_Field = 0
			Print #streamNum, USFormat(real_Field, SxmlFormat)+" "+USFormat(imag_Field, SxmlFormat)+IIf(fn < fSamp-1, " ", "");
		Next fn
		'next line
		Print #streamNum
	Next ii

	Close #streamNum

	generateHDSCurrent_FieldData = 0

End Function

Sub generateXmlFile(x0 As Double, y0 As Double, z0 As Double, _
					fMin As Double, fMax As Double, fSamp As Double, _
					dl As Double, dw As Double, DipoleType As Integer )

	Dim sFace As String, sField As String, dx As Double, dy As Double, dz As Double

	' By default, the dipole orientation is in Z direction. The source rotation according to the GUI is handled at the end of the code! 
		dx    = 0
		dy    = dw
		dz    = dl
		sFace = "y"
	
	If DipoleType = 0 Then
		sField    = "H"
	Else
		sField    = "E"
	End If

	Dim streamNum As Long, fn As Long

	streamNum = FreeFile

	Dim sDipoleType(1) As String
	sDipoleType(0) = "Electric_Dipole"
	sDipoleType(1) = "Magnetic_Dipole"

	Open GetProjectPath("TempDS")+"HDMacro_"+sDipoleType(DipoleType)+".xml" For Output As #streamNum
		Print #streamNum, "<?xml version="+Chr(34)+"1.0"+Chr(34)+" encoding="+Chr(34)+"UTF-8"+Chr(34)+"?>"
		Print #streamNum, "<EmissionScan>"
		Print #streamNum, vbTab+"<Nfs_ver>1.0</Nfs_ver>"
		Print #streamNum, vbTab+"<Filename>HDMacro_"+sField+"_"+sFace+".xml</Filename>"
		Print #streamNum, vbTab+"<File_ver>1</File_ver>"
		Print #streamNum, vbTab+"<Probe>"
		Print #streamNum, vbTab+vbTab+"<Field>"+sField+sFace+"</Field>"
		Print #streamNum, vbTab+"</Probe>"
		Print #streamNum, vbTab+"<Data>"
		Print #streamNum, vbTab+vbTab+"<Coordinates>none</Coordinates>"
		' XML file coordinate data will be in meters
		Print #streamNum, vbTab+vbTab+"<X0>"+Usformat((x0-dx/2), SxmlFormat)+"m</X0>"
		If (dx > 0) Then
			Print #streamNum, vbTab+vbTab+"<Xstep>"+Usformat(dx, SxmlFormat)+"m</Xstep>"
			Print #streamNum, vbTab+vbTab+"<Xmax>"+Usformat((x0 + dx/2), SxmlFormat)+"m</Xmax>"
		End If
		Print #streamNum, vbTab+vbTab+"<Y0>"+Usformat((y0-dy/2), SxmlFormat)+"m</Y0>"
		If (dy > 0) Then
			Print #streamNum, vbTab+vbTab+"<Ystep>"+Usformat(dy, SxmlFormat)+"m</Ystep>"
			Print #streamNum, vbTab+vbTab+"<Ymax>"+Usformat((y0 + dy/2), SxmlFormat)+"m</Ymax>"
		End If
		Print #streamNum, vbTab+vbTab+"<Z0>"+Usformat((z0-dz/2), SxmlFormat)+"m</Z0>"
		If (dz > 0) Then
			Print #streamNum, vbTab+vbTab+"<Zstep>"+Usformat(dz, SxmlFormat)+"m</Zstep>"
			Print #streamNum, vbTab+vbTab+"<Zmax>"+Usformat((z0 + dz/2), SxmlFormat)+"m</Zmax>"
		End If
		Print #streamNum, vbTab+vbTab+"<Frequencies>"
		Print #streamNum, vbTab+vbTab+vbTab+"<List>";
		For fn = 0 To fSamp - 1
			If fSamp > 1 Then
				Print #streamNum, USFormat(fMax - fn*(fMax-fMin)/(fSamp-1), SxmlFormat)+IIf(fn < fSamp-1, " ", "");
			Else
				Print #streamNum, USFormat(fMin, SxmlFormat);
			End If
		Next fn
		Print #streamNum, "</List>"
		Print #streamNum, vbTab+vbTab+"</Frequencies>"
		Print #streamNum, vbTab+vbTab+"<Measurement>"
		Print #streamNum, vbTab+vbTab+vbTab+"<Unit>V/m</Unit>"
		Print #streamNum, vbTab+vbTab+vbTab+"<Format>ri</Format>"
		Print #streamNum, vbTab+vbTab+vbTab+"<Data_files>HD_"+sDipoleType(DipoleType)+"_FieldData.dat</Data_files>"
		Print #streamNum, vbTab+vbTab+"</Measurement>"
		Print #streamNum, vbTab+"</Data>"
		Print #streamNum, "</EmissionScan>"
	Close #streamNum
End Sub

Function FieldValueforDipoleSource(Freq As Double, AmpT As Double, EpsR As Double, MuR As Double, _
									dl As Double, dw As Double, DipoleType As Integer) As Double

	Dim lambda As Double, k_vector As Double, Zimp As Double, I_dipole As Double

	Zimp = (mu0*MuR/eps0*EpsR)^0.5

	Dim Js As Double, FieldValue As Double, Pow_av As Double
	Pow_av = AmpT/2 'Convert the input peak power to ave power

	Dim ii As Integer, n_bg As Double

	lambda     = Clight/Freq

	'refractive_index_bg
	n_bg       = (EpsR*MuR)^0.5

	k_vector   = 2*Pi/lambda*n_bg

	I_dipole   = (Pow_av*48*Pi/Zimp)^0.5/k_vector/dl

	Js         = I_dipole/dw


 	Select Case  DlgText("PowerOrEFieldOrMomentDLB")
 	Case "Power:"
		If DipoleType = 0 Then
		'electric dipole
			FieldValueforDipoleSource = Js/2*(EpsR)^0.5
		Else
			FieldValueforDipoleSource = Js*Zimp/2/(EpsR)^0.5
		End If

	Case "E-Field at 1 meter:"
		If DipoleType = 0 Then
		'electric dipole
			FieldValueforDipoleSource = AmpT*4*pi/Zimp/k_vector/dl/dw
		End If

	Case "H-Field at 1 meter:"
		If DipoleType = 1 Then
		'electric dipole
			FieldValueforDipoleSource = AmpT*4*pi*Zimp/k_vector/dl/dw
		End If

	Case "Dipole moment:"
		If DipoleType = 0 Then
		'electric dipole
			FieldValueforDipoleSource = AmpT/dl/dw
		Else
			FieldValueforDipoleSource = AmpT/dl/dw
		End If
	End Select

End Function
