' This macro calculates the parameters to set up a Drude material from Gas type and pressure
'
' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' --------------------------------------------------------------------------------------------------------------------
' 12-Jul-2011 fsr: Initial version
' --------------------------------------------------------------------------------------------------------------------
'

Option Explicit

Public Const ElectronMass = 9.10938188e-31
Public Const ElectronCharge = -1.60217646e-19
Public Const kBoltzmann = 1.3806503e-23
' Relative polarizabilities for some atoms and molecules, from Smirnov 1981 via Lieberman Table 3.2
Public Const Polarizabilities = Array("H|4.5", _
										"C|12", _
										"N|7.5", _
										"O|5.4", _
										"Ar|11.08", _
										"CCl4|69", _
										"CF4|19", _
										"CO|13.2", _
										"CO2|17.5", _
										"Cl2|31", _
										"H2O|9.8", _
										"NH3|14.8", _
										"O2|10.6", _
										"SF6|30")

Public Const PressureUnitsArray = Array("(Pa):", "(bar):", "(Torr):")
Public Const PressureUnitsToSI = Array(1, 1e5, 133.322)

Sub Main

	Dim i As Long
	Dim sGasTypeArray() As String

	ReDim sGasTypeArray(UBound(Polarizabilities))
	For i = 0 To UBound(Polarizabilities)
		sGasTypeArray(i) = Split(Polarizabilities(i),"|")(0)
	Next

	Begin Dialog UserDialog 460,455,"Drude Material for Plasma Applications",.DlgFunc ' %GRID:10,7,1,1
		GroupBox 10,7,440,210,"Main settings",.GroupBox1
		PushButton 240,427,100,21,"Create",.CreatePB
		GroupBox 10,364,440,56,"Calculate collision rate constant",.GroupBox3
		GroupBox 10,224,440,133,"Calculate neutral and electron densities",.GroupBox2
		Text 30,280,140,14,"Gas temperature (K):",.Text4
		Text 30,308,220,14,"Degree of ionization (0 ... 1):",.Text5
		TextBox 320,245,100,21,.GasPressureT
		DropListBox 120,245,70,192,PressureUnitsArray(),.PressureUnitsDLB
		Text 30,154,220,14,"e-n collision rate constant (m^3/s):",.Text7
		TextBox 320,147,100,21,.CollisionRateT
		TextBox 320,119,100,21,.ElectronDensityT
		TextBox 320,91,100,21,.NeutralDensityT
		Text 30,126,210,14,"Electron density n_e (1/m^3):",.Text1
		Text 20,56,410,28,"The following three parameters are all that is needed. If available, enter their values here and use 'Create' to create the material.",.Text8
		Text 20,182,410,28,"If NOT available, you may calculate/estimate their values using the boxes below:",.Text9
		Text 30,98,210,14,"Neutral density n_n (1/m^3):",.Text10
		Text 30,252,90,14,"Gas pressure",.Text3
		TextBox 320,273,100,21,.GasTempT
		PushButton 320,329,100,21,"Calculate",.CalculateDensitiesPB
		TextBox 320,301,100,21,.IonizationDegreeT
		Text 30,392,180,14,"Select primary gas species:",.Text2
		DropListBox 320,385,100,192,sGasTypeArray(),.GasTypeDLB
		PushButton 350,427,100,21,"Close",.ClosePB
		PushButton 140,427,90,21,"Display only",.DisplayPB
		Text 20,35,100,14,"Material name:",.Text6
		TextBox 130,28,290,21,.MaterialNameT
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Dim dOmegaPe As Double, dCollisionFreq As Double
	Dim dNeutralDensity As Double, dElectronDensity As Double, dCollisionRate As Double

	'ReportInformationToWindow(DlgItem + " : " + CStr(Action))
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("MaterialNameT", "Plasma")
		DlgText("GasPressureT", "1")
		DlgText("GasTempT", "293")
		DlgText("IonizationDegreeT", "0.001")
		CalculateDensities()
		CalculateCollisionRate()
	Case 2 ' Value changing or button pressed
		Rem DlgFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "ClosePB"
				Exit All
			Case "CalculateDensitiesPB"
				DlgFunc = True
				CalculateDensities()
			Case "GasTypeDLB"
				DlgFunc = True
				CalculateCollisionRate()
			Case "CreatePB"
				dNeutralDensity = Evaluate(DlgText("NeutralDensityT"))
				dElectronDensity = Evaluate(DlgText("ElectronDensityT"))
				dCollisionRate = Evaluate(DlgText("CollisionRateT"))
				dOmegaPe = Sqr(dElectronDensity*ElectronCharge^2/Eps0/ElectronMass)
				dCollisionFreq = dNeutralDensity*dCollisionRate
				CreatePlasmaMaterial(DlgText("MaterialNameT"), dOmegaPe, dCollisionFreq)
			Case "DisplayPB"
				DlgFunc = True
				dNeutralDensity = Evaluate(DlgText("NeutralDensityT"))
				dElectronDensity = Evaluate(DlgText("ElectronDensityT"))
				dCollisionRate = Evaluate(DlgText("CollisionRateT"))
				dOmegaPe = Sqr(dElectronDensity*ElectronCharge^2/Eps0/ElectronMass)
				dCollisionFreq = dNeutralDensity*dCollisionRate
				MsgBox("Plasma frequency: " + Format(dOmegaPe, "0.000e+00") + " rad/s" + vbNewLine + "Collision frequency: " + Format(dCollisionFreq, "0.000e+00") + " 1/s" + vbNewLine, "Frequency display")
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DlgFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub CalculateDensities()
	DlgText("NeutralDensityT",Format(Evaluate(DlgText("GasPressureT"))*PressureUnitsToSI(DlgValue("PressureUnitsDLB"))/Evaluate(DlgText("GasTempT"))/kBoltzmann,"0.000e+00"))
	DlgText("ElectronDensityT",Format(Evaluate(DlgText("IonizationDegreeT"))*Evaluate(DlgText("GasPressureT"))*PressureUnitsToSI(DlgValue("PressureUnitsDLB"))/Evaluate(DlgText("GasTempT"))/kBoltzmann,"0.000e+00"))
End Sub

Sub CalculateCollisionRate()
	DlgText("CollisionRateT", Format(3.85e-14*Sqr(Evaluate(Split(Polarizabilities(DlgValue("GasTypeDLB")),"|")(1))), "0.000e+00"))
End Sub

Function CreatePlasmaMaterial(sMaterialName As String, dOmegaPe As Double, dCollisionFreq As Double) As Integer

	Dim sHistoryString As String

	CreatePlasmaMaterial = 0 ' Error

	sHistoryString = ""

	sHistoryString = sHistoryString + "With Material" + vbNewLine
	sHistoryString = sHistoryString + "     .Reset" + vbNewLine
	sHistoryString = sHistoryString + "     .Name "+Chr(34)+sMaterialName+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .FrqType "+Chr(34)+"All"+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .Type "+Chr(34)+"Normal"+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .DispModelEps  "+Chr(34)+"Drude"+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .EpsInfinity "+Chr(34)+"1.0"+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .DispCoeff1Eps "+Chr(34)+Format(dOmegaPe,"0.000e+00")+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .DispCoeff2Eps "+Chr(34)+Format(dCollisionFreq,"0.000e+00")+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .Colour "+Chr(34)+"0"+Chr(34)+", "+Chr(34)+"1"+Chr(34)+", "+Chr(34)+"1"+Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .Create" + vbNewLine
	sHistoryString = sHistoryString + "End With"

	AddToHistory("define material: " + sMaterialName, sHistoryString)

	CreatePlasmaMaterial = 1 ' All OK

End Function
