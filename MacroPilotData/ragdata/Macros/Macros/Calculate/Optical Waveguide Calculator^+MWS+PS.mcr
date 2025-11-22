'#Language "WWB-COM"

'----------------------------------------------------------------------------------------------------------------------------------------------
' This macro builds and calculates mode data for optical waveguides
'----------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 09-Aug-2017 fsr: Fixed 3D construction of asymmetric slab waveguide
' 28-Sep-2015 fsr: Finishing touches to prepare for distribution
' 28-Jul-2015 jfl: initial version
'----------------------------------------------------------------------------------------------------------------------------------------------


Option Explicit

Dim sMacroPath As String
Dim sMaterials() As String
Dim iMaterialIndex As Integer

Dim bRestartFlag As Boolean ' Used to refresh the dialog after results are calculated
Dim bFirstOpenFlag As Boolean ' Used to prevent default values from being loaded when dialog refreshes
Dim sLastFreqWave As String ' Stores the last state of the WaveFreq drop list
Dim Results ' Stores the results to be displayed on refresh

Type Solution
	dBeta As Double
	iM As Integer
	iNn As Integer
	sType As String
End Type

Sub Main
	Dim WaveguideSelect, WaveFreq
	WaveguideSelect = Array( "Buried Channel Waveguide", "Asymmetrical Slab Waveguide", "Strip-Loaded Waveguide", _
							 "Ridge Waveguide", "Rib Waveguide", "Diffused Waveguide", "Cylindrical Optical Fiber")
	WaveFreq = Array("Wavelength", "Frequency")
	Results = Array("|  Mode  |        neff        |        theta        |         beta         |")
	Dim i As Integer
	ReDim sMaterials(Material.GetNumberOfMaterials-1)

	For i = 1 To UBound(sMaterials)
		sMaterials(i) = Material.GetNameOfMaterialFromIndex(i)
	Next

	bFirstOpenFlag = True
	sLastFreqWave = "Wavelength"
	sMacroPath = GetInstallPath + "\Library\Macros\Calculate\"
	Begin Dialog UserDialog 940,490,"Optical Waveguide Calculator",.DialogFunc ' %GRID:10,7,1,1
		Picture 10,7,620,315,"",0,.Picture
		GroupBox 10,329,760,154,"Results:",.GbResults
		GroupBox 640,28,290,245,"Structure Parameters:",.GbStructure
		TextBox 690,49,90,21,.TbW1
		TextBox 830,49,90,21,.TbH1
		TextBox 690,84,90,21,.TbW2
		TextBox 830,84,90,21,.TbH2
		TextBox 690,119,90,21,.TbL
		TextBox 830,119,90,21,.TbH3
		TextBox 690,154,90,21,.Tbn1
		TextBox 690,189,90,21,.Tbn2
		TextBox 690,224,90,21,.Tbn3
		Text 650,49,30,14,"W1:",.LbW1
		Text 660,154,20,14,"n1:",.Lbn1
		Text 660,189,20,14,"n2:",.Lbn2
		Text 660,224,20,14,"n3:",.Lbn3
		Text 790,49,30,14,"H1:",.LbH1
		Text 790,84,30,14,"H2:",.LbH2
		Text 790,119,30,14,"H3:",.LbH3
		Text 650,84,30,14,"W2:",.LbW2
		DropListBox 640,7,290,121,WaveguideSelect(),.DlWaveguideSelect
		PushButton 800,154,100,21,"Use Material",.PbMaterial0
		PushButton 800,189,100,21,"Use Material",.PbMaterial1
		PushButton 800,224,100,21,"Use Material",.PbMaterial2
		Text 660,119,20,14,"L:",.LbL
		Text 650,252,270,14,"Unknown units",.TbUnits,2
		GroupBox 640,280,280,49,"Calculation Parameters:",.Gbcalculate
		TextBox 790,301,110,21,.TbWavelength
		Text 30,385,30,14,"NA:",.LbNA
		Text 30,420,30,14,"V:",.LbV
		TextBox 60,378,90,21,.TbNA
		TextBox 60,413,90,21,.TbV
		Text 340,329,80,14,"Mode Data:",.LbModeData
		ListBox 340,343,390,126,Results(),.LbResults,1
		Text 30,455,180,14,"Calculating... [||||          ]",.LbCalculating
		Text 170,420,50,14,"Modes:",.LbModes
		TextBox 220,413,90,21,.TbModes
		DropListBox 660,301,110,121,WaveFreq(),.DlWaveFreq
		PushButton 180,343,130,21,"Export Mode Data",.PbExport
		TextBox 220,378,90,21,.TbDelta
		Text 170,385,40,14,"delta:",.LbDelta
		PushButton 220,448,90,21,"Abort",.PbAbort

		PushButton 820,343,90,21,"Build 3D",.PbBuild
		PushButton 820,378,90,21,"Calculate",.PbCalculate
		PushButton 820,413,90,21,"Help",.PbHelp
		PushButton 820,448,90,21,"Close",.PbClose

		OKButton 0,0,0,0
	End Dialog
	Dim dlg As UserDialog
	Do
		bRestartFlag = False
		Dialog dlg
	Loop While (bRestartFlag)

End Sub

Function DialogFunc(DlgItem As String, Action As Integer, SuppValue As Integer) As Integer
	Select Case Action%
    Case 1, 2
    	If (Action% = 1) Then
    		DlgEnable("PbExport", UBound(Results)>0)
    		DlgEnable("LbResults", UBound(Results)>0)
    		DlgEnable("PbAbort",False)

' To-Do: Remove once help file exists
' ------------------------------------------------------------------------------------------------------
    		DlgEnable("PbHelp",False)
' ------------------------------------------------------------------------------------------------------

			DlgText("TbUnits", "Geometry unit: " + Units.GetUnit("Length") + " | Frequency unit: "+ Units.GetUnit("Frequency"))
			If (Left(DlgText("TbModes"),2) = ">=") Then
				DlgText("LbCalculating","Aborted! Mode data is incomplete!")
			Else
				DlgText("LbCalculating","")
			End If
			DlgItem = "DlWaveguideSelect"
			If (bFirstOpenFlag) Then
				bFirstOpenFlag = False

				DlgText("TbW1", CStr(500e-9*Units.GetGeometrySIToUnit))
				DlgText("Tbn1", "1.87")
				DlgText("Tbn2", "1.45")
				DlgText("Tbn3", "1")
				DlgText("TbH1", CStr(100e-9*Units.GetGeometrySIToUnit))
				DlgText("TbH2", CStr(2000e-9*Units.GetGeometrySIToUnit))
				DlgText("TbH3", CStr(100e-9*Units.GetGeometrySIToUnit))
				DlgText("TbW2", CStr(2500e-9*Units.GetGeometrySIToUnit))
				DlgText("TbL", CStr(10000e-9*Units.GetGeometrySIToUnit))
				If (Solver.GetFMax > 0) Then
					DlgText("TbWavelength", CInt(cLight/((Solver.GetFMax + Solver.GetFMin)/2 * Units.GetFrequencyUnitToSI) * Units.GetGeometrySIToUnit))
				Else
					DlgText("TbWavelength", CStr(1550e-6*Units.GetGeometrySIToUnit))
				End If
			End If
    	End If
		Select Case DlgItem
		Case "DlWaveguideSelect"
			DlgSetPicture "Picture", sMacroPath+"OptWG-"+DlgText("DlWaveguideSelect")+".bmp", 0
			Select Case DlgText("DlWaveguideSelect")
			Case "Symmetrical Slab Waveguide"
				DlgEnable("TbH1", True)
				DlgEnable("TbH2", True)
				DlgEnable("TbH3", False)
				DlgEnable("TbW2", False)
				DlgEnable("Tbn3", False)
				DlgEnable("PbMaterial2", False)
				DlgEnable("PbCalculate", True)
			Case "Asymmetrical Slab Waveguide"
				DlgEnable("TbH1", True)
				DlgEnable("TbH2", True)
				DlgEnable("TbH3", False)
				DlgEnable("TbW2", False)
				DlgEnable("Tbn3", True)
				DlgEnable("PbMaterial2", True)
				DlgEnable("PbCalculate", True)
			Case "Cylindrical Optical Fiber"
				DlgEnable("TbH1", False)
				DlgEnable("TbH2", False)
				DlgEnable("TbH3", False)
				DlgEnable("TbW2", True)
				DlgEnable("Tbn3", False)
				DlgEnable("PbMaterial2", False)
				DlgEnable("PbCalculate", True)
			Case "Buried Channel Waveguide", "Ridge Waveguide", "Diffused Waveguide"
				DlgEnable("TbH1", True)
				DlgEnable("TbH2", True)
				DlgEnable("TbH3", False)
				DlgEnable("TbW2", True)
				DlgEnable("Tbn3", False)
				DlgEnable("PbMaterial2", False)
				DlgEnable("PbCalculate", False)
			Case "Strip-Loaded Waveguide"
				DlgEnable("TbH1", True)
				DlgEnable("TbH2", True)
				DlgEnable("TbH3", True)
				DlgEnable("TbW2", True)
				DlgEnable("Tbn3", True)
				DlgEnable("PbMaterial2", True)
				DlgEnable("PbCalculate", False)
			Case "Rib Waveguide"
				DlgEnable("TbH1", True)
				DlgEnable("TbH2", True)
				DlgEnable("TbH3", True)
				DlgEnable("TbW2", True)
				DlgEnable("Tbn3", False)
				DlgEnable("PbMaterial2", False)
				DlgEnable("PbCalculate", False)
			End Select
		Case "PbMaterial0", "PbMaterial1", "PbMaterial2"
			DialogFunc = True
	Begin Dialog UserDialog 250,238,"Select Material",.MaterialSelectFunc ' %GRID:10,7,1,1
		ListBox 10,7,230,196,sMaterials(),.MlbMaterial
		OKButton 50,210,90,21
		CancelButton 150,210,90,21
	End Dialog
			Dim dlg As UserDialog
			If (Dialog(dlg) = 0) Then DlgItem = ""
			If (iMaterialIndex = -1) Then DlgItem = ""

			Select Case DlgItem
			Case "PbMaterial0"
				DlgText("Tbn1", sMaterials(iMaterialIndex))
			Case "PbMaterial1"
				DlgText("Tbn2", sMaterials(iMaterialIndex))
			Case "PbMaterial2"
				DlgText("Tbn3", sMaterials(iMaterialIndex))
			End Select
		Case "PbBuild"
			On Error GoTo Catch ' Try
			If (buildWaveguide(DlgText("DlWaveguideSelect"), DlgText("Tbn1"), DlgText("Tbn2"), DlgText("Tbn3"), _
						myEval(DlgText("TbW1")), myEval(DlgText("TbW2")), myEval(DlgText("TbH1")), myEval(DlgText("TbH2")), _
						myEval(DlgText("TbH3")), myEval(DlgText("TbL"))) = 0) Then
				DialogFunc = False
			Else
				DialogFunc = True
			End If
			On Error GoTo 0 ' End Try
		Case "PbCalculate"
			Dim dN1 As Double, dN2 As Double, dN3 As Double, dD As Double, dLambda As Double, dNA As Double, dV As Double, dW0 As Double, dMFD As Double
			Dim iIterator As Integer
			Dim dTMThetas() As Double, dTEThetas() As Double, sSolutions() As Solution
			dN1 = GetNValue(DlgText("Tbn1"))
			dN2 = GetNValue(DlgText("Tbn2"))
			If (dN1 <= 0 Or dN2 <= 0) Then
				MsgBox("You have entered a zero valued refractive index.", "Invalid Input")
				DialogFunc = True
				Exit Function
			End If
			If (dN2 >= dN1) Then
				MsgBox("n1 (" + Str(dN1) + ") must be greater than n2 ("+ Str(dN2) +") to guide light.", "Invalid Input")
				DialogFunc = True
				Exit Function
			End If
			If (Not IsNumeric(DlgText("TbWavelength"))) Then
				MsgBox("Field for "+DlgText("DlWaveFreq")+" is invalid.", vbOkOnly, "Invalid Input")
				DialogFunc = True
				Exit Function
			End If
			If (DlgText("DlWaveFreq") = "Wavelength") Then
				dLambda = Eval(DlgText("TbWavelength")) * Units.GetGeometryUnitToSI
			Else
				If (Eval(DlgText("TbWavelength")) <= 0) Then
					MsgBox("Field for "+DlgText("DlWaveFreq")+" is invalid.", vbOkOnly, "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				dLambda = cLight/(Eval(DlgText("TbWavelength")) * Units.GetFrequencyUnitToSI)
			End If
			If (dLambda <= 0) Then
				MsgBox("Field for "+DlgText("DlWaveFreq")+" is invalid.", vbOkOnly, "Invalid Input")
				DialogFunc = True
				Exit Function
			End If
			dNA = Sqr(dN1^2-dN2^2)
			DlgText("TbNA", Format(dNA, "0.00000"))
			If (DlgText("DlWaveguideSelect")="Asymmetrical Slab Waveguide") Then

				dN3 = GetNValue(DlgText("Tbn3"))
				If (dN3 <= 0) Then
					MsgBox("You have entered a zero valued refractive index.", "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				If (dN3 >= dN1) Then
					MsgBox("n1 (" + Str(dN1) + ") must be greater than n3 ("+ Str(dN3) +") to guide light.", "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				If (dN3 > dN2) Then
					MsgBox("It is conventional for the index in the cover (n3 = " + Str(dN3) + ") to be no greater than the index in the substrate (n2 = " + _
							Str(dN2) + "). These two values will now be swapped.", vbOkOnly, "n3 > n2")
					Dim sTemp As String, dTemp As Double
					sTemp = DlgText("Tbn2")
					DlgText("Tbn2", DlgText("Tbn3"))
					DlgText("Tbn3", sTemp)
					dTemp = dN2
					dN2 = dN3
					dN3 = dTemp
				End If
				If (Not IsNumeric(DlgText("TbH1"))) Then
					MsgBox("Field for H1 is invalid.", vbOkOnly, "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				dD = Eval(DlgText("TbH1")) * Units.GetGeometryUnitToSI
				If (dD <= 0) Then
					MsgBox("Field for H1 is invalid.", vbOkOnly, "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				dV = 2*pi*dD*dNA/dLambda
				DlgText("TbV", Format(dV, "0.00000"))
				DlgText("TbDelta", Format((dN2^2-dN3^2)/(dN1^2-dN2^2), "0.00000"))
				sSolutions = FindModesAsymmetricSlab(dN1, dN2, dN3, 2*pi/dLambda, dD/2)
				ReDim Preserve Results(0) ' Clear old results
				If (sSolutions(0).dBeta > 0) Then ReDim Preserve Results(UBound(sSolutions)+1)
				DlgText("TbModes", Format(UBound(Results), "0"))

				For iIterator = 0 To UBound(Results)-1
					Results(iIterator+1) = "  " + Trim(sSolutions(iIterator).sType) + Trim(Str(sSolutions(iIterator).iM)) + _
					vbTab + Format(sSolutions(iIterator).dBeta/2/pi*dLambda, "0.000000") + "        " + _
					Format(180/pi*Atn(sSolutions(iIterator).dBeta/dN1/2/pi*dLambda/Sqr(1-(sSolutions(iIterator).dBeta/dN1/2/pi*dLambda)^2)), "0.000000°") + _
					"        " + Format(((sSolutions(iIterator).dBeta/2/pi*dLambda)^2 - dN2^2)/(dN1^2-dN2^2), "0.000000")
				Next iIterator
			Else
				If (Not IsNumeric(DlgText("TbW1"))) Then
					MsgBox("Field for W1 is invalid.", vbOkOnly, "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				dD = Eval(DlgText("TbW1"))/2 * Units.GetGeometryUnitToSI
				If (dD <= 0) Then
					MsgBox("Field for W1 is invalid.", vbOkOnly, "Invalid Input")
					DialogFunc = True
					Exit Function
				End If
				dV = 2*pi*dD*dNA/dLambda
				DlgText("TbV", Format(dV, "0.00000"))
				DlgText("TbDelta", "N/A")
				sSolutions = FindModesCylidricalOptical(dN1, dN2, 2*pi/dLambda, dD)
				ReDim Preserve Results(0) ' Clear old results
				If (sSolutions(0).dBeta > 0) Then ReDim Preserve Results(UBound(sSolutions)+1)

				For iIterator = 0 To UBound(Results)-1
					Results(iIterator+1) = " " + Trim(sSolutions(iIterator).sType) + Trim(Str(sSolutions(iIterator).iM)) +","+ _
					Trim(Str(sSolutions(iIterator).iNn)) + vbTab + Format(sSolutions(iIterator).dBeta/2/pi*dLambda, "0.000000") + "        " + _
					Format(180/pi*Atn(sSolutions(iIterator).dBeta/dN1/2/pi*dLambda/Sqr(1-(sSolutions(iIterator).dBeta/dN1/2/pi*dLambda)^2)), "0.000000°") + _
					"        " + Format((sSolutions(iIterator).dBeta/2/pi*dLambda - dN2)/(dN1-dN2), "0.000000")
				Next iIterator
			End If
			bRestartFlag = True
		Case "LbResults", "PbExport"
			Dim sFileName As String, sLine As String
			Dim iStreamNum As Integer
			sFileName = GetFilePath(DlgText("DlWaveguideSelect"),"txt",GetProjectPath("Root"), "Export Results", 3)
			If (Len(sFileName) > 0) Then
				iStreamNum = FreeFile
				Open sFileName For Output As #iStreamNum
					Print #iStreamNum, "Mode"+vbTab+"neff"+vbTab+vbTab+"theta"+vbTab+vbTab+"b"
					For iIterator = 1 To UBound(Results)
						sLine = Results(iIterator)
						While (InStr(1, sLine, "  "))
							sLine = Replace(sLine, "  ", " ")
						Wend
						sLine = Trim(sLine)
						sLine = Replace(sLine, " ", vbTab)
						Print #iStreamNum, sLine
					Next iIterator
				Close #iStreamNum
			End If
			DialogFunc = True
		Case "PbClose"
			DialogFunc = False

' To-Do: Set PbHelp to open help file
' ------------------------------------------------------------------------------------------------------------------
		Case "PbHelp"
			MsgBox("I don't exitst yet.",vbOkOnly,"No Help File")
			DialogFunc = True
' ------------------------------------------------------------------------------------------------------------------

		Case "DlWaveFreq"
			If (DlgText("DlWaveFreq") <> sLastFreqWave) Then
				sLastFreqWave = DlgText("DlWaveFreq")
				If (Not IsNumeric(DlgText("TbWavelength"))) Then
					Exit Function
				End If
				If (Eval(DlgText("TbWavelength")) = 0) Then
					Exit Function
				End If
				If (sLastFreqWave = "Frequency") Then
					DlgText("TbWavelength", Trim(Str(Units.GetFrequencySIToUnit*cLight/(Eval(DlgText("TbWavelength"))*Units.GetGeometryUnitToSI))))
				Else
					DlgText("TbWavelength", Trim(Str(Units.GetGeometrySIToUnit*cLight/(Eval(DlgText("TbWavelength"))*Units.GetFrequencyUnitToSI))))
				End If
			End If
		End Select

    End Select
    Exit Function

	' General error handling for Build 3D function
Catch:
	MsgBox("Invalid expression. Please check inputs.",vbOkOnly,"Invalid Input")
	On Error GoTo 0
	DialogFunc = True
End Function
Function MaterialSelectFunc(DlgItem As String, Action As Integer, SuppValue As Integer) As Integer
	Select Case Action
	Case 1
		iMaterialIndex = 1
	Case 2
		Select Case DlgItem
		Case "MlbMaterial"
			iMaterialIndex = DlgValue("MlbMaterial")+1
		End Select
	End Select
End Function

Function FindModesAsymmetricSlab(dN1 As Double, dNC As Double, dNS As Double, dK0 As Double, dD As Double) As Solution()
	Dim sSolutions() As Solution
	ReDim sSolutions(0)
	Dim dP As Double, dQ1 As Double, dQ2 As Double
	Dim dBetaMin As Double, dBetaMax As Double
	Dim dDeltaBeta As Double, dBetaGuessTE As Double, dBetaGuessTM As Double
	Dim iM As Double
	Dim bHadSolution As Boolean

	dBetaMin = dK0*IIf(dNC > dNS, dNC, dNS)
	dBetaMax = dK0*dN1
	bHadSolution = True
	While (bHadSolution)
		dBetaGuessTE = (dBetaMax + dBetaMin)/2
		dBetaGuessTM = (dBetaMax + dBetaMin)/2
		dDeltaBeta = (dBetaMax - dBetaMin)/4
		While ((Abs(dDeltaBeta) > 0.0000000000000001))
			On Error GoTo Break ' Truncation errors sometimes cause the functions to be evaluated outside of their domain.
								' Would be better to find max truncation error and adjust dDeltaBeta accordingly.
			dP = Sqr(dN1^2*dK0^2-dBetaGuessTE^2)
			dQ1 = Sqr(dBetaGuessTE^2-dNC^2*dK0^2)
			dQ2 = Sqr(dBetaGuessTE^2-dNS^2*dK0^2)
			dBetaGuessTE = dBetaGuessTE - Sgn(iM*pi+Atn(dQ1/dP)+Atn(dQ2/dP)-2*dD*dP)*dDeltaBeta
			dP = Sqr(dN1^2*dK0^2-dBetaGuessTM^2)
			dQ1 = Sqr(dBetaGuessTM^2-dNC^2*dK0^2)
			dQ2 = Sqr(dBetaGuessTM^2-dNS^2*dK0^2)
			dBetaGuessTM = dBetaGuessTM - Sgn(iM*pi+Atn(dN1^2/dNC^2*dQ1/dP)+Atn(dN1^2/dNS^2*dQ2/dP)-2*dD*dP)*dDeltaBeta
			dDeltaBeta = dDeltaBeta/2
		Wend
Break:
		On Error GoTo 0
		If (dBetaGuessTE = dBetaGuessTM) Then
			bHadSolution = False
		Else
			sSolutions(UBound(sSolutions)).dBeta = dBetaGuessTE
			sSolutions(UBound(sSolutions)).sType = "TE"
			sSolutions(UBound(sSolutions)).iM = iM
			ReDim Preserve sSolutions(UBound(sSolutions) + 1)
			sSolutions(UBound(sSolutions)).dBeta = dBetaGuessTM
			sSolutions(UBound(sSolutions)).sType = "TM"
			sSolutions(UBound(sSolutions)).iM = iM
			ReDim Preserve sSolutions(UBound(sSolutions) + 1)
		End If
		iM = iM+1
	Wend
	If (UBound(sSolutions) > 0) Then
		ReDim Preserve sSolutions(UBound(sSolutions) - 1)
	End If
	FindModesAsymmetricSlab = sSolutions

End Function

Function FindModesCylidricalOptical(dN1 As Double, dN2 As Double, dK0 As Double, dD As Double) As Solution()
	Dim dBetaMin As Double, dBetaMax As Double, dBetaGuess As Double, dBetaLastGuess As Double

	Dim dXi As Double, dEta As Double

	Dim iEHCounter As Integer, iHECounter As Integer, iSolutionCounter As Integer, iNoSolutionCounter As Integer

	Dim dF As Double, dFLast As Double
	Dim iCounter As Integer, iM As Integer
	Dim sSolutions() As Solution
	ReDim sSolutions(0)
	Dim iFSign As Integer
	Dim dJump As Double
	Dim dThreshold As Double
	Dim lProgress As Long

	dThreshold = 1E-10
	dJump = dK0*(dN1-dN2)/30000 ' Take 30000 steps. Roots can be very close together so small partitions are required.
	dFLast = 0
	dBetaMin = dN2*dK0
	dBetaMax = dN1*dK0 - dJump

	DlgText("TbModes", Format(0, "0"))
	DlgEnable("PbAbort",True)

	iNoSolutionCounter = 0
	iM = 0
	While (iNoSolutionCounter < 3) ' Continue until there are no solutions for 3 consecutive iterations
		dBetaGuess = dBetaMax
		iFSign = 0
		iEHCounter = 1
		iHECounter = 1
		iSolutionCounter = 0

		Do
			lProgress = lProgress + 1
			If (lProgress Mod 1501 = 0) Then
				If (lProgress = 2147480700) Then lProgress = 0
				DlgText("LbCalculating", "Calculating... [" + Right("||||          ", lProgress Mod 15) + Left("||||           ", 14 - (lProgress Mod 15)) + "]")
			End If
			If (DlgFocus = "PbAbort") Then
				' User pressed Abort
				DlgEnable("PbAbort",False)
				DlgText("TbModes", ">=" + DlgText("TbModes"))
				DlgText("LbCalculating", "")
				If (UBound(sSolutions) > 0) Then
					ReDim Preserve sSolutions(UBound(sSolutions) - 1)
				End If
				FindModesCylidricalOptical = sSolutions
				Exit Function
			ElseIf (DlgFocus <> "PbCalculate") Then
				DlgFocus("PbCalculate")
			End If

			dF = EvalF(dBetaGuess, dN1, dN2, dK0, dD, iM)

			If (iFSign = 0) Then
				iFSign = Sgn(dF)
			ElseIf (iFSign <> Sgn(dF)) Then
				If (Abs(dF) > 1) Then
					' Assume change in sign is due to an asymptote
					iFSign = -iFSign
				Else
					While (Abs(dF) > dThreshold And Abs(dBetaGuess - dBetaLastGuess) > 1E-15)
						' We are straddling a root. Switch to bisection method
						Dim dTempB As Double, dTempF As Double
						dTempF = EvalF((dBetaGuess + dBetaLastGuess)/2, dN1, dN2, dK0, dD, iM)
						dTempB = dBetaLastGuess
						If (Sgn(dTempF) <> Sgn(dF)) Then
							dBetaLastGuess = dBetaGuess
						End If
						dBetaGuess = (dBetaGuess + dTempB)/2
						dF = dTempF
					Wend
					' We have converged sufficently close to a solution
					iSolutionCounter = iSolutionCounter + 1
					sSolutions(UBound(sSolutions)).dBeta = dBetaGuess
					sSolutions(UBound(sSolutions)).iM = iM
					dXi = dD*Sqr((dN1*dK0)^2 - dBetaGuess^2)
					dEta = dD * Sqr(dBetaGuess^2-dN2^2*dK0^2)
					If (iM = 0) Then
						' For m=0 we have TE and TM modes
						If (Abs(.5*dD*(cyl_Bessel_j(iM-1, dXi)-cyl_Bessel_j(iM+1, dXi))/dXi/cyl_Bessel_j(iM, dXi)+ _
							.5*dD*(-cyl_Bessel_k(iM-1, dEta)-cyl_Bessel_k(iM+1, dEta))/dEta/cyl_Bessel_k(iM, dEta)) > _
							Abs(.5*dD*(cyl_Bessel_j(iM-1, dXi)-cyl_Bessel_j(iM+1, dXi))/dXi/cyl_Bessel_j(iM, dXi)+ _
							(dN2^2/dN1^2)*.5*dD*(-cyl_Bessel_k(iM-1, dEta)-cyl_Bessel_k(iM+1, dEta))/dEta/cyl_Bessel_k(iM, dEta))) Then
							' TM mode dominates
							sSolutions(UBound(sSolutions)).sType = "TM"
							sSolutions(UBound(sSolutions)).iNn = iEHCounter
							iEHCounter = iEHCounter + 1
						Else
							' TE mode dominates
							sSolutions(UBound(sSolutions)).sType = "TE"
							sSolutions(UBound(sSolutions)).iNn = iHECounter
							iHECounter = iHECounter + 1
						End If
					Else
						' For m =/= 0 we have hybrid modes
						If (Abs(cyl_Bessel_j(iM-1,dXi)/(dXi*cyl_Bessel_j(iM,dXi)) - cyl_Bessel_k(iM-1,dEta)/(dEta*cyl_Bessel_k(iM,dEta))) > _
							Abs(cyl_Bessel_j(iM+1,dXi)/(dXi*cyl_Bessel_j(iM,dXi)) + cyl_Bessel_k(iM+1,dEta)/(dEta*cyl_Bessel_k(iM,dEta)))) Then
							' EH mode dominates
							sSolutions(UBound(sSolutions)).sType = "EH"
							sSolutions(UBound(sSolutions)).iNn = iEHCounter
							iEHCounter = iEHCounter + 1
						Else
							' HE mode dominates
							sSolutions(UBound(sSolutions)).sType = "HE"
							sSolutions(UBound(sSolutions)).iNn = iHECounter
							iHECounter = iHECounter + 1
						End If
					End If

					ReDim Preserve sSolutions(UBound(sSolutions) + 1)
					DlgText("TbModes", Format(UBound(sSolutions), "0"))
					dBetaGuess = IIf(dBetaGuess < dBetaLastGuess, dBetaGuess, dBetaLastGuess)
					iFSign = -iFSign
				End If
			End If
			dBetaLastGuess = dBetaGuess
			dBetaGuess = dBetaGuess - dJump
		Loop While (dBetaGuess > dBetaMin)
		If (iSolutionCounter = 0) Then
			iNoSolutionCounter = iNoSolutionCounter + 1
		Else
			iNoSolutionCounter = 0
		End If
		iM = iM + 1
	Wend
	DlgText("LbCalculating", "")
	If (UBound(sSolutions) > 0) Then
		ReDim Preserve sSolutions(UBound(sSolutions) - 1)
	Else
		MsgBox("Calculation could not find any accepting modes. Cylindrical Optical Fibers should always have at least the HE11 mode. " + _
				"Ensure your inputs are correct and try again.",vbOkOnly,"No Solutions Found")
	End If
	FindModesCylidricalOptical = sSolutions

End Function

Function EvalF(dBetaGuess As Double, dN1 As Double, dN2 As Double, dK0 As Double, dD As Double, iM As Integer)
	Dim dP As Double, dQ As Double
	Dim dJ As Double, dK As Double, dJPrime As Double, dKPrime As Double
	Dim dJRatio As Double, dKRatio As Double

	dP = Sqr(dN1^2*dK0^2-dBetaGuess^2)
	dQ = Sqr(dBetaGuess^2-dN2^2*dK0^2)

	' Evaluate Bessel functions
	dJ = cyl_Bessel_j(iM, dP*dD)
	dK = cyl_Bessel_k(iM, dQ*dD)
	dJPrime = .5*(cyl_Bessel_j(iM-1, dP*dD)-cyl_Bessel_j(iM+1, dP*dD))
	dKPrime = .5*(-cyl_Bessel_k(iM-1, dQ*dD)-cyl_Bessel_k(iM+1, dQ*dD))

	dJRatio = dJPrime/(dP*dJ)
	dKRatio = dKPrime/(dQ*dK)

	' The equation for Cylidrical Optical Waveguides
	EvalF = (dJRatio + dKRatio)*(dJRatio + dN2^2/dN1^2*dKRatio) - iM^2/dD^2*(1/dP^2+1/dQ^2)*(1/dP^2+dN2^2/dN1^2/dQ^2)
End Function


Sub newMaterial(sName As String, dEpsilon As Double, dRed As Double, dGreen As Double, dBlue As Double)
	StoreParameter(sName+"_E", dEpsilon)
	Dim sHistoryString As String
	sHistoryString = sHistoryString + "With Material " + vbNewLine
	sHistoryString = sHistoryString + "     .Reset " + vbNewLine
	sHistoryString = sHistoryString + "     .Name " + Chr(34) + sName + Chr(34) + "" + vbNewLine
	sHistoryString = sHistoryString + "     .Folder " + Chr(34) + "" + Chr(34) + "" + vbNewLine
	sHistoryString = sHistoryString + "     .Type " + Chr(34) + "Normal" + Chr(34) + "" + vbNewLine
	sHistoryString = sHistoryString + "     .Epsilon " + Chr(34) + sName+"_E" + Chr(34) + "" + vbNewLine
	sHistoryString = sHistoryString + "     .Colour " + Str(dRed) + ", " + Str(dGreen) + ", " + Str(dBlue) +vbNewLine
	sHistoryString = sHistoryString + "     .Create" + vbNewLine
	sHistoryString = sHistoryString + "End With " + vbNewLine
	AddToHistory("define material: "+sName, sHistoryString)
End Sub

Function buildWaveguide(sType As String, sMaterial1 As String, sMaterial2 As String, sMaterial3 As String, _
					dW0 As Double, dW1 As Double, dH0 As Double, dH1 As Double, _
					dH2 As Double, dLength As Double)
	Dim sHistoryString As String

	If (IsNumeric(sMaterial1)) Then
		If (Eval(sMaterial1) <= 0) Then
			MsgBox("Refractive index of the core (n1) must be assigned a positive number", "Invalid Parameter")
			buildWaveguide = -2
			Exit Function
		End If
		newMaterial("Core",Eval(sMaterial1)^2,0,1,.25)
		sMaterial1 = "Core"
	ElseIf (Not Material.Exists(sMaterial1)) Then
		MsgBox("Unknown material: " + sMaterial1, "Unknown Material")
		buildWaveguide = -2
		Exit Function
	End If
	If (IsNumeric(sMaterial2)) Then
		If (Eval(sMaterial2) <= 0) Then
			MsgBox("Refractive index of the substrate (n2) must be assigned a positive number", "Invalid Parameter")
			buildWaveguide = -2
			Exit Function
		End If
		newMaterial("Substrate",Eval(sMaterial2)^2,.25,.25,.25)
		sMaterial2 = "Substrate"
	ElseIf (Not Material.Exists(sMaterial2)) Then
		MsgBox("Unknown material: " + sMaterial2, "Unknown Material")
		buildWaveguide = -2
		Exit Function
	End If
	StoreParameter("Waveguide_Length", dLength)
	Select Case sType
	Case "Asymmetrical Slab Waveguide"
		If (IsNumeric(sMaterial3)) Then
			If (Eval(sMaterial3) <= 0) Then
				MsgBox("Refractive index of the cover (n3) must be assigned a positive number", "Invalid Parameter")
				buildWaveguide = -2
				Exit Function
			End If
			newMaterial("Cover",Eval(sMaterial3)^2,.25,.25,.25)
			sMaterial3 = "Cover"
		ElseIf (Not Material.Exists(sMaterial3)) Then
			MsgBox("Unknown material: " + sMaterial3, "Unknown Material")
			buildWaveguide = -2
			Exit Function
		End If
		If (dH0 >= dH1) Then
			MsgBox("Waveguide Height (H2) must be latger than Core Height (H1)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		StoreParameter("Waveguide_Height",dH1)
		StoreParameter("Core_Width",dW0)
		StoreParameter("Core_Height",dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Cover" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial3 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Core_Height/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Core_Height/2" + Chr(34) + ", " + Chr(34) + "Core_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Waveguide_Height/2" + Chr(34) + ", " + Chr(34) + "-Core_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
	Case "Cylindrical Optical Fiber"
		If (dW0 >= dW1) Then
			MsgBox("Outter Diameter (W1) must be latger than Inner Diameter (W0)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		StoreParameter("Inner_Diameter", dW0)
		StoreParameter("Outter_Diameter", dW1)
		sHistoryString = sHistoryString + "With Cylinder " + vbNewLine
		sHistoryString = sHistoryString + "     .Reset " + vbNewLine
		sHistoryString = sHistoryString + "     .Name " + Chr(34) + "Core" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Component " + Chr(34) + sType + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Material " + Chr(34) + sMaterial1 + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .OuterRadius " + Chr(34) + "Inner_Diameter/2" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .InnerRadius " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Axis " + Chr(34) + "z" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Xcenter " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Ycenter " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Segments " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Create " + vbNewLine
		sHistoryString = sHistoryString + "End With " + vbNewLine
		sHistoryString = sHistoryString + "With Cylinder " + vbNewLine
		sHistoryString = sHistoryString + "     .Reset " + vbNewLine
		sHistoryString = sHistoryString + "     .Name " + Chr(34) + "Substrate" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Component " + Chr(34) + sType + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Material " + Chr(34) + sMaterial2 + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .OuterRadius " + Chr(34) + "Outter_Diameter/2" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .InnerRadius " + Chr(34) + "Inner_Diameter/2" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Axis " + Chr(34) + "z" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Xcenter " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Ycenter " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Segments " + Chr(34) + "0" + Chr(34) + " " + vbNewLine
		sHistoryString = sHistoryString + "     .Create " + vbNewLine
		sHistoryString = sHistoryString + "End With " + vbNewLine
	Case "Buried Channel Waveguide"
		If (dW0 >= dW1) Then
			MsgBox("Substrate Width (W2) must be latger than Core Width (W1)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		If (dH0 >= dH1) Then
			MsgBox("Substrate Height (H2) must be latger than Core Height (H1)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		StoreParameter("Substrate_Width", dW1)
		StoreParameter("Substrate_Height", dH1)
		StoreParameter("Core_Width",dW0)
		StoreParameter("Core_Height",dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Core_Height/2" + Chr(34) + ", " + Chr(34) + "Core_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "Solid.Insert "+ Chr(34) +sType+ ":Substrate" + Chr(34) + ", "+ Chr(34) +sType+ ":Core" + Chr(34) + "" + vbNewLine
	Case "Strip-Loaded Waveguide"
		If (IsNumeric(sMaterial3)) Then
			If (Eval(sMaterial3) <= 0) Then
				MsgBox("Refractive index of the cover (n3) must be assigned a positive number", "Invalid Parameter")
				buildWaveguide = -2
				Exit Function
			End If
			newMaterial("Cover",Eval(sMaterial3)^2,.25,.25,.25)
			sMaterial3 = "Cover"
		ElseIf (Not Material.Exists(sMaterial3)) Then
			MsgBox("Unknown material: " + sMaterial3, "Unknown Material")
			buildWaveguide = -2
			Exit Function
		End If
		StoreParameter("Substrate_Width", dW1)
		StoreParameter("Substrate_Height", dH2)
		StoreParameter("Core_Height",dH1-dH0)
		StoreParameter("Cover_Width",dW0)
		StoreParameter("Cover_Height",dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Cover" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial3 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Cover_Width/2" + Chr(34) + ", " + Chr(34) + "Cover_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2+Core_Height" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2+Core_Height+Cover_Height" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2+Core_Height" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine

	Case "Ridge Waveguide"
		StoreParameter("Substrate_Width", dW1)
		StoreParameter("Substrate_Height", dH1)
		StoreParameter("Core_Width", dW0)
		StoreParameter("Core_Height", dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2+Core_Height" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
	Case "Rib Waveguide"
		StoreParameter("Substrate_Width", dW1)
		StoreParameter("Substrate_Height", dH2)
		StoreParameter("Core_Height",dH1-dH0)
		StoreParameter("Rib_Width",dW0)
		StoreParameter("Rib_Height",dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Rib_Width/2" + Chr(34) + ", " + Chr(34) + "Rib_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2+Core_Height" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2+Core_Height+Rib_Height" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core1" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2+Core_Height" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "Solid.Add " + Chr(34) +sType+ ":Core"  + Chr(34) + ", " + Chr(34) +sType+ ":Core1" + Chr(34) + vbNewLine
	Case "Diffused Waveguide"
		If (dW0 >= dW1) Then
			MsgBox("Substrate Width (W2) must be latger than Core Width (W1)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		If (dH0 >= dH1) Then
			MsgBox("Substrate Height (H2) must be latger than Core Height (H1)", "Invalid Parameter")
			buildWaveguide = -1
			Exit Function
		End If
		StoreParameter("Substrate_Width", dW1)
		StoreParameter("Substrate_Height", dH1)
		StoreParameter("Core_Width",dW0)
		StoreParameter("Core_Height",dH0)
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Core" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial1 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Core_Width/2" + Chr(34) + ", " + Chr(34) + "Core_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "Substrate_Height/2-Core_Height" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "With Brick" + vbNewLine
		sHistoryString = sHistoryString + "	.Reset" + vbNewLine
		sHistoryString = sHistoryString + "	.Name " + Chr(34) + "Substrate" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Component " + Chr(34) + sType + Chr(34) + vbNewLine
		sHistoryString = sHistoryString + "	.Material " + Chr(34) + sMaterial2 + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Xrange " + Chr(34) + "-Substrate_Width/2" + Chr(34) + ", " + Chr(34) + "Substrate_Width/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Yrange " + Chr(34) + "-Substrate_Height/2" + Chr(34) + ", " + Chr(34) + "Substrate_Height/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Zrange " + Chr(34) + "-Waveguide_Length/2" + Chr(34) + ", " + Chr(34) + "Waveguide_Length/2" + Chr(34) + "" + vbNewLine
		sHistoryString = sHistoryString + "	.Create" + vbNewLine
		sHistoryString = sHistoryString + "End With" + vbNewLine
		sHistoryString = sHistoryString + "Solid.Insert "+ Chr(34) +sType+ ":Substrate" + Chr(34) + ", "+ Chr(34) +sType+ ":Core" + Chr(34) + "" + vbNewLine
	End Select
	AddToHistory("build: "+sType, sHistoryString)
	buildWaveguide = 0

End Function

Function GetNValue(sMaterial As String) As Double
	Dim dEx As Double, dEy As Double, dEz As Double, dMx As Double, dMy As Double, dMz As Double
	If IsNumeric(sMaterial) Then
		GetNValue = Eval(sMaterial)
	ElseIf Material.Exists(sMaterial) Then
		Material.GetEpsilon(sMaterial, dEx, dEy, dEz)
		Material.GetMu(sMaterial, dMx, dMy, dMz)
		GetNValue = Sqr(dEx*dMx)
	Else
		GetNValue = -1
		MsgBox("Unknown material: " + sMaterial,vbOkOnly,"Unknown Material")
	End If
End Function


Function myEval(s As String) As Variant
	If (s = "") Then
		myEval = 0.0
	Else
		myEval = Eval(s)
	End If
End Function
