' CreateColeColeMaterial

' ================================================================================================
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 18-Jan-2022 mha: Fixed scaling regarding units
' 30-Jul-2021 mha: Added in option for magnetic relaxation, made parametrization possible, updated some material object methods
' 23-Feb-2011 fsr: Made the macro history rebuild compatible
' 28-May-2010 fsr: Added logarithmic sampling
' 24-May-2010 fsr: Initial version

'#include "vba_globals_all.lib"

' global variables
Dim sDefaultMaterialName  As String
Dim sDefaultTBFMin        As String
Dim sDefaultTBFMax        As String
Dim sDefaultTBNoOfSamples As String
Dim sDefaultTBEpsDC       As String
Dim sDefaultTBEpsInf      As String
Dim sDefaultTBDielecTau   As String
Dim sDefaultTBDielecAlpha As String
Dim sDefaultTBMueDC       As String
Dim sDefaultTBMueInf      As String
Dim sDefaultTBMagTau      As String
Dim sDefaultTBMagAlpha    As String

Sub Main ()
	' Initialize global defaults
	sDefaultMaterialName  =  "ColeColeMaterial"
	sDefaultTBFMin        = IIf(Solver.GetFmin=0, "1e-10", CStr(Solver.GetFmin))
	sDefaultTBFMax        = IIf(Solver.GetFmax=0, "100", CStr(Solver.GetFmax))
	sDefaultTBNoOfSamples = "501"
	sDefaultTBEpsDC       = "2"
	sDefaultTBEpsInf      = "1"
	sDefaultTBDielecTau   = "1e-10"
	sDefaultTBDielecAlpha = "0"
	sDefaultTBMueDC       = "2"
	sDefaultTBMueInf      = "1"
	sDefaultTBMagTau      = "1e-10"
	sDefaultTBMagAlpha    = "0"

	Begin Dialog UserDialog 600, 421, "Create Cole-Cole Material", .DialogFunc ' %GRID:10,7,1,1

		' Groupbox - Basic settings
		GroupBox  10,  10, 580, 100, "Basic settings",                .GBBasicSettings
		Text      30,  34, 100,  14, "Material name:",                .TMaterialName
		TextBox  140,  28, 410,  21,                                  .TBMaterialName
		CheckBox  30,  60, 192,  14, "Specify dielectric relaxation", .CBSpecifyDielectricRelax
		CheckBox  30,  83, 194,  14, "Specify magnetic relaxation",   .CBSpecifyMagneticRelax
		Text     250,  60, 320,  50, "Please be aware that if relaxation is not specified," & vbNewLine & "Eps_DC / Mue_DC will be used to specify the" & vbNewLine & "non-dispersive Eps / Mue of the material.", .TWarningRegardingEpsMuDC

		' Groupbox - Frequency and sampling
		GroupBox  10, 120, 580,  80, "Frequency and sampling",   .GBFrequAndSampling
		Text      30, 144,  40,  14, "FMin:",                    .TFMin
		TextBox   95, 138, 100,  21,                             .TBFMin
		Text     205, 144,  40,  14, Units.GetUnit("Frequency"), .TUnitFMin
		Text      30, 172,  40,  14, "FMax:",                    .TFMax
		TextBox   95, 166, 100,  21,                             .TBFMax
		Text     205, 172,  40,  14, Units.GetUnit("Frequency"), .TUnitFMax
		Text     340, 144,  97,  14, "No. of samples:",          .TNoOfSamples
		TextBox  450, 138, 100,  21,                             .TBNoOfSamples
		CheckBox 340, 172, 210,  14, "Logarithmic sampling",     .CBLogSamples

		' Groupbox - Dielectric relaxation
		GroupBox  10, 210, 580,  80, "Dielectric relaxation",  .GBDielectricRelaxation
		Text      30, 234,  55,  14, "Eps_DC:",                .TEpsDC
		TextBox   95, 228, 100,  21,                           .TBEpsDC
		Text      30, 262,  55,  14, "Eps_inf:",               .TEpsInf
		TextBox   95, 256, 100,  21,                           .TBEpsInf
		Text     340, 234,  50,  14, "tau:",                   .TDielecTau
		TextBox  400, 228, 150,  21,                           .TBDielecTau
		Text     560, 234,  10, 14, "s",                       .TDielecTauUnit
		Text     340, 262,  50,  14, "alpha:",                 .TDielecAlpha
		TextBox  400, 256, 150,  21,                           .TBDielecAlpha

		' Groupbox - Magnetic relaxation
		GroupBox  10, 300, 580,  80, "Magnetic relaxation",    .GBMagneticRelaxation
		Text      30, 324,  55,  14, "Mue_DC:",                .TMueDC
		TextBox   95, 318, 100,  21,                           .TBMueDC
		Text      30, 352,  55,  14, "Mue_inf:",               .TMueInf
		TextBox   95, 346, 100,  21,                           .TBMueInf
		Text     340, 324,  50,  14, "tau:",                   .TMagTau
		TextBox  400, 318, 150,  21,                           .TBMagTau
		Text     560, 324,  10, 14, "s",                       .TMagTauUnit
		Text     340, 352,  50,  14, "alpha:",                 .TMagAlpha
		TextBox  400, 346, 150,  21,                           .TBMagAlpha

		OKButton      20, 390, 90, 21
		CancelButton 120, 390, 90, 21

	End Dialog
	Dim dlg As UserDialog
	Dialog dlg


End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		' Groupbox - Basic settings
		DlgText "TBMaterialName", sDefaultMaterialName
		' Groupbox - Frequency and sampling
		DlgText "TBFMin",         sDefaultTBFMin
		DlgText "TBFMax",         sDefaultTBFMax
		DlgText "TBNoOfSamples",  sDefaultTBNoOfSamples
		' Groupbox - Dielectric relaxation
		DlgText "TBEpsDC",        sDefaultTBEpsDC
		DlgText "TBEpsInf",       sDefaultTBEpsInf
		DlgText "TBDielecTau",    sDefaultTBDielecTau
		DlgText "TBDielecAlpha",  sDefaultTBDielecAlpha
		' Groupbox - Magnetic relaxation
		DlgText "TBMueDC",        sDefaultTBMueDC
		DlgText "TBMueInf",       sDefaultTBMueInf
		DlgText "TBMagTau",       sDefaultTBMagTau
		DlgText "TBMagAlpha",     sDefaultTBMagAlpha
		' Checkmarked
		DlgValue "CBSpecifyDielectricRelax", 1
		DlgValue "CBSpecifyMagneticRelax",   0
		DlgValue "CBLogSamples",             0

		' Enabled values
		' Groupbox - Frequency settings
		DlgEnable "TBFMin",        True
		DlgEnable "TBFMax",        True
		DlgEnable "TBNoOfSamples", True
		DlgEnable "CBLogSamples",  True
		' Groupbox - Dielectric relaxation
		DlgEnable "TBEpsDC",       True
		DlgEnable "TBEpsInf",      True
		DlgEnable "TBDielecTau",   True
		DlgEnable "TBDielecAlpha", True
		' Groupbox - Magnetic relaxation
		DlgEnable "TBMueDC",       True
		DlgEnable "TBMueInf",      False
		DlgEnable "TBMagTau",      False
		DlgEnable "TBMagAlpha",    False
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "CBSpecifyDielectricRelax"
				If ( DlgValue("CBSpecifyDielectricRelax") = 0 ) Then
					If ( DlgValue("CBSpecifyMagneticRelax") = 0 ) Then
						' No relaxation used at all... Disable most entries
						' Groupbox - Frequency settings
						DlgEnable "TBFMin",        False
						DlgEnable "TBFMax",        False
						DlgEnable "TBNoOfSamples", False
						DlgEnable "CBLogSamples",  False
						' Groupbox - Dielectric relaxation
						DlgEnable "TBEpsDC",       True
						DlgEnable "TBEpsInf",      False
						DlgEnable "TBDielecTau",   False
						DlgEnable "TBDielecAlpha", False
						' Groupbox - Magnetic relaxation
						DlgEnable "TBMueDC",       True
						DlgEnable "TBMueInf",      False
						DlgEnable "TBMagTau",      False
						DlgEnable "TBMagAlpha",    False
					Else
						' Only dielectric relaxation is disabled.
						' Groupbox - Frequency settings
						DlgEnable "TBFMin",        True
						DlgEnable "TBFMax",        True
						DlgEnable "TBNoOfSamples", True
						DlgEnable "CBLogSamples",  True
						' Groupbox - Dielectric relaxation
						DlgEnable "TBEpsDC",       True
						DlgEnable "TBEpsInf",      False
						DlgEnable "TBDielecTau",   False
						DlgEnable "TBDielecAlpha", False
						' Groupbox - Magnetic relaxation
						DlgEnable "TBMueDC",       True
						DlgEnable "TBMueInf",      True
						DlgEnable "TBMagTau",      True
						DlgEnable "TBMagAlpha",    True
					End If
				Else
					' Dielectric relaxation enabled.
					' Only dielectric relaxation is disabled.
					' Groupbox - Frequency settings
					DlgEnable "TBFMin",        True
					DlgEnable "TBFMax",        True
					DlgEnable "TBNoOfSamples", True
					DlgEnable "CBLogSamples",  True
					' Groupbox - Dielectric relaxation
					DlgEnable "TBEpsDC",       True
					DlgEnable "TBEpsInf",      True
					DlgEnable "TBDielecTau",   True
					DlgEnable "TBDielecAlpha", True
				End If
			Case "CBSpecifyMagneticRelax"
				If ( DlgValue("CBSpecifyMagneticRelax") = 0 ) Then
					If ( DlgValue("CBSpecifyDielectricRelax") = 0 ) Then
						' No relaxation used at all... Disable most entries
						' Groupbox - Frequency settings
						DlgEnable "TBFMin",        False
						DlgEnable "TBFMax",        False
						DlgEnable "TBNoOfSamples", False
						DlgEnable "CBLogSamples",  False
						' Groupbox - Dielectric relaxation
						DlgEnable "TBEpsDC",       True
						DlgEnable "TBEpsInf",      False
						DlgEnable "TBDielecTau",   False
						DlgEnable "TBDielecAlpha", False
						' Groupbox - Magnetic relaxation
						DlgEnable "TBMueDC",       True
						DlgEnable "TBMueInf",      False
						DlgEnable "TBMagTau",      False
						DlgEnable "TBMagAlpha",    False
					Else
						' Only magnetic relaxation is disabled.
						' Groupbox - Frequency settings
						DlgEnable "TBFMin",        True
						DlgEnable "TBFMax",        True
						DlgEnable "TBNoOfSamples", True
						DlgEnable "CBLogSamples",  True
						' Groupbox - Dielectric relaxation
						DlgEnable "TBEpsDC",       True
						DlgEnable "TBEpsInf",      True
						DlgEnable "TBDielecTau",   True
						DlgEnable "TBDielecAlpha", True
						' Groupbox - Magnetic relaxation
						DlgEnable "TBMueDC",       True
						DlgEnable "TBMueInf",      False
						DlgEnable "TBMagTau",      False
						DlgEnable "TBMagAlpha",    False
					End If
				Else
					' Magnetic relaxation enabled.
					' Groupbox - Frequency settings
					DlgEnable "TBFMin",        True
					DlgEnable "TBFMax",        True
					DlgEnable "TBNoOfSamples", True
					DlgEnable "CBLogSamples",  True
					' Groupbox - Dielectric relaxation
					DlgEnable "TBMueDC",       True
					DlgEnable "TBMueInf",      True
					DlgEnable "TBMagTau",      True
					DlgEnable "TBMagAlpha",    True
				End If
			Case "OK"
				' First, make sure that the textboxes are filled with valid values...
				Dim ProblemWithATB As Boolean
				ProblemWithATB = False
				If DlgText("TBMaterialName") = "" Then
					MsgBox "You have entered an empty string for the material name. Please enter a valid name.", vbExclamation
					DialogFunc     = True	' There is an error in the settings -> Don't close the dialog box.
					ProblemWithATB = True
'				ElseIf Material.Exists(DlgText("TBMaterialName")) Then
'					MsgBox "The material name you have specified already belongs to an existing material. Please specify a new material name.", vbExclamation
'					DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
				End If
				If Not ProblemWithATB Then
					CreateEntryInHLForColeColeMaterial()
				Else
					DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
				End If
			Case "Cancel"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
			Case "TBMaterialName"
				' Check whether an empty string has been entered
				If DlgText("TBMaterialName") = "" Then
					MsgBox "You have entered an empty string for the material name. Please enter a valid name.", vbExclamation
					DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
				End If
				' Check for forbidden filenames
				Dim sNewName As String
				sNewName = DlgText("TBMaterialName")
				sNewName = NoForbiddenFilenameCharacters(sNewName)
				If sNewName <> DlgText("TBMaterialName") Then
					DlgText "TBMaterialName", sNewName
					MsgBox "Please be aware that the name you have entered for the material contained some forbidden characters. These characters have been removed automatically.", vbExclamation
					DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
				End If
				' Check if material name already exists...
				If Material.Exists(DlgText("TBMaterialName")) Then
					MsgBox "The material name you have specified already belongs to an existing material. Please be aware that continuing will cause the existing material to be overwritten.", vbExclamation
					DialogFunc = True	' There is an error in the settings -> Don't close the dialog box.
				End If
			Case "TBFMin"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBFMin") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBFMin"))
					If Err.Number <> 0 Then
						DlgText "TBFMin", sDefaultTBFMin
						MsgBox "The value you have entered for the minimal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBFMin"))) Then
							' Make sure frequency is greater then zero
							If Evaluate(DlgText("TBFMin")) <= 0 Then
								DlgText "TBFMin", sDefaultTBFMin
								MsgBox "The value you have entered for the minimal frequency is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBFMin", sDefaultTBFMin
							MsgBox "The value you have entered for the minimal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBFMin", sDefaultTBFMin
					MsgBox "The value you have entered for the minimal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBFMax"
				If DlgText("TBFMax") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBFMax"))
					If Err.Number <> 0 Then
						DlgText "TBFMax", sDefaultTBFMax
						MsgBox "The value you have entered for the maximal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBFMax"))) Then
							' Make sure frequency is greater then zero
							If Evaluate(DlgText("TBFMax")) <= 0 Then
								DlgText "TBFMax", sDefaultTBFMax
								MsgBox "The value you have entered for the maximal frequency is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBFMax", sDefaultTBFMax
							MsgBox "The value you have entered for the maximal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBFMin", sDefaultTBFMax
					MsgBox "The value you have entered for the maximal frequency cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBNoOfSamples"
				If DlgText("TBNoOfSamples") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBNoOfSamples"))
					If Err.Number <> 0 Then
						DlgText "TBNoOfSamples", sDefaultTBNoOfSamples
						MsgBox "The value you have entered for the number of samples cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBNoOfSamples"))) Then
							' Make sure frequency is greater then 1
							If Evaluate(DlgText("TBNoOfSamples")) <= 1 Then
								DlgText "TBNoOfSamples", sDefaultTBNoOfSamples
								MsgBox "The value you have entered for the number of samples is smaller or equal to one and has been replaced by the default value.", vbExclamation
							ElseIf Not Evaluate(DlgText("TBNoOfSamples")) = Int(Evaluate(DlgText("TBNoOfSamples"))) Then
								DlgText "TBNoOfSamples", IIf(Int(Evaluate(DlgText("TBNoOfSamples")))<2, "2", CStr(Int(Evaluate(DlgText("TBNoOfSamples")))))
								MsgBox "The value you have entered for the number of samples is not a natural number and has been replaced by the corresponding integer value.", vbExclamation
							End If
						Else
							' Input text cannot be intepreted as a number.
							DlgText "TBNoOfSamples", sDefaultTBNoOfSamples
							MsgBox "The value you have entered for the number of samples cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBNoOfSamples", sDefaultTBNoOfSamples
					MsgBox "The value you have entered for the number of samples cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBEpsDC"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBEpsDC") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBEpsDC"))
					If Err.Number <> 0 Then
						DlgText "TBEpsDC", sDefaultTBEpsDC
						MsgBox "The value you have entered for EpsDC cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBEpsDC"))) Then
							If Evaluate(DlgText("TBEpsDC")) <= 0 Then
								DlgText "TBEpsDC", sDefaultTBEpsDC
								MsgBox "The value you have entered for EpsDC is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBEpsDC", sDefaultTBEpsDC
							MsgBox "The value you have entered for EpsDC cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBEpsDC", sDefaultTBEpsDC
					MsgBox "The value you have entered for EpsDC cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBEpsInf"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBEpsInf") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBEpsInf"))
					If Err.Number <> 0 Then
						DlgText "TBEpsInf", sDefaultTBEpsInf
						MsgBox "The value you have entered for EpsInf cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBEpsInf"))) Then
							If Evaluate(DlgText("TBEpsInf")) <= 0 Then
								DlgText "TBEpsInf", sDefaultTBEpsInf
								MsgBox "The value you have entered for EpsInf is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBEpsInf", sDefaultTBEpsInf
							MsgBox "The value you have entered for for EpsInf cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBEpsInf", sDefaultTBEpsInf
					MsgBox "The value you have entered for for EpsInf cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBDielecTau"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBDielecTau") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBDielecTau"))
					If Err.Number <> 0 Then
						DlgText "TBDielecTau", sDefaultTBDielecTau
						MsgBox "The value you have entered for the dielectric tau cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBDielecTau"))) Then
							If Evaluate(DlgText("TBDielecTau")) <= 0 Then
								DlgText "TBDielecTau", sDefaultTBDielecTau
								MsgBox "The value you have entered for dielectric tau is smaller then zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBDielecTau", sDefaultTBDielecTau
							MsgBox "The value you have entered for for dielectric tau cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBDielecTau", sDefaultTBDielecTau
					MsgBox "The value you have entered for for dielectric tau cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBDielecAlpha"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBDielecAlpha") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBDielecAlpha"))
					If Err.Number <> 0 Then
						DlgText "TBDielecAlpha", sDefaultTBDielecAlpha
						MsgBox "The value you have entered for the dielectric alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBDielecAlpha"))) Then
							If Evaluate(DlgText("TBDielecAlpha")) < 0 Then
								DlgText "TBDielecAlpha", sDefaultTBDielecAlpha
								MsgBox "The value you have entered for dielectric alpha is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBDielecAlpha", sDefaultTBDielecAlpha
							MsgBox "The value you have entered for for dielectric alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBDielecAlpha", sDefaultTBDielecAlpha
					MsgBox "The value you have entered for for dielectric alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBMueDC"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBMueDC") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBMueDC"))
					If Err.Number <> 0 Then
						DlgText "TBMueDC", sDefaultTBMueDC
						MsgBox "The value you have entered for MueDC cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBMueDC"))) Then
							If Evaluate(DlgText("TBMueDC")) <= 0 Then
								DlgText "TBMueDC", sDefaultTBMueDC
								MsgBox "The value you have entered for MueDC is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBMueDC", sDefaultTBMueDC
							MsgBox "The value you have entered for MueDC cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBMueDC", sDefaultTBMueDC
					MsgBox "The value you have entered for MueDC cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBMueInf"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBMueInf") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBMueInf"))
					If Err.Number <> 0 Then
						DlgText "TBMueInf", sDefaultTBMueInf
						MsgBox "The value you have entered for MueInf cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBMueInf"))) Then
							If Evaluate(DlgText("TBMueInf")) <= 0 Then
								DlgText "TBMueInf", sDefaultTBMueInf
								MsgBox "The value you have entered for MueInf is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBMueInf", sDefaultTBMueInf
							MsgBox "The value you have entered for for MueInf cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBMueInf", sDefaultTBMueInf
					MsgBox "The value you have entered for for MueInf cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBMagTau"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBMagTau") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBMagTau"))
					If Err.Number <> 0 Then
						DlgText "TBMagTau", sDefaultTBMagTau
						MsgBox "The value you have entered for the magnetic tau cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBMagTau"))) Then
							If Evaluate(DlgText("TBMagTau")) <= 0 Then
								DlgText "TBMagTau", sDefaultTBMagTau
								MsgBox "The value you have entered for magnetic tau is smaller then zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBMagTau", sDefaultTBMagTau
							MsgBox "The value you have entered for for magnetic tau cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBMagTau", sDefaultTBMagTau
					MsgBox "The value you have entered for for magnetic tau cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
			Case "TBMagAlpha"
				' Check to see whether the entry is in fact a valid numerical value. If not, set value back to default entry
				If DlgText("TBMagAlpha") <> "" Then
					On Error Resume Next
						Evaluate(DlgText("TBMagAlpha"))
					If Err.Number <> 0 Then
						DlgText "TBMagAlpha", sDefaultTBMagAlpha
						MsgBox "The value you have entered for the magnetic alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(DlgText("TBMagAlpha"))) Then
							If Evaluate(DlgText("TBMagAlpha")) < 0 Then
								DlgText "TBMagAlpha", sDefaultTBMagAlpha
								MsgBox "The value you have entered for magnetic alpha is smaller or equal to zero and has been replaced by the default value.", vbExclamation
							End If
						Else
							' Input text cannot be interpreted as a number.
							DlgText "TBMagAlpha", sDefaultTBMagAlpha
							MsgBox "The value you have entered for for magnetic alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
						End If
					End If
				Else
					DlgText "TBMagAlpha", sDefaultTBMagAlpha
					MsgBox "The value you have entered for for magnetic alpha cannot be interpreted and has been replaced by the default value.", vbExclamation
				End If
		End Select
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function



Private Function CreateEntryInHLForColeColeMaterial() As Boolean
	Dim sCreateMatStr  As String
	Dim bAddToHLSucces As Boolean

	sCreateMatStr = "" & _
	"' Declare variables for creation of Cole-Cole material" & vbNewLine & _
	"Dim lRunningIndex  As Long" & vbNewLine & _
	"Dim dFreq          As Double" & vbNewLine & _
	"Dim dFreq2         As Double" & vbNewLine & _
	"Dim bLogSamples    As Boolean" & vbNewLine & _
	"Dim lNoOfSamples   As Long" & vbNewLine & _
	"Dim dFmin          As Double" & vbNewLine & _
	"Dim dFmax          As Double" & vbNewLine & _
	"Dim bDielecCole    As Boolean" & vbNewLine & _
	"Dim bMagCole       As Boolean" & vbNewLine & _
	"Dim dOmega         As Double" & vbNewLine & _
	"Dim dColeColeDenom As Double" & vbNewLine & _
	"Dim dColeColeReal  As Double" & vbNewLine & _
	"Dim dColeColeImag  As Double" & vbNewLine & _
	"Dim dEpsDC         As Double" & vbNewLine & _
	"Dim dEpsInf        As Double" & vbNewLine & _
	"Dim dDielecTau     As Double" & vbNewLine & _
	"Dim dDielecAlpha   As Double" & vbNewLine & _
	"Dim dMueDC         As Double" & vbNewLine & _
	"Dim dMueInf        As Double" & vbNewLine & _
	"Dim dMagTau        As Double" & vbNewLine & _
	"Dim dMagAlpha      As Double" & vbNewLine

	sCreateMatStr = sCreateMatStr & _
	vbNewLine & _
	"' Initialize variables" & vbNewLine & _
	"bLogSamples  = " & IIf(DlgValue("CBLogSamples") = 0, "False", "True") & vbNewLine & _
	"lNoOfSamples = " & DlgText("TBNoOfSamples") & vbNewLine & _
	"dFmin        = " & DlgText("TBFMin") & vbNewLine & _
	"dFmax        = " & DlgText("TBFMax") & vbNewLine & _
	"bDielecCole  = " & IIf(DlgValue("CBSpecifyDielectricRelax") = 0, "False", "True") & vbNewLine & _
	"bMagCole     = " & IIf(DlgValue("CBSpecifyMagneticRelax") = 0, "False", "True") & vbNewLine & _
	"dEpsDC       = " & DlgText("TBEpsDC") & vbNewLine & _
	"dEpsInf      = " & DlgText("TBEpsInf") & vbNewLine & _
	"dDielecTau   = " & DlgText("TBDielecTau") & vbNewLine & _
	"dDielecAlpha = " & DlgText("TBDielecAlpha") & vbNewLine & _
	"dMueDC       = " & DlgText("TBMueDC") & vbNewLine & _
	"dMueInf      = " & DlgText("TBMueInf") & vbNewLine & _
	"dMagTau      = " & DlgText("TBMagTau") & vbNewLine & _
	"dMagAlpha    = " & DlgText("TBMagAlpha") & vbNewLine

	sCreateMatStr = sCreateMatStr & _
	vbNewLine & _
	"With Material" & vbNewLine & _
    "    .Reset" & vbNewLine & _
	"    .Name " & Chr(34) & DlgText("TBMaterialName") & Chr(34) & vbNewLine & _
	"    .Folder " & Chr(34) & Chr(34) & vbNewLine & _
	"    .Rho " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .ThermalType " & Chr(34) & "Normal" & vbNewLine & _
	"    .ThermalConductivity " & Chr(34) & "0" & vbNewLine & _
	"    .SpecificHeat " & Chr(34) & "0" & Chr(34) & ", " & Chr(34) & "J/K/kg" & Chr(34) & vbNewLine & _
	"    .DynamicViscosity " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .Emissivity " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .MetabolicRate " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .VoxelConvection " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .BloodFlow " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .MechanicsType " & Chr(34) & "Unused" & Chr(34) & vbNewLine & _
	"    .FrqType " & Chr(34) & "all" & Chr(34) & vbNewLine & _
	"    .Type " & Chr(34) & "Normal" & Chr(34) & vbNewLine & _
	"    .MaterialUnit " & Chr(34) & "Frequency" & Chr(34) & ", " & Chr(34) & Units.GetUnit("Frequency") & Chr(34) & vbNewLine & _
	"    .MaterialUnit " & Chr(34) & "Geometry" & Chr(34) & ", " & Chr(34) & Units.GetUnit("Length") & Chr(34) & vbNewLine & _
	"    .MaterialUnit " & Chr(34) & "Time" & Chr(34) & ", " & Chr(34) & Units.GetUnit("Time") & Chr(34) & vbNewLine & _
	"    .MaterialUnit " & Chr(34) & "Temperature" & Chr(34) & ", " & Chr(34) & Units.GetUnit("Temperature") & Chr(34) & vbNewLine & _
	"    .Epsilon " & Chr(34) & DlgText("TBEpsDC") & Chr(34) & vbNewLine & _
	"    .Mu " & Chr(34) & DlgText("TBMueDC") & Chr(34) & vbNewLine & _
	"    .Sigma " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .TanD " & Chr(34) & "1" & Chr(34) & vbNewLine & _
	"    .TanDFreq " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .TanDGiven " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .TanDModel " & Chr(34) & "ConstTanD" & Chr(34) & vbNewLine & _
	"    .SetConstTanDStrategyEps " & Chr(34) & "AutomaticOrder" & Chr(34) & vbNewLine & _
	"    .ConstTanDModelOrderEps " & Chr(34) & "3" & Chr(34) & vbNewLine & _
	"    .DjordjevicSarkarUpperFreqEps " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .SetElParametricConductivity " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .ReferenceCoordSystem " & Chr(34) & "Global" & Chr(34) & vbNewLine & _
	"    .CoordSystemType " & Chr(34) & "Cartesian" & Chr(34) & vbNewLine & _
	"    .SigmaM  " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .TanDM " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .TanDMFreq " & Chr(34) & "0.0" & Chr(34) & vbNewLine & _
	"    .TanDMGiven " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .TanDMModel " & Chr(34) & "ConstTanD" & Chr(34) & vbNewLine & _
	"    .SetConstTanDStrategyMu " & Chr(34) & "AutomaticOrder" & Chr(34) & vbNewLine & _
	"    .ConstTanDModelOrderMu " & Chr(34) & "3" & Chr(34) & vbNewLine & _
	"    .DjordjevicSarkarUpperFreqMu " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .SetMagParametricConductivity " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .DispModelEps " & Chr(34) & "None" & Chr(34) & vbNewLine & _
	"    .DispModelMu " & Chr(34) & "None" & Chr(34) & vbNewLine & _
	"    .DispersiveFittingSchemeEps " & Chr(34) & "Nth Order" & Chr(34) & vbNewLine & _
	"    .MaximalOrderNthModelFitEps " & Chr(34) & "10" & Chr(34) & vbNewLine & _
	"    .ErrorLimitNthModelFitEps " & Chr(34) & "0.01" & Chr(34) & vbNewLine & _
	"    .UseOnlyDataInSimFreqRangeNthModelEps " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .DispersiveFittingSchemeMu " & Chr(34) & "Nth Order" & Chr(34) & vbNewLine & _
	"    .MaximalOrderNthModelFitMu " & Chr(34) & "10" & Chr(34) & vbNewLine & _
	"    .ErrorLimitNthModelFitMu " & Chr(34) & "0.01" & Chr(34) & vbNewLine & _
	"    .UseOnlyDataInSimFreqRangeNthModelMu " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    If bDielecCole Then" & vbNewLine & _
	"        .DispersiveFittingFormatEps " & Chr(34) & "Real_Imag" & Chr(34) & vbNewLine & _
	"        ' Add in dispersion data..." & vbNewLine & _
	"        For lRunningIndex = 0 To lNoOfSamples-1 STEP 1" & vbNewLine & _
	"            ' Define frequency" & vbNewLine & _
	"            If bLogSamples Then" & vbNewLine & _
	"            If ( dFmin = 0 ) Then dFmin = 1e-10" & vbNewLine & _
	"                dFreq = dFmin*10^(lRunningIndex*Log(dFmax/dFmin)/Log(10)/(lNoOfSamples-1))" & vbNewLine & _
	"            Else" & vbNewLine & _
	"                dFreq = dFmin + lRunningIndex*(dFmax-dFmin)/(lNoOfSamples-1)" & vbNewLine & _
	"            End If" & vbNewLine & _
	"            ' Define material values at frequency" & vbNewLine & _
	"            dFreq2         = dFreq*Units.GetFrequencyUnitToSI" & vbNewLine & _
	"            dOmega         = 2 * pi * dFreq2" & vbNewLine & _
	"            dColeColeDenom = ((1+(Cos((1-dDielecAlpha)*pi/2))*(dOmega*dDielecTau)^(1-dDielecAlpha))^2+((Sin((1-dDielecAlpha)*pi/2))*(dOmega*dDielecTau)^(1-dDielecAlpha))^2)" & vbNewLine & _
	"            dColeColeReal  = dEpsInf+(dEpsDC-dEpsInf)*(1+(Cos((1-dDielecAlpha)*pi/2))*(dOmega*dDielecTau)^(1-dDielecAlpha))/dColeColeDenom" & vbNewLine & _
	"            dColeColeImag  = (dEpsDC-dEpsInf)*Sin((1-dDielecAlpha)*pi/2)*(dOmega*dDielecTau)^(1-dDielecAlpha)/dColeColeDenom" & vbNewLine & _
	"            .AddDispersionFittingValueEps Replace(CStr(dFreq), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), Replace(CStr(dColeColeReal), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), Replace(CStr(dColeColeImag), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), " & Chr(34) & "1"  & Chr(34) & vbNewLine & _
	"        Next lRunningIndex" & vbNewLine & _
	"        ' Activate dispersion fitting" & vbNewLine & _
	"        .UseGeneralDispersionEps " & Chr(34) & "True" & Chr(34) & vbNewLine & _
	"    Else" & vbNewLine & _
	"        .UseGeneralDispersionEps " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    End If" & vbNewLine & _
	"    If bMagCole Then" & vbNewLine & _
	"        .DispersiveFittingFormatMu " & Chr(34) & "Real_Imag" & Chr(34) & vbNewLine & _
	"        ' Add in dispersion data..." & vbNewLine & _
	"        For lRunningIndex = 0 To lNoOfSamples-1 STEP 1" & vbNewLine & _
	"            ' Define frequency" & vbNewLine & _
	"            If bLogSamples Then" & vbNewLine & _
	"                If ( dFmin = 0 ) Then dFmin = 1e-10" & vbNewLine & _
	"                dFreq = dFmin*10^(lRunningIndex*Log(dFmax/dFmin)/Log(10)/(lNoOfSamples-1))" & vbNewLine & _
	"            Else" & vbNewLine & _
	"                dFreq = dFmin + lRunningIndex*(dFmax-dFmin)/(lNoOfSamples-1)" & vbNewLine & _
	"           End If" & vbNewLine & _
	"           ' Define material values at frequency" & vbNewLine & _
	"           dFreq2         = dFreq*Units.GetFrequencyUnitToSI" & vbNewLine & _
	"           dOmega         = 2 * pi * dFreq2" & vbNewLine & _
	"           dColeColeDenom = ((1+(Cos((1-dMagAlpha)*pi/2))*(dOmega*dMagTau)^(1-dMagAlpha))^2+((Sin((1-dMagAlpha)*pi/2))*(dOmega*dMagTau)^(1-dMagAlpha))^2)" & vbNewLine & _
	"           dColeColeReal  = dMueInf+(dMueDC-dMueInf)*(1+(Cos((1-dMagAlpha)*pi/2))*(dOmega*dMagTau)^(1-dMagAlpha))/dColeColeDenom" & vbNewLine & _
	"           dColeColeImag  = (dMueDC-dMueInf)*Sin((1-dMagAlpha)*pi/2)*(dOmega*dMagTau)^(1-dMagAlpha)/dColeColeDenom" & vbNewLine & _
	"           .AddDispersionFittingValueMu Replace(CStr(dFreq), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), Replace(CStr(dColeColeReal), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), Replace(CStr(dColeColeImag), " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & "), " & Chr(34) & "1"  & Chr(34) & vbNewLine & _
	"        Next lRunningIndex" & vbNewLine & _
	"        ' Activate dispersion fitting" & vbNewLine & _
	"        .UseGeneralDispersionMu " & Chr(34) & "True" & Chr(34) & vbNewLine & _
	"    Else" & vbNewLine & _
	"        .UseGeneralDispersionMu " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    End If" & vbNewLine & _
	"    .NLAnisotropy " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .NLAStackingFactor " & Chr(34) & "1" & Chr(34) & vbNewLine & _
	"    .NLADirectionX " & Chr(34) & "1" & Chr(34) & vbNewLine & _
	"    .NLADirectionY " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .NLADirectionZ " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .Colour " & Chr(34) & "0" & Chr(34) & ", " & Chr(34) & "1" & Chr(34) & ", " & Chr(34) & "1" & Chr(34) & vbNewLine & _
	"    .Wireframe " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .Reflection " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .Allowoutline " & Chr(34) & "True" & Chr(34) & vbNewLine & _
	"    .Transparentoutline " & Chr(34) & "False" & Chr(34) & vbNewLine & _
	"    .Transparency " & Chr(34) & "0" & Chr(34) & vbNewLine & _
	"    .Create" & vbNewLine & _
	"End With" & vbNewLine

	' ReportInformationToWindow(sCreateMatStr)

	bAddToHLSucces = AddToHistory("Define Cole-Cole material: " & DlgText("TBMaterialName"), sCreateMatStr)
	CreateEntryInHLForColeColeMaterial = bAddToHLSucces
End Function
