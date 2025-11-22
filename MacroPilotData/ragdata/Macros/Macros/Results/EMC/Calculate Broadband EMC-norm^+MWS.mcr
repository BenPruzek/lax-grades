' *EMC / Calculate Broadband EMC-norm
' !!! Do not change the line above !!!
' macro.963
'--------------------------------------------------------------------------------------------
' Copyright 2001-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 10-Aug-2018 yta: extends the limit curve to project Fmax if Fmax is greater than the highest freq in the standard
' 30-Jan-2017 ube,rsh: "Farfields\Farfield Cuts" is skipped
' 10-Feb-2012 rsj: "Farfields\farfield (broadband)" is skipped when stepping through singlfe-frq-farfields
' 31-Aug-2011 rsj: added the FCC emission norm
' 09-Feb-2008 ty : Small correction is added for broadband farfield extraction
' 09-Nov-2007 ube: new broadband farfield monitor included
' 23-Jul-2007 ube: adapted to 2008
' 13-Oct-2005 ube: Included into Online Help
' 26-May-2004 ube: completely rewritten (new choice of excitation string, now completely without DataCache-GlobalDataValue )
' 29-Sep-2003 ube: new name and new checks
' 19-Aug-2002 ube: some cosmetics (delete old, unused lines)
' 26-Jul-2002 ube: now with global data cache (no longer MWS-parameters, visible under Edit->Parameters)
' 22-Apr-2002 ube: bugfix switch off linear farfield scaling
' 09-Oct-2001 ube: only for MWS 4
'--------------------------------------------------------------------------------------------
Option Explicit
'#include "vba_globals_all.lib"

' constants define class A and class B norm

Public Const dFrqNormHz = Array( 30.0e6, 230.0e6, 230.0e6, 1000.0e6)

Public Const dBuVm_A03m = Array( 50.0  ,  50.0  ,  57.0  ,   57.0  )
Public Const dBuVm_A10m = Array( 40.0  ,  40.0  ,  47.0  ,   47.0  )
Public Const dBuVm_B03m = Array( 40.0  ,  40.0  ,  47.0  ,   47.0  )
Public Const dBuVm_B10m = Array( 30.0  ,  30.0  ,  37.0  ,   37.0  )

' FCC emission curve for digital devices
' only available for class A 10m distance and class B 03m distance

Public Const emc_norm_list = Array ( "EN/IEC", "FCC" )
'Public Const dFrqNormHz_FCC = Array ( 30.0e6, 88.0e6, 88.0e6, 216.0e6, 216.0e6, 960.0e6, 960.0e6, 40.0e9) '20180810 redefining as dynamic array
Public Const dBuVm_A10m_FCC = Array ( 39.08 , 39.08 ,  43.52,  43.52 ,  46.44 ,  46.44 , 49.54  , 49.54 )
Public Const dBuVm_B03m_FCC = Array ( 40.00 , 40.00 ,  43.52,  43.52 ,  46.02 ,  46.02 , 53.98  , 53.98 )



Sub Main


	Dim nexci_far As Long
	nexci_far = 0

	Dim ffname As String
	Dim Excitation_Names$()
	Dim sExci2 As String
	Dim sMoniName As String, sCompName As String

	Dim cst_emc_norm As Integer

	sCompName = ""

	Dim b_broadband_farfield_exists As Boolean, b_single_frq_farfield_exists As Boolean
	b_broadband_farfield_exists = False
	b_single_frq_farfield_exists = False

	ffname = Resulttree.GetFirstChildName ("Farfields")

	While ffname <> ""
		If (ffname = "Farfields\Farfield Cuts") Then ffname = Resulttree.GetNextItemName(ffname) 'Skip Farfield Cuts

		sMoniName = Mid(ffname,1,InStr(ffname,"[")-2)

		If sMoniName = "Farfields\farfield (broadband)" Then
			b_broadband_farfield_exists = True
		Else
			b_single_frq_farfield_exists = True
		End If

		If sCompName = "" Then
			' first monitor
			sCompName = sMoniName
		End If

		If sCompName = sMoniName Then
			sExci2 = Mid(ffname,InStr(ffname,"["))
			nexci_far = nexci_far + 1
			ReDim Preserve Excitation_Names$(nexci_far-1)
			Excitation_Names$(nexci_far-1) = sExci2
		End If

		ffname=Resulttree.GetNextItemName (ffname)

	Wend

	If (nexci_far = 0) Then
		MsgBox "No Farfield is calculated so far.", vbCritical
		Exit All
	End If

	Dim sExcit As String

'	If (nexci_far = 1) Then
'		sExcit = Excitation_Names$(0)
'	Else

	Begin Dialog UserDialog 380,245,"Calculate Broadband EMC-norm",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,14,180,56,"Choose excitation string",.GroupBox1
		DropListBox 40,35,140,192,Excitation_Names(),.iexc
		GroupBox 20,77,340,126,"",.GroupBox2
		OKButton 30,217,90,21
		CancelButton 130,217,90,21
		PushButton 230,217,90,21,"Help",.Help
		OptionGroup .bbff
			OptionButton 40,91,280,14,"Evaluate all single frq farfield monitors",.OptionButton1
			OptionButton 40,112,260,14,"Evaluate farfield (broadband) monitor",.OptionButton2
		Text 70,136,70,14,"Frq range:",.Text_frq
		Text 145,136,40,14,"min:",.Text_fmin
		Text 145,157,60,14,"max:",.Text_fmax
		Text 145,178,70,14,"stepsize:",.Text_fstep
		TextBox 210,133,120,21,.fmin
		TextBox 210,154,120,21,.fmax
		TextBox 210,175,120,21,.fstep
		GroupBox 210,14,150,56,"Choose EMC limits",.emc_norm
		DropListBox 230,35,110,192,emc_norm_list(),.norm_list
	End Dialog
	Dim dlg As UserDialog

	dlg.bbff = IIf(b_broadband_farfield_exists,1,0)

	dlg.fmin = CStr(Solver.GetFmin)
	dlg.fmax = CStr(Solver.GetFmax)
	dlg.fstep = CStr((Solver.GetFmax-Solver.GetFmin)/10.0)

	dlg.norm_list = 0

	If (Dialog(dlg) >= 0) Then Exit All
	sExcit = Excitation_Names$(dlg.iexc)

	' loop through all existing monitors (with correct excitation string)

	Dim EMCdata03m   As Object
	Dim EMCdata10m   As Object

	Set EMCdata03m = Result1D("")
	Set EMCdata10m = Result1D("")

	cst_emc_norm=dlg.norm_list

	If dlg.bbff Then

		If Not b_broadband_farfield_exists Then
			MsgBox "Farfield result ""farfield (broadband)"" not found. Macro stops.", vbCritical
			Exit All
		End If

		' --- only consider farfield (broadband) (with correct excitation string)

		Dim dfmin As Double, dfmax As Double, dfstep As Double, dfnow As Double

		dfmin  = Evaluate(dlg.fmin)
		dfmax  = Evaluate(dlg.fmax)
		dfstep = Evaluate(dlg.fstep)

		If dfmax < dfmin Or dfstep <= 0 Or dfmin < 0 Then
			MsgBox "Incorrect frequency range settings. Macro stops.", vbCritical
			Exit All
		End If

		ffname = Resulttree.GetFirstChildName ("Farfields")
		While ffname <> ""

			If (ffname = "Farfields\farfield (broadband) " + sExcit) Then
				' CorrectExcitationString, now extract data values

				SelectTreeItem ffname
				With FarfieldPlot
					.Plottype "3d"
					.SetPlotMode "efield"
			        .SetScaleLinear "False"
					.DBUnit "120"

					.SetTimeDomainFF False

					dfnow = dfmin

					If dfnow = 0 Then dfnow = dfstep

					While dfnow < dfmax

						.Setfrequency CStr(dfnow)

						.Distance "03"
						.Plot
						Wait 0.02
						EMCdata03m.AppendXY dfnow, .Getmax

						.Distance "10"
						.Plot
						Wait 0.02
						EMCdata10m.AppendXY dfnow, .Getmax

						dfnow = dfnow + dfstep

					Wend
				End With
			End If

			ffname=Resulttree.GetNextItemName (ffname)
		Wend

	Else

		If Not b_single_frq_farfield_exists Then
			MsgBox "No single frq farfield monitors found. Macro stops.", vbCritical
			Exit All
		End If

		' --- now loop all single frq farfield monitors (with correct excitation string)

		ffname = Resulttree.GetFirstChildName ("Farfields")

		While ffname <> ""
			If (ffname = "Farfields\Farfield Cuts") Then ffname = Resulttree.GetNextItemName(ffname) 'Skip Farfield Cuts

			If Left(ffname,30)="Farfields\farfield (broadband)" Then
				'
				' "Farfields\farfield (broadband)" is skipped when stepping through singlfe-frq-farfields
				'
				ffname=Resulttree.GetNextItemName (ffname)
			End If

			sExci2 = Mid(ffname,InStr(ffname,"["))

			If (sExcit = sExci2) Then
				' CorrectExcitationString, now extract data values

				SelectTreeItem ffname
				With FarfieldPlot
					.Plottype "3d"
					.SetPlotMode "efield"
			        .SetScaleLinear "False"
					.DBUnit "120"
					.Distance "03"
					.Plot
					Wait 0.02
					EMCdata03m.AppendXY GetFieldFrequency, .Getmax

					.Distance "10"
					.Plot
					Wait 0.02
					EMCdata10m.AppendXY GetFieldFrequency, .Getmax
				End With
			End If

			ffname=Resulttree.GetNextItemName (ffname)
		Wend
	End If

	' finally create reference curves (type A and B, 3 and 10m distance)


	Dim divFact As Double
	divFact = Units.GetFrequencyUnitToSI

	Dim classA03m   As Object
	Dim classA10m   As Object
	Dim classB03m   As Object
	Dim classB10m   As Object

	Set classA03m = Result1D("")
	Set classA10m = Result1D("")
	Set classB03m = Result1D("")
	Set classB10m = Result1D("")

	Dim iii As Long

	' rsj: added the option for different EMC norm curve
	If cst_emc_norm=0 Then
		For iii= 0 To UBound(dFrqNormHz)
			classA03m.AppendXY dFrqNormHz(iii)/divFact, dBuVm_A03m(iii)
			classA10m.AppendXY dFrqNormHz(iii)/divFact, dBuVm_A10m(iii)
			classB03m.AppendXY dFrqNormHz(iii)/divFact, dBuVm_B03m(iii)
			classB10m.AppendXY dFrqNormHz(iii)/divFact, dBuVm_B10m(iii)
		Next iii
			'yta: added to extend the EMC norm to project Fmax
			If Solver.GetFmax*Units.GetFrequencyUnitToSI > dFrqNormHz(3) Then
				classA03m.AppendXY Solver.GetFmax, dBuVm_A03m(iii-1)
				classA10m.AppendXY Solver.GetFmax, dBuVm_A10m(iii-1)
				classB03m.AppendXY Solver.GetFmax, dBuVm_B03m(iii-1)
				classB10m.AppendXY Solver.GetFmax, dBuVm_B10m(iii-1)
			End If
	Else
			'20180810yta: added to use solver fmax instead of the default 40GHz to cover typical cases
			Dim dFrqNormHz_FCC(8) As Double
			dFrqNormHz_FCC(0) = 30e6
			dFrqNormHz_FCC(1) = 88e6
			dFrqNormHz_FCC(2) = 88e6
			dFrqNormHz_FCC(3) = 216e6
			dFrqNormHz_FCC(4) = 216e6
			dFrqNormHz_FCC(5) = 960e6
			dFrqNormHz_FCC(6) = 960e6
			dFrqNormHz_FCC(7) = IIf(Solver.GetFmax*Units.GetFrequencyUnitToSI <= dFrqNormHz_FCC(6), 1e9, Solver.GetFmax*Units.GetFrequencyUnitToSI) 'plot extends to 1GHz if Fmax<960MHz
		'For iii= 0 To UBound(dFrqNormHz_FCC)
		'20180810yta: changed to UBound -1
		For iii= 0 To UBound(dFrqNormHz_FCC)-1
			If iii=UBound(dFrqNormHz_FCC) And EMCdata03m.getx(EMCdata03m.getn-1)<dFrqNormHz_FCC(UBound(dFrqNormHz_FCC))/divFact Then
				classA10m.AppendXY EMCdata03m.getx(EMCdata03m.getn-1), dBuVm_A10m_FCC(iii)
				classB03m.AppendXY EMCdata03m.getx(EMCdata03m.getn-1), dBuVm_B03m_FCC(iii)
			Else
				classA10m.AppendXY dFrqNormHz_FCC(iii)/divFact, dBuVm_A10m_FCC(iii)
				classB03m.AppendXY dFrqNormHz_FCC(iii)/divFact, dBuVm_B03m_FCC(iii)
			End If
		Next iii
			'yta: added to extend the EMC norm to project Fmax
			If Solver.GetFmax*Units.GetFrequencyUnitToSI > dFrqNormHz_FCC(6) Then
				classA10m.AppendXY Solver.GetFmax, dBuVm_A10m_FCC(iii-1)
				classB03m.AppendXY Solver.GetFmax, dBuVm_B03m_FCC(iii-1)
			End If
	End If
	' save all files and add them in the 1D Results tree

	Dim sNewFileName    As String
	sNewFileName=GetProjectPath("Result")

	' 03 m

	'rsj : Additional FCC Norm

	With EMCdata03m
		.SetYLabelAndUnit "EMC: Efield at 3 meter distance", "dB(uV/m)"
		.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
		.Type "Linear"
		.Save sNewFileName + "EMC-level_03m_" + NoForbiddenFilenameCharacters(sExcit) + ".sig"
		.AddToTree "1D Results\EMC\03 meter\EMC " + sExcit
	End With

	If cst_emc_norm=0 Then
		With classA03m
			.SetYLabelAndUnit "EMC: Efield at 3 meter distance", "dB(uV/m)"
			.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
			.Type "Linear"
			.Save sNewFileName + "EMC-Class_A_03m.sig"
			.AddToTree "1D Results\EMC\03 meter\Class A"
		End With
	End If

	With classB03m
		.SetYLabelAndUnit "EMC: Efield at 3 meter distance" , "dB(uV/m)"
		.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
		.Type "Linear"
		.Save sNewFileName + "EMC-Class_B_03m.sig"
		.AddToTree "1D Results\EMC\03 meter\Class B"
	End With

	' 10 m

	With EMCdata10m
		.SetYLabelAndUnit "EMC: Efield at 10 meter distance", "dB(uV/m)"
		.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
		.Type "Linear"
		.Save sNewFileName + "EMC-level_10m_" + NoForbiddenFilenameCharacters(sExcit) + ".sig"
		.AddToTree "1D Results\EMC\10 meter\EMC " + sExcit
	End With

	With classA10m
		.SetYLabelAndUnit "EMC: Efield at 10 meter distance", "dB(uV/m)"
		.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
		.Type "Linear"
		.Save sNewFileName + "EMC-Class_A_10m.sig"
		.AddToTree "1D Results\EMC\10 meter\Class A"
	End With

	If cst_emc_norm=0 Then
		With classB10m
			.SetYLabelAndUnit "EMC: Efield at 10 meter distance", "dB(uV/m)"
			.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
			.Type "Linear"
			.Save sNewFileName + "EMC-Class_B_10m.sig"
			.AddToTree "1D Results\EMC\10 meter\Class B"
		End With
	End If

	Wait 0.1
	SelectTreeItem "1D Results\EMC\03 meter"

End Sub
Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

	If (Action%=1 Or Action%=2) Then

		DlgEnable "Text_frq"  , DlgValue("bbff")
		DlgEnable "Text_fmin" , DlgValue("bbff")
		DlgEnable "Text_fmax" , DlgValue("bbff")
		DlgEnable "Text_fstep", DlgValue("bbff")
		DlgEnable "fmin" , DlgValue("bbff")
		DlgEnable "fmax" , DlgValue("bbff")
		DlgEnable "fstep", DlgValue("bbff")

		Select Case Item
		Case "Help"
			StartHelp "common_preloadedmacro_emc_calculate_broadband_emc-norm"
			DialogFunc = True
		End Select
	End If

End Function
