'#include "vba_globals_all.lib"
'--------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------------------------------------------------------
' 28-Nov-2016 fsr: Converted from mcs to mcr + AppendToHistory; bugfixes; added option to enable/disable monitor (groups) via parameters
' 08-May-2012 ube,ty: dialog handling for 2d monitors corrected
' 26-Mar-2012 fhi: added "log"-Steps for frequency samples, corrected 2D-Plane Monitors
' 12-Mar-2012 ube: on german windows, komma was shown in dialogue, now replaced by point
' 18-Dec-2011 ube: automatic skipping of 0-frq for farfield monitor, added Surface current (TLM Only)
' 04-Aug-2010 ube: make macro readable with new farfield template / adjustment of upper bound of for-loop
' 28-Jun-2007 ube: always replace kommas by points (eg german windows produces komma as separator)
' 27-Oct-2006 ube: tiny changes (handling of cancel case in history)
'        2006 ty@aet : 1st version
'--------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim array_frq_unit(7) As String
Dim array_wlg_unit(7) As String
Dim array_unit(7) As String

Sub Main

	array_frq_unit(0) = "Hz"
	array_frq_unit(1) = "kHz"
	array_frq_unit(2) = "MHz"
	array_frq_unit(3) = "GHz"
	array_frq_unit(4) = "THz"
	array_frq_unit(5) = "PHz"

	array_wlg_unit(0) = "m"
	array_wlg_unit(1) = "cm"
	array_wlg_unit(2) = "mm"
	array_wlg_unit(3) = "um"
	array_wlg_unit(4) = "nm"

	Begin Dialog UserDialog 580,364,"Set discrete broadband field monitors",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,7,240,217,"Type",.GroupBox1
		OptionGroup .Field_type_select
			OptionButton 34,28,90,14,"E-Field",.ef_select
			OptionButton 34,56,180,14,"H-Field/Surface current",.hf_select
			OptionButton 34,84,200,14,"Surface current (TLM Only)",.sff_select
			OptionButton 34,112,110,14,"PowerFlow",.pf_select
			OptionButton 34,140,120,14,"Farfield/RCS",.ff_select
		GroupBox 270,7,290,217,"Specification",.GroupBox2
		OptionGroup .Mon_mode_select
			OptionButton 290,28,100,14,"Frequency",.Set_mode_freq
			OptionButton 290,56,130,14,"Wave length",.Set_mode_wlg
		DropListBox 420,21,120,189,array_frq_unit(),.List_frq_unit
		DropListBox 420,49,120,189,array_wlg_unit(),.List_wlg_unit
		TextBox 370,77,100,21,.mon_val_min
		TextBox 370,105,100,21,.mon_val_max
		TextBox 370,168,100,21,.mon_val_stp
		Text 300,84,50,14,"Min:",.lab_val_min
		Text 300,112,50,14,"Max:",.lab_val_max
		Text 300,175,60,14,"Stepsize:",.lab_val_stp
		GroupBox 20,231,240,98,"2D Plane",.GroupBox3
		Text 40,280,80,14,"Orientation:",.Text4
		OptionGroup .Coord_select
			OptionButton 130,273,40,14,"X",.x_coord
			OptionButton 170,273,40,14,"Y",.y_coord
			OptionButton 210,273,40,14,"Z",.z_coord
		CheckBox 40,252,90,14,"Activate",.Check2d
		Text 40,308,60,14,"Position:",.Text5
		TextBox 130,301,70,21,.val_pos
		Text 480,84,30,14,"Hz",.txt_val_unit1
		Text 480,112,30,14,"Hz",.txt_val_unit2
		Text 480,175,30,14,"Hz",.txt_val_unit3
		Text 200,308,40,14,Units.GetUnit("Length"),.Text6
		OptionGroup .Lin_log
			OptionButton 290,140,90,21,"Lin Steps",.lin_steps
			OptionButton 400,140,100,21,"Log Steps",.log_steps
		Text 300,203,60,14,"Samples:",.Text1
		TextBox 370,196,100,21,.mon_val_stp_log

		GroupBox 270,231,290,98,"Special Settings",.GroupBox4
		CheckBox 290,252,250,14,"Add parameter to enable",.CreateParametersCB
		Text 310,266,120,14,"or disable monitors",.Text2
		CheckBox 290,287,240,14,"Use individual parameters",.IndividualParametersCB
		Text 310,301,110,14,"for each monitor",.IndividualParametersExtendedLabel

		OKButton 270,336,90,21
		PushButton 370,336,90,21,"Apply",.ApplyPB
		CancelButton 470,336,90,21


	End Dialog

	Dim dlg As UserDialog

	If (Dialog(dlg)=0) Then
		Exit All
	End If

End Sub

Function DialogFunc(Item As String, Action As Integer, Value As Integer) As Boolean

	Select Case Action
		Case 1 ' Dialog box initialization

			DlgEnable "Set_mode_freq", True
			DlgEnable "List_frq_unit", True
			DlgEnable "List_wlg_unit", False
			DlgEnable "mon_val_stp_log", False
				DlgValue "lin_log", 0	'lin as default

			DlgText "txt_val_unit1", Units.GetUnit("Frequency")
			DlgText "txt_val_unit2", Units.GetUnit("Frequency")
			DlgText "txt_val_unit3", Units.GetUnit("Frequency")

			DlgEnable "val_pos", False
			DlgEnable "x_coord", False
			DlgEnable "y_coord", False
			DlgEnable "z_coord", False
			DlgText "val_pos", "0"

			Select Case Units.GetUnit("Frequency")
				Case "Hz"
					DlgValue "List_frq_unit", 0
				Case "kHz"
					DlgValue "List_frq_unit", 1
				Case "MHz"
					DlgValue "List_frq_unit", 2
				Case "GHz"
					DlgValue "List_frq_unit", 3
				Case "THz"
					DlgValue "List_frq_unit", 4
				Case "PHz"
					DlgValue "List_frq_unit", 5
			End Select

			Select Case Units.GetUnit("Length")
				Case "m"
					DlgValue "List_wlg_unit", 0
				Case "cm"
					DlgValue "List_wlg_unit", 1
				Case "mm"
					DlgValue "List_wlg_unit", 2
				Case "um"
					DlgValue "List_wlg_unit", 3
				Case "nm"
					DlgValue "List_wlg_unit", 4
			End Select

			DlgValue "Mon_mode_select", 0
			DlgText "mon_val_min", CStr(Format(Solver.GetFmin,set_digit(Solver.GetFmin)))
			DlgText "mon_val_max", CStr(Format(Solver.GetFmax,set_digit(Solver.GetFmax)))
			DlgText "mon_val_stp", CStr(Format((Solver.GetFmax-Solver.GetFmin)/10,set_digit((Solver.GetFmax-Solver.GetFmin)/10)))
			DlgText "mon_val_stp_log", cstr(10)

			DlgValue("CreateParametersCB", False)
			DlgEnable("IndividualParametersCB", False)
			DlgEnable("IndividualParametersExtendedLabel", DlgEnable("IndividualParametersCB"))
			DlgValue("IndividualParametersCB", True)
		Case 2 ' Value changing or button pressed

			Select Case Item

				Case "Field_type_select"
					Select Case Value
						Case 2, 3, 4 'in case of farfield and powerflow
							DlgEnable "check2d", False
							DlgEnable "val_pos", False
							DlgEnable "x_coord", False
							DlgEnable "y_coord", False
							DlgEnable "z_coord", False
						Case Else
							If DlgValue("Check2d") = 1 Then
								DlgEnable "check2d", True
								DlgEnable "val_pos", True
								DlgEnable "x_coord", True
								DlgEnable "y_coord", True
								DlgEnable "z_coord", True
							Else
								DlgEnable "check2d", True
								DlgEnable "val_pos", False
								DlgEnable "x_coord", False
								DlgEnable "y_coord", False
								DlgEnable "z_coord", False
							End If
					End Select
					DialogFunc= True

				Case "Check2d"
					Select Case Value
						Case 1
							DlgEnable "val_pos", True
							DlgEnable "x_coord", True
							DlgEnable "y_coord", True
							DlgEnable "z_coord", True
							DlgEnable "ef_select", True
							DlgEnable "hf_select", True
							DlgEnable "sff_select", False
							DlgEnable "ff_select", False
							DlgEnable "pf_select", False
							DlgText "hf_select", "H-Field"
						Case 0
							DlgEnable "val_pos", False
							DlgEnable "x_coord", False
							DlgEnable "y_coord", False
							DlgEnable "z_coord", False
							DlgEnable "ef_select", True
							DlgEnable "hf_select", True
							DlgEnable "sff_select", True
							DlgEnable "ff_select", True
							DlgEnable "pf_select", True
							DlgText "hf_select", "H-Field/Surface current"
					End Select
					DialogFunc= True

				Case "Mon_mode_select"
					Select Case Value
						Case 0
							DlgEnable "List_frq_unit", True
							DlgEnable "List_wlg_unit", False
							DlgText "txt_val_unit1", DlgText("List_frq_unit")
							DlgText "txt_val_unit2", DlgText("List_frq_unit")
							DlgText "txt_val_unit3", DlgText("List_frq_unit")
						Case 1
							DlgEnable "List_frq_unit", False
							DlgEnable "List_wlg_unit", True
							DlgText "txt_val_unit1", DlgText("List_wlg_unit")
							DlgText "txt_val_unit2", DlgText("List_wlg_unit")
							DlgText "txt_val_unit3", DlgText("List_wlg_unit")
					End Select
					DialogFunc= True

				Case "List_frq_unit"
					DlgText "txt_val_unit1", DlgText("List_frq_unit")
					DlgText "txt_val_unit2", DlgText("List_frq_unit")
					DlgText "txt_val_unit3", DlgText("List_frq_unit")

				Case "List_wlg_unit"
					DlgText "txt_val_unit1", DlgText("List_wlg_unit")
					DlgText "txt_val_unit2", DlgText("List_wlg_unit")
					DlgText "txt_val_unit3", DlgText("List_wlg_unit")

				Case "Lin_log"
					Select Case Value
						Case 0	'lin
								DlgEnable "mon_val_stp_log", False
								DlgEnable "mon_val_stp", True
								DlgValue "lin_log", 0
						Case 1	'log
								DlgEnable "mon_val_stp_log", True
								DlgEnable "mon_val_stp", False
								DlgValue "lin_log", 1
					End Select
					DialogFunc= True

				Case "OK"
					CreateMonitors()
					DialogFunc = False
				Case "ApplyPB"
					CreateMonitors()
					DialogFunc = True

			End Select

		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
	End Select
	DlgEnable("IndividualParametersCB", DlgValue("CreateParametersCB") = 1)
	DlgEnable("IndividualParametersExtendedLabel", DlgEnable("IndividualParametersCB"))
End Function

Function CreateMonitors() As Integer

	Dim zz_val_min As Double, zz_lin_log As Integer, zz_mon_val_stp_log As Integer
	Dim zz_val_max As Double
	Dim zz_val_stp As Double
	Dim zz_val_mon As Double
	Dim zz_val_unit As String
	Dim zz_val_mode As String
	Dim zz_val_digit As String

	Dim zz_mon_freq As Double
	Dim zz_mon_name As String
	Dim zz_mon_name_symbol As String
	Dim zz_mon_type As String

	Dim zz_cut_yes As Boolean
	Dim zz_cut_axis As String
	Dim zz_cut_pos As Double

	Dim zz_unit_factor As Double

	Dim sNumber As String

	Dim sMonitorGroupParameterName As String, sMonitorIndividualParameterName As String
	Dim i As Long

	zz_val_min = Evaluate(DlgText("mon_val_min"))
	zz_val_max = Evaluate(DlgText("mon_val_max"))
	zz_val_stp = Evaluate(DlgText("mon_val_stp"))
	zz_lin_log = DlgValue("lin_log")
	zz_mon_val_stp_log=Evaluate(DlgText("mon_val_stp_log"))

	Select Case DlgValue("Mon_mode_select")
		Case 0
			zz_val_mode = "Frequency"
			zz_val_unit = array_frq_unit(DlgValue("List_frq_unit"))
			zz_unit_factor = 10^(3*DlgValue("List_frq_unit"))
		Case 1
			zz_val_mode = "WaveLength"
			zz_val_unit = array_wlg_unit(DlgValue("List_wlg_unit"))
			zz_val_min = IIf(zz_val_min = 0, zz_val_max/1e9, zz_val_min) ' prevent division by zero
			Select Case DlgValue("List_wlg_unit")
				Case 0
					zz_unit_factor = 1
				Case 1
					zz_unit_factor = 1e-2
				Case Else
					zz_unit_factor = 10^(-3*(DlgValue("List_wlg_unit")-1))
			End Select
	End Select

	Select Case DlgValue("Field_type_select")
		Case 0
			zz_mon_type = "Efield"
			zz_mon_name_symbol = "e-field "
		Case 1
			zz_mon_type = "Hfield"
			zz_mon_name_symbol = "h-field "
		Case 2
			zz_mon_type = "Surfacecurrent"
			zz_mon_name_symbol = "surface-current "
		Case 3
			zz_mon_type = "Powerflow"
			zz_mon_name_symbol = "power "
		Case 4
			zz_mon_type = "Farfield"
			zz_mon_name_symbol = "farfield "
	End Select

	i = 0
	Do
		sMonitorGroupParameterName = Replace(Replace(Replace(Replace(Replace("EnableMonitor_" & zz_mon_name_symbol & "_Group_" & CStr(i), "(", "_"), ")", ""), " ", ""), "=", ""), "-", "")
		i = i + 1
	Loop While (RestoreParameterExpression(sMonitorGroupParameterName) <> "")

	Select Case DlgValue("check2d")
		Case 0
			zz_cut_yes = False
		Case 1
			zz_cut_yes = True
			zz_cut_pos = Evaluate(DlgText("val_pos"))
			zz_cut_axis = Array("x", "y", "z")(DlgValue("Coord_select"))
	End Select

	Dim bModifyMonitorName As Boolean
	bModifyMonitorName = (Units.GetUnit("Frequency")<>zz_val_unit) And (zz_mon_type="Farfield")
	If bModifyMonitorName Then
		If (MsgBox "Please note, that broadband farfield result template will only work on monitors, "+vbCrLf+  _
			"defined using the global Frequency Unit (" + Units.GetUnit("Frequency") + ")" + vbCrLf +  _
			vbCrLf +"Do you want to continue anyway?",vbExclamation+vbYesNo)=vbNo Then
			Exit Function
		End If
	End If

	Dim array_of_frequencies() As Double
	Dim nr_of_steps As Integer, cst_index As Integer, cst_freq_step As Double

	'linear or logarithmic steps
	If zz_lin_log = 0 Then
		'lin
		' upper bound slightly increased by 0.000001*zz_val_stp to ensure it is always considered
		nr_of_steps = (zz_val_max+0.000001*zz_val_stp - zz_val_min) / zz_val_stp + 1
		ReDim array_of_frequencies (nr_of_steps)
		For cst_index = 0 To nr_of_steps-1
			array_of_frequencies (cst_index) = zz_val_min + zz_val_stp*cdbl(cst_index)
		Next
	Else
		ReDim array_of_frequencies (cint(zz_mon_val_stp_log)+1)
		If zz_val_min = 0 Then zz_val_min = zz_val_max/1e9
	    cst_freq_step = (Log(zz_val_max) - Log(zz_val_min))/(cdbl(zz_mon_val_stp_log)-1)
	    For cst_index = 0 To cdbl(zz_mon_val_stp_log)-1
			array_of_frequencies (cst_index) = Exp(Log(zz_val_min)+(cst_index)*cst_freq_step)
	    Next
		nr_of_steps=cint(zz_mon_val_stp_log)
	End If

	'loop thru samples:
	For cst_index = 0 To nr_of_steps-1
		zz_val_mon = array_of_frequencies (cst_index)

		If Abs(zz_val_mon) >= 10000 Then
			zz_val_digit = "0.0####E+###"
		ElseIf Abs(zz_val_mon) = 0 Then
			zz_val_digit = ""
		ElseIf Abs(zz_val_mon) <= 0.001 Then
			zz_val_digit = "0.0####E+###"
		Else
			zz_val_digit = "00.0000"
		End If

		sNumber = Replace(Format(zz_val_mon,zz_val_digit),",",".")

		Select Case zz_val_mode
			Case "Frequency"
				zz_mon_freq = zz_val_mon*zz_unit_factor*Units.GetFrequencySIToUnit
				If (Units.GetUnit("Frequency") = zz_val_unit) Then
					zz_mon_name = zz_mon_name_symbol & "(f="  & sNumber
				Else
					' it is not recommended to use local frequency unit different to global frq-unit,
					zz_mon_name = zz_mon_name_symbol & "(f="  & sNumber & " " & zz_val_unit
				End If
			Case "WaveLength"
				zz_mon_freq = clight/(zz_val_mon*zz_unit_factor)*Units.GetFrequencySIToUnit
				zz_mon_name = zz_mon_name_symbol & "(wl=" & sNumber & " " & zz_val_unit
		End Select

		Dim sHistoryCommand As String

		sHistoryCommand = ""

		AppendHistoryLine_LIB(sHistoryCommand, "With Monitor")
		AppendHistoryLine_LIB(sHistoryCommand, "	.Reset")

		If zz_cut_yes = True Then
			If zz_mon_type <> "Farfield" Then
				If zz_mon_type <> "Powerflow" Then
					If zz_mon_type <> "Surfacecurrent" Then
					zz_mon_name = zz_mon_name & ";" & zz_cut_axis & "=" & zz_cut_pos & ")"
					AppendHistoryLine_LIB(sHistoryCommand, "		.Dimension", "Plane")
					AppendHistoryLine_LIB(sHistoryCommand, "		.PlaneNormal", zz_cut_axis)
					AppendHistoryLine_LIB(sHistoryCommand, "		.PlanePosition", zz_cut_pos)
					End If
				End If
			End If
		Else
			zz_mon_name = zz_mon_name  & ")"
			AppendHistoryLine_LIB(sHistoryCommand, ".Dimension", "Volume")
		End If

		sMonitorIndividualParameterName = "EnableMonitor_" & zz_mon_name
		sMonitorIndividualParameterName = Replace(Replace(Replace(Replace(Replace(Replace(sMonitorIndividualParameterName, "(", ""), ")", ""), " ", ""), "=", ""), "-", ""), ".", "_")

		AppendHistoryLine_LIB(sHistoryCommand, ".Name", zz_mon_name)
		AppendHistoryLine_LIB(sHistoryCommand, ".Domain", "Frequency")
		AppendHistoryLine_LIB(sHistoryCommand, ".FieldType", zz_mon_type)
		AppendHistoryLine_LIB(sHistoryCommand, ".Frequency", Replace(Format(zz_mon_freq),",","."))
		If Not (zz_mon_type = "Farfield" And zz_mon_freq = 0.0) Then
			If CBool(DlgValue("CreateParametersCB")) Then
				MakeSureParameterExists(sMonitorGroupParameterName, "1")
				AppendHistoryLine_LIB(sHistoryCommand, "If (RestoreParameter(" & Chr(34) & sMonitorGroupParameterName & Chr(34) & ") = 1) Then")
				If CBool(DlgValue("IndividualParametersCB")) Then
					MakeSureParameterExists(sMonitorIndividualParameterName, "1")
					AppendHistoryLine_LIB(sHistoryCommand, "If (RestoreParameter(" & Chr(34) & sMonitorIndividualParameterName & Chr(34) & ") = 1) Then")
				End If
			End If
			AppendHistoryLine_LIB(sHistoryCommand, ".Create")
			If CBool(DlgValue("CreateParametersCB")) Then
				AppendHistoryLine_LIB(sHistoryCommand, "End If")
				If CBool(DlgValue("IndividualParametersCB")) Then
					AppendHistoryLine_LIB(sHistoryCommand, "End If")
				End If
			End If
			AppendHistoryLine_LIB(sHistoryCommand, "End With")
			AddToHistory("define " & IIf(zz_mon_type = "Farfield", "farfield ", "") & "monitor: " & zz_mon_name, sHistoryCommand)
		End If

	Next cst_index

End Function

Function set_digit(Value As Double) As String
	If Abs(Value) >= 10000 Then
			set_digit = "0.0####E+###"
		ElseIf Abs(Value) = 0 Then
			set_digit = ""
		ElseIf Abs(Value) <= 0.001 Then
			set_digit = "0.0####E+###"
		Else
			set_digit = ""
		End If
End Function


