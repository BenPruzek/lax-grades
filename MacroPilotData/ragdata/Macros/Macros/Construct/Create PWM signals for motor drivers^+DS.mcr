'#Language "WWB-COM"

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 20-Sep-2023 rsj: code rework with dialog func, GUI modification, port assignment drop down, port excitation signal name,
'                  ASCII file name contains also task name, use SaveSettings, ASCII with periodic function
' 02-May-2022 rsj: added the option for third harmonic adder, space vector, ii Index start from 0, carrier wave
'                  shift parameter, Tmax assignment for existing transient task, use transistion time based as Fmax estimator
'                  n_samples as Long, Update the help content
' 18-Aug-2021 fhi: added inner-resistance to the task, create a trans.task if not existing and sets a proper end-time
' 18-Sep-2017 ube: added Online Help and Help Button
' 07-Jul-2017 ube: included in official version
' 26-May-2017 aba: First version
'-----------------------------------------------------------------------------------------------------------------------------

Private Const HelpFileName = "common_preloadedmacro_PWM_signals_for_motor_drivers"
Const ListArray_ModTech = Array ("Sine Triangle", "3rd Harmonic Adder", "Space Vector")
Const k_factor = 1/6 '3rd Harmonic fraction of amplitude allows to have maximum modulation
Dim ListArray_Port() As String
Dim numExtPort As Integer

Sub Main

	numExtPort=FillPortNameArray()

	Begin Dialog UserDialog 710,273,"PWM Generator",.DialogFunction ' %GRID:10,7,1,1

		GroupBox 20,7,330,252,"PWM signal properties",.GroupBox1
		Text 30,35,140,14,"PWM frequency [Hz]",.Text1
		TextBox 190,28,150,21,.pwm_dialouge
		Text 30,63,140,14,"Sine frequency [Hz]",.Text2
		TextBox 190,56,150,21,.sine_dialouge
		Text 30,91,130,14,"Modulation degree",.Text16
		TextBox 190,84,150,21,.sine_ratio
		Text 30,119,150,14,"Modulation schemes",.Text17
		DropListBox 190,112,150,49,ListArray_ModTech(),.DropListBox_ModTech
		Text 30,147,130,14,"Blanking time [s]",.Text4
		TextBox 190,140,150,21,.dead_dialouge
		Text 30,175,100,14,"Stepwidth [s]",.Text5
		TextBox 190,168,150,21,.step_dialouge
		Text 30,203,120,14,"High voltage [v]",.Text6
		TextBox 190,196,150,21,.peak_dialouge
		Text 30,231,120,14,"Low voltage [v]",.Text15
		TextBox 190,224,150,21,.idle_dialouge
		GroupBox 360,7,330,84,"Transient task properties",.GroupBox2
		Text 370,35,140,14,"Transient task name",.Text7
		TextBox 530,28,150,21,.task_dialouge
		Text 370,63,160,14,"Number of sine period(s)",.Text8
		TextBox 530,56,150,21,.task_num_period
		GroupBox 360,98,160,112,"Port name high side",.GroupBox3
		GroupBox 530,98,160,112,"Port name low side",.GroupBox4
		Text 380,126,40,21,"u high",.Text9
		DropListBox 440,119,70,35,ListArray_Port(),.DropListBox_uhigh
		Text 380,154,40,21,"v high",.Text10
		DropListBox 440,147,70,35,ListArray_Port(),.DropListBox_vhigh
		Text 380,182,50,21,"w high",.Text11
		DropListBox 440,175,70,35,ListArray_Port(),.DropListBox_whigh
		Text 550,126,50,21,"u low",.Text12
		DropListBox 610,119,70,35,ListArray_Port(),.DropListBox_ulow
		Text 550,154,50,21,"v low",.Text13
		DropListBox 610,147,70,35,ListArray_Port(),.DropListBox_vlow
		Text 550,182,50,21,"w low",.Text14
		DropListBox 610,175,70,35,ListArray_Port(),.DropListBox_wlow
		Text 360,217,240,14,"",.TextInfo
		Text 620,217,70,14,"unknown",.tmax_trans_task,1
		OKButton 360,238,90,21
		CancelButton 460,238,90,21
		PushButton 600,238,90,21,"Help",.HelpPB


	End Dialog

	Dim dlg As UserDialog
	If Dialog(dlg)=0 Then Exit All

End Sub
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Dim dfreq_saw As Double
	Dim dfreq_sine As Double
	Dim dstepw As Double
	Dim dt_end As Double
	Dim dn_t As Double
	Dim dpeak As Double
	Dim didle As Double
	Dim dperiod As Double
	Dim dsine_ampl As Double
	Dim ddead_time As Double
	Dim port_counter As Integer
	Dim stask_name As String
	Dim iModTech As Integer
	Dim iNumPeriod As Integer
	'RSJ: add port number
	Dim sport_uhigh As String, sport_vhigh As String, sport_whigh As String
	Dim sport_ulow As String, sport_vlow As String, sport_wlow As String


	Select Case Action%
	Case 1 ' Dialog box initialization
		If numExtPort<=5 Then
			DlgEnable("OK",False)
			DlgText("TextInfo","Error: This macro requires 6 external ports definition")
			DlgText("tmax_trans_task","")
			DlgEnable("DropListBox_uhigh",False)
			DlgEnable("DropListBox_vhigh",False)
			DlgEnable("DropListBox_whigh",False)
			DlgEnable("DropListBox_ulow",False)
			DlgEnable("DropListBox_vlow",False)
			DlgEnable("DropListBox_wlow",False)
		Else
			DlgText("TextInfo","Tmax transient task [s]")
		End If

		DlgText("pwm_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "spwm_dialouge", CStr(10e3)))
		DlgText("peak_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "speak_dialouge", CStr(10)))
		DlgText("step_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sstep_dialouge", Format(CStr(0.05e-6),"Scientific")))
		DlgText("dead_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sdead_dialouge", Format(CStr(0.25e-6),"Scientific")))
		DlgText("task_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "stask_dialouge", "Trans_PWM"))
		DlgText("idle_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sidle_dialouge", CStr(0)))
		DlgText("sine_ratio",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "ssine_ratio", CStr(1)))
		DlgValue("DropListBox_ModTech",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_ModTech", "0")))
		DlgValue("DropListBox_uhigh",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_uhigh", IIf(numExtPort-6>=0,Cstr(numExtPort-6),Cstr(0)))))
		DlgValue("DropListBox_vhigh",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_vhigh", IIf(numExtPort-6>=0,Cstr(numExtPort-5),Cstr(0)))))
		DlgValue("DropListBox_whigh",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_whigh", IIf(numExtPort-6>=0,Cstr(numExtPort-4),Cstr(0)))))
		DlgValue("DropListBox_ulow",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_ulow", IIf(numExtPort-6>=0,Cstr(numExtPort-3),Cstr(0)))))
		DlgValue("DropListBox_vlow",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_vlow", IIf(numExtPort-6>=0,Cstr(numExtPort-2),Cstr(0)))))
		DlgValue("DropListBox_wlow",evaluate(GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_wlow", IIf(numExtPort-6>=0,Cstr(numExtPort-1),Cstr(0)))))
		DlgText("sine_dialouge",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "ssine_dialouge", CStr(500)))
		DlgText("task_num_period",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "stask_num_period", cstr(1)))
		DlgText("tmax_trans_task",GetSetting("CST STUDIO SUITE", "PWM Motor Generator", "stmax_trans_task", Format(cstr(1/500),"0.000e+00")))

	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "OK"
				'Save the dialog box input
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "spwm_dialouge", DlgText("pwm_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "speak_dialouge", DlgText("peak_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sstep_dialouge", DlgText("step_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sdead_dialouge", DlgText("dead_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "stask_dialouge", DlgText("task_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sstep_dialouge", DlgText("step_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sidle_dialouge", DlgText("idle_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "ssine_ratio", DlgText("sine_ratio")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "ssine_dialouge", DlgText("sine_dialouge")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "stask_num_period", DlgText("task_num_period")
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_ModTech", Cstr(DlgValue("DropListBox_ModTech"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_uhigh", Cstr(DlgValue("DropListBox_uhigh"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_vhigh", Cstr(DlgValue("DropListBox_vhigh"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_whigh", Cstr(DlgValue("DropListBox_whigh"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_ulow", Cstr(DlgValue("DropListBox_ulow"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_vlow", Cstr(DlgValue("DropListBox_vlow"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "sDropListBox_wlow", Cstr(DlgValue("DropListBox_wlow"))
				SaveSetting  "CST STUDIO SUITE", "PWM Motor Generator", "stmax_trans_task", DlgText("tmax_trans_task")

				'Disable all input fields
				DlgEnable("pwm_dialouge",False)
				DlgEnable("peak_dialouge",False)
				DlgEnable("step_dialouge",False)
				DlgEnable("dead_dialouge",False)
				DlgEnable("task_dialouge",False)
				DlgEnable("step_dialouge",False)
				DlgEnable("idle_dialouge",False)

				DlgEnable("sine_ratio",False)
				DlgEnable("sine_dialouge",False)
				DlgEnable("step_dialouge",False)
				DlgEnable("task_num_period",False)

				DlgEnable("DropListBox_ModTech",False)
				DlgEnable("DropListBox_uhigh",False)
				DlgEnable("DropListBox_whigh",False)
				DlgEnable("DropListBox_vhigh",False)

				DlgEnable("DropListBox_ulow",False)
				DlgEnable("DropListBox_vlow",False)
				DlgEnable("DropListBox_wlow",False)

				DlgEnable("OK",False)
				DlgEnable("Cancel",False)
				DlgEnable("HelpPB",False)


				dfreq_saw = evaluate(DlgText("pwm_dialouge")) 'pwm base frequency
				stask_name = DlgText("task_dialouge")
				dfreq_sine = evaluate(DlgText("sine_dialouge")) 'sine modulation frequency
				't_end = CDbl(dlg.total_dialouge) 'total signal time
				dstepw = evaluate(DlgText("step_dialouge")) 'stepwidth for signals
				dpeak = evaluate(DlgText("peak_dialouge")) 'peak value for trasistor voltage
				dsine_ampl = evaluate(DlgText("sine_ratio")) 'sine amplitude allows for pwm signals with no full on or full off behaviour
				ddead_time = evaluate(DlgText("dead_dialouge")) 'dead time between signals, make sure that it is a multiple of stepwidth
				didle = evaluate(DlgText("idle_dialouge"))
				iNumPeriod = evaluate(DlgText("task_num_period"))
				dt_end = IIf(ModDiv(dfreq_saw,dfreq_sine)=0,CDbl(1/dfreq_sine),iNumPeriod*CDbl(1/dfreq_sine))
				'dt_end = iNumPeriod*CDbl(1/dfreq_sine)
				dperiod = dt_end * (1/dfreq_saw * 1/dstepw)
				iModTech = evaluate(DlgValue("DropListBox_ModTech"))

				'RSJ: Add port selection
				sport_uhigh = ListArray_Port(DlgValue("DropListBox_uhigh"))
				sport_vhigh = ListArray_Port(DlgValue("DropListBox_vhigh"))
				sport_whigh = ListArray_Port(DlgValue("DropListBox_whigh"))
				sport_ulow = ListArray_Port(DlgValue("DropListBox_ulow"))
				sport_vlow = ListArray_Port(DlgValue("DropListBox_vlow"))
				sport_wlow = ListArray_Port(DlgValue("DropListBox_wlow"))

				DlgText("TextInfo","Progress: Generating PWM signals")
				DlgText("tmax_trans_task","")
				CreatePWMSignal(dfreq_saw,stask_name,dfreq_sine,dstepw,dpeak,dsine_ampl,ddead_time,didle,dt_end,dperiod,iNumPeriod,iModTech,sport_uhigh,sport_vhigh,sport_whigh,sport_ulow,sport_vlow,sport_wlow)
				DlgText("TextInfo","Progress: Done")
				DlgText("tmax_trans_task","")
				Wait 1

			Case "Cancel"
				DialogFunction=False
				Exit All
			Case "HelpPB"
				DialogFunction = True
				StartHelp HelpFileName
		End Select
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
			Case "task_num_period"
				If Cint(DlgText("task_num_period")) <= 0 Then
					MsgBox("Please assign number of period > 0",vbInformation,"PWM Generator")
					DialogFunction = True
				End If
		End Select
		DlgText("tmax_trans_task",Format(cstr(1/cdbl(DlgText("sine_dialouge"))*Cint(DlgText("task_num_period"))),"0.000e+00"))
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select

End Function

Sub CreatePWMSignal(freq_saw As Double,task_name As String,freq_sine As Double,stepw As Double,peak As Double,sine_ampl As Double,dead_time As Double,idle As Double, _
                    t_end As Double,period As Double,NumPeriod As Integer,ModTech As Integer,port_uhigh As String,port_vhigh As String,port_whigh As String,port_ulow As String,port_vlow As String,port_wlow As String)
	Dim SigSourceName As String
	'RSJ: n_samples has to be in Long
	Dim n_samples As Long
	Dim ii As Long
	Dim out_dir As String
	Dim Tmax As Double

	SigSourceName="PWM="+Cstr(freq_saw)+"Hz_"+"Sin="+Cstr(freq_sine)+"Hz_"+"Mod="+ListArray_ModTech(ModTech)
	n_samples = Round(t_end/stepw)

	'out_dir = "D:\Desktop\Conferences\2017 PWM Macro\"
	out_dir = 	GetProjectPath("ModelDS")

	Dim ruh(),rul(),rvh(),rvl(),rwh(),rwl()
	ReDim ruh(n_samples) : ReDim rul(n_samples) : ReDim rvh(n_samples) : ReDim rvl(n_samples) : ReDim rwh(n_samples) : ReDim rwl(n_samples)
	Dim ruh2(),rul2(),rvh2(),rvl2(),rwh2(),rwl2()
	ReDim ruh2(n_samples) : ReDim rul2(n_samples) : ReDim rvh2(n_samples) : ReDim rvl2(n_samples) : ReDim rwh2(n_samples) : ReDim rwl2(n_samples)
	Dim TriWaveShift As Double

	'RSJ: Added carrier wave shift parameter
	TriWaveShift = 0 'Triwave start at 1
	'TriWaveShift = 3/freq_saw/4 'Triwave start at 0
	'TriWaveShift = 2/freq_saw/4 'Triwave start at -1

	'create sawtooth

	For ii = 0 To n_samples-1
		ruh(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rul(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)
		rvh(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rvl(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)
		rwh(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rwl(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)

		ruh(ii) = IIf(Abs(ruh(ii))<1e-6,0,ruh(ii))
		rul(ii) = IIf(Abs(rul(ii))<1e-6,0,rul(ii))
		rvh(ii) = IIf(Abs(rvh(ii))<1e-6,0,rvh(ii))
		rvl(ii) = IIf(Abs(rvl(ii))<1e-6,0,rvl(ii))
		rwh(ii) = IIf(Abs(rwh(ii))<1e-6,0,rwh(ii))
		rwl(ii) = IIf(Abs(rwl(ii))<1e-6,0,rwl(ii))
	Next

	For ii = 0 To n_samples-1
		ruh2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rul2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)
		rvh2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rvl2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)
		rwh2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,1,-1, 0.5/freq_saw, 0.5/freq_saw)
		rwl2(ii) = Tri_Wave(ii*stepw+(0*period*stepw)+dead_time+TriWaveShift,-1,1, 0.5/freq_saw, 0.5/freq_saw)
		ruh2(ii) = IIf(Abs(ruh2(ii))<1e-6,0,ruh2(ii))
		rul2(ii) = IIf(Abs(rul2(ii))<1e-6,0,rul2(ii))
		rvh2(ii) = IIf(Abs(rvh2(ii))<1e-6,0,rvh2(ii))
		rvl2(ii) = IIf(Abs(rvl2(ii))<1e-6,0,rvl2(ii))
		rwh2(ii) = IIf(Abs(rwh2(ii))<1e-6,0,rwh2(ii))
		rwl2(ii) = IIf(Abs(rwl2(ii))<1e-6,0,rwl2(ii))
	Next

	Dim cuh(),cul(),cvh(),cvl(),cwh(),cwl()
	ReDim cuh(n_samples) : ReDim cul(n_samples) : ReDim cvh(n_samples) : ReDim cvl(n_samples) : ReDim cwh(n_samples) : ReDim cwl(n_samples)

	Dim Vk As Double, amp_factor As Double, uu As Double, vv As Double, ww As Double, Vcommon_voltage As Double

	If ModTech >= 1 Then
		sine_ampl = sine_ampl*(2/Sqr(3)) ' Maximum modulation index to have similar amplitude as Sinusoidal PWM
	End If

	For ii = 0 To n_samples-1
		'RSJ: added for THIPWM (Third Harmonic Injection/Adder PWM) and SVPWM
		If ModTech<2 Then
			If ModTech = 1 Then
				Vk = k_factor * Sin(3*freq_sine*2*pi*ii*stepw)
			ElseIf ModTech = 0 Then
				Vk = 0
			End If
			cuh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw)+Vk)
			cul(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+pi)-Vk)
			cvh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+2/3*pi)+Vk)
			cvl(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+2/3*pi+pi)-Vk)
			cwh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+4/3*pi)+Vk)
			cwl(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+4/3*pi+pi)-Vk)
		Else
			'SVPWM
			uu = Sin(freq_sine*2*pi*ii*stepw)
			vv = Sin(freq_sine*2*pi*ii*stepw+2/3*pi)
			ww = Sin(freq_sine*2*pi*ii*stepw+4/3*pi)
			Vcommon_voltage = CalcVcommon(uu,vv,ww)
			cuh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw) - Vcommon_voltage)
			cul(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+pi) + Vcommon_voltage)
			cvh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+2/3*pi) - Vcommon_voltage)
			cvl(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+2/3*pi+pi) + Vcommon_voltage)
			cwh(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+4/3*pi) - Vcommon_voltage)
			cwl(ii) = sine_ampl * (Sin(freq_sine*2*pi*ii*stepw+4/3*pi+pi) + Vcommon_voltage)
		End If
		cuh(ii)=IIf (Abs(cuh(ii))<1e-6,0,cuh(ii))
		cul(ii)=IIf (Abs(cul(ii))<1e-6,0,cul(ii))
		cvh(ii)=IIf (Abs(cvh(ii))<1e-6,0,cvh(ii))
		cvl(ii)=IIf (Abs(cvl(ii))<1e-6,0,cvl(ii))
		cwh(ii)=IIf (Abs(cwh(ii))<1e-6,0,cwh(ii))
		cwl(ii)=IIf (Abs(cwl(ii))<1e-6,0,cwl(ii))
	Next

	Dim pwm_uh(),pwm_ul(),pwm_vh(),pwm_vl(),pwm_wh(),pwm_wl()
	ReDim pwm_uh(n_samples) : ReDim pwm_ul(n_samples) : ReDim pwm_vh(n_samples) : ReDim pwm_vl(n_samples) : ReDim pwm_wh(n_samples) : ReDim pwm_wl(n_samples) :

	For ii = 0 To n_samples-1
		If (cuh(ii) >= ruh(ii) And cuh(ii) >= ruh2(ii)) Then pwm_uh(ii) = peak Else pwm_uh(ii) = idle
		If (cul(ii) >= rul(ii) And cul(ii) >= rul2(ii)) Then pwm_ul(ii) = peak Else pwm_ul(ii) = idle
		If (cvh(ii) >= rvh(ii) And cvh(ii) >= rvh2(ii)) Then pwm_vh(ii) = peak Else pwm_vh(ii) = idle
		If (cvl(ii) >= rvl(ii) And cvl(ii) >= rvl2(ii)) Then pwm_vl(ii) = peak Else pwm_vl(ii) = idle
		If (cwh(ii) >= rwh(ii) And cwh(ii) >= rwh2(ii)) Then pwm_wh(ii) = peak Else pwm_wh(ii) = idle
		If (cwl(ii) >= rwl(ii) And cwl(ii) >= rwl2(ii)) Then pwm_wl(ii) = peak Else pwm_wl(ii) = idle
	Next

	'---------------------------------r
	Dim my_plot_ruh As Object,my_plot_rul As Object, my_plot_rvh As Object,my_plot_rvl As Object,my_plot_rwh As Object,my_plot_rwl As Object
	Set my_plot_ruh = DS.Result1D("") : Set my_plot_rul = DS.Result1D("") : Set my_plot_rvh = DS.Result1D("") : Set my_plot_rvl = DS.Result1D("") : Set my_plot_rwh = DS.Result1D("") :Set my_plot_rwl = DS.Result1D("")
	my_plot_ruh.Initialize(n_samples):my_plot_rul.Initialize(n_samples):my_plot_rvh.Initialize(n_samples):my_plot_rvl.Initialize(n_samples):my_plot_rwh.Initialize(n_samples):my_plot_rwl.Initialize(n_samples):

	Dim my_plot_rvl2 As Object
	Set my_plot_rvl2 = DS.Result1D("")
	my_plot_rvl2.Initialize(n_samples)

	For ii = 0 To n_samples-1
		my_plot_ruh.SetXY (ii,ii*stepw,ruh(ii))
		my_plot_rul.SetXY (ii,ii*stepw,rul(ii))
		my_plot_rvh.SetXY (ii,ii*stepw,rvh(ii))
		my_plot_rvl.SetXY (ii,ii*stepw,rvl(ii))
		my_plot_rvl2.SetXY (ii,ii*stepw,rvl2(ii))
		my_plot_rwh.SetXY (ii,ii*stepw,rwh(ii))
		my_plot_rwl.SetXY (ii,ii*stepw,rwl(ii))
	Next
	my_plot_ruh.AddToTree ( "Results/uh/ruh" )
	my_plot_rul.AddToTree ( "Results/ul/rul" )
	my_plot_rvh.AddToTree ( "Results/vh/rvh" )
	my_plot_rvl.AddToTree ( "Results/vl/rvl" )
	my_plot_rvl2.AddToTree ( "Results/vl/rvl2" )
	my_plot_rwh.AddToTree ( "Results/wh/rwh" )
	my_plot_rwl.AddToTree ( "Results/wl/rwl" )

	'--------------------------------c
	Dim my_plot_cuh As Object, my_plot_cul As Object, my_plot_cvh As Object,my_plot_cvl As Object,my_plot_cwh As Object,my_plot_cwl As Object
	Set my_plot_cuh = DS.Result1D("") : Set my_plot_cul = DS.Result1D("") : Set my_plot_cvh = DS.Result1D("") : Set my_plot_cvl = DS.Result1D("") : Set my_plot_cwh = DS.Result1D("") :Set my_plot_cwl = DS.Result1D("")
	my_plot_cuh.Initialize(n_samples) : my_plot_cul.Initialize(n_samples):my_plot_cvh.Initialize(n_samples):my_plot_cvl.Initialize(n_samples):my_plot_cwh.Initialize(n_samples):my_plot_cwl.Initialize(n_samples):

	For ii = 0 To n_samples-1
		my_plot_cuh.SetXY (ii,ii*stepw,cuh(ii))
		my_plot_cul.SetXY (ii,ii*stepw,cul(ii))
		my_plot_cvh.SetXY (ii,ii*stepw,cvh(ii))
		my_plot_cvl.SetXY (ii,ii*stepw,cvl(ii))
		my_plot_cwh.SetXY (ii,ii*stepw,cwh(ii))
		my_plot_cwl.SetXY (ii,ii*stepw,cwl(ii))
	Next

	my_plot_cuh.AddToTree ( "Results/uh/cuh" )
	my_plot_cul.AddToTree ( "Results/ul/cul" )
	my_plot_cvh.AddToTree ( "Results/vh/cvh" )
	my_plot_cvl.AddToTree ( "Results/vl/cvl" )
	my_plot_cwh.AddToTree ( "Results/wh/cwh" )
	my_plot_cwl.AddToTree ( "Results/wl/cwl" )


	'-------------------------------------------pwm
	Dim my_plot_pwm_uh As Object,my_plot_pwm_ul As Object, my_plot_pwm_vh As Object,my_plot_pwm_vl As Object,my_plot_pwm_wh As Object,my_plot_pwm_wl As Object
	Set my_plot_pwm_uh = DS.Result1D("") : Set my_plot_pwm_ul = DS.Result1D("") : Set my_plot_pwm_vh = DS.Result1D("") : Set my_plot_pwm_vl = DS.Result1D("") : Set my_plot_pwm_wh = DS.Result1D("") :Set my_plot_pwm_wl = DS.Result1D("")
	my_plot_pwm_uh.Initialize(n_samples):my_plot_pwm_ul.Initialize(n_samples):my_plot_pwm_vh.Initialize(n_samples):my_plot_pwm_vl.Initialize(n_samples):my_plot_pwm_wh.Initialize(n_samples):my_plot_pwm_wl.Initialize(n_samples):

	'RSJ: ASCII file contains also task name, to avoid overwritting for different settings.
	Open out_dir + task_name + "_pwm_uh.txt" For Output As #1
	Open out_dir + task_name + "_pwm_ul.txt" For Output As #2
	Open out_dir + task_name + "_pwm_vh.txt" For Output As #3
	Open out_dir + task_name + "_pwm_vl.txt" For Output As #4
	Open out_dir + task_name + "_pwm_wh.txt" For Output As #5
	Open out_dir + task_name + "_pwm_wl.txt" For Output As #6


	For ii = 0 To n_samples-1

		my_plot_pwm_uh.SetXY (ii,ii*stepw,pwm_uh(ii)) : Print #1, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_uh(ii))
		my_plot_pwm_ul.SetXY (ii,ii*stepw,pwm_ul(ii)) : Print #2, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_ul(ii))
		my_plot_pwm_vh.SetXY (ii,ii*stepw,pwm_vh(ii)) : Print #3, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_vh(ii))
		my_plot_pwm_vl.SetXY (ii,ii*stepw,pwm_vl(ii)) : Print #4, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_vl(ii))
		my_plot_pwm_wh.SetXY (ii,ii*stepw,pwm_wh(ii)) : Print #5, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_wh(ii))
		my_plot_pwm_wl.SetXY (ii,ii*stepw,pwm_wl(ii)) : Print #6, CStr(ii*stepw)  + vbTab + vbTab + vbTab  +CStr(pwm_wl(ii))
	Next

	Close 1 : Close 2 : Close 3 : Close 4 : Close 5 : Close 6

	my_plot_pwm_uh.AddToTree ( "Results/uh/pwm_uh" ) : my_plot_pwm_uh.AddToTree ( "Results/pwm/pwm_uh" )
	my_plot_pwm_ul.AddToTree ( "Results/ul/pwm_ul" ) : my_plot_pwm_ul.AddToTree ( "Results/pwm/pwm_ul" )
	my_plot_pwm_vh.AddToTree ( "Results/vh/pwm_vh" ) : my_plot_pwm_vh.AddToTree ( "Results/pwm/pwm_vh" )
	my_plot_pwm_vl.AddToTree ( "Results/vl/pwm_vl" ) : my_plot_pwm_vl.AddToTree ( "Results/pwm/pwm_vl" )
	my_plot_pwm_wh.AddToTree ( "Results/wh/pwm_wh" ) : my_plot_pwm_wh.AddToTree ( "Results/pwm/pwm_wh" )
	my_plot_pwm_wl.AddToTree ( "Results/wl/pwm_wl" ) : my_plot_pwm_wl.AddToTree ( "Results/pwm/pwm_wl" )

	Dim ASCII_import_argument(1) As String

	If ModDiv(freq_saw,freq_sine)=0 Then
		ASCII_import_argument(1) = "true"   'activate the periodicity
		Tmax = t_end*Units.GetTimeSIToUnit * NumPeriod
	Else
		ASCII_import_argument(1) = "false"  'don't use the periodicity
		Tmax = t_end*Units.GetTimeSIToUnit
	End If

	'Check if task exists
	SimulationTask.Name(task_name)
	If SimulationTask.DoesExist Then
		ReportInformationToWindow("Task " +task_name+ " exists")
	Else
		With SimulationTask
		 .Reset
		 .Type ("Transient")
		 .Name (task_name)
		 .SetProperty ( "tmax", Tmax) ' set Tmax for the transient task
		 .SetProperty ( "fmax estimator", "transitiontime" )
		 .Create
		End With
		ReportInformationToWindow("Task "+ task_name+" created")
	End If
		'RSJ: Write also Tmax for existing Transient project and use transition time based as fmax estimator. Automatic fmax estimator has a bug with ASCII import
		'RSJ: Assign port number from the selection list
		ASCII_import_argument(0) = out_dir + task_name +"_pwm_uh.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetProperty ( "tmax", Tmax) ' set Tmax for the transient task
			If .GetProperty ( "fmax estimator") = "automatic" Then
				.SetProperty ( "fmax estimator", "transitiontime" )
			Else
					'do nothing
			End If
			.SetPortSignal (port_uhigh,"Import", ASCII_import_argument )
			.SetPortSignalName (port_uhigh,"uhigh "+SigSourceName)
			.SetPortInnerResistance ( port_uhigh, "0.0" )
			.SetPortSourceType(port_uhigh, "Voltage")
		End With
		ASCII_import_argument(0) = out_dir + task_name +"_pwm_vh.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetPortSignal (port_vhigh,"Import", ASCII_import_argument )
			.SetPortSignalName (port_vhigh,"vhigh "+SigSourceName)
			.SetPortInnerResistance ( port_vhigh, "0.0" )
			.SetPortSourceType(port_vhigh, "Voltage")
		End With
		ASCII_import_argument(0) = out_dir + task_name +"_pwm_wh.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetPortSignal (port_whigh,"Import", ASCII_import_argument )
			.SetPortSignalName (port_whigh,"whigh "+SigSourceName)
			.SetPortInnerResistance ( port_whigh, "0.0" )
			.SetPortSourceType(port_whigh, "Voltage")
		End With
		ASCII_import_argument(0) = out_dir + task_name +"_pwm_ul.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetPortSignal (port_ulow,"Import", ASCII_import_argument )
			.SetPortSignalName (port_ulow,"ulow "+SigSourceName)
			.SetPortInnerResistance ( port_ulow, "0.0" )
			.SetPortSourceType(port_ulow, "Voltage")
		End With
		ASCII_import_argument(0) = out_dir + task_name +"_pwm_vl.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetPortSignal (port_vlow,"Import", ASCII_import_argument )
			.SetPortSignalName (port_vlow,"vlow "+SigSourceName)
			.SetPortInnerResistance ( port_vlow, "0.0" )
			.SetPortSourceType(port_vlow, "Voltage")
		End With
		ASCII_import_argument(0) = out_dir + task_name + "_pwm_wl.txt"
		With SimulationTask
			.Reset
			.Name (task_name)
			.SetPortSignal (port_wlow,"Import", ASCII_import_argument )
			.SetPortSignalName (port_wlow,"wlow "+SigSourceName)
			.SetPortInnerResistance (port_wlow, "0.0" )
			.SetPortSourceType(port_wlow, "Voltage")
		End With


	ReportInformationToWindow("Macro ""Create PWM-signals"" finished")

End Sub

Function Tri_Wave(t, V1, V2, T1, T2)

	' *************************************************************
	' Generate Triangle Wave
	'
	' t - time
	' V1 - voltage level 1 (initial voltage)
	' V2 - voltage level 2
	' T1 - period ramping from V1 to V2
	' T2 - period ramping from V2 to V1
	'***************************************************************

	Dim t_tri, dV_dt1, dV_dt2 As Double
	Dim N As Single

	' Calculate voltage rates of change (slopes) during T1 and T2
	dV_dt1 = (V2 - V1) / T1
	dV_dt2 = (V1 - V2) / T2

	' given t, how many full cycles have occurred
	N = Int(t / (T1 + T2))

	' calc the time point in the current triangle wave
	t_tri = t - (T1 + T2) * N

	' if during T1, calculate triangle value using V1 and dV_dt1
	If t_tri <= T1 Then
	    Tri_Wave = V1 + dV_dt1 * t_tri

	' if during T2, calculate triangle value using V2 and dV_dt2
	Else
	   Tri_Wave = V2 + dV_dt2 * (t_tri - T1)

	End If


End Function

Function CalcVcommon(uuu As Double,vvv As Double,www As Double) As Double
	Dim ValMin As Double
	Dim ValMax As Double
	If uuu < vvv Then
		ValMax = vvv
		ValMin = uuu
	Else
		ValMax = uuu
		ValMin = vvv
	End If

	If ValMax < www Then
		ValMax = www
	ElseIf ValMin > www Then
		ValMin = www
	End If
	CalcVcommon = (ValMax + ValMin)*0.5
End Function

Function FillPortNameArray () As Integer
	Dim num_port As Integer, portname As String, ii As Integer
	num_port = ExternalPort.StartPortNameIteration
	ReDim ListArray_Port(num_port)

	For ii=0 To num_port-1
		ListArray_Port(ii)=ExternalPort.GetNextPortName
	Next
	FillPortNameArray = num_port
End Function

Function ModDiv(a As Double, b As Double) As Double
	ModDiv=(a/b)-Fix(a/b)
End Function
