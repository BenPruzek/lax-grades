'#Language "WWB-COM"

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2021-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 06-Dec-2023 rsj: Move the command evaluate(dlgtext) to section pushbutton1, and show error message for evaluating unknown input parameter.
' 06-Oct-2023 rsj: Add option to delay the signal
'                  0% spread -> use updated digital signal definition from CST-v2024.
'                  Modulation period repetition is now only used to determine Tmax. 
'                  CST-v2024: ASCII import can be set to periodic and delayed, ASCII filename now contains task name to avoid overwriting.
' 19-May-2022 rsj: Slightly GUI label rework
' 19-Apr-2022 rsj: Create port in schematic, if this is not yet available
' 29-Mar-2022 rsj: GUI rework, add dead time function, initial amplitude definition, adapt the online help content,
'                  Set fmax transition time estimator in transient task if not yet set.
'                  Set signal name accordingly for better overview
'                  Select Treeitem after apply for quick signal review
' 11-Jan-2022 ube: tiny changes, online Help page added
' 14-Dec-2021 rsj: First version
'-----------------------------------------------------------------------------------------------------------------------------

Const ListSpreadProfile = Array ("Up-Spread", "Center-Spread", "Down-Spread")
Const SHOW_SIGNAL_IN_TREE = 1  'Set 1 to plot the time signal in the result tree.
Type TimeSignalProp
	Vlow As Double
	Vhigh As Double
	trise As Double
	tfall As Double
	tdead As Double
    dutycycle As Double
    tdelay As Double
    freqsw As Double
End Type
Type TransTaskProp
	taskname As String
	portname As String
	portnamehigh As String
	portnamelow As String
End Type
Type SpreadProp
	freqmod As Double
	rate As Double
	profile As String
End Type
Dim out_dir As String


Sub Main
	Begin Dialog UserDialog 750,392,"Spread Spectrum Clock Generation (SSCG)",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,360,196,"Switching signal properties",.SWSigProp
		GroupBox 10,210,360,112,"Spread spectrum properties",.SS_prop
		Text 30,35,120,14,"Frequency [Hz]",.diag_freq_sw2
		TextBox 230,28,130,21,.text_freqsw
		Text 30,63,110,14,"Duty cycle in %",.diag_dutycycle
		TextBox 230,56,130,21,.text_dutycycle
		Text 30,91,110,14,"Rise time [s]",.diag_risetime
		TextBox 230,84,130,21,.text_risetime
		Text 30,119,110,14,"Fall time [s]",.diag_falltime
		TextBox 230,112,130,21,.text_falltime
		Text 30,147,120,14,"Amplitude high",.diag_vhigh
		TextBox 230,140,130,21,.text_vhigh
		Text 30,175,120,14,"Amplitude low",.diag_vlow
		TextBox 230,168,130,21,.text_vlow
		GroupBox 380,7,360,315,"Excitation settings",.GroupBox2
		Text 30,238,180,14,"Frequency modulation [Hz]",.diag_freq_mod
		TextBox 230,231,130,21,.text_fmod
		Text 30,266,180,14,"Spreading variation in %",.diag_spreadrate
		TextBox 230,259,130,21,.text_spreadpercent
		Text 30,294,120,14,"Spreading type",.diag_spreadprofile
		DropListBox 230,287,130,63,ListSpreadProfile(),.ListBoxSpreadProfile
		OptionGroup .GroupExcitationType
			OptionButton 400,84,180,14,"Single switch",.OptionButtonSE
			OptionButton 400,161,210,14,"Complementary switch pair(s)",.OptionButtonComplement
		OptionGroup .GroupInitAmp
			OptionButton 600,140,60,14,"High",.OptionButtonInitHigh
			OptionButton 680,140,50,14,"Low",.OptionButtonInitLow
		Text 420,140,120,14,"Initial amplitude",.diag_init_amp
		TextBox 610,182,120,21,.text_deadtime
		Text 420,189,180,14,"Dead time for low side [s]",.Text_dead_time
		Text 400,35,140,14,"Transient task name",.diag_taskname
		TextBox 610,28,120,21,.text_taskname
		Text 420,112,90,14,"External port",.diag_assignport
		Text 420,217,180,14,"External port(s) for high side",.diag_assignport_high
		Text 420,245,170,14,"External port(s) for low side",.diag_assignport_low
		TextBox 610,105,120,21,.text_assignport
		TextBox 610,210,120,21,.text_assignport_high
		TextBox 610,238,120,21,.text_assignport_low
		Text 400,63,190,14,"Number of modulation periods",.diag_modrep
		TextBox 610,56,120,21,.text_modrep
		CheckBox 400,294,90,14,"Use delay",.CheckBoxUseDelay
		PushButton 610,287,120,21,"Delay properties",.PushButtonDelayProp
		Text 380,329,190,14,"Tmax used in trans task [s]",.diag_tmaxused
		Text 630,329,110,14,"",.text_tmaxused
		PushButton 380,357,90,21,"Apply",.PushButton1
		PushButton 480,357,90,21,"Close",.PushButton2
		Text 20,329,350,28,"Note: Please use resolution bandwidth in transient task for at least half of the modulation frequency",.Text1
		PushButton 660,357,80,21,"Help",.Help
		CancelButton 590,357,60,21
		Text 400,266,320,14,"Note: Use semicolon as separator for multiple ports",.Text2


	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	'out_dir = 	GetProjectPath("ModelDS")

	Static timesig As TimeSignalProp
	Static ssprop As SpreadProp
	Static transprop As TransTaskProp
	Static fspread_list() As Double
	Static t_end As Double
	Static fmod_rep As Long
	Static nsample_fspread As Long
	Static InitAmp As Integer
	Static ComplementSig As Integer
	Dim sTime As String, separationstring As String, modifiedportname_Delay As String, modifiedportname_HS As String, modifiedportname_LS As String, modifiedportname_HSDelay As String, modifiedportname_LSDelay As String
	Dim modifiedportname As String
	Dim SignalName As String
	Dim ii As Integer
	Dim portnameArray_HS() As String, portnameArray_LS() As String, portnameArray_HSDelay() As String, portnameArray_LSDelay() As String, portnameArray() As String, portnameArrayDelay() As String
	Dim UseDelay As Integer
	Static DelayType As Integer
	Static DelayLen As Double
	Static DelayWhichPorts As Integer


	DlgVisible("Cancel", False) ' hide it, but it needs to be there for the red "X" to work

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("text_freqsw",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_freqsw", "425000"))
		DlgText("text_dutycycle",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_dutycycle", "50"))
		DlgText("text_risetime",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_risetime", "20e-9"))
		DlgText("text_falltime",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_falltime", "20e-9"))
		DlgText("text_deadtime",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_deadtime", "0"))
		DlgText("text_vhigh",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_vhigh", "1"))
		DlgText("text_vlow",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_vlow", "0"))
		DlgText("text_fmod",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_fmod", "2250"))
		DlgText("text_spreadpercent",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_spreadpercent", "12"))
		DlgValue("ListBoxSpreadProfile",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "ListBoxSpreadProfile", 1))
		DlgText("text_taskname",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_taskname", "Tran1"))
		DlgText("text_assignport",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport", "1"))
		DlgText("text_modrep",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_modrep", "2"))
		DlgText("text_tmaxused","unknown")
		DlgValue("GroupInitAmp",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupInitAmp", 0))
		DlgValue("GroupExcitationType",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupExcitationType", 0))
		DlgText("text_assignport_high",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport_high", "1"))
		DlgText("text_assignport_low",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport_low", "2"))
		DlgValue("CheckBoxUseDelay",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "CheckBoxUseDelay", "0"))
		ComplementSig=evaluate(DlgValue("GroupExcitationType"))
		DialogFunc_EcitationType(ComplementSig)
		UseDelay=evaluate(DlgValue("CheckBoxUseDelay"))

		If UseDelay=1 Then
			DlgEnable("PushButtonDelayProp",True)
		Else
			DlgEnable("PushButtonDelayProp",False)
		End If

	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "PushButton1"
			On Error GoTo ShowError
				timesig.freqsw=evaluate(DlgText("text_freqsw"))
				timesig.dutycycle=evaluate(DlgText("text_dutycycle"))
				timesig.trise=evaluate(DlgText("text_risetime"))
				timesig.tfall=evaluate(DlgText("text_falltime"))
				timesig.tdead=evaluate(DlgText("Text_deadtime"))
				If timesig.tdead <0 Then
					MsgBox("Deadtime has to be larger than 0",vbCritical,"Spread Spectrum Clock Generation (SSCG)")
					DlgText("text_deadtime",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "text_deadtime", "0"))
					GoTo BackToDialog
				End If
				timesig.Vhigh=evaluate(DlgText("text_vhigh"))
				timesig.Vlow=evaluate(DlgText("text_vlow"))
				ssprop.freqmod=evaluate(DlgText("text_fmod"))
				ssprop.rate=evaluate(DlgText("text_spreadpercent"))
				ssprop.profile=ListSpreadProfile(DlgValue("ListBoxSpreadProfile"))
				transprop.taskname=DlgText("text_taskname")
				transprop.portname=DlgText("text_assignport")
				transprop.portnamehigh=DlgText("text_assignport_high")
				transprop.portnamelow=DlgText("text_assignport_low")
				fmod_rep=Round(evaluate(DlgText("text_modrep")))
				nsample_fspread=Round(timesig.freqsw/ssprop.freqmod)
				InitAmp=evaluate(DlgValue("GroupInitAmp"))
			
			On Error GoTo 0
				GoTo EverythingOK
				
			ShowError:
				MsgBox("Error in evaluating the input fields. Please check all input expressions,"+vbNewLine+"especially if all used parameters are properly defined in the parameter list.",vbCritical,"Spread Spectrum Clock Generation (SSCG)")
				DialogFunc = True
				Exit Function
				
			EverythingOK:	

			'Save the dialog box input
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_freqsw", DlgText("text_freqsw")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_dutycycle", DlgText("text_dutycycle")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_risetime", DlgText("text_risetime")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_falltime", DlgText("text_falltime")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_deadtime", DlgText("text_deadtime")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_vhigh", DlgText("text_vhigh")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_vlow", DlgText("text_vlow")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_fmod", DlgText("text_fmod")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_spreadpercent", DlgText("text_spreadpercent")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_taskname", DlgText("text_taskname")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport", DlgText("text_assignport")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_modrep", DlgText("text_modrep")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "ListBoxSpreadProfile", DlgValue("ListBoxSpreadProfile")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "GroupInitAmp", DlgValue("GroupInitAmp")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "GroupExcitationType", DlgValue("GroupExcitationType")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport_high", DlgText("text_assignport_high")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "text_assignport_low", DlgText("text_assignport_low")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "CheckBoxUseDelay", DlgValue("CheckBoxUseDelay")

			'Grouping port naming for HS, LS, HS_delay and LS_delay
			portnameArray_HS=Split(transprop.portnamehigh,";")
			portnameArray_LS=Split(transprop.portnamelow,";")

			'Delay Properties
			UseDelay=Evaluate(GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "CheckBoxUseDelay", "0"))
			If UseDelay=1 Then
				DelayType=evaluate(GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayType", "0"))
				DelayLen=evaluate(GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "DelayLength", "0"))
				If DelayType=0 Then
					DelayLen=(1/timesig.freqsw)*(DelayLen/360)
				End If
				DelayWhichPorts=evaluate(GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayWhichPort", "0")) '0=all ports; 1=user defined port label
				Select Case DelayWhichPorts
				Case 0
					portnameArray_HSDelay=portnameArray_HS
					portnameArray_LSDelay=portnameArray_LS
					modifiedportname_HSDelay=Replace(transprop.portnamehigh,";","_")+"_"
					modifiedportname_LSDelay=Replace(transprop.portnamelow,";","_")+"_"
					modifiedportname_HS=""
					modifiedportname_LS=""
				Case 1
					modifiedportname_Delay=GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "DelayPortLabel", "")
					portnameArray_HSDelay=FindPortNameWithDelay(portnameArray_HS,Split(modifiedportname_Delay,";"))              'portnameArray_HS will be modifed
					portnameArray_LSDelay=FindPortNameWithDelay(portnameArray_LS,Split(modifiedportname_Delay,";"))              'portnameArray_LS will be modified

					If UBound(portnameArray_HSDelay)=0 Then
						modifiedportname_HSDelay=portnameArray_HSDelay(0)+"_"
					Else
						For ii=0 To UBound(portnameArray_HSDelay)
							modifiedportname_HSDelay=modifiedportname_HSDelay+portnameArray_HSDelay(ii)+"_"
						Next
					End If
					If UBound(portnameArray_LSDelay)=0 Then
						modifiedportname_LSDelay=portnameArray_LSDelay(0)+"_"
					Else
						For ii=0 To UBound(portnameArray_LSDelay)
							modifiedportname_LSDelay=modifiedportname_LSDelay+portnameArray_LSDelay(ii)+"_"
						Next
					End If
					If modifiedportname_HSDelay="" And modifiedportname_LSDelay="" Then
						MsgBox("No valid port name to apply the delay. Please check the setting.",vbCritical,"Spread Spectrum Clock Generation (SSCG)")
						GoTo BackToDialog
					End If
				End Select
				timesig.tdelay=DelayLen
			Else
				timesig.tdelay=0.0
			End If

			'RSJ: 06-sept-2023: for 0% spread, use a standard signal definition V2024
			If ssprop.rate<>0 Then
				fspread_list=GenerateFspreadList(timesig.freqsw,ssprop,nsample_fspread)
			End If

			'RSJ: 05-sept-2023: V2024 ASCII import can now use periodic option and delay
			'                   t_end*fmod_rep used now for setting the transient task tmax.
			'                   CreateSpreadSignal argument fmod_rep set to 1
			'Single switch signal definition
			If ComplementSig = 0 Then
				portnameArray=Split(transprop.portname,";")
				If UBound(portnameArray)=0 Then
					'Definition correct. Only one port name is defined.
				Else
					MsgBox("Please define only one port name for single switch excitation.",vbCritical,"Spread Spectrum Clock Generation (SSCG)")
					GoTo BackToDialog
				End If

				out_dir = GetProjectPath("ModelDS")
				out_dir = out_dir+transprop.taskname+"_Port_"+transprop.portname+"_"

				If ssprop.rate<>0 Then
					t_end=CreateSpreadSignal(fspread_list,timesig,1,InitAmp,0.0)
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_"+cstr(ssprop.rate)+"%_"+ssprop.profile
					SignalName = SignalName + IIf(timesig.tdelay=0,"","_delayed")
					AssignTransientTask (ComplementSig,transprop.portname,transprop.taskname,t_end*fmod_rep,timesig.tdelay,out_dir,SignalName)
					If timesig.tdelay<>0 Then
						DS.SelectTreeItem("Results\SpreadSignal\TimeSignal_"+IIf(CBool(InitAmp),"InitLow_Delayed","InitHigh_Delayed"))
					Else
						DS.SelectTreeItem("Results\SpreadSignal\TimeSignal_"+IIf(CBool(InitAmp),"InitLow","InitHigh"))
					End If
					With Plot1D
						.XRange(0,3/timesig.freqsw)
						.Plot
					End With
					DlgText("text_tmaxused",cstr(Format(t_end*fmod_rep+timesig.tdelay,"0.0000e+00")))
				Else
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_NoSpread"
					SignalName = SignalName + IIf(timesig.tdelay=0,"","_delayed")
					AssignTransienttask_DigitalSignal(transprop,timesig,InitAmp,fmod_rep,"","",SignalName)
					DlgText("text_tmaxused",cstr(Format((1/(timesig.freqsw))*fmod_rep+timesig.tdelay,"0.0000e+00")))
				End If
			Else
				'Complementary switch signal definition

				If UBound(portnameArray_HS)=0 Then
					modifiedportname_HS=portnameArray_HS(0)+"_"
				Else
					For ii=0 To UBound(portnameArray_HS)
						modifiedportname_HS=modifiedportname_HS+portnameArray_HS(ii)+"_"
					Next
				End If

				If UBound(portnameArray_LS)=0 Then
					modifiedportname_LS=portnameArray_LS(0)+"_"
					Else
					For ii=0 To UBound(portnameArray_LS)
						modifiedportname_LS=modifiedportname_LS+portnameArray_LS(ii)+"_"
					Next
				End If

				'0% spread, use standard digital signal definition
				If ssprop.rate=0 Then
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_NoSpread_HS"
					If modifiedportname_HS<>"" Then	AssignTransienttask_DigitalSignal(transprop,timesig,InitAmp,fmod_rep,modifiedportname_HS,"HS",SignalName)
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_NoSpread_LS"
					If modifiedportname_LS<>"" Then	AssignTransienttask_DigitalSignal(transprop,timesig,InitAmp,fmod_rep,modifiedportname_LS,"LS",SignalName)
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_NoSpread_HS_delayed"
					If modifiedportname_HSDelay<>"" Then AssignTransienttask_DigitalSignal(transprop,timesig,InitAmp,fmod_rep,modifiedportname_HSDelay,"HS",SignalName)
					SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_NoSpread_LS_delayed"
					If modifiedportname_LSDelay<>"" Then AssignTransienttask_DigitalSignal(transprop,timesig,InitAmp,fmod_rep,modifiedportname_LSDelay,"LS",SignalName)
					DlgText("text_tmaxused",cstr(Format((1/(timesig.freqsw))*fmod_rep+timesig.tdelay,"0.0000e+00")))
				Else
					'Creation HighSide Signal without DELAY
					If modifiedportname_HS <>"" Then
						out_dir = GetProjectPath("ModelDS")
						out_dir = out_dir+transprop.taskname+"_Port_"+ modifiedportname_HS
						SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_"+cstr(ssprop.rate)+"%_"+ssprop.profile+"_HS"
						t_end=CreateSpreadSignal(fspread_list,timesig,1,0,0.0)
						For ii=0 To UBound(portnameArray_HS)
							AssignTransientTask (ComplementSig,portnameArray_HS(ii),transprop.taskname,t_end*fmod_rep,0.0,out_dir,SignalName)
						Next
					End If
					'Creation HighSide Signal with DELAY
					If modifiedportname_HSDelay <>"" Then
						out_dir = GetProjectPath("ModelDS")
						out_dir=out_dir+transprop.taskname+"_Port_"+ modifiedportname_HSDelay
						SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_"+cstr(ssprop.rate)+"%_"+ssprop.profile+"_HS_delayed"
						t_end=CreateSpreadSignal(fspread_list,timesig,1,0,0.0)
						For ii=0 To UBound(portnameArray_HSDelay)
							AssignTransientTask (ComplementSig,portnameArray_HSDelay(ii),transprop.taskname,t_end*fmod_rep,timesig.tdelay,out_dir,SignalName)
						Next
					End If
					'Second LowSide Signal without DELAY
					timesig.dutycycle=100-timesig.dutycycle
					If modifiedportname_LS <>"" Then
						out_dir = GetProjectPath("ModelDS")
						out_dir = out_dir+transprop.taskname+"_Port_"+ modifiedportname_LS
						SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_"+cstr(ssprop.rate)+"%_"+ssprop.profile+"_LS"
						t_end=CreateSpreadSignal(fspread_list,timesig,1,1,timesig.tdead)
						For ii=0 To UBound(portnameArray_LS)
							AssignTransientTask (ComplementSig,portnameArray_LS(ii),transprop.taskname,t_end*fmod_rep,0.0,out_dir,SignalName)
						Next
					End If
					'Creation LowSide Signal with DELAY
					If modifiedportname_LSDelay <>"" Then
						out_dir = GetProjectPath("ModelDS")
						out_dir=out_dir+transprop.taskname+"_Port_"+ modifiedportname_LSDelay
						SignalName = cstr(timesig.freqsw*Units.GetFrequencySIToUnit)+Units.GetUnit("Frequency") +"_"+cstr(ssprop.rate)+"%_"+ssprop.profile+"_LS_delayed"
						t_end=CreateSpreadSignal(fspread_list,timesig,1,1,timesig.tdead)
						For ii=0 To UBound(portnameArray_LSDelay)
							AssignTransientTask (ComplementSig,portnameArray_LSDelay(ii),transprop.taskname,t_end*fmod_rep,timesig.tdelay,out_dir,SignalName)
						Next
					End If

					DS.SelectTreeItem("Results\SpreadSignal")
					With Plot1D
						.XRange(0,3/timesig.freqsw)
						.Plot
					End With
					DlgText("text_tmaxused",cstr(Format(t_end*fmod_rep+timesig.tdelay,"0.0000e+00")))
				End If
			End If

		Case "ListBoxSpreadProfile"
			ssprop.profile=ListSpreadProfile(DlgValue("ListBoxSpreadProfile"))

		Case "GroupInitAmp"
			InitAmp=evaluate(DlgValue("GroupInitAmp"))

		Case "GroupExcitationType"
			ComplementSig=evaluate(DlgValue("GroupExcitationType"))
			DialogFunc_EcitationType(ComplementSig)
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "GroupExcitationType", DlgValue("GroupExcitationType")

		Case "CheckBoxUseDelay"
			If DlgValue("CheckBoxUseDelay")="0" Then
				DlgEnable("PushButtonDelayProp",False)
			Else
				DlgEnable("PushButtonDelayProp",True)
			End If
			UseDelay=evaluate(DlgValue("CheckBoxUseDelay"))

		Case "PushButton2"
			Exit All

		Case "PushButtonDelayProp"
			DialogFunc = True
			DialogDelayProperty

		Case "Help"
			DialogFunc = True
			StartHelp "common_preloadedmacro_Spread_Spectrum_Clock_Generation_SSCG"

		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
		BackToDialog:
			DialogFunc = True ' Prevent button press from closing the dialog box

	End Select
End Function

Sub DialogFunc_EcitationType(excitation_type As Integer)
	If excitation_type=1 Then
		DlgEnable("GroupInitAmp",False)
		DlgEnable("diag_init_amp",False)
		DlgEnable("diag_assignport",False)
		DlgEnable("text_assignport",False)
		DlgEnable("Text_dead_time",True)
		DlgEnable("text_deadtime",True)
		DlgEnable("text_assignport_low",True)
		DlgEnable("text_assignport_high",True)
		DlgEnable("diag_assignport_high",True)
		DlgEnable("diag_assignport_low",True)

	Else
		DlgEnable("GroupInitAmp",True)
		DlgEnable("diag_init_amp",True)
		DlgEnable("diag_assignport",True)
		DlgEnable("text_assignport",True)
		DlgEnable("Text_dead_time",False)
		DlgEnable("text_deadtime",False)
		DlgEnable("text_assignport_low",False)
		DlgEnable("text_assignport_high",False)
		DlgEnable("diag_assignport_high",False)
		DlgEnable("diag_assignport_low",False)
	End If
End Sub

Function DialogDelayProperty()
	Begin Dialog UserDialog 370,175,"Delay Properties",.DelayProp ' %GRID:10,7,1,1
		Text 20,14,110,14,"Type of delay:",.Text_DelayType
		OptionGroup .GroupDelayType
			OptionButton 210,14,80,14,"Degree",.OptionButtonDegree
			OptionButton 290,14,60,14,"Time",.OptionButtonTime
		Text 20,42,100,14,"Delay length: ",.TextDelay
		Text 300,42,40,14,"[Deg]",.Text_DelayDimension,1
		TextBox 210,35,90,21,.DelayLength
		OptionGroup .GroupDelayWhichPort
			OptionButton 20,70,180,14,"Delay signal for all port(s)",.OptionButton_DelayAll
			OptionButton 20,98,180,14,"Delay signal port name(s):",.OptionButton_DelaySignalPort
		TextBox 210,91,130,21,.DelayPortLabel
		Text 20,126,330,14,"Note: Use semicolon as separator for multiple ports",.Text1
		OKButton 20,147,90,21
		CancelButton 130,147,90,21
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then 'Do Nothing
	End If

End Function

Rem See DialogFunc help topic for more information.
Private Function DelayProp(DlgItem$, Action%, SuppValue?) As Boolean

	Select Case Action%
	Case 1 ' Dialog box initialization

		DlgValue("GroupDelayType",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayType", "0"))
		DlgText("DelayLength",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "DelayLength", "0"))
		DlgValue("GroupDelayWhichPort",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayWhichPort", "0"))
		DlgText("DelayPortLabel",GetSetting("CST STUDIO SUITE", "SpreadSigGeneration", "DelayPortLabel", ""))

		If DlgValue("GroupDelayType")<>0 Then
			DlgText("Text_DelayDimension","[s]")
		Else
			DlgText("Text_DelayDimension","[Deg]")
		End If

		If DlgValue("GroupDelayWhichPort")=0 Then
			DlgEnable("DelayPortLabel",False)
		Else
			DlgEnable("DelayPortLabel",True)
		End If

	Case 2 ' Value changing or button pressed
		Rem DelayProp = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
		Case "OK"
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayType", DlgValue("GroupDelayType")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "DelayLength", DlgText("DelayLength")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "GroupDelayWhichPort", DlgValue("GroupDelayWhichPort")
			SaveSetting  "CST STUDIO SUITE", "SpreadSigGeneration", "DelayPortLabel", DlgText("DelayPortLabel")

		Case "GroupDelayType"
			If DlgValue("GroupDelayType")<>0 Then
				DlgText("Text_DelayDimension","[s]")
			Else
				DlgText("Text_DelayDimension","[Deg]")
			End If
		Case "GroupDelayWhichPort"
			If DlgValue("GroupDelayWhichPort")=0 Then
				DlgEnable("DelayPortLabel",False)
			Else
				DlgEnable("DelayPortLabel",True)
			End If
		End Select

	Case 3 ' TextBox or ComboBox text changed

	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DelayProp = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub AssignTransientTask (intComplementSig As Integer, sportname As String, staskname As String, intEndTime As Double,dTdelay As Double,sout_dir As String, sPortSigName As String)
	Dim ssTime As String
	ssTime = Replace(Evaluate(intEndTime*Units.GetTimeSIToUnit+dTdelay*Units.GetTimeSIToUnit),",",".")
	Static MyPos As Double, incr As Double

	'RSJ: 05-sept-2023: CST V2024 ASCII import can now use periodic option and delay
	Dim ASCII_import_argument(1) As String
	ASCII_import_argument(0) = sout_dir +""
	ASCII_import_argument(1) = "true"


	ExternalPort.name (sportname)
	If ExternalPort.DoesExist = False Then
		If MyPos=0 Then
			MyPos=50000
		End If
		MyPos=MyPos+incr
		With ExternalPort
			.Reset
			.Name (sportname)
			.Position(MyPos, MyPos)
			.Create
		End With
		incr=100
	End If

	'
	With SimulationTask
		.Name (staskname)
		'Check if task exists
		If .DoesExist Then
			DS.ReportInformationToWindow("Task " +staskname+ " already exists")
		Else
			.Reset
			.Type ("Transient")
			.Name (staskname)
			.Create
			DS.ReportInformationToWindow("Task "+ staskname+" newly created")
		End If
		.Reset
		.Name (staskname)
		.SetPortSignal (sportname,"Import",ASCII_import_argument )
		.SetPortSignalDelay(sportname, IIf(dTdelay=0,"0",Cstr(dTdelay*Units.GetTimeSIToUnit)))
		.SetPortSignalName (sportname,sPortSigName)
		.SetProperty ( "tmax", ssTime ) ' set Tmax for the transient task
		If .GetProperty ( "fmax estimator") = "automatic" Then
			.SetProperty ( "fmax estimator", "transitiontime" )
		Else
			'do nothing
		End If
		.SetPortInnerResistance ( sportname, "0.0" )
		.SetPortSourceType(sportname, "Voltage")
	End With

	DS.ReportInformationToWindow("Excitation signal for Task "+staskname + " with port "+ sportname+" has been sucessfully updated.")

End Sub


Function GenerateFspreadList(temp_freq_sw As Double, temp_ssprop As SpreadProp, temp_nsample_fspread As Long) As Double()
	Dim output_array() As Double
	ReDim output_array(temp_nsample_fspread)
	Dim ii As Long
	ii=0
	Dim Deltafmod As Double
	Deltafmod=temp_ssprop.freqmod/temp_nsample_fspread
	Dim spread_rate As Double
	spread_rate=temp_ssprop.rate

	Select Case temp_ssprop.profile
	Case "Up-Spread"
		spread_rate=spread_rate*0.01
		For ii=0 To temp_nsample_fspread-1
			If ii<=temp_nsample_fspread/2 Then
				output_array(ii)=spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * (ii*Deltafmod) + temp_freq_sw
			Else
				output_array(ii)=-spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * (ii*Deltafmod - 0.5*temp_ssprop.freqmod) + temp_freq_sw*(1+spread_rate)
			End If
		Next
	Case "Center-Spread"
		spread_rate=0.5*spread_rate*0.01
		For ii=0 To temp_nsample_fspread-1
			If ii<=temp_nsample_fspread/2 Then
				output_array(ii)=2*spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * ii*Deltafmod + (1-spread_rate)*temp_freq_sw
			Else
				output_array(ii)=-2*spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * (ii*Deltafmod - 0.5*temp_ssprop.freqmod) + temp_freq_sw*(1+spread_rate)
			End If
		Next
	Case "Down-Spread"
		spread_rate=spread_rate*0.01
		For ii=0 To temp_nsample_fspread-1
			If ii<=temp_nsample_fspread/2 Then
				output_array(ii)=spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * ii*Deltafmod + temp_freq_sw*(1-spread_rate)
			Else
				output_array(ii)=-spread_rate*temp_freq_sw/(0.5*temp_ssprop.freqmod) * (ii*Deltafmod - 0.5*temp_ssprop.freqmod) + temp_freq_sw
			End If
		Next
	End Select
	GenerateFspreadList=output_array
End Function


Function CreateSpreadSignal(temp_fspread_list() As Double, temp_timesig As TimeSignalProp, temp_fmod_rep As Long, vInitAmp As Integer, temp_deadtime As Double) As Double
	Dim jj As Long
	Dim kk As Long, index As Long
	jj=0

	Dim time_out_y() As Double
    Dim time_out_x() As Double
    Dim tstart As Double
    Dim tperiod As Double
	Dim thigh As Double
	Dim tlow As Double

	Dim y1 As Double, y2 As Double,x1 As Double, x2 As Double

	'Waveform for Init amplitude High
	'0-----1           4-----5          8-----9            12----
	'       \         /       \        /       \          /
	'        \       /         \      /         \        /
	'         2-----3           6----7           10----11
    '====Periode 1=====|====Periode 2===|====Periode 3===| etc..

	'Waveform for Init amplitude Low
	'         2-----3           6----7           10---11
	'        /       \         /      \         /       \
	'       /         \       /        \       /         \
	'0-----1           4-----5          8-----9           12----


	'out_dir = 	GetProjectPath("ModelDS")
	out_dir = out_dir + "spread_sig_" + IIf(CBool(vInitAmp),"InitLow","InitHigh") + ".txt"
	Open out_dir For Output As #1

	If temp_deadtime<>0 Then
		If vInitAmp = 0 Then
			index=4
		Else
			index=6
		End If
	Else
		index=5 'One periode of pulse is defined by 5 points
	End If

	'RSJ: 05-sept-2023: V2024 ASCII import can now use periodic option
	'                   temp_fmod_rep set 1 and activate the periodic in ASCII import
	'                   Delay the ASCII is done inside port excitation using signal delay option (supported from CST v2024)
	'                   as consequence, the signal review shows only WITHOUT the delay.
    For jj=1 To temp_fmod_rep
		For kk=0 To UBound(temp_fspread_list)-1
			If kk=0 And jj=1 Then
			Else
				If temp_deadtime <> 0 Then
					index=index+5
				Else
					index=index+4
				End If
			End If

			ReDim Preserve time_out_y(index)
			ReDim Preserve time_out_x(index)

			tperiod=1/temp_fspread_list(kk)

			If temp_deadtime=0 Then
				thigh=tperiod*temp_timesig.dutycycle*0.01-temp_timesig.tfall
				tlow=tperiod*((100-temp_timesig.dutycycle)*0.01)-temp_timesig.trise

				If vInitAmp = 0 Then
					y1=temp_timesig.Vhigh
					y2=temp_timesig.Vlow
					x1=thigh
					x2=tlow
				Else
					y1=temp_timesig.Vlow
					y2=temp_timesig.Vhigh
					x1=tlow
					x2=thigh
				End If

				If kk=0 And jj=1 Then
					time_out_x(index-5)=0
					time_out_y(index-5)=y1
					Print #1, CStr(time_out_x(index-5))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-5))
					time_out_x(index-4)=x1
					time_out_y(index-4)=y1
					Print #1, CStr(time_out_x(index-4))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-4))
					time_out_x(index-3)=time_out_x(index-4)+temp_timesig.tfall
					time_out_y(index-3)=y2
					Print #1, CStr(time_out_x(index-3))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-3))
					time_out_x(index-2)=time_out_x(index-3)+x2
					time_out_y(index-2)=y2
					Print #1, CStr(time_out_x(index-2))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-2))
					time_out_x(index-1)=time_out_x(index-2)+temp_timesig.trise
					time_out_y(index-1)=y1
					Print #1, CStr(time_out_x(index-1))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-1))
				Else

					time_out_x(index-4)=time_out_x(index-5)+x1
					time_out_y(index-4)=y1
					Print #1, CStr(time_out_x(index-4))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-4))
					time_out_x(index-3)=time_out_x(index-4)+temp_timesig.tfall
					time_out_y(index-3)=y2
					Print #1, CStr(time_out_x(index-3))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-3))
					time_out_x(index-2)=time_out_x(index-3)+x2
					time_out_y(index-2)=y2
					Print #1, CStr(time_out_x(index-2))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-2))
					time_out_x(index-1)=time_out_x(index-2)+temp_timesig.trise
					time_out_y(index-1)=y1
					Print #1, CStr(time_out_x(index-1))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-1))
				End If
			Else
				y1=temp_timesig.Vhigh
				y2=temp_timesig.Vlow
				thigh=(tperiod*temp_timesig.dutycycle*0.01)-temp_timesig.tfall-2*temp_timesig.trise-2*temp_deadtime
				tlow=tperiod*((100-temp_timesig.dutycycle)*0.01)+temp_deadtime
				If kk=0 And jj=1 Then
					time_out_x(index-6)=0
					time_out_y(index-6)=y2
					Print #1, CStr(time_out_x(index-6))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-6))
				End If
					time_out_x(index-5)=time_out_x(index-6)+tlow
					time_out_y(index-5)=y2
					Print #1, CStr(time_out_x(index-5))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-5))
					time_out_x(index-4)=time_out_x(index-5)+temp_timesig.trise
					time_out_y(index-4)=y1
					Print #1, CStr(time_out_x(index-4))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-4))
					time_out_x(index-3)=time_out_x(index-4)+thigh
					time_out_y(index-3)=y1
					Print #1, CStr(time_out_x(index-3))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-3))
					time_out_x(index-2)=time_out_x(index-3)+temp_timesig.tfall
					time_out_y(index-2)=y2
					Print #1, CStr(time_out_x(index-2))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-2))
					time_out_x(index-1)=time_out_x(index-6)+tperiod
					time_out_y(index-1)=y2
					Print #1, CStr(time_out_x(index-1))  + vbTab + vbTab + vbTab  +CStr(time_out_y(index-1))
			End If
		Next
	Next

	Close #1

	Dim tempsig1 As Object
	Set tempsig1 = DS.Result1D("")
	tempsig1.initialize(UBound(time_out_x))
	tempsig1.SetArray(time_out_x, "x")
	tempsig1.SetArray(time_out_y, "y")


	If SHOW_SIGNAL_IN_TREE=1 Then
		With tempsig1
			.Title "Spread Signal Review"
			.Xlabel "s"
			'.Ylabel "Volt"
			.Save GetProjectPath("ResultDS") + "spread_debug"
			.AddToTree "SpreadSignal\TimeSignal_"+IIf(CBool(vInitAmp),"InitLow","InitHigh")
		End With
	End If
	CreateSpreadSignal=tempsig1.getx(tempsig1.getN-1)
End Function

Function FindPortNameWithDelay(array_no_delay() As String, array_with_delay() As String) As String()
	Dim ii As Integer
	Dim jj As Integer
	Dim temp_string_w_delay As String
	Dim temp_string_no_delay As String
	Dim bool_flag As Boolean

	temp_string_w_delay=""
	temp_string_no_delay=""
	bool_flag=False

	For ii=0 To UBound(array_no_delay)
		For jj=0 To UBound(array_with_delay)
			If array_no_delay(ii)=array_with_delay(jj) Then
				temp_string_w_delay=temp_string_w_delay+array_no_delay(ii)+";"
				bool_flag=True
				Exit For
			End If
		Next
		If bool_flag=False Then
			temp_string_no_delay=temp_string_no_delay+array_no_delay(ii)+";"
		End If
	Next
	If Right(temp_string_no_delay,1)=";" Then
		temp_string_no_delay=Left(temp_string_no_delay,Len(temp_string_no_delay)-1)
	End If
	If Right(temp_string_w_delay,1)=";" Then
		temp_string_w_delay=Left(temp_string_w_delay,Len(temp_string_w_delay)-1)
	End If

	array_no_delay=Split(temp_string_no_delay,";")
	FindPortNameWithDelay=Split(temp_string_w_delay,";")
End Function

Sub AssignTransienttask_DigitalSignal(temp_transprop As TransTaskProp, temp_timesig As TimeSignalProp, vInitAmp As Integer,temp_fmod_rep As Long,string_compl_port As String,string_HS_LS As String,sPortSigName As String)
	Dim ii As Integer
	Dim incr As Long
	Static MyPos As Double
	Dim portarray() As String
	Dim values() As String
	ReDim values(10) As String
    values(0) = "Duty cycle and frequency"
	values(1) = "Manual"  ' Pattern Type
    values(2) = evaluate(temp_timesig.Vlow)  ' Alow
    values(3) = evaluate(temp_timesig.Vhigh) ' Ahigh
    values(4) = evaluate(temp_timesig.trise*Units.GetTimeSIToUnit) ' Trise
    values(5) = evaluate(temp_timesig.tfall*Units.GetTimeSIToUnit) ' Tfall
    values(6) = evaluate(temp_timesig.dutycycle*0.01) ' Dutycycle
    values(7) = evaluate(temp_timesig.freqsw*Units.GetFrequencySIToUnit) ' Freq
    values(8) = IIf(vInitAmp=0,"1","0") ' Initbit
    values(9) = IIf(vInitAmp=0,"10","01") ' Bitsequence
    values(10) = "true" ' Periodic

	Dim one_period As Double, Tdfall As Double
	one_period = 1/(temp_timesig.freqsw*Units.GetFrequencySIToUnit)

	Dim ssTime As String
	ssTime = Replace(Evaluate(one_period*temp_fmod_rep)+temp_timesig.tdelay*Units.GetTimeSIToUnit,",",".")  'Consider Tdelay in Tmax transient task

	If string_compl_port="" Then 'single port excitation
		ExternalPort.name (temp_transprop.portname)
		If ExternalPort.DoesExist = False Then
			If MyPos=0 Then
				MyPos=50000
			End If
			MyPos=MyPos+incr
			With ExternalPort
				.Reset
				.Name (temp_transprop.portname)
				.Position(MyPos, MyPos)
				.Create
			End With
			incr=100
		End If
		With SimulationTask
			.Name (temp_transprop.taskname)
			'Check if task exists
			If .DoesExist Then
				DS.ReportInformationToWindow("Task " +temp_transprop.taskname+ " already exists")
			Else
				.Reset
				.Type ("Transient")
				.Name (temp_transprop.taskname)
				.SetProperty ( "tmax", ssTime ) ' set Tmax for the transient task
				.Create
				DS.ReportInformationToWindow("Task "+ temp_transprop.taskname+" newly created")
			End If
			.Reset
			.Name (temp_transprop.taskname)
			.SetPortSignal (temp_transprop.portname,"Digital",values)
			.SetPortSignalDelay(temp_transprop.portname, IIf(temp_timesig.tdelay=0,"0",Cstr(temp_timesig.tdelay*Units.GetTimeSIToUnit)))
			.SetPortSignalName (temp_transprop.portname,sPortSigName)
			.SetProperty ( "tmax", ssTime ) ' set Tmax for the transient task
			If .GetProperty ( "fmax estimator") = "automatic" Then
				.SetProperty ( "fmax estimator", "transitiontime" )
			Else
				'do nothing
			End If
			.SetPortInnerResistance( temp_transprop.portname, "0.0" )
			.SetPortSourceType(temp_transprop.portname, "Voltage")
		End With
		DS.ReportInformationToWindow("Excitation signal for Task "+temp_transprop.taskname + " with port "+ temp_transprop.portname+" has been sucessfully updated.")

	Else 'complementary excitation
		If Right(string_compl_port,1)="_" Then
			string_compl_port=Left(string_compl_port,Len(string_compl_port)-1)
		End If
		portarray=Split(string_compl_port,"_")
		Tdfall = one_period*((temp_timesig.dutycycle*0.01)-0.5)

		ReDim values(11) As String
		values(0) = "Generic"
		values(1) = "Manual"  ' Pattern Type
		values(2) = evaluate(temp_timesig.Vlow)  ' Alow
	    values(3) = evaluate(temp_timesig.Vhigh) ' Ahigh
	    values(4) = evaluate(temp_timesig.trise*Units.GetTimeSIToUnit) ' Trise
	    values(5) = evaluate(temp_timesig.tfall*Units.GetTimeSIToUnit) ' Tfall
	    values(8) = evaluate(one_period)  'ttotal
	    values(11) = "true" 'periodic

		Select Case string_HS_LS
			Case "HS" 'High side
			    values(6) = evaluate(-0.5*temp_timesig.trise*Units.GetTimeSIToUnit)       'Tdrise, slightly shift to compensate small delay from initbit 1 --> consistent with online help description
			    values(7) = evaluate(Tdfall-0.5*temp_timesig.tfall*Units.GetTimeSIToUnit) 'Tdfall
				values(9) = "1" ' Initbit
				values(10) = "10" ' Bitseq
		 	Case "LS" 'Low side
				If temp_timesig.tdead=0 Then
					values(6) = evaluate(Tdfall-0.5*temp_timesig.trise*Units.GetTimeSIToUnit)  'slightly shift to compensate small delay from initbit 0 --> consistent with online help description
					values(7) = evaluate(-0.5*temp_timesig.tfall*Units.GetTimeSIToUnit)
				Else
					'Tdelay rise  'shift required as Tdead defined at point where HS initial low
					values(6) = evaluate(Tdfall+0.5*temp_timesig.tfall*Units.GetTimeSIToUnit+temp_timesig.tdead*Units.GetTimeSIToUnit)
					'Tdelay fall  'shift required as Tdead defined at point where HS initial high
				    values(7) = evaluate((-temp_timesig.tdead*Units.GetTimeSIToUnit)-(0.5*temp_timesig.trise*Units.GetTimeSIToUnit)-(temp_timesig.tfall*Units.GetTimeSIToUnit))
				End If
				values(9) = "0" ' Initbit
				values(10) = "01" ' Bitseq
	    End Select
	    For ii=0 To UBound(portarray)
			ExternalPort.name (portarray(ii))
			If ExternalPort.DoesExist = False Then
				If MyPos=0 Then
					MyPos=50000
				End If
				MyPos=MyPos+incr
				With ExternalPort
					.Reset
					.Name (portarray(ii))
					.Position(MyPos, MyPos)
					.Create
				End With
				incr=100
			End If
			With SimulationTask
				.Name (temp_transprop.taskname)
				'Check if task exists
				If .DoesExist Then
					DS.ReportInformationToWindow("Task " +temp_transprop.taskname+ " already exists")
				Else
					.Reset
					.Type ("Transient")
					.Name (temp_transprop.taskname)
					.Create
					DS.ReportInformationToWindow("Task "+ temp_transprop.taskname+" newly created")
				End If
				.Reset
				.Name (temp_transprop.taskname)
				.SetPortSignal (portarray(ii),"Digital",values)
				.SetPortSignalDelay(portarray(ii), IIf(InStr(sPortSigName,"delayed")=0,"0",Cstr(temp_timesig.tdelay*Units.GetTimeSIToUnit)))
				.SetPortSignalName (portarray(ii),sPortSigName )
				.SetProperty ( "tmax", ssTime) ' set Tmax for the transient task
				If .GetProperty ( "fmax estimator") = "automatic" Then
					.SetProperty ( "fmax estimator", "transitiontime" )
				Else
					'do nothing
				End If
				.SetPortInnerResistance( portarray(ii), "0.0" )
				.SetPortSourceType(portarray(ii), "Voltage")
			End With
			DS.ReportInformationToWindow("Excitation signal for Task "+temp_transprop.taskname + " with port "+ portarray(ii)+" has been sucessfully updated.")
		Next
	End If
End Sub



