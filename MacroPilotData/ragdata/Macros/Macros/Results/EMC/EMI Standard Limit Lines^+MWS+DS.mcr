'#Language "WWB-COM"

'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2022-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 18-Feb-2025 rsj: new command Plot1D.SetdBUnit ("dBu") applied to activate dBu log y-axis
' 03-Jan-2025 rsj: Add CISPR 25 CE I-method.
'                  CISPR 32 CE method correction and adding several classes into it.
'                  FCCb 15 CE limit fmax correction to 30MHz.
'                  Renaming the y-axis label according to limit component (e.g. field, voltage or current).
'                  Add function PlotLogAxis to set log x-axis and dB magnitude in y-axis.
' 06-Mar-2024 rsj: Add limit unit selection for user defined standard limit import.																			   
' 06-Sep-2023 rsj: Example format section now contains hint for Title keyword.
'                  Result folder will be named "untitled" if title keyword is missing.
' 28-Apr-2023 rsj: Set index to zero to start a new xrange in sub ReadAndPlotUserLimitFile after line input="[----end----]"
' 09-May-2022 ssp: Renamed macro
' 01-Apr-2022 rsj: First version
'-----------------------------------------------------------------------------------------------------------------------------

Option Explicit

Const ListArrayEMCStandard = Array ("UNSELECTED","CISPR25 CE(V-method)","CISPR25 CE(I-method)","CISPR25 RE-ALSE","CISPR25 RE-TEM","CISPR25 RE-Stripline","FCC15b - CE","FCC15b - RE", _
									"CISPR32-EN5032 - CE", "CISPR32-EN5032 - RE", "CISPR11-EN5011 - CE(Group1)" , "CISPR11-EN5011 - RE(Group1)", "MIL-STD-461G CE102", _
								    "MIL-STD-461G RE102")
Const RGBcolpeak ="color=255;0;0 linetype=Solid nocheck"
Const RGBcolqp   ="color=0;0;255 linetype=Solid nocheck"
Const RGBcolavg  ="color=0;128;0 linetype=Solid nocheck"

Const dFreqUnit = Units.GetFrequencySIToUnit

Const CISPR25_CEPeakArray_Class5 = Array (70,54,53,38,34,44,38)
Const CISPR25_CECurrentPeakArray_Class5 = Array (50,26,19,4,0,10,4)
Const CISPR25_ALSEPeakArray_Class5 = Array (46,40,38,28,32,26,41,45,28,34,40,35,38,32,44)
Const CISPR25_TEMPeakArray_Class5 = Array (26,20,26,16,10,20)
Const CISPR25_StriplinePeakArray_Class5 = Array (47,41,32,22,16,22,26,41,32,26,20,32)
Const CISPR32_CE_ClassA_QP = Array (79,73)
Const CISPR32_CE_ClassB_QP = Array (66,56,60)
Const CISPR32_CE_ClassA_QP_V_telecom = Array (97,87)
Const CISPR32_CE_ClassB_QP_V_telecom = Array (84,74)
Const CISPR32_CE_ClassA_QP_I_telecom = Array (53,43)
Const CISPR32_RE_ClassA_3mQP = Array (50.5,57.5)
Const CISPR32_RE_ClassA_3mPeak = Array (76,80)
Const CISPR11_CE_ClassA_MainPorts_QP_greater75kVA = Array(130,125,115)
Const CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA = Array(100,86,90,73)
Const CISPR11_CE_ClassA_MainPorts_QP_less20kVA = Array(79,73)
Const CISPR11_CE_ClassB_MainPorts_QP = Array(66,56,60)
Const CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim = Array(132,122,105)
Const CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim = Array(116,106,89)
Const CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim = Array(97,89)
Const CISPR11_CE_ClassB_DCPorts_QP_Vlim = Array(84,74)
Const CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim = Array(88,78,61)
Const CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim = Array(72,62,45)
Const CISPR11_RE_ClassA_QP3m_greater20kVA = Array(60)
Const CISPR11_RE_ClassA_QP3m_less20kVA = Array(50,57)
Const MIL_STD_461_CE102 = Array (94,60)
Const MIL_STD_461_RE102_ShipApp = Array (90,56,102)
Const MIL_STD_461_RE102_SubmarAppInternal = Array (88,50,95)
Const MIL_STD_461_RE102_SubmarAppExternal = Array (60,24,69)
Const MIL_STD_461_RE102_AirSpaceApp = Array (60,24,69)
Const MIL_STD_461_RE102_GroundApp = Array (44,89)
Const LimitUnit = Array ("A","V","A/m","V/m")
Const LimitLabel = Array ("Current","Voltage","H-field","E-field")					 


Dim WhereAmI As String
Dim EMCLimitClassApp () As Variant
Dim intEMCStandard As Integer
Dim strEMCClassApp As String, out_tree As String, ForTreeName As String, UserLimitFilePath As String
Dim BoolUserStandard As Boolean
Dim ListArrayHistory() As String

Sub Main

	ReDim ListArrayHistory(1)
	BoolUserStandard=False
	EMCLimitClassApp = Array ("Class 1", "Class 2", "Class 3", "Class 4", "Class 5")
	Begin Dialog UserDialog 710,217,"EMI Standard Limit Lines",.DialogFuncEMCStand ' %GRID:10,7,1,1
		GroupBox 20,7,330,63,"EMC Standard",.GroupBox3
		GroupBox 360,7,330,63,"Limit Class / Application",.GroupBox4
		DropListBox 30,28,310,35,ListArrayEMCStandard(),.DropListBoxEMCStandard
		DropListBox 370,28,310,28,EMCLimitClassApp(),.DropListBoxClassApp
		CancelButton 240,182,90,21
		PushButton 20,182,90,21,"Ok",.PushButtonApply
		PushButton 130,182,90,21,"Close",.PushButtonClose
		GroupBox 20,77,670,98,"",.GroupBox1
		TextBox 30,119,650,21,.TextBoxLimitLoc
		CheckBox 30,77,170,14,"User defined standard",.CheckBoxUserDefined
		PushButton 590,91,90,21,"Load file",.PushButtonLoad
		PushButton 580,182,110,21,"Example format",.PushButtonExample
		Text 30,98,120,14,"Limit file location:",.TextLimitfileLoc
		Text 30,154,70,14,"Limit unit:",.Text_user_limit_unit
		OptionGroup .UserInputLimitUnit
			OptionButton 100,154,50,14,"A",.OptionButton_unit_A
			OptionButton 160,154,50,14,"V",.OptionButton_unit_V
			OptionButton 220,154,60,14,"A/m",.OptionButton_unit_Am
			OptionButton 280,154,60,14,"V/m",.OptionButton_unit_Vm
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFuncEMCStand(DlgItem$, Action%, SuppValue?) As Boolean

	Dim main_emc_limit_peak As Object, main_emc_limit_qp As Object, main_emc_limit_avg As Object
	Static TextForHistory As String

	DlgVisible("Cancel",False)  'For proper function the red "X" button. Dialog will be closed.

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgValue("DropListBoxEMCStandard","1")
		DlgValue("DropListBoxClassApp","0")
		DlgText("TextBoxLimitLoc",GetSetting("CST STUDIO SUITE", "EMI Standard Limit", "TextBoxLimitLoc", ""))
		UserLimitFilePath=(DlgText("TextBoxLimitLoc"))
		DlgValue("CheckBoxUserDefined",False)
		DlgEnable("PushButtonLoad",False)
		DlgEnable("TextBoxLimitLoc",False)
		DlgEnable("PushButtonExample",False)
		DlgEnable("UserInputLimitUnit",False)
		DlgValue("UserInputLimitUnit",3)				   
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "PushButtonApply"
			DlgEnable("PushButtonApply",False)
			DlgEnable("PushButtonClose",False)
			SaveSetting  "CST STUDIO SUITE", "EMI Standard Limit", "TextBoxLimitLoc", DlgText("TextBoxLimitLoc")

			WhereAmI     = Left(GetApplicationName,2)

			If BoolUserStandard=False Then

				intEMCStandard = DlgValue("DropListBoxEMCStandard")
				strEMCClassApp = EMCLimitClassApp(DlgValue("DropListBoxClassApp"))
				ForTreeName  = ListArrayEMCStandard(intEMCStandard)+" "+strEMCClassApp+"\"

				If WhereAmI ="DS" Then
					Set main_emc_limit_peak = DS.Result1DComplex("")
					Set main_emc_limit_qp = DS.Result1DComplex("")
					Set main_emc_limit_avg = DS.Result1DComplex("")
					out_tree = "Results\"+ForTreeName
				Else
					Set main_emc_limit_peak = Result1DComplex("")
					Set main_emc_limit_qp = Result1DComplex("")
					Set main_emc_limit_avg = Result1DComplex("")
					out_tree = "1D Results\EMCStandard_Limit\"+ForTreeName
				End If

				If intEMCStandard>0 And intEMCStandard<6 Then
					'Plot the CISPR25
					PlotEMILimitCISPR25
					ElseIf intEMCStandard=6 Then
					'Plot the FCC15b - CE
					PlotEMILimitFCCbCE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=7 Then
					'Plot the FCC15b - RE
					PlotEMILimitFCCbRE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=8 Then
					'Plot CISPR32/EN5032 - CE
					PlotEMILimitCISPR32_CE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=9 Then
					'Plot CISPR32/EN5032 - RE
					PlotEMILimitCISPR32_RE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=10 Then
					'Plot CISPR11-EN5011 - CE(Group1)
					PlotEMILimitCISPR11_CE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=11 Then
					'Plot CISPR11-EN5011 - RE(Group1)
					PlotEMILimitCISPR11_RE(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				ElseIf intEMCStandard=12 Then
					'Plot MIL-STD-461G CE102
					PlotEMILimitMILSTD_CE102(main_emc_limit_peak)
				ElseIf intEMCStandard=13 Then
					'Plot MIL-STD-461G RE102
					PlotEMILimitMILSTD_RE102(main_emc_limit_peak, main_emc_limit_qp, main_emc_limit_avg)
				End If
				If WhereAmI ="DS" Then
					DS.SelectTreeItem out_tree
				Else
					SelectTreeItem out_tree
				End If
			Else
				If UserLimitFilePath="" Then
					MsgBox("No input file has been selected.",vbCritical,"EMI Standard Limit")
					DialogFuncEMCStand = True
				Else
					ReadAndPlotUserLimitFile(Evaluate(DlgValue("UserInputLimitUnit")))
				End If

			End If

			Plot1D.XLogarithmic(True)
			Plot1D.plot

			DlgEnable("PushButtonApply",True)
			DlgEnable("PushButtonClose",True)
			DialogFuncEMCStand = False

		Case "DropListBoxEMCStandard"
			If DlgValue("DropListBoxEMCStandard")>0 And DlgValue("DropListBoxEMCStandard")<6 Then
				EMCLimitClassApp = Array ("Class 1", "Class 2", "Class 3", "Class 4", "Class 5")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 6 Then 'FCC15b - CE
				EMCLimitClassApp = Array ("Class A Digital Device")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 7 Or DlgValue("DropListBoxEMCStandard")= 9 Then 'FCC15b - RE or CISPR32-EN5032 - RE
				EMCLimitClassApp = Array ("Class A 3m", "Class A 10m", "Class B 3m", "Class B 10m")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 8 Then 'CISPR32-EN5032 - CE
				EMCLimitClassApp = Array ("Class A Voltage limit (Main Ports)", "Class A Voltage limit (Telecom-LAN Ports)", "Class A Current limit (Telecom-LAN Ports)", "Class B Voltage limit (Main Ports)", _
				                          "Class B Voltage limit (Telecom-LAN Ports)", "Class B Current limit (Telecom-LAN Ports)")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 10 Then 'CISPR11-EN5011 - CE(Group1)
				EMCLimitClassApp = Array ("Class A Main Ports (>75kVA)", "Class A Main Ports (>20kVA & <=75kVA)", "Class A Main Ports (<=20kVA)","Class B Main Ports", _
										  "Class A DCPorts V-Limit (>75kVA)", "Class A DCPorts V-Limit (>20kVA & <=75kVA)", "Class A DCPorts V-Limit (<=20kVA)", "Class B DCPorts V-Limit", _
				                          "Class A DCPorts I-Limit (>75kVA)", "Class A DCPorts I-Limit (>20kVA & <=75kVA)")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 11 Then 'CISPR11-EN5011 - RE(Group1)
				EMCLimitClassApp = Array ("Class A 3m (>20kVA)", "Class A 10m (>20kVA)","Class A 3m (<=20kVA)", "Class A 10m (<=20kVA)", "Class B 3m", "Class B 10m")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 12 Then 'MIL-STD-461G CE102
				EMCLimitClassApp = Array ("SourceVoltage 28V","SourceVoltage 115V","SourceVoltage 220V","SourceVoltage 270V","SourceVoltage 440V")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			ElseIf DlgValue("DropListBoxEMCStandard")= 13 Then 'MIL-STD-461G RE102
				EMCLimitClassApp = Array ("Surface ship","Submarine","Aircraft and space system","Ground")
				DlgListBoxArray("DropListBoxClassApp", EMCLimitClassApp())
				DlgValue("DropListBoxClassApp","0")
				DlgEnable("DropListBoxClassApp",True)
				DialogFuncEMCStand = True
			End If
		Case "PushButtonLoad"
			DialogFuncEMCStand=True
			InitDialogLoadSetting()
		Case "PushButtonExample"
			DialogExample
			DialogFuncEMCStand = True
		Case "CheckBoxUserDefined"
			BoolUserStandard=DlgValue("CheckBoxUserDefined")
			If BoolUserStandard=True Then
				DlgEnable("DropListBoxEMCStandard",False)
				DlgEnable("DropListBoxClassApp",False)
				DlgEnable("TextBoxLimitLoc",True)
				DlgEnable("PushButtonLoad",True)
				DlgEnable("PushButtonExample",True)
				DlgEnable("UserInputLimitUnit",True)						
			Else
				DlgEnable("DropListBoxEMCStandard",True)
				DlgEnable("DropListBoxClassApp",True)
				DlgEnable("TextBoxLimitLoc",False)
				DlgEnable("PushButtonLoad",False)
				DlgEnable("PushButtonExample",False)
				DlgEnable("UserInputLimitUnit",False)						 
			End If

		End Select

		Rem DialogFuncEMCStand = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed

	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFuncEMCStand = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function InitDialogLoadSetting() As Integer
	Dim tempfilepath As String

	tempfilepath = GetFilePath("*.*", "All files|*.*|Text files|*.txt", "", "Please select data file", 0+4)

	If tempfilepath = "" Then
		'Exit All ' User pressed cancel in file selection dialog
	Else
		UserLimitFilePath=tempfilepath
		DlgText("TextBoxLimitLoc",UserLimitFilePath)
	End If

End Function

Sub ReadAndPlotUserLimitFile (UserLimitUnit As Integer)
	Dim cst_line_input As String
	Dim bandname As String, detectorname As String
	Dim stemp() As String
	Dim xrange() As Double, yrange() As Double
	Dim counter As Long, ii As Long
	Dim dFreqPeakUnitScale As Double, dFreqQPUnitScale As Double, dFreqAVGUnitScale As Double
	Dim YAxisUnit As String, XAxisUnit As String, sTitle As String

	counter=1

	Open UserLimitFilePath For Input As #2

	sTitle="Untitled"

	While Not EOF(2)
		Line Input #2,cst_line_input

		If Left(cst_line_input,1)="*" Or Left(cst_line_input,1)="" Then
			'This is a comment. Do nothing
			'GoTo SKIP
		End If

		If Left(cst_line_input,6)="[Title" Or Left(cst_line_input,6)="[title" Or Left(cst_line_input,6)="[TITLE" Then
			sTitle=Mid(cst_line_input,8,Len(cst_line_input)-8)
		End If

		If Left(cst_line_input,6)="[Peak]" Or Left(cst_line_input,6)="[QP]" Or Left(cst_line_input,6)="[AVG]" Then
			detectorname=Mid(Left(cst_line_input,6),2,Len(Left(cst_line_input,6))-2)
			Line Input #2,cst_line_input

			'Now read the frequency and amplitude unit
			stemp=Split(cst_line_input,";")
			XAxisUnit=Right(stemp(0),Len(stemp(0))-1)
			dFreqPeakUnitScale=FreqUniScale(XAxisUnit)
			If dFreqPeakUnitScale=-1 Then
				'Warning Error Freq.unit
				MsgBox("Incorrect frequency unit definition",vbCritical,"EMI Standard Limit")
				GoTo ErrorSetting
			End If
			YAxisUnit=Left(stemp(1),Len(stemp(1))-1)

			'Now read the frequency content for each band
			bandname=""
			Line Input #2,cst_line_input
			Do
				stemp=Split(cst_line_input,";")
				If stemp(0)<>bandname And bandname<>"" Then
					'Now Plot
					PlotUserLimit (detectorname,bandname,xrange,yrange,sTitle,UserLimitUnit,counter)
					counter=counter+1
					ii=0
					ReDim xrange(ii+1)
					ReDim yrange(ii+1)
					bandname=stemp(0)
					xrange(ii)=cdbl(stemp(1))*dFreqPeakUnitScale
					xrange(ii+1)=cdbl(stemp(2))*dFreqPeakUnitScale
					Select Case YAxisUnit
					Case "linear", "LINEAR", "Linear"
						yrange(ii)=cdbl(stemp(3))
						yrange(ii+1)=cdbl(stemp(4))
					Case "dBu", "DBU", "dBU", "dbu"
						yrange(ii)=10^((cdbl(stemp(3))-120)/20)
						yrange(ii+1)=10^((cdbl(stemp(4))-120)/20)
					Case "dB", "DB", "db"
						yrange(ii)=10^(cdbl(stemp(3))/20)
						yrange(ii+1)=10^(cdbl(stemp(4))/20)
					End Select
				Else
					ReDim Preserve xrange(ii+1)
					ReDim Preserve yrange(ii+1)
					bandname=stemp(0)
					xrange(ii)=cdbl(stemp(1))*dFreqPeakUnitScale
					xrange(ii+1)=cdbl(stemp(2))*dFreqPeakUnitScale
					Select Case YAxisUnit
					Case "linear", "LINEAR", "Linear"
						yrange(ii)=cdbl(stemp(3))
						yrange(ii+1)=cdbl(stemp(4))
					Case "dBu", "DBU", "dBU", "dbu"
						yrange(ii)=10^((cdbl(stemp(3))-120)/20)
						yrange(ii+1)=10^((cdbl(stemp(4))-120)/20)
					Case "dB", "DB", "db"
						yrange(ii)=10^(cdbl(stemp(3))/20)
						yrange(ii+1)=10^(cdbl(stemp(4))/20)
					End Select
					ii=ii+2
				End If
				Line Input #2,cst_line_input
				If cst_line_input="[----end----]" Then
					PlotUserLimit (detectorname,bandname,xrange,yrange,sTitle,UserLimitUnit,counter)
					'2023-04-28 rsj: set index ii to zero in order to start new xrange
					ii=0
				End If
			Loop Until cst_line_input="[----end----]"

		End If

SKIP:
	Wend
ErrorSetting:
	Close #2

	If WhereAmI ="DS" Then
		SelectTreeItem "Results\"+sTitle+"\"
	Else
		SelectTreeItem "1D Results\EMC User Standard Limit\"+sTitle+"\"
	End If

End Sub

Function FreqUniScale (sinput As String) As Double

	Select Case sinput
	Case "Hz", "HZ", "hz"
		FreqUniScale=dFreqUnit
	Case "kHz", "khz", "KHZ"
		FreqUniScale=1e3*dFreqUnit
	Case "MHz", "mhz", "MHZ"
		FreqUniScale=1e6*dFreqUnit
	Case "GHz", "ghz", "GHZ"
		FreqUniScale=1e9*dFreqUnit
	Case "THz", "thz", "THZ"
		FreqUniScale=1e12*dFreqUnit
	Case Else
		FreqUniScale=-1
	End Select

End Function


Sub PlotUserLimit (sdetectorname As String,sbandname As String,dxrange() As Double,dyrerange() As Double,sTitleName As String,iUserLimitUnit As Integer,lcounter As Long)
	Dim UserLimitPlot() As Object
	ReDim Preserve UserLimitPlot(lcounter)
	Dim out_treename As String
	Dim dyimrange() As Double
	ReDim dyimrange(UBound(dyrerange))


	If WhereAmI ="DS" Then
		Set UserLimitPlot(lcounter-1) = DS.Result1DComplex("")
		out_treename="Results\"+sTitleName+"\"
	Else
		Set UserLimitPlot(lcounter-1) = Result1DComplex("")
		out_treename="1D Results\EMC User Standard Limit\"+sTitleName+"\"
	End If

	UserLimitPlot(lcounter-1).initialize(UBound(dyrerange)+1)
	With UserLimitPlot(lcounter-1)
		.SetArray(dxrange,"x")
		.SetArray(dyrerange,"yre")
		.SetArray(dyimrange,"yim")
		.SetYLabelAndUnit (LimitLabel(iUserLimitUnit),LimitUnit(iUserLimitUnit))
		.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
		.Title (sTitleName)
		.Save GetProjectPath("Result") + sTitleName+"_"+sdetectorname+ "-Band_"+ sbandname
		.AddToTree out_treename + sdetectorname+ "-Band_"+ sbandname
	End With

	If WhereAmI ="DS" Then
		If sdetectorname = "Peak" Then
			DS.SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolpeak)
		ElseIf sdetectorname = "QP" Then
			DS.SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolqp)
		ElseIf sdetectorname = "AVG" Then
			DS.SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolavg)
		End If
		PlotLogAxis (out_treename,True)
	Else
		If sdetectorname = "Peak" Then
			SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolpeak)
		ElseIf sdetectorname = "QP" Then
			SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolqp)
		ElseIf sdetectorname = "AVG" Then
			SetPlotStyleForTreeItem(out_treename +sdetectorname+ "-Band_"+ sbandname,RGBcolavg)
		End If
		PlotLogAxis (out_treename,False)
	End If

End Sub

Sub PlotLogAxis (treename As String, IsDES As Boolean)
	IIf (IsDES,DS.Selecttreeitem(treename),Selecttreeitem(treename))
	Plot1D.XLogarithmic(True)
	Plot1D.SetdBUnit ("dBu")
	Plot1D.PlotView("magnitudedb")
End Sub


Sub PlotEMILimitCISPR25 ()
	Dim temp_val As Double, ii As Long, qpfactor As Double, avgfactor As Double, limitclass As Integer
	Dim emc_limit_peak (30) As Object, emc_limit_qp (30)As Object, emc_limit_avg (30) As Object
	Dim LimitTitle As String

	qpfactor=13
	avgfactor=20
	limitclass = 5-Cint(Right(strEMCClassApp,1))


	If (WhereAmI = "DS") Then
		For ii=0 To UBound(emc_limit_peak)-1
			Set emc_limit_peak(ii) = DS.Result1DComplex("")
			Set emc_limit_qp(ii) = DS.Result1DComplex("")
			Set emc_limit_avg(ii) = DS.Result1DComplex("")
		Next
	Else
		For ii=0 To UBound(emc_limit_peak)-1
			Set emc_limit_peak(ii) = Result1DComplex("")
			Set emc_limit_qp(ii) = Result1DComplex("")
			Set emc_limit_avg(ii) = Result1DComplex("")
		Next
	End If

	Select Case intEMCStandard
	Case "1" 'CISPR 25 CE V-Method
		LimitTitle="Conducted Emission (V-Method)"
		'LW
		temp_val=limitclass*10
		PlotLimit (emc_limit_peak(0), "Peak", "-Band_LW", LimitTitle , CISPR25_CEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(0), "QP", "-Band_LW", LimitTitle,CISPR25_CEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=limitclass*10-avgfactor
		PlotLimit (emc_limit_avg(0), "AVG", "-Band_LW", LimitTitle,CISPR25_CEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)

		'MW
		temp_val=limitclass*8
		PlotLimit (emc_limit_peak(1), "Peak", "-Band_MW",LimitTitle, CISPR25_CEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(1), "QP", "-Band_MW",LimitTitle,CISPR25_CEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=limitclass*8-avgfactor
		PlotLimit (emc_limit_avg(1), "AVG", "-Band_MW", LimitTitle,CISPR25_CEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)

		'SW
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(2), "Peak", "-Band_SW",LimitTitle, CISPR25_CEPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(2), "QP", "-Band_SW",LimitTitle, CISPR25_CEPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(2), "AVG", "-Band_SW",LimitTitle, CISPR25_CEPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)

		'CB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(3), "Peak", "-Band_CB", LimitTitle, CISPR25_CEPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(3), "QP", "-Band_CB",LimitTitle,CISPR25_CEPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(3), "AVG", "-Band_CB", LimitTitle,CISPR25_CEPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)

		'FM
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(4), "Peak", "-Band_FM",LimitTitle, CISPR25_CEPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(4), "QP", "-Band_FM",LimitTitle, CISPR25_CEPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(4), "AVG", "-Band_FM",LimitTitle, CISPR25_CEPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)

		'TVBandI
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(5), "Peak", "-Band_TVI",LimitTitle,CISPR25_CEPeakArray_Class5(4), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(5), "AVG", "-Band_TVI",LimitTitle, CISPR25_CEPeakArray_Class5(4), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)

		'VHF1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(6), "Peak", "-Band_VHF1",LimitTitle,CISPR25_CEPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(5), "QP", "-Band_VHF1",LimitTitle, CISPR25_CEPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(6), "AVG", "-Band_VHF1", LimitTitle,CISPR25_CEPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)

		'VHF2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(7), "Peak", "-Band_VHF2",LimitTitle,CISPR25_CEPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(7), "QP", "-Band_VHF2",LimitTitle, CISPR25_CEPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(7), "AVG", "-Band_VHF2",LimitTitle,CISPR25_CEPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)

	Case "2" 'CISPR 25 CE I-Method
		LimitTitle="Conducted Emission (I-Method)"
		'LW
		temp_val=limitclass*10
		PlotLimit (emc_limit_peak(0), "Peak", "-Band_LW", LimitTitle, CISPR25_CECurrentPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(0), "QP", "-Band_LW", LimitTitle,CISPR25_CECurrentPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=limitclass*10-avgfactor
		PlotLimit (emc_limit_avg(0), "AVG", "-Band_LW", LimitTitle,CISPR25_CECurrentPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)

		'MW
		temp_val=limitclass*8
		PlotLimit (emc_limit_peak(1), "Peak", "-Band_MW",LimitTitle, CISPR25_CECurrentPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(1), "QP", "-Band_MW",LimitTitle,CISPR25_CECurrentPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=limitclass*8-avgfactor
		PlotLimit (emc_limit_avg(1), "AVG", "-Band_MW", LimitTitle,CISPR25_CECurrentPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)

		'SW
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(2), "Peak", "-Band_SW",LimitTitle, CISPR25_CECurrentPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(2), "QP", "-Band_SW",LimitTitle, CISPR25_CECurrentPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(2), "AVG", "-Band_SW",LimitTitle, CISPR25_CECurrentPeakArray_Class5(2), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)

		'CB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(3), "Peak", "-Band_CB", LimitTitle, CISPR25_CECurrentPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(3), "QP", "-Band_CB",LimitTitle,CISPR25_CECurrentPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(3), "AVG", "-Band_CB", LimitTitle,CISPR25_CECurrentPeakArray_Class5(5), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)

		'FM
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(4), "Peak", "-Band_FM",LimitTitle, CISPR25_CECurrentPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(4), "QP", "-Band_FM",LimitTitle, CISPR25_CECurrentPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(4), "AVG", "-Band_FM",LimitTitle, CISPR25_CECurrentPeakArray_Class5(3), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)

		'TVBandI
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(5), "Peak", "-Band_TVI",LimitTitle,CISPR25_CECurrentPeakArray_Class5(4), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(5), "AVG", "-Band_TVI",LimitTitle, CISPR25_CECurrentPeakArray_Class5(4), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)

		'VHF1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(6), "Peak", "-Band_VHF1",LimitTitle,CISPR25_CECurrentPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(5), "QP", "-Band_VHF1",LimitTitle, CISPR25_CECurrentPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(6), "AVG", "-Band_VHF1", LimitTitle,CISPR25_CECurrentPeakArray_Class5(5), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)

		'VHF2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(7), "Peak", "-Band_VHF2",LimitTitle,CISPR25_CECurrentPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(7), "QP", "-Band_VHF2",LimitTitle, CISPR25_CECurrentPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(7), "AVG", "-Band_VHF2",LimitTitle,CISPR25_CECurrentPeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)

	Case "3" '"CISPR25: RE-ALSE"
		LimitTitle="RE-ALSE"
		'LW
		temp_val=limitclass*10
		PlotLimit (emc_limit_peak(0), "Peak", "-Band_LW",LimitTitle, CISPR25_ALSEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(0), "QP", "-Band_LW", LimitTitle,CISPR25_ALSEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=limitclass*10-avgfactor
		PlotLimit (emc_limit_avg(0), "AVG", "-Band_LW",LimitTitle, CISPR25_ALSEPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)

		'MW
		temp_val=limitclass*8
		PlotLimit (emc_limit_peak(1), "Peak", "-Band_MW", LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(1), "QP", "-Band_MW",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=limitclass*8-avgfactor
		PlotLimit (emc_limit_avg(1), "AVG", "-Band_MW",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)

		'SW
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(2), "Peak", "-Band_SW",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(2), "QP", "-Band_SW",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(2), "AVG", "-Band_SW", LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)

		'CB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(3), "Peak", "-Band_CB", LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(3), "QP", "-Band_CB", LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(3), "AVG", "-Band_CB",LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)

		'FM
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(4), "Peak", "-Band_FM",LimitTitle, CISPR25_ALSEPeakArray_Class5(2), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(4), "QP", "-Band_FM",LimitTitle, CISPR25_ALSEPeakArray_Class5(2), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(4), "AVG", "-Band_FM",LimitTitle, CISPR25_ALSEPeakArray_Class5(2), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)

		'TVBandI
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(5), "Peak", "-Band_TVI",LimitTitle, CISPR25_ALSEPeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(5), "AVG", "-Band_TVI", LimitTitle,CISPR25_ALSEPeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)

		'VHF1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(6), "Peak", "-Band_VHF1",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(5), "QP", "-Band_VHF1",LimitTitle,CISPR25_ALSEPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(6), "AVG", "-Band_VHF1",LimitTitle, CISPR25_ALSEPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)

		'VHF2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(7), "Peak", "-Band_VHF2",LimitTitle, CISPR25_ALSEPeakArray_Class5(11), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(6), "QP", "-Band_VHF2", LimitTitle,CISPR25_ALSEPeakArray_Class5(11), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(7), "AVG", "-Band_VHF2",LimitTitle, CISPR25_ALSEPeakArray_Class5(11), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)


		'VHF3
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(8), "Peak", "-Band_VHF3", LimitTitle,CISPR25_ALSEPeakArray_Class5(11), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(7), "QP", "-Band_VHF3", LimitTitle,CISPR25_ALSEPeakArray_Class5(11), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(8), "AVG", "-Band_VHF3",LimitTitle, CISPR25_ALSEPeakArray_Class5(11), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)

		'TVBandIII
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(9), "Peak", "-Band_TVBandIII",LimitTitle, CISPR25_ALSEPeakArray_Class5(4), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(9), "AVG", "-Band_TVBandIII",LimitTitle, CISPR25_ALSEPeakArray_Class5(4), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)

		'DABIII
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(10), "Peak", "-Band_DABIII", LimitTitle,CISPR25_ALSEPeakArray_Class5(5), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(10), "AVG", "-Band_DABIII", LimitTitle,CISPR25_ALSEPeakArray_Class5(5), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)

		'TVBandIV
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(11), "Peak", "-Band_TVBandIV",LimitTitle, CISPR25_ALSEPeakArray_Class5(6), temp_val, 468e6*dFreqUnit, 944e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(11), "AVG", "-Band_TVBandIV", LimitTitle,CISPR25_ALSEPeakArray_Class5(6), temp_val, 468e6*dFreqUnit, 944e6*dFreqUnit, out_tree)

		'DTTV
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(12), "Peak", "-Band_DTTV",LimitTitle, CISPR25_ALSEPeakArray_Class5(6), temp_val, 470e6*dFreqUnit, 770e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(12), "AVG", "-Band_DTTV", LimitTitle,CISPR25_ALSEPeakArray_Class5(6), temp_val, 470e6*dFreqUnit, 770e6*dFreqUnit, out_tree)

		'UHF(1)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(13), "Peak", "-Band_UHF(1)", LimitTitle,CISPR25_ALSEPeakArray_Class5(2), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(8), "QP", "-Band_UHF(1)",LimitTitle, CISPR25_ALSEPeakArray_Class5(2), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(13), "AVG", "-Band_UHF(1)",LimitTitle, CISPR25_ALSEPeakArray_Class5(2), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)


		'UHF(2)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(14), "Peak", "-Band_UHF(2)",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(9), "QP", "-Band_UHF(2)",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(14), "AVG", "-Band_UHF(2)", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)


		'RKE_1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(15), "Peak", "-Band_RKE(1)", LimitTitle,CISPR25_ALSEPeakArray_Class5(4), temp_val, 300e6*dFreqUnit, 330e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+6
		PlotLimit (emc_limit_avg(15), "AVG", "-Band_RKE(1)",LimitTitle, CISPR25_ALSEPeakArray_Class5(4), temp_val, 300e6*dFreqUnit, 330e6*dFreqUnit, out_tree)

		'RKE_2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(16), "Peak", "-Band_RKE(2)",LimitTitle, CISPR25_ALSEPeakArray_Class5(4), temp_val, 420e6*dFreqUnit, 450e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+6
		PlotLimit (emc_limit_avg(16), "AVG", "-Band_RKE(2)", LimitTitle,CISPR25_ALSEPeakArray_Class5(4), temp_val, 420e6*dFreqUnit, 450e6*dFreqUnit, out_tree)

		'GSM800
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(17), "Peak", "-Band_GSM800", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 860e6*dFreqUnit, 895e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(17), "AVG", "-Band_GSM800",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 860e6*dFreqUnit, 895e6*dFreqUnit, out_tree)

		'EGSM/GSM900
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(18), "Peak", "-Band_EGSM-GSM900",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 925e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(18), "AVG", "-Band_EGSM-GSM900", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 925e6*dFreqUnit, 960e6*dFreqUnit, out_tree)

		'DAB L Band
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(19), "Peak", "-Band_DAB L Band", LimitTitle,CISPR25_ALSEPeakArray_Class5(8), temp_val, 1447e6*dFreqUnit, 1494e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(19), "AVG", "-Band_DAB L Band",LimitTitle, CISPR25_ALSEPeakArray_Class5(8), temp_val, 1447e6*dFreqUnit, 1494e6*dFreqUnit, out_tree)

		'SDARS
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(20), "Peak", "-Band_SDARS", LimitTitle,CISPR25_ALSEPeakArray_Class5(9), temp_val, 2320e6*dFreqUnit, 2345e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(20), "AVG", "-Band_SDARS", LimitTitle,CISPR25_ALSEPeakArray_Class5(9), temp_val, 2320e6*dFreqUnit, 2345e6*dFreqUnit, out_tree)

		'GSM1800
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(21), "Peak", "-Band_GSM1800",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 1803e6*dFreqUnit, 1882e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(21), "AVG", "-Band_GSM1800",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 1803e6*dFreqUnit, 1882e6*dFreqUnit, out_tree)

		'GSM1900
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(22), "Peak", "-Band_GSM1900",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 1850e6*dFreqUnit, 1990e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(22), "AVG", "-Band_GSM1900", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 1850e6*dFreqUnit, 1990e6*dFreqUnit, out_tree)

		'3G/IMT2000(1)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(23), "Peak", "-Band_3G-IMT2000", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 1900e6*dFreqUnit, 1992e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(23), "AVG", "-Band_3G-IMT2000", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 1900e6*dFreqUnit, 1992e6*dFreqUnit, out_tree)

		'3G/IMT2000(2)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(24), "Peak", "-Band_3G-IMT2000(2)", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 2010e6*dFreqUnit, 2025e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(24), "AVG", "-Band_3G-IMT2000(2)",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 2010e6*dFreqUnit, 2025e6*dFreqUnit, out_tree)

		'3G/IMT2000(3)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(25), "Peak", "-Band_3G-IMT2000(3)",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 2108e6*dFreqUnit, 2172e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(25), "AVG", "-Band_3G-IMT2000(3)",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 2108e6*dFreqUnit, 2172e6*dFreqUnit, out_tree)

		'Bluetooth/802.11
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(26), "Peak", "-Band_Bluetooth-802.11",LimitTitle, CISPR25_ALSEPeakArray_Class5(14), temp_val, 2400e6*dFreqUnit, 2500e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(26), "AVG", "-Band_Bluetooth-802.11", LimitTitle,CISPR25_ALSEPeakArray_Class5(14), temp_val, 2400e6*dFreqUnit, 2500e6*dFreqUnit, out_tree)

	Case "4" '"CISPR25: RE-TEM"
		LimitTitle="RE-TEM"
		'LW
		temp_val=limitclass*10
		PlotLimit (emc_limit_peak(0), "Peak", "-Band_LW", LimitTitle,CISPR25_TEMPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(0), "QP", "-Band_LW", LimitTitle,CISPR25_TEMPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=limitclass*10-avgfactor
		PlotLimit (emc_limit_avg(0), "AVG", "-Band_LW",LimitTitle, CISPR25_TEMPeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)

		'MW
		temp_val=limitclass*8
		PlotLimit (emc_limit_peak(1), "Peak", "-Band_MW", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(1), "QP", "-Band_MW", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=limitclass*8-avgfactor
		PlotLimit (emc_limit_avg(1), "AVG", "-Band_MW", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)


		'SW
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(2), "Peak", "-Band_SW", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(2), "QP", "-Band_SW",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(2), "AVG", "-Band_SW", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)

		'CB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(3), "Peak", "-Band_CB",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(3), "QP", "-Band_CB", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(3), "AVG", "-Band_CB", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)

		'VHF1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(4), "Peak", "-Band_VHF1",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(4), "QP", "-Band_VHF1",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(4), "AVG", "-Band_VHF1", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)

		'VHF2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(5), "Peak", "-Band_VHF2", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(5), "QP", "-Band_VHF2", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(5), "AVG", "-Band_VHF2", LimitTitle,CISPR25_TEMPeakArray_Class5(1), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)


		'VHF3
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(6), "Peak", "-Band_VHF3",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(6), "QP", "-Band_VHF3",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(6), "AVG", "-Band_VHF3",LimitTitle, CISPR25_TEMPeakArray_Class5(1), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)

		'FM
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(7), "Peak", "-Band_FM",LimitTitle, CISPR25_TEMPeakArray_Class5(0), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(7), "QP", "-Band_FM",LimitTitle, CISPR25_TEMPeakArray_Class5(0), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(7), "AVG", "-Band_FM", LimitTitle,CISPR25_TEMPeakArray_Class5(0), temp_val, 76e6*dFreqUnit, 108e6*dFreqUnit, out_tree)

		'TVI
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(8), "Peak", "-Band_TVI", LimitTitle,CISPR25_TEMPeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(8), "AVG", "-Band_TVI",LimitTitle, CISPR25_TEMPeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)

		'TVIII
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(9), "Peak", "-Band_TVIII",LimitTitle, CISPR25_TEMPeakArray_Class5(3), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(9), "AVG", "-Band_TVIII",LimitTitle, CISPR25_TEMPeakArray_Class5(3), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)

		'DAB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(10), "Peak", "-Band_DAB",LimitTitle, CISPR25_TEMPeakArray_Class5(4), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(10), "AVG", "-Band_DAB",LimitTitle, CISPR25_TEMPeakArray_Class5(4), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)

	Case "5" '"CISPR25: RE-Stripline"
		LimitTitle="RE-Stripline"
		'LW
		temp_val=limitclass*10
		PlotLimit (emc_limit_peak(0), "Peak", "-Band_LW",LimitTitle, CISPR25_StriplinePeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(0), "QP", "-Band_LW",LimitTitle, CISPR25_StriplinePeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)
		temp_val=limitclass*10-avgfactor
		PlotLimit (emc_limit_avg(0), "AVG", "-Band_LW", LimitTitle,CISPR25_StriplinePeakArray_Class5(0), temp_val, 150000*dFreqUnit, 300000*dFreqUnit, out_tree)

		'MW
		temp_val=limitclass*8
		PlotLimit (emc_limit_peak(1), "Peak", "-Band_MW",LimitTitle, CISPR25_StriplinePeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(1), "QP", "-Band_MW",LimitTitle, CISPR25_StriplinePeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)
		temp_val=limitclass*8-avgfactor
		PlotLimit (emc_limit_avg(1), "AVG", "-Band_MW",LimitTitle, CISPR25_StriplinePeakArray_Class5(1), temp_val, 530000*dFreqUnit, 1.8e6*dFreqUnit, out_tree)

		'SW
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(2), "Peak", "-Band_SW", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(2), "QP", "-Band_SW", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(2), "AVG", "-Band_SW", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 5.9e6*dFreqUnit, 6.2e6*dFreqUnit, out_tree)

		'CB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(3), "Peak", "-Band_CB", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(3), "QP", "-Band_CB", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(3), "AVG", "-Band_CB", LimitTitle,CISPR25_StriplinePeakArray_Class5(1), temp_val, 26e6*dFreqUnit, 28e6*dFreqUnit, out_tree)

		'VHF1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(4), "Peak", "-Band_VHF1", LimitTitle,CISPR25_StriplinePeakArray_Class5(2), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(4), "QP", "-Band_VHF1", LimitTitle,CISPR25_StriplinePeakArray_Class5(2), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(4), "AVG", "-Band_VHF1",LimitTitle, CISPR25_StriplinePeakArray_Class5(2), temp_val, 30e6*dFreqUnit, 54e6*dFreqUnit, out_tree)

		'VHF2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(5), "Peak", "-Band_VHF2",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(5), "QP", "-Band_VHF2",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(5), "AVG", "-Band_VHF2", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 68e6*dFreqUnit, 87e6*dFreqUnit, out_tree)


		'VHF3
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(6), "Peak", "-Band_VHF3", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(6), "QP", "-Band_VHF3", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(6), "AVG", "-Band_VHF3", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 142e6*dFreqUnit, 175e6*dFreqUnit, out_tree)

		'UHF(1)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(7), "Peak", "-Band_UHF(1)", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(7), "QP", "-Band_UHF(1)", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(7), "AVG", "-Band_UHF(1)", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 380e6*dFreqUnit, 512e6*dFreqUnit, out_tree)

		'UHF(2)
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(8), "Peak", "-Band_UHF(2)",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=temp_val-qpfactor
		PlotLimit (emc_limit_qp(8), "QP", "-Band_UHF(2)", LimitTitle,CISPR25_StriplinePeakArray_Class5(6), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(8), "AVG", "-Band_UHF(2)",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 820e6*dFreqUnit, 960e6*dFreqUnit, out_tree)

		'TVI
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(9), "Peak", "-Band_TVI", LimitTitle,CISPR25_StriplinePeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(9), "AVG", "-Band_TVI", LimitTitle,CISPR25_StriplinePeakArray_Class5(3), temp_val, 41e6*dFreqUnit, 88e6*dFreqUnit, out_tree)

		'TVIII
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(10), "Peak", "-Band_TVIII", LimitTitle,CISPR25_StriplinePeakArray_Class5(3), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(10), "AVG", "-Band_TVIII", LimitTitle,CISPR25_StriplinePeakArray_Class5(3), temp_val, 174e6*dFreqUnit, 230e6*dFreqUnit, out_tree)

		'DAB
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(11), "Peak", "-Band_DAB",LimitTitle, CISPR25_StriplinePeakArray_Class5(4), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(11), "AVG", "-Band_DAB", LimitTitle,CISPR25_StriplinePeakArray_Class5(4), temp_val, 171e6*dFreqUnit, 245e6*dFreqUnit, out_tree)

		'TVBandIV
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(12), "Peak", "-Band_TVBandIV", LimitTitle,CISPR25_StriplinePeakArray_Class5(3), temp_val, 468e6*dFreqUnit, 944e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(12), "AVG", "-Band_TVBandIV",LimitTitle, CISPR25_StriplinePeakArray_Class5(3), temp_val, 468e6*dFreqUnit, 944e6*dFreqUnit, out_tree)

		'DTTV
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(13), "Peak", "-Band_DTTV",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 470e6*dFreqUnit, 770e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+10
		PlotLimit (emc_limit_avg(13), "AVG", "-Band_DTTV",LimitTitle, CISPR25_StriplinePeakArray_Class5(6), temp_val, 470e6*dFreqUnit, 770e6*dFreqUnit, out_tree)

		'RKE_1
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(14), "Peak", "-Band_RKE(1)",LimitTitle, CISPR25_StriplinePeakArray_Class5(10), temp_val, 300e6*dFreqUnit, 330e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+6
		PlotLimit (emc_limit_avg(14), "AVG", "-Band_RKE(1)",LimitTitle, CISPR25_StriplinePeakArray_Class5(10), temp_val, 300e6*dFreqUnit, 330e6*dFreqUnit, out_tree)

		'RKE_2
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(15), "Peak", "-Band_RKE(2)", LimitTitle,CISPR25_StriplinePeakArray_Class5(10), temp_val, 420e6*dFreqUnit, 450e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor+6
		PlotLimit (emc_limit_avg(15), "AVG", "-Band_RKE(2)", LimitTitle,CISPR25_StriplinePeakArray_Class5(10), temp_val, 420e6*dFreqUnit, 450e6*dFreqUnit, out_tree)

		'GSM800
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(16), "Peak", "-Band_GSM800", LimitTitle,CISPR25_StriplinePeakArray_Class5(2), temp_val, 860e6*dFreqUnit, 895e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(16), "AVG", "-Band_GSM800", LimitTitle,CISPR25_StriplinePeakArray_Class5(2), temp_val, 860e6*dFreqUnit, 895e6*dFreqUnit, out_tree)

		'EGSM/GSM900
		temp_val=limitclass*6
		PlotLimit (emc_limit_peak(17), "Peak", "-Band_EGSM-GSM900",LimitTitle, CISPR25_StriplinePeakArray_Class5(2), temp_val, 925e6*dFreqUnit, 960e6*dFreqUnit, out_tree)
		temp_val=limitclass*6-avgfactor
		PlotLimit (emc_limit_avg(17), "AVG", "-Band_EGSM-GSM900", LimitTitle,CISPR25_StriplinePeakArray_Class5(2), temp_val, 925e6*dFreqUnit, 960e6*dFreqUnit, out_tree)

	End Select

End Sub

Sub PlotLimit (LimitObject As Object, detector As String, band_name As String, sYAxisLabel As String,LimitVal As Double, factor As Double, freq_start As Double, freq_stop As Double, outtree As String)

	Dim YAxisUnit As String, YAxisLabel As String, ResultFileName As String

	Select Case sYAxisLabel
		Case "RE-ALSE"
			YAxisUnit="V/m"
			YAxisLabel="E-field"
            ResultFileName="RE-ALSE"
        Case "RE-TEM"
        	YAxisUnit="V/m"
        	YAxisLabel="E-field"
            ResultFileName="RE-TEM"
        Case "RE-Stripline"
        	YAxisUnit="V/m"
        	YAxisLabel="E-field"
            ResultFileName="RE-Stripline"
		Case "Conducted Emission (I-Method)"
			YAxisUnit="A"
			YAxisLabel="Current"
			ResultFileName="CE_I"
		Case "Conducted Emission (V-Method)"
	        YAxisUnit="V"
	        YAxisLabel="Voltage"
	        ResultFileName="CE_V"
	End Select

	With LimitObject
		.appendxydouble(freq_start,10^((LimitVal+factor-120)/20),0)
		.appendxydouble(freq_stop,10^((LimitVal+factor-120)/20),0)
		.SetYLabelAndUnit (YAxisLabel, YAxisUnit)
		.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
		.Save GetProjectPath("Result") +  "CISPR25_"+ResultFileName+"_"+strEMCClassApp+"_"+detector+band_name
		.AddToTree outtree + detector+ band_name
	End With

	If WhereAmI ="DS" Then
		If detector = "Peak" Then
			DS.SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolpeak)
		ElseIf detector = "QP" Then
			DS.SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolqp)
		ElseIf detector = "AVG" Then
			DS.SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolavg)
		End If
		PlotLogAxis (outtree,True)
	Else
		If detector = "Peak" Then
			SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolpeak)
		ElseIf detector = "QP" Then
			SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolqp)
		ElseIf detector = "AVG" Then
			SetPlotStyleForTreeItem(outtree +detector+band_name,RGBcolavg)
		End If
		PlotLogAxis (outtree,False)
	End If

End Sub

Sub PlotEMILimitFCCbCE (fccbCE_emc_limit_peak As Object, fccbCE_emc_limit_qp As Object, fccbCE_emc_limit_avg As Object)
	With fccbCE_emc_limit_qp
		.appendxydouble(150000*dFreqUnit,10^((79-120)/20),0)
		.appendxydouble(500000*dFreqUnit,10^((79-120)/20),0)
		.appendxydouble(500000*dFreqUnit,10^((73-120)/20),0)
		.appendxydouble(30000000*dFreqUnit,10^((73-120)/20),0)
		.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
		.SetYLabelAndUnit ("Voltage","V")
		.Save GetProjectPath("Result") +  "QP-FCC15b_CE"
		.AddToTree out_tree + "QP"
	End With
	With fccbCE_emc_limit_avg
		.appendxydouble(150000*dFreqUnit,10^((66-120)/20),0)
		.appendxydouble(500000*dFreqUnit,10^((66-120)/20),0)
		.appendxydouble(500000*dFreqUnit,10^((60-120)/20),0)
		.appendxydouble(30000000*dFreqUnit,10^((60-120)/20),0)
		.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
		.SetYLabelAndUnit ("Voltage","V")
		.Save GetProjectPath("Result") +  "AVG-FCC15b_CE"
		.AddToTree out_tree + "AVG"
	End With

	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		DS.SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,False)
	End If
End Sub


Sub PlotEMILimitFCCbRE (fccbRE_emc_limit_peak As Object, fccbRE_emc_limit_qp As Object, fccbRE_emc_limit_avg As Object)

	Select Case strEMCClassApp
		Case "Class A 3m"
			With fccbRE_emc_limit_qp
				.appendxydouble(30e6*dFreqUnit,10^((49.5-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((49.5-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((54-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((54-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((57-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((57-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((60-120)/20),0)
				.appendxydouble(1000e6*dFreqUnit,10^((60-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "QP-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "QP"
			End With
			With fccbRE_emc_limit_avg
				.appendxydouble(1000e6*dFreqUnit,10^((60-120)/20),0)
				.appendxydouble(40e9*dFreqUnit,10^((60-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "AVG-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "AVG"
			End With
		Case "Class A 10m"
			With fccbRE_emc_limit_qp
				.appendxydouble(30e6*dFreqUnit,10^((39-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((39-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((46.5-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((46.5-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((49.5-120)/20),0)
				.appendxydouble(1000e6*dFreqUnit,10^((49.5-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "QP-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "QP"
			End With
			With fccbRE_emc_limit_avg
				.appendxydouble(1000e6*dFreqUnit,10^((49.5-120)/20),0)
				.appendxydouble(40e9*dFreqUnit,10^((49.5-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "AVG-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "AVG"
			End With
		Case "Class B 3m"
			With fccbRE_emc_limit_qp
				.appendxydouble(30e6*dFreqUnit,10^((40-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((40-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((46-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((46-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((54-120)/20),0)
				.appendxydouble(1000e6*dFreqUnit,10^((54-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "QP-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "QP"
			End With
			With fccbRE_emc_limit_avg
				.appendxydouble(1000e6*dFreqUnit,10^((54-120)/20),0)
				.appendxydouble(40e9*dFreqUnit,10^((54-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "AVG-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "AVG"
			End With
		Case "Class B 10m"
			With fccbRE_emc_limit_qp
				.appendxydouble(30e6*dFreqUnit,10^((29.5-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((29.5-120)/20),0)
				.appendxydouble(88e6*dFreqUnit,10^((33-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((33-120)/20),0)
				.appendxydouble(216e6*dFreqUnit,10^((35.5-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((35.5-120)/20),0)
				.appendxydouble(960e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(1000e6*dFreqUnit,10^((43.5-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "QP-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "QP"
			End With
			With fccbRE_emc_limit_avg
				.appendxydouble(1000e6*dFreqUnit,10^((43.5-120)/20),0)
				.appendxydouble(40e9*dFreqUnit,10^((43.5-120)/20),0)
				.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
				.SetYLabelAndUnit ("E-field","V/m")
				.Save GetProjectPath("Result") +  "AVG-FCC15b_RE_" + strEMCClassApp
				.AddToTree out_tree + "AVG"
			End With
	End Select

	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		DS.SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,False)
	End If


End Sub

Sub PlotEMILimitCISPR32_CE(csipr32CE_emc_limit_peak As Object,csipr32CE_emc_limit_qp As Object,csipr32CE_emc_limit_avg As Object)
	Select Case strEMCClassApp
	Case "Class A Voltage limit (Main Ports)"
		With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP(1)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(0)-13-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(0)-13-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP(1)-13-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A Voltage limit (Telecom-LAN Ports)"
       With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(1)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(0)-13-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_V_telecom(1)-13-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A Current limit (Telecom-LAN Ports)"
       With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Current","A")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(0)-13-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-13-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Current","A")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	Case "Class B Voltage limit (Main Ports)"
		With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassB_QP(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassB_QP(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(2)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(2)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassB_QP(0)-10-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassB_QP(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(2)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP(2)-10-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	Case "Class B Voltage limit (Telecom-LAN Ports)"
       With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(0)-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(1)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(0)-10-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(1)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassB_QP_V_telecom(1)-10-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Voltage","V")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	Case "Class B Current limit (Telecom-LAN Ports)"
       With csipr32CE_emc_limit_qp
			.appendxydouble(150000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(0)-13-120)/20),0)
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-13-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Current","A")
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32CE_emc_limit_avg
			.appendxydouble(500000*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-23-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_CE_ClassA_QP_I_telecom(1)-23-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("Current","A")
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_CE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
	End Select

	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		DS.SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,False)
	End If
End Sub
Sub PlotEMILimitCISPR32_RE(csipr32RE_emc_limit_peak As Object,csipr32RE_emc_limit_qp As Object,csipr32RE_emc_limit_avg As Object)

	Select Case strEMCClassApp
	Case "Class A 3m"
		With csipr32RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32RE_emc_limit_avg
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-20-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-20-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-20-120)/20),0)
			.appendxydouble(6e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-20-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
		With csipr32RE_emc_limit_peak
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-120)/20),0)
			.appendxydouble(6e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "Peak-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "Peak"
		End With
	Case "Class A 10m"
		With csipr32RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-10.5-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-10.5-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-10.5-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-10.5-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
	Case "Class B 3m"
		With csipr32RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-10-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-10-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
		With csipr32RE_emc_limit_avg
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-26-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-26-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-26-120)/20),0)
			.appendxydouble(6e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-26-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "AVG"
		End With
		With csipr32RE_emc_limit_peak
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-6-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(0)-6-120)/20),0)
			.appendxydouble(3e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-6-120)/20),0)
			.appendxydouble(6e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mPeak(1)-6-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "Peak-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "Peak"
		End With
	Case "Class B 10m"
		With csipr32RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-20.5-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(0)-20.5-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-20.5-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR32_RE_ClassA_3mQP(1)-20.5-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR32-EN5032_RE_"+strEMCClassApp
			.AddToTree out_tree + "QP"
		End With
	End Select
	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		DS.SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		DS.SetPlotStyleForTreeItem(out_tree + "Peak",RGBcolpeak)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		SetPlotStyleForTreeItem(out_tree + "Peak",RGBcolpeak)
		PlotLogAxis (out_tree,False)
	End If
End Sub

Sub PlotEMILimitCISPR11_CE (csipr11CE_emc_limit_peak, csipr11CE_emc_limit_qp, csipr11CE_emc_limit_avg)
	Dim tempclass As String
	tempclass=strEMCClassApp

	Select Case strEMCClassApp
	Case "Class A Main Ports (>75kVA)"
		tempclass=Replace(tempclass,">","greater")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(2)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(2)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(2)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater75kVA(2)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With

	Case "Class A Main Ports (>20kVA & <=75kVA)"
		tempclass=Replace(tempclass,">","greater")
		tempclass=Replace(tempclass,"<=","less")
		tempclass=Replace(tempclass," & ","")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(2)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(3)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(2)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_greater20kVAless75kVA(3)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A Main Ports (<=20kVA)"
		tempclass=Replace(tempclass,"<=","less")
		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(1)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(0)-13-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(0)-13-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_MainPorts_QP_less20kVA(1)-13-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class B Main Ports"
		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(1)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(2)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(2)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(1)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(2)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassB_MainPorts_QP(2)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A DCPorts V-Limit (>75kVA)"
		tempclass=Replace(tempclass,">","greater")
		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(0)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(2)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(0)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(1)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Vlim(2)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A DCPorts V-Limit (>20kVA & <=75kVA)"
		tempclass=Replace(tempclass,">","greater")
		tempclass=Replace(tempclass,"<=","less")
		tempclass=Replace(tempclass," & ","")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(0)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(2)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(0)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(1)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Vlim(2)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A DCPorts V-Limit (<=20kVA)"
		tempclass=Replace(tempclass,"<=","less")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(0)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(1)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(0)-13-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(1)-13-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_less20kVA_Vlim(1)-13-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class B DCPorts V-Limit"
		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(0)-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(1)-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(0)-10-120)/20),0)
			.appendxydouble(0.5e6*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(1)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassB_DCPorts_QP_Vlim(1)-10-120)/20),0)
			.SetYLabelAndUnit ("Voltage","V")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With
	Case "Class A DCPorts I-Limit (>75kVA)"
		tempclass=Replace(tempclass,">","greater")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(0)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(2)-120)/20),0)
			.SetYLabelAndUnit ("Current","A")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(0)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(1)-11-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater75kVA_Ilim(2)-13-120)/20),0)
			.SetYLabelAndUnit ("Current","A")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With

	Case "Class A DCPorts I-Limit (>20kVA & <=75kVA)"
		tempclass=Replace(tempclass,">","greater")
		tempclass=Replace(tempclass,"<=","less")
		tempclass=Replace(tempclass," & ","")

		With csipr11CE_emc_limit_qp
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(0)-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(1)-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(2)-120)/20),0)
			.SetYLabelAndUnit ("Current","A")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
		With csipr11CE_emc_limit_avg
			.appendxydouble(150e3*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(0)-10-120)/20),0)
			.appendxydouble(5e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(1)-10-120)/20),0)
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_CE_ClassA_DCPorts_QP_greater20kVAless75kVA_Ilim(2)-13-120)/20),0)
			.SetYLabelAndUnit ("Current","A")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "AVG-CISPR11_CE_"+tempclass
			.AddToTree out_tree + "AVG"
		End With

	End Select
	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		DS.SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		SetPlotStyleForTreeItem(out_tree + "AVG",RGBcolavg)
		PlotLogAxis (out_tree,False)
	End If

End Sub

Sub PlotEMILimitCISPR11_RE (csipr11RE_emc_limit_peak, csipr11RE_emc_limit_qp, csipr11RE_emc_limit_avg)

	Dim tempclass As String
	tempclass=strEMCClassApp

	Select Case strEMCClassApp
	Case "Class A 3m (>20kVA)"
		tempclass=Replace(tempclass,">","greater")
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_greater20kVA(0)-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_greater20kVA(0)-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	Case "Class A 10m (>20kVA)"
		tempclass=Replace(tempclass,">","greater")
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_greater20kVA(0)-10-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_greater20kVA(0)-10-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	Case "Class A 3m (<=20kVA)"
		tempclass=Replace(tempclass,"<=","less")
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	Case "Class A 10m (<=20kVA)"
		tempclass=Replace(tempclass,"<=","less")
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-10-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-10-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	Case "Class B 3m"
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-10-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-10-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-10-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	Case "Class B 10m"
		With csipr11RE_emc_limit_qp
			.appendxydouble(30e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-20-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(0)-20-120)/20),0)
			.appendxydouble(230e6*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-20-120)/20),0)
			.appendxydouble(1e9*dFreqUnit,10^((CISPR11_RE_ClassA_QP3m_less20kVA(1)-20-120)/20),0)
			.SetYLabelAndUnit ("E-field","V/m")
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.Save GetProjectPath("Result") +  "QP-CISPR11_RE_"+tempclass
			.AddToTree out_tree + "QP"
		End With
	End Select
	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "QP",RGBcolqp)
		PlotLogAxis (out_tree,False)
	End If

End Sub

Sub PlotEMILimitMILSTD_CE102(ce102_emc_limit As Object)
	Dim limit_relax As Double

	Select Case strEMCClassApp
		Case "SourceVoltage 28V"
			limit_relax=0
		Case "SourceVoltage 115V"
			limit_relax=6
		Case "SourceVoltage 220V"
			limit_relax=9
		Case "SourceVoltage 270V"
			limit_relax=10
		Case "SourceVoltage 440V"
			limit_relax=12
	End Select

	With ce102_emc_limit
		.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_CE102(0)+limit_relax-120)/20),0)
		.appendxydouble(500e3*dFreqUnit,10^((MIL_STD_461_CE102(1)+limit_relax-120)/20),0)
		.appendxydouble(10e6*dFreqUnit,10^((MIL_STD_461_CE102(1)+limit_relax-120)/20),0)
		.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
		.SetYLabelAndUnit ("Voltage","V")
		.Save GetProjectPath("Result") +  "MIL-CE102_" + strEMCClassApp
		.AddToTree out_tree + "CE102Limit"
	End With

	If WhereAmI ="DS" Then
		DS.SetPlotStyleForTreeItem(out_tree + "CE102Limit",RGBcolqp)
		PlotLogAxis (out_tree,True)
	Else
		SetPlotStyleForTreeItem(out_tree + "CE102Limit",RGBcolqp)
		PlotLogAxis (out_tree,False)
	End If
End Sub

Sub PlotEMILimitMILSTD_RE102(re102_emc_limit_a As Object ,re102_emc_limit_b As Object ,re102_emc_limit_c As Object)

	Select Case strEMCClassApp
	Case "Surface ship"
		With re102_emc_limit_a
			.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(0)-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(1)-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(2)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_belowdeck_" + strEMCClassApp
			.AddToTree out_tree + "Below deck"
		End With
		With re102_emc_limit_b
			.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(0)-20-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(1)-20-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_ShipApp(2)-20-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_abovedeck_" + strEMCClassApp
			.AddToTree out_tree + "Above deck and exposed below deck"
		End With
		If WhereAmI ="DS" Then
			DS.SetPlotStyleForTreeItem(out_tree + "Below deck",RGBcolpeak)
			DS.SetPlotStyleForTreeItem(out_tree + "Above deck and exposed below deck",RGBcolqp)
			PlotLogAxis (out_tree,True)
		Else
			SetPlotStyleForTreeItem(out_tree + "Below deck",RGBcolpeak)
			SetPlotStyleForTreeItem(out_tree + "Above deck and exposed below deck",RGBcolqp)
			PlotLogAxis (out_tree,False)
		End If
	Case "Submarine"
		With re102_emc_limit_a
			.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppInternal(0)-120)/20),0)
			.appendxydouble(800e3*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppInternal(1)-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppInternal(1)-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppInternal(2)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_interal_" + strEMCClassApp
			.AddToTree out_tree + "Internal to pressure hull"
		End With
		With re102_emc_limit_b
			.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppExternal(0)-120)/20),0)
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppExternal(1)-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppExternal(1)-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_SubmarAppExternal(2)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_external_" + strEMCClassApp
			.AddToTree out_tree + "External to pressure hull"
		End With
		If WhereAmI ="DS" Then
			DS.SetPlotStyleForTreeItem(out_tree + "Internal to pressure hull",RGBcolpeak)
			DS.SetPlotStyleForTreeItem(out_tree + "External to pressure hull",RGBcolqp)
			PlotLogAxis (out_tree,True)
		Else
			SetPlotStyleForTreeItem(out_tree + "Internal to pressure hull",RGBcolpeak)
			SetPlotStyleForTreeItem(out_tree + "External to pressure hull",RGBcolqp)
			PlotLogAxis (out_tree,False)
		End If
	Case "Aircraft and space system"
		With re102_emc_limit_a
			.appendxydouble(10e3*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(0)-120)/20),0)
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(2)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_external_" + strEMCClassApp
			.AddToTree out_tree + "Fixed wing external and helicopter"
		End With
		With re102_emc_limit_b
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)+10-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)+10-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(2)+10-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_internal_less25m_" + strEMCClassApp
			.AddToTree out_tree + "Fixed wing internal (<25m nose to tail)"
		End With
		With re102_emc_limit_c
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)+20-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(1)+20-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_AirSpaceApp(2)+20-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_internal_greater25m_" + strEMCClassApp
			.AddToTree out_tree + "Fixed wing internal (>25m nose to tail)"
		End With
		If WhereAmI ="DS" Then
			DS.SetPlotStyleForTreeItem(out_tree + "Fixed wing external and helicopter",RGBcolpeak)
			DS.SetPlotStyleForTreeItem(out_tree + "Fixed wing internal (<25m nose to tail)",RGBcolqp)
			DS.SetPlotStyleForTreeItem(out_tree + "Fixed wing internal (>25m nose to tail)",RGBcolqp)
			PlotLogAxis (out_tree,True)
		Else
			SetPlotStyleForTreeItem(out_tree + "Fixed wing external and helicopter",RGBcolpeak)
			SetPlotStyleForTreeItem(out_tree + "Fixed wing internal (<25m nose to tail)",RGBcolqp)
			SetPlotStyleForTreeItem(out_tree + "Fixed wing internal (>25m nose to tail)",RGBcolqp)
			PlotLogAxis (out_tree,False)
		End If
	Case "Ground"
		With re102_emc_limit_a
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(0)-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(0)-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(1)-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_fixed_" + strEMCClassApp
			.AddToTree out_tree + "Navy fixed and air force"
		End With
		With re102_emc_limit_b
			.appendxydouble(2e6*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(0)-20-120)/20),0)
			.appendxydouble(100e6*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(0)-20-120)/20),0)
			.appendxydouble(18e9*dFreqUnit,10^((MIL_STD_461_RE102_GroundApp(1)-20-120)/20),0)
			.SetXLabelAndUnit ("Freq",Units.GetUnit("Frequency"))
			.SetYLabelAndUnit ("E-field","V/m")
			.Save GetProjectPath("Result") +  "MIL-RE102_mobile_" + strEMCClassApp
			.AddToTree out_tree + "Navy mobile and army"
		End With
		If WhereAmI ="DS" Then
			DS.SetPlotStyleForTreeItem(out_tree + "Navy fixed and air force",RGBcolpeak)
			DS.SetPlotStyleForTreeItem(out_tree + "Navy mobile and army",RGBcolqp)
			PlotLogAxis (out_tree,True)
		Else
			SetPlotStyleForTreeItem(out_tree + "Navy fixed and air force",RGBcolpeak)
			SetPlotStyleForTreeItem(out_tree + "Navy mobile and army",RGBcolqp)
			PlotLogAxis (out_tree,False)
		End If
	End Select
End Sub


Function DialogExample()
	Begin Dialog UserDialog 570,441,"EMI Standard Limit",.DialogFuncExampleData ' %GRID:10,7,1,1
		TextBox 20,14,530,385,.TextBoxExample,2
		PushButton 460,413,90,21,"Close",.PushButtonClose
		PushButton 320,413,130,21,"Copy to clipboard",.PushButtonClipBoard
		CancelButton 20,413,90,21
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then 'Do Nothing
	End If

End Function

Rem See DialogFunc help topic for more information.
Private Function DialogFuncExampleData(DlgItem$, Action%, SuppValue?) As Boolean
	DlgVisible("Cancel",False)
	Dim ExampleContent As String

	FillContent(ExampleContent)

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("TextBoxExample",ExampleContent)
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "PushButtonClose"
			Exit Function
		Case "PushButtonClipBoard"
			Clipboard ExampleContent
			DialogFuncExampleData = True
		End Select
		Rem DialogFuncExampleData = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFuncExampleData = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub FillContent (sContent As String)
	sContent="*This is a comment section"+vbNewLine
	sContent=sContent+"*Several comments can be inserted for user information"+vbNewLine
	sContent=sContent+"*Colum content: band;freq start;freq stop;amplitude start;amplitude Stop"+vbNewLine
	sContent=sContent+"*Supported amplitude unit: dBu, dB, or linear"+vbNewLine
	sContent=sContent+"*Supported frequency unit: Hz, kHz, MHz, GHz, THz"+vbNewLine
	sContent=sContent+"*Keyword Title is required"+vbNewLine
	sContent=sContent+"[Title OEM RE 10m Distance]"+vbNewLine
	sContent=sContent+"[Peak]"+vbNewLine
	sContent=sContent+"[Hz;dBu]"+vbNewLine
	sContent=sContent+"LW;0.15e6;0.2e6;89;89"+vbNewLine
	sContent=sContent+"LW;0.2e6;0.28e6;90;90"+vbNewLine
	sContent=sContent+"AM;0.52e6;30e6;73;73"+vbNewLine
	sContent=sContent+"FM;76e6;108e6;34;37"+vbNewLine
	sContent=sContent+"TVIII;174e6;230e6;67;67"+vbNewLine
	sContent=sContent+"[----end----]"+vbNewLine
	sContent=sContent+"[QP]"+vbNewLine
	sContent=sContent+"[Hz;dBu]"+vbNewLine
	sContent=sContent+"LW;0.15e6;0.2e6;80;80"+vbNewLine
	sContent=sContent+"LW;0.2e6;0.28e6;91;91"+vbNewLine
	sContent=sContent+"AM;0.52e6;30e6;63;63"+vbNewLine
	sContent=sContent+"FM;76e6;108e6;24;27"+vbNewLine
	sContent=sContent+"TVIII;174e6;230e6;57;57"+vbNewLine
	sContent=sContent+"[----end----]"+vbNewLine
	sContent=sContent+"[AVG]"+vbNewLine
	sContent=sContent+"[Hz;dBu]"+vbNewLine
	sContent=sContent+"LW;0.15e6;0.2e6;69;69"+vbNewLine
	sContent=sContent+"LW;0.2e6;0.28e6;70;70"+vbNewLine
	sContent=sContent+"AM;0.52e6;30e6;53;53"+vbNewLine
	sContent=sContent+"FM;76e6;108e6;14;17"+vbNewLine
	sContent=sContent+"TVIII;174e6;230e6;47;47"+vbNewLine
	sContent=sContent+"[----end----]"+vbNewLine
End Sub
