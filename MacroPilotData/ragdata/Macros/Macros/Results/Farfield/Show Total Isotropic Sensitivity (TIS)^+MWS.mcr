' Show Total Isotropic Sensitivity (TIS)
'
' This macro computes and displays the Total Isotropic Sensitivity (TIS) of an antenna. 
' It can be specified at which port the receiver gets connected, and in case of a 
' waveguide port which mode is used (for discrete ports mode number equals 1). 
' The power can be specified in Watt (average) or dBmW (average).
'
' In case of a system simulation, the combined result should be calculated before 
' the macro gets started. This can be done under "Results -> Combine Results") 
' or within an S-Parameter Task in CST DESIGN STUDIO. In the second case the farfields 
' need to be combined within the combine results tab in the S-Parameter Task definition.

' The TIS value finally gets displayed within the farfield monitor either in Watt or dBm.

' ================================================================================================
' Copyright 2008-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------------
' 17-Jul-2018 ube: SetSpeciels changed into SetSpecials
' 22-Jan-2014 twi: r m s  labels removed from dialogue
' 14-Jun-2010 ube: slight change in dialog text (simult.excitation included)
' 06-Jan-2010 ube: same fix as in TRP macro concerning DialogFunc
' 02-Dec-2008 twi: first version based on TRP macro
'------------------------------------------------------------------------------------------

Const HelpFileName = "common_preloadedmacro_farfield_TIS"

Sub Main ()

	Dim Read_Power As String
	Dim Power As Double
	Dim Scaled_Power As Double
	Dim Port_Number As String
	Dim Mode_Number As String
	Dim Port As Integer
	Dim Mode As Integer
	Dim Mon_Name As String

	Begin Dialog UserDialog 400,266,"TIS",.DialogFunc ' %GRID:10,2,1,1
		Text 30,14,360,14,"Show Total Isotropic Sensitivity (TIS) in Farfield Plot:",.hhh
		GroupBox 20,154,360,70,"Show TIS",.GroupBox2
		OKButton 20,238,90,21
		OptionGroup .Group1
			OptionButton 40,176,110,14,"in Watt",.linear
			OptionButton 220,176,100,14,"in dBmW",.db
			OptionButton 40,203,130,14,"Switch TIS Off",.swoff
		CancelButton 120,238,90,21
'not yet		PushButton 290,238,90,21,"Help",.Help
		GroupBox 20,36,360,106,"Define Receiver Sensitivity",.GroupBox1
		OptionGroup .GroupCombine
			OptionButton 40,64,50,14,"Use",.OptionButton1
			OptionButton 40,92,50,14,"Use",.OptionButton2
			OptionButton 40,119,330,14,"Use existing Combined Result / Simult.Excitation",.OptionButton3
		TextBox 100,60,40,20,.pWatt
		TextBox 100,88,40,20,.dbm
		Text 150,63,70,14,"pW",.TextWatt
		Text 150,91,70,14,"dBmW",.Textdbm
		Text 210,78,20,14,"at",.Text4
		Text 260,63,30,14,"Port:",.Text5
		Text 260,91,40,14,"Mode:",.Text6
		TextBox 310,60,30,20,.Port
		TextBox 310,88,30,20,.Mode
	End Dialog


	Dim dlg As UserDialog

	dlg.Group1 = 1
	dlg.GroupCombine = 1
	dlg.pWatt = "1"
	dlg.dbm = "-90"
	dlg.Port = "1"
	dlg.Mode = "1"

	If (Dialog(dlg) = 0) Then Exit All

	Port_Number = dlg.Port
	Mode_Number = dlg.Mode
	Port = Evaluate (Port_Number)
	Mode = Evaluate (Mode_Number)


		Select Case dlg.GroupCombine
			Case 0 ' linear power
				Read_Power = dlg.pWatt
				Power = Evaluate (Read_Power)
				Scaled_Power = Sqr(1e-12*Power*2)
				Mon_Name = Read_Power + "pW at Port" + Port_Number + ",Mode" + Mode_Number
				With CombineResults
   					 .Reset
    				 .SetMonitorType ("frequency")
    				 .FarfieldsOnly (True)
    				 .EnableAutomaticLabeling (False)
    				 .SetLabel (Mon_Name)
    				 .SetPortModeValues (Port, Mode, Scaled_Power, 0)
            		 .Run
				End With


			Case 1 ' in dB power
				Read_Power = dlg.dbm
				Power = Evaluate (Read_Power)
				Scaled_Power = Sqr(0.001*10^(Power/10)*2)
				Mon_Name = Read_Power + "dBmW at Port" + Port_Number + ",Mode" + Mode_Number
				With CombineResults
   					 .Reset
    				 .SetMonitorType ("frequency")
    				 .FarfieldsOnly (True)
    				 .EnableAutomaticLabeling (False)
    				 .SetLabel (Mon_Name)
    				 .SetPortModeValues (Port, Mode, Scaled_Power, 0)
            		 .Run
				End With

		End Select

		Select Case dlg.Group1
			Case 0 ' linear
				AddToHistoryNoModelChange ("Plot TRP", "Farfieldplot.SetSpecials " + Chr(34) +"showtis" + Chr(34))

			Case 1 ' in dB
				AddToHistoryNoModelChange ("Plot TRP", "Farfieldplot.SetSpecials " + Chr(34) +"showtisdb" + Chr(34))

			Case 2 ' given rbeta
				AddToHistoryNoModelChange ("Plot TRP", "Farfieldplot.SetSpecials " + Chr(34) +"showtisoff" + Chr(34))


		End Select

End Sub
Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	If (Action% = 1 Or Action% = 2) Then

		Dim iCOMB As Integer
		iCOMB = DlgValue("GroupCombine")

		DlgEnable "pWatt"    , IIf(iCOMB = 0, 1, 0)
		DlgEnable "TextWatt", IIf(iCOMB = 0, 1, 0)
		DlgEnable "dbm"     , IIf(iCOMB = 1, 1, 0)
		DlgEnable "Textdbm" , IIf(iCOMB = 1, 1, 0)

		DlgEnable "Text4" , IIf(iCOMB = 2, 0, 1)
		DlgEnable "Text5" , IIf(iCOMB = 2, 0, 1)
		DlgEnable "Text6" , IIf(iCOMB = 2, 0, 1)
		DlgEnable "Port" , IIf(iCOMB = 2, 0, 1)
		DlgEnable "Mode" , IIf(iCOMB = 2, 0, 1)

		If (DlgItem = "Help") Then
			StartHelp HelpFileName
			DialogFunc = True
		End If

	End If

End Function
