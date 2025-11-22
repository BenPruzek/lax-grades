' Time gating for output signal

' ================================================================================================
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 22-Mar-2007 jwa: add item "Time Gating"
' 14-Mar-2007 jwa: initial time gating macro, change original out put signal .
'------------------------------------------------------------------------------------

Const HelpFileName = ""

Option Explicit
'#include "vba_globals_all.lib"

Public SignalList(50) As String
Public cst_from_port() As String
Public cst_to_port() As String
Public cst_from_mode() As String
Public cst_to_mode() As String
Public CST_NrOfPorts As Integer
Public cst_signalsname As String
Public cst_treename As String
Sub Main
	Dim cst_input_signal As Object, cst_output_signal As Object
	Dim cst_signal_name As String
	Dim cst_runindex As Long
	Dim cst_nfsteps As Long
	Dim cst_tstart As Double, cst_tend As Double
	Dim cst_nfreq As Long, cst_ii As Long
	Dim cst_icount As Integer, cst_strtmp As String
	Dim cst_iii As Integer
	Dim cst_outmodes As Long, cst_inmodes As Long
	Dim cst_port As Integer
	Dim cst_inport As String, cst_inmode As String
	Dim cst_outport As String, cst_outmode As String
	Dim cst_prestring As String
	Dim cst_nports As Long
	Dim in_tname As String, in_fname As String,out_tname As String, out_fname As String
	Dim BInvertselect As Boolean
	Dim i_inport As Long, i_inmode As Long, i_outport As Long, i_outmode As Long
	Dim cst_portmode As String
	Dim inport_low As Long, inport_high As Long, inmode_low As Long, inmode_high As Long
	Dim outport_low As Long, outport_high As Long, outmode_low As Long, outmode_high As Long


	Begin Dialog UserDialog 400,243,"Time gating for time signal",.DialogFunc ' %GRID:10,3,1,1
		OKButton 10,213,90,21
		GroupBox 10,6,380,99,"Excitation Settings",.GroupBox1
		CancelButton 110,213,90,21
		DropListBox 70,48,130,192,cst_from_port(),.FromPort
		DropListBox 260,48,120,192,cst_to_port(),.ToPort
		DropListBox 70,75,130,192,cst_from_mode(),.FromMode
		DropListBox 260,75,120,192,cst_to_mode(),.ToMode
		Text 20,51,40,15,"Port",.Text4
		Text 210,51,40,15,"Port",.Text10
		Text 20,78,40,15,"Mode",.Text5
		Text 210,75,40,15,"Mode",.Text11
		GroupBox 10,114,380,90,"Time window",.GroupBox4
		CheckBox 40,171,140,30,"Invert selection",.InverSelect
		TextBox 40,147,90,21,.tstart
		TextBox 200,147,90,21,.tend
		Text 40,129,100,15,"Window Start:",.Text6
		Text 200,129,140,15,"Window End",.Text7
		Text 20,27,90,15,"From",.Text8
		Text 210,27,90,15,"To",.Text9
	'	PushButton 270,213,90,21,"Help",.Help
	End Dialog

	'--- Get number of Ports -----------------------------------------------------
		cst_nports	= Solver.GetNumberOfPorts
		ReDim cst_from_port(cst_nports-1)
		ReDim cst_to_port(cst_nports-1)
		For cst_runindex = 0 To cst_nports-1 	'assign Portnumbers to array
	 		cst_from_port(cst_runindex) = CStr(cst_runindex+1)
	  	   	cst_to_port(cst_runindex) = CStr(cst_runindex+1)
		Next cst_runindex

	 'Set number of modes For Port Nr. 1
		Enlarge_Mode_List cst_from_mode(), "1"
		Enlarge_Mode_List cst_to_mode(), "1"


	Dim dlg As UserDialog

         dlg.tstart    =  "0"
         dlg.tend   =  "0"
         dlg.InverSelect=False
	'--- Open Dialogue
	If (Dialog(dlg) = 0) Then Exit All

		cst_inport = cst_from_port(dlg.FromPort)
		cst_inmode = cst_from_mode(dlg.FromMode)
		cst_outport = cst_to_port(dlg.ToPort)
		cst_outmode = cst_to_mode(dlg.ToMode)
        cst_tstart    = Eval(dlg.tstart)
        cst_tend  = Eval(dlg.tend)
        BInvertselect=dlg.InverSelect

       If cst_tstart > cst_tend Then

        MsgBox "Ending time is smaller than starting time!"
        Exit All

       End If


	Dim Sinp As Object,Soutp As Object, sInput As String, sOutput As String
	Dim ii As Long,nn As Long,tx As Double,ty As Double

    sInput = CStr(cst_inport) + "(" + CStr(cst_inmode) + ")"
	sOutput = CStr(	cst_outport ) + "(" + CStr(cst_outmode) + ")" + CStr(cst_inport) + "(" + CStr(cst_inmode) + ")"
    Set Sinp  = Result1D("i" + sInput)
	Set Soutp = Result1D("o" + sOutput)
	
	Sinp.AddToTree("1D Results\Time Gating\Port signals\i" + sInput)
	With Soutp
		nn = .GetN

		For ii = 0 To nn-1
			tx=.GetX(ii)
			ty= .GetY(ii)

			If BInvertselect Then
				.SetY ( ii, IIf(tx > cst_tstart And tx < cst_tend, 0, ty ))
			Else
				.SetY ( ii, IIf(tx > cst_tstart And tx < cst_tend, ty, 0))
			End If
		Next ii
		.Save ("^o"+sOutput +"_gating"+ ".sig")
		.AddToTree("1D Results\Time Gating\Port signals\o" + sOutput)
	End With

    ' Now determine the frequency range used for the simulation
	Dim dFmin As Double, dFmax As Double
    dFmin = Solver.GetFmin()
	dFmax = Solver.GetFmax()
    Dim nSamples As Long
	nSamples = 1001
	
	Dim SinpC as Object
	Dim SoutpC as Object
	Set SinpC = Result1DComplex("")
	Set SoutpC = Result1DComplex("")
	
	SinpC.Initialize(nn)
	SoutpC.Initialize(nn)
	
	For ii=0 To nn-1
		SinpC.SetX(ii, Sinp.GetX(ii))
		SinpC.SetYRe(ii, Sinp.GetY(ii))
		SoutpC.SetX(ii, Soutp.GetX(ii))
		SoutpC.SetYRe(ii, Soutp.GetY(ii))
	Next ii
	
	' Calculate the input and output spectrums by using DFT's
	SetIntegrationMethod "trapezoidal"
	CalculateFourierComplex(SinpC,  "time", SInpC, "frequency", "-1", "1.0", dFmin, dFmax,nSamples)
	CalculateFourierComplex(SoutpC, "time", SOutpC, "frequency", "-1", "1.0", dFmin, dFmax,nSamples)

	' Divide the output spectrum by the input spectrum in order to get the sparameters
	SoutpC.ComponentDiv(SinpC)
	SoutpC.SetLogarithmicFactor(20.0)
	SoutpC.SetXLabelAndUnit( "Frequency" ,  Units.GetUnit("Frequency"))
	SoutpC.YLabel ""
	SoutpC.Title "S-Parameters with Timegating" 
	
	SoutpC.Save("^Sc" + sOutput + "_gating.sig")
					
	SoutpC.AddToTree("1D Results\Time Gating\S-Parameters\S" + sOutput)
	
	SelectTreeItem "1D Results\Time Gating\Port signals"

End Sub
'--------------------------------------------------------------------------------------------
Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "Help"
			StartHelp HelpFileName
			DialogFunc = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
Sub Enlarge_Mode_List (Modelist() As String, portindex As String)
Dim cst_runindex As Integer
Dim nr_of_modes_found As Integer
Dim mode_type As String
Dim maxmodenr As Integer
Dim maxmodeatport As Integer
Dim Number_of_Modes As Integer
Dim len_of_Field As Integer
Dim run_index As Integer
Dim mode_impedance As Double

len_of_Field = 0
maxmodeatport=100

'-------------------------------------------------------------------------------
' Waveguide-Port Check
'-------------------------------------------------------------------------------

For Number_of_Modes = 1 To maxmodeatport		'loop over all Modes
  mode_type = ""
  On Error Resume Next
  mode_type = CStr(Port.GetModeType(portindex, Number_of_Modes))
  On Error GoTo 0
  If (mode_type = "") Then Exit For
  len_of_Field = len_of_Field+1
'Next Number_of_Modes
ReDim  Preserve Modelist (len_of_Field-1)
For  run_index = 0 To len_of_Field-1     'Set Portnumbers consequ. from 1,2,etc.
   Modelist (run_index) = CStr(run_index+1)
Next run_index

Next Number_of_Modes

' end waveguide-ports

'--------------------------------------------------------------------------------------
' Discrete-Port Check
'-------------------------------------------------------------------------------
   On Error Resume Next
      mode_impedance = Port.GetWaveImpedance(portindex, 1)
      On Error GoTo 0
      If mode_impedance <> 0 Then
       mode_impedance=0 '´Reset, remain On old value If Error
      GoTo nextloopindex
      End If
      len_of_Field = len_of_Field+1
      ReDim Preserve Modelist(len_of_Field-1)
      For  run_index = 0 To len_of_Field-1
        Modelist (run_index) = CStr(run_index+1)
      Next run_index


     nextloopindex:

   ' end discrete-ports

End Sub
