'#Language "WWB-COM"

' *Matching Circuite / Mini Match advanced
' !!!
' 
'--------------------------------------------------------------------------------------------
' Matching Macro
' ================================================================================================
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 01-Sep-2021 iqa: Replaced deprecated 'ExternalPort.Number'-commands with the 'ExternalPort.Name'-command
' 31-Aug-2021 tsz: Added nextUniquePortName() Function
' 09-Feb-2021 tsz: Replaced port_name with pin_name. Temporary fix until buscalls containing VARIANTS can be transfered from MWS to DS (then replace with Net-command)
' 24-Nov-2018 fde: Fixed Problems with wrong units in Capacitor
' 16-Feb-2016 fde: Fixed problem with crash when dialog box for file selection was exited.
' 22-Nov-2016 fde: Fixed problem with Broadband File
' 16-Nov-2016 fde: Change For 2017, sig-files no longer written
' 15-Oct-2015 gba: do not assume that schematic block is named "MWSSCHEM1"
' 15-Oct-2015 gba: speed up routing by calling port.position and port.rotate before port.create
' 02-May-2014 ube: button cosmetic
' 31-Jan-2013 fde: button cosmetic
' 26-Oct-2012 fde: fixed bug in Topology 3
' 17-Jul-2012 ube: small dialog cosmetics
' 19-Jun-2012 ube: small dialog cosmetics
' 08-Jun-2012 fde: Fixed Problem with center frequncy
' 17-May-2012 fde: First Version
'--------------------------------------------------------------------------------------------


Option Explicit

'#include "vba_globals_all.lib"
'#include "template_conversions.lib"
'#include "infix_postfix.lib"
'#include "complex.lib"


Public p_ar_on As Boolean
Public p_mode As String
Public p_port As String
Public portnamearray() As String
Public nports As Long
Public macropath As String
Public resultdir As String
Public cst_filename As String
Public ModeType As String
Public Line_Impedance As Double
Public Line_Impedance_s As String
Public ListArray_mode() As String
Public Allow_match As Boolean
Public Relative_values As Boolean
Public UseARFilter As Boolean
Public TopologyArray() As Integer
Public BandwidthArray() As Double
Public SerialArray() As Double
Public ParalleArray() As Double
Public TopologyProper() As Boolean
Public cst_freq As Double
Public SchematicBlockName As String

Private Sub GetSchematicBlockName()
	Dim nBlocks As Long
	Dim nIndex As Long
	Dim BlockNameArray() As String
	Dim BlockName As String
	Dim BlockType As String

	SchematicBlockName = "MWSSCHEM1"
	nBlocks = Block.StartBlockNameIteration
	ReDim BlockNameArray(nBlocks)

	For nIndex=0 To nBlocks-1
		BlockNameArray(nIndex) = Block.GetNextBlockName
	Next nIndex

	For nIndex=0 To nBlocks-1
		With Block
			.Reset
			.Name BlockNameArray(nIndex)
			BlockType = .GetTypeShortName
		End With
		If Right$(BlockType, 5) = "SCHEM" Then
			SchematicBlockName = BlockNameArray(nIndex)
		End If
	Next nIndex
End Sub



Sub Main

 Dim ListArray_port() As String
 Dim cst_runindex As Integer
 Dim nports As Integer
 Dim nmodes As Integer
 resultdir     = GetProjectPath ("Result")
 GetSchematicBlockName()


macropath = GetInstallPath + "\Library\Macros"
Dim ModeType As String
Dim Line_Impedance As String


 FillPortNameArray		'get all port names
 ' get the modes of the first available port as default in the popup
 enlarge_mode_list ListArray_mode(),portnamearray(0)
Solver.CalculateZandYMatrices


	Begin Dialog UserDialog 900,245,"Calculate LC-Matching Circuit",.DialogFunc ' %GRID:10,7,1,1
		Text 30,14,50,14,"Port",.Text1
		Text 120,14,60,14,"Mode",.Text3
		DropListBox 30,28,60,192,portnamearray(),.from_port
		DropListBox 120,28,60,192,ListArray_mode(),.from_mode
		PushButton 30,189,250,21,"Place LC Elements in Schematic",.baby_doit
        'OKButton 310,147,110,21
		CancelButton 290,189,90,21
		Text 30,56,150,14,"Port Impedance",.Text4
		TextBox 30,77,130,21,.Lineimpedance
		CheckBox 30,161,180,14,"Use AR Filter if Present",.UseAR
		OptionGroup .Match_choice
			OptionButton 220,28,140,14,"Match Option 1",.OptionButton1
			OptionButton 220,56,140,14,"Match Option 2",.OptionButton2
			OptionButton 220,84,140,14,"Match Option 3",.OptionButton3
			OptionButton 220,112,140,14,"Match Option 4",.OptionButton4
		Text 30,112,150,14,"Matching Frequency",.Text5
		TextBox 30,133,130,21,.freq
		Picture 380,21,320,147,"Picture1",0,.Picture1
		Text 730,84,30,14,"C1",.MatchPart1
		Text 730,28,30,14,"L1",.MatchPart2
		TextBox 730,49,90,21,.ValueTextPart2
		TextBox 730,105,90,21,.ValueTextPart1
		Text 830,49,60,14,"micro H",.LableMatchPart2
		Text 830,105,50,14,"pico F",.LableMatchPart1
		TextBox 730,161,90,21,.BW
		Text 730,140,140,14,"Bandwidth",.Text2
		Text 380,175,330,14,"(Match Option 1 has largest Bandwidth Potential)",.Text6
		PushButton 30,217,250,21,"Write Broadband Match File",.baby_doit2

	End Dialog

 Dim dlg As UserDialog



 Do
  If Dialog(dlg)=0   Then Exit All
 Loop Until  (ListArray_port(0) <> "")

End Sub



'--------------------------------------
Function DialogFunc%(DlgItem As String, Action As Integer, SuppValue As Integer)

    Dim m As String
    Dim file As String
    Dim basepath As String
    Dim cst_freq_s As String


    Debug.Print "Action=";Action
    Debug.Print DlgItem
    Debug.Print "SuppValue=";SuppValue


    Select Case Action
    Case 1 ' Dialog box initialization
            'Print First Impedance

            Relative_values = False
            p_port = DlgText ("from_port")
            p_mode = DlgText ("from_mode")

            DlgText ("freq"), Format ((Solver.GetFMin+(Solver.GetFmax-Solver.GetFMin)/2),"Fixed")
            If (Port.GetType(p_port) = "Waveguide") Then
	            If p_mode = "" Then p_mode = "1"
	            ' Make sure that the port mode has been calculated
	            If Not(SelectTreeItem("2D/3D Results\Port Modes\Port"+CStr(p_port)+"\e"+Cstr(p_mode))) Then
	            	MsgBox("Cannot find selected port mode. Please ensure that port modes have been calculated.")
	            	Exit All
	            End If
	            ModeType = CStr(Port.GetModeType(p_port, p_mode))
	            If (ModeType = "TEM" Or ModeType = "QTEM") Then
		            Line_Impedance = Port.GetLineImpedance (CInt(p_port), CInt(p_mode))
		            DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
		            Calculate_Potential
		            SetDialogBox
	            Else
	            	MsgBox("Selected mode is of type " + ModeType + ". Matching is only possible for TEM or QTEM modes.", "Error")
		            DlgText "Lineimpedance" , "Not Def."
		            Allow_match = False
	            End If
            End If
            If (Port.GetType(p_port) = "Discrete") Then
	            Line_Impedance = Port.GetLineImpedance (CInt(p_port), 1)
	            DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
	            Calculate_Potential
	            SetDialogBox
            End If

        'Beep

    Case 2 ' Value changing or button pressed
      Select Case DlgItem
		Case "Help"
			StartHelp "common_preloadedmacro_filter_analysis_group_delay_computation"
			DialogFunc = True

        Case "from_port" 	'
        	enlarge_mode_list ListArray_mode(), DlgText ("from_port")
            DlgListBoxArray "from_mode", ListArray_mode()
            DlgValue "from_mode", 0
            p_port = DlgText ("from_port")
            p_mode = DlgText ("from_mode")
            If (Port.GetType(p_port) = "Waveguide") Then
            	If p_mode = "" Then p_mode = "1"
				' Make sure that the port mode has been calculated
	            If Not(SelectTreeItem("2D/3D Results\Port Modes\Port"+CStr(p_port)+"\e"+Cstr(p_mode))) Then
	            	MsgBox("Cannot find selected port mode. Please ensure that port modes have been calculated.")
	            	Exit All
	            End If
	            ModeType = CStr(Port.GetModeType(p_port, p_mode))
	            If (ModeType = "TEM" Or ModeType = "QTEM") Then
		            Line_Impedance = Port.GetLineImpedance (CInt(p_port), CInt(p_mode))
		            DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
		            Calculate_Potential
		            SetDialogBox
	            Else
	            	MsgBox("Selected mode is of type " + ModeType + ". Matching is only possible for TEM or QTEM modes.", "Error")
		            DlgText "Lineimpedance" , "Not Def."
			        Allow_match = False
	            End If
            End If
            If (Port.GetType(p_port) = "Discrete") Then
	            Line_Impedance = Port.GetLineImpedance (CInt(p_port), 1)
	            DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
	            Calculate_Potential
	            SetDialogBox
            End If


            DialogFunc = True

        Case "from_mode"
        	p_port = DlgText ("from_port")
            p_mode = DlgText ("from_mode")
			If (Port.GetType(p_port) = "Waveguide") Then
            	If p_mode = "" Then p_mode = "1"
				' Make sure that the port mode has been calculated
	            If Not(SelectTreeItem("2D/3D Results\Port Modes\Port"+CStr(p_port)+"\e"+Cstr(p_mode))) Then
	            	MsgBox("Cannot find selected port mode. Please ensure that port modes have been calculated.")
	            	Exit All
	            End If
	            ModeType = CStr(Port.GetModeType(p_port, p_mode))
	            If (ModeType = "TEM" Or ModeType = "QTEM") Then
		            Line_Impedance = Port.GetLineImpedance (CInt(p_port), CInt(p_mode))
		            DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
		            Calculate_Potential
		            SetDialogBox
	            Else
	            	MsgBox("Selected mode is of type " + ModeType + ". Matching is only possible for TEM or QTEM modes.", "Error")
		            DlgText "Lineimpedance" , "Not Def."
			        Allow_match = False
	            End If
            End If
            If (Port.GetType(p_port) = "Discrete") Then
            	Line_Impedance = Port.GetLineImpedance (CInt(p_port), 1)
            	DlgText "Lineimpedance" , Format (Line_Impedance,"Fixed")
            	Calculate_Potential
            	SetDialogBox
            End If


            DialogFunc = True

        Case "RelativeValues"
        	Relative_values = SuppValue
        	DialogFunc = True

        Case "UseAR"
        	UseARFilter = SuppValue
        	DialogFunc = True

        Case "baby_doit"
            ' update global values
            p_port = DlgText ("from_port")
            p_mode = DlgText ("from_mode")
            ' update Dialog
            DlgText "from_port" ,p_port
            Create_Circuit
            SetDialogBox

        Case "baby_doit2"
        	' update global values
            p_port = DlgText ("from_port")
            p_mode = DlgText ("from_mode")
            ' update Dialog
            DlgText "from_port" ,p_port
            WriteZFile
            SetDialogBox
            DialogFunc=True

        Case "Match_choice"
        	SetDialogBox

            DialogFunc=False

        Case Else

      End Select
    Case 3   ' Text box changed
      		Select Case DlgItem
        		Case "Lineimpedance"
                Line_Impedance_s = DlgText ("Lineimpedance")
                If (CDBl("4,4")= 44) Then
                Replace(Line_Impedance_s,",",".")
      			End If
      			If (CDBl("4.4") = 44) Then
                Replace(Line_Impedance_s,".",",")
      			End If
				Line_Impedance = CDBL (Line_Impedance_s)
				DlgText ("Lineimpedance"), Format (Line_Impedance,"Fixed")
                If (Val (Line_Impedance_s) = 0) Then
               	Allow_match = False
               	Else
               	Calculate_Potential
               	SetDialogBox
               	End If
        		DlgText ("Lineimpedance"), Format (Line_Impedance_s,"Fixed")
        	    Case "freq"
      			cst_freq_s = DlgText ("freq")
      			If (CDBl("4,4")= 44) Then
                Replace(cst_freq_s,",",".")
      			End If
      			If (CDBl("4.4") = 44) Then
                Replace(cst_freq_s,".",",")
      			End If
      			cst_freq = CDBL(cst_freq_s)
				DlgText ("freq"), cstr(cst_freq)
                Calculate_Potential
                SetDialogBox


      		End Select

      DialogFunc = True 				'do not exit the dialog
	  'm = DlgText ("ComboBox1")

    Case 4 ' Focus changed
       Debug.Print "DlgFocus=""";DlgFocus();""""
    End Select
End Function

 '-----------------------
Sub SetDialogBox
	Dim Matchchoice As Integer
	Dim Topology As Integer
	Dim Ps As Double
	Dim Pp As Double
	Dim Bw As Double
	Dim cst_filenam As String


    Matchchoice = DlgValue("Match_choice")
    Topology = TopologyArray(Matchchoice)
    Ps = SerialArray (Matchchoice)
    Pp = ParalleArray (Matchchoice)
    Bw = BandwidthArray (Matchchoice)/cst_freq

	Select Case Topology
	Case 1 'Cs,Lp
		DlgText ("MatchPart1", "C")
		DlgText ("LableMatchPart1", "pico F")
		DlgText ("ValueTextPart1", Format (Ps/1e-12, "Scientific"))

		DlgText ("MatchPart2", "L")
		DlgText ("LableMatchPart2", "micro H")
		DlgText ("ValueTextPart2", Format (Pp/1e-6, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match1.bmp"
        DlgSetPicture "Picture1",cst_filename,0

	Case 2 'Ls,Cp
		DlgText ("MatchPart1", "L")
		DlgText ("ValueTextPart1", Format (Ps/1e-6, "Scientific"))
		DlgText ("LableMatchPart1", "micro H")

		DlgText ("MatchPart2", "C")
		DlgText ("LableMatchPart2", "pico F")
		DlgText ("ValueTextPart2", Format (Pp/1e-12, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match2.bmp"
        DlgSetPicture "Picture1",cst_filename,0

	Case 3 'Cs,Cp
		DlgText ("MatchPart1", "Cs")
		DlgText ("LableMatchPart1", "pico F")
		DlgText ("ValueTextPart1", Format (Ps/1e-12, "Scientific"))

		DlgText ("MatchPart2", "Cp")
		DlgText ("LableMatchPart2", "pico F")
		DlgText ("ValueTextPart2", Format (Pp/1e-12, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match3.bmp"
        DlgSetPicture "Picture1",cst_filename,0

	Case 4 'Ls,Lp
		DlgText ("MatchPart1", "Ls")
		DlgText ("LableMatchPart1", "micro H")
		DlgText ("ValueTextPart1", Format (Ps/1e-6, "Scientific"))

		DlgText ("MatchPart2", "Lp")
		DlgText ("LableMatchPart2", "micro H")
		DlgText ("ValueTextPart2", Format (Pp/1e-6, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match4.bmp"
        DlgSetPicture "Picture1",cst_filename,0


	Case 5 'Cp,Ls
		DlgText ("MatchPart1", "C")
		DlgText ("LableMatchPart1", "pico F")
		DlgText ("ValueTextPart1", Format (Pp/1e-12, "Scientific"))


		DlgText ("MatchPart2", "L")
		DlgText ("LableMatchPart2", "micro H")
		DlgText ("ValueTextPart2", Format (Ps/1e-6, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match5.bmp"
        DlgSetPicture "Picture1",cst_filename,0


	Case 6 'Lp,Cs
		DlgText("MatchPart1", "L")
		DlgText ("LableMatchPart1", "micro H")
		DlgText ("ValueTextPart1", Format (Pp/1e-6, "Scientific"))


		DlgText ("MatchPart2", "C")
		DlgText ("LableMatchPart2", "pico F")
		DlgText ("ValueTextPart2", Format (Ps/1e-12, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match6.bmp"
        DlgSetPicture "Picture1",cst_filename,0


	Case 7
		DlgText ("MatchPart1", "Cp")
		DlgText ("LableMatchPart1", "pico F")
		DlgText ("ValueTextPart1", Format (Pp/1e-12, "Scientific"))

		DlgText ("MatchPart2", "Cs")
		DlgText ("LableMatchPart2", "pico F")
		DlgText ("ValueTextPart2", Format (Ps/1e-12, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match7.bmp"
        DlgSetPicture "Picture1",cst_filename,0


	Case 8
		DlgText("MatchPart1", "Lp")
		DlgText ("LableMatchPart1", "micro H")
		DlgText ("ValueTextPart1", Format (Pp/1e-6, "Scientific"))

		DlgText ("MatchPart2", "Ls")
		DlgText ("LableMatchPart2", "micro H")
		DlgText ("ValueTextPart2", Format (Ps/1e-6, "Scientific"))
		cst_filename = macropath+"\Matching Circuits\Match8.bmp"
        DlgSetPicture "Picture1",cst_filename,0

	End Select
	DlgText ("Bw", Format (Bw, "Percent"))

 End Sub



 '-----------------------

Private Sub Calculate_Potential

Dim one As Complex

Dim Z_ant As Object
Dim Z_re As Object
Dim X_ant As Double
Dim R_ant As Double
Dim sFile As String
Dim sR1D As String

Dim X_match_neg As Double
Dim X_match_pos As Double
Dim B_match_neg As Double
Dim B_match_pos As Double
Dim Z_port As Double
Dim Z_port_comp As Complex
Dim L1_match As Double
Dim C1_match As Double
Dim L2_match As Double
Dim C2_match As Double
Dim Z_net As Complex
Dim X_net As Complex
Dim B_net As Complex
Dim cst_iii As Long
Dim cst_i As Long
Dim freq As Double
Dim cst_low_match As Long
Dim cst_high_match As Long
Dim cst_sym_match As Long
Dim reflection As Complex
Dim reflection_den As Complex
Dim reflection_num As Complex
Dim reflection_abs As Double
Dim calculation_works As Boolean
Dim bandwidth_potential_s As Double
Dim bandwidth_potential_sym_s As Double
Dim Matching_topology_s As Integer
Dim Matching_topology_sym_s As Integer
Dim parallel_l_s As Double
Dim parallel_l_sym_s As Double
Dim parallel_c_s As Double
Dim parallel_c_sym_s As Double
Dim serial_l_s As Double
Dim serial_l_sym_s As Double
Dim serial_c_s As Double
Dim serial_c_sym_s As Double
Dim second_matching_elements_s As Double
Dim second_matching_elements_sym_s As Double
Dim bandwidth_potential_best_so_far As Double
Dim bandwidth_potential_sym_best_so_far As Double
Dim Matching_topology_best_so_far As Integer
Dim Matching_topology_sym_best_so_far As Integer
Dim bandwidth_potential As Object
Dim bandwidth_potential_sym As Object
Dim Matching_topology As Object
Dim Matching_topology_sym As Object
Dim parallel_l As Object
Dim parallel_c As Object
Dim serial_l As Object
Dim serial_c As Object
Dim parallel_l_sym As Object
Dim parallel_c_sym As Object
Dim serial_l_sym As Object
Dim serial_c_sym As Object
Dim cst_freq_no As String
Dim Topology_count As Integer


Line_Impedance = 0

p_port = DlgText ("from_port")
p_mode = DlgText ("from_mode")
Line_Impedance_s = DlgText ("LineImpedance")
Line_Impedance = CDbl (Line_Impedance_s)

Dim ZOutIn As String
	' Determine prefix for Z matrix file
    	ZOutIn =  p_port + "(" + p_mode + ")" + p_port+ "(" + p_mode + ")"
	    ZOutIn = Replace(ZOutIn, "()", "")
	    ' Load m Z matrix file
	    sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix\Z" + ZOutIn)
	    If sFile = "" Then 'try without mode numbers
        ZOutIn =  p_port + "," + p_port
        sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix\Z" + ZOutIn)
        End If
        If (sFile = "")  Then
	    	MsgBox("Cannot find Z matrix results. Please make sure that the Z matrix has been calculated.","Error")
            Exit All
        Else
        	Set Z_ant = Result1DComplex(sFile)
        End If

		' Load Zre and Zim from Z matrix file - AR version
	    If UseARFilter Then
	    	sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix (AR)\Z" + ZOutIn)
	    	If sFile = "" Then 'try without mode numbers
        		ZOutIn =  p_port + "," + p_port
        		sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix (AR)\Z" + ZOutIn)
            End If
        	If (sFile <> "")  Then
        		Set Z_ant = Result1DComplex(sFile)
        	End If
	    End If



one.re = 1
one.im = 0
Z_port = Line_Impedance
Z_port_comp.re = Z_port
Z_port_comp.im = 0


'Main Loop

cst_freq_no = DlgText ("freq")
cst_freq = (RealVal_old(cst_freq_no))


'Get match for serial/parallel (maximum 2 Sets) Possible topologies: 1 CL 2 LC 3 CC 4 LL

calculation_works = True

bandwidth_potential_s = 0
bandwidth_potential_best_so_far = 0
Matching_topology_s = 0
Matching_topology_best_so_far = 0

bandwidth_potential_sym_s = 0
bandwidth_potential_sym_best_so_far = 0
Matching_topology_sym_s = 0
Matching_topology_sym_best_so_far = 0
Topology_count = 0
Set Z_re = Z_ant.Real
cst_i=Z_re.GetClosestIndexFromX(cst_freq)

R_ant=Z_ant.GetYRe(cst_i)
X_ant=Z_ant.GetYIm(cst_i)

If (R_ant^2+X_ant^2 >= Z_port*R_ant) Then  ' Check if match for serial/parallel is possible: Possible topologies:

B_match_pos=(X_ant+Sqr(R_ant/Z_port)*Sqr(R_ant^2+X_ant^2-Z_port*R_ant))/(R_ant^2+X_ant^2)
B_match_neg=(X_ant-Sqr(R_ant/Z_port)*Sqr(R_ant^2+X_ant^2-Z_port*R_ant))/(R_ant^2+X_ant^2)
X_match_pos = 1/B_match_pos + X_ant*Z_port/R_ant - Z_port/(B_match_pos*R_ant)
X_match_neg= 1/B_match_neg + X_ant*Z_port/R_ant - Z_port/(B_match_neg*R_ant)


'Calculate First Set------------------------------------------------------------------------
ReDim Preserve TopologyArray(Topology_count+1)
ReDim Preserve BandwidthArray(Topology_count+1)
ReDim Preserve SerialArray(Topology_count+1)
ReDim Preserve ParalleArray(Topology_count+1)
ReDim Preserve TopologyProper(Topology_count+1)


L1_match = 0
C1_match = 0
L2_match = 0
C2_match = 0

If B_match_pos < 0 Then 'Inductor
L1_match = -1/(B_match_pos*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
ParalleArray(Topology_count) = L1_match
Else 'cap
C1_match = B_match_pos/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
ParalleArray(Topology_count) = C1_match
End If
If X_match_pos < 0 Then 'Capacitor
C2_match = -1/(X_match_pos*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
SerialArray(Topology_count) = C2_match
Else
L2_match = X_match_pos/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
SerialArray(Topology_count) = L2_match
End If


' Just for topology check
If L1_match*L2_match > 0 Then Matching_topology_s = 4
If C1_match*C2_match > 0 Then Matching_topology_s = 3
If C1_match*L2_match > 0 Then Matching_topology_s = 2
If C2_match*L1_match > 0 Then Matching_topology_s = 1

TopologyArray(Topology_count) = Matching_topology_s



'Calculate Reflection and lower bandwidth limit cst_low_match

cst_low_match = cst_i

While ((cst_low_match > 1) And (reflection_abs < 0.501))
cst_low_match = cst_low_match -1
freq = Z_ant.GetX(cst_low_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_low_match)
Z_net.im = Z_ant.GetYIm (cst_low_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If
 'Calculate Reflection
 'Impedance first
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)
  Z_net = plus (X_net, Z_net)
  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend

reflection_abs = 0 'reset
'Calculate Reflection and higher bandwidth limit cst_high_match

cst_high_match = cst_i -1
reflection_abs = 0
While ((cst_high_match < Z_ant.GetN-1 ) And (reflection_abs < 0.501))
cst_high_match = cst_high_match + 1
freq = Z_ant.GetX(cst_high_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_high_match)
Z_net.im = Z_ant.GetYIm (cst_high_match)
B_net.re = 0
X_net.re = 0

If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If
 'Calculate Reflection
 'Impedance first
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)
  Z_net = plus (X_net, Z_net)
  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend

reflection_abs = 0 'reset
If ((cst_low_match = 0) Or (cst_high_match = Z_ant.GetN-1))  Then
TopologyProper(Topology_count) = False 'Matching works but BW Potential is not proper..
Else
TopologyProper(Topology_count) = True
End If

BandwidthArray(Topology_count) = Z_ant.Getx(cst_high_match)-Z_ant.Getx(cst_low_match)

Topology_count = Topology_count + 1

'Calculate second set ------------------------------------------------------------------------

L1_match = 0
C1_match = 0
L2_match = 0
C2_match = 0

ReDim Preserve TopologyArray(Topology_count+1)
ReDim Preserve BandwidthArray(Topology_count+1)
ReDim Preserve SerialArray(Topology_count+1)
ReDim Preserve ParalleArray(Topology_count+1)
ReDim Preserve TopologyProper(Topology_count+1)


If B_match_neg < 0 Then 'Inductor
L1_match = -1/(B_match_neg*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
ParalleArray(Topology_count) = L1_match
Else 'cap
C1_match = B_match_neg/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
ParalleArray(Topology_count) = C1_match
End If
If X_match_neg < 0 Then 'Capacitor
C2_match = -1/(X_match_neg*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
SerialArray(Topology_count) = C2_match
Else
L2_match = X_match_neg/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)
SerialArray(Topology_count) = L2_match
End If

' Just for topology check
If L1_match*L2_match > 0 Then Matching_topology_s = 4
If C1_match*C2_match > 0 Then Matching_topology_s = 3
If C1_match*L2_match > 0 Then Matching_topology_s = 2
If C2_match*L1_match > 0 Then Matching_topology_s = 1

TopologyArray(Topology_count) = Matching_topology_s

'Calculate Reflection and lower bandwidth limit cst_low_match

cst_low_match = cst_i+1

While ((cst_low_match > 1) And (reflection_abs < 0.501))
cst_low_match = cst_low_match -1
freq = Z_ant.GetX(cst_low_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_low_match)
Z_net.im = Z_ant.GetYIm (cst_low_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If
 'Calculate Reflection
 'Impedance first
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)
  Z_net = plus (X_net, Z_net)
  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend

reflection_abs = 0 'reset
'Calculate Reflection and higher bandwidth limit cst_high_match

cst_high_match = cst_i

While ((cst_high_match < Z_ant.GetN-1 ) And (reflection_abs < 0.501))
cst_high_match = cst_high_match + 1
freq = Z_ant.GetX(cst_high_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_high_match)
Z_net.im = Z_ant.GetYIm (cst_high_match)
B_net.re = 0
X_net.re = 0

If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If
 'Calculate Reflection
 'Impedance first
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)
  Z_net = plus (X_net, Z_net)
  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend

reflection_abs = 0 'reset
'MsgBox "exit first with: " + Cstr(cst_i) + "    "+ Cstr(cst_low_match) + "   " + (cstr(cst_high_match))  + "    " + _
'CStr(L1_match) + "   " + CStr(L2_match) + "   " + CStr(C1_match) + "   " + CStr(C2_match)

If ((cst_low_match = 0) Or (cst_high_match = Z_ant.GetN-1))  Then
TopologyProper(Topology_count) = False 'Matching works but BW Potential is not proper..
Else
TopologyProper(Topology_count) = True
End If

BandwidthArray(Topology_count) = Z_ant.Getx(cst_high_match)-Z_ant.Getx(cst_low_match)

Topology_count = Topology_count + 1
End If ' for serial parallel


'------------------------------------------------------
'------------------------------------------------------
'------------------------------------------------------

If (R_ant <= Z_port) Then  ' Check if match for parallel/serial is possible
' Get match for serial/parallel (maximum 2 Sets)
' Possible topologies: 5 CL 6 LC 7 CC 8 LL


ReDim Preserve TopologyArray(Topology_count+1)
ReDim Preserve BandwidthArray(Topology_count+1)
ReDim Preserve SerialArray(Topology_count+1)
ReDim Preserve ParalleArray(Topology_count+1)
ReDim Preserve TopologyProper(Topology_count+1)


B_match_pos = Sqr((Z_port-R_ant)/R_ant)/Z_port
B_match_neg = -Sqr((Z_port-R_ant)/R_ant)/Z_port
X_match_pos = Sqr(R_ant*(Z_port-R_ant))-X_ant
X_match_neg= -Sqr(R_ant*(Z_port-R_ant))-X_ant


'Calculate First Set------------------------------------------------------------------------

L1_match = 0
C1_match = 0
L2_match = 0
C2_match = 0


If B_match_pos < 0 Then 'Inductor
L1_match = -1/(B_match_pos*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI) 'Y
ParalleArray(Topology_count) = L1_match
Else 'cap
C1_match = B_match_pos/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Y
ParalleArray(Topology_count) = C1_match
End If
If X_match_pos < 0 Then 'Capacitor
C2_match = -1/(X_match_pos*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Z
SerialArray(Topology_count) = C2_match
Else 'inductor
L2_match = X_match_pos/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Z
SerialArray(Topology_count) = L2_match
End If

' Just for topology check
If L1_match*L2_match > 0 Then Matching_topology_s = 8
If C1_match*C2_match > 0 Then Matching_topology_s = 7
If C1_match*L2_match > 0 Then Matching_topology_s = 5
If C2_match*L1_match > 0 Then Matching_topology_s = 6

TopologyArray(Topology_count) = Matching_topology_s

'Calculate Reflection and lower bandwidth limit cst_low_match

cst_low_match = cst_i+1
reflection_abs  = 0

While ((cst_low_match > 1) And (reflection_abs < 0.501))
cst_low_match = cst_low_match -1
freq = Z_ant.GetX(cst_low_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_low_match)
Z_net.im = Z_ant.GetYIm (cst_low_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If

 'Calculate Reflection

 'Impedance first

  Z_net.im = X_net.im + Z_net.im
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)

  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend
reflection_abs = 0 'reset
' MsgBox "End low loop 1"


'Calculate Reflection and higher bandwidth limit cst_high_match

cst_high_match = cst_i-1

While ((cst_high_match < Z_ant.GetN-1 ) And (reflection_abs < 0.501))
cst_high_match = cst_high_match + 1
freq = Z_ant.GetX(cst_high_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_high_match)
Z_net.im = Z_ant.GetYIm (cst_high_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If

 'Calculate Reflection

 'Impedance first

  Z_net.im = X_net.im + Z_net.im
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)

  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)

Wend
reflection_abs = 0 'reset


If ((cst_low_match = 0) Or (cst_high_match = Z_ant.GetN-1))  Then
TopologyProper(Topology_count) = False 'Matching works but BW Potential is not proper..
Else
TopologyProper(Topology_count) = True
End If

BandwidthArray(Topology_count) = Z_ant.Getx(cst_high_match)-Z_ant.Getx(cst_low_match)


Topology_count = Topology_count + 1

'Calculate second set------------------------------------------------------------------------


ReDim Preserve TopologyArray(Topology_count+1)
ReDim Preserve BandwidthArray(Topology_count+1)
ReDim Preserve SerialArray(Topology_count+1)
ReDim Preserve ParalleArray(Topology_count+1)
ReDim Preserve TopologyProper(Topology_count+1)


L1_match = 0
C1_match = 0
L2_match = 0
C2_match = 0


If B_match_neg < 0 Then 'Inductor
L1_match = -1/(B_match_neg*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI) 'Y
ParalleArray(Topology_count) = L1_match
Else 'cap
C1_match = B_match_neg/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Y
ParalleArray(Topology_count) = C1_match
End If
If X_match_neg < 0 Then 'Capacitor
C2_match = -1/(X_match_neg*2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Z
SerialArray(Topology_count) = C2_match
Else 'inductor
L2_match = X_match_neg/(2*Pi*Z_ant.Getx(cst_i)*Units.GetFrequencyUnitToSI)  'Z
SerialArray(Topology_count) = L2_match
End If

' Just for topology check

If (L1_match*L2_match > 0) Then Matching_topology_s = 8
If (C1_match*C2_match > 0) Then Matching_topology_s = 7
If (C1_match*L2_match > 0) Then Matching_topology_s = 5
If (C2_match*L1_match > 0) Then Matching_topology_s = 6

TopologyArray(Topology_count) = Matching_topology_s

'Calculate Reflection and lower bandwidth limit cst_low_match

cst_low_match = cst_i+1
reflection_abs = 0 'reset
While ((cst_low_match > 1) And (reflection_abs < 0.501))
cst_low_match = cst_low_match -1
freq = Z_ant.GetX(cst_low_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_low_match)
Z_net.im = Z_ant.GetYIm (cst_low_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If

 'Calculate Reflection

 'Impedance first

  Z_net.im = X_net.im + Z_net.im
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)
  Z_net.re = Z_net.re
  Z_net.im = Z_net.im

  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)
Wend


reflection_abs = 0 'reset
'Calculate Reflection and higher bandwidth limit cst_high_match

cst_high_match = cst_i -1

While ((cst_high_match < Z_ant.GetN-1 ) And (reflection_abs < 0.501))
cst_high_match = cst_high_match + 1
freq = Z_ant.GetX(cst_high_match)*Units.GetFrequencyUnitToSI
Z_net.re = Z_ant.GetYRe (cst_high_match)
Z_net.im = Z_ant.GetYIm (cst_high_match)
B_net.re = 0
X_net.re = 0

  If L1_match <> 0 Then
  	B_net.im = -1/(L1_match*2*Pi*freq) 'Y not Z
  Else
    B_net.im = C1_match*(2*Pi*freq)  'Y not Z
  End If
  If C2_match <> 0 Then
  	X_net.im = -1/(C2_match*2*Pi*freq) ' this is Z
  Else
    X_net.im = L2_match*2*Pi*freq  ' this is Z
  End If

 'Calculate Reflection

 'Impedance first

  Z_net.im = X_net.im + Z_net.im
  Z_net = div (one ,Z_net)
  Z_net = plus (B_net, Z_net)
  Z_net = div (one,Z_net)

  'now refelction
  reflection_num = minus (Z_net,Z_port_comp)
  reflection_den = plus (Z_net, Z_port_comp)
  reflection = div (reflection_num,reflection_den)
  reflection_abs = absolute (reflection)

Wend
reflection_abs = 0 'reset

If ((cst_low_match = 0) Or (cst_high_match = Z_ant.GetN-1))  Then
TopologyProper(Topology_count) = False 'Matching works but BW Potential is not proper..
Else
TopologyProper(Topology_count) = True
End If


BandwidthArray(Topology_count) = Z_ant.Getx(cst_high_match)-Z_ant.Getx(cst_low_match)

Topology_count = Topology_count + 1

End If ' for paralle/serial


' ----------------------------------------------   All set calculated

' Now sort for Bandwidth and Proper


Dim Bandwidth_temp As Double
Dim Serial_temp As Double
Dim Parallel_temp As Double
Dim Proper_temp As Boolean
Dim Togology_temp As Integer
Dim cst_ii As Integer


For cst_i = 1 To 9
For cst_ii = 0 To Topology_count -1

If ((BandwidthArray(cst_ii) < (BandwidthArray(cst_ii+1))) And TopologyProper(cst_ii+1))  Then
    Bandwidth_temp = BandwidthArray(cst_ii)
    BandwidthArray(cst_ii) = BandwidthArray(cst_ii+1)
    BandwidthArray(cst_ii+1)=Bandwidth_temp
    Serial_temp  = SerialArray(cst_ii)
    SerialArray(cst_ii)=SerialArray(cst_ii+1)
    SerialArray(cst_ii+1) = Serial_temp
    Parallel_temp = ParalleArray(cst_ii)
    ParalleArray(cst_ii) = ParalleArray(cst_ii+1)
    ParalleArray(cst_ii+1)=Parallel_temp
    Togology_temp = TopologyArray(cst_ii)
    TopologyArray(cst_ii)=TopologyArray(cst_ii+1)
    TopologyArray(cst_ii+1)=Togology_temp
    Proper_temp = TopologyProper(cst_ii)
    TopologyProper(cst_ii) = TopologyProper(cst_ii+1)
    TopologyProper(cst_ii+1)=Proper_temp
End If

Next
Next
If Topology_count = 2 Then
	DlgEnable ("OptionButton3", False)
	DlgEnable ("OptionButton4", False)
	DlgValue ("Match_choice", 0)
	'MsgBox ( CStr(TopologyArray(0)) + "  " + CStr(TopologyArray(1)))
Else
   	DlgEnable ("OptionButton3", True)
	DlgEnable ("OptionButton4", True)
	'MsgBox ( CStr(TopologyArray(0)) + "  " + CStr(TopologyArray(1))  + "  " + CStr(TopologyArray(2))  + "  " + CStr(TopologyArray(3)))
End If


'---- End Main Loop


End Sub

Private Sub WriteZFile

Dim S_ant As Object
Dim S_re As Object
Dim X_ant As Double
Dim R_ant As Double
Dim sFile As String
Dim sR1D As String

Dim cst_iii As Long
Dim cst_i As Long
Dim freq As Double
Dim calculation_works As Boolean
Dim sOutputFile As String
Dim Zstring As String
Dim cst_freq_no As String
Dim sOutputFormat As String

Line_Impedance = 0

p_port = DlgText ("from_port")
p_mode = DlgText ("from_mode")
Line_Impedance_s = DlgText ("LineImpedance")
Line_Impedance = CDbl (Line_Impedance_s)

sOutputFormat = "00000.0000"

Dim SOutIn As String
	' Determine prefix for Z matrix file
    	SOutIn =  p_port + "(" + p_mode + ")" + p_port+ "(" + p_mode + ")"
	    SOutIn = Replace(SOutIn, "()", "")
	    ' Load m Z matrix file
        sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix\Z" + SOutIn)
	    If sFile = "" Then 'try without mode number
	            SOutIn =  p_port + "," + p_port
	            sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix\Z" + SOutIn)
        End If
       Set S_ant = Result1DComplex(sFile)

		' Load Zre and Zim from Z matrix file - AR version
	    If UseARFilter Then
	    	sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix (AR)\Z" + SOutIn)
	    	If sFile = "" Then 'try without mode numbers
        		SOutIn =  p_port + "," + p_port
        		sFile=Resulttree.GetFileFromTreeItem("1D Results\Z Matrix (AR)\Z" + SOutIn)
            End If
        	If (sFile <> "")  Then
        		Set S_ant = Result1DComplex(sFile)
        	End If
	    End If


cst_iii = S_ant.GetN
sOutputFile = GetFilePath("*.s1p","*",GetProjectPath("Result"),"Select output file",7)
If (sOutputFile = "") Then Exit Sub

Open sOutputFile For Output As #2
'WriteHeader for Z
Print #2, "! TOUCHSTONE file generated by CST STUDIO Suite Macro Mini Match"
Print #2, "! " + Cstr(Date) +  "     " + Cstr(Time)
Print #2, "! Creates Broadband Matching file Ports"
Print #2, "# " + Cstr(Units.GetUnit("Frequency")) + " Z RI R 1"
'Loop
S_ant.Conjugate
For cst_i = 0 To cst_iii-1
  Zstring = USFormat(S_ant.GetX(cst_i),"000000.000000") + "   " + USFormat(S_ant.GetYRe(cst_i),"0.000000") + "  " + USFormat(S_ant.GetYim(cst_i),"0.000000")
  Print #2, Zstring
Next
Close #2

Exit All

End Sub


Public Function nextUniquePortName() As String
	Dim portNumber As Integer
	
	nextUniquePortName = "1"
	portNumber = 1
	
	With ExternalPort
		.Reset
		.Name nextUniquePortName
		While .DoesExist()
			portNumber = portNumber + 1
			.Reset
			.Name CStr(portNumber)
		Wend
	End With
	
	nextUniquePortName = Cstr(portNumber)
End Function

Private Sub Create_Circuit

     'DS

Dim mws_block_port_position_x As  Long, mws_block_port_position_y As  Long
Dim mws_block_center_position_x As Long, mws_block_center_position_y As  Long, make_abs As  String
Dim x_offset As Integer, y_offset As Integer, size_offset As Integer
Dim blockname_variable As String, rel_x_pos As Double, rel_y_pos As Double, rot_angle As Double
Dim differential_flag As Boolean, orientation As String
Dim L_a As Double, glength As Double, iii As Integer, port_name As String, pin_name As String, index_p As Long , iii2 As Integer
Dim np As Integer, strg_pos As Integer, Strg_length As Integer
Dim Next_port_name As String


Dim Matchchoice As Integer
Dim Topology As Integer
Dim Ps As Double
Dim Pp As Double



Matchchoice = DlgValue("Match_choice")
Topology = TopologyArray(Matchchoice)
Ps =SerialArray (Matchchoice)
Pp =ParalleArray (Matchchoice)



differential_flag = False

'Returns a Block's number of ports.

With Block
	.name SchematicBlockName
    np=.GetNumberOfPorts
End With

'Get Port Index and compare Port names with Matching names
For iii = 0 To np
     With Block
     	.name SchematicBlockName
     	port_name =.GetPortName (iii)
     End With
    If Right$(port_name,1)=")" Then		'multipin"
    	strg_pos = InStr(port_name,"(")
    	Strg_length = Len(port_name)
    	If (p_port = Left$(port_name,strg_pos-1)) Then
    		If ((p_mode + ")" = Right$(port_name, (Strg_length-strg_pos)))) Then
                index_p = iii
    		End If
    	End If

    Else
       If port_name = p_port Then
       		index_p = iii
       End If
    End If
Next

With Block
	.name SchematicBlockName
	index_p = .GetPinIndexFromPortIndex(index_p)
	If .GetBusSize(index_p) <> 1 Then
		ReportError "Buspins are not supported."
	End If
	pin_name = .GetPinName (index_p)
End With



With Block
	.name SchematicBlockName
	.SetDifferentialPorts differential_flag
	mws_block_port_position_x = .GetPortPositionX (   (index_p  ) )
	mws_block_port_position_y = .GetPortPositionY (   (index_p ) )
	mws_block_center_position_x = .GetPositionX
	mws_block_center_position_y = .GetPositionY
End With
  rel_x_pos =  mws_block_center_position_x - mws_block_port_position_x
  rel_y_pos =  mws_block_center_position_y - mws_block_port_position_y


' chekc if port is connected

Dim Port_connected

With Block
	.name SchematicBlockName
	Port_connected = .IsPortConnected (index_p)
End With

If Port_connected Then
		MsgBox("Port is already connected - Please de-connect ports -  Exiting Macro", "Error")
		Exit All
End If
size_offset = 200	'position offset for drawing the el ements
  'compute the orientation of the Matching Elemments; either left / right /up /down
  If Abs(rel_x_pos) >= Abs(rel_y_pos) Then 'horizontal orientated
   If rel_x_pos <0 Then
  	x_offset =  size_offset : y_offset =  0 : 	rot_angle = 0 : orientation="R" 'to the right
   Else
	x_offset = - size_offset : y_offset =  0 : 	rot_angle = 180 : orientation="L" 'to the left
   End If
  End If
  If Abs(rel_x_pos) < Abs(rel_y_pos) Then ' vertical orientated
   If rel_y_pos <0 Then
  	x_offset =   0 : y_offset =   size_offset : 	rot_angle = 90 : orientation="D" 	'   down
   Else
	x_offset =   0 : y_offset =  - size_offset : 	rot_angle = -90	: orientation="U" ' up
   End If
  End If


' ---------
If (Topology = 1) Then
   blockname_variable = "C"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Capacitance",  Round(Ps/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
   End With
   blockname_variable = "L"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+90)
    	Case "L"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+270)
    	Case  "D"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset )
            .Rotate (rot_angle+270)
		Case "U"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset)
            .Rotate (rot_angle+90)
    End Select
    .SetDoubleProperty ("Inductance",  Round(Pp/1e-6,6)  )
	.create
    .SetLocalUnitForProperty ("Inductance", "uH")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND" + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
    			Case  "D"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and C
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("C"+ p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between C and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("C"+ p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort("L"+ p_port + "(" + p_mode + ")" ,"1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("C"+ p_port + "(" + p_mode + ")" ,"1",False)
	 .create
    End With
	With Link	'between L and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("L"+ p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")", "GND" ,False)
	 .create
    End With

End If

' ---- Topology 2

If (Topology = 2) Then
   blockname_variable = "L"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Inductance",  Round(Ps/1e-6,6)  )
    .create
    .SetLocalUnitForProperty ("Inductance", "uH")

   End With
   blockname_variable = "C"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+90)
    	Case "L"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+270)
    	Case  "D"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset )
            .Rotate (rot_angle+270)
		Case "U"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset)
            .Rotate (rot_angle+90)
    End Select
    .SetDoubleProperty ("Capacitance",  Round(Pp/1e-12,6) )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND" + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
    			Case  "D"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset   )
   				End Select
			.Create
			End With

    Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and L
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between L and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort ("C" + p_port + "(" + p_mode + ")","1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With
	With Link	'between L and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("C" + p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")","GND",False)
	 .create
    End With

End If


'---- Tolopogy 3 --ext 1 ----------------

If (Topology = 3 ) Then
   blockname_variable = "Cs"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Capacitance",  Round(Ps/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
   End With
   blockname_variable = "Cp"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+90)
    	Case "L"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+270)
    	Case  "D"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset )
            .Rotate (rot_angle+270)
		Case "U"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset)
            .Rotate (rot_angle+90)
    End Select
    .SetDoubleProperty ("Capacitance",  Round(Pp/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND" + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
    			Case  "D"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and C
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("Cs"+ p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between C and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("Cs"+ p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Cp"+ p_port + "(" + p_mode + ")" ,"1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("Cs"+ p_port + "(" + p_mode + ")" ,"1",False)
	 .create
    End With
	With Link	'between L and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Cp"+ p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")", "GND" ,False)
	 .create
    End With

End If


' ------  Topology 4   -----------------

If (Topology = 4) Then
   blockname_variable = "Ls"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Inductance",  Round(Ps/1e-6,6)  )
    .create
    .SetLocalUnitForProperty ("Inductance", "uH")
   End With
   blockname_variable = "Lp"+ p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+90)
    	Case "L"
    		.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+270)
    	Case  "D"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset )
            .Rotate (rot_angle+270)
		Case "U"
			.position (mws_block_port_position_x +0.5*x_offset +size_offset/2, mws_block_port_position_y+0.5*y_offset)
            .Rotate (rot_angle+90)
    End Select
    .SetDoubleProperty ("Inductance",  Round(Pp/1e-6,6)  )
	.create
    .SetLocalUnitForProperty ("Inductance", "uH")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND" + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +0.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
    			Case  "D"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2.2*x_offset +size_offset , mws_block_port_position_y+0.5*y_offset + 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and Ls
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("Ls"+ p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between L and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("Ls"+ p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between Ls and Lp -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Lp"+ p_port + "(" + p_mode + ")" ,"1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("Ls"+ p_port + "(" + p_mode + ")" ,"1",False)
	 .create
    End With
	With Link	'between L and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Lp"+ p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")", "GND" ,False)
	 .create
    End With

End If



'  -----------------   Topology 5   ----------------------


If (Topology = 5) Then
   blockname_variable = "L" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Inductance",  Round(Ps/1e-6,6)  )
    .create
    .SetLocalUnitForProperty ("Inductance", "uH")
   End With
   blockname_variable = "C" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+270)
    	Case "L"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+90)
    	Case "D"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset + size_offset/2  )
    		.Rotate (rot_angle+90)
		Case "U"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset - size_offset/2  )
    		.Rotate (rot_angle+270)
    End Select
    .SetDoubleProperty ("Capacitance",  Round(Pp/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND"  + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
 			    Case  "D"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset + size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset - 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and L
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between L and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort("L" + p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("C" + p_port + "(" + p_mode + ")","2",False)
	 .create
    End With
	With Link	'between C and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("C" + p_port + "(" + p_mode + ")","1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")","GND",False)
	 .create
    End With

End If


' ------------- Topology 6  ---------------------------


If (Topology = 6) Then
   blockname_variable = "C" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Capacitance", Round(Ps/1e-12,6)  )
    .create
   End With
   blockname_variable = "L" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+270)
       	Case "L"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    	    .Rotate (rot_angle+90)
    	Case  "D"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset + size_offset/2  )
            .Rotate (rot_angle+90)
		Case "U"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset - size_offset/2  )
            .Rotate (rot_angle+270)
    End Select

    .SetDoubleProperty ("Inductance",  Round(Pp/1e-6,6)  )
 	.create
    .SetLocalUnitForProperty ("Inductance", "uH")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND"  + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
    			Case  "D"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset + size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset - 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and C
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("C" + p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between C and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("C"  + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort("C" + p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("L" + p_port + "(" + p_mode + ")","2",False)
	 .create
    End With
	With Link	'between L and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("L" + p_port + "(" + p_mode + ")","1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")","GND",False)
	 .create
    End With
   End If


' ------- Topology 7 form 5

If (Topology = 7) Then
   blockname_variable = "Cs" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Capacitance",  Round(Ps/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
   End With
   blockname_variable = "Cp" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Capacitor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+270)
    	Case "L"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+90)
    	Case "D"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset + size_offset/2  )
    		.Rotate (rot_angle+90)
		Case "U"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset - size_offset/2  )
    		.Rotate (rot_angle+270)
    End Select
    .SetDoubleProperty ("Capacitance",  Round(Pp/1e-12,6)  )
    .create
    .SetLocalUnitForProperty ("Capacitance", "pF")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND"  + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
 			    Case  "D"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset + size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset - 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and Cs
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("Cs" + p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between L and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("Cs" + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between Cs and Cp -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Cs" + p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("Cp" + p_port + "(" + p_mode + ")","2",False)
	 .create
    End With
	With Link	'between C and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Cp" + p_port + "(" + p_mode + ")","1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")","GND",False)
	 .create
    End With

End If

' ------- Topology 8 form 5


If (Topology = 8) Then
   blockname_variable = "Ls" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    .position (mws_block_port_position_x +x_offset , mws_block_port_position_y+y_offset )
    .Rotate (rot_angle)
    .SetDoubleProperty ("Inductance",  Round(Ps/1e-6,6)  )
    .create
    .SetLocalUnitForProperty ("Inductance", "uH")
   End With
   blockname_variable = "Lp" + p_port + "(" + p_mode + ")"
   With Block
    .Reset
    .type "CircuitBasic\Inductor"
    .name blockname_variable
    Select Case orientation
    	Case "R"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+270)
    	Case "L"
    		.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset/2 )
    		.Rotate (rot_angle+90)
    	Case "D"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset + size_offset/2  )
    		.Rotate (rot_angle+90)
		Case "U"
			.position (mws_block_port_position_x +2*x_offset +size_offset/2, mws_block_port_position_y+y_offset - size_offset/2  )
    		.Rotate (rot_angle+270)
    End Select
    .SetDoubleProperty ("Inductance",  Round(Pp/1e-6,6)  )
    .create
    .SetLocalUnitForProperty ("Inductance", "uH")
    End With

   	With Block ' check existance
    	.name "GND" + p_port + "(" + p_mode + ")"
		If  .doesexist Then
 		 .delete
 		End If
   		End With
  		With Block
			.Reset
			.Type ("Ground")
			.name "GND"  + p_port + "(" + p_mode + ")"

			Select Case orientation
    			Case "L","R"
    				.position (mws_block_port_position_x +1.5*x_offset , mws_block_port_position_y+y_offset +size_offset  )
 			    Case  "D"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset + size_offset/1  )
					'.rotate (90)
				Case "U"
					.position (mws_block_port_position_x +2*x_offset +size_offset , mws_block_port_position_y+y_offset - 0.2*size_offset   )
   				End Select
			.Create
			End With

	Next_port_name = nextUniquePortName()
  With ExternalPort
    .Name Next_port_name
	.position (mws_block_port_position_x+2*x_offset,mws_block_port_position_y+2*y_offset )
    .create
    .SetDifferential differential_flag
	.SetFixedImpedance( True)
	.SetImpedance Line_Impedance

  End With
   With Link	'Links between ext.port and L
	 .Reset
	 .setsourceportfromexternalport  (Next_port_name,False)
	 .settargetportfromblockport ("Ls" + p_port + "(" + p_mode + ")","2", False)
	 .create
    End With
    With Link	'between L and MWS-ports
	 .Reset
	 .SetSourcePortFromBlockPort(SchematicBlockName, pin_name,  False)  ' DS_port_index,False
	 .settargetportfromblockport ("Ls" + p_port + "(" + p_mode + ")","1",False)
	 .create
    End With

	With Link	'between C and L -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Ls" + p_port + "(" + p_mode + ")","2",False)  ' DS_port_index,False
	 .settargetportfromblockport ("Lp" + p_port + "(" + p_mode + ")","2",False)
	 .create
    End With
	With Link	'between C and GND -ports
	 .Reset
	 .SetSourcePortFromBlockPort("Lp" + p_port + "(" + p_mode + ")","1",False)  ' DS_port_index,False
	 .settargetportfromblockport ("GND" + p_port + "(" + p_mode + ")","GND",False)
	 .create
    End With

End If



Exit All
End Sub




Private Sub FillPortNameArray()
' -------------------------------------------------------------------------------------------------
' FillPortNameArray: This function fills the global array of port names
' -------------------------------------------------------------------------------------------------

	Dim nIndex As Long, nCount As Long

	' Determine the total number of ports first and reset the enumberation to the beginning of the
	' ports list.

	nports = Port.StartPortNumberIteration
   If nports = 0 Then
    	MsgBox("No ports defined, exiting.","Error")
    	Exit All
    End If

	' Make the port name array large enough to hold all port names

	ReDim portnamearray(nports)

	' Now loop over all ports and add the ports to the port array

	nCount = 0
	Dim strPortName As Integer
	For nIndex = 0 To nports-1


		strPortName = Port.GetNextPortNumber

		portnamearray(nCount) = CStr(strPortName)
		nCount = nCount + 1

	Next nIndex

	' Adjust the length of the port name array to the actual number

	If (nCount < nports) Then
		ReDim Preserve portnamearray(nCount)
	End If

End Sub

Sub enlarge_mode_list (modelist() As String, portindex As String)

 Dim cst_runindex As Integer, nr_of_modes_found As Integer
 Dim mode_type As String, maxmodenr As Integer, maxmodeatport As Integer, Number_of_Modes As Integer
 Dim len_of_Field As Integer, run_index As Integer, mode_impedance As Double

 len_of_Field = 0
 maxmodeatport=100

' Waveguide-Port Check

 For Number_of_Modes = 1 To maxmodeatport		'loop over all Modes
  mode_type = ""
  On Error Resume Next
  mode_type = CStr(Port.GetModeType(portindex, Number_of_Modes))
  On Error GoTo 0
  If (mode_type = "") Then Exit For
  len_of_Field = len_of_Field+1
  'Next Number_of_Modes
  ReDim  Preserve modelist (len_of_Field-1)
  For  run_index = 0 To len_of_Field-1     'Set Portnumbers consequ. from 1,2,etc.
    modelist (run_index) = CStr(run_index+1)
  Next run_index
 Next Number_of_Modes

' end waveguide-ports

  ' Discrete-Port Check
 On Error Resume Next
 mode_impedance = Port.GetWaveImpedance(portindex, 1)
 On Error GoTo 0
 If mode_impedance <> 0 Then
    mode_impedance=0 'Reset, remain On old value If Error
    GoTo nextloopindex
 End If
 len_of_Field = len_of_Field+1
 ReDim Preserve modelist(len_of_Field-1)
 For  run_index = 0 To len_of_Field-1
        modelist (run_index) = CStr(run_index+1)
 Next run_index

 nextloopindex:

 ' end discrete-ports

End Sub


Function RealVal_old(lib_Text As Variant) As Double

        If (CDbl("0.5") > 1) Then
                RealVal_old = CDbl(Replace(lib_Text, ".", ","))
        Else
                On Error Resume Next
                        RealVal_old = CDbl(lib_Text)
                On Error GoTo 0
        End If

End Function


