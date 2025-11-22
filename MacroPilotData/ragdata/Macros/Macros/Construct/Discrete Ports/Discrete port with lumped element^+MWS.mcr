' *Construct / Discrete Ports / Discrete port with lumped element
' !!! Do not change the line above !!!
' macro.801

' ================================================================================================
' Copyright 2005-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 15-Mar-2006 msc: Redesign using the AddToHistory command (no update necessary) 
' 01-Mar-2006 imu: Corrected problem with local coordinates 
'		(discrete port was defined in global, lumped elem. in local coords)
' 24-Oct-2005 ube: Included into Online Help
' 26-Aug-2005 ube: small fix: cancel button
' 29-Jul-2005 imu: first version
'--------------------------------------------------------------------------------------------
Dim macropath As String

Const MacroName = "Construct~Discrete_Port_with_Lumped_Elem"

'#include "vba_globals_all.lib"

Sub Main
Dim n_of_ppoints As Integer
Dim i As Integer
Dim Array_x (3) As Double
Dim Array_y (3) As Double
Dim Array_z (3) As Double
Dim portno As Integer
Dim valL As String, valR As String, valC As String
Dim mon_iu_port As Boolean, mon_iu_lumped As Boolean
Dim lumpedType As String
Dim portimp As String

n_of_ppoints=Pick.GetNumberOfPickedPoints
If n_of_ppoints <> 2 Then
	MsgBox _
		"Please pick two points (and only two). "+vbCrLf+"Aborting Macro", _
		vbOkOnly + vbCritical, _
		"Construct / Discrete port with lumped element"
	Exit All
End If

For i = 1 To 2
	If Pick.GetPickpointCoordinates (i, Array_x(i-1), Array_y(i-1), Array_z(i-1))  = True Then
    Else
    	MsgBox "failed to get pickpoint coordinates"
    	Exit All
    End If
Next i

Array_x(2) = (Array_x(0) + Array_x(1))/2
Array_y(2) = (Array_y(0) + Array_y(1))/2
Array_z(2) = (Array_z(0) + Array_z(1))/2

port_nr_offset=Solver.GetNumberOfPorts

If port_nr_offset <> 0 Then
	Port.StartPortNumberIteration
	For i = 1 To Solver.GetNumberOfPorts
		aaa=Port.GetNextPortNumber
	Next i
	portno = aaa+1
Else
	portno = 1
End If

macropath = GetInstallPath + "\Library\Macros"
	Begin Dialog UserDialog 410,441,"Define discrete port in series with lumped element",.dialogfunc ' %GRID:10,7,1,1
		GroupBox 10,77,390,140,"Discrete port with real impedance",.GroupBox1

		Text 20,161,90,14,"Impedance",.Text1
		TextBox 100,154,120,21,.PortImp
		Text 230,154,40,14,"Ohm",.Text2
		GroupBox 10,231,390,175,"Series lumped element",.GroupBox2
		Text 20,105,130,14,"Type: S-parameter",.Text3
		Text 20,126,100,14,"Port Name: ",.Text4
		'ListBox 110,42,90,21,ListArray(),.ListBox1
		CheckBox 30,189,240,14,"Monitor voltage and current",.MonVoltCurrentPort
		OKButton 50,413,90,21
		CancelButton 170,413,90,21
		GroupBox 20,280,320,42,"RLC Parallel",.GroupBox3
		Text 30,301,90,14,"Type",.LumpedTypetext
		OptionGroup .LumpedType
			OptionButton 80,301,100,14,"RLC Serial",.OptionButton1
			OptionButton 200,301,110,14,"RLC Parallel",.OptionButton2
		Text 30,329,90,14,"R [Ohms]:",.aa
		Text 160,329,90,14,"L [H]:",.aa1
		Text 280,329,90,14,"C [F]:",.aa2
		TextBox 20,350,120,21,.ValR
		TextBox 150,350,120,21,.ValL
		TextBox 280,350,110,21,.ValC
		CheckBox 40,385,320,14,"Monitor voltage and current lumped element",.MonVoltCurrentLumped
		Text 110,126,90,14,CStr(portno),.Text5
		Text 30,252,90,14,"Name",.Text6
		TextBox 90,252,150,21,.elname
		Picture 30,7,320,63,"Picture1",0,.Picture1
		PushButton 300,413,90,21,"Help",.Help
	End Dialog
	Dim dlg As UserDialog

	dlg.portimp = "50.0"
	dlg.ValL="0": dlg.ValC="0": dlg.ValR="0":
	dlg.elname = "element"
	If (Dialog(dlg) = 0) Then Exit All

	If dlg.LumpedType = 0 Then
		lumpedType = "RLCSerial"
	Else
		lumpedType = "RLCParallel"
	End If
	mon_iu_lumped = dlg.MonVoltCurrentLumped
	mon_iu_port = dlg.MonVoltCurrentPort
	valR = dlg.valR: valC = dlg.valC: valL = dlg.valL
	portimp = dlg.portimp

 Dim sCommand As String
 sCommand = ""

'Define discrete port
 sCommand = sCommand +  vbCrLf + " " + vbCrLf + "'@ activate global coordinates "+ vbCrLf + "WCS.ActivateWCS ""global""" + vbCrLf
 sCommand = sCommand + vbCrLf + " " + vbCrLf + "'@ define discrete port: "+ CStr(portno) + vbCrLf

 sCommand = sCommand + "With DiscretePort" + vbCrLf + "     .Reset" + vbCrLf + "     .Portnumber " + """ "+ CStr(portno)+ """" + vbCrLf
 sCommand = sCommand + "     .Type ""Sparameter""" + vbCrLf
 sCommand = sCommand + "     .Point1 " + """" + StrValue(Array_x(0)) + """, """+StrValue(Array_y(0))+""", """ +StrValue(Array_z(0)) + """" + vbCrLf
 sCommand = sCommand + "     .Point2 " + """" + StrValue(Array_x(2)) + """, """+StrValue(Array_y(2))+""", """ +StrValue(Array_z(2)) + """" +vbCrLf
 sCommand = sCommand + "     .Impedance """ + portimp+"""" + vbCrLf + "     .UsePickedPoints  False" + vbCrLf
 If mon_iu_port = False Then
    sCommand = sCommand + "     .Monitor False" + vbCrLf
  Else
    sCommand = sCommand + "     .Monitor True" + vbCrLf
  End If

 sCommand = sCommand + "     .Create" + vbCrLf + "End With" + vbCrLf + " " + vbCrLf
 AddToHistory "define discrete port: "+ CStr(portno), sCommand
 sCommand = ""

'Define lumped element
 sCommand = vbCrLf + " " + vbCrLf + "'@ define lumped element: "+ dlg.elname + vbCrLf
 sCommand = sCommand + "With LumpedElement" + vbCrLf + "     .Reset" + vbCrLf + "     .SetName """ + dlg.elname + """" + vbCrLf

 sCommand = sCommand + "     .SetType """ + lumpedType + """" + vbCrLf

 sCommand = sCommand + "     .SetR """ + valR + """" + vbCrLf + "     .SetL  """ + valL + """" + vbCrLf + "     .SetC  """ + valC + """" + vbCrLf

 sCommand = sCommand + "     .SetP1 ""False"", " + """" + StrValue(Array_x(2)) + """, """+StrValue(Array_y(2))+""",""" +StrValue(Array_z(2)) + """" + vbCrLf
 sCommand = sCommand + "     .SetP2 ""False"", " + """" + StrValue(Array_x(1)) + """, """+StrValue(Array_y(1))+""",""" +StrValue(Array_z(1)) + """" +vbCrLf
 sCommand = sCommand + "     .SetInvert ""False""" + vbCrLf
 If mon_iu_lumped = False Then
    sCommand = sCommand + "     .SetMonitor False" + vbCrLf
  Else
    sCommand = sCommand + "     .SetMonitor True" + vbCrLf
  End If

 sCommand = sCommand + "     .Create" + vbCrLf + "End With" + vbCrLf + " " + vbCrLf
'ub do not switch back to local WCS
'ub sCommand = sCommand + vbCrLf + " " + vbCrLf + "'@ activate local coordinates "+ vbCrLf +"WCS.ActivateWCS ""local""" + vbCrLf

 AddToHistory "define lumped element: "+ dlg.elname, sCommand
 sCommand = ""

End Sub

Function smith_BaseName (lib_path As String) As String

        Dim lib_dircount As Integer, lib_extcount As Integer, lib_filename As String

        lib_dircount = InStrRev(lib_path, "\")
        lib_filename = Mid$(lib_path, lib_dircount+1)
        lib_extcount = InStrRev(lib_filename, ".")
        smith_BaseName = Left$(lib_filename, IIf(lib_extcount > 0, lib_extcount-1, 999))

End Function
'---------------------------------------------------------------------
Function StrValue (getthedouble As Double) As String
  StrValue = Replace(CStr(getthedouble),",",".")
End Function

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
    'Dim filename As String, extension As String, index As Integer

	DlgSetPicture "Picture1", macropath + "\Construct\Discrete Ports\Discrete port with lumped element.bmp", 0
    Select Case Action
		Case 1 ' Dialog box initialization
        Case 2 ' Value changing or button pressed
        	Select Case Item
            	Case "Help"
					StartHelp "common_preloadedmacro_construct___discrete_port_with_lumped_element"
                    DialogFunc = True
            End Select
        Case 3 ' ComboBox or TextBox Value changed
        Case 4 ' Focus changed
        Case 5 ' Idle
    End Select

End Function

