' *Construct / Discrete Ports / Multiple discrete Ports
' !!! Do not change the line above !!!

' Pick an arbitrary number of points via "Pick Points" (P)
' acting as 1st points for the discrete ports.
' the 2nd Points are defined by a plane-definition either x,y,z = c
'
' macro.840
'
' ================================================================================================
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 28-Sep-2007 ube: fixed small bug (did not work any longer in 2008) ...
' 23-May-2006 ube: no longer structure macro, assign commands removed, ...
' 24-Oct-2005 ube: Included into Online Help
' 14-Nov-2003 ube: mwsversion4 multiple defined (removed here, is still in vbaglobals.lib)
' 06-Mar-2003 ube: save file (mod+...) before dialogue is opened
' 24-Sep-2002 ube: global lib included
' 12-Aug-2002 fhi: Append Port-definitions to the mod-file and start rebuild
' 29-Jul-2002 fhi: initial version; limitation: only global coordiante system

Option Explicit

'#include "vba_globals_all.lib"

Sub Main
 Dim cst_type_of_port As String, cst_coordinate_xyz As String, cst_footpoint As Double
 Dim cst_impedance As Double , cst_voltagecurrent As Double , cst_checkbox1 As Integer

 Dim port_nr_offset As Integer
 Dim n_of_ppoints As Integer
 Dim i As Integer
 Dim Array_x () As Double
 Dim Array_y () As Double
 Dim Array_z () As Double
 Dim type_array(3) As String
 Dim list_array_x(3) As String, projectname As String
 
 n_of_ppoints=Pick.GetNumberOfPickedPoints
 If n_of_ppoints = 0 Then
		MsgBox _
			"No Points are picked - aborting Macro", _
			vbOkOnly + vbCritical, _
			"Construct / Multiple discrete Ports"
		Exit All

 End If

 ReDim Array_x(n_of_ppoints)
 ReDim Array_y(n_of_ppoints)
 ReDim Array_z(n_of_ppoints)
	
 For i = 1 To n_of_ppoints
       If Pick.GetPickpointCoordinates (i, Array_x(i-1), Array_y(i-1), Array_z(i-1))  = True Then  
       Else 
        MsgBox "failed to get pickpoint coordinates"
       End If
 Next i
    
    'dialog
	Begin Dialog UserDialog 390,224,"Generating multiple Discrete Ports",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 210,28,50,192,list_array_x(),.DropListBox1
		TextBox 210,77,110,21,.footpoint
		DropListBox 20,28,130,192,type_array(),.type_of_port
		TextBox 20,77,130,21,.voltagecurrent
		TextBox 20,126,130,21,.impedance
		'CheckBox 330,91,70,21,"local_coord",.CheckBox1
		Text 20,7,90,14,"Type of Port",.Text1
		Text 210,7,100,14,"Port orientation",.Text2
		Text 20,56,140,21,"Voltage or Current",.Text3
		Text 20,105,140,21,"Impedance:",.Text4
		Text 210,56,140,14,"Footpoint Definition",.Text5
		CheckBox 40,161,230,14,"Monitor Voltage and Current",.CheckBox1
		OKButton 40,189,90,21
		CancelButton 150,189,90,21
		Text 210,105,170,14,"(=Groundplane coordinate)",.Text6
		PushButton 260,189,90,21,"Help",.Help
		'Text 330,63,210,21,"Use local/global coordinates",.Text6
	End Dialog
	Dim dlg As UserDialog
	
 list_array_x(0) = "x"
 list_array_x(1) = "y"
 list_array_x(2) = "z"
 dlg.footpoint ="0.0"
 type_array(0)= "S_parameter"
 type_array(1)= "Voltage"
 type_array(2)= "Current"
 dlg.impedance="50."
 dlg.voltagecurrent ="1."
 dlg.checkbox1 = 0
	
 If (Dialog(dlg) = 0) Then Exit All
    
 cst_type_of_port = type_array(dlg.type_of_port)
 cst_coordinate_xyz = list_array_x(dlg.droplistbox1)
 cst_footpoint = RealVal(dlg.footpoint)
 cst_voltagecurrent = RealVal(dlg.voltagecurrent)
 cst_impedance = RealVal(dlg.impedance)
 cst_checkbox1=CInt(dlg.checkbox1)

 port_nr_offset=Solver.GetNumberOfPorts

 Dim sCommand As String

 For i = 0 To n_of_ppoints-1

	sCommand = ""
	sCommand = sCommand + "With DiscretePort" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Portnumber """+ CStr(i+1+port_nr_offset)+ """" + vbLf

	sCommand = sCommand + "     .Voltage """+ StrValue(cst_voltagecurrent)+"""" + vbLf
	sCommand = sCommand + "     .Current """+ StrValue(cst_voltagecurrent)+"""" + vbLf
	Select Case cst_type_of_port
	Case "S_parameter"
		sCommand = sCommand + "     .Type ""Sparameter""" + vbLf
	Case "Voltage"
		sCommand = sCommand + "     .Type ""Voltage""" + vbLf
	Case "Current"
		sCommand = sCommand + "     .Type ""Current""" + vbLf
	Case Else
		sCommand = sCommand + "     .Type ""Sparameter""" + vbLf
	End Select

	sCommand = sCommand + "     .Point1 """ +StrValue(Array_x(i))+""", """+StrValue(Array_y(i))+""", """ +StrValue(Array_z(i)) + """" + vbLf

	Select Case  cst_coordinate_xyz
	Case "x"
		sCommand = sCommand + "     .Point2 """ + StrValue(cst_footpoint) + """, """+StrValue(Array_y(i))+""", """+StrValue( Array_z(i)) + """" + vbLf
	Case "y"
		sCommand = sCommand + "     .Point2 """ + StrValue(Array_x(i))+ """, """+StrValue(cst_footpoint)+""", """+StrValue( Array_z(i)) + """" + vbLf
	Case"z"
		sCommand = sCommand + "     .Point2 """ + StrValue(Array_x(i))+""", """+StrValue( Array_y(i))+""", """+StrValue(cst_footpoint) + """" + vbLf
	Case Else
	End Select

    sCommand = sCommand + "     .Impedance """ + StrValue(cst_impedance)+"""" + vbLf
    sCommand = sCommand + "     .UsePickedPoints False" + vbLf

    If cst_checkbox1 = 0 Then
       sCommand = sCommand + "     .Monitor False" + vbLf
    Else
       sCommand = sCommand + "     .Monitor True" + vbLf
    End If

    sCommand = sCommand + "     .Create" + vbLf
    sCommand = sCommand + "End With" + vbLf

    AddToHistory "define discrete port: "+ CStr(i+1+port_nr_offset), sCommand

 Next i

End Sub
'---------------------------------------------------------------------
Function StrValue (getthedouble As Double) As String
  StrValue = Replace(CStr(getthedouble),",",".")
End Function

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "Help"
			StartHelp "common_preloadedmacro_Construct_Multiple_Discrete_Ports"
			DialogFunc = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
