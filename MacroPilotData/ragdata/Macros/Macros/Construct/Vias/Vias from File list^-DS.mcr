' *Construct / Vias / Vias from File list
' !!! Do not change the line above !!!
' macro.516
'
' macro creates cylinders from a File list in the format
'	x1	y1	z11	z12
'	x2	y2	z21	z22
'	x3	y3	z31	z31
'  .....
'
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 22-Nov-2016 mki: changed from 'print' to 'AddtoHistory'
' 30-Jul-2009 ube: Split replaced by CSTSplit, since otherwise competing with standard VBA-Split function
' 19-Oct-2007 ube: converted to 2008
' 24-Mar-2003 reh: first version
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit
'#include "vba_globals_all.lib"

Sub Main

	Dim cst_infile As String, cst_outfile As String, cstfile As String
	Dim cst_inline As String
	Dim cst_line_array(4) As String
	Dim cst_macro_via_outer_radius As Double, cst_macro_via_inner_radius As Double
	Dim cst_noc As Integer, cst_via_nr As Long
	Dim cst_mod_file As String, cst_sat_file As String, cst_backup_sat_file As String
	Dim cst_tmp As Double, cst_z1 As Double, cst_z2 As Double
	Dim command_contents As String
	Dim command_name As String
    command_name = "Execute Macro: Vias from File List"
    command_contents = ""

	On Error GoTo ERROR_VR_MISSING
	cst_macro_via_outer_radius = RestoreDoubleParameter("via_outer_radius")
	'cst_macro_via_outer_radius = 100
	cst_macro_via_inner_radius = RestoreDoubleParameter("via_inner_radius")
	'cst_macro_via_inner_radius = 90
	On Error GoTo 0

	If cst_macro_via_outer_radius <= 0 Or cst_macro_via_outer_radius <= cst_macro_via_inner_radius Then
		MsgBox "Please check variables, defining via radius.", vbCritical
		Exit All
	End If

	Begin Dialog UserDialog 600,84,"Automatic via creation from file list",.DialogFunc ' %GRID:10,7,1,1
		Text 20,21,90,14,"Input File",.Text1
		TextBox 100,21,390,21,.Infile
		PushButton 500,21,90,21,"Browse",.Browseinputfile
		OKButton 380,56,90,21
		CancelButton 480,56,90,21
		Text 20,56,260,14,"one line per via:   x1   y1   z1   z2",.Text2
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All

	cst_infile = dlg.Infile

	Open cst_infile For  Input As #1

	cst_via_nr = 0

	While Not EOF(1)

		'--- read data from file
		Line Input #1,cst_inline

		'--- split line into parts and export data; check if numeric line before split
		If IsNumeric(Left(LTrim(cst_inline),1)) Or IsNumeric(Left(LTrim(cst_inline),2)) Then

			cst_via_nr = cst_via_nr + 1
			cst_noc = CSTSplit(cst_inline, cst_line_array())
			cst_z1 = RealVal(cst_line_array(2))
			cst_z2 = Realval(cst_line_array(3))
			If cst_z2 < cst_z1 Then
				cst_tmp = cst_z2
				cst_z2 = cst_z1
				cst_z1 = cst_tmp
			End If
			If cst_noc < 4 Then
				MsgBox "Following file format is required:" + vbCrLf + _
						"x1   y1   z1   z2"
				Exit All
			End If

     		command_contents =	command_contents + "Cylinder.Name Solid.GetNextFreeName " + vbLf
			command_contents =	command_contents + "Cylinder.Component ""PEC""" + vbLf
     		command_contents =	command_contents + "Cylinder.Material ""PEC""" + vbLf
     		command_contents =	command_contents + "Cylinder.OuterRadius """+ Replace(Evaluate(cst_macro_via_outer_radius),",",".")+"""" + vbLf
     		command_contents =	command_contents + "Cylinder.InnerRadius """+ Replace(Evaluate(cst_macro_via_inner_radius),",",".")+"""" + vbLf
     		command_contents =	command_contents + "Cylinder.Axis ""z""" + vbLf
     		command_contents =	command_contents + "Cylinder.Zrange """+Eval("cst_z1")+""", """+Eval("cst_z2")+"""" + vbLf
     		command_contents =	command_contents + "Cylinder.Xcenter """+ Eval("cst_line_array(0)")+"""" + vbLf
     		command_contents =	command_contents + "Cylinder.Ycenter """+ Eval("cst_line_array(1)")+"""" + vbLf
     		command_contents =	command_contents + "Cylinder.Create" + vbLf
		End If
	Wend

	Close #1

	AddToHistory(command_name, command_contents)

	On Error Resume Next
	Exit Sub

ERROR_VR_MISSING:
	MsgBox "Please define variables with proper values first: 'via_inner_radius' and 'via_outer_radius' first",vbOkOnly+vbExclamation,"Macro Execution stopped"
	MakeSureParameterExists "via_inner_radius", "0"
	MakeSureParameterExists "via_outer_radius", "0"
	Exit Sub

End Sub
'-------------------------------------------------------------------------------------
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean

	Dim Extension As String, projectdir As String, filename As String
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        Select Case DlgItem
        	Case "Browseinputfile"
        		Extension = "txt;dat"
                projectdir = Dirname(GetProjectbasename)
                filename = GetFilePath(,Extension, projectdir, "Specify geometry file", 0)
                If (filename <> "") Then
                    DlgText "Infile", FullPath(filename, projectdir)
                End If
        	dialogfunc = True
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
