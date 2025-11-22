' Set Port Target Cut Off Frequency

' Option Explicit

' ================================================================================================
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------------
' 23-Jun-2015 mbk: first version
'------------------------------------------------------------------------------------------

Sub Main ()

Dim scommand As String
Dim sheader As String

DeleteResults 'Otherwise to subsequent runs with different frequencies will not work

Begin Dialog UserDialog 340,110,"Set Port Target Cut Off Frequency" ' %GRID:5,5,1,1
	OKButton 10,90,150,20
	CancelButton 170,90,150,20
	GroupBox 5,10,325,75,"",.GroupBox1
	Text 15,25,90,15,"Port Number:",.Text1
	TextBox 205,20,90,20,.pnum
	Text 15,45,180,15,"Target Cut Off Frequency",.Text6
	Text 15,60,100,15,"(in project units)",.Text7
	TextBox 205,45,90,20,.target
'	PushButton 220,310,90,20,"Help",.Help
End Dialog
Dim dlg As UserDialog

dlg.pnum   = "1"
dlg.target = "0"

cst_result = Dialog(dlg)

target_port	     = dlg.pnum
target_frequency = Cstr(Evaluate(dlg.target)*Units.GetFrequencyUnitToSI)

cst_result = Evaluate(cst_result)
If (cst_result =0) Then Exit All   ' if cancel/help is clicked, exit all
If (cst_result =1) Then Exit All

'The Target Cut Off Frequency only works with the Generalized Port Mode Solver
'Compared to many port modes, it can only give the same results, if lower order port modes are then still absorbed
'Therefore the absorption of unconsidered mode fields is activated

sheader = "(*) Set Target Cut Off Frequency for Port " + dlg.pnum
scommand =	"With Solver" + vbCrLf + _
			".WaveguidePortGeneralized ""True""" + vbCrLf + _
			".AbsorbUnconsideredModeFields ""Activate""" + vbCrLf + _
			"End With" + vbCrLf + _
			vbCrLf + _
			"Port.SetTargetCutOffFrequency(""" + target_port  + """,""" + target_frequency + """)"


AddToHistory sheader, scommand

End Sub
