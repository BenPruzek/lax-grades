' Show PBA-FPBA-log-file

Sub Main () 
	Shell("notepad.exe " + Chr$(34) + GetProjectPath("Result") + "MCalc.log" + Chr$(34), 1)
End Sub
