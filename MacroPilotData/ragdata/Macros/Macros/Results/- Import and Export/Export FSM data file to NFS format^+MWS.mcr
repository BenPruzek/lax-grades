
' ================================================================================================
' This macro converts a Field source monitor (.fsm) file to .nfs format
'
'
' ================================================================================================
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 20-Dec-2021 dta: released in the distribution
' 01-Oct-2013 dta: added coarsening factor and NFS directory choice
' 05-Apr-2013 dta: initial version
' ================================================================================================

'#include "complex.lib"

Const bDebug = False 'Debug flag
Dim GlobalDataFileName As String, NFS_Dir_Suffix As String
Dim CoarseningFactor As Double


Option Explicit
Sub Main

	'Check if there's any field source monitor defined

Dim iii As Long, counter As Long
Dim Nmoni_Fieldsource As Long
Dim Fieldsource_name()  As String

	'Calculate the total number of field source monitors

    Nmoni_Fieldsource=0

    For iii= 0 To Monitor.GetNumberOfMonitors-1
            If (Monitor.GetMonitorTypeFromIndex(iii) = "Fieldsource") Then
                    Nmoni_Fieldsource = Nmoni_Fieldsource + 1

            End If
    Next

    If (Nmoni_Fieldsource=0) Then
        MsgBox "The current project has no field source monitor defined"
    'Exit All

	End If

	Begin Dialog UserDialog 370,147,"Specify NFS export settings",.DialogFunc2 ' %GRID:10,7,1,1
		GroupBox 10,7,330,91,"",.GroupBox1
		Text 20,35,170,14,"Add NFS folder suffix",.Text1
		TextBox 220,28,90,21,.NFSsuffix
		Text 20,70,140,14,"Set coarsening factor",.Text2
		TextBox 220,63,90,21,.cfactor
		OKButton 30,112,90,21
		CancelButton 140,112,90,21
	End Dialog
	Dim dlg2 As UserDialog
	'Dialog dlg2

	' set defaults for dialog
 	dlg2.NFSsuffix="_1"
 	dlg2.cfactor= "1"

 	If (Dialog(dlg2) = 0) Then Exit All ' User pressed Cancel

	NFS_Dir_Suffix=dlg2.NFSsuffix
	CoarseningFactor=evaluate(dlg2.cfactor)




	GlobalDataFileName = GetFilePath("*.fsm;","Field source Files|*.fsm|All Files|*.*",GetProjectPath("Result"),"Select Field source monitor file to load",0)
	If GlobalDataFileName = "" Then Exit All

	Begin Dialog UserDialog 600,98,"Convert Field Source file",.DialogFunc ' %GRID:10,7,1,1
		Text 20,14,580,14,"Converting '"+Split(GlobalDataFileName,"\")(UBound(Split(GlobalDataFileName,"\")))+"'",.FileNameT
		Text 20,42,490,14,"",.OutputT
		OKButton 260,70,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Private Function DialogFunc2(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable("OK", False)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		ExportFSMtoNFS(GlobalDataFileName)
	Case 6 ' Function key
	End Select
End Function



Function ExportFSMtoNFS(DataFileName As String) As Integer

Dim NFSDirectoryPath As String
Dim index As Long


index=InStrRev (DataFileName, ".")			'Returns the index of the last "\" in the datafilename path
NFSDirectoryPath= Left$(DataFileName,index-1)		'Trim the project path starting from the given index+1

With NFSFile
    .Reset
    .SetCoarsening (CoarseningFactor)
    .Write(DataFileName, NFSDirectoryPath+NFS_Dir_Suffix)
End With



'With Monitor
'.Reset
'.Export ("nfs" ,"" ,DataFileName, True)
'End With

DlgText("OutputT", "Done!")
DlgEnable("OK", True)

End Function
