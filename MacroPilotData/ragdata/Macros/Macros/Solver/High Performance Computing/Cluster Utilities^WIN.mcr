'#Language "WWB-COM"
'#Uses "CSTClusterConf.cls"
''' @version 20180705
' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 04-May-2018 ube: First version
' ================================================================================================
Option Explicit

Sub Main ()
	On Error GoTo HandleError_noconf

	' This Macro submits the currently opened CST project a Cluster. See the Cluster Integration FAQ for more documentation.

	Shell("msg * /TIME:3 Configuring and connecting to scheduler system. Please wait... Status will be displayed in Messages-window.")
	Dim ClusterConf As CSTClusterConf
	Set ClusterConf = New CSTClusterConf

	ClusterConf.Configure

	On Error GoTo HandleError

	ClusterConf.Run

ExitProc:

	Set ClusterConf = Nothing 'does not work when ConfigGui was tested
	Exit Sub
HandleError:
	ClusterConf.Log(CStr(Err.Description) + CStr(vbExclamation) + " Error " + CStr(Err.Number) + CStr(Err.Description),1)
	ReportInformationToWindow ( CStr(Err.Description) + CStr(vbExclamation) + " Error " + CStr(Err.Number) + CStr(Err.Description) + " in Interact With Scheduler")
	MsgBox CStr(Err.Description)
	Resume ExitProc

HandleError_noconf:

	MsgBox CStr(Err.Description)
	ReportError (" Error " + CStr(Err.Number) + ": " + CStr(Err.Description) + " in Interact With Scheduler")
	Resume ExitProc
	
End Sub

