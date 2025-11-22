' Launch BetaMatch from MWS and DS
' ================================
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 08-May-2012 ube: First version
' ================================================================================================
'
' For CST 2012 (will not work for earlier versions)
'
' Launch BetaMatch with the current simulation results from CST MWS or DS.
' If more than one port is defined the user will be asked to select a port.
' AR Filter Data can be used (choice) if available (MWS only).
' DS: If more than one S-Parameter Task is defined the user must chose which
' task to use (via selection dialog).
'
' -------------------------------------------
' Version: 120127.0240.23
' Date: 2012-01-27, 02:40
' 
' BetaMatch: 3.1.D60 (build: 09ac65867e94)
' Copyright (C) 2008-2012 MNW Scan Pte Ltd
' -------------------------------------------
'
Option Explicit
Sub Main ()
	Dim bmmacro, macrokey As String
	If GetApplicationName = "DS" Then
		macrokey = "LaunchFromCSTDS"
	ElseIf (Left(GetApplicationName,2)="DS") Then
		macrokey = "LaunchFromCSTDSfromMWS"
	Else
		macrokey = "LaunchFromCSTMWS"
	End If
	bmmacro = GetSetting("..\MNW Scan", "BetaMatch", macrokey)
	If bmmacro <> "" Then
		If Dir(bmmacro) <> "" Then
			RunScript(bmmacro)
		Else
			MsgBox "A script is missing from your BetaMatch installation. " + _
			"This indicates that the current installation most likely is corrupt. " + _
			"Please download the latest version of BetaMatch from http://mnw-scan.com " + _
			"and redo the installation." + vbCrLf + vbCrLf + _
			"If the problem persists please contact support@mnw-scan.com for help. " + vbCrLf + vbCrLf + _
			"Name of missing script: " + vbCrLf+ bmmacro, vbOkOnly, _
			"Script could not be found!"
			Exit All
		End If
	Else
		bmmacro = GetSetting("..\MNW Scan", "BetaMatch", "Installpath")
		If bmmacro = "" Then
			NoBetaMatchDlg
		Else
			MsgBox "The installed version of BetaMatch is outdated and does not support CST. " + _
			"Go to www.mnw-scan.com or contact betamatch@mnw-scan.com to obtain the latest version.", vbOkOnly, _
			"BetaMatch version too old!"
		End If
	End If
End Sub
Private Sub NoBetaMatchDlg()
	Dim message, label As String
	Dim button As Integer
	message = vbCrLf + "BetaMatch is a software tool that calculates and optimizes matching networks " + _
	"for antennas and other 1-port devices. " + _
	"With BetaMatch you can easily calculate an optimized matching network and then " + _
	"transfer the network back to CST for further processing." + vbCrLf + vbCrLf + _
	"BetaMatch could not be launched because either it is not installed or " + _
	"the installed version does not support CST. " + _
	"Contact MNW Scan or go to the website to obtain the latest version." + vbCrLf + vbCrLf+ _
	"To find out more about BetaMatch go to www.mnw-scan.com or contact betamatch@mnw-scan.com"
	label = "BetaMatch could not be found!"
	Begin Dialog UserDialog 510,224,label,.DialogFunction ' %GRID:10,7,1,1
		PushButton 30,196,90,21,"Close",.PushButton3
		PushButton 30,168,230,21,"Open BetaMatch web-site",.PushButton2
		PushButton 260,168,230,21,"Send mail to betamatch@mnw-scan.com",.PushButton1
		Text 30,7,460,154,message,.Text1
		CancelButton 140,196,80,21,.cancel
	End Dialog
	Dim dlg As UserDialog
	button = Dialog(dlg)
	If button = 3 Then
		Shell "cmd /c start mailto:betamatch@mnw-scan.com"
	End If
	If button = 2 Then
		Shell "cmd /c start www.mnw-scan.com"
	End If
End Sub
Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgVisible "cancel", False
	Case 2 ' Value changing or button pressed
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
