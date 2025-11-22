'#Language "WWB-COM"

'----------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2019-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 06-Jul-2022 ube: added warning about deleting results
' 03-Jul-2019 ube: initial version
'----------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Function dialogfunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization

		Select Case Monitor.GetGlobalOverrideNearfieldCalculation
		Case "per_monitor"
	    	DlgValue("GroupFFCombine",0)
		Case "always_off"
	    	DlgValue("GroupFFCombine",1)
		Case "always_on"
	    	DlgValue("GroupFFCombine",2)
		End Select

    Case 2 ' Value changing or button pressed

    Case 3 ' TextBox or ComboBox text changed

    Case 4 ' Focus changed

    Case 5 ' Idle

    Case 6 ' Function key

    End Select
End Function

Sub Main ()

	Begin Dialog UserDialog 670,238,"Farfield Monitors - Activate fast Combine Results",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,14,630,84,"",.GroupBox1
		OptionGroup .GroupFFCombine
			OptionButton 40,28,520,14,"Automatic (individual setting of each Farfield monitor is considered)",.OptionButton1
			OptionButton 40,49,580,14,"Fast Combine = ON   (for all farfield monitors)",.OptionButton2
			OptionButton 40,70,510,14,"Fast Combine = OFF   (for all farfield monitors)",.OptionButton3
		Text 30,112,610,84,"The Fast Combine Feature drastically reduces disc storage and improves combine result performance for farfield monitors, eg when driven through Schematic AC-Combine Task." + vbCrLf  + vbCrLf +  "Note: This feature is only supported by T-FIT and F solver without unitcell boundaries. When activating it, no nearfields can be calculated (no radial farfield components). Also manual decoupling plane can't be defined.",.Text5
		OKButton 30,210,90,21
		CancelButton 130,210,90,21
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = -1) Then

		Dim sCommand As String
		sCommand = ""

		Select Case dlg.GroupFFCombine
		Case 0
			sCommand = sCommand + "Monitor.SetGlobalOverrideNearfieldCalculation  ""per_monitor""" + vbCrLf
		Case 1
			sCommand = sCommand + "Monitor.SetGlobalOverrideNearfieldCalculation  ""always_off""" + vbCrLf
		Case 2
			sCommand = sCommand + "Monitor.SetGlobalOverrideNearfieldCalculation  ""always_on""" + vbCrLf
		End Select

		If MsgBox("Changing this setting will require a new solver run." + vbCrLf + "Existing results will be deleted, are you sure to continue?",vbExclamation+ vbYesNo) = vbYes Then
			AddToHistory "(*) set fast Calculation Mode for FF Monitors", sCommand
		End If

	End If

End Sub
