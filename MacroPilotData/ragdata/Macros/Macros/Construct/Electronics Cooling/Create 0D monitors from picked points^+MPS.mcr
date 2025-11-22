'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"

' ================================================================================================
' Macro: Creates 0D Monitors from all picked points.
'
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 28-Apr-2023 mha: Former behavior: If no pick points are set, everything is greyed out.
'                  Behavior now: If no pick points are set, a corresponding note is shown, explainign that pick points are needed.
' 16-Mar-2022 mha: added option to keep connection to pick coordinates alive
' 15-Mar-2022 mha: first version
' ================================================================================================

' *** global variables
Dim lNumberOfPickedPoints As Long ' Number of picked points

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Establish number of picked points
	Dim sHistoryListString  As String
	Dim sHistoryListCaption As String
	Dim lType               As Long    ' indicates type of monitor
	Dim sName               As String  ' Name of monitor (if specified)
	Dim bSpecifiedName      As Boolean ' Indicates whether specified name should be used for monitor
	Dim lRoundToDec         As Long    ' Indicates naming scheme for coordinates in case of automatic naming
	Dim bCouple             As Boolean ' Indicates whether or not coordinates should be coupled to pick points

	lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, lNumberOfPickedPoints, lType, sName, bSpecifiedName, lRoundToDec, bCouple) ) Then
		If lNumberOfPickedPoints > 0 Then
			CreateHistoryListEntryForMonitorCreation(lNumberOfPickedPoints, sHistoryListString, sHistoryListCaption, lType, sName, bSpecifiedName, lRoundToDec, bCouple)
			Resulttree.EnableTreeUpdate(False)
			AddToHistory(sHistoryListCaption, sHistoryListString)
			Resulttree.EnableTreeUpdate(True)
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, lNumberOfPoints As Long, lType As Long, sNameMon As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean) As Boolean
	Begin Dialog UserDialog 361, 241, "Create 0D monitors from picked points",.DialogFunc ' %GRID:3,3,1,1
		' Groupbox
		GroupBox       9,   6, 343, 141, "Monitor settings", .GBMonitor
		OptionGroup                                          .OGMonitorType
		OptionButton  18,  24, 102,  14, "Temperature",      .OBTemperature
		OptionButton 140,  24,  72,  14, "Velocity",         .OBVelocity
		OptionButton 232,  24,  76,  14, "Pressure",         .OBPressure
		OptionGroup                                          .OGNaming
		OptionButton  18,  45, 134,  14, "Automatic naming", .OBAutomaticNaming
		OptionButton  18,  98, 108,  14, "Specify name",     .OBSpecifyName
		Text          39,  62, 207,  14, "Round coordinates to ... decimal", .TRoundTo01
		OptionGroup                                          .OGDecimal
		OptionButton  50,  78,  45,  14, "1st,",             .OBDecimal01
		OptionButton 104,  78,  50,  14, "2nd,",             .OBDecimal02
		OptionButton 163,  78,  45,  14, "3rd,",             .OBDecimal03
		OptionButton 217,  78,  50,  14, "6th,",             .OBDecimal04
		OptionButton 274,  78,  41,  14, "9th",              .OBDecimal05
		TextBox       39, 118, 207,  21,                     .TBMonName
		Text         251, 121,  81,  14, "+ ""-001"" etc.",  .TNameExplanation

		GroupBox       9, 150, 343,  58, "Coupling to picks", .GBCouplingToPicks
		OptionGroup                                           .OGCouplingToPicks
		OptionButton  18, 168, 297,  14, "Couple monitor coordinates to picked points",    .OBCouplingToPicks
		OptionButton  18, 185, 326,  14, "Decouple monitor coordinates and picked points", .OBNoCouplingToPicks

		GroupBox       9,   6, 343, 202, "Please take note:", .GBNote
		Text          18,  22, 348, 108, "This macro can only be utilized in a thermal simulation environment, and not with the Mechanics solver.", .TNote1
		Text          18,  22, 348, 108, "Currently there are no points picked. This macro requires picked points before it can be used.", .TNote2

		' OK and Cancel buttons
		OKButton      18, 215,  90,  21
		CancelButton 126, 215,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define         = True
		lType          = CLng(dlg.OGMonitorType)
		sNameMon       = dlg.TBMonName
		bSpecifiedName = CBool(dlg.OGNaming)
		bCouple        = Not CBool(dlg.OGCouplingToPicks)
		Select Case CLng(dlg.OGDecimal)
		Case 0
			lRoundToDec = 1
		Case 1
			lRoundToDec = 2
		Case 2
			lRoundToDec = 3
		Case 3
			lRoundToDec = 6
		Case 4
			lRoundToDec = 9
		Case Else
			lRoundToDec = 2
		End Select
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunc: This function defines the dialog box behaviour. It is automatically called
'             whenever the user changes some settings in the dialog box, presses any button
'             or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim lNumPickedPoints As Long
	Dim sValueSet        As String
	Dim sCurrentSolver   As String
	Dim bClassicSolver   As Boolean
	Dim bCHTSolver       As Boolean

	lNumPickedPoints = Pick.GetNumberOfPickedPoints
	sCurrentSolver   = GetSolverType

	If LCase(sCurrentSolver) = LCase("Conjugate Heat Transfer") Then
		bCHTSolver     = True
		bClassicSolver = False
	Else
		bCHTSolver     = False
		If LCase(sCurrentSolver) = LCase("Thermal Steady State") Or LCase(sCurrentSolver) = LCase("Thermal Transient") Then
			bClassicSolver = True
		Else
			bClassicSolver = False
		End If
	End If

	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			If Not ( bClassicSolver Or bCHTSolver ) Or ( lNumberOfPickedPoints <= 0 ) Then
				' everything greyed out / hidden...
				DlgEnable "OGMonitorType",        False
				DlgValue  "OGMonitorType",        0
				DlgEnable "OGNaming",             False
				DlgValue  "OGNaming",             0
				DlgEnable "OGDecimal",            False
				DlgValue  "OGDecimal",            1
				DlgEnable "TBMonName",            False
				DlgEnable "OGCouplingToPicks",    False
				DlgVisible "GBMonitor",           False
				DlgVisible "OBTemperature",       False
				DlgVisible "OBVelocity",          False
				DlgVisible "OBPressure",          False
				DlgVisible "OBAutomaticNaming",   False
				DlgVisible "OBSpecifyName",       False
				DlgVisible "TRoundTo01",          False
				DlgVisible "OBDecimal01",         False
				DlgVisible "OBDecimal02",         False
				DlgVisible "OBDecimal03",         False
				DlgVisible "OBDecimal04",         False
				DlgVisible "OBDecimal05",         False
				DlgVisible "TBMonName",           False
				DlgVisible "TNameExplanation",    False
				DlgVisible "GBCouplingToPicks",   False
				DlgVisible "OBCouplingToPicks",   False
				DlgVisible "OBNoCouplingToPicks", False
				DlgVisible "GBNote",              True
				If Not ( bClassicSolver Or bCHTSolver ) Then
					DlgVisible "TNote1",          True
					DlgVisible "TNote2",          False
				Else
					DlgVisible "TNote1",          False
					DlgVisible "TNote2",          True
				End If
			Else
				DlgVisible "GBNote",           False
				DlgVisible "TNote1",           False
				DlgVisible "TNote2",           False
				DlgEnable "OGMonitorType",     True
				DlgEnable "OBVelocity",        bCHTSolver
				DlgEnable "OBPressure",        bCHTSolver
				DlgValue  "OGMonitorType",     0
				DlgEnable "OGNaming",          True
				DlgValue  "OGNaming",          0
				DlgEnable "OGDecimal",         True
				DlgValue  "OGDecimal",         1
				DlgEnable "TBMonName",         False
				DlgEnable "OGCouplingToPicks", True
			End If

		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "OGNaming" ) Then
				If ( lSuppValue = 0 ) Then
					DlgEnable "TBMonName", False
					DlgEnable "OGDecimal", True
				Else
					DlgEnable "TBMonName", True
					DlgEnable "OGDecimal", False
				End If
			End If

		Case 3 ' TextBox or ComboBox text changed
			If ( sDlgItem = "TBMonName" ) Then
				Dim sNoForbiddenChars As String
				sValueSet    = DlgText(sDlgItem)
				If sValueSet = "" Then
					MsgBox "Please be aware that you seem to have entered an empty string for the monitor name. If this value is retained, the monitor names will solely consist of the numbering.", vbExclamation
				Else
					' Remove any forbidden characters and eliminate leading "."
					sNoForbiddenChars = NoForbiddenFilenameCharacters(sValueSet)
					sNoForbiddenChars = LTrim(sValueSet)
					sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
					If sNoForbiddenChars = "" Then
						DlgText sDlgItem, ""
						MsgBox "Please be aware that the monitor name you have entered consisted only of characters not allowed to be used or resulted in leading periods. As such, the entry has been reduced to an empty string. If this value is retained, the monitor names will solely consist of the numbering.", vbExclamation
					ElseIf sNoForbiddenChars <> sValueSet Then
						DlgText sDlgItem, sNoForbiddenChars
						MsgBox "Please be aware that the monitor name you have entered contained some characters not allowed to be used. As such, the entry has been slightly altered.", vbExclamation
					End If
				End If
			End If

		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function



' Removes leading periods in string
Private Function RemoveLeadingPeriods(sMyString As String) As String
	Dim sRemainingStr As String

	If Left(sMyString, 1) = "." Then
		sRemainingStr = Right(sMyString, Len(sMyString)-1)
		While Left(sRemainingStr, 1) = "."
			sRemainingStr = Right(sRemainingStr, Len(sRemainingStr)-1)
		Wend
	Else
		sRemainingStr = sMyString
	End If

	RemoveLeadingPeriods = sRemainingStr
End Function



' --------------------------------------------------------------------------------
' CreateHistoryListEntryForMonitorCreation
' Creates history list entry to create 0D monitors
' lNumberOfPoints:     Number of picked points
' sHistoryListString:  String is filled with history list entry
' sHistoryListCaption: String is filled with caption for history list entry
' lType:               Determines type of monitor (0 -> Temperature, 1 -> Velocity
'                      2 -> Pressure
' sName                Name of monitor (in case automatic naming is not used
' bSpecifiedName       Indicates whether automatic naming should be used or not
' lRoundToDec:         Indicates up to which decimal place coordinates should be
'                      rounded in case of automatic naming
' bCouple              Indicates whether or not monitor position should be coupled
'                      to picked points or not
' No error checking so far...
' --------------------------------------------------------------------------------
Private Function CreateHistoryListEntryForMonitorCreation(lNumberOfPoints As Long, sHistoryListString As String, sHistoryListCaption As String, lType As Long, sName As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean) As Boolean
	Dim dXCoor          As Double ' x-coordinate
	Dim dYCoor          As Double ' y-coordinate
	Dim dZCoor          As Double ' z-coordinate
	Dim lRunningIndex01 As Long   ' running index
	Dim lNumPointsMin1  As Long   ' One less then number of points (for fractions)
	Dim lCounter        As Long   ' Counter...
	Dim sFormatString   As String ' Concerns leading zeros for naming
	Dim sAutoName       As String ' Name used in the end...
	Dim sType           As String

	' Housekeeping (type)
	Select Case lType
	Case 1
		sType     = "Velocity"
		sAutoName = "velocity"
	Case 2
		sType     = "Pressure"
		sAutoName = "pressure"
	Case Else
		sType     = "Temperature"
		sAutoName = "temp"
	End Select

	' Set format for naming...
	sFormatString      = ""
	For lRunningIndex01 = 1 To Len(lNumberOfPoints) STEP 1
		sFormatString = sFormatString & "0"
	Next lRunningIndex01


	sHistoryListString = "" & _
	"' Declare" & vbCrLf & _
	"Dim lRunningIndex   As Long" & vbCrLf & _
	"Dim dXCoordinates() As Double" & vbCrLf & _
	"Dim dYCoordinates() As Double" & vbCrLf & _
	"Dim dZCoordinates() As Double" & vbCrLf & _
	"Dim sCurrentName()  As String" & vbCrLf & vbCrLf & _
	"' Initialize" & vbCrLf & _
	"ReDim dXCoordinates(" & CStr(lNumberOfPoints) & ")" & vbCrLf & _
	"ReDim dYCoordinates(" & CStr(lNumberOfPoints) & ")" & vbCrLf & _
	"ReDim dZCoordinates(" & CStr(lNumberOfPoints) & ")" & vbCrLf & _
	"ReDim sCurrentName(" & CStr(lNumberOfPoints) & ")" & vbCrLf & vbCrLf

	If bCouple Then
		sHistoryListString = sHistoryListString & _
		"For lRunningIndex = 0 To " & CStr(lNumberOfPoints-1) & " STEP 1" & vbCrLf & _
		"    Pick.GetPickpointCoordinatesByIndex(lRunningIndex, dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex))" & vbCrLf & _
		"Next lRunningIndex" & vbCrLf & vbCrLf
	Else
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			Pick.GetPickpointCoordinatesByIndex(lRunningIndex01, dXCoor, dYCoor, dZCoor)
			sHistoryListString = sHistoryListString & _
			"dXCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dXCoor) & vbCrLf & _
			"dYCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dYCoor) & vbCrLf & _
			"dZCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dZCoor) & vbCrLf
		Next lRunningIndex01
		sHistoryListString = sHistoryListString & vbCrLf
	End If

	If bSpecifiedName Then
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			sHistoryListString = sHistoryListString & _
			"sCurrentName(" & CStr(lRunningIndex01) & ") = """ & sName & "-" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
		Next lRunningIndex01
	Else
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			Pick.GetPickpointCoordinatesByIndex(lRunningIndex01, dXCoor, dYCoor, dZCoor)
			sHistoryListString = sHistoryListString & _
			"sCurrentName(" & CStr(lRunningIndex01) & ") = """ & sAutoName & " (" & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
		Next lRunningIndex01
	End If
	sHistoryListString = sHistoryListString & vbCrLf

	sHistoryListString = sHistoryListString & _
	"For lRunningIndex = 0 To " & CStr(lNumberOfPoints) & "-1 STEP 1" & vbCrLf & _
	"    With TimeMonitor0D" & vbCrLf & _
	"        .Reset" & vbCrLf & _
	"        .Name sCurrentName(lRunningIndex)" & vbCrLf & _
	"        .FieldType """ & sType & """" & vbCrLf & _
	"        .Component ""X""" & vbCrLf & _
	"        .CoordinateSystem ""Cartesian""" & vbCrLf & _
	"        .Position dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex)" & vbCrLf & _
	"        .Create" & vbCrLf & _
	"    End With" & vbCrLf & _
	"Next lRunningIndex" & vbCrLf & vbCrLf

	' Take care of caption:
	sHistoryListCaption = "define time monitor 0d: "
	If bSpecifiedName Then
		sHistoryListCaption = sHistoryListCaption & _
		sName & "-" & Format(Cstr(1), sFormatString) & " to " & sName & "-" & Format(Cstr(lNumberOfPoints), sFormatString)
	Else
		Pick.GetPickpointCoordinatesByIndex(0, dXCoor, dYCoor, dZCoor)
		sHistoryListCaption = sHistoryListCaption & _
		sAutoName & " (" & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ") to "
		Pick.GetPickpointCoordinatesByIndex(lNumberOfPoints-1, dXCoor, dYCoor, dZCoor)
		sHistoryListCaption = sHistoryListCaption & _
		sAutoName & " (" & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")"
	End If

	CreateHistoryListEntryForMonitorCreation = True
End Function
