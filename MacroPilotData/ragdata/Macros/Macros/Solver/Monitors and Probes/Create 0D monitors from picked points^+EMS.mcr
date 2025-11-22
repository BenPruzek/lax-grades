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
' 05-Nov-2024 mha: Fixed issue with hiding GUI elements in case of solver other then LT caused by last change.
' 13-Mar-2024 mha: Switched from radio buttons for components to check boxes for easier setup with multiple components.
'                  In case of scalar fields, automatically check mark X, uncheck others and grey out.
' 28-Apr-2023 mha: Former behavior: If solver other then LT is active, everything is greyed out.
'                  Also: If no pick points are set, everything is greyed out.
'                  Behavior now: If solver other then LT is active only a corresponding note is shown, explaining that LT solver must be used.
'                  Also: If LT solver is active and no pick points are set, a corresponding note is shown, explainign that pick points are needed.
' 06-Oct-2022 mha: first version
' ================================================================================================

' *** global variables
Dim lNumberOfPickedPoints As Long    ' Number of picked points
Dim bLTSolver             As Boolean ' True if LT solver is active, False otherwise

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Check the solver type:
	Dim sCurrentSolver As String
	sCurrentSolver = GetSolverType
	If LCase(sCurrentSolver) = LCase("LF Time Domain (MQS)") Then
		bLTSolver = True
	Else
		bLTSolver = False
	End If

	' Establish number of picked points
	Dim sHistoryListString  As String
	Dim sHistoryListCaption As String
	Dim lType               As Long    ' indicates type of monitor
	Dim sName               As String  ' Name of monitor (if specified)
	Dim bSpecifiedName      As Boolean ' Indicates whether specified name should be used for monitor
	Dim lRoundToDec         As Long    ' Indicates naming scheme for coordinates in case of automatic naming
	Dim bCouple             As Boolean ' Indicates whether or not coordinates should be coupled to pick points
	Dim bXComponent         As Boolean ' Indicates whether or not a monitor for the the x-component should be used
	Dim bYComponent         As Boolean ' Indicates whether or not a monitor for the the y-component should be used
	Dim bZComponent         As Boolean ' Indicates whether or not a monitor for the the z-component should be used
	Dim bAbsComponent       As Boolean ' Indicates whether or not a monitor for the the absolute value should be used

	lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, lNumberOfPickedPoints, lType, sName, bSpecifiedName, lRoundToDec, bCouple, bXComponent, bYComponent, bZComponent, bAbsComponent) ) Then
		If bLTSolver Then
			If lNumberOfPickedPoints > 0 Then
				If bXComponent Or bYComponent Or bZComponent Or bAbsComponent Then
					CreateHistoryListEntryForMonitorCreation(lNumberOfPickedPoints, sHistoryListString, sHistoryListCaption, lType, bXComponent, bYComponent, bZComponent, bAbsComponent, sName, bSpecifiedName, lRoundToDec, bCouple)
					ResultTree.EnableTreeUpdate(False)
					AddToHistory(sHistoryListCaption, sHistoryListString)
					ResultTree.EnableTreeUpdate(True)
				End If
			End If
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, lNumberOfPoints As Long, lType As Long, sNameMon As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean, bXComponent As Boolean, bYComponent As Boolean, bZComponent As Boolean, bAbsComponent As Boolean) As Boolean
	Begin Dialog UserDialog 384, 379, "Create 0D monitors from picked points",.DialogFunc ' %GRID:3,3,1,1
		' Groupbox
		GroupBox       9,   6, 366, 278, "Monitor settings", .GBMonitor
		GroupBox      20,  21, 343,  85, "Type",             .GBMonitorType
		OptionGroup                                          .OGMonitorType
		OptionButton  30,  40,  61,  14, "B-field",          .OBBField
		OptionButton 110,  40,  63,  14, "H-field",          .OBHField
		OptionButton 192,  40,  61,  14, "E-field",          .OBEField
		OptionButton 272,  40,  62,  14, "D-field",          .OBDField
		OptionButton  30,  61, 145,  14, "Cond. current dens.", .OBCOndCurrDens
		OptionButton 192,  61,  76,  14, "Material",         .OBMaterial
		OptionButton 272,  61,  76,  14, "Potential",        .OBPotential
		OptionButton  30,  82,  90,  14, "Ohmic loss",       .OBOhmicLoss
		GroupBox      20, 108, 343,  43, "Component",        .GBMonitorComponent
		CheckBox     272, 127,  61,  14, "Abs",              .CBComponentAbs
		CheckBox      30, 127,  61,  14, "X",                .CBComponentX
		CheckBox     110, 127,  63,  14, "Y",                .CBComponentY
		CheckBox     192, 127,  61,  14, "Z",                .CBComponentZ
		GroupBox      20, 153, 343, 123, "Naming",           .GBMonitorNaming
		OptionGroup                                          .OGNaming
		OptionButton  30, 172, 134,  14, "Automatic naming", .OBAutomaticNaming
		OptionButton  30, 225, 108,  14, "Specify name",     .OBSpecifyName
		Text          51, 189, 207,  14, "Round coordinates to ... decimal", .TRoundTo01
		OptionGroup                                          .OGDecimal
		OptionButton  62, 205,  45,  14, "1st,",             .OBDecimal01
		OptionButton 116, 205,  50,  14, "2nd,",             .OBDecimal02
		OptionButton 175, 205,  45,  14, "3rd,",             .OBDecimal03
		OptionButton 229, 205,  50,  14, "6th,",             .OBDecimal04
		OptionButton 286, 205,  41,  14, "9th",              .OBDecimal05
		TextBox       51, 245, 207,  21,                     .TBMonName
		Text         263, 248,  81,  14, "+ ""_X_001"" etc.",  .TNameExplanation

		GroupBox       9, 288, 366,  58, "Coupling to picks", .GBCouplingToPicks
		OptionGroup                                           .OGCouplingToPicks
		OptionButton  18, 306, 297,  14, "Couple monitor coordinates to picked points",    .OBCouplingToPicks
		OptionButton  18, 323, 326,  14, "Decouple monitor coordinates and picked points", .OBNoCouplingToPicks

		GroupBox       9,   6, 366, 340, "Please take note:", .GBNote
		Text          18,  22, 348, 108, "This macro can only be used when the LF Time Domain solver is active...", .TNote1
		Text          18,  22, 348, 108, "Currently there are no points picked. This macro requires picked points before it can be used.", .TNote2

		' OK and Cancel buttons
		OKButton      18, 353,  90,  21
		CancelButton 126, 353,  90,  21
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
		bXComponent    = CBool(dlg.CBComponentX)
		bYComponent    = CBool(dlg.CBComponentY)
		bZComponent    = CBool(dlg.CBComponentZ)
		bAbsComponent  = CBool(dlg.CBComponentAbs)
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

	lNumPickedPoints = Pick.GetNumberOfPickedPoints

	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			If Not ( bLTSolver ) Or ( lNumberOfPickedPoints <= 0 ) Then
				' everything greyed out / hidden...
				DlgEnable "OGMonitorType",          False
				DlgValue  "OGMonitorType",          0
				DlgEnable "OGNaming",               False
				DlgValue  "OGNaming",               0
				DlgEnable "OGDecimal",              False
				DlgValue  "OGDecimal",              1
				DlgEnable "TBMonName",              False
				DlgEnable "OGCouplingToPicks",      False
				DlgEnable "CBComponentAbs",         False
				DlgEnable "CBComponentX",           False
				DlgEnable "CBComponentY",           False
				DlgEnable "CBComponentZ",           False
				DlgValue  "CBComponentAbs",         True
				DlgValue  "CBComponentX",           False
				DlgValue  "CBComponentY",           False
				DlgValue  "CBComponentZ",           False
				DlgVisible "GBMonitor",             False
				DlgVisible "GBMonitorType",         False
				DlgVisible "OBBField",              False
				DlgVisible "OBHField",              False
				DlgVisible "OBEField",              False
				DlgVisible "OBDField",              False
				DlgVisible "OBCOndCurrDens",        False
				DlgVisible "OBMaterial",            False
				DlgVisible "OBPotential",           False
				DlgVisible "OBOhmicLoss",           False
				DlgVisible "GBMonitorComponent",    False
				DlgVisible "CBComponentAbs",        False
				DlgVisible "CBComponentX",          False
				DlgVisible "CBComponentY",          False
				DlgVisible "CBComponentZ",          False
				DlgVisible "GBMonitorNaming",       False
				DlgVisible "OBAutomaticNaming",     False
				DlgVisible "OBSpecifyName",         False
				DlgVisible "TRoundTo01",            False
				DlgVisible "OBDecimal01",           False
				DlgVisible "OBDecimal02",           False
				DlgVisible "OBDecimal03",           False
				DlgVisible "OBDecimal04",           False
				DlgVisible "OBDecimal05",           False
				DlgVisible "TBMonName",             False
				DlgVisible "TNameExplanation",      False
				DlgVisible "GBCouplingToPicks",     False
				DlgVisible "OBCouplingToPicks",     False
				DlgVisible "OBNoCouplingToPicks",   False
				DlgVisible "GBNote",                True
				If Not bLTSolver Then
					DlgVisible "TNote1",            True
					DlgVisible "TNote2",            False
				Else
					DlgVisible "TNote1",            False
					DlgVisible "TNote2",            True
				End If
			Else
				DlgVisible "GBNote",            False
				DlgVisible "TNote1",            False
				DlgVisible "TNote2",            False
				DlgEnable "CBComponentAbs",     True
				DlgEnable "CBComponentX",       True
				DlgEnable "CBComponentY",       True
				DlgEnable "CBComponentZ",       True
				DlgValue  "CBComponentAbs",     True
				DlgValue  "CBComponentX",       False
				DlgValue  "CBComponentY",       False
				DlgValue  "CBComponentZ",       False
				DlgEnable "OGNaming",           True
				DlgValue  "OGNaming",           0
				DlgEnable "OGDecimal",          True
				DlgValue  "OGDecimal",          1
				DlgEnable "TBMonName",          False
				DlgEnable "OGCouplingToPicks",  True
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

			ElseIf ( sDlgItem = "OGMonitorType" ) Then
				If ( lSuppValue < 5 ) Then
					DlgEnable "CBComponentAbs", True
					DlgEnable "CBComponentX",   True
					DlgEnable "CBComponentY",   True
					DlgEnable "CBComponentZ",   True
				Else
					DlgValue  "CBComponentAbs", False
					DlgEnable "CBComponentAbs", False
					DlgValue  "CBComponentX",   True
					DlgEnable "CBComponentX",   True
					DlgValue  "CBComponentY",   False
					DlgEnable "CBComponentY",   False
					DlgValue  "CBComponentZ",   False
					DlgEnable "CBComponentZ",   False
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
' lType:               Determines type of monitor
'                      0 -> B-field
'                      1 -> H-field
'                      2 -> E-field
'                      3 -> D-field
'                      4 -> Cond. current dens.
'                      5 -> Material
'                      6 -> Potential
'                      7 -> Ohmic loss
' bXComponent          If True, monitors for the x-component should be set up
' bYComponent          If True, monitors for the y-component should be set up
' bZComponent          If True, monitors for the z-component should be set up
' bAbsComponent        If True, monitors for the absolute value should be set up
' sName                Name of monitor (in case automatic naming is not used)
' bSpecifiedName       Indicates whether automatic naming should be used or not
' lRoundToDec:         Indicates up to which decimal place coordinates should be
'                      rounded in case of automatic naming
' bCouple              Indicates whether or not monitor position should be coupled
'                      to picked points or not
' No error checking so far...
' --------------------------------------------------------------------------------
Private Function CreateHistoryListEntryForMonitorCreation(lNumberOfPoints As Long, sHistoryListString As String, sHistoryListCaption As String, lType As Long, bXComponent As Boolean, bYComponent As Boolean, bZComponent As Boolean, bAbsComponent As Boolean, sName As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean) As Boolean
	Dim dXCoor          As Double ' x-coordinate
	Dim dYCoor          As Double ' y-coordinate
	Dim dZCoor          As Double ' z-coordinate
	Dim lRunningIndex01 As Long   ' running index
	Dim lNumPointsMin1  As Long   ' One less then number of points (for fractions)
	Dim lCounter        As Long   ' Counter...
	Dim sFormatString   As String ' Concerns leading zeros for naming
	Dim sType           As String ' concerns the type...
	Dim sAutoName       As String ' Name used in the end...
	Dim lXNameIndex     As Integer ' concerns naming of monitors points (which index to use to save the name for the x-component)
	Dim lYNameIndex     As Integer ' concerns naming of monitors points (which index to use to save the name for the y-component)
	Dim lZNameIndex     As Integer ' concerns naming of monitors points (which index to use to save the name for the z-component)
	Dim lAbsNameIndex   As Integer ' concerns naming of monitors points (which index to use to save the name for the absolute value)
	Dim lNumNameIndices As Integer ' concerns naming of monitors points (which index to use to save the name for the x-component)

	Select Case lType
	Case 1
		sType     = "H-Field"
		sAutoName = "h-field"
	Case 2
		sType     = "E-Field"
		sAutoName = "e-field"
	Case 3
		sType     = "D-Field"
		sAutoName = "d-field"
	Case 4
		sType     = "Cond. Current Dens."
		sAutoName = "cond. current dens."
	Case 5
		sType     = "Material"
		sAutoName = "material"
	Case 6
		sType     = "Potential"
		sAutoName = "potential"
	Case 7
		sType     = "Ohmic Losses"
		sAutoName = "ohmic losses"
	Case Else
		sType     = "B-Field"
		sAutoName = "b-field"
	End Select

	' Set format for naming...
	sFormatString      = ""
	For lRunningIndex01 = 1 To Len(lNumberOfPoints) STEP 1
		sFormatString = sFormatString & "0"
	Next lRunningIndex01

	' Number of names in second array direction
	lNumNameIndices = 0
	If bXComponent Then
		lXNameIndex     = lNumNameIndices
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bYComponent Then
		lYNameIndex     = lNumNameIndices
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bZComponent Then
		lZNameIndex     = lNumNameIndices
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bAbsComponent Then
		lAbsNameIndex   = lNumNameIndices
		lNumNameIndices = lNumNameIndices + 1
	End If

	sHistoryListString = "" & _
	"' Declare" & vbCrLf & _
	"Dim lRunningIndex   As Long" & vbCrLf & _
	"Dim dXCoordinates() As Double" & vbCrLf & _
	"Dim dYCoordinates() As Double" & vbCrLf & _
	"Dim dZCoordinates() As Double" & vbCrLf & _
	"Dim sCurrentName()  As String" & vbCrLf & vbCrLf & _
	"' Initialize" & vbCrLf & _
	"ReDim dXCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim dYCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim dZCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim sCurrentName(" & CStr(lNumberOfPoints-1) & ", " & CStr(lNumNameIndices-1) & ")" & vbCrLf & vbCrLf

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
			lNumNameIndices = 0
			If bXComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sName & "_X_" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bYComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sName & "_Y_" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bZComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sName & "_Z_" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bAbsComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sName & "_Abs_" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
		Next lRunningIndex01
	Else
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			Pick.GetPickpointCoordinatesByIndex(lRunningIndex01, dXCoor, dYCoor, dZCoor)
			' Set up multi-dimensional array of names... Initialize lNumNameIndices first...
			lNumNameIndices = 0
			If bXComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sAutoName & " (X; " & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bYComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sAutoName & " (Y; " & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bZComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sAutoName & " (Z; " & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
				lNumNameIndices = lNumNameIndices + 1
			End If
			If bAbsComponent Then
				sHistoryListString = sHistoryListString & _
				"sCurrentName(" & CStr(lRunningIndex01) & ", " & CStr(lNumNameIndices) & ") = """ & sAutoName & " (Abs; " & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
			End If
		Next lRunningIndex01
	End If
	sHistoryListString = sHistoryListString & vbCrLf

	sHistoryListString = sHistoryListString & _
	"For lRunningIndex = 0 To " & CStr(lNumberOfPoints-1) & " STEP 1" & vbCrLf

	lNumNameIndices = 0
	If bXComponent Then
		sHistoryListString = sHistoryListString & _
		"    With TimeMonitor0D" & vbCrLf & _
		"        .Reset" & vbCrLf & _
		"        .Name sCurrentName(lRunningIndex, " & CStr(lNumNameIndices) & ")" & vbCrLf & _
		"        .FieldType """ & sType & """" & vbCrLf & _
		"        .Component ""X""" & vbCrLf & _
		"        .UsePickedPoint ""False""" & vbCrLf & _
		"        .CoordinateSystem ""Cartesian""" & vbCrLf & _
		"        .Position dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex)" & vbCrLf & _
		"        .Create" & vbCrLf & _
		"    End With" & vbCrLf
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bYComponent Then
		sHistoryListString = sHistoryListString & _
		"    With TimeMonitor0D" & vbCrLf & _
		"        .Reset" & vbCrLf & _
		"        .Name sCurrentName(lRunningIndex, " & CStr(lNumNameIndices) & ")" & vbCrLf & _
		"        .FieldType """ & sType & """" & vbCrLf & _
		"        .Component ""Y""" & vbCrLf & _
		"        .UsePickedPoint ""False""" & vbCrLf & _
		"        .CoordinateSystem ""Cartesian""" & vbCrLf & _
		"        .Position dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex)" & vbCrLf & _
		"        .Create" & vbCrLf & _
		"    End With" & vbCrLf
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bZComponent Then
		sHistoryListString = sHistoryListString & _
		"    With TimeMonitor0D" & vbCrLf & _
		"        .Reset" & vbCrLf & _
		"        .Name sCurrentName(lRunningIndex, " & CStr(lNumNameIndices) & ")" & vbCrLf & _
		"        .FieldType """ & sType & """" & vbCrLf & _
		"        .Component ""Z""" & vbCrLf & _
		"        .UsePickedPoint ""False""" & vbCrLf & _
		"        .CoordinateSystem ""Cartesian""" & vbCrLf & _
		"        .Position dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex)" & vbCrLf & _
		"        .Create" & vbCrLf & _
		"    End With" & vbCrLf
		lNumNameIndices = lNumNameIndices + 1
	End If
	If bAbsComponent Then
		sHistoryListString = sHistoryListString & _
		"    With TimeMonitor0D" & vbCrLf & _
		"        .Reset" & vbCrLf & _
		"        .Name sCurrentName(lRunningIndex, " & CStr(lNumNameIndices) & ")" & vbCrLf & _
		"        .FieldType """ & sType & """" & vbCrLf & _
		"        .Component ""Abs""" & vbCrLf & _
		"        .UsePickedPoint ""False""" & vbCrLf & _
		"        .CoordinateSystem ""Cartesian""" & vbCrLf & _
		"        .Position dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex)" & vbCrLf & _
		"        .Create" & vbCrLf & _
		"    End With" & vbCrLf
	End If

	sHistoryListString = sHistoryListString & _
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
