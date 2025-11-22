'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"

' ================================================================================================
' Macro: Creates 0D Monitors along curve element.
'
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 08-Aug-2023 mha: Fixed problem that local coordinate system was not considered for picking.
' 28-Apr-2023 mha: Former behavior: If non-thermal sovler is active, everything is greyed out.
'                  Behavior now: If non-thermal solver is active only a corresponding note is shown, explaining that a thermal solver must be used.
' 21-Mar-2023 mha: Split into two history list entries to be conform with FastLoading option
' 06-Oct-2022 mha: Added global variable to set default number of points
' 28-Sep-2022 mha: fixed a typo
' 15-Mar-2022 mha: first version
' ================================================================================================

' *** global variables
Dim lDefaultNumberOfPoints As Long
Dim bClassicSolver         As Boolean
Dim bCHTSolver             As Boolean

Sub Main
	' Set global variables:
	lDefaultNumberOfPoints = 10

	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Determine active solver
	Dim sCurrentSolver As String
	sCurrentSolver = GetSolverType
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

	' Populate array with curve element names and declare corresponding variables
	Dim sCurves()            As String  ' Array with all curves
	Dim lNumCurves           As Long    ' Number of curves present in model
	Dim sChosenCurveElement  As String  ' Curve element which is chosen in the end
	Dim lNumberOfPoints      As Long    ' Number of points/coordinates to be extracted
	Dim sHistoryListString1  As String
	Dim sHistoryListCaption1 As String
	Dim sHistoryListString2  As String
	Dim sHistoryListCaption2 As String
	Dim lType                As Long    ' indicates type of monitor
	Dim sName                As String  ' Name of monitor (if specified)
	Dim bSpecifiedName       As Boolean ' Indicates whether specified name should be used for monitor
	Dim lRoundToDec          As Long    ' Indicates naming scheme for coordinates in case of automatic naming
	Dim bCouple              As Boolean ' Indicates whether monitor coordinates should be coupled to curve or not...

	lNumCurves = PopulateStringArrayWithNamesOfCurveElements(sCurves)
	If lNumCurves <= 0 Then
		ReDim sCurves(0)
		sCurves(0) = "No curve elements available..."
	End If

	' Store relevant data / default settings
	StoreScriptSetting("SSlNumCurves", Cstr(lNumCurves))
	StoreScriptSetting("SSlNumPoints", CStr(lDefaultNumberOfPoints))
	StoreScriptSetting("SSsFileName",  "CurveElementExport.txt")
	StoreScriptSetting("SSsDelimiter",  ", ")

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, sCurves, lNumCurves, sChosenCurveElement, lNumberOfPoints, lType, sName, bSpecifiedName, lRoundToDec, bCouple) ) Then
		If bCHTSolver Or bClassicSolver Then
			If lNumCurves > 0 Then
				CreateHistoryListEntryForPickPointCreation(sChosenCurveElement, lNumberOfPoints, sHistoryListString1, sHistoryListCaption1, lType)
				CreateHistoryListEntryForMonitorCreation(sChosenCurveElement, lNumberOfPoints, sHistoryListString2, sHistoryListCaption2, lType, sName, bSpecifiedName, lRoundToDec, bCouple)
				ResultTree.EnableTreeUpdate(False)
				If bCouple Then
					AddToHistory(sHistoryListCaption1, sHistoryListString1)
				End If
				AddToHistory(sHistoryListCaption2, sHistoryListString2)
				ResultTree.EnableTreeUpdate(True)
			End If
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, sCurvesArray() As String, lNumCurves As Long, sChosenCurveElementName As String, lNumberOfPoints As Long, lType As Long, sNameMon As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean) As Boolean
	Begin Dialog UserDialog 390, 307, "Create 0D monitors along curve element",.DialogFunc ' %GRID:3,3,1,1
		' Groupbox
		GroupBox       9,   6, 372,  69, "Curve element settings",               .GBCurveElement
		Text          18,  24, 138,  14, "Select curve element:",                .TSelectCurveElement
		DropListBox  161,  21, 210,  21, sCurvesArray(),                         .DLBCurveElements
		Text          18,  51, 235,  14, "Specify number of points to be used:", .TSpecifyNumberOfPoints
		TextBox      258,  48, 113,  21,                     .TBNumberOfPoints

		GroupBox       9,  77, 372, 141, "Monitor settings", .GBMonitor
		OptionGroup                                          .OGMonitorType
		OptionButton  18,  95, 102,  14, "Temperature",      .OBTemperature
		OptionButton 140,  95,  72,  14, "Velocity",         .OBVelocity
		OptionButton 232,  95,  76,  14, "Pressure",         .OBPressure
		OptionGroup                                          .OGNaming
		OptionButton  18, 116, 134,  14, "Automatic naming", .OBAutomaticNaming
		OptionButton  18, 169, 108,  14, "Specify name",     .OBSpecifyName
		Text          39, 133, 207,  14, "Round coordinates to ... decimal", .TRoundTo01
		OptionGroup                                          .OGDecimal
		OptionButton  50, 149,  45,  14, "1st,",             .OBDecimal01
		OptionButton 104, 149,  50,  14, "2nd,",             .OBDecimal02
		OptionButton 163, 149,  45,  14, "3rd,",             .OBDecimal03
		OptionButton 217, 149,  50,  14, "6th,",             .OBDecimal04
		OptionButton 274, 149,  41,  14, "9th",              .OBDecimal05
		TextBox       39, 189, 207,  21,                     .TBMonName
		Text         251, 192, 108,  14, "+ ""-001"" etc.",  .TNameExplanation

		GroupBox       9, 220, 372,  55, "Coupling to curve", .GBCoupling
		OptionGroup                                           .OGCouplingCoors
		OptionButton  18, 238, 249,  14, "Couple monitor coordinates to curve",    .OBCouplingCoorsYes
		OptionButton  18, 256, 278,  14, "Decouple monitor coordinates and curve", .OBCouplingCoorsNo

		GroupBox       9,   6, 372, 269, "Please take note:", .GBNote
		Text          18,  22, 354, 108, "This macro can only be used in a thermal simulation environment.", .TNote

		' OK and Cancel buttons
		OKButton      18, 280,  90,  21
		CancelButton 126, 280,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define                  = True
		sChosenCurveElementName = sCurvesArray(dlg.DLBCurveElements)
		lNumberOfPoints         = CLng(dlg.TBNumberOfPoints)
		lType                   = CLng(dlg.OGMonitorType)
		sNameMon                = dlg.TBMonName
		bSpecifiedName          = CBool(dlg.OGNaming)
		bCouple                 = Not CBool(dlg.OGCouplingCoors)
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
	Dim lNumCurveElements As Long
	Dim sValueSet         As String
	lNumCurveElements = CLng(GetScriptSetting("SSlNumCurves", "0"))

	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			If Not ( bClassicSolver Or bCHTSolver ) Then
				' everything greyed out / hidden...

				DlgEnable "TSelectCurveElement",     False
				DlgEnable "DLBCurveElements",        False
				DlgEnable "TSpecifyNumberOfPoints",  False
				DlgEnable "TBNumberOfPoints",        False
				DlgText   "TBNumberOfPoints",        GetScriptSetting("SSlNumPoints", CStr(lDefaultNumberOfPoints))
				DlgEnable "OGMonitorType",           False
				DlgValue  "OGMonitorType",           0
				DlgEnable "OGNaming",                False
				DlgValue  "OGNaming",                0
				DlgEnable "OGDecimal",               False
				DlgValue  "OGDecimal",               1
				DlgEnable "TBMonName",               False
				DlgEnable "OGCouplingCoors",         False
				DlgVisible "GBCurveElement",         False
				DlgVisible "TSelectCurveElement",    False
				DlgVisible "DLBCurveElements",       False
				DlgVisible "TSpecifyNumberOfPoints", False
				DlgVisible "TBNumberOfPoints",       False
				DlgVisible "GBMonitor",              False
				DlgVisible "OBTemperature",          False
				DlgVisible "OBVelocity",             False
				DlgVisible "OBPressure",             False
				DlgVisible "OBAutomaticNaming",      False
				DlgVisible "OBSpecifyName",          False
				DlgVisible "TRoundTo01",             False
				DlgVisible "OBDecimal01",            False
				DlgVisible "OBDecimal02",            False
				DlgVisible "OBDecimal03",            False
				DlgVisible "OBDecimal04",            False
				DlgVisible "OBDecimal05",            False
				DlgVisible "TBMonName",              False
				DlgVisible "TNameExplanation",       False
				DlgVisible "GBCoupling",             False
				DlgVisible "OBCouplingCoorsYes",     False
				DlgVisible "OBCouplingCoorsNo",      False
				DlgVisible "GBNote",                 True
				DlgVisible "TNote",                  True
			Else
				DlgVisible "GBNote",                 False
				DlgVisible "TNote",                  False
				If lNumCurveElements <= 0 Then
					DlgEnable "TSelectCurveElement",    False
					DlgEnable "DLBCurveElements",       False
					DlgEnable "TSpecifyNumberOfPoints", False
					DlgEnable "TBNumberOfPoints",       False
					DlgText   "TBNumberOfPoints",       GetScriptSetting("SSlNumPoints", CStr(lDefaultNumberOfPoints))
					DlgEnable "OGMonitorType",          False
					DlgEnable "OGNaming",               False
					DlgEnable "OGDecimal",              False
					DlgValue  "OGDecimal",              1
					DlgValue  "OGNaming",               0
					DlgEnable "TBMonName",              False
					DlgEnable "OGCouplingCoors",        False
				Else
					DlgEnable "TSelectCurveElement",    True
					DlgEnable "DLBCurveElements",       True
					DlgEnable "TSpecifyNumberOfPoints", True
					DlgEnable "TBNumberOfPoints",       True
					DlgText   "TBNumberOfPoints",       GetScriptSetting("SSlNumPoints", CStr(lDefaultNumberOfPoints))
					DlgEnable "OGMonitorType",          True
					DlgEnable "OBVelocity",             bCHTSolver
					DlgEnable "OBPressure",             bCHTSolver
					DlgValue  "OGMonitorType",          0
					DlgEnable "OGNaming",               True
					DlgValue  "OGNaming",               0
					DlgEnable "OGDecimal",              True
					DlgValue  "OGDecimal",              1
					DlgEnable "TBMonName",              False
					DlgEnable "OGCouplingCoors",        True
				End If
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
			If ( sDlgItem = "TBNumberOfPoints" ) Then
				' Check once to to see whether the entry is in fact a valid numerical value. If not, set value back to default value.
				sValueSet   = DlgText(sDlgItem)
				If sValueSet <> "" Then
					On Error Resume Next
						Evaluate(sValueSet)
					If Err.Number <> 0 Then
						DlgText sDlgItem, CStr(lDefaultNumberOfPoints)
						MsgBox "The value you have entered for the number of points cannot be interpreted and has been set back to the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(sValueSet)) Then
							' Input text can be interpreted as a number, now check whether it is compliant with given constraints (integer and greater then 3)
							If Clng(Evaluate(sValueSet)) >= 3 Then
								DlgText "TBNumberOfPoints", Clng(Evaluate(sValueSet))
							Else
								DlgText sDlgItem, CStr(lDefaultNumberOfPoints)
								MsgBox "Please be aware that the number of points is required to be greater or equal to three. For now the value has been set back to the default value.", vbExclamation
							End If
						Else
							DlgText sDlgItem, CStr(lDefaultNumberOfPoints)
							MsgBox "Please be aware that the entry you have made for the number of points cannot be interpreted. The value has been set back to the default value.", vbExclamation
						End If
					End If
				Else
					DlgText sDlgItem, CStr(lDefaultNumberOfPoints)
					MsgBox "Please be aware that the entry you have made for the number of points cannot be interpreted. The value has been set back to the default value.", vbExclamation
				End If

			ElseIf ( sDlgItem = "TBMonName" ) Then
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



' Populates given string array with curves (or rather curve elements) present in the file
' Returns the number of curve elements
Private Function PopulateStringArrayWithNamesOfCurveElements(sCurves() As String) As Long
	Dim lNumOfCurveElementsPreSet As Long   ' Initial size of array / used for padding later on
	Dim lNumOfCurveElementsActual As Long   ' Actual number of curve elements
	Dim sNameCurveElement         As String ' Name of current curve element
	Dim sTreePathOutside          As String ' Treepath of curves
	Dim sTreePathInside           As String ' Treepath of curves elements

	' Initialize
	lNumOfCurveElementsActual = 0
	lNumOfCurveElementsPreSet = 50
	ReDim sCurves(lNumOfCurveElementsPreSet)

	sTreePathOutside = ResultTree.GetFirstChildName("Curves")

	If sTreePathOutside <> "" Then
		While sTreePathOutside <> ""
			sTreePathInside = ResultTree.GetFirstChildName(sTreePathOutside)
			While sTreePathInside <> ""
				sCurves(lNumOfCurveElementsActual) = Replace(Right(sTreePathInside, Len(sTreePathInside) - Len("Curves\")), "\", ":")
				lNumOfCurveElementsActual          = lNumOfCurveElementsActual + 1
				sTreePathInside                    = ResultTree.GetNextItemName(sTreePathInside)
				If lNumOfCurveElementsActual = UBound(sCurves) Then
					ReDim Preserve sCurves(UBound(sCurves)+lNumOfCurveElementsPreSet)
				End If
			Wend
			sTreePathOutside = ResultTree.GetNextItemName(sTreePathOutside)
		Wend
	Else
		lNumOfCurveElementsActual = 0
	End If

	If lNumOfCurveElementsActual > 0 Then
		ReDim Preserve sCurves(lNumOfCurveElementsActual - 1)
	Else
		ReDim sCurves(lNumOfCurveElementsActual)
	End If

	PopulateStringArrayWithNamesOfCurveElements = lNumOfCurveElementsActual
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
' CreateHistoryListEntryForPickPointCreation
' Creates history list entry to pick points for 0D monitors
' sChosenCurveElement: Name of curve element which was chosen for 0D Monitors
' lNumberOfPoints:     Determines how many monitors should be distributed on the
'                      curve element
' sHistoryListString:  String is filled with history list entry
' sHistoryListCaption: String is filled with caption for history list entry
' lType:               Determines type of monitor (0 -> Temperature, 1 -> Velocity
'                      2 -> Pressure
' No error checking so far...
' --------------------------------------------------------------------------------
Private Function CreateHistoryListEntryForPickPointCreation(sChosenCurveElement As String, lNumberOfPoints As Long, sHistoryListString As String, sHistoryListCaption As String, lType As Long) As Boolean
	Dim dCurvePos       As Double ' Relative position on curve element
	Dim dXCoor          As Double ' x-coordinate
	Dim dYCoor          As Double ' y-coordinate
	Dim dZCoor          As Double ' z-coordinate
	Dim lRunningIndex01 As Long   ' running index
	Dim lNumPointsMin1  As Long   ' One less then number of points (for fractions)
	Dim lCounter        As Long   ' Counter...
	Dim sFormatString   As String ' Concerns leading zeros for naming
	Dim bLWCSActive     As Boolean ' Indicates whether a local working coordinate system is active

	lNumPointsMin1 = lNumberOfPoints-1
	bLWCSActive    = IIf(LCase(WCS.IsWCSActive)="global", False, True)

	sHistoryListString = "" & _
	"' Declare" & vbCrLf & _
	"Dim lRunningIndex As Long" & vbCrLf & _
	"Dim dXCoordinate  As Double" & vbCrLf & _
	"Dim dYCoordinate  As Double" & vbCrLf & _
	"Dim dZCoordinate  As Double" & vbCrLf & _
	"Dim dCurvePos     As Double" & vbCrLf  & vbCrLf

	' In case local working coordinate system is active, it needs to be turned off for coordinates determination...
	If bLWCSActive Then
		sHistoryListString = sHistoryListString & _
		"WCS.ActivateWCS(""global"")" & vbCrLf & vbCrLf
	End If

	sHistoryListString = sHistoryListString & _
	"If Not IsCurrentlyFastLoading Then" & vbCrLf & _
	"    For lRunningIndex = 0 To " & CStr(lNumPointsMin1) & " STEP 1" & vbCrLf & _
	"        dCurvePos = lRunningIndex / " & CStr(lNumPointsMin1) & vbCrLf & _
	"        Curve.SampleCoordinates(""" & sChosenCurveElement & """, dCurvePos, dXCoordinate, dYCoordinate, dZCoordinate)" & vbCrLf & _
	"        Pick.PickPointFromCoordinates(dXCoordinate, dYCoordinate, dZCoordinate)" & vbCrLf & _
	"    Next lRunningIndex" & vbCrLf & _
	"End If" & vbCrLf & vbCrLf

	' In case local working coordinate system was deactivated for determination of coordinates, turn it back on again...
	If bLWCSActive Then
		sHistoryListString = sHistoryListString & _
		"WCS.ActivateWCS(""local"")" & vbCrLf & vbCrLf
	End If

	' Take care of caption:
	Select Case lType
	Case 1
		sHistoryListCaption = "pick point: pick points for time monitor 0d, velocity along curve element """ & sChosenCurveElement & """"
	Case 2
		sHistoryListCaption = "pick point: pick points for time monitor 0d, pressure along curve element """ & sChosenCurveElement & """"
	Case Else
		sHistoryListCaption = "pick point: pick points for time monitor 0d, temp along curve element """ & sChosenCurveElement & """"
	End Select

	CreateHistoryListEntryForPickPointCreation = True
End Function



' --------------------------------------------------------------------------------
' CreateHistoryListEntryForMonitorCreation
' Creates history list entry to create 0D monitors
' sChosenCurveElement: Name of curve element which was chosen for 0D Monitors
' lNumberOfPoints:     Determines how many monitors should be distributed on the
'                      curve element
' sHistoryListString:  String is filled with history list entry
' sHistoryListCaption: String is filled with caption for history list entry
' lType:               Determines type of monitor (0 -> Temperature, 1 -> Velocity
'                      2 -> Pressure
' sName                Name of monitor (in case automatic naming is not used
' bSpecifiedName       Indicates whether automatic naming should be used or not
' lRoundToDec:         Indicates up to which decimal place coordinates should be
'                      rounded in case of automatic naming
' bCouple              Indicates whether coordinates should be coupled to curve
'                      element or not...
' No error checking so far...
' --------------------------------------------------------------------------------
Private Function CreateHistoryListEntryForMonitorCreation(sChosenCurveElement As String, lNumberOfPoints As Long, sHistoryListString As String, sHistoryListCaption As String, lType As Long, sName As String, bSpecifiedName As Boolean, lRoundToDec As Long, bCouple As Boolean) As Boolean
	Dim dCurvePos       As Double ' Relative position on curve element
	Dim dXCoor          As Double ' x-coordinate
	Dim dYCoor          As Double ' y-coordinate
	Dim dZCoor          As Double ' z-coordinate
	Dim lRunningIndex01 As Long   ' running index
	Dim lNumPointsMin1  As Long   ' One less then number of points (for fractions)
	Dim lCounter        As Long   ' Counter...
	Dim sFormatString   As String ' Concerns leading zeros for naming
	Dim sType           As String ' type...
	Dim sAutoName       As String ' concerns naming of monitor points

	lNumPointsMin1 = lNumberOfPoints-1

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
	"Dim dCurvePos       As Double" & vbCrLf & _
	"Dim sCurrentName()  As String" & vbCrLf & vbCrLf & _
	"' Initialize" & vbCrLf & _
	"ReDim dXCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim dYCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim dZCoordinates(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & _
	"ReDim sCurrentName(" & CStr(lNumberOfPoints-1) & ")" & vbCrLf & vbCrLf

	If bCouple Then
		sHistoryListString = sHistoryListString & _
		"For lRunningIndex = 0 To " & CStr(lNumPointsMin1) & " STEP 1" & vbCrLf & _
		"    Pick.GetPickpointCoordinatesByIndex(lRunningIndex, dXCoordinates(lRunningIndex), dYCoordinates(lRunningIndex), dZCoordinates(lRunningIndex))" & vbCrLf & _
		"Next lRunningIndex" & vbCrLf & vbCrLf
	Else
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			dCurvePos = lRunningIndex01 / lNumPointsMin1
			Curve.SampleCoordinates(sChosenCurveElement, dCurvePos, dXCoor, dYCoor, dZCoor)
			sHistoryListString = sHistoryListString & _
			"dXCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dXCoor) & vbCrLf & _
			"dYCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dYCoor) & vbCrLf & _
			"dZCoordinates(" & CStr(lRunningIndex01) & ") = " & CStr(dZCoor) & vbCrLf
		Next lRunningIndex01
		sHistoryListString = sHistoryListString & vbCrLf & vbCrLf
	End If

	If bSpecifiedName Then
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			sHistoryListString = sHistoryListString & _
			"sCurrentName(" & CStr(lRunningIndex01) & ") = """ & sName & "-" & Format(Cstr(lRunningIndex01+1), sFormatString) & """" & vbCrLf
		Next lRunningIndex01
	Else
		For lRunningIndex01 = 0 To lNumberOfPoints-1 STEP 1
			dCurvePos = lRunningIndex01 / (lNumberOfPoints-1)
			Curve.SampleCoordinates(sChosenCurveElement, dCurvePos, dXCoor, dYCoor, dZCoor)
			sHistoryListString = sHistoryListString & _
			"sCurrentName(" & CStr(lRunningIndex01) & ") = """ & sAutoName & " (" & CStr(Round(dXCoor, lRoundToDec)) & " " & CStr(Round(dYCoor, lRoundToDec)) & " " & CStr(Round(dZCoor, lRoundToDec)) & ")""" & vbCrLf
		Next lRunningIndex01
	End If
	sHistoryListString = sHistoryListString & vbCrLf

	sHistoryListString = sHistoryListString & _
	"For lRunningIndex = 0 To " & CStr(lNumberOfPoints-1) & " STEP 1" & vbCrLf & _
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
	Select Case lType
	Case 1
		sHistoryListCaption = "define time monitor 0d: velocity along curve element """ & sChosenCurveElement & """"
	Case 2
		sHistoryListCaption = "define time monitor 0d: pressure along curve element """ & sChosenCurveElement & """"
	Case Else
		sHistoryListCaption = "define time monitor 0d: temp along curve element """ & sChosenCurveElement & """"
	End Select

	CreateHistoryListEntryForMonitorCreation = True
End Function



' -----------------------------------------------------------------------------
' CreateStringFor0DThermalMonitor
' Function returns string with methods for creating 0D time monitor (thermally
' related) Must be based on coordinates, not pickes, only cartesian version
' (currently)
' sName: String, name of monitor
' lType: Type of monitor (0 -> temperature, 1 -> velocity, 2 -> pressure, else -> temperature)
' sPositionX: Position on x-axis
' sPositionY: Position on y-axis
' sPositionZ: Position on z-axis
' lRoundToForName: Number of
' -----------------------------------------------------------------------------
Private Function CreateStringFor0DThermalMonitor(sName As String, bSpecifiedName As Boolean, lType As Long, dPositionX As Double, dPositionY As Double, dPositionZ As Double, lRoundToForName As Long) As String
	Dim sType     As String
	Dim sAutoName As String
	Dim sReturn   As String

	Select Case lType
	Case 1
		sType     = "Velocity"
		sAutoName = "velocity (" & CStr(Round(dPositionX, lRoundToForName)) & " " & CStr(Round(dPositionY, lRoundToForName))  & " " & CStr(Round(dPositionZ, lRoundToForName)) & ")"
	Case 2
		sType     = "Pressure"
		sAutoName = "pressure (" & CStr(Round(dPositionX, lRoundToForName)) & " " & CStr(Round(dPositionY, lRoundToForName))  & " " & CStr(Round(dPositionZ, lRoundToForName)) & ")"
	Case Else
		sType     = "Temperature"
		sAutoName = "temp (" & CStr(Round(dPositionX, lRoundToForName)) & " " & CStr(Round(dPositionY, lRoundToForName))  & " " & CStr(Round(dPositionZ, lRoundToForName)) & ")"
	End Select

	If Not bSpecifiedName Then
		sName = sAutoName
	End If

	sReturn = "With TimeMonitor0D" & vbCrLf  & _
	"    .Reset" & vbCrLf  & _
	"    .Name """ & sName & """" & vbCrLf  & _
	"    .FieldType """ & sType & """" & vbCrLf  & _
	"    .Component ""X""" & vbCrLf  & _
	"    .CoordinateSystem ""Cartesian""" & vbCrLf  & _
	"    .Position """ & CStr(dPositionX) & """, """ & CStr(dPositionY) & """, """ & CStr(dPositionZ) & """" & vbCrLf  & _
	"    .Create" & vbCrLf  & _
	"End With" & vbCrLf & vbCrLf

	CreateStringFor0DThermalMonitor = sReturn
End Function
