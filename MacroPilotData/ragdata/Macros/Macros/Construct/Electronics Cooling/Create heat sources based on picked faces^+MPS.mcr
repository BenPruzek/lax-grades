'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"

' Macro: Creates heat sources based on picked faces.

' ================================================================================================
' Copyright 2024-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 03-Mar-2024 jga: first version
'------------------------------------------------------------------------------------

' *** global variables
' No global variables specified...

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

'In Case of same Solid names, make sure that faces are Not On the same body...
'perhaps face chain and ensure, id does not match any other???

	' Establish number of picked points
	Dim lNumHeatSourceCandidates As Long    ' Number of heat source candidates
	Dim bWrongMaterial           As Boolean ' Indicates that a picked face belonged to a solid with unsuitable material properties
	Dim sListOfSolidNames()      As String  ' Solid names for the heat source candidates
	Dim lListOfFaceIDs()         As Long    ' Face IDs for the heat source candidates
	Dim sHistoryListString       As String
	Dim sHistoryListCaption      As String
	Dim sName                    As String  ' Name of heat source
	Dim lType                    As Long    ' Indicated type of source
	Dim dPowerLoss               As Double  ' Powerloss or density, depending on type

	lNumHeatSourceCandidates = Pick.GetNumberOfPickedFaces
	bWrongMaterial           = False

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, lNumHeatSourceCandidates, sName, lType, dPowerLoss) ) Then
		' If no faces are picked this can be seen in the dialog - therefore there is no need to issue a warning at this point in time...
		If lNumHeatSourceCandidates > 0 Then
			SetUpListOfSolidNamesAndFaceIDs(sListOfSolidNames, lListOfFaceIDs)
'			lNumHeatSourceCandidates = RemoveEntriesWithUnsuitableMaterial(sListOfSolidNames, lListOfFaceIDs) ' Not available at this stage for technical reasons...
			If lNumHeatSourceCandidates > 0 Then
				CreateHistoryListEntryForHeatSourceCreation(sHistoryListString, sHistoryListCaption, sListOfSolidNames, lListOfFaceIDs, sName, lType, dPowerLoss)
				Resulttree.EnableTreeUpdate(False)
				AddToHistory(sHistoryListCaption, sHistoryListString)
				Resulttree.EnableTreeUpdate(True)
			Else
				ReportWarningToWindow("Unfortunately no potential heat sources candidates with ")
			End If
		End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, lNumberOfPickedFaces As Long, sHeatSourceName As String, lType As Long, dPowerLoss As Double) As Boolean
	Dim sTypeArray(1) As String
	sTypeArray(0) = "Total power per picked solid element (W)"
	sTypeArray(1) = "Power density (W/m^3)"

	Begin Dialog UserDialog 325, 249, "Create heat sources from picked faces", .DialogFunc ' %GRID:3,3,1,1
		' Groupbox
		GroupBox       9,   6, 307, 111, "Basic heat source settings", .GBMonitor
		Text          18,  28,  41,  18, "Name:",                      .TSourceName
		TextBox       69,  24, 149,  21,                               .TBSourceName
		Text         225,  28,  84,  15, "+ ""-001"" etc.",            .TNameExplanation
		DropListBox   18,  55, 290,  21, sTypeArray(),                 .DLBType
		Text          18,  90,  40,  14, "Value:",                     .TLossPower
		TextBox       68,  86, 240,  21,                               .TBLossPower
		GroupBox       9, 120, 307,  95, "Please note",                .GBPleaseNote
		Text          18, 135, 289,  78, "Please ensure that no solid " & _
		                                 "elements are picked twice or " & _
		                                 "that solid elements are " & _
		                                 "picked which already contain " & _
		                                 "a source and that materials " & _
										 "have positive thermal " & _
										 "conductivity. At present no " & _
										 "automatic error " & _
		                                 "checking is performed.",     .TPleaseNote
		' OK and Cancel buttons
		OKButton      18, 221,  90,  21
		CancelButton 126, 221,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define          = True
		sHeatSourceName = dlg.TBSourceName
		lType           = CLng(dlg.DLBType)
		dPowerLoss      = CDbl(dlg.TBLossPower)
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunc: This function defines the dialog box behaviour. It is automatically called
'             whenever the user changes some settings in the dialog box, presses any button
'             or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim lNumPickedFaces As Long
	Dim sValueSet       As String
	Dim sCurrentSolver  As String
	Dim bThermalSolver  As Boolean
	Dim bValidEntry     As Boolean

	lNumPickedFaces = Pick.GetNumberOfPickedFaces
	sCurrentSolver  = GetSolverType

	If LCase(sCurrentSolver) = LCase("Conjugate Heat Transfer") Then
		bThermalSolver = True
	Else
		If LCase(sCurrentSolver) = LCase("Thermal Steady State") Then
			bThermalSolver = True
		Else
			If LCase(sCurrentSolver) = LCase("Thermal Transient") Then
				bThermalSolver = True
			Else
				bThermalSolver = False
			End If
		End If
	End If

	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			If Not ( bThermalSolver ) Then
				' everything greyed out...
				DlgEnable "TBSourceName", False
				DlgText   "TBSourceName", "No thermal solver???"
				DlgEnable "DLBType",      False
				DlgEnable "TBLossPower",  False
			Else
				If lNumPickedFaces <= 0 Then
					DlgEnable "TBSourceName", False
					DlgText   "TBSourceName", "No faces selected..."
					DlgEnable "DLBType",      False
					DlgEnable "TBLossPower",  False
				Else
					DlgEnable "TBSourceName", True
					DlgText   "TBSourceName", "Heatsource"
					DlgEnable "DLBType",      True
					DlgValue  "DLBType",      0
					DlgEnable "TBLossPower",  True
					DlgText   "TBLossPower", "1"
				End If
			End If

		Case 2 ' Value changing or button pressed

		Case 3 ' TextBox or ComboBox text changed
			If ( sDlgItem = "TBSourceName" ) Then
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

			ElseIf (sDlgItem = "TBLossPower") Then
				sValueSet   = DlgText(sDlgItem)
				bValidEntry = CheckValueEnteredForNumeric(sValueSet)

				If Not bValidEntry Then
					DlgText sDlgItem, "1"
					MsgBox "The value you have entered for the power loss or power loss density cannot be interpreted as a numerical value and has been set to the default value." & vbCrLf & "Please enter a different expression.", vbExclamation
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
' CreateHistoryListEntryForHeatSourceCreation
' Creates history list entry to create heat sources
' sHistoryListString:  String is filled with history list entry
' sHistoryListCaption: String is filled with caption for history list entry
' sListOfSolidNames:   String array, conatins names of solids
' lListOfFaceIDs:      Long array, contains face IDs
' sName                Name of heat source (will get counting postfix
' lType:               Determines type of heat source (power specified or power
'                      density)
' dPowerLoss           Power loss per lump, or density
' No error checking so far...
' --------------------------------------------------------------------------------
Private Function CreateHistoryListEntryForHeatSourceCreation(sHistoryListString As String, sHistoryListCaption As String, sListOfSolidNames() As String, lListOfFaceIDs() As Long, sName As String, lType As Long, dPowerLoss As Double) As Boolean
	Dim lRunningIndex01 As Long   ' running index
	Dim lCounter        As Long   ' Counter...
	Dim sFormatString   As String ' Concerns leading zeros for naming

	' Go through coordinates
	sHistoryListString = ""
	lCounter           = 0
	sFormatString      = ""
	For lRunningIndex01 = 1 To Len(UBound(sListOfSolidNames)+1) STEP 1
		sFormatString = sFormatString & "0"
	Next lRunningIndex01

	' Get name of solid and face id and set up heat sources
	For lRunningIndex01 = 0 To UBound(sListOfSolidNames) STEP 1
		lCounter           = lCounter + 1
		sHistoryListString = sHistoryListString & CreateStringForHeatSource(sListOfSolidNames(lRunningIndex01), lListOfFaceIDs(lRunningIndex01), sName & "-" & Format(Cstr(lCounter), sFormatString), lType, dPowerLoss)
	Next lRunningIndex01

	If UBound(sListOfSolidNames) > 0 Then
		sHistoryListCaption = "define heat sources: " & sName & "-" & Format(Cstr(1)) & " to " & sName & "-" & Format(Cstr(lCounter))
	Else
		sHistoryListCaption = "define heat sources: " & sName & "-" & Format(Cstr(1))
	End If

	CreateHistoryListEntryForHeatSourceCreation = True
End Function



' -----------------------------------------------------------------------------
' CreateStringForHeatSource
' Function returns string with methods for creating 0D time monitor (thermally
' related) Must be based on coordinates, not pickes, only cartesian version
' (currently)
' sSolidName: String, name of solid
' lFaceID:    Long, face id
' sName:      String, name of monitor
' lType:      Long, type of heat source (overall power or density specified)
'             0 -> overall power, 1 -> density
' dPowerLoss: Double, value of power loss (total or density)
' -----------------------------------------------------------------------------
Private Function CreateStringForHeatSource(sSolidName As String, lFaceID As Long, sName As String, lType As Long, dPowerLoss As Double) As String
	Dim sType     As String
	Dim sReturn   As String

	Select Case lType
	Case 1
		sType     = "Density"
	Case Else
		sType     = "Integral"
	End Select

	sReturn = "With HeatSource" & vbCrLf  & _
	"    .Reset" & vbCrLf  & _
	"    .Name """ & sName & """" & vbCrLf  & _
	"    .Value """ & CStr(dPowerLoss) & """" & vbCrLf  & _
	"    .ValueType """ & sType & """" & vbCrLf  & _
	"    .Face """ & sSolidName & """, """ & CStr(lFaceID) & """" & vbCrLf  & _
	"    .Create" & vbCrLf  & _
	"End With" & vbCrLf & vbCrLf

	CreateStringForHeatSource = sReturn
End Function



' -------------------------------------------------------------------------------------------------
' CheckValueEnteredForNumeric:
' This function checks given entries for correctnes.
' In case problems are detected with a given entry, the function returns "False", otherwise "True".
' -------------------------------------------------------------------------------------------------
Private Function CheckValueEnteredForNumeric(sValueSet As String) As Boolean
	Dim dValueSet  As Double
	Dim bReturnVal As Boolean

	' Check once to to see whether the entry is in fact a valid numerical value. If not, set value back to default
	On Error Resume Next
	dValueSet = evaluate(sValueSet)
	If Err.Number <> 0 Then
		bReturnVal = False
	Else
		bReturnVal = True
	End If

	CheckValueEnteredForNumeric = bReturnVal
End Function



' -------------------------------------------------------------------------------------------------
' SetUpListOfSolidNamesAndFaceIDs:
' Assigns solid names and face ids from picked faces to arrays
' Returns overall number of entries
' -------------------------------------------------------------------------------------------------
Private Function SetUpListOfSolidNamesAndFaceIDs(sListOfSolidNames() As String, lListOfFaceIDs() As Long) As Long
	Dim lNumPickedFaces As Long
	Dim lRunningIndex   As Long
	Dim lFaceID         As Long

	lNumPickedFaces = Pick.GetNumberOfPickedFaces
	If lNumPickedFaces > 0 Then
		ReDim sListOfSolidNames(lNumPickedFaces-1)
		ReDim lListOfFaceIDs(lNumPickedFaces-1)

		If lNumPickedFaces > 0 Then
			For lRunningIndex = 0 To lNumPickedFaces - 1 STEP 1
				sListOfSolidNames(lRunningIndex) = Pick.GetPickedFaceByIndex(lRunningIndex, lFaceID)
				lListOfFaceIDs(lRunningIndex)    = lFaceID
			Next lRunningIndex
		End If
	End If

	SetUpListOfSolidNamesAndFaceIDs = lNumPickedFaces
End Function



' -------------------------------------------------------------------------------------------------
' RemoveEntriesWithUnsuitableMaterial
' This function removes all entries in the arrays, where the corresponding material does not have a
' non-zero thermal conductivity
' It returns the remaining number of entries
' It is assumed that the arrays have been initialized and contain at least one entry
' -------------------------------------------------------------------------------------------------
Private Function RemoveEntriesWithUnsuitableMaterial(sListOfSolidNames() As String, lListOfFaceIDs() As Long) As Long
	Dim lRunningIndex As Long
	Dim lReturnValue  As Long
	Dim lOrigUBound   As Long

	lReturnValue = UBound(sListOfSolidNames) + 1
	lOrigUBound  = UBound(sListOfSolidNames)

	For lRunningIndex = lOrigUBound To 0 STEP -1
		If Not MaterialHasPositiveThermalConductivity(sListOfSolidNames(lRunningIndex)) Then
			lReturnValue = lReturnValue - 1
			If lReturnValue > 0 Then
				RemoveEntryFromArrays(sListOfSolidNames, lListOfFaceIDs, lRunningIndex)
			End If
		End If
	Next lRunningIndex

	RemoveEntriesWithUnsuitableMaterial = lReturnValue
End Function



' -------------------------------------------------------------------------------------------------
' MaterialHasPositiveThermalConductivity
' This function receives a solid name and checks the corresponding material for positive thermal
' conductivity.
' If this is given, it returns True, otherwise False
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' At the moment not usable, since no way to check for PTC!
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' -------------------------------------------------------------------------------------------------
Private Function MaterialHasPositiveThermalConductivity(sSolidName) As Boolean
	Dim bReturn       As Boolean
	Dim sMaterialName As String
	Dim sMaterialType As String
	Dim dThermalCond1 As Double
	Dim dThermalCond2 As Double
	Dim dThermalCond3 As Double

	sMaterialName = Solid.GetMaterialNameForShape(sSolidName)
	sMaterialType = Material.GetTypeOfMaterial(sMaterialName)
	Material.GetThermalConductivity(sMaterialName, dThermalCond1, dThermalCond2, dThermalCond3)

	If dThermalCond1 = 0 And dThermalCond2 = 0 And dThermalCond3 = 0 Then
		bReturn = False
	Else
		bReturn = True
	End If

	MaterialHasPositiveThermalConductivity = bReturn
End Function



' -------------------------------------------------------------------------------------------------
' RemoveEntryFromArrays
' This function removes entries at the given index from given arrays and redims them.
' As such, only dynamic arrays with more then one entry are expected as arguments
' -------------------------------------------------------------------------------------------------
Private Function RemoveEntryFromArrays(sListOfSolidNames() As String, lListOfFaceIDs() As Long, lIndex As Long) As Boolean
	Dim lRunningIndex01 As Long
	Dim lCurrentUBound  As Long

	lCurrentUBound = UBound(sListOfSolidNames)

	For lRunningIndex01 = lIndex + 1 To lCurrentUBound STEP 1
		sListOfSolidNames(lRunningIndex01 - 1) = sListOfSolidNames(lRunningIndex01)
		lListOfFaceIDs(lRunningIndex01 - 1)    = lListOfFaceIDs(lRunningIndex01)
	Next lRunningIndex01

	lCurrentUBound = lCurrentUBound - 1
	ReDim Preserve sListOfSolidNames(lCurrentUBound)
	ReDim Preserve lListOfFaceIDs(lCurrentUBound)

	RemoveEntryFromArrays = True
End Function
