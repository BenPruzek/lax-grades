'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"
'#include "exports.lib"

' ================================================================================================
' Macro: Exports coordinates for selected curve element.
'
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 19-Jan-2022 mha: first version
' ================================================================================================

' *** global variables
' No global variables specified...

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Populate array with curve element names and declare corresponding variables
	Dim sCurves()           As String  ' Array with all curves
	Dim lNumCurves          As Long    ' Number of curves present in model
	Dim sChosenCurveElement As String  ' Curve element which is chosen in the end
	Dim lNumberOfPoints     As Long    ' Number of points/coordinates to be extracted
	Dim sExportFileName     As String  ' Name of export file
	Dim bSpecifyDelimiter   As Boolean ' Specifies whether a user defined delimiter should be used.
	Dim sDelimiter          As String  ' Define delimiter
	Dim bHeader             As Boolean ' Indicates whether header should be included in export.
	lNumCurves = PopulateStringArrayWithNamesOfCurveElements(sCurves)
	If lNumCurves <= 0 Then
		ReDim sCurves(0)
		sCurves(0) = "No curve elements available..."
	End If

	' Store relevant data / default settings
	StoreScriptSetting("SSlNumCurves", Cstr(lNumCurves))
	StoreScriptSetting("SSlNumPoints", Cstr(100))
	StoreScriptSetting("SSsFileName",  "CurveElementExport.txt")
	StoreScriptSetting("SSsDelimiter",  ", ")

	' Call the define method and check whether it is completed successfully
	If ( Define("test", True, False, sCurves, lNumCurves, sChosenCurveElement, lNumberOfPoints, sExportFileName, bSpecifyDelimiter, sDelimiter, bHeader) ) Then
		' Export coordinates of curve element
		If Not bSpecifyDelimiter Then
			sDelimiter = GetScriptSetting("SSsDelimiter",  ", ")
		End If
		ExportCoordinatesOfCurveElement(sChosenCurveElement, lNumberOfPoints, sExportFileName, sDelimiter, bHeader)
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean, sCurvesArray() As String, lNumCurves As Long, sChosenCurveElementName As String, lNumberOfPoints As Long, sExportFileName As String, bSpecifyDelimiter As Boolean, sDelimiter As String, bHeader As Boolean) As Boolean
	Begin Dialog UserDialog 389, 178, "Export coordinates from curve element", .DialogFunc ' %GRID:3,3,1,1
		' Groupbox - select solids
		GroupBox       9,   6, 371, 143, "Specify settings",                          .GBSettings
		Text          18,  24, 138,  14, "Select curve element:",                     .TSelectCurveElement
		DropListBox  161,  21, 210,  21, sCurvesArray(),                              .DLBCurveElements
		Text          18,  51, 258,  14, "Specify number of points to be extracted:", .TSpecifyNumberOfPoints
		TextBox      281,  48,  90,  21,                                              .TBNumberOfPoints
		Text          18,  78, 178,  14, "Specify filename (for export):",            .TFileName
		TextBox      201,  75, 170,  21,                                              .TBFileName
		CheckBox      18, 105, 198,  14, "Specify delimiter (for export):",           .CBDelimiter
		TextBox      221, 102, 150,  21,                                              .TBDelimiter
		OptionGroup                                                                   .OGHeaderOrNot
		OptionButton  18, 127, 119,  14, "Include header",                            .OBHeader
		OptionButton 160, 127, 164,  14, "Do not include header",                     .OBNoHeader

		' OK and Cancel buttons
		OKButton      15, 154,  90,  21
		CancelButton 115, 154,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		sChosenCurveElementName = sCurvesArray(dlg.DLBCurveElements)
		lNumberOfPoints         = CLng(dlg.TBNumberOfPoints)
		sExportFileName         = dlg.TBFileName
		bSpecifyDelimiter       = CBool(dlg.CBDelimiter)
		sDelimiter              = dlg.TBDelimiter
		bHeader                 = Not CBool(dlg.OGHeaderOrNot)
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
			If lNumCurveElements <= 0 Then
				DlgEnable "TSelectCurveElement",    False
				DlgEnable "DLBCurveElements",       False
				DlgEnable "TSpecifyNumberOfPoints", False
				DlgEnable "TBNumberOfPoints",       False
				DlgText   "TBNumberOfPoints",       GetScriptSetting("SSlNumPoints", "100")
				DlgEnable "TFileName",              False
				DlgEnable "TBFileName",             False
				DlgText   "TBFileName",             GetScriptSetting("SSsFileName",  "CurveElementExport.txt")
				DlgEnable "CBDelimiter",            False
				DlgEnable "TBDelimiter",            False
				DlgText   "TBDelimiter",            GetScriptSetting("SSsDelimiter",  ", ")
				DlgEnable "OGHeaderOrNot",          False
			Else
				DlgEnable "TSelectCurveElement",    True
				DlgEnable "DLBCurveElements",       True
				DlgEnable "TSpecifyNumberOfPoints", True
				DlgEnable "TBNumberOfPoints",       True
				DlgText   "TBNumberOfPoints",       GetScriptSetting("SSlNumPoints", "100")
				DlgEnable "TFileName",              True
				DlgEnable "TBFileName",             True
				DlgText   "TBFileName",             GetScriptSetting("SSsFileName",  "CurveElementExport.txt")
				DlgEnable "CBDelimiter",            True
				DlgEnable "TBDelimiter",            False
				DlgText   "TBDelimiter",            GetScriptSetting("SSsDelimiter",  ", ")
				DlgEnable "OGHeaderOrNot",          True
			End If
		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "CBDelimiter" ) Then
				If ( lSuppValue = 0 ) Then
					DlgEnable "TBDelimiter", False
				Else
					DlgEnable "TBDelimiter", True
				End If
			End If
		Case 3 ' TextBox or ComboBox text changed
			If ( sDlgItem = "TBFileName" ) Then
				Dim sNoForbiddenChars As String
				sValueSet         = DlgText(sDlgItem)
				If sValueSet = "" Then
					DlgText sDlgItem, "CurveElementExport.txt"
					MsgBox "Please be aware that you seem to have entered an empty string for the filename. As such, the entry has been replaced with the original default value.", vbExclamation
				Else
					' Remove any forbidden characters and eliminate leading "."
					sNoForbiddenChars = NoForbiddenFilenameCharacters(sValueSet)
					sNoForbiddenChars = RemoveLeadingPeriods(sNoForbiddenChars)
					If sNoForbiddenChars = "" Then
						DlgText sDlgItem, "CurveElementExport.txt"
						MsgBox "Please be aware that the filename you have entered consisted only of characters not allowed to be used for filenames or resulted in leading periods. As such, the entry has been replaced with the original default value.", vbExclamation
					ElseIf sNoForbiddenChars <> sValueSet Then
						DlgText sDlgItem, sNoForbiddenChars
						MsgBox "Please be aware that the filename you have entered contained some characters not allowed to be used for filenames. As such, the entry has been slightly altered.", vbExclamation
					End If
				End If
			ElseIf ( sDlgItem = "TBNumberOfPoints" ) Then
				' Check once to to see whether the entry is in fact a valid numerical value. If not, set value back to default value.
				sValueSet   = DlgText(sDlgItem)
				If sValueSet <> "" Then
					On Error Resume Next
						Evaluate(sValueSet)
					If Err.Number <> 0 Then
						DlgText sDlgItem, "100"
						MsgBox "The value you have entered for the number of points cannot be interpreted and has been set back to the default value.", vbExclamation
					Else
						If IsNumeric(Evaluate(sValueSet)) Then
							' Input text can be interpreted as a number, now check whether it is compliant with given constraints (integer and greater then 3)
							If Clng(Evaluate(sValueSet)) >= 3 Then
								DlgText "TBNumberOfPoints", Clng(Evaluate(sValueSet))
							Else
								DlgText sDlgItem, "100"
								MsgBox "Please be aware that the number of points is required to be greater or equal to three. For now the value has been set back to the default value.", vbExclamation
							End If
						Else
							DlgText sDlgItem, "100"
							MsgBox "Please be aware that the entry you have made for the number of points cannot be interpreted. The value has been set back to the default value.", vbExclamation
						End If
					End If
				Else
					DlgText sDlgItem, "100"
					MsgBox "Please be aware that the entry you have made for the number of points cannot be interpreted. The value has been set back to the default value.", vbExclamation
				End If
			ElseIf ( sDlgItem = "TBDelimiter" ) Then
				' Only make sure that the delimiter is not empty
				sValueSet = DlgText(sDlgItem)
				If sValueSet = "" Then
					DlgText sDlgItem, GetScriptSetting("SSsDelimiter",  ", ")
					MsgBox "Please be aware that you have entered an empty string as delimiter. As such the delimter has been reset to the default of """ & GetScriptSetting("SSsDelimiter",  ", ") & """", vbExclamation
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

	sTreePathOutside = Resulttree.GetFirstChildName("Curves")

	If sTreePathOutside <> "" Then
		While sTreePathOutside <> ""
			sTreePathInside = Resulttree.GetFirstChildName(sTreePathOutside)
			While sTreePathInside <> ""
				sCurves(lNumOfCurveElementsActual) = Replace(Right(sTreePathInside, Len(sTreePathInside) - Len("Curves\")), "\", ":")
				lNumOfCurveElementsActual          = lNumOfCurveElementsActual + 1
				sTreePathInside                    = Resulttree.GetNextItemName(sTreePathInside)
				If lNumOfCurveElementsActual = UBound(sCurves) Then
					ReDim Preserve sCurves(UBound(sCurves)+lNumOfCurveElementsPreSet)
				End If
			Wend
			sTreePathOutside = Resulttree.GetNextItemName(sTreePathOutside)
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



' Export coordinates of curve elements
Private Function ExportCoordinatesOfCurveElement(sChosenCurveElement As String, lNumberOfPoints As Long, sExportFileName As String, sDelimiter As String, bHeader As Boolean) As Boolean
	Dim dCurvePos       As Double ' Relative position on curve element
	Dim dXCoor()        As Double ' x-coordinate
	Dim dYCoor()        As Double ' y-coordinate
	Dim dZCoor()        As Double ' z-coordinate
	Dim lRunningIndex01 As Long   ' running index
	Dim oCurveCoors     As Object ' Will be used as 1DC object, where x -> x-coordinates, yReal -> y-coordinates, yImag -> z-coordinates
	Dim lNumPointsMin1  As Long   ' One less then number of points (for fractions)
	Dim sOutPutPath     As String ' Full path of exported file
	Dim sHeader         As String ' Header of file

	' Initialize
	lNumPointsMin1 = lNumberOfPoints-1
	ReDim dXCoor(lNumPointsMin1)
	ReDim dYCoor(lNumPointsMin1)
	ReDim dZCoor(lNumPointsMin1)

	' Determine coordinates
	For lRunningIndex01 = 0 To lNumPointsMin1 STEP 1
		dCurvePos = lRunningIndex01 / lNumPointsMin1
		Curve.SampleCoordinates(sChosenCurveElement, dCurvePos, dXCoor(lRunningIndex01), dYCoor(lRunningIndex01), dZCoor(lRunningIndex01))
	Next lRunningIndex01

	' Write coordinates to resultobject (for easier export through pre-defined function)
	Set oCurveCoors = Result1DComplex("")
	oCurveCoors.Initialize(lNumberOfPoints)
	oCurveCoors.SetArray(dXCoor, "x")
	oCurveCoors.SetArray(dYCoor, "yre")
	oCurveCoors.SetArray(dZCoor, "yim")

	' Make sure export path exists
	sOutPutPath = GetProjectPathMaster_LIB() + "\Export\3d\"
	CST_MkDir sOutPutPath

	' Export data
	If bHeader Then
		' Define header
		sHeader = "" & _
		"# Export curve element coordinates for: " & sChosenCurveElement & vbNewLine & _
		"# x-coordinates in " & Units.GetUnit("Length") & sDelimiter & " y-coordinates in " & Units.GetUnit("Length") & sDelimiter & " z-coordinates in " & Units.GetUnit("Length")
		' Export
		write1DComplexData_LIB(oCurveCoors, sOutPutPath & sExportFileName, sDelimiter, "ReIm", sHeader)
		ReportInformationToWindow("Coordinates for curve element " & """" & sChosenCurveElement & """" & " have been exported to " & """" & sOutPutPath & sExportFileName & """.")
		DS.ReportInformationToWindow("Coordinates for curve element " & """" & sChosenCurveElement & """" & " have been exported to " & """" & sOutPutPath & sExportFileName & """.")
	Else
		write1DComplexData_LIB(oCurveCoors, sOutPutPath & sExportFileName, sDelimiter, "ReIm")
		ReportInformationToWindow("Coordinates for curve element " & """" & sChosenCurveElement & """" & " have been exported to " & """" & sOutPutPath & sExportFileName & """.")
		DS.ReportInformationToWindow("Coordinates for curve element " & """" & sChosenCurveElement & """" & " have been exported to " & """" & sOutPutPath & sExportFileName & """.")
	End If

End Function
