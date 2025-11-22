
' This macro builds an mdif file in AWR or ADS format CST S-Parameter parametric results or tables
'
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 17-Nov-2021 fsr: Sorting bugfixes & performance improvement
' 08-Oct-2021 fsr: Only consider selected parameters for sorting
' 05-Oct-2021 fsr: Fixed problem with sorting for more than 1 dimension
' 30-Sep-2021 fsr: Fixed problem for cases with only a single parameter combination
' 25-Aug-2021 fsr: Check for result tree entry only when checking for complete S-Parameters, not data file behind it as that only exists if current result is present
' 12-Aug-2021 fsr: Added additional output information about file location and reference impedance handling. Fixed ADS format. Implemented rudimentary
'					rudimentary data sorting in increasing parameter values
' 05-Apr-2021 fsr: Index of selected parameter combination was shifted by 1 in some cases. Fixed.
' 10-Mar-2021 fsr: Disable dialog while running; provide progress info in GUI, use buffered write and manual formating vs. USFormat for better performance
' 05-Mar-2019 fsr: All data Access going through parametric tree results Now
' 14-Feb-2017 fsr: Fixed AWR format (units now within ACDATA block)
' 31-Jan-2014 fsr: Parametric results now supported
' 02-Dec-2013 fsr: Fixed AWR format
' 25-Feb-2013 fsr: Added check to make sure at least 1 parameter is selected
' 09-Mar-2012 fsr: Included max number of parameters (7 for AWR, 9 in general); data now written in US format, independent of locale settings
' 27-Feb-2012 fsr: Initial version, from pre-existing internal macros
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'#include "vba_globals_all.lib"
'#include "template_conversions.lib"
'#include "exports.lib"

Option Explicit

Public Const FormatOptions=Array("AWR", "ADS")
Public bUseParameter() As Boolean, sParameterNames() As String, dParameterReferenceValues() As Double
Public Const sPrintFormat = "0.0000000E+00" ' output format for numbers written to file

Sub Main

	Dim sFormatOptions() As String

	FillArray(sFormatOptions, FormatOptions)

	If (InStr(GetSelectedTreeItem, "\S-Parameters") > 0) Then
		' User manually selected an item to export
	ElseIf SelectTreeItem("1D Results\S-Parameters") Then
		' Automatically select from primary results. If it exists, use it.
	ElseIf SelectTreeItem("Tables\1D Results\S-Parameters") Then
		' this works too
	Else
		ReportError("Cannot find S-Parameter results.")
	End If

	Begin Dialog UserDialog 350,140,"Create MDIF File",.DialogFunc ' %GRID:10,7,1,1
		Text 20,21,80,14,"File Format:",.Text1
		DropListBox 260,14,70,177,sFormatOptions(),.FileFormatDLB
		Text 20,49,200,14,"Number of Frequency Samples:",.Text2
		TextBox 260,42,70,21,.NSamplesT
		Text 20,77,230,14,"Reference Impedance:",.Text3
		TextBox 260,70,70,21,.ReferenceImpT
		Text 20,112,90,14,"",.ProgressL
		OKButton 140,105,90,21
		CancelButton 240,105,90,21
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg)=0 Then ' user pressed cancel
		Exit All
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim i As Integer
	Dim vUseParameter As Variant ' to be used as a Boolean Array, determines if parameter at a given index is to be exported or not.
	Dim bSortData As Boolean

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("NSamplesT", "101")
		DlgText("ReferenceImpT", "50")
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "OK"
				' Disable dialog
				For i = 0 To DlgCount-1
					DlgEnable(i, False)
				Next
					If (MsgBox("Please note that all S-Parameters need to be" + vbNewLine _
								+"normalized to the specified reference impedance. Data will not be renormalized automatically. Continue?", vbYesNo, "Port Impedance Check") = vbNo) Then
						Exit All ' Port renormalization may be implemented here later
				End If
				ParameterSetup(DlgText("FileFormatDLB"), vUseParameter, bSortData)
				Select Case DlgText("FileFormatDLB")
					Case "ADS"
						Main_ADS(vUseParameter, bSortData, CLng(DlgText("NSamplesT")), CDbl(DlgText("ReferenceImpT")), "ProgressL")
					Case "AWR"
						Main_AWR(vUseParameter, bSortData, CLng(DlgText("NSamplesT")), CDbl(DlgText("ReferenceImpT")), "ProgressL")
				End Select
			Case "Cancel"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub ParameterSetup(sFormat As String, vUseParameter As Variant, bSortData)

	Dim i As Long
	Dim nParameters As Long, nParameterSets As Long, nParametersUsed As Long
	Dim sIncludeList() As String, sExcludeList() As String
	Dim sListOfRunIDs() As String, vParameterNames As Variant, vParameterValues As Variant

	sListOfRunIDs = GetListOfRunIDs_LIB(Resulttree.GetFirstChildName("1D Results\S-Parameters"))
	If UBound(sListOfRunIDs) = 0 Then ReportError("No parametric S-parameter data found.")
	' "Current" run ID could have any index, and it does not have parameter listed. Go through full list until a valid option is found (usually the first or the second)
	For i = 0 To UBound(sListOfRunIDs)
		If GetParameterCombination(sListOfRunIDs(i), vParameterNames, vParameterValues) Then Exit For
	Next

	nParameters = UBound(vParameterNames)+1
	nParameterSets = UBound(sListOfRunIDs) ' run ID 0 is current, non-parametric result

	ReDim sParameterNames(nParameters-1)
	ReDim bUseParameter(nParameters-1)
	ReDim dParameterReferenceValues(nParameters-1)

	For i = 0 To nParameters-1
		sParameterNames(i) = CStr(vParameterNames(i))
		' Store parameter values of first object as reference; possibly allow user to select reference value in future
		dParameterReferenceValues(i) = CDbl(vParameterValues(i))
	Next

	nParametersUsed = 0
	For i = 0 To nParameters-1
		' include first 7 parameters for AWR by default, first 9 for ADS
		If ((sFormat = "AWR") And (i>6) _
			Or (sFormat = "ADS") And (i>8)) Then
			bUseParameter(i) = False
		Else
			bUseParameter(i) = True
		End If
	Next

	ReDim sIncludeList(UBound(bUseParameter))
	ReDim sExcludeList(UBound(bUseParameter))
	For i = 0 To UBound(bUseParameter)
		If bUseParameter(i) Then
			sIncludeList(i) = sParameterNames(i)
		Else
			sExcludeList(i) = sParameterNames(i)
		End If
	Next

	Begin Dialog UserDialog 400,280,"Parameter Setup",.ParameterSetupDialogFunc ' %GRID:10,7,1,1
		TextBox 20,7,90,21,.InvisFileFormatT
		ListBox 20,35,150,147,sIncludeList(),.IncludeLB
		ListBox 230,35,150,147,sExcludeList(),.ExcludeLB
		Text 20,14,90,14,"Include:",.Text1
		Text 230,14,90,14,"Exclude:",.Text2
		PushButton 180,77,40,21,">",.MoveToExcludePB
		PushButton 180,105,40,21,"<",.MoveToIncludePB
		PushButton 180,49,40,21,">>",.ExcludeAllPB
		PushButton 180,133,40,21,"<<",.IncludeAllPB
		CheckBox 20,217,350,14,"Sort parameter values in increasing order",.SortDataCB
		OKButton 190,245,90,21
		CancelButton 290,245,90,21
		Text 20,182,360,28,"Note: Only unvaried and dependent parameters can be excluded.",.Note
	End Dialog
	Dim dlg As UserDialog

	dlg.InvisFileFormatT = sFormat

	If Dialog(dlg, 999)=0 Then
		Exit Sub
	End If

	vUseParameter = bUseParameter
	bSortData = CBool(dlg.SortDataCB)

End Sub

Private Function ParameterSetupDialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim i As Long
	Dim sIncludeList() As String, sExcludeList() As String, sFormat As String, nParametersUsed As Long

	sFormat = DlgText("InvisFileFormatT")

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgVisible("InvisFileFormatT", False)
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "Cancel"
				Exit All
			Case "OK"
				nParametersUsed = 0
				For i = 0 To UBound(bUseParameter)
					nParametersUsed = nParametersUsed + IIf(bUseParameter(i), 1, 0)
				Next
				If (nParametersUsed = 0) Then
					MsgBox("Please select at least one parameter. The Touchstone format should be used for non-parametric exports.", "Input Error")
					ParameterSetupDialogFunc = True
				Else
					ParameterSetupDialogFunc = False ' Done, close window
				End If
			Case "MoveToExcludePB"
				ParameterSetupDialogFunc = True
				bUseParameter(GetParameterIndexFromName(DlgText("IncludeLB"))) = False
				ReDim sIncludeList(UBound(bUseParameter))
				ReDim sExcludeList(UBound(bUseParameter))
				For i = 0 To UBound(bUseParameter)
					If bUseParameter(i) Then
						sIncludeList(i) = sParameterNames(i)
					Else
						sExcludeList(i) = sParameterNames(i)
					End If
				Next
				DlgListBoxArray("IncludeLB", sIncludeList)
				DlgListBoxArray("ExcludeLB", sExcludeList)
			Case "MoveToIncludePB"
				ParameterSetupDialogFunc = True
				nParametersUsed = 0
				For i = 0 To UBound(bUseParameter)
					nParametersUsed = nParametersUsed + IIf(bUseParameter(i), 1, 0)
				Next
				If (sFormat = "ADS" And nParametersUsed = 9) Then
					MsgBox("Only up to 9 parameters are currently supported, please exclude another parameter first.","Warning")
				ElseIf (sFormat = "AWR" And nParametersUsed = 7) Then
					MsgBox("Only up to 7 parameters are supported by AWR, please exclude another parameter first.","Warning")
				Else
					bUseParameter(GetParameterIndexFromName(DlgText("ExcludeLB"))) = True
					ReDim sIncludeList(UBound(bUseParameter))
					ReDim sExcludeList(UBound(bUseParameter))
					For i = 0 To UBound(bUseParameter)
						If bUseParameter(i) Then
							sIncludeList(i) = sParameterNames(i)
						Else
							sExcludeList(i) = sParameterNames(i)
						End If
					Next
					DlgListBoxArray("IncludeLB", sIncludeList)
					DlgListBoxArray("ExcludeLB", sExcludeList)
				End If
			Case "ExcludeAllPB"
				ParameterSetupDialogFunc = True
				ReDim sIncludeList(UBound(bUseParameter))
				ReDim sExcludeList(UBound(bUseParameter))
				For i = 0 To UBound(bUseParameter)
					bUseParameter(i) = False
					sIncludeList(i) = ""
					sExcludeList(i) = sParameterNames(i)
				Next
				DlgListBoxArray("IncludeLB", sIncludeList)
				DlgListBoxArray("ExcludeLB", sExcludeList)
			Case "IncludeAllPB"
				' Include as many parameters as allowed
				ParameterSetupDialogFunc = True
				ReDim sIncludeList(UBound(bUseParameter))
				ReDim sExcludeList(UBound(bUseParameter))
				For i = 0 To UBound(bUseParameter)
					If ((sFormat = "AWR") And (i>6) _
						Or (sFormat = "ADS") And (i>8)) Then
						bUseParameter(i) = False
						sIncludeList(i) = ""
						sExcludeList(i) = sParameterNames(i)
					Else
						bUseParameter(i) = True
						sIncludeList(i) = sParameterNames(i)
						sExcludeList(i) = ""
					End If
				Next
				DlgListBoxArray("IncludeLB", sIncludeList)
				DlgListBoxArray("ExcludeLB", sExcludeList)
			End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : ParameterSetupDialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select

End Function

Function GetParameterIndexFromName(sParameterName As String) As Long

	If sParameterName = "" Then
		GetParameterIndexFromName = 0
	Else
		For GetParameterIndexFromName = 0 To UBound(sParameterNames)
			If sParameterNames(GetParameterIndexFromName) = sParameterName Then Exit For
		Next
	End If

End Function

Sub Main_ADS(vUseParameter As Variant, bSortData As Boolean, nFrequencySamples As Long, dImpedance As Double, Optional sProgressLabel As String)

	Dim i As Long, j As Long, k As Long, m As Long, N As Long
	Dim nParameters As Long, nIndependentParameters, nParameterSets As Long, nPorts As Long
	Dim sListOfRunIDs() As String, sLastResultID As String
	Dim iMap As Long, iTemp As Long, dParameterMap() As Double, oSortedMap As Object, oTempSort As Object, nSubStart As Long, nSubEnd As Long ' used to sort parameter monotonically
	Dim bDoneSorting As Boolean, bIsFirstParameter As Boolean
	Dim oParameterValues() As Object

	Dim sOutputFileName As String
	Dim iOutputFile As Integer

	Dim sTreeResultNames() As String
	Dim oSParameterData() As Object
	Dim vParameterNames As Variant, vParameterValues As Variant
	Dim vPreviousParameterNames As Variant, vPreviousParameterValues As Variant
	Dim dCurrentParameterValue As Double

	sOutputFileName = GetFilePath("sparameter.mdf", "mdf", GetProjectPath("Result"), "Save GMDIF File", 3)

	nPorts = Solver.GetNumberOfPorts
	If (nPorts > 6) Then
		ReportError("Only up to 6 ports are currently supported, aborting.")
	End If

	ReDim oSParameterData(nPorts-1,nPorts-1)
	' prepare parametric data
	ReDim sTreeResultNames(nPorts-1, nPorts-1)
	For i = 0 To nPorts-1
		For j = 0 To nPorts-1
			sTreeResultNames(j,i) = "1D Results\S-Parameters\S"+CStr(j+1)+","+CStr(i+1)
			If Not Resulttree.DoesTreeItemExist(sTreeResultNames(j,i)) Then
				ReportError("Could not find result for S"+CStr(j+1)+","+CStr(i+1)+". The full S-Parameter matrix is needed for the export.")
			End If
		Next
	Next

	sListOfRunIDs = GetListOfRunIDs_LIB(Resulttree.GetFirstChildName("1D Results\S-Parameters"))
	sLastResultID = GetLastResultID()

	nParameterSets = UBound(sListOfRunIDs) ' one of the IDs is for the current, non-parametric result
	If nParameterSets = 0 Then ReportError("No parametric S-parameter data found.")

	For i = 0 To nParameterSets ' one additional to check because of current result
		If GetParameterCombination(sListOfRunIDs(i), vParameterNames, vParameterValues) Then Exit For
	Next
	nParameters = UBound(vParameterNames)+1

	iOutputFile = OpenBufferedFile_LIB(sOutputFileName, "Output")

	BufferedFileWriteLine_LIB(iOutputFile, "! " + CStr(Now))
	BufferedFileWriteLine_LIB(iOutputFile, "! Generated by CST Studio Suite" + vbNewLine)

	iMap = 0
	ReDim dParameterMap(nParameterSets)
	For i = 0 To nParameterSets
		If sListOfRunIDs(i) <> sLastResultID Then
			dParameterMap(iMap) = i
			iMap = iMap + 1
		End If
	Next
	If iMap = 0 Then
		ReportError("No suitable parameter sets detected.")
	End If
	' Trim map as needed
	ReDim Preserve dParameterMap(iMap - 1)
	nParameterSets = UBound(dParameterMap)

	ReDim oParameterValues(nParameters-1)
	For j = 0 To nParameters-1
		Set oParameterValues(j) = Result1D("")
	Next

	If bSortData Then
		' Now start the sorting/mapping, ascending order
		Set oSortedMap = Result1D("")
		' First parameter gets sorted everywhere
		bIsFirstParameter = True
		For j = 0 To UBound(vParameterNames)
			While Not vUseParameter(j)
				j = j + 1
				If (j = UBound(vParameterNames) + 1) Then Exit For
			Wend
			' Use fast sorting option of Result1D object - set up with current level parameter values as X axis, iMap index as Y axis.
			' Then go up level by level for each selected parameter and resort for subrange where previous level parameter values are the same
			' Determine first subinterval
			If bIsFirstParameter Then
				nSubStart = 0
				nSubEnd = nParameterSets
			Else
				nSubStart = 0
				For iMap = nSubStart To nParameterSets
					If oSortedMap.GetX(iMap) <> oSortedMap.GetX(nSubStart) Then Exit For
				Next
				nSubEnd = iMap - 1
			End If
			' ReportInformation("Parameter " & vParameterNames(j) & ": " & CStr(nSubStart) & "..." & CStr(nSubEnd))
			Do
				Set oTempSort = Result1D("")
				For iMap = nSubStart To nSubEnd
					GetParameterCombination(sListOfRunIDs(dParameterMap(iMap)), vParameterNames, vParameterValues)
					oTempSort.AppendXY(vParameterValues(j), dParameterMap(iMap))
				Next
				' Sort subinterval
				oTempSort.SortByX
				If bIsFirstParameter Then
					Set oSortedMap = oTempSort
				Else
					' Copy sorted subinterval back to main list
					For iMap = nSubStart To nSubEnd
						oSortedMap.SetXYDouble(iMap, oTempSort.GetX(iMap - nSubStart), oTempSort.GetY(iMap - nSubStart))
					Next
				End If
				' Find next subinterval
				If bIsFirstParameter Then
					bIsFirstParameter = False ' done here
				ElseIf nSubEnd = nParameterSets Then
					nSubEnd = nSubEnd + 1 ' also done
				Else
					nSubStart = nSubEnd + 1
					For iMap = nSubStart To nParameterSets
						If oSortedMap.GetX(iMap) <> oSortedMap.GetX(nSubStart) Then Exit For
					Next
					nSubEnd = iMap - 1
				End If
				' ReportInformation("Parameter " & vParameterNames(j) & ": " & CStr(nSubStart) & "..." & CStr(nSubEnd))
			Loop While (nSubEnd <= nParameterSets)
			' Copy sorted list back to dParameterMap
			dParameterMap = oSortedMap.GetArray("y")
		Next
	End If

	' For all result entries
	For iMap = 0 To nParameterSets
		If ((sProgressLabel <> "") And (nParameterSets > 0)) Then DlgText(sProgressLabel, Format(iMap / nParameterSets * 100, "0.00") & "%")
		' Write down parameter values, format: VAR <name>(real)=<value>
		For j = 0 To nParameters-1
			GetParameterCombination(sListOfRunIDs(dParameterMap(iMap)), vParameterNames, vParameterValues)
			dCurrentParameterValue = CDbl(vParameterValues(j))
			If vUseParameter(j)Then
				BufferedFileWriteLine_LIB(iOutputFile, "VAR "+sParameterNames(j)+"(real)="+ Replace(Format(dCurrentParameterValue, sPrintFormat), ",", "."))
				oParameterValues(j).AppendXY(iMap, dCurrentParameterValue)
			End If
		Next j ' next parameter name
		BufferedFileWriteLine_LIB(iOutputFile, "BEGIN ACDATA")
		' Write % freq(real) S1,1(complex) S2,1(complex) ... SN,1(complex) S1,2(complex) ... SN,2(complex) ... SN,N(complex) Z1(complex) ... ZN(complex)
		BufferedFileWrite_LIB(iOutputFile, "% freq(real) "+vbTab)
		For j = 0 To nPorts-1
			For k = 0 To nPorts-1
				BufferedFileWrite_LIB(iOutputFile, "S"+CStr(k+1)+","+CStr(j+1)+"(complex)"+vbTab+vbTab+vbTab+vbTab+vbTab)
			Next
		Next
		For j = 0 To nPorts-1
			BufferedFileWrite_LIB(iOutputFile, "Z"+CStr(j+1)+"(complex)"+vbTab+vbTab+vbTab+vbTab+vbTab)
		Next
		BufferedFileWriteLine_LIB(iOutputFile, "")

		' Write data in format described above
		' Read S-Parameter data for each parameter set
		For j = 0 To nPorts-1
			For k = 0 To nPorts-1
				Set oSParameterData(k,j) = Resulttree.GetResultFromTreeItem(sTreeResultNames(k, j), sListOfRunIDs(dParameterMap(iMap)))
				oSParameterData(k,j).ResampleTo(oSParameterData(k,j).GetX(0), oSParameterData(k,j).GetX(oSParameterData(k,j).GetN-1),nFrequencySamples)
			Next
		Next

		' Go through all frequency samples
		For j = 0 To nFrequencySamples-1
			' Print frequency value
			BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(0,0).GetX(j)*Units.GetFrequencyUnitToSI(), sPrintFormat), ",", ".") & vbTab)
			' Print S-Parameter values
			For k = 0 To nPorts-1
				For m = 0 To nPorts-1
					BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(m,k).GetYRe(j), sPrintFormat), ",", ".") & vbTab)
					BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(m,k).GetYIm(j), sPrintFormat), ",", ".") & vbTab)
				Next
			Next
			' Print port impedances
			For k = 0 To nPorts-1
				BufferedFileWrite_LIB(iOutputFile, Replace(Format(dImpedance, sPrintFormat), ",", ".")+vbTab+"0"+vbTab)
			Next
			BufferedFileWriteLine_LIB(iOutputFile, "")
		Next
		BufferedFileWriteLine_LIB(iOutputFile, "END ACDATA"+vbNewLine)
		EndOfParameterSetLoop:
	Next iMap ' Next parameter set

	For j = 0 To nParameters-1
		If vUseParameter(j)Then
			oParameterValues(j).SortByX
			AddPlotToTree_LIB(oParameterValues(j), "MDIF export sorting\" & sParameterNames(j), False)
		End If
	Next

	CloseBufferedFile_LIB(iOutputFile)
	Clipboard(sOutputFileName)
	MsgBox("Export completed. Exported file can be found at:" & vbNewLine & vbNewLine & sOutputFileName & vbNewLine & vbNewLine & "(Path has been added to clipboard.)")

End Sub

Sub Main_AWR(vUseParameter As Variant, bSortData As Boolean, nFrequencySamples As Long, dImpedance As Double, Optional sProgressLabel As String)

	Dim i As Long, j As Long, k As Long, m As Long, N As Long
	Dim nParameters As Long, nExportedParameters As Long, nParameterSets As Long, nPorts As Long
	Dim sListOfRunIDs() As String, sLastResultID As String
	Dim iMap As Long, iTemp As Long, dParameterMap() As Double, oSortedMap As Object, oTempSort As Object, nSubStart As Long, nSubEnd As Long ' used to sort parameter monotonically
	Dim bDoneSorting As Boolean, bIsFirstParameter As Boolean
	Dim oParameterValues() As Object

	Dim sOutputFileName As String
	Dim iOutputFile As Integer

	Dim sTreeResultNames() As String
	Dim oSParameterData() As Object
	Dim vParameterNames As Variant, vParameterValues As Variant
	Dim vPreviousParameterNames As Variant, vPreviousParameterValues As Variant
	Dim dCurrentParameterValue As Double

	' do NOT use the same file name as for ADS, since the files types are not compatible!
	sOutputFileName = GetFilePath("sparameter_AWR.mdf", "mdf", GetProjectPath("Result"), "Save GMDIF File", 3)
	If sOutputFileName = "" Then
		ReportInformationToWindow("GMDIF export was canceled.")
		Exit All
	End If

	nPorts = Solver.GetNumberOfPorts

	ReDim oSParameterData(nPorts-1,nPorts-1)
	' prepare parametric data
	ReDim sTreeResultNames(nPorts-1, nPorts-1)
	For i = 0 To nPorts-1
		For j = 0 To nPorts-1
			sTreeResultNames(j,i) = "1D Results\S-Parameters\S"+CStr(j+1)+","+CStr(i+1)
			If Not Resulttree.DoesTreeItemExist(sTreeResultNames(j,i)) Then
				ReportError("Could not find result for S"+CStr(j+1)+","+CStr(i+1)+". The full S-Parameter matrix is needed for the export.")
			End If
		Next
	Next

	sListOfRunIDs = GetListOfRunIDs_LIB(Resulttree.GetFirstChildName("1D Results\S-Parameters"))
	sLastResultID = GetLastResultID()

	nParameterSets = UBound(sListOfRunIDs) ' one of the IDs is for the current, non-parametric result
	If nParameterSets = 0 Then ReportError("No parametric S-parameter data found.")

	For i = 0 To nParameterSets ' one additional to check because of current result
		If GetParameterCombination(sListOfRunIDs(i), vParameterNames, vParameterValues) Then Exit For
	Next
	nParameters = UBound(vParameterNames)+1

	iOutputFile = OpenBufferedFile_LIB(sOutputFileName, "Output")
	BufferedFileWriteLine_LIB(iOutputFile, "! " + CStr(Now))
	BufferedFileWriteLine_LIB(iOutputFile, "! Generated by CST Studio Suite" + vbNewLine)

	iMap = 0
	ReDim dParameterMap(nParameterSets)
	For i = 0 To nParameterSets
		If sListOfRunIDs(i) <> sLastResultID Then
			dParameterMap(iMap) = i
			iMap = iMap + 1
		End If
	Next
	If iMap = 0 Then
		ReportError("No suitable parameter sets detected.")
	End If
	' Trim map as needed
	ReDim Preserve dParameterMap(iMap - 1)
	nParameterSets = UBound(dParameterMap)

	ReDim oParameterValues(nParameters-1)
	For j = 0 To nParameters-1
		Set oParameterValues(j) = Result1D("")
	Next

	If bSortData Then
		' Now start the sorting/mapping, ascending order
		Set oSortedMap = Result1D("")
		' First parameter gets sorted everywhere
		bIsFirstParameter = True
		For j = 0 To UBound(vParameterNames)
			While Not vUseParameter(j)
				j = j + 1
				If (j = UBound(vParameterNames) + 1) Then Exit For
			Wend
			' Use fast sorting option of Result1D object - set up with current level parameter values as X axis, iMap index as Y axis.
			' Then go up level by level for each selected parameter and resort for subrange where previous level parameter values are the same
			' Determine first subinterval
			If bIsFirstParameter Then
				nSubStart = 0
				nSubEnd = nParameterSets
			Else
				nSubStart = 0
				For iMap = nSubStart To nParameterSets
					If oSortedMap.GetX(iMap) <> oSortedMap.GetX(nSubStart) Then Exit For
				Next
				nSubEnd = iMap - 1
			End If
			' ReportInformation("Parameter " & vParameterNames(j) & ": " & CStr(nSubStart) & "..." & CStr(nSubEnd))
			Do
				Set oTempSort = Result1D("")
				For iMap = nSubStart To nSubEnd
					GetParameterCombination(sListOfRunIDs(dParameterMap(iMap)), vParameterNames, vParameterValues)
					oTempSort.AppendXY(vParameterValues(j), dParameterMap(iMap))
				Next
				' Sort subinterval
				oTempSort.SortByX
				If bIsFirstParameter Then
					Set oSortedMap = oTempSort
				Else
					' Copy sorted subinterval back to main list
					For iMap = nSubStart To nSubEnd
						oSortedMap.SetXYDouble(iMap, oTempSort.GetX(iMap - nSubStart), oTempSort.GetY(iMap - nSubStart))
					Next
				End If
				' Find next subinterval
				If bIsFirstParameter Then
					bIsFirstParameter = False ' done here
				ElseIf nSubEnd = nParameterSets Then
					nSubEnd = nSubEnd + 1 ' also done
				Else
					nSubStart = nSubEnd + 1
					For iMap = nSubStart To nParameterSets
						If oSortedMap.GetX(iMap) <> oSortedMap.GetX(nSubStart) Then Exit For
					Next
					nSubEnd = iMap - 1
				End If
				' ReportInformation("Parameter " & vParameterNames(j) & ": " & CStr(nSubStart) & "..." & CStr(nSubEnd))
			Loop While (nSubEnd <= nParameterSets)
			' Copy sorted list back to dParameterMap
			dParameterMap = oSortedMap.GetArray("y")
		Next
	End If

	' For all result entries
	For iMap = 0 To nParameterSets
		If ((sProgressLabel <> "") And (nParameterSets > 0)) Then DlgText(sProgressLabel, Format(iMap / nParameterSets * 100, "0.00") & "%")
		' Write down parameter values, format: VAR <name>=<value>
		For j = 0 To nParameters-1
			GetParameterCombination(sListOfRunIDs(dParameterMap(iMap)), vParameterNames, vParameterValues)
			dCurrentParameterValue = CDbl(vParameterValues(j))

			If vUseParameter(j) Then
				BufferedFileWriteLine_LIB(iOutputFile, "VAR "+sParameterNames(j)+"="+Replace(Format(dCurrentParameterValue, sPrintFormat), ",", "."))
				oParameterValues(j).AppendXY(iMap, dCurrentParameterValue)
			End If
		Next j ' next parameter name

		BufferedFileWriteLine_LIB(iOutputFile, "BEGIN ACDATA")
		BufferedFileWriteLine_LIB(iOutputFile, "# GHz S RI R " + Replace(Format(dImpedance, "0.00"), ",", "."))

		' Please note format: S11-S21-S12-S22, in agreement with format for s2p format
		' In contrast, snp format is S11-S12-S13-... for n>2; AWR GMDIF documentation does not mention if it follows the s2p or snp convention for n>2; assuming s2p convention for now

		' Write % F n11x n11y n21x n21y n31x n31y ... n12x n12y n22x n22y n32x n32y ... nNNx nNNy
		BufferedFileWrite_LIB(iOutputFile, "% F ")
		For j = 0 To nPorts-1
			For k = 0 To nPorts-1
				BufferedFileWrite_LIB(iOutputFile, "n"+CStr(k+1)+CStr(j+1)+"x "+vbTab+vbTab+"n"+CStr(k+1)+CStr(j+1)+"y "+vbTab+vbTab)
			Next
		Next

		BufferedFileWriteLine_LIB(iOutputFile, "")
		' Write data in format desribed above
		' Read S-Parameter data for each parameter set
		For k = 0 To nPorts-1
			For j = 0 To nPorts-1
				Set oSParameterData(k,j) = Resulttree.GetResultFromTreeItem(sTreeResultNames(k, j), sListOfRunIDs(dParameterMap(iMap)))
				oSParameterData(k,j).ResampleTo(oSParameterData(k,j).GetX(0), oSParameterData(k,j).GetX(oSParameterData(k,j).GetN-1),nFrequencySamples)
			Next
		Next

		' Go through all frequency samples
		For j = 0 To nFrequencySamples-1
			' Print frequency value
			BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(0,0).GetX(j)*Units.GetFrequencyUnitToSI()/1e9, sPrintFormat), ",", ".") & vbTab)
			' Print S-Parameter values
			For m = 0 To nPorts-1
				For k = 0 To nPorts-1
					BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(m,k).GetYRe(j), sPrintFormat), ",", ".") & vbTab)
					BufferedFileWrite_LIB(iOutputFile, Replace(Format(oSParameterData(m,k).GetYIm(j), sPrintFormat), ",", ".") & vbTab)
				Next
			Next
			BufferedFileWriteLine_LIB(iOutputFile, "")
		Next
		BufferedFileWriteLine_LIB(iOutputFile, "END ACDATA" & vbNewLine)
		'EndOfParameterSetLoop:
	Next iMap ' Next parameter set

	For j = 0 To nParameters-1
		If vUseParameter(j)Then
			oParameterValues(j).SortByX
			AddPlotToTree_LIB(oParameterValues(j), "MDIF export sorting\" & sParameterNames(j), False)
		End If
	Next

	CloseBufferedFile_LIB(iOutputFile)
	Clipboard(sOutputFileName)
	MsgBox("Export completed. Exported file can be found at:" & vbNewLine & vbNewLine & sOutputFileName & vbNewLine & vbNewLine & "(Path has been added to clipboard.)")

End Sub
