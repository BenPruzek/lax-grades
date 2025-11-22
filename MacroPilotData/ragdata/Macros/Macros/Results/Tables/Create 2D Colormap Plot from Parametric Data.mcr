Option Explicit

'#include "vba_globals_all.lib"
'#include "template_conversions.lib"
'#include "template_results.lib"

' =============================================================================================================================================================
' This macro takes parametric 0D or 1D data (table or parametric results) and plots the 0D values over two parameters (one parameter for 1D table).
' The two parameters represent x and y axes. In case of 1D data, the x axis of the data will be the x axis of the 2D plot
'
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' =============================================================================================================================================================
' History of Changes
' ------------------
' 05-Oct-2020 fsr: Additional bugfixes
' 01-Oct-2020 fsr: Parameter sorting not necessarily consistent between 3D, Schematic, and results. Introduced index mapping to address this.
' 26-May-2020 fsr: Re-established compatibility with Schematic. Results are now shown in Schematic tree while in Schematic.
' 11-Dec-2019 thn: Fix index error
' 25-Sep-2019 fsr: Removed Table code that was already commented out
' 05-Mar-2019 fsr: Changed data access from Table object to ResultTree entry
' 13-Mar-2018 fsr: Added error message if less than two parameters defined in project
' 17-Jun-2016 fsr: Added DS/Schematic compatibility
' 04-Jan-2016 fsr: Complex data now supported; resample input data if number of x-samples varies in original data; user may specify result name manually
' 11-Dec-2015 fsr: Added support for parametric tree entries
' 10-Dec-2015 fsr: Small bugfix; added resampling option; drop-down list to select result
' 03-Dec-2015 fsr: Added support for 1D tables
' 15-May-2014 fsr: Initial version
' =============================================================================================================================================================

Dim nNumberOfParameters As Long
Dim sParameterList() As String

Dim aRTemplateName() As String
Dim aRTemplateType() As String
Dim aRTemplateDefault() As String

Dim sResultFolder As String

Sub Main()

	Dim i As Long
	Dim sListOfSelectionSettings As String, sListOfSelectionTypeSettings As String
	Dim nOutputNameBoxSizeAdjustment As Integer

'	If (Left(GetApplicationName, 2) = "DS") Then
'		MsgBox("This macro is currently only functional within the '3D' tab.")
'		Exit All
'	End If

	sResultFolder = IIf(Left(GetApplicationName, 2) = "DS", "Results\Colormap Plots\", "2D/3D Results\Colormap Plots\")
	nOutputNameBoxSizeAdjustment = 7 * (Len(sResultFolder)) ' 1 average character is approximately 8 units wide, 7 seems to work better here due to narrow characters

	FillResultList_LIB(aRTemplateName, aRTemplateType, "EMPTY", "ALL", "ALL", "ALL", sListOfSelectionSettings, sListOfSelectionTypeSettings)

	Begin Dialog UserDialog 720,196,"2D Colormap Plot from Parametric Data",.DialogFunc ' %GRID:10,7,1,1
		Text 20,14,160,14,"Create colormap plot from:",.Text1
		DropListBox 20,35,680,128,aRTemplateName(),.aRTemplateA
		Text 20,70,320,14,"Store result under: " & sResultFolder,.Text6
		TextBox 150+nOutputNameBoxSizeAdjustment,63,550-nOutputNameBoxSizeAdjustment,21,.OutputNameT
		Text 20,105,50,14,"x-axis:",.Text2
		DropListBox 70,98,170,170,sParameterList(),.XParameterDLB
		CheckBox 250,105,180,14,"Resample/interpolate to",.ResampleXCB
		TextBox 440,98,90,21,.XSamplesT
		Text 540,105,60,14,"samples.",.Text4
		Text 20,133,40,14,"y-axis:",.Text3
		DropListBox 70,126,170,156,sParameterList(),.YParameterDLB
		CheckBox 250,133,180,14,"Resample/interpolate to",.ResampleYCB
		TextBox 440,126,90,21,.YSamplesT
		Text 540,133,60,14,"samples.",.Text5
		Text 20,168,470,14,"",.OutputLabel
		OKButton 510,161,90,21
		CancelButton 610,161,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.OutputNameT = "[Auto]"

	dlg.XParameterDLB = 0
	dlg.YParameterDLB = 1

	dlg.ResampleXCB = 0
	dlg.XSamplesT = "21"
	dlg.ResampleYCB = 0
	dlg.YSamplesT = "21"

	If (Dialog(dlg) = 0) Then
		Exit All
	End If

End Sub

Function CreateColormapPlot(sSelectedResultName As String, sSelectedResultType As String, sOutputName As String, nXSamples As Long, nYSamples As Long) As Long

	' Creates a Colormap plot for sSelectedResultName of sSelectedResultType
	' If nXSamples or nYSamples are > 0, the data will be resampled accordingly prior to plotting

	Dim dTableDataList() As Double, nNumberOfDataItems As Long
	Dim oXValues As Object, oYValues As Object ' Use objects first because they have built-in sorting; then switch to double arrays
	Dim dXValues() As Double, dYValues() As Double, dZReValues() As Double, dZImValues() As Double ' use objects for list as the object provides a quick SortByX
	Dim nXParameterIndex As Long, nYParameterIndex As Long, nXParameterIndexData As Long, nYParameterIndexData As Long, dSamplingDeviationX As Double, dSamplingDeviationY As Double
	Dim dXMin As Double, dXMax As Double, dYMin As Double, dYMax As Double
	Dim i As Long, j As Long, k As Long
	Dim sXLabel As String, sYLabel As String
	Dim oTempObject As Object, dTempArrayRe() As Double, dTempArrayIm() As Double
	Dim sResultTreePath As String
	Dim sListOfRunIDs() As String, sListOfParameterNames As Variant, sListOfParameterValues As Variant
	Dim colourfile As String
	Dim bComplex As Boolean ' is the data to be displayed complex?
	Dim dLogScaleFactor As Double
	Dim nXSamplesReference As Long, nYSamplesReference As Long
	Dim sTreeLocation As String
	Dim oTreeObject As Object, sTreeResultFolder As String
	Dim bUseThisRunID As Boolean

	Dim dStartTime As Double

	If (Left(GetApplicationName, 2) = "DS") Then
		Set oTreeObject = DSResultTree
		sTreeResultFolder = "Results\"
	Else
		Set oTreeObject = Resulttree
		sTreeResultFolder = "2D/3D Results\Colormap Plots\"
	End If

	If (sSelectedResultType = "0D") And (nNumberOfParameters < 2) Then
		ReportError("This macro requires at least 2 parameters to be varied.")
	ElseIf (sSelectedResultType = "0D") And (nNumberOfParameters > 2) Then
		ReportWarning("More than 2 parameters detected. The macro will only work properly if only 2 parameters are varied. The rest need to remain constant.")
	End If

	colourfile = IIf(sOutputName = "[Auto]", NoForbiddenFilenameCharacters(sSelectedResultName), NoForbiddenFilenameCharacters(sOutputName)) & ".dat"
	sResultTreePath = GetTreePathFromResultName(sSelectedResultName, sSelectedResultType)

	nXParameterIndex = DlgValue("XParameterDLB")
	nYParameterIndex = DlgValue("YParameterDLB")

	' Set up x and y axes
	Set oXValues = Result1D("")
	Set oYValues = Result1D("")

	sYLabel = sParameterList(nYParameterIndex)

	sListOfRunIDs = GetListOfRunIDs_LIB(sResultTreePath)
	nNumberOfDataItems = UBound(sListOfRunIDs) + 1
	If (InStr(UCase(sSelectedResultType), "0D")>0) Then
		sXLabel = sParameterList(nXParameterIndex)
		For i = 0 To nNumberOfDataItems-1
			If (Left(GetApplicationName, 2) = "DS") Then
				bUseThisRunID = DS.GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
			Else
				bUseThisRunID = GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
			End If
			If bUseThisRunID Then
				' Order of parameters in project is not necessarily identical to order of parameters in results. Make sure they are properly mapped.
				For j = 0 To UBound(sListOfParameterNames)
					If (sListOfParameterNames(j) = sParameterList(nXParameterIndex)) Then nXParameterIndexData = j
					If (sListOfParameterNames(j) = sParameterList(nYParameterIndex)) Then nYParameterIndexData = j
				Next

				' Build list of unique x and y values in preparation for converting data to matrix
				If i = 0 Then
					oXValues.AppendXY(sListOfParameterValues(nXParameterIndexData), 0)
					oYValues.AppendXY(sListOfParameterValues(nYParameterIndexData), 0)
				End If
				For j = 0 To oXValues.GetN()-1
					If sListOfParameterValues(nXParameterIndexData) = oXValues.GetX(j) Then Exit For
				Next
				If j = oXValues.GetN() Then ' For loop finished -> unique value
					oXValues.AppendXY(sListOfParameterValues(nXParameterIndexData), 0)
				End If
				For k = 0 To oYValues.GetN()-1
					If sListOfParameterValues(nYParameterIndexData) = oYValues.GetX(k) Then Exit For
				Next
				If k = oYValues.GetN() Then ' For loop finished  -> unique value
					oYValues.AppendXY(sListOfParameterValues(nYParameterIndexData), 0)
				End If
			Else
				' Do nothing, skip to next
			End If
		Next
	Else
		If (Left(GetApplicationName, 2) = "DS") Then
			Set oXValues = DSResultTree.GetResultFromTreeItem(sResultTreePath, sListOfRunIDs(nNumberOfDataItems-1))
		Else
			Set oXValues = Resulttree.GetResultFromTreeItem(sResultTreePath, sListOfRunIDs(nNumberOfDataItems-1))
		End If

		sXLabel = oXValues.GetXLabel()
		For i = 0 To nNumberOfDataItems - 1
			If (Left(GetApplicationName, 2) = "DS") Then
				bUseThisRunID = DS.GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
			Else
				bUseThisRunID = GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
			End If
			If bUseThisRunID Then
				' Order of parameters in project is not necessarily identical to order of parameters in results. Make sure they are properly mapped.
				For j = 0 To UBound(sListOfParameterNames)
					If (sListOfParameterNames(j) = sParameterList(nXParameterIndex)) Then nXParameterIndexData = j
					If (sListOfParameterNames(j) = sParameterList(nYParameterIndex)) Then nYParameterIndexData = j
				Next

				If i = 0 Then
					oYValues.AppendXY(sListOfParameterValues(nYParameterIndexData), 0)
				Else
					For k = 0 To oYValues.GetN()-1
						If sListOfParameterValues(nYParameterIndexData) = oYValues.GetX(k) Then Exit For
					Next
					If k = oYValues.GetN() Then ' For loop finished  -> unique value
						oYValues.AppendXY(sListOfParameterValues(nYParameterIndexData), 0)
					End If
				End If
			Else
				' Do nothing, skip to next
			End If
		Next
	End If

	oXValues.SortByX
	oYValues.SortByX

	dXValues = oXValues.GetArray("x")
	dYValues = oYValues.GetArray("x")

	dXMin = dXValues(0)
	dXMax = dXValues(UBound(dXValues))
	dYMin = dYValues(0)
	dYMax = dYValues(UBound(dYValues))

	nXSamplesReference = UBound(dXValues) + 1
	nYSamplesReference = UBound(dYValues) + 1

	If (nXSamplesReference < 3) Then ReportError("Colormap plot requires at least 3 data points on the x axis.")
	If (nXSamplesReference < 3) Then ReportError("Colormap plot requires at least 3 data points on the y axis.")

	' Fill data matrices
	ReDim dZReValues(oYValues.GetN()-1, oXValues.GetN()-1)
	ReDim dZImValues(oYValues.GetN()-1, oXValues.GetN()-1)
	dStartTime = Timer()
	For i = 0 To nNumberOfDataItems-1
		If ((i-1) Mod 10 = 0) Then DlgText("OutputLabel", "Processing data: " & Cstr(i) & "/" & CStr(nNumberOfDataItems))
		' Find the right coordinates in the matrix
		Select Case sSelectedResultType
			Case "0D", "M0D", "real0D", "TABLE0d real"
				bComplex = False
				If (Left(GetApplicationName, 2) = "DS") Then
					bUseThisRunID = DS.GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				Else
					bUseThisRunID = GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				End If
				If bUseThisRunID Then
					' Order of parameters in project is not necessarily identical to order of parameters in results. Make sure they are properly mapped.
					For j = 0 To UBound(sListOfParameterNames)
						If (sListOfParameterNames(j) = sParameterList(nXParameterIndex)) Then nXParameterIndexData = j
						If (sListOfParameterNames(j) = sParameterList(nYParameterIndex)) Then nYParameterIndexData = j
					Next
					For j = 0 To oXValues.GetN()-1
						If oXValues.GetX(j) = sListOfParameterValues(nXParameterIndexData) Then Exit For
					Next
					For k = 0 To oYValues.GetN()-1
						If oYValues.GetX(k) = sListOfParameterValues(nYParameterIndexData) Then Exit For
					Next
					dZReValues(k, j) = GetResultByRunID_LIB(sListOfRunIDs(i), sSelectedResultName, sSelectedResultType, "0D")
					dZImValues(k, j) = 0
				Else
					' Do nothing, skip to next
				End If
			Case "1D", "M1D", "real"
				bComplex = False
				If (Left(GetApplicationName, 2) = "DS") Then
					bUseThisRunID = DS.GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				Else
					bUseThisRunID = GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				End If
				If bUseThisRunID Then
					' Order of parameters in project is not necessarily identical to order of parameters in results. Make sure they are properly mapped.
					For j = 0 To UBound(sListOfParameterNames)
						If (sListOfParameterNames(j) = sParameterList(nXParameterIndex)) Then nXParameterIndexData = j
						If (sListOfParameterNames(j) = sParameterList(nYParameterIndex)) Then nYParameterIndexData = j
					Next
					For k = 0 To oYValues.GetN()-1
						If oYValues.GetX(k) = sListOfParameterValues(nYParameterIndexData) Then Exit For
					Next
					Set oTempObject = GetResultByRunID_LIB(sListOfRunIDs(i), sSelectedResultName, sSelectedResultType, "1D")
					oTempObject.SortByX ' this is important
					' Resample if number of samples in x varies for some reason
					If oTempObject.GetN() <> nXSamples Then
						oTempObject.ResampleTo(dXMin, dXMax, nXSamplesReference)
					End If
					dTempArrayRe = oTempObject.GetArray("y")
					For j = 0 To UBound(dXValues)
						dZReValues(k, j) = dTempArrayRe(j)
						dZImValues(k, j) = 0
					Next
				Else
					' Do nothing, skip to next
				End If
			Case "1DC", "M1DC", "complex"
				bComplex = True
				If (Left(GetApplicationName, 2) = "DS") Then
					bUseThisRunID = DS.GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				Else
					bUseThisRunID = GetParameterCombination(sListOfRunIDs(i), sListOfParameterNames, sListOfParameterValues)
				End If
				If bUseThisRunID Then
					' Order of parameters in project is not necessarily identical to order of parameters in results. Make sure they are properly mapped.
					For j = 0 To UBound(sListOfParameterNames)
						If (sListOfParameterNames(j) = sParameterList(nXParameterIndex)) Then nXParameterIndexData = j
						If (sListOfParameterNames(j) = sParameterList(nYParameterIndex)) Then nYParameterIndexData = j
					Next
					For k = 0 To oYValues.GetN()-1
						If oYValues.GetX(k) = sListOfParameterValues(nYParameterIndexData) Then Exit For
					Next
					Set oTempObject = GetResultByRunID_LIB(sListOfRunIDs(i), sSelectedResultName, sSelectedResultType, "1DC")
					oTempObject.SortByX ' this is important
					' Resample if number of samples in x varies for some reason
					If oTempObject.GetN() <> nXSamples Then
						oTempObject.ResampleTo(dXMin, dXMax, nXSamplesReference)
					End If
					dTempArrayRe = oTempObject.GetArray("yre")
					dTempArrayIm = oTempObject.GetArray("yim")
					For j = 0 To UBound(dXValues)
						dZReValues(k, j) = dTempArrayRe(j)
						dZImValues(k, j) = dTempArrayIm(j)
					Next
					dLogScaleFactor = oTempObject.GetLogarithmicFactor()
				Else
					' Do nothing, skip to next
				End If
			Case Else
				MsgBox("Data type '" & sSelectedResultType & "' is currently not supported.")
				Exit All
		End Select
	Next
	' DS.ReportInformationToWindow(Timer() - dStartTime)

	' If resampling in either direction is selected, or if data is not equidistant, resample data
	dSamplingDeviationX = CalculateDeviationFromEquidistantSampling(oXValues)
	dSamplingDeviationY = CalculateDeviationFromEquidistantSampling(oYValues)
	If ((dSamplingDeviationX > 0.1) And (nXSamples = 0)) Then
		nXSamples = 101
		ReportWarning("Colormap plot from parametric data: Data is not equidistantly sampled in x direction and will be resampled automatically using 101 samples.")
	End If
	If ((dSamplingDeviationY > 0.1) And (nYSamples = 0)) Then
		nYSamples = 21
		ReportWarning("Colormap plot from parametric data: Data is not equidistantly sampled in y direction and will be resampled automatically using 21 samples.")
	End If

	' Re-check dSamplingVariation and set to zero if it will be resampled
	dSamplingDeviationX = IIf(nXSamples > 0, 0, dSamplingDeviationX)
	dSamplingDeviationY = IIf(nYSamples > 0, 0, dSamplingDeviationY)

	If ((dSamplingDeviationX > 0.01) Or (dSamplingDeviationY > 0.01)) Then
		ReportWarning("Colormap plot from parametric data: Samples deviate from equidistant sampling by more than 1% (but less than 10%). This will negatively affect the 2D plot accuracy. Please consider to activate manual resampling.")
	End If

	If ((nXSamples > 0) Or (nYSamples > 0)) Then
		Resample2DData(dXValues, dYValues, dZReValues, dZImValues, nXSamples, nYSamples)
	End If

	nXSamples = UBound(dXValues) + 1
	nYSamples = UBound(dYValues) + 1

	' Labels must not be empty
	If sXLabel = "" Then sXLabel = "x-axis"
	If sYLabel = "" Then sYLabel = "y-axis"

	sTreeLocation = sTreeResultFolder & IIf(sOutputName = "[Auto]", Mid(sSelectedResultName, InStrRev(sSelectedResultName, "\")+1), sOutputName)

	'====================================================================================
	Dim o As Object

	If (Left(GetApplicationName, 2) = "DS") Then
		Set o = DS.Result2D("")
	Else
		Set o = Result2D("")
	End If

	If bComplex Then
		o.InitializeComplex(nXSamples, nYSamples)
		o.SetLogarithmicFactor(dLogScaleFactor)
	Else
		o.Initialize(nXSamples, nYSamples)
	End If

	o.SetTitle(sSelectedResultName)
	o.SetXLabel(sXLabel)
	o.SetXUnit("1")
	o.SetYLabel(sYLabel)
	o.SetYUnit("1")
	o.SetXmin(dXMin)
	o.SetXmax(dXMax)
	o.SetYmin(dYMin)
	o.SetYmax(dYMax)

	If bComplex Then
		For k = 0 To nYSamples - 1
			For j = 0 To nXSamples - 1
				o.SetValueComplex(j, k, dZReValues(k, j), dZImValues(k, j))
			Next
		Next
	Else
		For k = 0 To nYSamples - 1
			For j = 0 To nXSamples - 1
				o.SetValue(j, k, dZReValues(k, j))
			Next
		Next
	End If

	o.Save(GetProjectPath("Result") + "\" + colourfile)
	o.AddToTree(sTreeLocation)
	'====================================================================================

	If (Left(GetApplicationName, 2) = "DS") Then
		DS.SelectTreeItem(sTreeLocation)
	Else
		SelectTreeItem(sTreeLocation)
	End If

End Function

Function CalculateDeviationFromEquidistantSampling(oInputObject As Object) As Double

	' This function calculates the normalized maximum deviation from an equidistant sampling step. Works with Result1D and Result1DComplex

	Dim dAverageStep As Double, i As Long, dMaxDeviation As Double

	If (oInputObject.GetN() <= 2) Then
		CalculateDeviationFromEquidistantSampling = 0
		Exit Function
	End If

	oInputObject.SortByX
	dAverageStep = (oInputObject.GetX(oInputObject.GetN()-1)-oInputObject.GetX(0))/(oInputObject.GetN()-1)
	If (dAverageStep = 0) Then
		ReportWarning("Equidistant check: First and last point have the value, exiting.")
		CalculateDeviationFromEquidistantSampling = 0
		Exit Function
	End If
	dMaxDeviation = (dAverageStep-(oInputObject.GetX(1)-oInputObject.GetX(0)))/dAverageStep
	For i = 1 To oInputObject.GetN()-2
		If ((oInputObject.GetX(i+1)-oInputObject.GetX(i))/dAverageStep > dMaxDeviation) Then
			dMaxDeviation = (dAverageStep-(oInputObject.GetX(i+1)-oInputObject.GetX(i)))/dAverageStep
		End If
	Next

	CalculateDeviationFromEquidistantSampling = dMaxDeviation

End Function

Function Resample2DData(ByRef dXValues() As Double, ByRef dYValues() As Double, ByRef dZReValues() As Double, ByRef dZImValues() As Double, nXSamples As Long, nYSamples As Long) As Long

	' dXValues() is a list containing the x data (N samples)
	' dYValues() is a list containing the y data (M samples)
	' dZReValues() and dZImValues() are 2D arrays MxN
	' Data will be resampled to nXSamples x nYSamples
	' If nXSamples or nYSamples is 0, M or N remains unchanged

	Dim nM As Long, nN As Long
	Dim oDataRows() As Object, oDataColumns() As Object
	Dim dTempArrayRe() As Double, dTempArrayIm() As Double
	Dim dLocalXValues() As Double, dLocalYValues() As Double, dLocalZReValues() As Double, dLocalZImValues() As Double ' needed because of redimming
	Dim i As Long, j As Long

	If (UBound(dZReValues, 1) <> UBound(dZImValues, 1)) Or (UBound(dZReValues, 2) <> UBound(dZImValues, 2)) Then ReportError("Error: Real and imaginary data need to have the same number of samples.")
	nM = UBound(dZReValues, 1) + 1
	nN = UBound(dZReValues, 2) + 1

	dLocalXValues = dXValues
	dLocalYValues = dYValues
	dLocalZReValues = dZReValues
	dLocalZImValues = dZImValues

	If (nXSamples > 0) Then
		' Formulate 2D matrix as an array of Result1Ds, use the built-in resampling function for Result1Ds
		ReDim oDataRows(nM-1)
		For i = 0 To nM - 1
			Set oDataRows(i) = Result1DComplex("")
			oDataRows(i).Initialize(nN)
			ReDim dTempArrayRe(nN - 1)
			ReDim dTempArrayIm(nN - 1)
			For j = 0 To nN - 1
				dTempArrayRe(j) = dLocalZReValues(i, j)
				dTempArrayIm(j) = dLocalZImValues(i, j)
			Next
			oDataRows(i).SetArray(dLocalXValues, "x")
			oDataRows(i).SetArray(dTempArrayRe, "yre")
			oDataRows(i).SetArray(dTempArrayIm, "yim")
			' AddPlotToTree_LIB(oDataRows(i), "ColormapPlotTest\" & CStr(i) & "_BEFORE")
			oDataRows(i).ResampleTo(dLocalXValues(0), dLocalXValues(nN - 1), nXSamples)
			' AddPlotToTree_LIB(oDataRows(i), "ColormapPlotTest\" & CStr(i) & "_AFTER")
		Next
		' Create new dZValues from rows
		ReDim dLocalZReValues(nM - 1, nXSamples - 1)
		ReDim dLocalZImValues(nM - 1, nXSamples - 1)
		For i = 0 To nM - 1
			dTempArrayRe = oDataRows(i).GetArray("yre")
			dTempArrayIm = oDataRows(i).GetArray("yim")
			For j = 0 To nXSamples - 1
				dLocalZReValues(i, j) = dTempArrayRe(j)
				dLocalZImValues(i, j) = dTempArrayIm(j)
			Next
		Next
		nN = nXSamples
		dLocalXValues = oDataRows(0).GetArray("x")
	End If

	If (nYSamples > 0) Then
		' Formulate 2D matrix as an array of Result1Ds, use the built-in resampling function for Result1Ds
		ReDim oDataColumns(nN - 1)
		For i = 0 To nN - 1
			Set oDataColumns(i) = Result1DComplex("")
			oDataColumns(i).Initialize(nM)
			ReDim dTempArrayRe(nM - 1)
			ReDim dTempArrayIm(nM - 1)
			For j = 0 To nM - 1
				dTempArrayRe(j) = dLocalZReValues(j, i)
				dTempArrayIm(j) = dLocalZImValues(j, i)
			Next
			oDataColumns(i).SetArray(dLocalYValues, "x")
			oDataColumns(i).SetArray(dTempArrayRe, "yre")
			oDataColumns(i).SetArray(dTempArrayIm, "yim")
			oDataColumns(i).ResampleTo(dLocalYValues(0), dLocalYValues(nM - 1), nYSamples)
		Next
		' Create new dZValues from rows
		ReDim dLocalZReValues(nYSamples - 1, nN - 1)
		ReDim dLocalZImValues(nYSamples - 1, nN - 1)
		For i = 0 To nN - 1
			dTempArrayRe = oDataColumns(i).GetArray("yre")
			dTempArrayIm = oDataColumns(i).GetArray("yim")
			For j = 0 To nYSamples - 1
				dLocalZReValues(j, i) = dTempArrayRe(j)
				dLocalZImValues(j, i) = dTempArrayIm(j)
			Next
		Next
		nM = nYSamples
		dLocalYValues = oDataColumns(0).GetArray("x")
	End If

	dXValues = dLocalXValues
	dYValues = dLocalYValues
	dZReValues = dLocalZReValues
	dZImValues = dLocalZImValues

End Function


Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim i As Long
	Dim sSelectedResultName As String, sSelectedResultType As String
	Dim nXSamples As Long, nYSamples As Long
	Dim sXAxisArray(0) As String

	sXAxisArray(0) = "x-axis"

	Select Case Action%
	Case 1 ' Dialog box initialization
		If (Left(GetApplicationName, 2) = "DS") Then
			nNumberOfParameters = DS.GetNumberOfParameters
			If (nNumberOfParameters < 1) Then ReportError("This macro requires at least 1 parameter to be defined in the project.")
			ReDim sParameterList(nNumberOfParameters-1)
			For i = 0 To nNumberOfParameters-1
				sParameterList(i) = DS.GetParameterName(i)
			Next
		Else
			nNumberOfParameters = GetNumberOfParameters
			If (nNumberOfParameters < 1) Then ReportError("This macro requires at least 1 parameter to be defined in the project.")
			ReDim sParameterList(nNumberOfParameters-1)
			For i = 0 To nNumberOfParameters-1
				sParameterList(i) = GetParameterName(i)
			Next
		End If

		DlgListBoxArray("XParameterDLB", sParameterList())
		DlgListBoxArray("YParameterDLB", sParameterList())
		DlgValue("XParameterDLB", 0)
		DlgValue("YParameterDLB", 0)
		DlgEnable("XParameterDLB", False)
		DlgEnable("YParameterDLB", False)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "aRTemplateA"
				sSelectedResultName = aRTemplateName(DlgValue("aRTemplateA"))
				sSelectedResultType = aRTemplateType(DlgValue("aRTemplateA"))
				DlgEnable("YParameterDLB", True)
				If (InStr(UCase(sSelectedResultType), "0D")>0) Then
					DlgListBoxArray("XParameterDLB", sParameterList())
					DlgEnable("XParameterDLB", True)
				Else
					DlgListBoxArray("XParameterDLB", sXAxisArray())
					DlgEnable("XParameterDLB", False)
				End If
				DlgValue("XParameterDLB", 0)
				DlgListBoxArray("YParameterDLB", sParameterList())
				DlgValue("YParameterDLB", IIf(DlgEnable("XParameterDLB"), 1, 0))
			Case "OK"
				sSelectedResultName = aRTemplateName(DlgValue("aRTemplateA"))
				sSelectedResultType = aRTemplateType(DlgValue("aRTemplateA"))

				nXSamples = IIf(DlgValue("ResampleXCB") = 1, CLng(DlgText("XSamplesT")), 0)
				nYSamples = IIf(DlgValue("ResampleYCB") = 1, CLng(DlgText("YSamplesT")), 0)

				DlgEnable("OK", False)
				DlgEnable("Cancel", False)
				CreateColormapPlot(sSelectedResultName, sSelectedResultType, DlgText("OutputNameT"), nXSamples, nYSamples)
				DlgEnable("Cancel", True)
				DlgEnable("OK", True)
			Case "Cancel"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select

	DlgEnable("XSamplesT", DlgEnable("ResampleXCB") And DlgValue("ResampleXCB") = 1)
	DlgEnable("YSamplesT", DlgEnable("ResampleYCB") And DlgValue("ResampleYCB") = 1)

End Function

