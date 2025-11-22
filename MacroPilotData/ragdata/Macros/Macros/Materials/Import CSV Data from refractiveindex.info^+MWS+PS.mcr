'#Language "WWB-COM"
'#include "vba_globals_all.lib"

Option Explicit


' This file reads in material data in csv format from refractiveindex.info
' Format: First 2 columns wl/n, then 2 columns wl/k (optional). wl in um

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' --------------------------------------------------------------------------------------
' 23-Mar-2015 fde: Fixed a problem with data files that do not contain k data
' 30-Jan-2015 wch: Show (fmin,fmax) and check the frequency range settings
' 01-Jul-2014 jfl: Linear interpolation added for data of different sampling
' 07-May-2014 fsr: Initial version
' --------------------------------------------------------------------------------------

Sub Main

	Dim sFilePath As String, sFileName As String, sFileContents As String, sFileLines() As String
	Dim dNWavelength() As Double, dKWavelength() As Double, dNData() As Double, dKData() As Double
	Dim i As Long
	Dim nLastIndex As Double
	Dim sHistoryString As String, sFrequency As String, sEpsRe As String, sEpsIm As String

	'Load in data file
	sFilePath = GetFilePath("*.csv", "*.csv", "", "Please select material file in csv format", 0)
	sFileName = Mid(sFilePath, InStrRev(sFilePath, "\")+1)
	sFileContents = TextFileToString_LIB(sFilePath)
	sFileLines = Split(sFileContents, vbNewLine)

	ReDim dNData(0)
	ReDim dNWavelength(0)
	ReDim dKData(0)
	ReDim dKWavelength(0)

	i = 0
	Do
		i = i + 1
	Loop While sFileLines(i - 1) <> "wl,n" 'Jump to n data

	While ((i <= UBound(sFileLines)) And (sFileLines(i) <> "") And (sFileLines(i) <> ",") And (sFileLines(i) <> "wl,k"))
		nLastIndex = UBound(dNData)
		ReDim Preserve dNData(nLastIndex + 1)
		ReDim Preserve dNWavelength(nLastIndex + 1)
		dNData(nLastIndex) = Evaluate(Split(sFileLines(i), ",")(1))
		dNWavelength(nLastIndex) = Evaluate(Split(sFileLines(i), ",")(0))
		i = i + 1
	Wend


	If (InStr(sFileContents, "wl,k")) Then
		While sFileLines(i - 1) <> "wl,k" 'Jump to k data if it exists
			i = i + 1
		Wend
	End If


	If (i<UBound(sFileLines)) Then ' k data to follow
		While (i <= UBound(sFileLines))
			If (sFileLines(i) = "") Then Exit While
			nLastIndex = UBound(dKData)
			ReDim Preserve dKData(nLastIndex + 1)
			ReDim Preserve dKWavelength(nLastIndex + 1)
			dKData(nLastIndex) = Evaluate(Split(sFileLines(i), ",")(1))
			dKWavelength(nLastIndex) = Evaluate(Split(sFileLines(i), ",")(0))
			i = i + 1
		Wend
	Else
		ReDim Preserve dKData(UBound(dNData))
		ReDim Preserve dKWavelength(UBound(dNData))
		For i = 0 To UBound(dNData) 'No k data; default all k values to 1e-6
			dKData(i) = 1e-6
			dKWavelength(i) = dNWavelength(i)
		Next
	End If

	'All arrays have a UBound 1 higher than needed
	ReDim Preserve dNData(UBound(dNData)-1)
	ReDim Preserve dNWavelength(UBound(dNWavelength)-1)
	ReDim Preserve dKData(UBound(dKData)-1)
	ReDim Preserve dKWavelength(UBound(dKWavelength)-1)

	'Sort data in case presented out of order
	SortWithChild(dNWavelength, dNData)
	SortWithChild(dKWavelength, dKData)

	If (dNWavelength(0) = 0) Then dNWavelength(0) = 1e-12 ' prevent division by 0
	If (dKWavelength(0) = 0) Then dKWavelength(0) = 1e-12 ' prevent division by 0

	'Find the intersection of the k and n wavelengths and trim off all data outside
	Dim nUpBound As Integer, nLowBound As Integer
	nLowBound = 0
	nUpBound = UBound(dNWavelength)
	For i = 1 To UBound(dNWavelength) - 1
		If (dNWavelength(i) <= dKWavelength(0)) Then nLowBound = i
		If (dNWavelength(i) >= dKWavelength(UBound(dKWavelength))) Then nUpBound = i
	Next
	For i = 0 To nUpBound - nLowBound
		dNWavelength(i) = dNWavelength(i + nLowBound)
		dNData(i) = dNData(i + nLowBound)
	Next
	nLowBound = 0
	nUpBound = UBound(dKWavelength)
	For i = 1 To UBound(dKWavelength) - 1
		If (dKWavelength(i) <= dNWavelength(0)) Then nLowBound = i
		If (dKWavelength(i) >= dNWavelength(UBound(dNWavelength))) Then nUpBound = i
	Next
	For i = 0 To nUpBound - nLowBound
		dKWavelength(i) = dKWavelength(i + nLowBound)
		dKData(i) = dKData(i + nLowBound)
	Next

	'Ensure each set of data includes all wavelengths from the other
	MakeComparableTo(dNWavelength, dNData, dKWavelength)
	MakeComparableTo(dKWavelength, dKData, dNWavelength)

	'Handles any extra boundary wavelengths present in only one set of data
	If dNWavelength(UBound(dNWavelength)) > dKWavelength(UBound(dKWavelength)) Then
		ReDim Preserve dNWavelength(UBound(dNWavelength)-1)
	Else
		If dKWavelength(UBound(dKWavelength)) > dNWavelength(UBound(dNWavelength)) Then
			ReDim Preserve dKWavelength(UBound(dKWavelength)-1)
		End If
	End If

	If dNWavelength(0) < dKWavelength(0) Then
		For i = 1 To UBound(dNWavelength)
			dNWavelength(i-1) = dNWavelength(i)
		Next
		ReDim Preserve dNWavelength(UBound(dNWavelength)-1)
	Else
		If dKWavelength(0) < dNWavelength(0) Then
			For i = 1 To UBound(dKWavelength)
				dKWavelength(i-1) = dKWavelength(i)
			Next
			ReDim Preserve dKWavelength(UBound(dKWavelength)-1)
		End If
	End If

	'Check the frequency range and imported data range
	Dim fminToSi As Double, fmaxToSi As Double, ffactor As Double
	Dim UBFreqToSi As Double, LBFreqToSi As Double
	Dim sMsg As String,fUnit As String, sUBFreq As String, sLBFreq As String

	With Solver
	  fminToSi = .GetFmin*Units.GetFrequencyUnitToSi
	  fmaxToSi = .GetFmax*Units.GetFrequencyUnitToSi
	  fUnit = Units.GetUnit("Frequency")
	End With

	UBFreqToSi = CLight/(dNWavelength(1)*1e-6)
    LBFreqToSi = CLight/(dNWavelength(UBound(dKWavelength))*1e-6)

    'Read the imported frequency range in Msgbox
	sUBFreq = CStr(Round(UBFreqToSi/Units.GetFrequencyUnitToSi))
	sLBFreq = CStr(Round(LBFreqToSi/Units.GetFrequencyUnitToSi))


	sHistoryString = "With Material" + vbNewLine
	sHistoryString = sHistoryString + "     .Reset" + vbNewLine
	sHistoryString = sHistoryString + "     .Name " + Chr(34) + sFileName + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .Folder " + Chr(34) + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .Type " + Chr(34) + "Normal"  + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .UseGeneralDispersionEps " + Chr(34) + "True"  + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .DispersiveFittingSchemeEps " + Chr(34) + "Nth Order"  + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitEps " + Chr(34) + "10"  + Chr(34) + vbNewLine
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitEps " + Chr(34) + "0.01"  + Chr(34) + vbNewLine
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelEps " + Chr(34) + "True"  + Chr(34) + vbNewLine

	For i = 0 To UBound(dNWavelength)
		sFrequency = CStr(CLight/(dNWavelength(i)*1e-6)/Units.GetFrequencyUnitToSi)
		sEpsRe = CStr(dNData(i)^2-dKData(i)^2)
		sEpsIm = CStr(2*dNData(i)*dKData(i))
		sHistoryString = sHistoryString + "     .AddDispersionFittingValueEps " + Chr(34) + sFrequency + Chr(34) + ", " + Chr(34) + sEpsRe + Chr(34) + ", "+ Chr(34) + sEpsIm + Chr(34) + ", "+ Chr(34) + "1" + Chr(34) + vbNewLine
	Next i

	sHistoryString = sHistoryString + "     .Create" + vbNewLine
	sHistoryString = sHistoryString + "End With" + vbNewLine

	AddToHistory("define material: " + sFileName, sHistoryString)

	If ((fminToSi > UBFreqToSi) Or (fmaxToSi < LBFreqToSi)) And (Solver.GetFMax > 0) Then
		sMsg = "Material data were imported successfully, but all " & vbNewLine _
				& "data points are outside the current frequency range." + vbNewLine
		sMsg = sMsg + "Simulation (fmin,fmax) = (" + CStr(Round(Solver.GetFmin)) + " ... " + CStr(Round(Solver.GetFmax)) + ") " + fUnit + vbNewLine
		sMsg = sMsg + "Imported (fmin,fmax) = (" + sLBFreq + " ... " + sUBFreq + ") " + fUnit + vbNewLine
		sMsg = sMsg + "Please check the solver frequency settings !"
		MsgBox(sMsg, "Please check your settings")
	Else
		sMsg = "Material data were imported successfully." + vbNewLine
		sMsg = sMsg + "Frequency range of data: (" + sLBFreq + " ... " + sUBFreq + ") " + fUnit + vbNewLine
		MsgBox(sMsg, "Success")
	End If

End Sub

'Sorts a parent array making the same adjustments to a child array to preserve the relationship between the arrays
Private Function SortWithChild(ByRef parent() As Double, ByRef child() As Double)
	If UBound(parent) <= 0 Or UBound(parent) <> UBound(child) Then Exit Function
	Dim j As Integer, k As Integer, l As Integer
	Dim dPlaceholderParent As Double, dPlaceholderChild As Double
	For j = 1 To UBound(parent)

		dPlaceholderChild = child(j)
		dPlaceholderParent = parent(j)
		k = j

		While (k > 0)
			If (dPlaceholderParent >= parent(k - 1)) Then Exit While
			k = k - 1
		Wend
		l = j - 1
		While l >= k
			parent(l + 1) = parent(l)
			child(l + 1) = child(l)
			l = l - 1
		Wend
		parent(k) = dPlaceholderParent
		child(k) = dPlaceholderChild

	Next j
End Function

Private Function MakeComparableTo (ByRef dataXValues() As Double, ByRef dataYValues() As Double, ByVal xValuesToMatch() As Double)
	Dim j As Integer, k As Integer
	For j = 0 To UBound(xValuesToMatch)
		If ((j = UBound(xValuesToMatch)) And (xValuesToMatch(UBound(xValuesToMatch)) > dataXValues(UBound(dataXValues)))) Then Exit For
		If Not (j = 0 And xValuesToMatch(0) < dataXValues(0)) Then
			k = 0
			While (k <= UBound(dataXValues))
				If dataXValues(k) >= xValuesToMatch(j) Then Exit While
				k = k + 1
			Wend
	 		If (dataXValues(k) <> xValuesToMatch(j)) Then
				ReDim Preserve dataXValues(UBound(dataXValues) + 1)
				ReDim Preserve dataYValues(UBound(dataYValues) + 1)
				dataXValues(UBound(dataXValues)) = xValuesToMatch(j)

				dataYValues(UBound(dataYValues)) = _
					((dataYValues(k) - dataYValues(k - 1))/(dataXValues(k) - dataXValues(k - 1))) _
					* (xValuesToMatch(j) - dataXValues(k - 1)) _
					+ dataYValues(k - 1)
				'Ynew = slope between the two points * delta-x from first point + y-value of first point

				SortWithChild(dataXValues, dataYValues)
			End If
		End If
	Next j
End Function
