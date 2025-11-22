Option Explicit

'#include "vba_globals_all.lib"
'#include "template_conversions.lib"

' ================================================================================================
' Macro to calculate min + max value of selected 0d table
'
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 25-Sep-2019 fsr: Replaced Table mechanism with new RunID infrastructure
' 29-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 06-Jul-2015 fsr: Replaced 'GetTypeOfDataItem' with new 'GetResultTypeOfDataItem'
' 25-Jul-2007 ube: first version
' ================================================================================================

Sub Main
	Dim stree As String
	Dim sListOfRunIDs() As String, sListOfParameterNames As Variant, sTempList As Variant, sListOfParameterValues() As Variant

	stree = GetSelectedTreeItem
	If Left$(stree, 18)<>"Tables\0D Results\" Then
		MsgBox "Macro only works with 0D Tables, containing real numbers."+vbCrLf+"Please select such a table in the tree first.", vbExclamation
		Exit All
	End If

	Dim dMinValue As Double, dMaxValue As Double, iData As Long, nData As Long, iPara As Integer, nPara As Integer, dv As Double, iMin As Long, iMax As Long

	sListOfRunIDs() = GetListOfRunIDs_LIB(stree)
	nData = UBound(sListOfRunIDs) + 1

	If nData=0 Then
		MsgBox "Table does not contain any data."+vbCrLf+"Exit.", vbExclamation
		Exit All
	End If

	ReDim sListOfParameterValues(nData - 1)
	GetParameterCombination(sListOfRunIDs(0), sListOfParameterNames, sTempList)
	sListOfParameterValues(0) = sTempList
	nPara = UBound(sListOfParameterNames) + 1

	iData = 0
	dv = GetResultByRunID_LIB(sListOfRunIDs(iData), Replace(stree, "Tables\0D Results\", ""), "0D", "0D")
	dMinValue = dv
	dMaxValue = dv
	iMin = iData
	iMax = iData

	For iData=1 To nData-1
		GetParameterCombination(sListOfRunIDs(iData), sListOfParameterNames, sTempList)
		sListOfParameterValues(iData) = sTempList
		dv = GetResultByRunID_LIB(sListOfRunIDs(iData), Replace(stree, "Tables\0D Results\", ""), "0D", "0D")
		If (dv < dMinValue) Then
			dMinValue = dv
			iMin = iData
		End If
		If (dv > dMaxValue) Then
			dMaxValue = dv
			iMax = iData
		End If
	Next iData

	' now output string
	Dim sout As String

	sout =        "---------------------------------------------" + vbCrLf
	sout = sout + "Maximum = " + CStr(dMaxValue) + vbCrLf
	sout = sout + "---------------------------------------------" + vbCrLf
	sout = sout + "Maximum reached at:" + vbCrLf

	For iPara=0 To nPara-1
		sout = sout + Chr$(9) + sListOfParameterNames(iPara) + " = " + CStr(sListOfParameterValues(iMax)(iPara)) + vbCrLf
	Next iPara
	sout = sout + vbCrLf

	sout = sout + "---------------------------------------------" + vbCrLf
	sout = sout + "Minimum = " + CStr(dMinValue) + vbCrLf
	sout = sout + "---------------------------------------------" + vbCrLf
	sout = sout + "Minimum reached at:" + vbCrLf

	For iPara=0 To nPara-1
		sout = sout + Chr$(9) + sListOfParameterNames(iPara) + " = " + CStr(sListOfParameterValues(iMin)(iPara)) + vbCrLf
	Next iPara

	stree = Mid(stree, 1+InStrRev(stree,"\"))
	MsgBox sout, ,"Table: " + stree

End Sub
