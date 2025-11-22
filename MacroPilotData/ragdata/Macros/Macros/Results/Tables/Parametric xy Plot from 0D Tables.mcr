' Parametric xy Plot

'#include "vba_globals_all.lib"
'#include "template_conversions.lib"

' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' -----------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 06-Jul-2020 fsr: Fixed result selection for Schematic
' 26-Sep-2019 fsr: Replaced Table mechanism with new RunID infrastructure
' 27-Feb-2019 fsr: Removed legacy code that was already commented out since 2015 without tripping an accompanying safety check
' 29-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 06-Jul-2015 fsr: Replaced 'GetTypeOfDataItem' with new 'GetResultTypeOfDataItem'
' 13-Feb-2014 fsr: Added support for DS
' 03-Feb-2014 fsr: replaced bubble sort With New .SortByX Command; changed file Name To tree entry Name such that multiple plots can be shown
' 01-Nov-2011 ube: added bubble sort and put into 2012 version (has also been part of v2011)
' 20-Sep-2010 ube: first version
' -----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main ()

	Dim s0DTables(1000) As String
	Dim ss As String
	Dim icount As Integer
	Dim sListOfRunIDs() As String
	Dim d1 As Double, d2 As Double
	Dim sInputResultType As String, sResultStringX As String, sResultStringY As String

	If (Left(GetApplicationName, 2) = "DS") Then
		If(MsgBox("Please ensure that the appropriate 0D subfolder (e.g. 'Tasks\PP1\0D') is selected. Continue?", vbYesNo) = vbNo) Then Exit All
		ss = DSResultTree.GetFirstChildName(DS.GetSelectedTreeItem())
	Else
		ss = Resulttree.GetFirstChildName("Tables\0D Results")
	End If
	icount = 0

	While ss <> ""
		s0DTables(icount)=ss
		icount = icount + 1
		If (Left(GetApplicationName, 2) = "DS") Then
			ss= DSResultTree.GetNextItemName (ss)
		Else
			ss = Resulttree.GetNextItemName(ss)
		End If
	Wend
	If (icount = 0) Then
		MsgBox "No 0D Tables exist", vbOkOnly
		Exit All
	End If

	Begin Dialog UserDialog 610,140,"Create Parametric xy Plot from 0D-Tables" ' %GRID:10,7,1,1
		Text 20,14,60,14,"x-axis",.Text1
		Text 20,42,50,14,"y-axis",.Text2
		DropListBox 120,14,470,192,s0DTables(),.x
		DropListBox 120,42,470,192,s0DTables(),.y
		Text 20,77,90,14,"Result Tree",.Text3
		TextBox 120,77,470,21,.Tree
		OKButton 20,112,90,21
		CancelButton 120,112,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.Tree="Parametric Plots\Plot"

	If (Dialog(dlg) = 0) Then Exit All

	Dim r1d As Object
	Dim nPara As Long, nData As Long, iPara As Long,	iData As Long

	If (Left(GetApplicationName, 2) = "DS") Then
		Set r1d = DS.Result1D("")
		sInputResultType = ""
		sResultStringX = Replace(s0DTables(dlg.x), "Tasks\", "")
		sResultStringY = Replace(s0DTables(dlg.y), "Tasks\", "")
	Else
		Set r1d = Result1D("")
		sInputResultType = "0D"
		sResultStringX = Replace(s0DTables(dlg.x), "Tables\0D Results\", "")
		sResultStringY = Replace(s0DTables(dlg.y), "Tables\0D Results\", "")
	End If

	sListOfRunIDs() = GetListOfRunIDs_LIB(s0DTables(dlg.x))
	nData = UBound(sListOfRunIDs) + 1
	If (nData <> UBound(GetListOfRunIDs_LIB(s0DTables(dlg.y))) + 1) Then
		MsgBox "x-axis table contains different amount of data points as y-axis" + vbCrLf + "Exit all",vbCritical
		Exit All
	End If

	With r1d

		For iData=0 To nData-1
			d1 = GetResultByRunID_LIB(sListOfRunIDs(iData), sResultStringX, sInputResultType, "0D")
			d2 = GetResultByRunID_LIB(sListOfRunIDs(iData), sResultStringY, sInputResultType, "0D")
			.AppendXY(d1, d2)
		Next iData

		.SortByX ' sorts data by increasing x values

		.Save "^"+NoForbiddenFilenameCharacters(dlg.Tree)+".sig"
		If (Left(GetApplicationName, 2) = "DS") Then
			.AddToTree("Results\"+dlg.Tree)
			SelectTreeItem("Results\"+dlg.Tree)
		Else
			.AddToTree("1D Results\"+dlg.Tree)
			SelectTreeItem("1D Results\"+dlg.Tree)
		End If
	End With

End Sub
