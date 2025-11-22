' *Get selected item path
' !!! Do not change the line above !!!
' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 31-Oct-2022 ube: First version
' ================================================================================================

Option Explicit

Sub Main

	Dim ItemName As String

	If Left(GetApplicationName, 2) = "DS" Then
		ItemName = DS.GetSelectedTreeItem
		DS.ReportInformation ItemName
	Else
		ItemName = GetSelectedTreeItem
		ReportInformation ItemName
	End If

	Clipboard ItemName
	
End Sub
