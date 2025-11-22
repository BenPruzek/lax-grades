' trasform_towcs_or_togcs

' ================================================================================================
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------------
' 01-Apr-2016 kko: improved macro to be able to deal with subcomponents and mixed selections (only working from Suite Version 2017 on)
' 17-Dec-2013 ube: submitted first version, created by tgl
'------------------------------------------------------------------------------------------

Option Explicit

Sub Main ()
	Dim nItem As Long
	Dim sItem As String
	Dim sItemQu As String
	Dim sHistEntry As String
	Dim direction As String
	Dim objects As String
	Dim ii As Integer
	Dim FoundObjects As Integer
	Dim FoundWrongSelection As Boolean ' default of Boolean is False.

	nItem = GetNumberOfSelectedTreeItems()

	Begin Dialog UserDialog 340,119, "Transform selected objects" ' %GRID:10,7,1,1
		Text 20,14,420,21,"This operation will transform/align " + CStr(nItem) + "objects.",.Text1
		OptionGroup .Group1
			OptionButton 40,42,170,14,"from Global to WCS" ',.toWCS
			OptionButton 40,63,170,14,"from WCS to Global" ',.toGCS
		OKButton 20,91,90,21
		CancelButton 120,91,90,21
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then Exit All

	If dlg.Group1 = 0 Then
		direction = ".Transform ""Mixed"", ""toWCS""" + vbLf
	ElseIf dlg.Group1 = 1 Then
		direction = ".Transform ""Mixed"", ""toGCS""" + vbLf
	Else
		ReportError("Something went wrong in the GUI")
	End If

	sHistEntry = "With Transform" + vbLf + _
     			".Reset" + vbLf + _
			    ".UsePickedPoints ""False""" + vbLf + _
     			".InvertPickedPoints ""False""" + vbLf + _
     			".MultipleObjects ""False""" + vbLf + _
     			".GroupObjects ""False""" + vbLf + _
			    ".Repetitions ""1""" + vbLf + _
     			".MultipleSelection ""False""" + vbLf

	sItem = GetSelectedTreeItem()
	sItemQu = GetQualifiedNameFromTreeName(sItem)
	Do
		If sItemQu <> "" Then
			If (FoundObjects < 1) Then
				objects = ".Name """ + sItemQu + """" + vbLf
			Else
				objects = objects + ".AddName """ + sItemQu + """" + vbLf
			End If

			FoundObjects = FoundObjects + 1
		Else
			FoundWrongSelection = True

		End If
		sItem =	GetNextSelectedTreeItem
		sItemQu = GetQualifiedNameFromTreeName(sItem)
	Loop While sItem <> ""

	If FoundWrongSelection Or FoundObjects < 1 Then
		ReportError("No valid selection found. Please select transformable objects from the tree first.")
		Exit All
	End If

	sHistEntry = sHistEntry + objects + direction + "End With" + vbLf
	addToHistory("transform selected objects to or from wcs", sHistEntry) ' please note: "transform" is a keyword that allows the history entry to be repeated several times!

End Sub
