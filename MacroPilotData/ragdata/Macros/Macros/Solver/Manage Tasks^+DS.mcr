'#Language "WWB-COM"
'#include "vba_globals_all.lib"
'#include "complex.lib"
'#include "infix_postfix.lib"

Option Explicit

' This macro simplifies managing the DS task list by allowing multiselection for enabling/disabling/deleting
'
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' -------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 03-Sep-2024 dre2: Removed all Run buttons, since it blocks the entire DE (see CST-71186)
' 31-Aug-2020 fsr: Commented out 'Run All Tasks' since aborting tasks is not supported through VBA
' 13-Apr-2018 fsr: Check if a task exist before trying to delete it. (Could have been deleted already as part of sequence/optimization/sweep task)
' 17-Jan-2018 fsr: Fixed index out of bounds error that was triggered when deleting the last task
' 03-Nov-2016 fsr: Added hidden Cancel button to enable X in top right corner; improved behavior of buttons when no tasks are selected
' 24-Oct-2016 ube: added Help page, check if at least one task is defined
' 21-Oct-2016 fsr: Minor GUI changes; added "Run All" option and help button
' 22-Sep-2016 fsr: Initial version
' -------------------------------------------------------------------------------------------------------------------------------------------------

Private sTaskList() As String

Sub Main

	If SimulationTask.StartTaskNameIteration < 1 Then
		MsgBox "No tasks defined. Please first define at least one task.", vbOkOnly, "Manage Tasks"
		Exit All
	End If

	Dim sTaskSelectionList() As String

	Begin Dialog UserDialog 550,455,"Manage Tasks",.DialogFunc ' %GRID:10,7,1,1
		MultiListBox 20,14,410,427,sTaskSelectionList(),.TaskSelectionMLB
		PushButton 450,14,90,21,"Move up",.MoveUpPB
		PushButton 450,42,90,21,"Move down",.MoveDownPB

		PushButton 450,105,90,21,"Enable",.EnablePB
		PushButton 450,133,90,21,"Disable",.DisablePB
		PushButton 450,161,90,21,"Delete",.DeletePB

		PushButton 450,392,90,21,"Help",.HelpPB
		CancelButton 450,420,90,21
		PushButton 450,420,90,21,"Close",.ClosePB

	End Dialog

	Dim dlg As UserDialog

	If Dialog(dlg, 0) = 0 Then
		Exit All
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim nTasks As Long, i As Long
	Dim nSelectedTaskIndexes() As Integer
	Dim sSelectedTask As String
	Dim bIsEnabled As Boolean

	nSelectedTaskIndexes = DlgValue("TaskSelectionMLB")

    Select Case Action%
    Case 1 ' Dialog box initialization
		' Do nothing except to update task array, which is always done below
		DlgVisible("Cancel", False) ' do not show, just used to enable behavior of "X" in dialog
	Case 2 ' Value changing or button pressed
    	Select Case DlgItem$
    		Case "MoveUpPB"
				MoveTasksUp(nSelectedTaskIndexes)
				DialogFunc = True 'do not exit the dialog
			Case "MoveDownPB"
				MoveTasksDown(nSelectedTaskIndexes)
				DialogFunc = True 'do not exit the dialog
			Case "EnablePB"
				For i = 0 To (UBound(nSelectedTaskIndexes))
					sSelectedTask = Replace(sTaskList(nSelectedTaskIndexes(i)), " [Disabled]", "")
					With SimulationTask
						.Name(sSelectedTask)
						.SetProperty("enabled","true")
					End With
				Next
				DialogFunc = True 'do not exit the dialog
			Case "DisablePB"
				For i = 0 To (UBound(nSelectedTaskIndexes))
					sSelectedTask = sTaskList(nSelectedTaskIndexes(i))
					With SimulationTask
						.Name(sSelectedTask)
						.SetProperty("enabled","false")
					End With
				Next
				DialogFunc = True 'do not exit the dialog
			Case "DeletePB"
				If (MsgBox("The selected tasks will be permanently deleted. Do you want to proceed?", vbYesNo, "Confirm Delete Tasks") = vbYes) Then
					For i = 0 To (UBound(nSelectedTaskIndexes))
						sSelectedTask = Replace(sTaskList(nSelectedTaskIndexes(i)), " [Disabled]", "")
						With SimulationTask
							.Name(sSelectedTask)
							If .DoesExist Then .Delete
						End With
					Next
				End If
				DialogFunc = True 'do not exit the dialog
			Case "TaskSelectionMLB"
				DialogFunc = True 'do not exit the dialog
			Case "HelpPB"
				StartHelp "common_preloadedmacro_Manage_Tasks"
				DialogFunc = True
			Case "ClosePB"
				Exit All
		End Select
	End Select

	UpdateTaskNamesArray(nSelectedTaskIndexes)

End Function

Function UpdateTaskNamesArray(nCurrentSelection() As Integer) As Integer

	Dim i As Long, nTasks As Long

	With SimulationTask
		nTasks = .StartTaskNameIteration
		If (nTasks > 0) Then
			ReDim sTaskList(nTasks-1)
			For i = 0 To nTasks-1
				sTaskList(i) = .GetNextTaskName
				.Name(sTaskList(i))
				If Not CBool(.GetProperty("enabled")) Then sTaskList(i) = sTaskList(i) + " [Disabled]"
			Next
		Else
			ReDim sTaskList(0)
			sTaskList(0) = ""
		End If
	End With
	DlgListBoxArray("TaskSelectionMLB", sTaskList)
	DlgValue("TaskSelectionMLB", nCurrentSelection)
	UpdateTaskNamesArray = 0 ' all went well

End Function

Function MoveTasksUp(nTaskIndexes() As Integer) As Integer

	' Move task up in list, within the same parent folder
	Dim sTaskFolder As String, sCurrentTaskName As String, sSelectedTasks() As String, sTasksInCurrentFolder() As String, nCurrentTaskPosition As Integer
	Dim sNextTaskInFolder As String
	Dim i As Long, j As Long

	If (UBound(nTaskIndexes) < LBound(nTaskIndexes)) Then Exit Function ' Do nothing as nothing is selected

	ReDim sSelectedTasks(UBound(nTaskIndexes))

	For i = 0 To UBound(nTaskIndexes)
		' Needs to run in a separate loop to avoid name mix-ups due to changing task positions later
		sSelectedTasks(i) = sTaskList(nTaskIndexes(i))
	Next

	For i = 0 To UBound(nTaskIndexes)
		' Determine task folder
		sCurrentTaskName = sSelectedTasks(i)
		If (InStr(sCurrentTaskName, "\") > 0) Then
			sTaskFolder = Left(sCurrentTaskName, InStrRev(sCurrentTaskName, "\") - 1)
		Else
			sTaskFolder = ""
		End If
		ReDim sTasksInCurrentFolder(0)
		' Find all tasks in this folder, determine position of current task in this list
		For j = 0 To UBound(sTaskList)
			' If task found, note its position and exit
			If (((UBound(Split(Replace(sTaskList(j), sTaskFolder, ""), "\")) = 1) And (Left(Replace(sTaskList(j), sTaskFolder, ""), 1) = "\")) _
					Or ((sTaskFolder = "") And (UBound(Split(sTaskList(j), "\")) = 0))) Then
				If (sTaskList(j) = sCurrentTaskName) Then nCurrentTaskPosition = UBound(sTasksInCurrentFolder)
				sTasksInCurrentFolder(UBound(sTasksInCurrentFolder)) = sTaskList(j)
				ReDim Preserve sTasksInCurrentFolder(UBound(sTasksInCurrentFolder) + 1)
			End If
		Next

		If nCurrentTaskPosition = 0 Then ' task is already on top of current folder
			' do nothing
		Else
			sNextTaskInFolder = Replace(sTasksInCurrentFolder(nCurrentTaskPosition - 1), sTaskFolder  & "\", "")
			If (i > 0) Then ' check if sNextTaskInFolder is the previously moved task; do nothing in this case to keep order
				If (((sTaskFolder = "") And (sNextTaskInFolder = sSelectedTasks(i - 1))) _
					Or (sTaskFolder & "\" & sNextTaskInFolder = sSelectedTasks(i - 1))) Then
					sNextTaskInFolder = Replace(sTasksInCurrentFolder(nCurrentTaskPosition + 1), sTaskFolder  & "\", "")
				End If
			End If
			With SimulationTask
				.Name(Replace(sCurrentTaskName, " [Disabled]", ""))
				.MoveInTree(sTaskFolder, Replace(sNextTaskInFolder, " [Disabled]", ""))
			End With
			UpdateTaskNamesArray(nTaskIndexes)
		End If
	Next

	For i = 0 To UBound(sSelectedTasks)
		' update list of indexes so that task selection in dialog remains
		' needs to be done in a separate loop after all tasks are moved (multiselection possible!)
		sCurrentTaskName = sSelectedTasks(i)
		nTaskIndexes(i) = FindListIndex(sTaskList, sCurrentTaskName)
	Next

End Function

Function MoveTasksDown(nTaskIndexes() As Integer) As Integer

	' Move task down in list, within the same parent folder
	Dim sTaskFolder As String, sCurrentTaskName As String, sSelectedTasks() As String, sTasksInCurrentFolder() As String, nCurrentTaskPosition As Integer
	Dim sNextTaskInFolder As String
	Dim i As Long, j As Long

	If (UBound(nTaskIndexes) < LBound(nTaskIndexes)) Then Exit Function ' Do nothing as nothing is selected

	ReDim sSelectedTasks(UBound(nTaskIndexes))

	For i = UBound(nTaskIndexes) To 0 STEP - 1
		' Determine task folder
		sCurrentTaskName = sTaskList(nTaskIndexes(i))
		sSelectedTasks(i) = sCurrentTaskName
		If (InStr(sCurrentTaskName, "\") > 0) Then
			sTaskFolder = Left(sCurrentTaskName, InStrRev(sCurrentTaskName, "\") - 1)
		Else
			sTaskFolder = ""
		End If
		ReDim sTasksInCurrentFolder(0)
		' Find all tasks in this folder, determine position of current task in this list
		For j = 0 To UBound(sTaskList)
			If (((UBound(Split(Replace(sTaskList(j), sTaskFolder, ""), "\")) = 1) And (Left(Replace(sTaskList(j), sTaskFolder, ""), 1) = "\")) _
					Or ((sTaskFolder = "") And (UBound(Split(sTaskList(j), "\")) = 0))) Then
				If (sTaskList(j) = sCurrentTaskName) Then nCurrentTaskPosition = UBound(sTasksInCurrentFolder)
				sTasksInCurrentFolder(UBound(sTasksInCurrentFolder)) = sTaskList(j)
				ReDim Preserve sTasksInCurrentFolder(UBound(sTasksInCurrentFolder) + 1)
			End If
		Next

		If ((UBound(sTasksInCurrentFolder) = 0) Or (nCurrentTaskPosition = UBound(sTasksInCurrentFolder) - 1)) Then ' task is already at bottom of current folder
			' do nothing
		Else
			sNextTaskInFolder = Replace(sTasksInCurrentFolder(nCurrentTaskPosition + 2), sTaskFolder  & "\", "")
			If (i < UBound(sSelectedTasks) - 1) Then ' check if sNextTaskInFolder is the previously moved task; do nothing in this case to keep order
				If (((sNextTaskInFolder = "") And (sSelectedTasks(i + 1) = sTasksInCurrentFolder(UBound(sTasksInCurrentFolder) - 1))) _
					Or (sSelectedTasks(i + 2) = sTaskFolder & "\" & sNextTaskInFolder) _
					Or ((sTaskFolder = "") And (sSelectedTasks(i + 2) = sNextTaskInFolder))) Then
					sNextTaskInFolder = Replace(sSelectedTasks(i + 1), sTaskFolder  & "\", "")
				End If
			ElseIf (i < UBound(sSelectedTasks)) Then ' check if sNextTaskInFolder is the previously moved task; do nothing in this case to keep order
				If ((sNextTaskInFolder = "") And (sSelectedTasks(i + 1) = sTasksInCurrentFolder(UBound(sTasksInCurrentFolder) - 1))) Then
					sNextTaskInFolder = Replace(sSelectedTasks(i + 1), sTaskFolder  & "\", "")
				End If
			End If
			With SimulationTask
				.Name(Replace(sCurrentTaskName, " [Disabled]", ""))
				.MoveInTree(sTaskFolder, Replace(sNextTaskInFolder, " [Disabled]", ""))
			End With
			UpdateTaskNamesArray(nTaskIndexes)
		End If
	Next

	For i = 0 To UBound(sSelectedTasks)
		' update list of indexes so that task selection in dialog remains
		' needs to be done in a separate loop after all tasks are moved (multiselection possible!)
		sCurrentTaskName = sSelectedTasks(i)
		nTaskIndexes(i) = FindListIndex(sTaskList, sCurrentTaskName)
	Next

End Function
