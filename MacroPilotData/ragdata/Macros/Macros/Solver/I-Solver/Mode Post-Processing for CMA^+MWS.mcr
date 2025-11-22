'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 07-Jul-2020 ube: use SelectModelView in addition to SelectTreeItem "Components" (to ensure, 2d3dplot windows become inactive)
' 29-May-2017 nnn: first version
' ================================================================================================

' ---- variables to store the mode indices ----
Dim numModePairs As Integer		' current number of mode pairs
Dim numModePairsMax As Integer	' maximal allowed number of mode pairs (There is no limit by the solver.
								' If more mode pairs are required, only this macro has to be adjusted.)
Dim modeIndices1() As Integer
Dim modeIndices2() As Integer


Sub Main
	' Check if folder with 1D-results of CMA exists
	Dim TreeItem As String
	TreeItem = "1D Results\Characteristic Mode Analysis"

	Dim CMAfolderExists As Variant
	CMAfolderExists = Resulttree.DoesTreeItemExist(TreeItem)

	If CMAfolderExists <> True Then
		MsgBox "Error in mode post-processing for CMA: " & vbNewLine & _
			"No results for the characteristic mode analysis are available."
		End
	End If


	' ---- Actual dialog ----

	numModePairsMax = 6

	numModePairs = 0
	IncreaseSizeArray()

	' define dialog
	Dim dialogH As Integer
	Dim dialogW As Integer
	Dim dialogSpacingLarge As Integer
	Dim dialogSpacingSmall As Integer
	Dim dialogHtext As Integer
	Dim dialogHbutton As Integer
	Dim dialogHgroupBox As Integer
	dialogW = 400			' width of the dialog
	dialogSpacingLarge = 7	' distance from the rim of the dialog box
	dialogSpacingSmall = 4	' distance between two rows
	dialogHtext = 15		' height of a default text box
	dialogHbutton = 21		' height of a button
	dialogHgroupBox = 1 * dialogSpacingLarge + 2 * dialogHtext + 2 * dialogSpacingSmall	' height of one group box
	dialogH = 3 * dialogSpacingLarge + dialogHbutton _
				+ numModePairs * dialogHgroupBox _
				+ (numModePairs - 1) * dialogSpacingLarge	' height of the dialog


	Dim iDialogResponse As Integer
	Dim exitDialog As Boolean
	exitDialog = False

	Dim isInputOK As Boolean


	While exitDialog <> True


		' As a workaround for a dynamic dialog, a separate dialog is defined for each number of mode pairs.
		Select Case numModePairs

		Case 1
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for one pair of modes ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "+", .AddModePairButton
				'PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
				'	dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg As UserDialog

			dlg.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg.modeIndex2Pair1 = CStr(modeIndices2(0))

			iDialogResponse = Dialog(dlg)


		Case 2
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for pair of modes 1 ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ---- group box for pair of modes 2 ----
				GroupBox dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 2"
				' mode index 1
				Text 2 * dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair2
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair2

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "+", .AddModePairButton
				PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg2 As UserDialog

			dlg2.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg2.modeIndex2Pair1 = CStr(modeIndices2(0))
			dlg2.modeIndex1Pair2 = CStr(modeIndices1(1))
			dlg2.modeIndex2Pair2 = CStr(modeIndices2(1))

			iDialogResponse = Dialog(dlg2)


		Case 3
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for pair of modes 1 ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ---- group box for pair of modes 2 ----
				GroupBox dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 2"
				' mode index 1
				Text 2 * dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair2
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair2

				' ---- group box for pair of modes 3 ----
				GroupBox dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 3"
				' mode index 1
				Text 2 * dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair3
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair3

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "+", .AddModePairButton
				PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg3 As UserDialog

			dlg3.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg3.modeIndex2Pair1 = CStr(modeIndices2(0))
			dlg3.modeIndex1Pair2 = CStr(modeIndices1(1))
			dlg3.modeIndex2Pair2 = CStr(modeIndices2(1))
			dlg3.modeIndex1Pair3 = CStr(modeIndices1(2))
			dlg3.modeIndex2Pair3 = CStr(modeIndices2(2))

			iDialogResponse = Dialog(dlg3)


		Case 4
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for pair of modes 1 ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ---- group box for pair of modes 2 ----
				GroupBox dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 2"
				' mode index 1
				Text 2 * dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair2
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair2

				' ---- group box for pair of modes 3 ----
				GroupBox dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 3"
				' mode index 1
				Text 2 * dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair3
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair3

				' ---- group box for pair of modes 4 ----
				GroupBox dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 4"
				' mode index 1
				Text 2 * dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair4
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair4

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "+", .AddModePairButton
				PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg4 As UserDialog

			dlg4.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg4.modeIndex2Pair1 = CStr(modeIndices2(0))
			dlg4.modeIndex1Pair2 = CStr(modeIndices1(1))
			dlg4.modeIndex2Pair2 = CStr(modeIndices2(1))
			dlg4.modeIndex1Pair3 = CStr(modeIndices1(2))
			dlg4.modeIndex2Pair3 = CStr(modeIndices2(2))
			dlg4.modeIndex1Pair4 = CStr(modeIndices1(3))
			dlg4.modeIndex2Pair4 = CStr(modeIndices2(3))

			iDialogResponse = Dialog(dlg4)


		Case 5
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for pair of modes 1 ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ---- group box for pair of modes 2 ----
				GroupBox dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 2"
				' mode index 1
				Text 2 * dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair2
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair2

				' ---- group box for pair of modes 3 ----
				GroupBox dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 3"
				' mode index 1
				Text 2 * dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair3
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair3

				' ---- group box for pair of modes 4 ----
				GroupBox dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 4"
				' mode index 1
				Text 2 * dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair4
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair4

				' ---- group box for pair of modes 5 ----
				GroupBox dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 5"
				' mode index 1
				Text 2 * dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair5
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair5

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "+", .AddModePairButton
				PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg5 As UserDialog

			dlg5.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg5.modeIndex2Pair1 = CStr(modeIndices2(0))
			dlg5.modeIndex1Pair2 = CStr(modeIndices1(1))
			dlg5.modeIndex2Pair2 = CStr(modeIndices2(1))
			dlg5.modeIndex1Pair3 = CStr(modeIndices1(2))
			dlg5.modeIndex2Pair3 = CStr(modeIndices2(2))
			dlg5.modeIndex1Pair4 = CStr(modeIndices1(3))
			dlg5.modeIndex2Pair4 = CStr(modeIndices2(3))
			dlg5.modeIndex1Pair5 = CStr(modeIndices1(4))
			dlg5.modeIndex2Pair5 = CStr(modeIndices2(4))

			iDialogResponse = Dialog(dlg5)


		Case 6
			Begin Dialog UserDialog dialogW, dialogH, "Mode Post-Processing for CMA"
				' ---- group box for pair of modes 1 ----
				GroupBox dialogSpacingLarge, dialogSpacingLarge, dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, _
					"Selected mode pair 1"
				' mode index 1
				Text 2 * dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex1Pair1
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, 100, dialogHtext, _
					"Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, dialogSpacingLarge + dialogSpacingSmall + dialogHtext, _
					30, dialogHtext, .modeIndex2Pair1

				' ---- group box for pair of modes 2 ----
				GroupBox dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 2"
				' mode index 1
				Text 2 * dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair2
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 2 * dialogSpacingLarge + dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 2 * dialogSpacingLarge + dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair2

				' ---- group box for pair of modes 3 ----
				GroupBox dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 3"
				' mode index 1
				Text 2 * dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair3
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 3 * dialogSpacingLarge + 2 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 3 * dialogSpacingLarge + 2 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair3

				' ---- group box for pair of modes 4 ----
				GroupBox dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 4"
				' mode index 1
				Text 2 * dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair4
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 4 * dialogSpacingLarge + 3 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 4 * dialogSpacingLarge + 3 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair4

				' ---- group box for pair of modes 5 ----
				GroupBox dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 5"
				' mode index 1
				Text 2 * dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair5
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 5 * dialogSpacingLarge + 4 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 5 * dialogSpacingLarge + 4 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair5

				' ---- group box for pair of modes 6 ----
				GroupBox dialogSpacingLarge, 6 * dialogSpacingLarge + 5 * dialogHgroupBox, _
					dialogW - 2 * dialogSpacingLarge, dialogHgroupBox, "Selected mode pair 6"
				' mode index 1
				Text 2 * dialogSpacingLarge, 6 * dialogSpacingLarge + 5 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 1:"
				TextBox 2 * dialogSpacingLarge + 100 + dialogSpacingLarge, 6 * dialogSpacingLarge + 5 * dialogHgroupBox + dialogSpacingSmall _
					+ dialogHtext, 30, dialogHtext, .modeIndex1Pair6
				' mode index 2
				Text 2 * dialogSpacingLarge + 200, 6 * dialogSpacingLarge + 5 * dialogHgroupBox + dialogSpacingSmall + dialogHtext, _
					100, dialogHtext, "Mode index 2:"
				TextBox 2 * dialogSpacingLarge + 300 + dialogSpacingLarge, 6 * dialogSpacingLarge + 5 * dialogHgroupBox _
					+ dialogSpacingSmall + dialogHtext, 30, dialogHtext, .modeIndex2Pair6

				' ----buttons----
				PushButton dialogW-3*90-3*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Run", .RunButton
				CancelButton dialogW-2*90-2*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton
				PushButton dialogW-90-1*dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, 90, dialogHbutton, "Help", .helpButton
				' make button have 0-size instead of removing it completely because otherwise the returned code when pressing the
				' following buttons would change compared to the other dialogs
				PushButton dialogSpacingLarge, dialogH-dialogHbutton-dialogSpacingLarge, _
					0 * dialogHbutton, 0 * dialogHbutton, "+", .AddModePairButton
				PushButton 2 * dialogSpacingLarge + dialogHbutton, dialogH-dialogHbutton-dialogSpacingLarge, _
					dialogHbutton, dialogHbutton, "-", .RemoveModePairButton

			End Dialog

			Dim dlg6 As UserDialog

			dlg6.modeIndex1Pair1 = CStr(modeIndices1(0))
			dlg6.modeIndex2Pair1 = CStr(modeIndices2(0))
			dlg6.modeIndex1Pair2 = CStr(modeIndices1(1))
			dlg6.modeIndex2Pair2 = CStr(modeIndices2(1))
			dlg6.modeIndex1Pair3 = CStr(modeIndices1(2))
			dlg6.modeIndex2Pair3 = CStr(modeIndices2(2))
			dlg6.modeIndex1Pair4 = CStr(modeIndices1(3))
			dlg6.modeIndex2Pair4 = CStr(modeIndices2(3))
			dlg6.modeIndex1Pair5 = CStr(modeIndices1(4))
			dlg6.modeIndex2Pair5 = CStr(modeIndices2(4))
			dlg6.modeIndex1Pair6 = CStr(modeIndices1(5))
			dlg6.modeIndex2Pair6 = CStr(modeIndices2(5))

			iDialogResponse = Dialog(dlg6)

		End Select


		' ---- Evaluate dialog response ----

		If iDialogResponse = 1 Then	'Run button

			' -------------------------------
			' ---- Check for valid input ----
			' -------------------------------
			isInputOK = True

			Select Case numModePairs
			Case 1
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg.modeIndex1Pair1)) Or Not(IsNumeric(dlg.modeIndex2Pair1)) Then
					isInputOK = False
				End If

			Case 2
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg2.modeIndex1Pair1)) Or Not(IsNumeric(dlg2.modeIndex2Pair1)) Then
					isInputOK = False
				End If
				' ---- mode pair 2 ----
				If Not(IsNumeric(dlg2.modeIndex1Pair2)) Or Not(IsNumeric(dlg2.modeIndex2Pair2)) Then
					isInputOK = False
				End If

			Case 3
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg3.modeIndex1Pair1)) Or Not(IsNumeric(dlg3.modeIndex2Pair1)) Then
					isInputOK = False
				End If
				' ---- mode pair 2 ----
				If Not(IsNumeric(dlg3.modeIndex1Pair2)) Or Not(IsNumeric(dlg3.modeIndex2Pair2)) Then
					isInputOK = False
				End If
				' ---- mode pair 3 ----
				If Not(IsNumeric(dlg3.modeIndex1Pair3)) Or Not(IsNumeric(dlg3.modeIndex2Pair3)) Then
					isInputOK = False
				End If

			Case 4
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg4.modeIndex1Pair1)) Or Not(IsNumeric(dlg4.modeIndex2Pair1)) Then
					isInputOK = False
				End If
				' ---- mode pair 2 ----
				If Not(IsNumeric(dlg4.modeIndex1Pair2)) Or Not(IsNumeric(dlg4.modeIndex2Pair2)) Then
					isInputOK = False
				End If
				' ---- mode pair 3 ----
				If Not(IsNumeric(dlg4.modeIndex1Pair3)) Or Not(IsNumeric(dlg4.modeIndex2Pair3)) Then
					isInputOK = False
				End If
				' ---- mode pair 4 ----
				If Not(IsNumeric(dlg4.modeIndex1Pair4)) Or Not(IsNumeric(dlg4.modeIndex2Pair4)) Then
					isInputOK = False
				End If

			Case 5
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg5.modeIndex1Pair1)) Or Not(IsNumeric(dlg5.modeIndex2Pair1)) Then
					isInputOK = False
				End If
				' ---- mode pair 2 ----
				If Not(IsNumeric(dlg5.modeIndex1Pair2)) Or Not(IsNumeric(dlg5.modeIndex2Pair2)) Then
					isInputOK = False
				End If
				' ---- mode pair 3 ----
				If Not(IsNumeric(dlg5.modeIndex1Pair3)) Or Not(IsNumeric(dlg5.modeIndex2Pair3)) Then
					isInputOK = False
				End If
				' ---- mode pair 4 ----
				If Not(IsNumeric(dlg5.modeIndex1Pair4)) Or Not(IsNumeric(dlg5.modeIndex2Pair4)) Then
					isInputOK = False
				End If
				' ---- mode pair 5 ----
				If Not(IsNumeric(dlg5.modeIndex1Pair5)) Or Not(IsNumeric(dlg5.modeIndex2Pair5)) Then
					isInputOK = False
				End If

			Case 6
				' ---- mode pair 1 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair1)) Or Not(IsNumeric(dlg6.modeIndex2Pair1)) Then
					isInputOK = False
				End If
				' ---- mode pair 2 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair2)) Or Not(IsNumeric(dlg6.modeIndex2Pair2)) Then
					isInputOK = False
				End If
				' ---- mode pair 3 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair3)) Or Not(IsNumeric(dlg6.modeIndex2Pair3)) Then
					isInputOK = False
				End If
				' ---- mode pair 4 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair4)) Or Not(IsNumeric(dlg6.modeIndex2Pair4)) Then
					isInputOK = False
				End If
				' ---- mode pair 5 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair5)) Or Not(IsNumeric(dlg6.modeIndex2Pair5)) Then
					isInputOK = False
				End If
				' ---- mode pair 6 ----
				If Not(IsNumeric(dlg6.modeIndex1Pair6)) Or Not(IsNumeric(dlg6.modeIndex2Pair6)) Then
					isInputOK = False
				End If

			End Select


			If Not(isInputOK) Then
				MsgBox "All mode indices must be positive integral values. Please check input.", _
					16, "Please check input"
			Else
				' ---- Convert input ----
				Select Case numModePairs
				Case 1
					modeIndices1(0) = CInt(dlg.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg.modeIndex2Pair1)
				Case 2
					modeIndices1(0) = CInt(dlg2.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg2.modeIndex2Pair1)
					modeIndices1(1) = CInt(dlg2.modeIndex1Pair2)
					modeIndices2(1) = CInt(dlg2.modeIndex2Pair2)
				Case 3
					modeIndices1(0) = CInt(dlg3.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg3.modeIndex2Pair1)
					modeIndices1(1) = CInt(dlg3.modeIndex1Pair2)
					modeIndices2(1) = CInt(dlg3.modeIndex2Pair2)
					modeIndices1(2) = CInt(dlg3.modeIndex1Pair3)
					modeIndices2(2) = CInt(dlg3.modeIndex2Pair3)
				Case 4
					modeIndices1(0) = CInt(dlg4.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg4.modeIndex2Pair1)
					modeIndices1(1) = CInt(dlg4.modeIndex1Pair2)
					modeIndices2(1) = CInt(dlg4.modeIndex2Pair2)
					modeIndices1(2) = CInt(dlg4.modeIndex1Pair3)
					modeIndices2(2) = CInt(dlg4.modeIndex2Pair3)
					modeIndices1(3) = CInt(dlg4.modeIndex1Pair4)
					modeIndices2(3) = CInt(dlg4.modeIndex2Pair4)
				Case 5
					modeIndices1(0) = CInt(dlg5.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg5.modeIndex2Pair1)
					modeIndices1(1) = CInt(dlg5.modeIndex1Pair2)
					modeIndices2(1) = CInt(dlg5.modeIndex2Pair2)
					modeIndices1(2) = CInt(dlg5.modeIndex1Pair3)
					modeIndices2(2) = CInt(dlg5.modeIndex2Pair3)
					modeIndices1(3) = CInt(dlg5.modeIndex1Pair4)
					modeIndices2(3) = CInt(dlg5.modeIndex2Pair4)
					modeIndices1(4) = CInt(dlg5.modeIndex1Pair5)
					modeIndices2(4) = CInt(dlg5.modeIndex2Pair5)
				Case 6
					modeIndices1(0) = CInt(dlg6.modeIndex1Pair1)
					modeIndices2(0) = CInt(dlg6.modeIndex2Pair1)
					modeIndices1(1) = CInt(dlg6.modeIndex1Pair2)
					modeIndices2(1) = CInt(dlg6.modeIndex2Pair2)
					modeIndices1(2) = CInt(dlg6.modeIndex1Pair3)
					modeIndices2(2) = CInt(dlg6.modeIndex2Pair3)
					modeIndices1(3) = CInt(dlg6.modeIndex1Pair4)
					modeIndices2(3) = CInt(dlg6.modeIndex2Pair4)
					modeIndices1(4) = CInt(dlg6.modeIndex1Pair5)
					modeIndices2(4) = CInt(dlg6.modeIndex2Pair5)
					modeIndices1(5) = CInt(dlg6.modeIndex1Pair6)
					modeIndices2(5) = CInt(dlg6.modeIndex2Pair6)
				End Select

				isInputOK = CheckInput()
			End If


			If isInputOK Then
				SelectTreeItem "Components"	'Result items should be deselected before starting the solver to avoid possible issues with the viewer.
				SelectModelView
				RunPP_CMA()
				exitDialog = True
			End If


		ElseIf iDialogResponse = 2 Then	'Help button

			StartHelp "common_preloadedmacro_mode_post-processing_for_the_characteristic_mode_analysis"
			exitDialog = True

		ElseIf iDialogResponse = 3 Then	'AddModePairButton

			If numModePairs < numModePairsMax Then
				dialogH += dialogHgroupBox + dialogSpacingLarge
				IncreaseSizeArray()
			End If

		ElseIf iDialogResponse = 4 Then	'RemoveModePairButton

			If numModePairs > 1 Then
				dialogH -= dialogHgroupBox + dialogSpacingLarge
				DecreaseSizeArray()
			End If

		ElseIf iDialogResponse = 0 Then	'Cancel button

			exitDialog = True

		End If

	Wend ' end dialog

	
End Sub


Function IncreaseSizeArray()
	ReDim Preserve modeIndices1(numModePairs + 1)
	ReDim Preserve modeIndices2(numModePairs + 1)

	modeIndices1(numModePairs) = 1
	modeIndices2(numModePairs) = 2

	numModePairs += 1
End Function

Function DecreaseSizeArray()
	ReDim Preserve modeIndices1(numModePairs - 1)
	ReDim Preserve modeIndices2(numModePairs - 1)

	numModePairs -= 1
End Function


Function RunPP_CMA()
	Dim indexModePair As Integer

	For indexModePair = 0 To numModePairs - 1 STEP 1
		With IESolver
			.AddModePairForPP_CMA(modeIndices1(indexModePair), modeIndices2(indexModePair))
		End With
	Next indexModePair

	With FDSolver
		.Start
	End With

	With IESolver
		.ClearModePairsForPP_CMA
	End With

End Function


' Checks all mode indices for valid values
' Returns True if all values are valid,
'		  False otherwise
Function CheckInput() As Boolean
	CheckInput = True

	Dim indexModePair As Integer
	Dim indexMode1 As Integer
	Dim indexMode2 As Integer

	Dim anyValueNegative As Boolean
	anyValueNegative = False
	Dim anyMode1AndMode2OfOnePairSame As Boolean
	anyMode1AndMode2OfOnePairSame = False


	For indexModePair = 0 To numModePairs - 1 STEP 1
		indexMode1 = modeIndices1(indexModePair)
		indexMode2 = modeIndices2(indexModePair)
		If indexMode1 < 1 Or indexMode2 < 1 Then
			anyValueNegative = True
		End If
		If indexMode1 = indexMode2 Then
			anyMode1AndMode2OfOnePairSame = True
		End If
	Next indexModePair

	If anyValueNegative Then
		MsgBox "All mode indices must be positive integral values. Please check input.", _
				16, "Please check input"
		CheckInput = False
	End If
	If anyMode1AndMode2OfOnePairSame Then
		MsgBox "Index 1 and index 2 of one mode must not be the same. Please check input.", _
				16, "Please check input"
		CheckInput = False
	End If
End Function
