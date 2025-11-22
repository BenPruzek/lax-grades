Option Explicit

'------------------------------------------------------------------------------
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' =============================================================================
' History of Changes
'------------------------------------------------------------------------------
' 23-Feb-2024 mha: Added in additional text pointing out new options
' 29-Sep-2021 mha: Added in text concerning material classes (explain difference between priorization schemes)
' 24-Apr-2020 ncu: remove "T-Compatible-F" since it is now default
' 17-Jul-2018 ube: included scheme "T-Compatible-F"
' 05-Feb-2015 osk: new vba-commands (no longer called "debug")
' 05-Feb-2013 ube: Initial version
'------------------------------------------------------------------------------

' define global variables
Dim lChoice As Long

Sub Main ()
	' Activate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings True

	' Clear the data
	ClearScriptSettings
	DS.ClearScriptSettings

	' Call the define method and check whether it is completed successfully
	If Define("test", True, False) Then
		' Determine current mesh scheme settings:
		Dim schemeName As String
		With MeshSettings
			.SetMeshType "All"
			schemeName = .Get("PrioritizationScheme")
		End With

		Dim schemeSelector As Integer
		If schemeName = "ElectricConductivity" Then
			schemeSelector = 1
		ElseIf schemeName = "ThermalConductivity" Then
			schemeSelector = 2
		Else
			schemeSelector = 0
		End If

	    If schemeSelector <> lChoice Then
			' the setting has been changed...
			Dim sCommand As String

			' set up history list entry...
			sCommand = ""
			sCommand = sCommand + "With MeshSettings " + vbLf
			sCommand = sCommand + ".SetMeshType ""All""" + vbLf

			Select Case lChoice
			Case 0
				sCommand = sCommand + ".Set ""PrioritizationScheme"", ""DefaultHF""" + vbLf
			Case 1
				sCommand = sCommand + ".Set ""PrioritizationScheme"", ""ElectricConductivity""" + vbLf
			Case 2
				sCommand = sCommand + ".Set ""PrioritizationScheme"", ""ThermalConductivity""" + vbLf
			End Select

			sCommand = sCommand + "End With"

			AddToHistory "(*) set mesh material priority", sCommand
	    End If
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False
End Sub



' -------------------------------------------------------------------------------------------------
' Define: This function defines the look of the dialog box
' -------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean) As Boolean
	Dim sPleaseNoteText As String
	sPleaseNoteText = ""
	sPleaseNoteText = sPleaseNoteText & "Please note that ""Electric Conductivity Sorting"" and ""Default"" can also be set directly in the CST Studio Suite GUI without any need for macro intervention." & vbCrLf
	sPleaseNoteText = sPleaseNoteText & "In order to do so, simply go to ""Simulation: Mesh > Mesh View"" and when the ""Mesh"" tab opens up to ""Mesh: Mesh Control > Advanced Control > Mesh Priority"" in order to call up the ""Mesh Priority"" dialog." & vbCrLf
	sPleaseNoteText = sPleaseNoteText & "This dialog allows to select a priority scheme and directly view the priorization as concerning shapes and material."

	Begin Dialog UserDialog 660, 346, "TET Mesh - Define Material Priority", .DialogFunc ' %GRID:10,7,1,1
		GroupBox      10,   7, 640, 110, "Please note",   .GBNote
		Text          30,  23, 600,  92, sPleaseNoteText, .TNote
		GroupBox      10, 117, 640, 194, "Options", .GroupBox4
		Text          30, 138, 560,  49, "In case of material overlapping, meshing process might require Boolean operations."+vbCrLf+vbCrLf+"Please choose the Material Priority scheme, avoiding manual Boolean operations:",.Text1
		OptionGroup .GroupPrio
			OptionButton 40, 194, 410, 14, "Default (same rules as for Transient Solver)",.none
			OptionButton 40, 215, 585, 14, "Electric Conductivity Sorting (larger conductivity value overwrites smaller value)", .electric
			OptionButton 40, 260, 585, 14, "Thermal Conductivity Sorting (larger conductivity value overwrites smaller value)",.thermal
			Text         60, 230, 585, 25, "(Material priority is executed inside individual material classes respectively. In general lossy metal has priority over PEC, PEC has priority of normal materials.)",.Text2
			Text         60, 275, 585, 25, "(Material priority is not constrained by material class. E.g. normal material with high thermal conductivity has priority to lossy metal with lower thermal conductivity.)",.Text3
		OKButton      10, 318, 90,  21
		CancelButton 120, 318, 90,  21
	End Dialog

	Dim dlg As UserDialog

	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		' In case of a result template settings would be stored as ScriptSettings, to be retreived again later. Can also be used for macros if settings should be stored!
		lChoice = dlg.GroupPrio
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunc(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Select Case iAction
		Case 1 ' Dialog box initialization
			' Grey out, enable, initialize...
			DlgValue  "GroupPrio", 0
		Case 2 ' Value changing or button pressed
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function
