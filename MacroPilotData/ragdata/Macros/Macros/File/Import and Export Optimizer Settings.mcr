'#Language "WWB-COM"
'#include "vba_globals_all.lib"

Option Explicit
' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 10-Mar-2022 mkr: First version
' ================================================================================================

Dim SettingsList() As String
Dim SettingString As String

Dim StudioList() As String
Dim TaskNameList() As String

Dim OptTypeString As String
Dim sFilename As String


Sub Main

	Dim nTasks As Integer
	nTasks = SimulationTask.StartTaskNameIteration

	Dim StudioInitLength As Integer
	If contains3D Then
		FillArray StudioList() ,  Array("3D", "Schematic\Home-Ribbon")
		StudioInitLength = 2
	Else
		FillArray StudioList() ,  Array("Schematic\Home-Ribbon")
		StudioInitLength = 1
	End If

	ReDim TaskNameList(2)
	TaskNameList(0) = ""
	TaskNameList(1) = ""

	Dim TaskName As String
	TaskName = SimulationTask.GetNextTaskName

	Dim iT As Integer
	iT = 0
	While TaskName <> ""
		If SimulationTask.GetTypeForTask(TaskName) = "Optimization" Then

			ReDim Preserve StudioList(iT+StudioInitLength+1)
			StudioList(iT+StudioInitLength) = "Schematic\" + TaskName

			ReDim Preserve TaskNameList(iT+StudioInitLength+1)
			TaskNameList(iT+StudioInitLength) = TaskName

			iT = iT + 1
		End If

		TaskName = SimulationTask.GetNextTaskName
	Wend

	FillArray SettingsList() ,  Array("All", "Parameters", "Goals", "Algorithm Settings")

	Begin Dialog UserDialog 400,168,"Import and Export Optimizer Settings",.dialogfunc ' %GRID:10,7,1,1
		CancelButton 300,140,80,21
		DropListBox 170,10,210,21,StudioList(),.iStudio
		PushButton 210,140,80,21,"Import",.ImportButton
		PushButton 120,140,80,21,"Export",.ExportButton
		Text 20,14,120,14,"Optimizer Location:",.Text1
		Text 20,49,100,14,"Optimizer Type:",.Text2
		OptionGroup .OptType
			OptionButton 180,49,90,14,"parametric",.Para
			OptionButton 180,70,120,14,"non-parametric",.Nonpara
		Text 20,105,120,14,"Handled Settings:",.Text3
		DropListBox 170,102,210,21,SettingsList(),.Setting
	End Dialog

	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then Exit All

End Sub



Function dialogfunc(DlgItem$, Action%, SuppValue%) _
	As Boolean
	Select Case Action%

		Case 1 ' Dialog Initialization

			DlgValue ("OptType",0)

			If contains3D Then
				If bDS Then
					DlgValue ("iStudio",1)
					DlgEnable  "OptType" , False
				End If
			Else
				DlgEnable  "OptType" , False
			End If

		Case 2 ' Value changing or button pressed

			If DlgValue ("OptType") = 0 Then
				OptTypeString = "Parametric"
			Else
				OptTypeString = "Non-Parametric"
			End If

			If DlgItem$ = "iStudio" Then

				If StudioList(DlgValue("iStudio")) = "3D" Then
					DlgEnable  "OptType" , True
				Else
					DlgValue ("OptType",0)
					DlgEnable  "OptType" , False
				End If

			End If

			If DlgItem$ = "OptType" Then

				Dim CurrentSetting As String
				CurrentSetting = DlgText("Setting")

				If DlgValue ("OptType") = 0 Then
					FillArray SettingsList() ,  Array("All", "Parameters", "Goals", "Algorithm Settings")
				Else
					FillArray SettingsList() ,  Array("All", "Parameters", "Design Responses")
				End If

				DlgListBoxArray("Setting", SettingsList)

				If CurrentSetting = "All" Then
						DlgValue ("Setting", 0)
				ElseIf CurrentSetting = "Parameters" Then
					DlgValue ("Setting", 1)
				End If

			End If


			If DlgItem$ = "ImportButton" Then

				sFilename = ""
				sFilename  = GetFilePath("", "txt", "", "Browse Optimizer Settings File", 0)

				If DlgValue ("Setting")= -1 Then
					MsgBox("Please choose a setting")
					dialogfunc = True
				ElseIf sFilename = ""  Then
					dialogfunc = True
				Else

					If StudioList(DlgValue("iStudio")) = "3D" Then
						Optimizer.ImportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					ElseIf StudioList(DlgValue("iStudio")) = "Schematic\Home-Ribbon" Then
						DSOptimizer.ImportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					Else
						DSOptimizer.SetSimulationType TaskNameList(DlgValue("iStudio"))
						DSOptimizer.ImportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					End If

				End If
			End If

			If DlgItem$ = "ExportButton" Then

				Dim sExportFolder , sFilePrefix As String
				sExportFolder = GetExportPathMaster_LIB()

				sFilePrefix = IIf( (StudioList(DlgValue("iStudio")) = "3D"), "3D", "Schematic")

				Dim i As Integer
				i=0

				Do
					i=i+1
					sFilename = sExportFolder + "OptimizerSettings_" + sFilePrefix + Format(i, "\_00") + ".txt"
				Loop Until Dir$(sFilename)=""

				If DlgValue ("Setting")= -1 Then
					MsgBox("Please choose a setting")
					dialogfunc = True
				Else

					If StudioList(DlgValue("iStudio")) = "3D" Then
						Optimizer.ExportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					ElseIf StudioList(DlgValue("iStudio")) = "Schematic\Home-Ribbon" Then
						DSOptimizer.ExportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					Else
						DSOptimizer.SetSimulationType TaskNameList(DlgValue("iStudio"))
						DSOptimizer.ExportSettings(OptTypeString, SettingsList(DlgValue ("Setting")), sFilename)
					End If

				End If

			End If


	End Select
End Function

Function contains3D As Boolean

	contains3D = (Right(GetApplicationName, 3) = "MWS") Or _
                 (Right(GetApplicationName, 2) = "CS") Or _
				  bEMS Or _
                 (Right(GetApplicationName, 11) = "EStatSolver") Or _
                  bPS Or _
				 (Right(GetApplicationName, 14) = "TrackingSolver") Or _
                  bMPS Or _
				 (Right(GetApplicationName, 11) = "TStatSolver")

End Function
