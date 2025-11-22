
' ================================================================================================
' This macro allows to set the steps per lambda used to write the Nearfield/FSM monitors in the I-solver
' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' -------------------
' 13-Sep-2017 dta: introduce sampling of fsm
' 31-Jan-2017 dta: corrected problem for projects with no monitors defined
' 07-Dec-2016 ube: dialogue cosmetics
' 18-Oct-2016 dta: more options introduced.
' 17-Mar-2014 dta: initial version
' ================================================================================================

Option Explicit

'#include "vba_globals_all.lib"

Dim N_Field2D_Monitors As Long
Dim Field2D_Monitorname() As String
Dim N_FS_Monitors As Long
Dim Field_Sampling As Boolean

Sub Main

Dim iii As Long, cst_iloop As Long

'=========================================
' calculate number of 2D Field monitors
'=========================================

N_Field2D_Monitors=0
N_FS_Monitors=0

cst_iloop=0
Field_Sampling=False

    For iii= 0 To Monitor.GetNumberOfMonitors-1
            If Monitor.GetMonitorTypeFromIndex(iii) = "E-Field 2D" Or Monitor.GetMonitorTypeFromIndex(iii) = "H-Field 2D" Or Monitor.GetMonitorTypeFromIndex(iii) = "E-Field 3D" Or Monitor.GetMonitorTypeFromIndex(iii) = "H-Field 3D" Then
                    N_Field2D_Monitors = N_Field2D_Monitors + 1

                    ReDim Preserve Field2D_Monitorname(N_Field2D_Monitors-1)
					Field2D_Monitorname(cst_iloop)=Monitor.GetMonitorNameFromIndex(iii)      'record in the array the 2D field monitor names
					'ReDim Preserve cst_ff_monitor_index((N_Field2D_Monitors-1))
					'cst_ff_monitor_index(cst_iloop)=iii								     'record in the array the 2D field monitor names

					If Field_Sampling=False Then
						If (Monitor.GetSubVolumeSampling(Field2D_Monitorname(cst_iloop))(0)="0" And Monitor.GetSubVolumeSampling(Field2D_Monitorname(cst_iloop))(1)="0" And Monitor.GetSubVolumeSampling(Field2D_Monitorname(cst_iloop))(2)="0") Then
							Field_Sampling=False
						Else
							Field_Sampling=True
						End If
					End If

					cst_iloop=cst_iloop+1

			ElseIf Monitor.GetMonitorTypeFromIndex(iii) = "Fieldsource" Then
					N_FS_Monitors=N_FS_Monitors+1

            End If
    Next iii

	If N_Field2D_Monitors=0 And N_FS_Monitors=0 Then   'pop up a message if no 2D Field Monitors, nor Field Source Monitors have been defined
			MsgBox("No Field Monitor or Field Source Monitor found. Please define at least one monitor before launching the macro.", 16, "Error")
			End
	End If
	
	Begin Dialog UserDialog 570,259,"Set Monitor Sampling in the I-Solver",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,14,550,154,"2D/3D Nearfield (E, H)",.GroupBox1
		GroupBox 10,175,550,49,"Field Source Monitor (.fsm)",.GroupBox2
		OKButton 20,231,90,21
		CancelButton 120,231,90,21
		TextBox 300,28,40,21,.StepsFieldMonitor
		GroupBox 20,56,530,105,"",.GroupBox4
		TextBox 330,73,40,21,.Global_Nx
		TextBox 410,73,40,21,.Global_Ny
		TextBox 490,73,40,21,.Global_Nz
		TextBox 300,189,40,21,.StepsFSM
		Text 30,196,250,21,"Set steps/wavelength globally",.Text2
		OptionGroup .SetSteps
			OptionButton 30,35,240,14,"Set steps/wavelength globally",.Step_lambda
			OptionButton 30,77,260,14,"Set absolute number of steps along",.Step_absolute
		Text 310,77,18,14,"X",.Text1
		Text 390,77,18,14,"Y",.Text3
		GroupBox 40,105,500,49,"Monitor Selection",.GroupBox3
		Text 470,77,18,14,"Z",.Text4
		DropListBox 250,126,280,21,Field2D_Monitorname(),.Field2DSelect
		CheckBox 50,126,160,14,"Apply to All Monitors",.FieldMonitorAll
	End Dialog
	Dim dlg As UserDialog


	If RestoreGlobalDataValue("StepsPerLambdaForFieldMonitor")="" Then
		dlg.StepsFieldMonitor="10"
	Else
		dlg.StepsFieldMonitor=RestoreGlobalDataValue("StepsPerLambdaForFieldMonitor")
	End If


	'default values
	dlg.Global_Nx="20"
	dlg.Global_Ny="20"
	dlg.Global_Nz="20"

	'StepX=CDbl(dlg.Global_Nx)
	'StepY=CDbl(dlg.Global_Ny)
	'StepZ=CDbl(dlg.Global_Nz)

	dlg.StepsFSM=RestoreGlobalDataValue("StepsPerLambdaForNFS")

	If (Dialog(dlg) = 0) Then Exit All

If dlg.SetSteps=0 Then			'set steps per lamba globally
	StoreGlobalDataValue("StepsPerLambdaForFieldMonitor", dlg.StepsFieldMonitor)

	'reset field monitor sampling definition
	For iii=0 To N_Field2D_Monitors-1
		With Monitor
		.ChangeSubVolumeSamplingToHistory(Field2D_Monitorname(iii), 0, 0, 0)
		End With
	Next iii

ElseIf dlg.SetSteps=1 And dlg.FieldMonitorAll=1 Then   'set absolute number of steps for all monitors
	For iii=0 To N_Field2D_Monitors-1
		With Monitor
		.ChangeSubVolumeSamplingToHistory(Field2D_Monitorname(iii), Cdbl(dlg.Global_Nx), Cdbl(dlg.Global_Ny), Cdbl(dlg.Global_Nz))
		End With
	Next iii

Else				'set absolute number of steps on selected monitor
	With Monitor
		.ChangeSubVolumeSamplingToHistory(Field2D_Monitorname(dlg.Field2DSelect), Cdbl(dlg.Global_Nx), Cdbl(dlg.Global_Ny), Cdbl(dlg.Global_Nz))
	End With

End If

StoreGlobalDataValue("StepsPerLambdaForNFS", dlg.StepsFSM)


End Sub

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	Select Case Action%
	    Case 1 ' Dialog box initialization
			If N_Field2D_Monitors=0 Then
				DlgEnable "SetSteps", False
				DlgEnable "StepsFieldMonitor", False
			End If

			If N_FS_Monitors=0 Then
				DlgEnable "StepsFSM", False
			End If

			If Field_Sampling Then
	    		DlgValue "FieldMonitorAll",0
	    		DlgValue "SetSteps",1
	    		DlgEnable "FieldMonitorAll", True
	    		DlgEnable "Global_Nx", True
				DlgEnable "Global_Ny", True
				DlgEnable "Global_Nz", True
				DlgEnable "FieldMonitorAll", True

				DlgText "Global_Nx", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(0))(0))
				DlgText "Global_Ny", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(0))(1))
				DlgText "Global_Nz", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(0))(2))

	    	Else
				DlgValue "FieldMonitorAll",1


	    		DlgEnable "FieldMonitorAll", False
				DlgEnable "Field2DSelect", False

				DlgEnable "Global_Nx", False
				DlgEnable "Global_Ny", False
				DlgEnable "Global_Nz", False

			End If

    	Case 2 ' Value changing or button pressed
    			Select Case DlgItem$
    				Case "SetSteps"
    					If SuppValue = 1 Then
							DlgEnable "Global_Nx", True
							DlgEnable "Global_Ny", True
							DlgEnable "Global_Nz", True
							DlgEnable "FieldMonitorAll", True
							DlgEnable "StepsFieldMonitor",False

						DialogFunc = True
		    			Else
		    			DlgEnable "StepsFieldMonitor",True
						DlgEnable "Global_Nx", False
						DlgEnable "Global_Ny", False
						DlgEnable "Global_Nz", False

						DlgEnable "FieldMonitorAll", False
						DlgEnable "StepsFieldMonitor",True
						DialogFunc = True
						End If
					Case "FieldMonitorAll"
						If SuppValue = 1 Then
							DlgEnable "Field2DSelect", False
						Else
							DlgEnable "Field2DSelect", True
							DlgText "Global_Nx", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(DlgValue("Field2DSelect")))(0))
							DlgText "Global_Ny", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(DlgValue("Field2DSelect")))(1))
							DlgText "Global_Nz", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(DlgValue("Field2DSelect")))(2))
						DialogFunc = True
						End If

					Case "Field2DSelect"
						If SuppValue <= N_Field2D_Monitors-1 Then
						DlgText "Global_Nx", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(SuppValue))(0))
						DlgText "Global_Ny", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(SuppValue))(1))
						DlgText "Global_Nz", CStr(Monitor.GetSubVolumeSampling(Field2D_Monitorname(SuppValue))(2))
						DialogFunc = True

						End If
    			End Select
    	Case 3 ' TextBox or ComboBox text changed
    	Case 4 ' Focus changed
    	Case 5 ' Idle
    	Case 6 ' Function key
    End Select

    If (Action%=1) Or (Action%=2) Then

		If (DlgValue("SetSteps")=1) And (DlgValue("FieldMonitorAll")=0) Then
			DlgEnable "Field2DSelect", True
		ElseIf (DlgValue("SetSteps")=0) And (DlgValue("FieldMonitorAll")=0) Then
			DlgEnable "Field2DSelect", False
		End If

	End If

End Function

