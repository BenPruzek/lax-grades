'--------------------------------------------------------------------------------------------
' This macro uses DS-Simulation Projects to compare 3D-Solvers and Meshes
'--------------------------------------------------------------------------------------------
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 17-Apr-2020 ube: fixed writing to Schematic Results tree, removed TLM20
' 29-Feb-2016 ube: changed TLM20 to 15/20/15 (same as FIT)
' 06-Jan-2016 ube: TLM20 test
' 29-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 26-Nov-2013 ube: added tet mesh equilibration test
' 14-Feb-2012 ube: Fres solver is now running gen-purpose mesh adaptation and broadband sweep with MOR
' 08-Jan-2012 ube: if AR-S-Parameters exist, those will be compared
' 23-Sep-2011 ube: added more solvers: M, I, Fres, TLM
' 20-Jul-2011 ube: now automatically closes Task-windows after simulation
'					(to not end up in too many modelers and windows open)
' 24-May-2011 ube: PBA comparison did not work
' 20-May-2011 ube: first version
'--------------------------------------------------------------------------------------------
Option Explicit
'#include "vba_globals_all.lib"

Sub Main ()

	Begin Dialog UserDialog 570,336,"Compare Solvers and Meshes" ' %GRID:10,7,1,1
		OKButton 10,308,90,21
		CancelButton 110,308,90,21

		GroupBox 10,7,550,105,"TD (Hexahedral Mesh)",.GroupT
		Text 30,28,20,14,"T",.Text2
		Text 30,49,60,14,"Tadapt",.Text3
		'Text 30,70,120,14,"TnewMesh",.Text4
		CheckBox 100,28,220,14,"FIT Solver",.T
		CheckBox 100,49,320,14,"Energy based mesh adaptation (max 3 passes)",.Tadapt

		GroupBox 10,112,550,112,"FD (Tetrahedral Mesh)",.GroupF
		Text 30,133,50,14,"F",.Text5
		Text 30,154,60,14,"Fflat",.Text6
		CheckBox 100,133,430,14,"default (F-general purpose sweep, curved elements)",.F
		CheckBox 100,154,300,14,"Flat (non-curved) tetrahedral elements",.Fflat

		GroupBox 10,231,550,63,"Specialized Solvers",.GroupBox1
		Text 30,175,50,14,"Fres",.Text1
		CheckBox 100,175,440,14,"F-fast resonant sweep, curved elements",.Fres
		Text 30,252,40,14,"M",.Text4
		CheckBox 100,252,180,14,"Multilayer Solver",.M
		Text 30,273,30,14,"I",.Text7
		CheckBox 100,273,180,14,"Integral Equation",.I
		Text 30,70,50,14,"TLM",.Text8
		CheckBox 100,70,200,14,"TLM Solver",.TLM
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) = 0) Then Exit All

	If (GetApplicationName = "DS") Then
		MsgBox "Currently this macro is only supported in the canvas (schematic editor) of a MWS-3D-Project.", vbExclamation
		Exit All
	End If

	' save project in order to make sure, that latest cst-blocks are stored
	Save

	Dim sHistoryTitle As String, sAddedVBACommands As String

	If dlg.T Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "Mesh.FPBAAccuracyEnhancement ""enable"" " + vbCrLf

		SetupAndRunSimulationProject("T", "HF_TRANSIENT", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.F Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "With MeshSettings" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .SetMeshType ""Tet"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .Set ""CurvatureOrder"", ""3"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "FDSolver.OrderTet ""Second"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "FDSolver.SetMethod ""Tetrahedral"", ""General purpose"" " + vbCrLf

		SetupAndRunSimulationProject("F", "HF_FREQUENCYDOMAIN", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.M Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "With MeshSettings" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf

		SetupAndRunSimulationProject("M", "HF_MULTILAYER", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.Fres Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "Mesh.MeshType ""Tetrahedral"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "With MeshSettings" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .SetMeshType ""Tet"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .Set ""CurvatureOrder"", ""3"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "FDSolver.SetMethod ""Tetrahedral"", ""Fast reduced order model"" " + vbCrLf

		SetupAndRunSimulationProject("Fres", "HF_FREQUENCYDOMAIN", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.I Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "With MeshSettings" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf

		SetupAndRunSimulationProject("I", "HF_INTEGRALEQUATION", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.TLM Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "Mesh.MeshType ""HexahedralTLM"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "Solver.Method ""Hexahedral TLM"" " + vbCrLf

		SetupAndRunSimulationProject("TLM", "HF_TRANSIENT", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.Fflat Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "With MeshSettings" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .SetMeshType ""Tet"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "  .Set ""CurvatureOrder"", ""1"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "FDSolver.OrderTet ""Second"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "FDSolver.SetMethod ""Tetrahedral"", ""General purpose"" " + vbCrLf

		SetupAndRunSimulationProject("Fflat", "HF_FREQUENCYDOMAIN", sHistoryTitle, sAddedVBACommands)
	End If

	If dlg.Tadapt Then
		sHistoryTitle = "set mesh properties"
		sAddedVBACommands = ""
		sAddedVBACommands = sAddedVBACommands + "Mesh.FPBAAccuracyEnhancement ""enable"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "With MeshAdaption3D" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "    .SetType ""Time"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "    .SetAdaptionStrategy ""Energy"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "    .MinPasses ""2"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "    .MaxPasses ""3"" " + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "End With" + vbCrLf
		sAddedVBACommands = sAddedVBACommands + "Solver.MeshAdaption ""True"" " + vbCrLf

		SetupAndRunSimulationProject("Tadapt", "HF_TRANSIENT", sHistoryTitle, sAddedVBACommands)
	End If

End Sub
Sub SetupAndRunSimulationProject (sTaskName As String, sSolverType As String, sHistoryTitle As String, sAddedVBACommands As String)

	Dim sTimer As String, dTimer As Double

	With SimulationProject
	    .ResetComponents
	    .SetAllComponents "3D"
	    .SetAllPorts "SCHEMATIC"

		.LoadReferenceDataFromBlock "Block: MWSSCHEM1"
		.SetUseReferenceData True

		Dim iii As Integer, s1 As String
		iii = 0
		s1 = sTaskName
		While .DoesExist(sTaskName)
			iii = iii+1
			sTaskName = s1+Cstr(iii)
		Wend
		If .DoesExist(sTaskName) Then
			' .Delete sTaskName
		End If

		.Create "MWS", sTaskName
		.SetSolverType sSolverType
		.Get3D.AddToHistory sHistoryTitle, sAddedVBACommands
		.EndCreation

		sTimer = Timer_LIB("reset",0.0)
		On Error Resume Next
		.Run sTaskName
		On Error GoTo 0
		sTimer = Timer_LIB("",dTimer)

		CopyResultsForComparison(sTaskName, sTimer)

		.Close sTaskName, False	' false means: save it before closing (do not discard any changes)
								' boolean discard flag specifies whether potentially unsaved project modifications shall be ignored.

	End With

End Sub
Sub CopyResultsForComparison(sTaskName As String, sTimer As String)

	Dim sTaskTime As String
	If sTimer <> "" Then
		sTaskTime = sTaskName + "-" + sTimer
	Else
		sTaskTime = sTaskName
	End If

	Dim result As Object
	Set result = DS.Result1DComplex("")

	Dim sfile As String, sfile2 As String, ssname As String, sPath As String

	sPath = DSResultTree.GetFirstChildName("Tasks\"+sTaskName+"\3D Model Results\1D Results\S-Parameters (AR)")
	If sPath="" Then
		' no AR-Results exist, take normal S-Parameters
		sPath = DSResultTree.GetFirstChildName("Tasks\"+sTaskName+"\3D Model Results\1D Results\S-Parameters")
	End If

	While sPath<>""
		'MsgBox sPath
		ssname = Mid$(sPath,1+InStrRev(sPath,"\"),)

		sfile  = DSResultTree.GetFileFromTreeItem(sPath)

		With result
			.Load(sfile)
			.Save GetProjectPath("Temp")+"\"+ NoForbiddenFilenameCharacters(sPath)+".sig"
			.AddToTree ("Results\"+ssname+"\"+Replace(sTaskTime,".",","))
		End With
		sPath = DSResultTree.GetNextItemName(sPath)
	Wend
End Sub
