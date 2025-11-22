'#Language "WWB-COM"
'#include "vba_globals_all.lib"

' This macro export geometries under in the current CST project to .sat or .sab files, and created a python script for Abaqus to import them with material properties
' When PCB geometries are detected, more options are offered to combine them into fewer shapes

' --------------------------------------------------------------------------------------------------------------
' Copyright 2022-2025 Dassault Systemes Deutschland GmbH
' ==============================================================================================================
' History of Changes
' ------------------
' 06-Feb-2025 ube: Help button activated.
' 19-Jun-2024 mha: Corrected wrong factor sYMFactor for Youngs Modulus to SI.
' 27-Oct-2023 mha: Switched permittivity, permeability from relative to absolute in abaqus import file
' 26-Oct-2023 mha: added units to csv-list and following material properties: RelativePermittivity, RelativePermeability, ElectricConductivity
' 25-Oct-2023 mha: added explicit query for PEC, in case of PEC copper properties are reported (added corresponding user-inaccessible option in dialog)
' 29-Sep-2023 mha: added option for export regarding Abaqus unit system (SI or mm-tonne-seconds)
' 22-Mar-2023 wxu: fixed an issue when layer option is used, some VIAs might have partially mached names (line 289->290)
' 07-Mar-2023 wxu: fixed a filename issue, under ExportAll, (instr ->InstrRev), add a preventive line to remove "/"
' 27-Feb-2023 wxu: fixed a problem: previous If command would skip material that partially match (in ExportAllUnder function)
' 15-Dec-2022 wxu: fixed an issue with ":" in the name in function "AbaqusName"
' 28-Nov-2022 wxu: Initial version
' ==============================================================================================================

Option Explicit

Private Type OptionsGeometry
	sExportFolder     As String
	nPCBType          As Integer	'0: do not exist; 1: regular detailed PCB; [-1: simplified PCB]
	nTraceOptions     As Integer	'0: one part; 1: net 2: layer 3: as is
	nSubstrateOptions As Integer	'0: one part; 1: slice each layer, 2: as is (will have to boolean out traces)
	bMaterial         As Boolean
'	bHeatSource       As Boolean
	bSelected         As Boolean
	nUnits            As Integer	'0: SI, 1: mm-tonne-seconds
	bPreserve         As Boolean
	bAbaqusScript     As Boolean
End Type

Dim GeomSettings          As OptionsGeometry
Dim sPartsNames()         As String
Dim nPCB                  As Long	'number of PCBs
Dim CSTVersion            As Integer
Dim dUnitConversionFactor As Double

Dim StreamNum As Integer, StreamNum_Script As Integer

Public Const sTitle = "Transfer Model to Abaqus"

Const PEC_replace_rho As Double = 8930.0
Const PEC_replace_kx As Double = 401.0
Const PEC_replace_Cp As Double = 390
Const PEC_replace_YM As Double = 120
Const PEC_replace_PR As Double = 0.33
Const PEC_replace_CTE As Double = 17
Const PEC_replace_EpsX As Double = 1
Const PEC_replace_MuX As Double = 1
Const PEC_replace_SigmaX As Double = 5.8e7


Sub Main

	CSTVersion = Cint(Mid(GetApplicationVersion,9,4))
	dUnitConversionFactor = 1
	nPCB = 0

	Begin Dialog UserDialog 580, 397, sTitle, .DialogFunc ' %GRID:10,7,1,1

		Text  90, 14, 380, 21, "This template exports geometries and properties to Abaqus", .Text4
		Text 110, 35, 350, 21, "PCB options are only available when a PCB is detected",     .Text5

		GroupBox  20,  77, 540, 84, "Model Export Options",               .GroupBox2
		CheckBox  40,  98, 190, 14, "Export material properties",         .ckMaterial
'		CheckBox  40, 116, 170, 14, "Export heat sources",                .ckHeatSource
		OptionGroup                                                       .Units
			OptionButton  40, 119, 170, 14, "Export to SI unit system",   .Units_SI
			OptionButton 300, 119, 202, 14, "Export to mm-tonne-seconds", .Units_mm_tonne_s
		CheckBox  40, 140, 220, 14, "Export selected geometries only",    .ckSelected
		CheckBox 300, 140, 240, 14, "Export PEC with copper properties",      .CBPECAsCopper

		GroupBox 20, 166, 540, 119, "PCB Export Options", .GroupBox1
		Text     30, 183, 130,  14, "Traces",             .Text1
		OptionGroup                                                            .Trace
			OptionButton 30, 201, 220, 14, "Combine all traces into one part", .opt1
			OptionButton 30, 222, 160, 14, "Each layer as a part",             .Opt2
			OptionButton 30, 243, 160, 14, "Each net as a part",               .opt3
			OptionButton 30, 264,  70, 14, "As is",                            .Opt4

		Text 300, 183, 130, 14, "Substrate", .Text2
		OptionGroup                                                                 .Substrate
			OptionButton 300, 201, 250, 14, "Combine all dielectric into one part", .Ops1
			OptionButton 300, 222, 250, 14, "As is",                                .Ops2

		CheckBox  30, 292, 220, 14, "Preserve original geometries", .ckPreserve
		CheckBox 310, 292, 220, 14, "Create Abaqus Import Script",  .ckAbaqusScript

		TextBox       20, 341, 470, 21,                     .tbExportPath
		Text          20, 320, 100, 18, "Export to folder", .Text3
		PushButton   500, 341,  70, 21, "Browse...",        .PBBrowse
		PushButton   470, 369, 100, 21, "Help",             .Help
		CancelButton 360, 369, 100, 21
		OKButton     250, 369, 100, 21
	End Dialog

	Dim dlg As UserDialog

	dlg.tbExportPath   = GetExportPathMaster_LIB() + "ABQ\"
	dlg.ckMaterial     = True
'	dlg.ckHeatSource   = True
	dlg.Units          = 0
	dlg.ckPreserve     = True
	dlg.ckAbaqusScript = True
	dlg.Trace          = 0
	dlg.Substrate      = 0
	dlg.ckSelected     = False

	GeomSettings.nPCBType = PCBType()


	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function
		' will cause the framework to cancel the creation or modification without storing
		' anything.
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function
		' will cause the framework to complete the creation or modification and store the corresponding
		' settings.

		GeomSettings.sExportFolder     = dlg.tbExportPath
		GeomSettings.nTraceOptions     = dlg.Trace
		GeomSettings.nSubstrateOptions = dlg.Substrate
		GeomSettings.bMaterial         = dlg.ckMaterial
'		GeomSettings.bHeatSource       = dlg.ckHeatSource
		GeomSettings.bSelected         = dlg.ckSelected
		GeomSettings.nUnits            = dlg.Units
		GeomSettings.bPreserve         = dlg.ckPreserve
		GeomSettings.bAbaqusScript     = dlg.ckAbaqusScript

		If ExecuteExport() = 1 Then
			ReportInformation("Done")
			Shell Environ("windir")+ "\explorer.exe """ & GeomSettings.sExportFolder & "", vbNormalFocus
		End If
	End If
End Sub



Function DialogFunc(DlgItem As String, action As Integer, Value As Integer) As Boolean
	Select Case action
		Case 1 ' Dialog box initialization
'			DlgVisible("ckHeatSource",False)
			DlgEnable("ckSelected", False)
			DlgValue("CBPECAsCopper", 1)
			DlgEnable("CBPECAsCopper", False)

			DlgEnable("Units", True)
			DlgValue("Units", 0)
			If GeomSettings.nPCBType = 0 Then	'0: do not exist; 1: regular detailed PCB; -1: simplified PCB
				DlgEnable("Trace", False)
				DlgEnable("Substrate", False)
				DlgEnable("opt1",False)
				DlgEnable("opt2",False)
				DlgEnable("opt3",False)
				DlgEnable("opt4",False)
				DlgEnable("ops1",False)
				DlgEnable("ops2",False)
			End If

		Case 2 ' Value changing or button pressed
			Select Case DlgItem
				Case "Units"
					If DlgValue("Units") = 1 Then
						' mm-tonne-seconds
						dUnitConversionFactor = 1000
						GeomSettings.nUnits   = 1
					Else
						' SI
						dUnitConversionFactor = 1
						GeomSettings.nUnits   = 0
					End If
					DialogFunc = True
				Case "PBBrowse"
					Dim sTmp As String
					 sTmp = GetFolder_Lib(GeomSettings.sExportFolder)
					 If sTmp <> "" Then GeomSettings.sExportFolder = sTmp
					DlgText("tbExportPath", GeomSettings.sExportFolder)
					DialogFunc = True
				Case "OK"
					DialogFunc = False
				Case "Help"
					StartHelp "common_preloadedmacro_CST_Abaqus_Connector"
					DialogFunc = True
			End Select

		Case 3 ' ComboBox or TextBox Value changed
			DialogFunc = True
		Case 4 ' Focus changed
		Case 5 ' Idle
	End Select
End Function



Function ExecuteExport() As Integer
	ScreenUpdating ( False )
	CST_MkDir GeomSettings.sExportFolder

	StreamNum = FreeFile

	Open GeomSettings.sExportFolder + "Material_Summary.csv" For Output As #StreamNum

	If GeomSettings.bAbaqusScript Then
		StreamNum_Script = FreeFile

		Open GeomSettings.sExportFolder + "Abaqus_import_CST.py" For Output As #StreamNum_Script

		'	i = Material.GetNumberOfMaterials()
		'	Material.GetThermalConductivity("Air", kx, ky,kz)
		'	Material.GetYoungsModulus("Copper (annealed)", kx)
		'	Material.GetPoissonsRatio("Copper (annealed)", ky)
		'	Material.GetThermalExpansionRate("Copper (annealed)", kz)
		Print #StreamNum_Script, "# -*- coding: mbcs -*-"
		Print #StreamNum_Script, "from part import *"
		Print #StreamNum_Script, "from material import *"
		Print #StreamNum_Script, "from section import *"
		Print #StreamNum_Script, "from assembly import *"
		Print #StreamNum_Script, "from step import *"
		Print #StreamNum_Script, "from interaction import *"
		Print #StreamNum_Script, "from load import *"
		Print #StreamNum_Script, "from mesh import *"
		Print #StreamNum_Script, "from optimization import *"
		Print #StreamNum_Script, "from job import *"
		Print #StreamNum_Script, "from sketch import *"
		Print #StreamNum_Script, "from visualization import *"
		Print #StreamNum_Script, "from connectorBehavior import *"
		Print #StreamNum_Script, "import regionToolset"
		Print #StreamNum_Script, "import os"
		'Print #StreamNum_Script, "os.chdir(os.getcwd())" trying to change Abaqus current working directory, but not working...
	End If

	If GeomSettings.nPCBType = 0 Then 'no detailed PCB, export all parts one by one
		If GeomSettings.bSelected Then
			ExportAllUnder(GetSelectedTreeItem)
		Else
			ExportAllUnder("Components")
		End If

	Else 'detailed PCBs
		Dim sPCB As String, sPCBRoot() As String

		sPCBRoot = GetAllPCBs()

		For Each sPCB In sPCBRoot
			nPCB = nPCB + 1
			ExportPCB(sPCB)
		Next
	End If

	Close #StreamNum
	If GeomSettings.bAbaqusScript Then	Close #StreamNum_Script

	ExecuteExport = 1
	ReportInformation("Model and script are located at " + GeomSettings.sExportFolder)
	ScreenUpdating ( True )
	MsgBox "Geometries and scripts have been exported to " + vbNewLine + _
			GeomSettings.sExportFolder + vbNewLine + _
			"Please set Abaqus CAE Working Directory to this folder" + vbNewLine + _
			"and use File->Run Script to import the model", vbInformation, "Model Export Finished"
End Function



Function ExportPCB(sPCB As String) As Integer
	Dim sCurrPCB As String, sCurrPCBPartsName As Variant
	Dim i As Long, j As Long

	If GeomSettings.bPreserve Then 'keeping original parts, making copies to boolean and export
		sCurrPCB = "#" + sPCB + "#"
		Component.New (sCurrPCB)
		With Transform
		    .Reset
		    .Name sPCB
		    .Vector "0", "0", "0"
		    .UsePickedPoints "False"
		    .InvertPickedPoints "False"
		    .MultipleObjects "True"
		    .GroupObjects "False"
		    .Repetitions "1"
		    .MultipleSelection "False"
		    .Destination sCurrPCB
		    .Material ""
		    .Transform "Shape", "Translate"
		End With
	Else
		sCurrPCB = sPCB
	End If

	'------------ Trace -------------------------------------------------------------------------------
	If GeomSettings.nTraceOptions = 0 Then '0: one part;
		ReportInformation("Boolean combine all traces together, this may take a while...")

		sCurrPCBPartsName = NamesUnder(sCurrPCB + "/Nets")

		For i = 1 To UBound(sCurrPCBPartsName)
			Solid.Add(sCurrPCBPartsName(0), sCurrPCBPartsName(i))
			ReportInformation("Working on part " + cstr(i) + " of " + cstr(UBound(sCurrPCBPartsName)) + "...")
		Next

		ReportInformation("Boolean combine all traces done!")
		Solid.Rename sCurrPCBPartsName(0), sCurrPCB + "/Nets:All"
		Component.DeleteAllEmptyComponents ( sCurrPCB + "/Nets")

	ElseIf GeomSettings.nTraceOptions = 1 Then '1: Layers
		Dim sLayerNames As Variant, sCurrentLayer As String
		ReportInformation("Boolean combine layers one by one, this may take a while...")

		sLayerNames = GetAllLayerNames(sCurrPCB)
		sCurrPCBPartsName = NamesUnder(sCurrPCB + "/Nets")

		For j = 0 To UBound(sLayerNames)
			ReportInformation("Working on Layer " + cstr(j) + " of " + cstr(UBound(sLayerNames))+ "...")
			sCurrentLayer = ""
			For i = 0 To UBound(sCurrPCBPartsName)
				'If InStr(sCurrPCBPartsName(i), sLayerNames(j)) >0 Then
				If Split(sCurrPCBPartsName(i),":")(1) = sLayerNames(j) Then '2023-03-22 fix: name of parts can be partially match, fails on a VIA part: VIA_SIGNAL_1_GND_2_1, because VIA_SIGNAL_1_GND_2 partially matches the name
					If sCurrentLayer ="" Then
						sCurrentLayer = sCurrPCBPartsName(i)
					Else
						Solid.Add(sCurrentLayer, sCurrPCBPartsName(i))
					End If
				End If
			Next i
			Solid.Rename(sCurrentLayer, sCurrPCB + "/Nets:" + sLayerNames(j))
		Next j
		Component.DeleteAllEmptyComponents ( sCurrPCB + "/Nets")

	ElseIf GeomSettings.nTraceOptions = 2 Then '2: Net
		Dim sNetNames As Variant, sNetParts As Variant

		ReportInformation("Boolean combine nets one by one, this may take a while...")
		sNetNames = GetAllNetNames(sCurrPCB)

		For i = 0 To UBound(sNetNames)
			sNetParts = NamesUnder(sNetNames(i))
			ReportInformation("Working on Net " + cstr(i) + " of " + cstr(UBound(sNetNames)) + "...")
			For j = 1 To UBound(sNetParts)
				Solid.Add(sNetParts(0), sNetParts(j))
			Next j
			Solid.Rename(sNetParts(0), Left(sNetNames(i),InStrRev(sNetNames(i),"/")-1) + ":" + Mid(sNetNames(i), InStrRev(sNetNames(i),"/")+1))
		Next i
		Component.DeleteAllEmptyComponents ( sCurrPCB + "/Nets")

	ElseIf GeomSettings.nTraceOptions = 3 Then '3: As Is
			'do nothing
	End If
	' ------------------------ Substrate ---------------------------------------------------------
	Dim	sSubNames As Variant

	sSubNames = NamesUnder(sCurrPCB + "/Substrates")
	If GeomSettings.nSubstrateOptions = 0 Then 'one block
		For i = 1 To UBound(sSubNames)
			Solid.Add(sSubNames(0), sSubNames(i))
		Next
		ReportInformation("Boolean combine all substrate done!")

		Solid.Rename sSubNames(0), sCurrPCB + "/Substrates:All"
	ElseIf GeomSettings.nSubstrateOptions = 1 Then 'as is
		'Do nothing
	End If

	Dim sSub As String, sNet As String

	sSubNames = NamesUnder(sCurrPCB + "/Substrates")
	sNetNames = NamesUnder(sCurrPCB + "/Nets")

	ReportInformation("Boolean insert traces into substrates...")

	For Each sSub In sSubNames
		For Each sNet In sNetNames
			Solid.Insert(sSub, sNet)
		Next
	Next

	ReportInformation("Exporting...")

	ExportAllUnder(sCurrPCB)

	If GeomSettings.bPreserve Then Component.Delete(sCurrPCB)  'bPreserve is on, then the sCurrPCB is a copy
End Function



Function ExportAllUnder(sRoot As String) As Integer
	Dim i As Long, sCADFileName As String, sMatFileName As String, sScriptFileName As String, sAbaqusPartName As String
	Dim sParts As Variant
	Dim sMat() As String, sTmp As String, iMat As Integer, sAbaqusMatName As String
	Dim rho As Double, kx As Double, ky As Double, kz As Double, Cp As Double
	Dim EpsX As Double, EpsY As Double, EpsZ As Double, MuX As Double, MuY As Double, MuZ As Double, SigmaX As Double,SigmaY As Double,SigmaZ As Double
	Dim YM As Double, PR As Double, CTE As Double
	Dim sYMFactor As String



	sParts = NamesUnder(sRoot)
	' loop over all solids
	ReDim sMat(0)
	iMat = 0
	sMat(0) = Chr(14)  'something impossible

	Dim ACISFileExt As String

	If CSTVersion < 2023 Then 	'v2022 and earlier had a bug that always export sat file
		ACISFileExt = ".sat"
	Else
		ACISFileExt = ".sab"
	End If

	For i = 0 To UBound(sParts)
		'remove the assembly name
		sCADFileName = Right(sParts(i), Len(sParts(i)) - InStrRev(sParts(i),"/")) '2023-03-07 fix
		'add a sequence number, and ensure parts do not start with a number
		If nPCB >0 Then
			sCADFileName = "PCB" + cstr(nPCB) + "_"+ cstr(i) + "_" + sCADFileName
		Else
			sCADFileName = "P" + cstr(i) + "_" + sCADFileName
		End If

		sCADFileName = AbaqusName(sCADFileName)
		sAbaqusPartName = Chr(39) + sCADFileName + Chr(39)

		With SAT
			.Reset
			.FileName (GeomSettings.sExportFolder + sCADFileName + ACISFileExt)
			.SaveVersion ("31.0") 'this version works with Abaqus 2022
			.ExportFromActiveCoordinateSystem (False)
			.Write (sParts(i))
		End With

		If GeomSettings.bAbaqusScript Then
			Print #StreamNum_Script, "# ----- Next part -----"
			Print #StreamNum_Script, "mdb.openAcis(" + Chr(39) + sCADFileName + ACISFileExt + Chr(39) +"," + "scaleFromFile=OFF)"
			Print #StreamNum_Script, "mdb.models[" + Chr(39) +"Model-1" + Chr(39) + "].PartFromGeometryFile(combine=False,"
	    	Print #StreamNum_Script, "	dimensionality=THREE_D, geometryFile=mdb.acis, mergeSolidRegions=True,"
	    	Print #StreamNum_Script, "	name=" + sAbaqusPartName + ","
	    	Print #StreamNum_Script, "	scale=" + cstr(Units.GetGeometryUnitToSI*dUnitConversionFactor)+ ", type=DEFORMABLE_BODY)"
		End If

	    sTmp = Solid.GetMaterialNameForShape(sParts(i))

	   	If nPCB >0 Then
			sAbaqusMatName = "PCB" + cstr(nPCB) + "_" + Right(sTmp, Len(sTmp) - InStr(sTmp,"/"))
		Else
			sAbaqusMatName = sTmp
	   	End If

		sAbaqusMatName = AbaqusName(sAbaqusMatName)
	   	sAbaqusMatName = Chr(39) +  sAbaqusMatName + Chr(39)

		If FindListIndex(sMat, sTmp) = -1 Then 'new material, bug fix 2/27/23, previous If command would skip material that partially match
			ReDim Preserve sMat(iMat)
			sMat(iMat) = sTmp
			iMat = iMat + 1

			If LCase(Material.GetTypeOfMaterial(sTmp)) = LCase("PEC") Then
				rho = PEC_replace_rho
				kx = PEC_replace_kx
				ky = PEC_replace_kx
				kz = PEC_replace_kx
				Cp = PEC_replace_Cp
				YM = PEC_replace_YM
				PR = PEC_replace_PR
				CTE = PEC_replace_CTE
				EpsX = PEC_replace_EpsX
				EpsY = PEC_replace_EpsX
				EpsZ = PEC_replace_EpsX
				MuX = PEC_replace_MuX
				MuY = PEC_replace_MuX
				MuZ = PEC_replace_MuX
				SigmaX = PEC_replace_SigmaX
				SigmaY = PEC_replace_SigmaX
				SigmaZ = PEC_replace_SigmaX
			Else
				Material.GetRho(sTmp, rho)
				Material.GetThermalConductivity(sTmp, kx, ky, kz)
				Material.GetSpecificHeat(sTmp, Cp)
				Material.GetYoungsModulus(sTmp, YM)
				Material.GetPoissonsRatio(sTmp, PR)
				Material.GetThermalExpansionRate(sTmp, CTE)
				Material.GetEpsilon(sTmp, EpsX, EpsY, EpsZ)
				Material.GetMu(sTmp, MuX, MuY, MuZ)
				Material.GetSigma(sTmp, SigmaX, SigmaY, SigmaZ)
			End If

			Print #StreamNum_Script, "mdb.models['Model-1'].Material(name=" + sAbaqusMatName  + ")"
			'epsilon and mu do not have units

			' vacuum permittivity is decreased by a factor of 1e-6 from SI to mm - tonne - seconds
			If EpsX <> 0.0 Or EpsY <> 0.0 Or EpsZ <> 0.0 Then
				If EpsX = EpsY And EpsX = EpsZ Then
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Dielectric(table=((" + cstr(EpsX*Eps0/dUnitConversionFactor^2) + ", ), ))"
				Else
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Dielectric(table=((" + cstr(EpsX*Eps0/dUnitConversionFactor^2) + "," + _
					cstr(EpsY*Eps0/dUnitConversionFactor^2) + "," + cstr(EpsZ*Eps0/dUnitConversionFactor^2)+"), ), type=ORTHOTROPIC)"
				End If
			End If

			' vacuum permeability is unchanged for SI vs. mm - tonne - seconds
			If MuX <> 0.0 Or MuY <> 0.0 Or MuZ <> 0.0 Then
				If MuX = MuY And MuX = MuZ And MuX <> 0.0 Then
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].MagneticPermeability(table=((" + cstr(MuX * Mu0) + ", ), ))"
				Else
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].MagneticPermeability(table=((" + cstr(MuX * Mu0) + "," + _
					cstr(MuY * Mu0) + "," + cstr(MuZ * Mu0)+"), ), type=ORTHOTROPIC)"
				End If
			End If

			' Sigma is originally given in SI units - either it stays, or is converted to S/mm
			If SigmaX <> 0.0 Or SigmaY <> 0.0 Or SigmaZ <> 0.0 Then
				If SigmaX = SigmaY And SigmaX = SigmaZ And SigmaX <> 0.0 Then
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].ElectricalConductivity(table=((" + cstr(SigmaX/dUnitConversionFactor) + ", ), ))"
				Else
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].ElectricalConductivity(table=((" + cstr(SigmaX/dUnitConversionFactor) + "," + _
					cstr(SigmaY/dUnitConversionFactor) + "," + cstr(SigmaZ/dUnitConversionFactor)+"), ), type=ORTHOTROPIC)"
				End If
			End If

			' Density is originally given in SI units - either it stays kg/m^3, or is converted to tonne/mm^3
			If rho <> 0.0 Then Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Density(table=((" + cstr(rho/dUnitConversionFactor^4) + ", ), ))"

			' Specific heat is originally given in SI units - either it stays in J/kg/K or is converted to mJ/tonne/K
			If  Cp <> 0.0 Then Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].SpecificHeat(table=((" + cstr(Cp*dUnitConversionFactor^2) + ", ), ))"

			' Thermal conductivity is originally given in SI units - the value doesn not change... (mW/mmK = W/mK)
			If kx <> 0.0 Or ky <> 0.0 Or kz <> 0.0 Then
				If kx = ky And kx = kz Then
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Conductivity(table=((" + cstr(kx) + ", ), ))"
				Else
					Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Conductivity(table=((" + cstr(kx) + "," + cstr(ky) + "," + cstr(kz)+"), ), type=ORTHOTROPIC)"
				End If
			End If

			' Youngs Modulas is originally given in GPa - either it is converted to Pa, or to MPa
			If YM <> 0.0 Or PR <> 0.0 Then
				If GeomSettings.nUnits = 1 Then
					' mm-tonne-seconds
					sYMFactor = "e3"
				Else
					' SI Units
					sYMFactor = "e9"
				End If
				Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Elastic(table=((" + cstr(YM) + sYMFactor + "," + cstr(PR) +" ), ))"
			End If

			'CTE in 1/K, CST in 1E-6
			If CTE <> 0.0 Then Print #StreamNum_Script, "mdb.models['Model-1'].materials[" + sAbaqusMatName + "].Expansion(table=((" + cstr(CTE) +"e-6" + ", ), ))"
			Print #StreamNum_Script, "mdb.models['Model-1'].HomogeneousSolidSection(material=" + sAbaqusMatName +", name=" + sAbaqusMatName + ", thickness=None)"
		End If

		Print #StreamNum_Script, "p = mdb.models['Model-1'].parts[" + sAbaqusPartName + "]"
		Print #StreamNum_Script, "ncells = len(p.cells)"
		Print #StreamNum_Script, "if ncells > 0:"
	    Print #StreamNum_Script, "	myregion = regionToolset.Region(cells = p.cells[0:ncells])"
	    Print #StreamNum_Script, "	p.SectionAssignment (region = myregion, sectionName = " + sAbaqusMatName+ ", offset = 0)"

	    Print #StreamNum_Script, "mdb.models['Model-1'].rootAssembly.Instance(dependent=ON, name=" + Left(sAbaqusPartName, Len(sAbaqusPartName)-1) + "-1" + Chr(39) + ", part= p)"
		Print #StreamNum, sCADFileName + "," + Solid.GetMaterialNameForShape(sParts(i))
	Next

	Print #StreamNum, "   "

	For i = 0 To UBound(sMat)
		If LCase(Material.GetTypeOfMaterial(sMat(i))) = LCase("PEC") Then
			rho = PEC_replace_rho
			kx = PEC_replace_kx
			ky = PEC_replace_kx
			kz = PEC_replace_kx
			Cp = PEC_replace_Cp
			YM = PEC_replace_YM
			PR = PEC_replace_PR
			CTE = PEC_replace_CTE
			EpsX = PEC_replace_EpsX
			EpsY = PEC_replace_EpsX
			EpsZ = PEC_replace_EpsX
			MuX = PEC_replace_MuX
			MuY = PEC_replace_MuX
			MuZ = PEC_replace_MuX
			SigmaX = PEC_replace_SigmaX
			SigmaY = PEC_replace_SigmaX
			SigmaZ = PEC_replace_SigmaX
		Else
			Material.GetRho(sMat(i), rho)
			Material.GetThermalConductivity( sMat(i), kx, ky, kz )
			Material.GetSpecificHeat( sMat(i), Cp )
			Material.GetYoungsModulus(sMat(i),YM)
			Material.GetPoissonsRatio(sMat(i),PR)
			Material.GetThermalExpansionRate(sMat(i), CTE)
			Material.GetEpsilon(sMat(i), EpsX, EpsY, EpsZ)
			Material.GetMu(sMat(i), MuX, MuY, MuZ)
			Material.GetSigma(sMat(i), SigmaX, SigmaY, SigmaZ)
		End If

		Print #StreamNum, sMat(i)
		Print #StreamNum, "Density," + cstr(rho) + ",kg/m^3"
		If kx = ky And kx = kz Then
			Print #StreamNum, "ThermalConductivity," + cstr(kx) + ",W/K/m"
		Else
			Print #StreamNum, "ThermalConductivity," + cstr(kx) + "," + cstr(ky) + "," + cstr(kz) + ",W/K/m"
		End If
		Print #StreamNum, "SpecificHeat," + cstr(Cp) + ",J/K/kg"
		Print #StreamNum, "YoungsModulus," + cstr(YM) + ",GPa"
		Print #StreamNum, "PoissonsRatio," + cstr(PR) + ","
		Print #StreamNum, "ThermalExpansionCoefficient," + cstr(CTE) + ",1e-6/K"
		If EpsX = EpsY And EpsX = EpsZ Then
			Print #StreamNum, "RelativePermittivity," + cstr(EpsX) + ","
		Else
			Print #StreamNum, "RelativePermittivity," + cstr(EpsX) + "," + cstr(EpsY) + "," + cstr(EpsZ) + ","
		End If
		If MuX = MuY And MuX = MuZ Then
			Print #StreamNum, "RelativePermeability ," + cstr(MuX) + ","
		Else
			Print #StreamNum, "RelativePermeability," + cstr(MuX) + "," + cstr(MuY) + "," + cstr(MuZ) + ","
		End If
		If SigmaX = SigmaY And SigmaX = SigmaZ Then
			Print #StreamNum, "ElectricConductivity ," + cstr(SigmaX) + ",S/m"
		Else
			Print #StreamNum, "ElectricConductivity," + cstr(SigmaX) + "," + cstr(SigmaY) + "," + cstr(SigmaZ) + ",S/m"
		End If
	Next

	Print #StreamNum, "   "
End Function



Function PCBType() As Integer
	Dim nSolids As Long, sSolid As String, i As Long, counter As Long
	Dim bNets As Boolean, bSubstrate As Boolean ', bSimplified As Boolean

	PCBType = 0

	counter = 0
	nSolids = Solid.GetNumberOfShapes
	ReDim sPartsNames(nSolids)

	' loop over all solids for names and types

	For i = 0 To nSolids - 1
	    sSolid = Solid.GetNameOfShapeFromIndex ( i )
		sPartsNames (i) = sSolid

	    bNets = (InStr(sSolid, "Nets") > 0)
	    bSubstrate = (InStr(sSolid, "Substrate") > 0)
	'	bSimplified = (InStr(sSolid, "Simplified Geometry") > 0)

	    If bNets Or bSubstrate Then  counter = counter + 1

	'	If bSimplified Then
	'			counter = counter -1
	'	End If
	Next i

	If counter >0 Then PCBType = 1
	'If counter <0 Then PCBType = -1
End Function



Function GetAllPCBs() As String()
	Dim bNew As Boolean
	Dim sParent As String, sItem As String, sPCB() As String
	Dim nPCB As Integer, i As Long

	nPCB = 0
	'ReDim sPCB(0)

	For i = 0 To UBound(sPartsNames)
		If (InStr(sPartsNames(i), "Nets") >0) Then
			sParent = Left(sPartsNames(i),InStr(sPartsNames(i), "Nets")-2)
			bNew = True
			For Each sItem In sPCB
				If sParent = sItem Then
					bNew = False
					Exit For
				End If
			Next
			If bNew Then
				ReDim Preserve sPCB(nPCB)
				sPCB(nPCB) = sParent
				nPCB = nPCB + 1
			End If
		End If

	Next

	'For i = 0 To UBound(sPCB)
	'	sPCB(i) = Mid(sPCB(i), InStrRev(sPCB(i),"/") + 1)
	'Next

	GetAllPCBs = sPCB
End Function



Function NamesUnder(sRoot As String) As Variant
	Dim i As Long, sSolid As String, nSolids As Long, sNames As Variant, nCount As Long
	Dim indexL As Integer, indexR As Integer, LS As String, RS As String

	nSolids = Solid.GetNumberOfShapes

	If sRoot = "Components" Then
		ReDim sNames(nSolids-1)
		For i = 0 To nSolids - 1
		    sSolid = Solid.GetNameOfShapeFromIndex ( i )
			sNames(i) = sSolid
		Next
	Else
		nCount = 0
		For i = 0 To nSolids - 1
		    sSolid = Solid.GetNameOfShapeFromIndex ( i )
		    indexL = InStr(sSolid, sRoot)
		    If  indexL > 0 Then
				indexR = indexL + Len(sRoot)
				If indexL = 1 Then
					LS =""
				Else
					LS = Mid(sSolid,indexL-1,1)
				End If

				RS = Mid(sSolid,indexR,1)

				If (LS = "/" Or LS = "" ) And ( RS = "/" Or RS = ":" ) Then
					ReDim Preserve sNames(nCount)
					sNames (nCount) = sSolid
					nCount = nCount + 1
				End If
		    End If
		Next i
	End If

	NamesUnder = sNames
End Function



Function GetAllNetNames(sCurrPCB As String) As Variant
	Dim bNew As Boolean
	Dim sChild As String, sItem As String, sNets() As String, sAllParts As Variant

	Dim nNets As Integer, i As Long

	nNets = 0
	sAllParts = NamesUnder(sCurrPCB + "/Nets")

	For i = 0 To UBound(sAllParts)
			sChild = Left(sAllParts(i),InStr(sAllParts(i), ":")-1)
			bNew = True
			For Each sItem In sNets
				If sChild = sItem Then
					bNew = False
					Exit For
				End If
			Next
			If bNew Then
				ReDim Preserve sNets(nNets)
				sNets(nNets) = sChild
				nNets = nNets + 1
			End If
	Next

	GetAllNetNames = sNets
End Function



Function GetAllLayerNames(sCurrPCB As String) As Variant
	Dim bNew As Boolean
	Dim sLayer As String, sItem As String, sLayers() As String, sAllParts As Variant

	Dim nLayers As Integer, i As Long

	nLayers = 0
	sAllParts = NamesUnder(sCurrPCB + "/Nets")

	For i = 0 To UBound(sAllParts)
		sLayer = Split(sAllParts(i),":")(1)
		bNew = True
		For Each sItem In sLayers
			If sLayer = sItem Then
				bNew = False
				Exit For
			End If
		Next
		If bNew Then
			ReDim Preserve sLayers(nLayers)
			sLayers(nLayers) = sLayer
			nLayers = nLayers + 1
		End If
	Next

	GetAllLayerNames = sLayers
End Function



Function GetTopLevelComponentNames() As Variant
	Dim sComp() As String, sNext As String, sNext_strip As String
	Dim iCount As Long
	ReDim sComp(0) As String

	sNext = ResultTree.GetFirstChildName("Components")
	If sNext <> "" Then
		sNext_strip = Replace(sNext,"Components\","",1,1)
	End If
	sComp(0) = sNext_strip
	iCount = 0

	sNext = ResultTree.GetNextItemName (sNext)
	While sNext <> ""
		iCount = iCount + 1
		ReDim Preserve sComp(iCount) As String
		sNext_strip = Replace(sNext,"Components\","",1,1)
		sComp(iCount) = sNext_strip
		sNext = ResultTree.GetNextItemName (sNext)
	Wend

GetTopLevelComponentNames = sComp
End Function



' Function: AbaqusName
' This function makes the sName compatible with Abaqus, which requires:
'  - Names must be 1-38 characters long,
'  - may not begin with a number
'  - may not begin or end with a space or an underscore
'  - may not contain double quotes, periods, backward slashs, or non-printable characters
Function AbaqusName(sName As String) As String
	Dim i As Integer, iASCIICode As Integer
	Dim sBegin As String, sEnd As String

	'taking care of non-printable characters and double quote ( = chr(34))
	For i = 1 To Len(sName)
		iASCIICode = Asc(Mid(sName,i,1))
		If iASCIICode <40 Or iASCIICode > 122 Then
			sName = Replace(sName, Chr(iASCIICode),"_")
		End If
	Next

	sName = Replace (sName, ".",    "_") 'replace period chr(46) with "_"
	sName = Replace (sName, "\",    "_") 'replace "\" chr(92) with "_"
	sName = Replace (sName, "/",    "_") 'replace "/"  with "_" added 2023-03-07
	sName = Replace (sName, ":",    "_") 'replace ":"  with "_"

	sBegin = Left(sName, 1)
	sEnd = Right(sName,1)

	If IsNumeric(sBegin) Or sBegin = " " Or sBegin = "_" Then	sName = "B" + Right(sName, Len(sName)-1)

	If sEnd = " " Or sEnd = "_" Then	sName = Left(sName, Len(sName)-1) + "E"

	AbaqusName  = Left(sName, 38)
End Function
