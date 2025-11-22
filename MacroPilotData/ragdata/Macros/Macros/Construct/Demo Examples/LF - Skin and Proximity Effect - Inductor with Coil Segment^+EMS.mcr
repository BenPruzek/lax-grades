'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Macro: Creates demo example for the Low Frequency Frequency Domain solver
'
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 31-Oct-2023 mha: Changed to more standard history list entries, set coil segment from solid and switched to peak values.
' ================================================================================================

' *** global variables
Dim sHistoryCaption As String
Dim sHistoryContent As String


Sub Main
	' define parameters...
	StoreParameterWithDescription ("current", "1/sqr(2)", "coil current")
	StoreParameterWithDescription ("freq", "1e4", "frequency")
	StoreParameterWithDescription ("n", "4", "number of turns [1], has to be >= 2")
	StoreParameterWithDescription ("r_wire", "1", "radius of inductor's wire")
	StoreParameterWithDescription ("height", "2.1*n*r_wire", "height of inductor")
	StoreParameterWithDescription ("r_blend", "2*r_wire", "radius inductor's blend")
	StoreParameterWithDescription ("r_i", "18", "inner radius of inductor")

	' define geometry and settings...
	sHistoryCaption = "define units"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Units" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Length"", ""mm""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Temperature"", ""degC""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Voltage"", ""V""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Current"", ""A""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Resistance"", ""Ohm""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Conductance"", ""S""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Capacitance"", ""F""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Inductance"", ""H""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Frequency"", ""Hz""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetUnit ""Time"", ""s""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetResultUnit ""frequency"", ""frequency"", """"" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "change solver type"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "ChangeSolverType ""LF Frequency Domain""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define boundaries"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Boundary" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Xmin ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Xmax ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Ymin ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Ymax ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Zmin ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Zmax ""electric""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Xsymmetry ""none""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Ysymmetry ""none""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Zsymmetry ""none""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ApplyInAllDirections ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XminPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XminPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XmaxPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XmaxPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YminPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YminPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YmaxPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YmaxPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZminPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZminPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZmaxPotential """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZmaxPotentialType ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define background"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Background" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ResetBackground" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XminSpace ""2*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .XmaxSpace ""2*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YminSpace ""2*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .YmaxSpace ""2*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZminSpace ""3*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ZmaxSpace ""3*r_i""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ApplyInAllDirections ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	sHistoryContent = sHistoryContent & "With Material" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Rho ""1.204""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ThermalType ""Normal""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ThermalConductivity ""0.026""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SpecificHeat ""1005"", ""J/K/kg""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DynamicViscosity ""1.84e-5""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseEmissivity ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Emissivity ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MetabolicRate ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .VoxelConvection ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .BloodFlow ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Absorptance ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MechanicsType ""Unused""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .IntrinsicCarrierDensity ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .FrqType ""all""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Type ""Normal""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaterialUnit ""Frequency"", ""Hz""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaterialUnit ""Geometry"", ""m""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaterialUnit ""Time"", ""s""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaterialUnit ""Temperature"", ""K""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Epsilon ""1.00059""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Mu ""1.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Sigma ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanD ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDFreq ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDGiven ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDModel ""ConstTanD""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetConstTanDStrategyEps ""AutomaticOrder""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ConstTanDModelOrderEps ""3""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DjordjevicSarkarUpperFreqEps ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetElParametricConductivity ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ReferenceCoordSystem ""Global""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .CoordSystemType ""Cartesian""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SigmaM ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDM ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDMFreq ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDMGiven ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TanDMModel ""ConstTanD""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetConstTanDStrategyMu ""AutomaticOrder""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ConstTanDModelOrderMu ""3""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DjordjevicSarkarUpperFreqMu ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMagParametricConductivity ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DispModelEps  ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DispModelMu ""None""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DispersiveFittingSchemeEps ""Nth Order""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaximalOrderNthModelFitEps ""10""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ErrorLimitNthModelFitEps ""0.1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseOnlyDataInSimFreqRangeNthModelEps ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DispersiveFittingSchemeMu ""Nth Order""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaximalOrderNthModelFitMu ""10""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ErrorLimitNthModelFitMu ""0.1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseOnlyDataInSimFreqRangeNthModelMu ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseGeneralDispersionEps ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseGeneralDispersionMu ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NLAnisotropy ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NLAStackingFactor ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NLADirectionX ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NLADirectionY ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NLADirectionZ ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Colour ""0.682353"", ""0.717647"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Wireframe ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reflection ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Allowoutline ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Transparentoutline ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Transparency ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ChangeBackgroundMaterial" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "set mesh properties (Tetrahedral)"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Mesh" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MeshType ""Tetrahedral""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetCreator ""Low Frequency""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Tet""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""Version"", 1%" & vbCrLf
	sHistoryContent = sHistoryContent & "    'MAX CELL - WAVELENGTH REFINEMENT" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""StepsPerWaveNear"", ""4""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""StepsPerWaveFar"", ""4""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""PhaseErrorNear"", ""0.02""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""PhaseErrorFar"", ""0.02""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""CellsPerWavelengthPolicy"", ""automatic""" & vbCrLf
	sHistoryContent = sHistoryContent & "    'MAX CELL - GEOMETRY REFINEMENT" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""StepsPerBoxNear"", ""10""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""StepsPerBoxFar"", ""3""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""ModelBoxDescrNear"", ""maxedge""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""ModelBoxDescrFar"", ""maxedge""" & vbCrLf
	sHistoryContent = sHistoryContent & "    'MIN CELL" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""UseRatioLimit"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""RatioLimit"", ""100""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""MinStep"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    'MESHING METHOD" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Unstr""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""Method"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Tet""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""CurvatureOrder"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""CurvatureOrderPolicy"", ""automatic""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""CurvRefinementControl"", ""NormalTolerance""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""NormalTolerance"", ""22.5""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""SrfMeshGradation"", ""1.5""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""SrfMeshOptimization"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Unstr""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""UseMaterials"",  ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""MoveMesh"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""All""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""AutomaticEdgeRefinement"",  ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Tet""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""UseAnisoCurveRefinement"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""UseSameSrfAndVolMeshGradation"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""VolMeshGradation"", ""1.5""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""VolMeshOptimization"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With MeshSettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMeshType ""Unstr""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""SmallFeatureSize"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""CoincidenceTolerance"", ""1e-06""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""SelfIntersectionCheck"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Set ""OptimizeForPlanarStructures"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "With Mesh" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetParallelMesherMode ""Tet"", ""maximum""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetMaxParallelMesherThreads ""Tet"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define lf solver frequency settings"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With LFSolver" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ResetFrequencySettings" & vbCrLf
	sHistoryContent = sHistoryContent & "    .AddFrequency ""freq""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define lf frequency domain solver special parameters"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With LFSolver" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaxLinIter ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Preconditioner ""ILU""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .SetTreeCotreeGauging ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .EnableDivergenceCheck ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LSESolverType ""Auto""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .BroadbandCalculation ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NonlinearEquivalentMu ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NonlinearEquivalentMuMaxIter ""10""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NonlinearEquivalentMuAccu ""1e-3""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TetSolverOrder ""2""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define lf frequency domain solver parameters"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With LFSolver" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Method ""Tetrahedral Mesh""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Accuracy ""1e-6""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .CalcImpedanceMatrix ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .StoreResultsInCache ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MeshAdaption ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .EquationType ""Magnetoquasistatic""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ValueScaling ""peak""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TetAdaption ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseMaxNumberOfThreads ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaxNumberOfThreads ""48""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MaximumNumberOfCPUDevices ""2""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UseDistributedComputing ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf
	sHistoryContent = sHistoryContent & "UseDistributedComputingForParameters ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "MaxNumberOfDistributedComputingParameters ""2""" & vbCrLf
	sHistoryContent = sHistoryContent & "UseDistributedComputingMemorySetting ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "MinDistributedComputingMemoryLimit ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "UseDistributedComputingSharedDirectory ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "OnlyConsider0D1DResultsForDC ""False""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define curve analytical: curve1:inner_coil"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With AnalyticalCurve" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""inner_coil""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Curve ""curve1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LawX ""r_i*cos(2*pi*n*t)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LawY ""r_i*sin(2*pi*n*t)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LawZ ""height*t""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ParameterRange ""0"", ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick end point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:inner_coil"", ""1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "align wcs with point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.AlignWCSWithSelected ""Point""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "rotate wcs around u +90 degrees"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.RotateWCS ""u"", ""90""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "store picked point: 1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.NextPickToDatabase ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:inner_coil"", ""1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "store picked point: 2"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.NextPickToDatabase ""2""" & vbCrLf
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:inner_coil"", ""2""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define curve polygon: curve1:termination"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Polygon" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""termination""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Curve ""curve1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Point ""xp(1)"", ""yp(1)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .RLine ""5*r_wire"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .RLine ""0"", ""yp(2)-yp(1)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LineTo ""xp(2)"", ""yp(2)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "activate global coordinates"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.ActivateWCS ""global""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick end point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:inner_coil"", ""1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick end point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:termination"", ""2""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick end point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:termination"", ""3""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick end point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveEndpointFromId ""curve1:inner_coil"", ""2""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define curve blend: :blend1 on: picked points"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With BlendCurve" & vbCrLf
	sHistoryContent = sHistoryContent & "  .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "  .Name ""blend1""" & vbCrLf
	sHistoryContent = sHistoryContent & "  .Radius ""r_blend""" & vbCrLf
	sHistoryContent = sHistoryContent & "  .UsePickedPoints" & vbCrLf
	sHistoryContent = sHistoryContent & "  .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "pick mid point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Pick.PickCurveMidpointFromId ""curve1:termination"", ""2""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "align wcs with point"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.AlignWCSWithSelected ""Point""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "rotate wcs around u +90 degrees"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.RotateWCS ""u"", ""90""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define curve circle: curve1:circle1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Circle" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""circle1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Curve ""curve1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Radius ""r_wire""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Xcenter ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Ycenter ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Segments ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "new component: component1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Component.New ""component1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define sweepprofile: component1:conductor"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With SweepCurve" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""conductor""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Component ""component1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Material ""Copper (annealed)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Twistangle ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Taperangle ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ProjectProfileToPathAdvanced ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .CutEndOff ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DeleteProfile ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .DeletePath ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Path ""curve1:inner_coil""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Curve ""curve1:circle1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "delete curve: curve1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Curve.DeleteCurve ""curve1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define cylinder: component1:Coil"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Cylinder" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""Coil""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Component ""component1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Material ""Copper (annealed)""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .OuterRadius ""1.01*r_wire""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .InnerRadius ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Axis ""z""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Zrange ""-0.1"", ""0.1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Xcenter ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Ycenter ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Segments ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "transform: translate component1:conductor"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Transform" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""component1:conductor""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Vector ""0"", ""0"", ""0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .UsePickedPoints ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .InvertPickedPoints ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MultipleObjects ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .GroupObjects ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Repetitions ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .MultipleSelection ""False""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Destination """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Material """"" & vbCrLf
	sHistoryContent = sHistoryContent & "    .AutoDestination ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Transform ""Shape"", ""Translate""" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "boolean insert shapes: component1:conductor, component1:Coil"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Solid.Insert ""component1:conductor"", ""component1:Coil""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "boolean intersect shapes: component1:Coil, component1:conductor_1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "Solid.Intersect ""component1:Coil"", ""component1:conductor_1""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "activate global coordinates"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "WCS.ActivateWCS ""global""" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	sHistoryCaption = "define coil: coil1"
	sHistoryContent = ""
	sHistoryContent = sHistoryContent & "With Coil" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Reset" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Name ""coil1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Type ""Coil Segment""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .OperationMode ""Current""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ConductorModel ""Solid""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ToolType ""CoilSegmentFromSolid""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Value ""current""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Phase ""0.0""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NTurns ""1""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Conductivity ""5.8e7""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .TerminationType ""PEC""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .NStrands ""120""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .StrandDiameter ""0.05""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .LengthExtension ""1.2""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .CurrentDirection ""Regular""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .ProjectProfileToPathAdvanced ""True""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .AddSolid ""component1:Coil""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Port0 ""component1:Coil"", ""7""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Port1 ""component1:Coil"", ""6""" & vbCrLf
	sHistoryContent = sHistoryContent & "    .Create" & vbCrLf
	sHistoryContent = sHistoryContent & "End With" & vbCrLf & vbCrLf
	AddToHistory(sHistoryCaption, sHistoryContent)

	' not in history list...
	ResetViewToStructure()
	Plot.DrawBox(True)
End Sub
