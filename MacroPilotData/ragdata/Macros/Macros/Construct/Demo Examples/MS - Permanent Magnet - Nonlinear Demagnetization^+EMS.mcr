Sub Main ()

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 31-Jan-2018 ube: First version
' ================================================================================================

Dim sCommand As String

'@ switch Bounding Box on
Plot.DrawBox True

'@ switch working plane off
Plot.DrawWorkplane "false"

BeginHide

	'@ add parameters
	StoreDoubleParameter "inner_radius", 0.5
	StoreDoubleParameter "outer_radius", 2
	StoreDoubleParameter "magnet_thickness", 1
	StoreDoubleParameter "plate_thickness", 0.2

	'@ change Solver Type
	sHistoryString = ""
	sHistoryString = sHistoryString + "" + vbLf
	sHistoryString = sHistoryString + "ChangeSolverType ""LF MStatic"" "
	AddToHistory "change solver type", sHistoryString


    '@ define units
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Units" + vbLf
    sHistoryString = sHistoryString + "    .Geometry ""mm""" + vbLf
    sHistoryString = sHistoryString + "    .Frequency ""Hz""" + vbLf
    sHistoryString = sHistoryString + "    .Voltage ""V""" + vbLf
    sHistoryString = sHistoryString + "    .Resistance ""Ohm""" + vbLf
    sHistoryString = sHistoryString + "    .Inductance ""NanoH""" + vbLf
    sHistoryString = sHistoryString + "    .TemperatureUnit  ""Kelvin""" + vbLf
    sHistoryString = sHistoryString + "    .Time ""s""" + vbLf
    sHistoryString = sHistoryString + "    .Current ""A""" + vbLf
    sHistoryString = sHistoryString + "    .Conductance ""Siemens""" + vbLf
    sHistoryString = sHistoryString + "    .Capacitance ""PikoF""" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define units", sHistoryString

    
    '@ define background
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Background " + vbLf
    sHistoryString = sHistoryString + "     .ResetBackground " + vbLf
    sHistoryString = sHistoryString + "     .XminSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .XmaxSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .YminSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .YmaxSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .ZminSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .ZmaxSpace ""3"" " + vbLf
    sHistoryString = sHistoryString + "     .ApplyInAllDirections ""True"" " + vbLf
    sHistoryString = sHistoryString + "End With " + vbLf
    sHistoryString = sHistoryString + "With Material " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Rho ""1.204""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalType ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalConductivity ""0.026""" + vbLf
    sHistoryString = sHistoryString + "     .SpecificHeat ""1005"", ""J/K/kg""" + vbLf
    sHistoryString = sHistoryString + "     .DynamicViscosity ""1.84e-5""" + vbLf
    sHistoryString = sHistoryString + "     .Emissivity ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .MetabolicRate ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .VoxelConvection ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .BloodFlow ""0""" + vbLf
    sHistoryString = sHistoryString + "     .MechanicsType ""Unused""" + vbLf
    sHistoryString = sHistoryString + "     .FrqType ""all""" + vbLf
    sHistoryString = sHistoryString + "     .Type ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Frequency"", ""Hz""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Geometry"", ""m""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Time"", ""s""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Temperature"", ""Kelvin""" + vbLf
    sHistoryString = sHistoryString + "     .Epsilon ""1.0""" + vbLf
    sHistoryString = sHistoryString + "     .Mu ""1.0""" + vbLf
    sHistoryString = sHistoryString + "     .Sigma ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanD ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDModel ""ConstSigma""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderEps ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetElParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ReferenceCoordSystem ""Global""" + vbLf
    sHistoryString = sHistoryString + "     .CoordSystemType ""Cartesian""" + vbLf
    sHistoryString = sHistoryString + "     .SigmaM ""0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDM ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMModel ""ConstSigma""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderMu ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetMagParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelEps  ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelMu ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeEps ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitEps ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitEps ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeMu ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitMu ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitMu ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NonlinearMeasurementError ""1e-1""" + vbLf
    sHistoryString = sHistoryString + "     .NLAnisotropy ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NLAStackingFactor ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionX ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionY ""0""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionZ ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Colour ""0.6"", ""0.6"", ""0.6"" " + vbLf
    sHistoryString = sHistoryString + "     .Wireframe ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Reflection ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Allowoutline ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparentoutline ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparency ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .ChangeBackgroundMaterial" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define background", sHistoryString

    
    '@ define boundaries
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Boundary" + vbLf
    sHistoryString = sHistoryString + "     .Xmin ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Xmax ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Ymin ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Ymax ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Zmin ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Zmax ""electric""" + vbLf
    sHistoryString = sHistoryString + "     .Xsymmetry ""none""" + vbLf
    sHistoryString = sHistoryString + "     .Ysymmetry ""none""" + vbLf
    sHistoryString = sHistoryString + "     .Zsymmetry ""none""" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define boundaries", sHistoryString

    
    '@ define material: HF 14.22 p
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Material " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Name ""HF 14.22 p""" + vbLf
    sHistoryString = sHistoryString + "     .Folder """"" + vbLf
    sHistoryString = sHistoryString + "     .FrqType ""all""" + vbLf
    sHistoryString = sHistoryString + "     .Type ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Frequency"", ""Hz""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Geometry"", ""mm""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Time"", ""s""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Temperature"", ""Kelvin""" + vbLf
    sHistoryString = sHistoryString + "     .Epsilon ""1""" + vbLf
    sHistoryString = sHistoryString + "     .Mu ""1""" + vbLf
    sHistoryString = sHistoryString + "     .Sigma ""0""" + vbLf
    sHistoryString = sHistoryString + "     .TanD ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDModel ""ConstTanD""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderEps ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetElParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ReferenceCoordSystem ""Global""" + vbLf
    sHistoryString = sHistoryString + "     .CoordSystemType ""Cartesian""" + vbLf
    sHistoryString = sHistoryString + "     .SigmaM ""0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDM ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMModel ""ConstTanD""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderMu ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetMagParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelEps  ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelMu ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeEps ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitEps ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitEps ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeMu ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitMu ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitMu ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NonlinearMeasurementError ""1e-3""" + vbLf
    sHistoryString = sHistoryString + "     .ResetHBList" + vbLf
    sHistoryString = sHistoryString + "     .SetNonlinearCurveType ""Hard-Magnetic-JH""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-245000"", ""0""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-230000"", "".1""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-220000"", "".155""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-210000"", "".195""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-200000"", "".220""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-180000"", "".245""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-160000"", "".260""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-100000"", "".270""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""0"", "".275""" + vbLf
    sHistoryString = sHistoryString + "     .GenerateNonlinearCurve" + vbLf
    sHistoryString = sHistoryString + "     .NLAnisotropy ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NLAStackingFactor ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionX ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionY ""0""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionZ ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Rho ""0""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalType ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalConductivity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .SpecificHeat ""0"", ""J/K/kg""" + vbLf
    sHistoryString = sHistoryString + "     .DynamicViscosity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Emissivity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .MetabolicRate ""0""" + vbLf
    sHistoryString = sHistoryString + "     .BloodFlow ""0""" + vbLf
    sHistoryString = sHistoryString + "     .VoxelConvection ""0""" + vbLf
    sHistoryString = sHistoryString + "     .MechanicsType ""Unused""" + vbLf
    sHistoryString = sHistoryString + "     .Colour ""1"", ""0.501961"", ""0.25098"" " + vbLf
    sHistoryString = sHistoryString + "     .Wireframe ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Reflection ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Allowoutline ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparentoutline ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparency ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Create" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define material: HF 14.22 p", sHistoryString    

    
    '@ define material: AlNiCo
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Material " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Name ""AlNiCo""" + vbLf
    sHistoryString = sHistoryString + "     .Folder """"" + vbLf
    sHistoryString = sHistoryString + "     .FrqType ""all""" + vbLf
    sHistoryString = sHistoryString + "     .Type ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Frequency"", ""Hz""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Geometry"", ""mm""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Time"", ""s""" + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit ""Temperature"", ""Kelvin""" + vbLf
    sHistoryString = sHistoryString + "     .Epsilon ""1""" + vbLf
    sHistoryString = sHistoryString + "     .Mu ""1""" + vbLf
    sHistoryString = sHistoryString + "     .Sigma ""0""" + vbLf
    sHistoryString = sHistoryString + "     .TanD ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDModel ""ConstTanD""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderEps ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetElParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ReferenceCoordSystem ""Global""" + vbLf
    sHistoryString = sHistoryString + "     .CoordSystemType ""Cartesian""" + vbLf
    sHistoryString = sHistoryString + "     .SigmaM ""0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDM ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMFreq ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMGiven ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TanDMModel ""ConstTanD""" + vbLf
    sHistoryString = sHistoryString + "     .EnableUserConstTanDModelOrderMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .ConstTanDModelOrderMu ""1""" + vbLf
    sHistoryString = sHistoryString + "     .SetMagParametricConductivity ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelEps  ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispModelMu ""None""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeEps ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitEps ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitEps ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeMu ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitMu ""10""" + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitMu ""0.1""" + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionEps ""False""" + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionMu ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NonlinearMeasurementError ""1e-3""" + vbLf
    sHistoryString = sHistoryString + "     .ResetHBList" + vbLf
    sHistoryString = sHistoryString + "     .SetNonlinearCurveType ""Hard-Magnetic-JH""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-47746.5"", "" 0""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-45757"", "" 0.45""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-43767.6"", "" 0.6""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-39788.7"", "" 0.78""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-31831"", "" 0.93""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-23873.2"", "" 0.99""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-15915.5"", "" 1.03""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue ""-7957.75"", "" 1.075""" + vbLf
    sHistoryString = sHistoryString + "     .AddNonlinearCurveValue "" 0"", "" 1.09""" + vbLf
    sHistoryString = sHistoryString + "     .GenerateNonlinearCurve" + vbLf
    sHistoryString = sHistoryString + "     .NLAnisotropy ""False""" + vbLf
    sHistoryString = sHistoryString + "     .NLAStackingFactor ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionX ""1""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionY ""0""" + vbLf
    sHistoryString = sHistoryString + "     .NLADirectionZ ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Rho ""0""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalType ""Normal""" + vbLf
    sHistoryString = sHistoryString + "     .ThermalConductivity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .SpecificHeat ""0"", ""J/K/kg""" + vbLf
    sHistoryString = sHistoryString + "     .DynamicViscosity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Emissivity ""0""" + vbLf
    sHistoryString = sHistoryString + "     .MetabolicRate ""0""" + vbLf
    sHistoryString = sHistoryString + "     .BloodFlow ""0""" + vbLf
    sHistoryString = sHistoryString + "     .VoxelConvection ""0""" + vbLf
    sHistoryString = sHistoryString + "     .MechanicsType ""Unused""" + vbLf
    sHistoryString = sHistoryString + "     .Colour ""0"", ""1"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Wireframe ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Reflection ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Allowoutline ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparentoutline ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparency ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Create" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define material: AlNiCo", sHistoryString    

    
    '@ define material: ST37
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Material" + vbLf
    sHistoryString = sHistoryString + "     .Reset" + vbLf
    sHistoryString = sHistoryString + "     .Name ""ST37""" + vbLf
    sHistoryString = sHistoryString + "     .Folder """"" + vbLf
    sHistoryString = sHistoryString + "     .FrqType ""all"" " + vbLf
    sHistoryString = sHistoryString + "     .Type ""Nonlinear"" " + vbLf
    sHistoryString = sHistoryString + "     .SetMaterialUnit ""Hz"", ""m"" " + vbLf
    sHistoryString = sHistoryString + "     .Mue ""1000"" " + vbLf
    sHistoryString = sHistoryString + "     .NonlinearMeasurementError ""1e-1"" " + vbLf
    sHistoryString = sHistoryString + "     .Kappa ""0""" + vbLf
    sHistoryString = sHistoryString + "     .ResetHBList" + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""0"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""50"", ""0.02"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""100"", ""0.045"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""150"", ""0.09"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""200"", ""0.15"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""300"", ""0.265"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""400"", ""0.38"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""500"", ""0.5"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""600"", ""0.62"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""700"", ""0.74"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""800"", ""0.85"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""900"", ""0.95"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1000"", ""1.033"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1100"", ""1.11"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1200"", ""1.17"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1400"", ""1.26"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1600"", ""1.32"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1800"", ""1.365"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2000"", ""1.4"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2400"", ""1.46"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2800"", ""1.51"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3000"", ""1.53"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3500"", ""1.575"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""4000"", ""1.62"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""5000"", ""1.69"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""6000"", ""1.75"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""7000"", ""1.79"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""8000"", ""1.82"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1.0e+004"", ""1.87"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1.2e+004"", ""1.9"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1.6e+004"", ""1.95"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""1.8e+004"", ""1.975"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2.0e+004"", ""2"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2.2e+004"", ""2.02"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2.4e+004"", ""2.04"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2.6e+004"", ""2.058"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""2.8e+004"", ""2.075"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3.0e+004"", ""2.09"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3.2e+004"", ""2.105"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3.4e+004"", ""2.12"" " + vbLf
    sHistoryString = sHistoryString + "     .AddHBValue ""3.6e+004"", ""2.136"" " + vbLf
    sHistoryString = sHistoryString + "     .Rho ""7850"" " + vbLf
    sHistoryString = sHistoryString + "     .ThermalType ""Normal"" " + vbLf
    sHistoryString = sHistoryString + "     .ThermalConductivity ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .SpecificHeat ""0"", ""J/K/kg""" + vbLf
    sHistoryString = sHistoryString + "     .MetabolicRate ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .BloodFlow ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .VoxelConvection ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .MechanicsType ""Unused"" " + vbLf
    sHistoryString = sHistoryString + "     .FrqType ""hf"" " + vbLf
    sHistoryString = sHistoryString + "     .Type ""Normal"" " + vbLf
    sHistoryString = sHistoryString + "     .SetMaterialUnit ""Hz"", ""m"" " + vbLf
    sHistoryString = sHistoryString + "     .Epsilon ""1"" " + vbLf
    sHistoryString = sHistoryString + "     .Mue ""1"" " + vbLf
    sHistoryString = sHistoryString + "     .Kappa ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanD ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDFreq ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDGiven ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDModel ""ConstTanD"" " + vbLf
    sHistoryString = sHistoryString + "     .KappaM ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDM ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDMFreq ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDMGiven ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .TanDMModel ""ConstTanD"" " + vbLf
    sHistoryString = sHistoryString + "     .DispModelEps ""None"" " + vbLf
    sHistoryString = sHistoryString + "     .DispModelMue ""None"" " + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeEps ""Nth Order"" " + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeMue ""Nth Order"" " + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionEps ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .UseGeneralDispersionMue ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Colour ""0.443137"", ""0.509804"", ""0.184314"" " + vbLf
    sHistoryString = sHistoryString + "     .Wireframe ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Reflection ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Allowoutline ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparentoutline ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Transparency ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Create" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define material: ST37", sHistoryString


    '@ new component: component1
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "Component.New ""component1"""
    AddToHistory "new component: component1", sHistoryString


    '@ define cylinder: component1:solid1
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Cylinder " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Name ""magnet"" " + vbLf
    sHistoryString = sHistoryString + "     .Component ""component1"" " + vbLf
    sHistoryString = sHistoryString + "     .Material ""HF 14.22 p"" " + vbLf
    sHistoryString = sHistoryString + "     .OuterRadius ""outer_radius"" " + vbLf
    sHistoryString = sHistoryString + "     .InnerRadius ""inner_radius"" " + vbLf
    sHistoryString = sHistoryString + "     .Axis ""z"" " + vbLf
    sHistoryString = sHistoryString + "     .Zrange ""0"", ""magnet_thickness"" " + vbLf
    sHistoryString = sHistoryString + "     .Xcenter ""-0"" " + vbLf
    sHistoryString = sHistoryString + "     .Ycenter ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Segments ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Create " + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define cylinder: component1:solid1", sHistoryString
    

    '@ define magnet: magnet1
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Magnet" + vbLf
    sHistoryString = sHistoryString + "     .Reset" + vbLf
    sHistoryString = sHistoryString + "     .Name ""magnet1""" + vbLf
    sHistoryString = sHistoryString + "     .SetMagnetType ""Constant""" + vbLf
    sHistoryString = sHistoryString + "     .Remanence ""1""" + vbLf
    sHistoryString = sHistoryString + "     .MagDir ""0"", ""0"", ""1""" + vbLf
    sHistoryString = sHistoryString + "     .InverseDir ""False""" + vbLf
    sHistoryString = sHistoryString + "     .Face ""component1:magnet"", ""3""" + vbLf
    sHistoryString = sHistoryString + "     .Transformable ""True""" + vbLf
    sHistoryString = sHistoryString + "     .Create" + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define magnet: magnet1", sHistoryString


    '@ pick face
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "Pick.PickFaceFromId ""component1:magnet"", ""3"""
    AddToHistory "pick face", sHistoryString    
    
    
    '@ define extrude: component1:solid2
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Extrude " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Name ""plate_upper"" " + vbLf
    sHistoryString = sHistoryString + "     .Component ""component1"" " + vbLf
    sHistoryString = sHistoryString + "     .Material ""ST37"" " + vbLf
    sHistoryString = sHistoryString + "     .Mode ""Picks"" " + vbLf
    sHistoryString = sHistoryString + "     .Height ""plate_thickness"" " + vbLf
    sHistoryString = sHistoryString + "     .Twist ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .Taper ""0.0"" " + vbLf
    sHistoryString = sHistoryString + "     .UsePicksForHeight ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .DeleteBaseFaceSolid ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .ClearPickedFace ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .Create " + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "define extrude: component1:solid2", sHistoryString
    
    
    '@ pick circle center point
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "Pick.PickCirclecenterFromId ""component1:plate_upper"", ""3"""
    AddToHistory "pick circle center point", sHistoryString

    
    '@ pick circle center point
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "Pick.PickCirclecenterFromId ""component1:magnet"", ""2"""
    AddToHistory "pick circle center point", sHistoryString


    '@ transform: translate component1:solid2
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Transform " + vbLf
    sHistoryString = sHistoryString + "     .Reset " + vbLf
    sHistoryString = sHistoryString + "     .Name ""component1:plate_upper"" " + vbLf
    sHistoryString = sHistoryString + "     .Vector ""0"", ""0"", ""-plate_thickness-magnet_thickness"" " + vbLf
    sHistoryString = sHistoryString + "     .UsePickedPoints ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .InvertPickedPoints ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .MultipleObjects ""True"" " + vbLf
    sHistoryString = sHistoryString + "     .GroupObjects ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Repetitions ""1"" " + vbLf
    sHistoryString = sHistoryString + "     .MultipleSelection ""False"" " + vbLf
    sHistoryString = sHistoryString + "     .Destination """" " + vbLf
    sHistoryString = sHistoryString + "     .Material """" " + vbLf
    sHistoryString = sHistoryString + "     .Transform ""Shape"", ""Translate"" " + vbLf
    sHistoryString = sHistoryString + "End With"
    AddToHistory "transform: translate component1:solid2", sHistoryString


    '@ set mesh properties (Planar)
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Mesh " + vbLf
    sHistoryString = sHistoryString + "     .MeshType ""Planar"" " + vbLf
    sHistoryString = sHistoryString + "     .SetCreator ""Low Frequency""" + vbLf
    sHistoryString = sHistoryString + "End With " + vbLf
    sHistoryString = sHistoryString + "With MeshSettings " + vbLf
    sHistoryString = sHistoryString + "     .SetMeshType ""Plane"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""Version"", 1%" + vbLf
    sHistoryString = sHistoryString + "     'MAX CELL - GEOMETRY REFINEMENT " + vbLf
    sHistoryString = sHistoryString + "     .Set ""StepsPerBoxNear"", ""50"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""StepsPerBoxFar"", ""5"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""MaxStepNear"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""MaxStepFar"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""ModelBoxDescrNear"", ""maxedge"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""ModelBoxDescrFar"", ""maxedge"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""UseMaxStepAbsolute"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""GeometryRefinementSameAsNear"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     'MIN CELL " + vbLf
    sHistoryString = sHistoryString + "     .Set ""UseRatioLimit"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""RatioLimit"", ""100"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""MinStep"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     '2D CUTPLANE " + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplaneType"", ""Rotational""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplaneRotationAxis"", ""0.0"", ""0.0"", ""1.0""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplaneRVector"", ""1.0"", ""0.0"", ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplanePositionTypeX"", ""Free""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplanePositionTypeY"", ""Center""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplanePositionTypeZ"", ""Center""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplanePosition"", ""0"", ""0.0"", ""0.0""" + vbLf
    sHistoryString = sHistoryString + "     .Set  ""CutplaneDepth"", ""1""" + vbLf
    sHistoryString = sHistoryString + " " + vbLf
    sHistoryString = sHistoryString + "End With " + vbLf
    sHistoryString = sHistoryString + "With MeshSettings " + vbLf
    sHistoryString = sHistoryString + "     .SetMeshType ""Plane"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""CurvatureOrder"", ""1"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""CurvatureOrderPolicy"", ""automatic"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""CurvRefinementControl"", ""NormalTolerance"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""NormalTolerance"", ""22.5"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""SrfMeshGradation"", ""1.2"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""SrfMeshOptimization"", ""1"" " + vbLf
    sHistoryString = sHistoryString + "End With " + vbLf
    sHistoryString = sHistoryString + "With MeshSettings " + vbLf
    sHistoryString = sHistoryString + "     .SetMeshType ""Unstr"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""UseMaterials"",  ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""MoveMesh"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "End With " + vbLf
    sHistoryString = sHistoryString + "With MeshSettings " + vbLf
    sHistoryString = sHistoryString + "     .SetMeshType ""Unstr"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""SmallFeatureSize"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""CoincidenceTolerance"", ""1e-06"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""SelfIntersectionCheck"", ""1"" " + vbLf
    sHistoryString = sHistoryString + "     .Set ""OptimizeForPlanarStructures"", ""0"" " + vbLf
    sHistoryString = sHistoryString + "End With" + vbLf
    sHistoryString = sHistoryString + ""
    AddToHistory "set mesh properties (Planar)", sHistoryString


    '@ define m-static solver parameters
    sHistoryString = ""
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With MStaticSolver" + vbLf
    sHistoryString = sHistoryString + "     .Reset" + vbLf
    sHistoryString = sHistoryString + "     .Method ""Planar Mesh""" + vbLf
    sHistoryString = sHistoryString + "     .Accuracy ""1e-6""" + vbLf
    sHistoryString = sHistoryString + "     .ApparentInductanceMatrix ""False""" + vbLf
    sHistoryString = sHistoryString + "     .IncrementalInductanceMatrix ""False""" + vbLf
    sHistoryString = sHistoryString + "     .StoreResultsInCache ""False""" + vbLf
    sHistoryString = sHistoryString + "     .MeshAdaption ""False""" + vbLf
    sHistoryString = sHistoryString + "     .PrecomputeStationaryCurrentSource ""False""" + vbLf
    sHistoryString = sHistoryString + "     .MaxLinIter ""0""" + vbLf
    sHistoryString = sHistoryString + "     .Preconditioner ""ILU""" + vbLf
    sHistoryString = sHistoryString + "     .IgnorePECMaterial ""False""" + vbLf
    sHistoryString = sHistoryString + "     .EnableDivergenceCheck ""True""" + vbLf
    sHistoryString = sHistoryString + "     .LSESolverType ""Auto""" + vbLf
    sHistoryString = sHistoryString + "     .TetSolverOrder ""2""" + vbLf
    sHistoryString = sHistoryString + "     .TetAdaption ""False""" + vbLf
    sHistoryString = sHistoryString + "     .TetAdaptionMinCycles ""2""" + vbLf
    sHistoryString = sHistoryString + "     .TetAdaptionMaxCycles ""6""" + vbLf
    sHistoryString = sHistoryString + "     .TetAdaptionAccuracy ""0.01""" + vbLf
    sHistoryString = sHistoryString + "     .TetAdaptionRefinementPercentage ""10""" + vbLf
    sHistoryString = sHistoryString + "     .SnapToGeometry ""True""" + vbLf
    sHistoryString = sHistoryString + "     .UseMaxNumberOfThreads ""True""" + vbLf
    sHistoryString = sHistoryString + "     .MaxNumberOfThreads ""96""" + vbLf
    sHistoryString = sHistoryString + "     .MaximumNumberOfCPUDevices ""2""" + vbLf
    sHistoryString = sHistoryString + "     .UseDistributedComputing ""False""" + vbLf
    sHistoryString = sHistoryString + "End With" + vbLf
    sHistoryString = sHistoryString + "UseDistributedComputingForParameters ""False""" + vbLf
    sHistoryString = sHistoryString + "MaxNumberOfDistributedComputingParameters ""2""" + vbLf
    sHistoryString = sHistoryString + "UseDistributedComputingMemorySetting ""False""" + vbLf
    sHistoryString = sHistoryString + "MinDistributedComputingMemoryLimit ""0""" + vbLf
    sHistoryString = sHistoryString + "UseDistributedComputingSharedDirectory ""False"""
    AddToHistory "define m-static solver parameters", sHistoryString    
   
	ResetViewToStructure()
EndHide

End Sub
