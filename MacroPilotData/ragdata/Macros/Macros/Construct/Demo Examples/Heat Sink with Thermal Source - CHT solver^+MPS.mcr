'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Macro: Creates demo example for the CHT solver with heat sink
'
' Copyright 2023-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 13-Sep-2023 mhh: first version
' ================================================================================================

' *** global variables
Dim sCaption As String
Dim sContent As String


Sub Main
	' define parameters...
	StoreDoubleParameter("Heatcurrent", 5)

	' define geometry and settings...
	sCaption = "define units"

	sContent = ""
	sContent = sContent & "With Units" & vbCrLf
	sContent = sContent & "    .SetUnit ""Length"", ""cm""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Temperature"", ""degC""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Voltage"", ""V""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Current"", ""A""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Resistance"", ""Ohm""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Conductance"", ""S""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Capacitance"", ""pF""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Inductance"", ""nH""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Frequency"", ""Hz""" & vbCrLf
	sContent = sContent & "    .SetUnit ""Time"", ""s""" & vbCrLf
	sContent = sContent & "    .SetResultUnit ""frequency"", ""frequency"", """"" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define material: Aluminum"

	sContent = ""
	sContent = sContent & "With Material" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Aluminum""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .FrqType ""static""" & vbCrLf
	sContent = sContent & "    .Type ""Normal""" & vbCrLf
	sContent = sContent & "    .SetMaterialUnit ""Hz"", ""mm""" & vbCrLf
	sContent = sContent & "    .Epsilon ""1""" & vbCrLf
	sContent = sContent & "    .Mu ""1.0""" & vbCrLf
	sContent = sContent & "    .Kappa ""3.56e+007""" & vbCrLf
	sContent = sContent & "    .TanD ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .KappaM ""0""" & vbCrLf
	sContent = sContent & "    .TanDM ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDMModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .DispModelEps ""None""" & vbCrLf
	sContent = sContent & "    .DispModelMu ""None""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeEps ""General 1st""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeMu ""General 1st""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionEps ""False""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionMu ""False""" & vbCrLf
	sContent = sContent & "    .FrqType ""all""" & vbCrLf
	sContent = sContent & "    .Type ""Lossy metal""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Frequency"", ""GHz""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Geometry"", ""mm""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Time"", ""s""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Temperature"", ""Kelvin""" & vbCrLf
	sContent = sContent & "    .Mu ""1.0""" & vbCrLf
	sContent = sContent & "    .Sigma ""3.56e+007""" & vbCrLf
	sContent = sContent & "    .Rho ""2700.0""" & vbCrLf
	sContent = sContent & "    .ThermalType ""Normal""" & vbCrLf
	sContent = sContent & "    .ThermalConductivity ""237.0""" & vbCrLf
	sContent = sContent & "    .SpecificHeat ""900"", ""J/K/kg""" & vbCrLf
	sContent = sContent & "    .MetabolicRate ""0""" & vbCrLf
	sContent = sContent & "    .BloodFlow ""0""" & vbCrLf
	sContent = sContent & "    .VoxelConvection ""0""" & vbCrLf
	sContent = sContent & "    .MechanicsType ""Isotropic""" & vbCrLf
	sContent = sContent & "    .YoungsModulus ""69""" & vbCrLf
	sContent = sContent & "    .PoissonsRatio ""0.33""" & vbCrLf
	sContent = sContent & "    .ThermalExpansionRate ""23""" & vbCrLf
	sContent = sContent & "    .ReferenceCoordSystem ""Global""" & vbCrLf
	sContent = sContent & "    .CoordSystemType ""Cartesian""" & vbCrLf
	sContent = sContent & "    .NLAnisotropy ""False""" & vbCrLf
	sContent = sContent & "    .NLAStackingFactor ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionX ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionY ""0""" & vbCrLf
	sContent = sContent & "    .NLADirectionZ ""0""" & vbCrLf
	sContent = sContent & "    .Colour ""1"", ""1"", ""0""" & vbCrLf
	sContent = sContent & "    .Wireframe ""False""" & vbCrLf
	sContent = sContent & "    .Reflection ""False""" & vbCrLf
	sContent = sContent & "    .Allowoutline ""True""" & vbCrLf
	sContent = sContent & "    .Transparentoutline ""False""" & vbCrLf
	sContent = sContent & "    .Transparency ""0""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define material: Thermal paste"

	sContent = ""
	sContent = sContent & "With Material" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Thermal paste""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .Rho ""2750""" & vbCrLf
	sContent = sContent & "    .ThermalType ""Normal""" & vbCrLf
	sContent = sContent & "    .ThermalConductivity ""8.0""" & vbCrLf
	sContent = sContent & "    .SpecificHeat ""850"", ""J/K/kg""" & vbCrLf
	sContent = sContent & "    .DynamicViscosity ""0""" & vbCrLf
	sContent = sContent & "    .UseEmissivity ""True""" & vbCrLf
	sContent = sContent & "    .Emissivity ""0""" & vbCrLf
	sContent = sContent & "    .MetabolicRate ""0.0""" & vbCrLf
	sContent = sContent & "    .VoxelConvection ""0.0""" & vbCrLf
	sContent = sContent & "    .BloodFlow ""0""" & vbCrLf
	sContent = sContent & "    .Absorptance ""0""" & vbCrLf
	sContent = sContent & "    .MechanicsType ""Unused""" & vbCrLf
	sContent = sContent & "    .IntrinsicCarrierDensity ""0""" & vbCrLf
	sContent = sContent & "    .FrqType ""all""" & vbCrLf
	sContent = sContent & "    .Type ""Normal""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Frequency"", ""Hz""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Geometry"", ""mm""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Time"", ""s""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Temperature"", ""degC""" & vbCrLf
	sContent = sContent & "    .Epsilon ""1""" & vbCrLf
	sContent = sContent & "    .Mu ""1""" & vbCrLf
	sContent = sContent & "    .Sigma ""0""" & vbCrLf
	sContent = sContent & "    .TanD ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .SetConstTanDStrategyEps ""AutomaticOrder""" & vbCrLf
	sContent = sContent & "    .ConstTanDModelOrderEps ""3""" & vbCrLf
	sContent = sContent & "    .DjordjevicSarkarUpperFreqEps ""0""" & vbCrLf
	sContent = sContent & "    .SetElParametricConductivity ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceCoordSystem ""Global""" & vbCrLf
	sContent = sContent & "    .CoordSystemType ""Cartesian""" & vbCrLf
	sContent = sContent & "    .SigmaM ""0""" & vbCrLf
	sContent = sContent & "    .TanDM ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDMModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .SetConstTanDStrategyMu ""AutomaticOrder""" & vbCrLf
	sContent = sContent & "    .ConstTanDModelOrderMu ""3""" & vbCrLf
	sContent = sContent & "    .DjordjevicSarkarUpperFreqMu ""0""" & vbCrLf
	sContent = sContent & "    .SetMagParametricConductivity ""False""" & vbCrLf
	sContent = sContent & "    .DispModelEps  ""None""" & vbCrLf
	sContent = sContent & "    .DispModelMu ""None""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeEps ""Nth Order""" & vbCrLf
	sContent = sContent & "    .MaximalOrderNthModelFitEps ""10""" & vbCrLf
	sContent = sContent & "    .ErrorLimitNthModelFitEps ""0.1""" & vbCrLf
	sContent = sContent & "    .UseOnlyDataInSimFreqRangeNthModelEps ""False""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeMu ""Nth Order""" & vbCrLf
	sContent = sContent & "    .MaximalOrderNthModelFitMu ""10""" & vbCrLf
	sContent = sContent & "    .ErrorLimitNthModelFitMu ""0.1""" & vbCrLf
	sContent = sContent & "    .UseOnlyDataInSimFreqRangeNthModelMu ""False""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionEps ""False""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionMu ""False""" & vbCrLf
	sContent = sContent & "    .NLAnisotropy ""False""" & vbCrLf
	sContent = sContent & "    .NLAStackingFactor ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionX ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionY ""0""" & vbCrLf
	sContent = sContent & "    .NLADirectionZ ""0""" & vbCrLf
	sContent = sContent & "    .Colour ""0"", ""1"", ""1""" & vbCrLf
	sContent = sContent & "    .Wireframe ""False""" & vbCrLf
	sContent = sContent & "    .Reflection ""False""" & vbCrLf
	sContent = sContent & "    .Allowoutline ""True""" & vbCrLf
	sContent = sContent & "    .Transparentoutline ""False""" & vbCrLf
	sContent = sContent & "    .Transparency ""0""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define background"

	sContent = ""
	sContent = sContent & "With Background" & vbCrLf
	sContent = sContent & "    .ResetBackground" & vbCrLf
	sContent = sContent & "    .XminSpace ""3""" & vbCrLf
	sContent = sContent & "    .XmaxSpace ""3""" & vbCrLf
	sContent = sContent & "    .YminSpace ""3""" & vbCrLf
	sContent = sContent & "    .YmaxSpace ""3""" & vbCrLf
	sContent = sContent & "    .ZminSpace ""3""" & vbCrLf
	sContent = sContent & "    .ZmaxSpace ""9""" & vbCrLf
	sContent = sContent & "    .ApplyInAllDirections ""False""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With Material" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Rho ""1.204""" & vbCrLf
	sContent = sContent & "    .ThermalType ""Normal""" & vbCrLf
	sContent = sContent & "    .ThermalConductivity ""0.026""" & vbCrLf
	sContent = sContent & "    .SpecificHeat ""1005"", ""J/K/kg""" & vbCrLf
	sContent = sContent & "    .DynamicViscosity ""1.84e-5""" & vbCrLf
	sContent = sContent & "    .UseEmissivity ""True""" & vbCrLf
	sContent = sContent & "    .Emissivity ""0.0""" & vbCrLf
	sContent = sContent & "    .MetabolicRate ""0.0""" & vbCrLf
	sContent = sContent & "    .VoxelConvection ""0.0""" & vbCrLf
	sContent = sContent & "    .BloodFlow ""0""" & vbCrLf
	sContent = sContent & "    .Absorptance ""0""" & vbCrLf
	sContent = sContent & "    .MechanicsType ""Unused""" & vbCrLf
	sContent = sContent & "    .IntrinsicCarrierDensity ""0""" & vbCrLf
	sContent = sContent & "    .FrqType ""all""" & vbCrLf
	sContent = sContent & "    .Type ""Normal""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Frequency"", ""Hz""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Geometry"", ""m""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Time"", ""s""" & vbCrLf
	sContent = sContent & "    .MaterialUnit ""Temperature"", ""K""" & vbCrLf
	sContent = sContent & "    .Epsilon ""1.00059""" & vbCrLf
	sContent = sContent & "    .Mu ""1.0""" & vbCrLf
	sContent = sContent & "    .Sigma ""0.0""" & vbCrLf
	sContent = sContent & "    .TanD ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .SetConstTanDStrategyEps ""AutomaticOrder""" & vbCrLf
	sContent = sContent & "    .ConstTanDModelOrderEps ""3""" & vbCrLf
	sContent = sContent & "    .DjordjevicSarkarUpperFreqEps ""0""" & vbCrLf
	sContent = sContent & "    .SetElParametricConductivity ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceCoordSystem ""Global""" & vbCrLf
	sContent = sContent & "    .CoordSystemType ""Cartesian""" & vbCrLf
	sContent = sContent & "    .SigmaM ""0""" & vbCrLf
	sContent = sContent & "    .TanDM ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMFreq ""0.0""" & vbCrLf
	sContent = sContent & "    .TanDMGiven ""False""" & vbCrLf
	sContent = sContent & "    .TanDMModel ""ConstTanD""" & vbCrLf
	sContent = sContent & "    .SetConstTanDStrategyMu ""AutomaticOrder""" & vbCrLf
	sContent = sContent & "    .ConstTanDModelOrderMu ""3""" & vbCrLf
	sContent = sContent & "    .DjordjevicSarkarUpperFreqMu ""0""" & vbCrLf
	sContent = sContent & "    .SetMagParametricConductivity ""False""" & vbCrLf
	sContent = sContent & "    .DispModelEps  ""None""" & vbCrLf
	sContent = sContent & "    .DispModelMu ""None""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeEps ""Nth Order""" & vbCrLf
	sContent = sContent & "    .MaximalOrderNthModelFitEps ""10""" & vbCrLf
	sContent = sContent & "    .ErrorLimitNthModelFitEps ""0.1""" & vbCrLf
	sContent = sContent & "    .UseOnlyDataInSimFreqRangeNthModelEps ""False""" & vbCrLf
	sContent = sContent & "    .DispersiveFittingSchemeMu ""Nth Order""" & vbCrLf
	sContent = sContent & "    .MaximalOrderNthModelFitMu ""10""" & vbCrLf
	sContent = sContent & "    .ErrorLimitNthModelFitMu ""0.1""" & vbCrLf
	sContent = sContent & "    .UseOnlyDataInSimFreqRangeNthModelMu ""False""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionEps ""False""" & vbCrLf
	sContent = sContent & "    .UseGeneralDispersionMu ""False""" & vbCrLf
	sContent = sContent & "    .NLAnisotropy ""False""" & vbCrLf
	sContent = sContent & "    .NLAStackingFactor ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionX ""1""" & vbCrLf
	sContent = sContent & "    .NLADirectionY ""0""" & vbCrLf
	sContent = sContent & "    .NLADirectionZ ""0""" & vbCrLf
	sContent = sContent & "    .Colour ""0.6"", ""0.6"", ""0.6""" & vbCrLf
	sContent = sContent & "    .Wireframe ""False""" & vbCrLf
	sContent = sContent & "    .Reflection ""False""" & vbCrLf
	sContent = sContent & "    .Allowoutline ""True""" & vbCrLf
	sContent = sContent & "    .Transparentoutline ""False""" & vbCrLf
	sContent = sContent & "    .Transparency ""0""" & vbCrLf
	sContent = sContent & "    .ChangeBackgroundMaterial" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "new component: component1"

	sContent = ""
	sContent = sContent & "Component.New ""component1""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define brick: component1:heatsource"

	sContent = ""
	sContent = sContent & "With Brick" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""heatsource""" & vbCrLf
	sContent = sContent & "    .Component ""component1""" & vbCrLf
	sContent = sContent & "    .Material ""Copper (annealed)""" & vbCrLf
	sContent = sContent & "    .Xrange ""-2"", ""2""" & vbCrLf
	sContent = sContent & "    .Yrange ""-2"", ""2""" & vbCrLf
	sContent = sContent & "    .Zrange ""0"", ""0.5""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "construct heatsink"

	sContent = ""
	sContent = sContent & "' pick face" & vbCrLf
	sContent = sContent & "Pick.PickFaceFromId ""component1:heatsource"", ""1""" & vbCrLf & vbCrLf
	sContent = sContent & "' define extrude: component1:heatsink" & vbCrLf & vbCrLf
	sContent = sContent & "With Extrude" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""heatsink""" & vbCrLf
	sContent = sContent & "    .Component ""component1""" & vbCrLf
	sContent = sContent & "    .Material ""Copper (annealed)""" & vbCrLf
	sContent = sContent & "    .Mode ""Picks""" & vbCrLf
	sContent = sContent & "    .Height ""0.125""" & vbCrLf
	sContent = sContent & "    .Twist ""0.0""" & vbCrLf
	sContent = sContent & "    .Taper ""0.0""" & vbCrLf
	sContent = sContent & "    .UsePicksForHeight ""False""" & vbCrLf
	sContent = sContent & "    .DeleteBaseFaceSolid ""False""" & vbCrLf
	sContent = sContent & "    .KeepMaterials ""False""" & vbCrLf
	sContent = sContent & "    .ClearPickedFace ""True""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "' store picked point: 1" & vbCrLf
	sContent = sContent & "Pick.NextPickToDatabase ""1""" & vbCrLf
	sContent = sContent & "Pick.PickEndpointFromId ""component1:heatsink"", ""2""" & vbCrLf & vbCrLf
	sContent = sContent & "' pick face" & vbCrLf
	sContent = sContent & "Pick.PickFaceFromId ""component1:heatsink"", ""5""" & vbCrLf & vbCrLf
	sContent = sContent & "' align wcs with face" & vbCrLf
	sContent = sContent & "WCS.AlignWCSWithSelected ""Face""" & vbCrLf & vbCrLf
	sContent = sContent & "' store picked point: 2" & vbCrLf
	sContent = sContent & "Pick.NextPickToDatabase ""2""" & vbCrLf
	sContent = sContent & "Pick.PickEndpointFromId ""component1:heatsink"", ""2""" & vbCrLf & vbCrLf
	sContent = sContent & "' define brick: component1:heatsink-2" & vbCrLf
	sContent = sContent & "With Brick" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""heatsink-2""" & vbCrLf
	sContent = sContent & "    .Component ""component1""" & vbCrLf
	sContent = sContent & "    .Material ""Copper (annealed)""" & vbCrLf
	sContent = sContent & "    .Xrange ""xp(2)"", ""xp(2) + 0.25""" & vbCrLf
	sContent = sContent & "    .Yrange ""yp(2)"", ""yp(2) + 1""" & vbCrLf
	sContent = sContent & "    .Zrange ""0"", ""2.875""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "' activate global coordinates" & vbCrLf
	sContent = sContent & "WCS.ActivateWCS ""global""" & vbCrLf & vbCrLf
	sContent = sContent & "' transform: translate component1:heatsink-2" & vbCrLf
	sContent = sContent & "With Transform" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""component1:heatsink-2""" & vbCrLf
	sContent = sContent & "    .Vector ""1.25"", ""0"", ""0""" & vbCrLf
	sContent = sContent & "    .UsePickedPoints ""False""" & vbCrLf
	sContent = sContent & "    .InvertPickedPoints ""False""" & vbCrLf
	sContent = sContent & "    .MultipleObjects ""True""" & vbCrLf
	sContent = sContent & "    .GroupObjects ""True""" & vbCrLf
	sContent = sContent & "    .Repetitions ""3""" & vbCrLf
	sContent = sContent & "    .MultipleSelection ""False""" & vbCrLf
	sContent = sContent & "    .Destination """"" & vbCrLf
	sContent = sContent & "    .Material """"" & vbCrLf
	sContent = sContent & "    .AutoDestination ""True""" & vbCrLf
	sContent = sContent & "    .Transform ""Shape"", ""Translate""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "' transform: translate component1:heatsink-2" & vbCrLf
	sContent = sContent & "With Transform" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""component1:heatsink-2""" & vbCrLf
	sContent = sContent & "    .Vector ""0"", ""1.5"", ""0""" & vbCrLf
	sContent = sContent & "    .UsePickedPoints ""False""" & vbCrLf
	sContent = sContent & "    .InvertPickedPoints ""False""" & vbCrLf
	sContent = sContent & "    .MultipleObjects ""True""" & vbCrLf
	sContent = sContent & "    .GroupObjects ""True""" & vbCrLf
	sContent = sContent & "    .Repetitions ""2""" & vbCrLf
	sContent = sContent & "    .MultipleSelection ""False""" & vbCrLf
	sContent = sContent & "    .Destination """"" & vbCrLf
	sContent = sContent & "    .Material """"" & vbCrLf
	sContent = sContent & "    .AutoDestination ""True""" & vbCrLf
	sContent = sContent & "    .Transform ""Shape"", ""Translate""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "' boolean add shapes: component1:heatsink, component1:heatsink-2" & vbCrLf
	sContent = sContent & "Solid.Add ""component1:heatsink"", ""component1:heatsink-2""" & vbCrLf & vbCrLf
	sContent = sContent & "' change material: component1:heatsink to: Aluminum" & vbCrLf
	sContent = sContent & "Solid.ChangeMaterial ""component1:heatsink"", ""Aluminum""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define contact properties: contactprops1"

	sContent = ""
	sContent = sContent & "With ContactProperties" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""contactprops1""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .Enable ""True""" & vbCrLf
	sContent = sContent & "    .Thickness ""0.1""" & vbCrLf
	sContent = sContent & "    .Material ""Thermal paste""" & vbCrLf
	sContent = sContent & "    .NumberOfLayers ""1""" & vbCrLf
	sContent = sContent & "    .AddFace ""component1:heatsource"", ""4"", ""1""" & vbCrLf
	sContent = sContent & "    .AddFace ""component1:heatsink"", ""74"", ""2""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define heat source: heatcurrent1"

	sContent = ""
	sContent = sContent & "With HeatSource" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""heatcurrent1""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .Enable ""True""" & vbCrLf
	sContent = sContent & "    .Value ""Heatcurrent""" & vbCrLf
	sContent = sContent & "    .ValueType ""Integral""" & vbCrLf
	sContent = sContent & "    .Face ""component1:heatsource"", ""4""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "change solver type"

	sContent = ""
	sContent = sContent & "ChangeSolverType ""Conjugate Heat Transfer""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "pick solid edge chain"

	sContent = ""
	sContent = sContent & "Pick.PickSolidEdgeChainFromId ""component1:heatsink"", ""151"", ""78""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "pick face chain"

	sContent = ""
	sContent = sContent & "Pick.PickFaceChainFromId ""component1:heatsink"", ""74""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define time monitor 2d: Top"

	sContent = ""
	sContent = sContent & "With TimeMonitor2D" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Top""" & vbCrLf
	sContent = sContent & "    .FieldType ""solid flux""" & vbCrLf
	sContent = sContent & "    .InvertOrientation ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceTemperature ""Ambient""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""76""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""75""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""74""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""73""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""79""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""12""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""10""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""9""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""7""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""24""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""21""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""19""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""18""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""16""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""15""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""13""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""4""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""3""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""1""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""53""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""52""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""49""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""66""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""65""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""64""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""61""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""72""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""71""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""67""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""60""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""59""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""58""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""55""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""29""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""28""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""27""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""25""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""42""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""41""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""40""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""39""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""37""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""48""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""47""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""45""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""43""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""36""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""35""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""34""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""33""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsink"", ""31""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "pick solid edge chain"

	sContent = ""
	sContent = sContent & "Pick.PickSolidEdgeChainFromId ""component1:heatsource"", ""3"", ""1""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "pick face chain"

	sContent = ""
	sContent = sContent & "Pick.PickFaceChainFromId ""component1:heatsource"", ""4""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define time monitor 2d: Bottom"

	sContent = ""
	sContent = sContent & "With TimeMonitor2D" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Bottom""" & vbCrLf
	sContent = sContent & "    .FieldType ""solid flux""" & vbCrLf
	sContent = sContent & "    .InvertOrientation ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceTemperature ""Ambient""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsource"", ""6""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsource"", ""5""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsource"", ""4""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsource"", ""3""" & vbCrLf
	sContent = sContent & "    .UsePickedFaceFromId ""solid$component1:heatsource"", ""2""" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "set mesh properties (CFD)"

	sContent = ""
	sContent = sContent & "With Mesh" & vbCrLf
	sContent = sContent & "    .MeshType ""CFDNew""" & vbCrLf
	sContent = sContent & "    .SetCreator ""Low Frequency""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .SetMeshType ""CfdNew""" & vbCrLf
	sContent = sContent & "    .Set ""Version"", 1%" & vbCrLf
	sContent = sContent & "    'MAX CELL - WAVELENGTH REFINEMENT" & vbCrLf
	sContent = sContent & "    .Set ""StepsPerWaveNear"", ""10""" & vbCrLf
	sContent = sContent & "    .Set ""StepsPerWaveFar"", ""10""" & vbCrLf
	sContent = sContent & "    .Set ""WavelengthRefinementSameAsNear"", ""1""" & vbCrLf
	sContent = sContent & "    'MAX CELL - GEOMETRY REFINEMENT" & vbCrLf
	sContent = sContent & "    .Set ""StepsPerBoxNear"", ""12""" & vbCrLf
	sContent = sContent & "    .Set ""StepsPerBoxFar"", ""12""" & vbCrLf
	sContent = sContent & "    .Set ""MaxStepNear"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""MaxStepFar"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""ModelBoxDescrNear"", ""maxedge""" & vbCrLf
	sContent = sContent & "    .Set ""ModelBoxDescrFar"", ""maxedge""" & vbCrLf
	sContent = sContent & "    .Set ""UseMaxStepAbsolute"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""GeometryRefinementSameAsNear"", ""1""" & vbCrLf
	sContent = sContent & "    'MIN CELL" & vbCrLf
	sContent = sContent & "    .Set ""UseRatioLimitGeometry"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""RatioLimitGeometry"", ""10""" & vbCrLf
	sContent = sContent & "    .Set ""MinStepGeometryX"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""MinStepGeometryY"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""MinStepGeometryZ"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""UseSameMinStepGeometryXYZ"", ""1""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .Set ""PlaneMergeVersion"", ""2""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .SetMeshType ""CfdNew""" & vbCrLf
	sContent = sContent & "    .Set ""Version"", 1%" & vbCrLf
	sContent = sContent & "    'OBJECT SETTINGS" & vbCrLf
	sContent = sContent & "    .Set  ""RefinementLevelNear"", ""2""" & vbCrLf
	sContent = sContent & "    .Set  ""RefinementLevelFar"", ""1""" & vbCrLf
	sContent = sContent & "    .Set  ""UseSameRefinementLevelForNearAndFar"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""ExtendRangePolicy"", ""RELATIVE""" & vbCrLf
	sContent = sContent & "    .Set  ""RelativeExtendNear"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""RelativeExtendFar"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""AbsoluteExtendNear"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""AbsoluteExtendFar"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""UseSameMaxStepNearXYZ"", ""True""" & vbCrLf
	sContent = sContent & "    .Set  ""MaxStepNearY"", ""0""" & vbCrLf
	sContent = sContent & "    .Set  ""MaxStepNearZ"", ""0""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .SetMeshType ""CfdNew""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementOn"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementPolicy"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementRatio"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementStep"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementNSteps"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""EllipseRefinementOn"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""EllipseRefinementPolicy"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""EllipseRefinementRatio"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""EllipseRefinementStep"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""EllipseRefinementNSteps"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""FaceRefinementBufferLines"", ""3""" & vbCrLf
	sContent = sContent & "    .Set ""EdgeRefinementOn"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""EdgeRefinementPolicy"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""EdgeRefinementRatio"", ""2""" & vbCrLf
	sContent = sContent & "    .Set ""EdgeRefinementStep"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""EdgeRefinementBufferLines"", ""3""" & vbCrLf
	sContent = sContent & "    .Set ""RefineEdgeMaterialGlobal"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""RefineAxialEdgeGlobal"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""BufferLinesNear"", ""3""" & vbCrLf
	sContent = sContent & "    .Set ""UseDielectrics"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""EquilibrateOn"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""Equilibrate"", ""1.2""" & vbCrLf
	sContent = sContent & "    .Set ""IgnoreThinPanelMaterial"", ""0""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .SetMeshType ""CfdNew""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToAxialEdges"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToPlanes"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToSpheres"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToEllipses"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToCylinders"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToCylinderCenters"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToEllipseCenters"", ""1""" & vbCrLf
	sContent = sContent & "    .Set ""SnapToTori"", ""1""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf & vbCrLf
	sContent = sContent & "With MeshSettings" & vbCrLf
	sContent = sContent & "    .SetMeshType ""CfdNew""" & vbCrLf
	sContent = sContent & "    .Set ""EnableNarrowChannelRefinement"", ""False""" & vbCrLf
	sContent = sContent & "    .Set ""MinimumNarrowChannelWidth"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""ConsiderMaximumNarrowChannelWidth"", ""False""" & vbCrLf
	sContent = sContent & "    .Set ""MaximumNarrowChannelWidth"", ""0""" & vbCrLf
	sContent = sContent & "    .Set ""ConsiderMinimumNumNarrowChannelCells"", ""True""" & vbCrLf
	sContent = sContent & "    .Set ""MinimumNumNarrowChannelCells"", ""5""" & vbCrLf
	sContent = sContent & "    .Set ""MaximumNumNarrowChannelRefinements"", ""3""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "set PBA version"

	sContent = ""
	sContent = sContent & "Discretizer.PBAVersion ""2023090824""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define thermal boundaries"

	sContent = ""
	sContent = sContent & "With Boundary" & vbCrLf
	sContent = sContent & "    .ThermalBoundary ""All"", ""open""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""X"", ""symmetric""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""Y"", ""symmetric""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""Z"", ""none""" & vbCrLf
	sContent = sContent & "    .ResetThermalBoundaryValues" & vbCrLf
	sContent = sContent & "    .WallFlow ""All"", ""NoSlip""" & vbCrLf
	sContent = sContent & "    .EnableThermalRadiation ""All"", ""True""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define CHT solver parameters"

	sContent = ""
	sContent = sContent & "With CHTSolver" & vbCrLf
	sContent = sContent & "    .SolverMode ""Steady-state""" & vbCrLf
	sContent = sContent & "    .FluidFlow ""True""" & vbCrLf
	sContent = sContent & "    .UseGravity ""True""" & vbCrLf
	sContent = sContent & "    .Gravity ""0"", ""0"", ""-9.81""" & vbCrLf
	sContent = sContent & "    .TurbulenceModelChoice ""Automatic""" & vbCrLf
	sContent = sContent & "    .ThermalConduction ""True""" & vbCrLf
	sContent = sContent & "    .AmbientTemperatureUnit ""20"", ""degC""" & vbCrLf
	sContent = sContent & "    .Radiation  ""False""" & vbCrLf
	sContent = sContent & "    .SolarRadiation  ""False""" & vbCrLf
	sContent = sContent & "    .AmbientRadiationTemperatureUnit ""293.15"", ""K""" & vbCrLf
	sContent = sContent & "    .TransientSolverDuration ""0""" & vbCrLf
	sContent = sContent & "    .SetTimeStep ""Method"", ""Adaptive""" & vbCrLf
	sContent = sContent & "    .SetRadiationExplicit ""True""" & vbCrLf
	sContent = sContent & "    .LocalSolidTimeStep ""False""" & vbCrLf
	sContent = sContent & "    .VTKOutput ""False""" & vbCrLf
	sContent = sContent & "    .InitialSteadyStateRelax ""0""" & vbCrLf
	sContent = sContent & "    .TurbulenceLVELMaxViscosity ""0""" & vbCrLf
	sContent = sContent & "    .TurbulenceRelaxation ""0""" & vbCrLf
	sContent = sContent & "    .TurbulenceKEpsMaxViscosity ""0""" & vbCrLf
	sContent = sContent & "    .TurbulenceRampIt ""100""" & vbCrLf
	sContent = sContent & "    .UseMcalcRasterizer ""0""" & vbCrLf
	sContent = sContent & "    .NarrowChannelToAmbient ""0"", ""0"", ""0"", ""0"", ""0"", ""0""" & vbCrLf
	sContent = sContent & "    .UseMaxNumberOfThreads ""True""" & vbCrLf
	sContent = sContent & "    .MaxNumberOfThreads ""4""" & vbCrLf
	sContent = sContent & "    .MaximumNumberOfCPUDevices ""2""" & vbCrLf
	sContent = sContent & "    .UseDistributedComputing ""False""" & vbCrLf
	sContent = sContent & "    .HardwareAcceleration ""False""" & vbCrLf
	sContent = sContent & "    .MaximumNumberOfGPUs ""1""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)

	' not in history list...
	ResetViewToStructure()
	Plot.DrawBox(True)
End Sub
