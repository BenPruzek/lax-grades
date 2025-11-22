'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Macro: Creates demo example for the classic thermal solver with heat sink
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
	StoreDoubleParameter("Top_convection", 8.5)
	StoreDoubleParameter("Bottom_convection", 10.2)
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
	sContent = sContent & "' define extrude: component1:heatsink" & vbCrLf
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


	sCaption = "pick face chain"

	sContent = ""
	sContent = sContent & "Pick.PickFaceChainFromId ""component1:heatsink"", ""74""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define thermal surface property: Top"

	sContent = ""
	sContent = sContent & "With ThermalSurfaceProperty" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Top""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .Enable ""True""" & vbCrLf
	sContent = sContent & "    .UseEmissivityValue ""False""" & vbCrLf
	sContent = sContent & "    .Emissivity ""0.0""" & vbCrLf
	sContent = sContent & "    .UseAbsorptanceValue ""False""" & vbCrLf
	sContent = sContent & "    .Absorptance ""0.0""" & vbCrLf
	sContent = sContent & "    .ConvectiveHeatTransferCoefficient ""Top_convection"", ""W/m^2/K""" & vbCrLf
	sContent = sContent & "    .UseSurrogateHeatTransfer ""False""" & vbCrLf
	sContent = sContent & "    .SurrogateHeatTransferCoefficient ""0.0"", ""W/m^2/K""" & vbCrLf
	sContent = sContent & "    .HeatTransferCoeffOnlyToVacuum ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceTemperatureType ""Ambient""" & vbCrLf
	sContent = sContent & "    .Coverage ""BackgroundOnly""" & vbCrLf
	sContent = sContent & "    .UsePickedFaces" & vbCrLf
	sContent = sContent & "    .Create" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "pick face chain"

	sContent = ""
	sContent = sContent & "Pick.PickFaceChainFromId ""component1:heatsource"", ""3""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define thermal surface property: Bottom"

	sContent = ""
	sContent = sContent & "With ThermalSurfaceProperty" & vbCrLf
	sContent = sContent & "    .Reset" & vbCrLf
	sContent = sContent & "    .Name ""Bottom""" & vbCrLf
	sContent = sContent & "    .Folder """"" & vbCrLf
	sContent = sContent & "    .Enable ""True""" & vbCrLf
	sContent = sContent & "    .UseEmissivityValue ""False""" & vbCrLf
	sContent = sContent & "    .Emissivity ""0.0""" & vbCrLf
	sContent = sContent & "    .UseAbsorptanceValue ""False""" & vbCrLf
	sContent = sContent & "    .Absorptance ""0.0""" & vbCrLf
	sContent = sContent & "    .ConvectiveHeatTransferCoefficient ""Bottom_convection"", ""W/m^2/K""" & vbCrLf
	sContent = sContent & "    .UseSurrogateHeatTransfer ""False""" & vbCrLf
	sContent = sContent & "    .SurrogateHeatTransferCoefficient ""0.0"", ""W/m^2/K""" & vbCrLf
	sContent = sContent & "    .HeatTransferCoeffOnlyToVacuum ""False""" & vbCrLf
	sContent = sContent & "    .ReferenceTemperatureType ""Ambient""" & vbCrLf
	sContent = sContent & "    .Coverage ""BackgroundOnly""" & vbCrLf
	sContent = sContent & "    .UsePickedFaces" & vbCrLf
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


	sCaption = "define thermal boundaries"

	sContent = ""
	sContent = sContent & "With Boundary" & vbCrLf
	sContent = sContent & "    .ThermalBoundary ""All"", ""adiabatic""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""X"", ""symmetric""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""Y"", ""symmetric""" & vbCrLf
	sContent = sContent & "    .ThermalSymmetry ""Z"", ""none""" & vbCrLf
	sContent = sContent & "    .ResetThermalBoundaryValues" & vbCrLf
	sContent = sContent & "    .WallFlow ""All"", ""NoSlip""" & vbCrLf
	sContent = sContent & "    .EnableThermalRadiation ""All"", ""True""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "change solver type"

	sContent = ""
	sContent = sContent & "ChangeSolverType ""Thermal Steady State""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "set tetrahedral mesh type"

	sContent = ""
	sContent = sContent & "Mesh.MeshType ""Tetrahedral""" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define thermal solver special parameters"

	sContent = ""
	sContent = sContent & "With ThermalSolver" & vbCrLf
	sContent = sContent & "    .NonlinearAccuracy ""1e-6""" & vbCrLf
	sContent = sContent & "    .MaxLinIter ""0""" & vbCrLf
	sContent = sContent & "    .Preconditioner ""ILU""" & vbCrLf
	sContent = sContent & "   .LSESolverType ""Auto""" & vbCrLf
	sContent = sContent & "    .ConsiderBioheat ""True""" & vbCrLf
	sContent = sContent & "    .PTCDefault ""Floating""" & vbCrLf
	sContent = sContent & "    .BloodTemperature ""37.0""" & vbCrLf
	sContent = sContent & "    .TetSolverOrder ""2""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf

	AddToHistory(sCaption, sContent)


	sCaption = "define thermal solver parameters"

	sContent = ""
	sContent = sContent & "With ThermalSolver" & vbCrLf
	sContent = sContent & "    .Accuracy ""1e-6""" & vbCrLf
	sContent = sContent & "    .StoreResultsInCache ""False""" & vbCrLf
	sContent = sContent & "    .AmbientTemperature ""20"", ""degC""" & vbCrLf
	sContent = sContent & "    .Method ""Tetrahedral Mesh""" & vbCrLf
	sContent = sContent & "    .MeshAdaption ""False""" & vbCrLf
	sContent = sContent & "    .CalcThermalConductanceMatrix ""False""" & vbCrLf
	sContent = sContent & "    .TetAdaption ""True""" & vbCrLf
	sContent = sContent & "    .UseMaxNumberOfThreads ""True""" & vbCrLf
	sContent = sContent & "    .MaxNumberOfThreads ""1024""" & vbCrLf
	sContent = sContent & "    .MaximumNumberOfCPUDevices ""2""" & vbCrLf
	sContent = sContent & "    .UseDistributedComputing ""False""" & vbCrLf
	sContent = sContent & "End With" & vbCrLf
	sContent = sContent & "UseDistributedComputingForParameters ""False""" & vbCrLf
	sContent = sContent & "MaxNumberOfDistributedComputingParameters ""2""" & vbCrLf
	sContent = sContent & "UseDistributedComputingMemorySetting ""False""" & vbCrLf
	sContent = sContent & "MinDistributedComputingMemoryLimit ""0""" & vbCrLf
	sContent = sContent & "UseDistributedComputingSharedDirectory ""False""" & vbCrLf
	sContent = sContent & "OnlyConsider0D1DResultsForDC ""False""" & vbCrLf & vbCrLf

	AddToHistory(sCaption, sContent)


	' not in history list...
	ResetViewToStructure()
	Plot.DrawBox(True)
End Sub
