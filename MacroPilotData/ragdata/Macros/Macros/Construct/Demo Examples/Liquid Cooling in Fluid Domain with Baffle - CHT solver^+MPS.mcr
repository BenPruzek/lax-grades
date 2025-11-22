'#Language "WWB-COM"

Option Explicit

' =====================================================================================================
' Macro: Creates demo example for the CHT solver with liquid cooling
'
' Copyright 2024-2024 Dassault Systemes Deutschland GmbH
' =====================================================================================================
' History of Changes
' ------------------
' 20-Feb-2024 mha: removed unused mesh groups, changed number/size of baffles, renamed components,
'                  added monitors etc.
' 15-Nov-2023 ywu: first version
' =====================================================================================================

' *** global variables
Dim sCommand As String
Dim sCaption As String

Sub Main
	' define parameters
	StoreDoubleParameter("Q", 2.5)
	StoreDoubleParameter("V_in", 0.5)

	' set up history list
	sCaption = "define units"
	sCommand = ""
	sCommand = sCommand + "With Units" + vbLf
	sCommand = sCommand + "     .Geometry ""cm""" + vbLf
	sCommand = sCommand + "     .Frequency ""Hz""" + vbLf
	sCommand = sCommand + "     .Time ""s""" + vbLf
	sCommand = sCommand + "     .TemperatureUnit ""Celsius""" + vbLf
	sCommand = sCommand + "     .Voltage ""V""" + vbLf
	sCommand = sCommand + "     .Current ""A""" + vbLf
	sCommand = sCommand + "     .Resistance ""Ohm""" + vbLf
	sCommand = sCommand + "     .Conductance ""Siemens""" + vbLf
	sCommand = sCommand + "     .Capacitance ""PikoF""" + vbLf
	sCommand = sCommand + "     .Inductance ""NanoH""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "change solver type"
	sCommand = ""
	sCommand = sCommand + "ChangeSolverType ""Conjugate Heat Transfer""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define material: Aluminum"
	sCommand = ""
	sCommand = sCommand + "With Material" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Aluminum""" + vbLf
	sCommand = sCommand + "     .FrqType ""Static""" + vbLf
	sCommand = sCommand + "     .Type ""Normal""" + vbLf
	sCommand = sCommand + "     .SetMaterialUnit ""Hz"", ""mm""" + vbLf
	sCommand = sCommand + "     .Epsilon ""1""" + vbLf
	sCommand = sCommand + "     .Mu ""1.0""" + vbLf
	sCommand = sCommand + "     .Kappa ""3.56e+007""" + vbLf
	sCommand = sCommand + "     .TanD ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDFreq ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDGiven ""False""" + vbLf
	sCommand = sCommand + "     .TanDModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "     .KappaM ""0""" + vbLf
	sCommand = sCommand + "     .TanDM ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDMFreq ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDMGiven ""False""" + vbLf
	sCommand = sCommand + "     .TanDMModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "     .DispModelEps ""None""" + vbLf
	sCommand = sCommand + "     .DispModelMu ""None""" + vbLf
	sCommand = sCommand + "     .DispersiveFittingSchemeEps ""General 1st""" + vbLf
	sCommand = sCommand + "     .DispersiveFittingSchemeMu ""General 1st""" + vbLf
	sCommand = sCommand + "     .UseGeneralDispersionEps ""False""" + vbLf
	sCommand = sCommand + "     .UseGeneralDispersionMu ""False""" + vbLf
	sCommand = sCommand + "     .FrqType ""All""" + vbLf
	sCommand = sCommand + "     .Type ""Lossy metal""" + vbLf
	sCommand = sCommand + "     .SetMaterialUnit ""GHz"", ""mm""" + vbLf
	sCommand = sCommand + "     .Rho ""2700.0""" + vbLf
	sCommand = sCommand + "     .ThermalType ""Normal""" + vbLf
	sCommand = sCommand + "     .ThermalConductivity ""237.0""" + vbLf
	sCommand = sCommand + "     .SpecificHeat ""900"", ""J/K/kg""" + vbLf
	sCommand = sCommand + "     .MechanicsType ""Isotropic""" + vbLf
	sCommand = sCommand + "     .YoungsModulus ""69""" + vbLf
	sCommand = sCommand + "     .PoissonsRatio ""0.33""" + vbLf
	sCommand = sCommand + "     .ThermalExpansionRate ""23""" + vbLf
	sCommand = sCommand + "     .Colour ""1"", ""1"", ""0""" + vbLf
	sCommand = sCommand + "     .Wireframe ""False""" + vbLf
	sCommand = sCommand + "     .Transparency ""0""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define material: Chip"
	sCommand = ""
	sCommand = sCommand + "With Material" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Chip""" + vbLf
	sCommand = sCommand + "     .Folder """ + vbLf
	sCommand = sCommand + "     .Rho ""1000.0""" + vbLf
	sCommand = sCommand + "     .ThermalType ""Normal""" + vbLf
	sCommand = sCommand + "     .ThermalConductivity ""5""" + vbLf
	sCommand = sCommand + "     .SpecificHeat ""1000"", ""J/K/kg""" + vbLf
	sCommand = sCommand + "     .DynamicViscosity ""0""" + vbLf
	sCommand = sCommand + "     .UseEmissivity ""True""" + vbLf
	sCommand = sCommand + "     .Emissivity ""0""" + vbLf
	sCommand = sCommand + "     .MetabolicRate ""0.0""" + vbLf
	sCommand = sCommand + "     .VoxelConvection ""0.0""" + vbLf
	sCommand = sCommand + "     .BloodFlow ""0""" + vbLf
	sCommand = sCommand + "     .MechanicsType ""Unused""" + vbLf
	sCommand = sCommand + "     .IntrinsicCarrierDensity  ""0""" + vbLf
	sCommand = sCommand + "     .FrqType ""all""" + vbLf
	sCommand = sCommand + "     .Type ""Normal""" + vbLf
	sCommand = sCommand + "     .MaterialUnit  ""Frequency"", ""Hz""" + vbLf
	sCommand = sCommand + "     .MaterialUnit  ""Geometry"", ""cm""" + vbLf
	sCommand = sCommand + "     .MaterialUnit  ""Time"", ""s""" + vbLf
	sCommand = sCommand + "     .MaterialUnit  ""Temperature"", ""Celsius""" + vbLf
	sCommand = sCommand + "     .Epsilon ""1""" + vbLf
	sCommand = sCommand + "     .Mu ""1.0""" + vbLf
	sCommand = sCommand + "     .Kappa ""3.56e+007""" + vbLf
	sCommand = sCommand + "     .TanD ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDFreq ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDGiven ""False""" + vbLf
	sCommand = sCommand + "     .TanDModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "     .KappaM ""0""" + vbLf
	sCommand = sCommand + "     .TanDM ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDMFreq ""0.0""" + vbLf
	sCommand = sCommand + "     .TanDMGiven ""False""" + vbLf
	sCommand = sCommand + "     .TanDMModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "     .DispModelEps ""None""" + vbLf
	sCommand = sCommand + "     .DispModelMu ""None""" + vbLf
	sCommand = sCommand + "     .DispersiveFittingSchemeEps ""General 1st""" + vbLf
	sCommand = sCommand + "     .DispersiveFittingSchemeMu ""General 1st""" + vbLf
	sCommand = sCommand + "     .UseGeneralDispersionEps ""False""" + vbLf
	sCommand = sCommand + "     .UseGeneralDispersionMu ""False""" + vbLf
	sCommand = sCommand + "     .Colour ""0.752941"", ""0.752941"", ""0.752941""" + vbLf
	sCommand = sCommand + "     .Wireframe ""False""" + vbLf
	sCommand = sCommand + "     .Reflection ""False""" + vbLf
	sCommand = sCommand + "     .Allowoutline ""True""" + vbLf
	sCommand = sCommand + "     .Transparentoutline ""False""" + vbLf
	sCommand = sCommand + "     .Transparency ""0""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "change solver type"
	sCommand = ""
	sCommand = sCommand + "ChangeSolverType ""Conjugate Heat Transfer""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define background"
	sCommand = ""
	sCommand = sCommand + "With Background" + vbLf
	sCommand = sCommand + "     .XminSpace 0" + vbLf
	sCommand = sCommand + "     .XmaxSpace 0" + vbLf
	sCommand = sCommand + "     .YminSpace 0" + vbLf
	sCommand = sCommand + "     .YmaxSpace 0" + vbLf
	sCommand = sCommand + "     .ZminSpace 0" + vbLf
	sCommand = sCommand + "     .ZmaxSpace 0" + vbLf
	sCommand = sCommand + "     .ApplyInAllDirections ""True""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "new component: Structure"
	sCommand = ""
	sCommand = sCommand + "Component.New ""Structure""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define brick: Structure:Enclosure"
	sCommand = ""
	sCommand = sCommand + "With Brick" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Enclosure""" + vbLf
	sCommand = sCommand + "     .Component ""Structure""" + vbLf
	sCommand = sCommand + "     .Material ""Aluminum""" + vbLf
	sCommand = sCommand + "     .Xrange ""-2.5"", ""2.5""" + vbLf
	sCommand = sCommand + "     .Yrange ""-1.5"", ""1.5""" + vbLf
	sCommand = sCommand + "     .Zrange ""-0.5"", ""0.5""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "shell object: Structure:Enclosure"
	sCommand = ""
	sCommand = sCommand + "Solid.ShellAdvanced ""Structure:Enclosure"", ""Inside"", ""0.1"", ""True""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick face"
	sCommand = ""
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Enclosure"", ""3""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "align wcs with face"
	sCommand = ""
	sCommand = sCommand + "WCS.AlignWCSWithSelected ""Face""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "move wcs"
	sCommand = ""
	sCommand = sCommand + "WCS.MoveWCS ""local"", -1.85, 0, 0" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define cylinder: Structure:solid4"
	sCommand = ""
	sCommand = sCommand + "With Cylinder" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Insert""" + vbLf
	sCommand = sCommand + "     .Component ""Structure""" + vbLf
	sCommand = sCommand + "     .Material ""Air""" + vbLf
	sCommand = sCommand + "     .OuterRadius 0.3" + vbLf
	sCommand = sCommand + "     .InnerRadius 0" + vbLf
	sCommand = sCommand + "     .Axis ""z""" + vbLf
	sCommand = sCommand + "     .Zrange -0.2, 0.1" + vbLf
	sCommand = sCommand + "     .Xcenter ""0""" + vbLf
	sCommand = sCommand + "     .Ycenter ""0""" + vbLf
	sCommand = sCommand + "     .Segments ""0""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define cylinder: Structure:pipe_in"
	sCommand = ""
	sCommand = sCommand + "With Cylinder" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Pipe_in""" + vbLf
	sCommand = sCommand + "     .Component ""Structure""" + vbLf
	sCommand = sCommand + "     .Material ""Aluminum""" + vbLf
	sCommand = sCommand + "     .OuterRadius 0.375" + vbLf
	sCommand = sCommand + "     .InnerRadius 0.3" + vbLf
	sCommand = sCommand + "     .Axis ""z""" + vbLf
	sCommand = sCommand + "     .Zrange 0, 1" + vbLf
	sCommand = sCommand + "     .Xcenter ""0""" + vbLf
	sCommand = sCommand + "     .Ycenter ""0""" + vbLf
	sCommand = sCommand + "     .Segments ""0""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "activate global coordinates"
	sCommand = ""
	sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick mid point"
	sCommand = ""
	sCommand = sCommand + "Pick.PickMidpointFromId ""Structure:Enclosure"", ""4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "align wcs with point"
	sCommand = ""
	sCommand = sCommand + "WCS.AlignWCSWithSelected ""Point""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "transform: mirror Structure:pipe_in"
	sCommand = ""
	sCommand = sCommand + "With Transform" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Structure:Pipe_in""" + vbLf
	sCommand = sCommand + "     .Origin ""Free""" + vbLf
	sCommand = sCommand + "     .Center ""0"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "     .PlaneNormal ""1"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "     .MultipleObjects ""True""" + vbLf
	sCommand = sCommand + "     .GroupObjects ""False""" + vbLf
	sCommand = sCommand + "     .Repetitions ""1""" + vbLf
	sCommand = sCommand + "     .MultipleSelection ""False""" + vbLf
	sCommand = sCommand + "     .Destination """"" + vbLf
	sCommand = sCommand + "     .Material """"" + vbLf
	sCommand = sCommand + "     .AutoDestination ""True""" + vbLf
	sCommand = sCommand + "     .Transform ""Shape"", ""Mirror""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "rename block:Structure:pipe_in_1 to Structure:pipe_out"
	sCommand = ""
	sCommand = sCommand + "Solid.Rename ""Structure:Pipe_in_1"", ""Pipe_out""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "transform: mirror Structure:Insert"
	sCommand = ""
	sCommand = sCommand + "With Transform" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Structure:Insert""" + vbLf
	sCommand = sCommand + "     .Origin ""Free""" + vbLf
	sCommand = sCommand + "     .Center ""0"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "     .PlaneNormal ""1"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "     .MultipleObjects ""True""" + vbLf
	sCommand = sCommand + "     .GroupObjects ""False""" + vbLf
	sCommand = sCommand + "     .Repetitions ""1""" + vbLf
	sCommand = sCommand + "     .MultipleSelection ""False""" + vbLf
	sCommand = sCommand + "     .Destination """"" + vbLf
	sCommand = sCommand + "     .Material """"" + vbLf
	sCommand = sCommand + "     .AutoDestination ""True""" + vbLf
	sCommand = sCommand + "     .Transform ""Shape"", ""Mirror""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "boolean subtract shapes: Structure:Enclosure, Structure:Insert"
	sCommand = ""
	sCommand = sCommand + "Solid.Subtract ""Structure:Enclosure"", ""Structure:Insert""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "boolean subtract shapes: Structure:Enclosure, Structure:Insert_1"
	sCommand = ""
	sCommand = sCommand + "Solid.Subtract ""Structure:Enclosure"", ""Structure:Insert_1""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "activate global coordinates"
	sCommand = ""
	sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick edge"
	sCommand = ""
	sCommand = sCommand + "Pick.PickEdgeFromId ""Structure:Pipe_in"", ""4"", ""4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define curve item from edges: curve1:edges3"
	sCommand = ""
	sCommand = sCommand + "With EdgeCurve" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""edges3""" + vbLf
	sCommand = sCommand + "     .Curve ""curve1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define coverprofile: Interior boundaries:inlet"
	sCommand = ""
	sCommand = sCommand + "With CoverCurve" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""inlet""" + vbLf
	sCommand = sCommand + "     .Component ""Interior boundaries""" + vbLf
	sCommand = sCommand + "     .Material ""Aluminum""" + vbLf
	sCommand = sCommand + "     .Curve ""curve1:edges3""" + vbLf
	sCommand = sCommand + "     .DeleteCurve ""True""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick edge"
	sCommand = ""
	sCommand = sCommand + "Pick.PickEdgeFromId ""Structure:Pipe_out"", ""4"", ""4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define curve item from edges: curve1:edges4"
	sCommand = ""
	sCommand = sCommand + "With EdgeCurve" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""edges4""" + vbLf
	sCommand = sCommand + "     .Curve ""curve1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define coverprofile: Interior boundaries:outlet"
	sCommand = ""
	sCommand = sCommand + "With CoverCurve" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""outlet""" + vbLf
	sCommand = sCommand + "     .Component ""Interior boundaries""" + vbLf
	sCommand = sCommand + "     .Material ""Aluminum""" + vbLf
	sCommand = sCommand + "     .Curve ""curve1:edges4""" + vbLf
	sCommand = sCommand + "     .DeleteCurve ""True""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick face"
	sCommand = ""
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Enclosure"", ""15""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "align wcs with face"
	sCommand = ""
	sCommand = sCommand + "WCS.AlignWCSWithSelected ""Face""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "store picked point: 1"
	sCommand = ""
	sCommand = sCommand + "Pick.NextPickToDatabase ""1""" + vbLf
	sCommand = sCommand + "Pick.PickMidpointFromId ""Structure:Enclosure"", ""24""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "store picked point: 2"
	sCommand = ""
	sCommand = sCommand + "Pick.NextPickToDatabase ""2""" + vbLf
	sCommand = sCommand + "Pick.PickMidpointFromId ""Structure:Enclosure"", ""18""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define brick: Interior boundaries:Baffle"
	sCommand = ""
	sCommand = sCommand + "With Brick" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Baffle""" + vbLf
	sCommand = sCommand + "     .Component ""Interior boundaries""" + vbLf
	sCommand = sCommand + "     .Material ""Aluminum""" + vbLf
	sCommand = sCommand + "     .Xrange ""xp(1) -0.05"", ""xp(1) + 0.05""" + vbLf
	sCommand = sCommand + "     .Yrange ""yp(2)"", ""yp(1)""" + vbLf
	sCommand = sCommand + "     .Zrange ""0"", ""1.8""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "activate global coordinates"
	sCommand = ""
	sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick face"
	sCommand = ""
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Enclosure"", ""7""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "align wcs with face"
	sCommand = ""
	sCommand = sCommand + "WCS.AlignWCSWithSelected ""Face""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define brick: Structure:Chip1"
	sCommand = ""
	sCommand = sCommand + "With Brick" + vbLf
	sCommand = sCommand + "    .Reset" + vbLf
	sCommand = sCommand + "    .Name ""Chip1""" + vbLf
	sCommand = sCommand + "    .Component ""Structure""" + vbLf
	sCommand = sCommand + "    .Material ""Chip""" + vbLf
	sCommand = sCommand + "    .Xrange ""-1.4"", ""-0.7""" + vbLf
	sCommand = sCommand + "    .Yrange ""-0.9"", ""-0.5""" + vbLf
	sCommand = sCommand + "    .Zrange ""0.0"", ""0.1""" + vbLf
	sCommand = sCommand + "    .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "transform: translate Structure:Chip1"
	sCommand = ""
	sCommand = sCommand + "With Transform" + vbLf
	sCommand = sCommand + "    .Reset" + vbLf
	sCommand = sCommand + "    .Name ""Structure:Chip1""" + vbLf
	sCommand = sCommand + "    .Vector ""0"", ""1.3"", ""0""" + vbLf
	sCommand = sCommand + "    .UsePickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .InvertPickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .MultipleObjects ""True""" + vbLf
	sCommand = sCommand + "    .GroupObjects ""False""" + vbLf
	sCommand = sCommand + "    .Repetitions ""1""" + vbLf
	sCommand = sCommand + "    .MultipleSelection ""False""" + vbLf
	sCommand = sCommand + "    .Destination """"" + vbLf
	sCommand = sCommand + "    .Material """"" + vbLf
	sCommand = sCommand + "    .AutoDestination ""True""" + vbLf
	sCommand = sCommand + "    .Transform ""Shape"", ""Translate""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "rename block: Structure:Chip1_1 to: Structure:Chip2"
	sCommand = ""
	sCommand = sCommand + "Solid.Rename ""Structure:Chip1_1"", ""Chip2""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "transform: translate Structure:Chip1"
	sCommand = ""
	sCommand = sCommand + "With Transform" + vbLf
	sCommand = sCommand + "    .Reset" + vbLf
	sCommand = sCommand + "    .Name ""Structure:Chip1""" + vbLf
	sCommand = sCommand + "    .Vector ""2.1"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "    .UsePickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .InvertPickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .MultipleObjects ""True""" + vbLf
	sCommand = sCommand + "    .GroupObjects ""False""" + vbLf
	sCommand = sCommand + "    .Repetitions ""1""" + vbLf
	sCommand = sCommand + "    .MultipleSelection ""True""" + vbLf
	sCommand = sCommand + "    .Destination """"" + vbLf
	sCommand = sCommand + "    .Material """"" + vbLf
	sCommand = sCommand + "    .AutoDestination ""True""" + vbLf
	sCommand = sCommand + "    .Transform ""Shape"", ""Translate""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "transform: translate Structure:Chip2"
	sCommand = ""
	sCommand = sCommand + "With Transform" + vbLf
	sCommand = sCommand + "    .Reset" + vbLf
	sCommand = sCommand + "    .Name ""Structure:Chip2""" + vbLf
	sCommand = sCommand + "    .Vector ""2.1"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "    .UsePickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .InvertPickedPoints ""False""" + vbLf
	sCommand = sCommand + "    .MultipleObjects ""True""" + vbLf
	sCommand = sCommand + "    .GroupObjects ""False""" + vbLf
	sCommand = sCommand + "    .Repetitions ""1""" + vbLf
	sCommand = sCommand + "    .MultipleSelection ""False""" + vbLf
	sCommand = sCommand + "    .Destination """"" + vbLf
	sCommand = sCommand + "    .Material """"" + vbLf
	sCommand = sCommand + "    .AutoDestination ""True""" + vbLf
	sCommand = sCommand + "    .Transform ""Shape"", ""Translate""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "rename block: Structure:Chip2_1 to: Structure:Chip3"
	sCommand = ""
	sCommand = sCommand + "Solid.Rename ""Structure:Chip2_1"", ""Chip3""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "rename block: Structure:Chip1_1 to: Structure:Chip4"
	sCommand = ""
	sCommand = sCommand + "Solid.Rename ""Structure:Chip1_1"", ""Chip4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "activate global coordinates"
	sCommand = ""
	sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define material: Fluid/Water liquid (20C) (CHT)"
	sCommand = ""
	sCommand = sCommand + "With Material" + vbLf
	sCommand = sCommand + "    .Reset" + vbLf
	sCommand = sCommand + "    .Name ""Water liquid (20C) (CHT)""" + vbLf
	sCommand = sCommand + "    .Folder ""Fluid""" + vbLf
	sCommand = sCommand + "    .FrqType ""all""" + vbLf
	sCommand = sCommand + "    .Type ""Normal""" + vbLf
	sCommand = sCommand + "    .MaterialUnit ""Frequency"", ""GHz""" + vbLf
	sCommand = sCommand + "    .MaterialUnit ""Geometry"", ""mm""" + vbLf
	sCommand = sCommand + "    .MaterialUnit ""Time"", ""s""" + vbLf
	sCommand = sCommand + "    .Epsilon ""78""" + vbLf
	sCommand = sCommand + "    .Mu ""0.999991""" + vbLf
	sCommand = sCommand + "    .Sigma ""1.59""" + vbLf
	sCommand = sCommand + "    .TanD ""0.0""" + vbLf
	sCommand = sCommand + "    .TanDFreq ""0.0""" + vbLf
	sCommand = sCommand + "    .TanDGiven ""False""" + vbLf
	sCommand = sCommand + "    .TanDModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "    .EnableUserConstTanDModelOrderEps ""False""" + vbLf
	sCommand = sCommand + "    .ConstTanDModelOrderEps ""1""" + vbLf
	sCommand = sCommand + "    .SetElParametricConductivity ""False""" + vbLf
	sCommand = sCommand + "    .ReferenceCoordSystem ""Global""" + vbLf
	sCommand = sCommand + "    .CoordSystemType ""Cartesian""" + vbLf
	sCommand = sCommand + "    .SigmaM ""0""" + vbLf
	sCommand = sCommand + "    .TanDM ""0.0""" + vbLf
	sCommand = sCommand + "    .TanDMFreq ""0.0""" + vbLf
	sCommand = sCommand + "    .TanDMGiven ""False""" + vbLf
	sCommand = sCommand + "    .TanDMModel ""ConstTanD""" + vbLf
	sCommand = sCommand + "    .EnableUserConstTanDModelOrderMu ""False""" + vbLf
	sCommand = sCommand + "    .ConstTanDModelOrderMu ""1""" + vbLf
	sCommand = sCommand + "    .SetMagParametricConductivity ""False""" + vbLf
	sCommand = sCommand + "    .DispModelEps  ""None""" + vbLf
	sCommand = sCommand + "    .DispModelMu ""None""" + vbLf
	sCommand = sCommand + "    .DispersiveFittingSchemeEps ""1st Order""" + vbLf
	sCommand = sCommand + "    .DispersiveFittingSchemeMu ""1st Order""" + vbLf
	sCommand = sCommand + "    .UseGeneralDispersionEps ""False""" + vbLf
	sCommand = sCommand + "    .UseGeneralDispersionMu ""False""" + vbLf
	sCommand = sCommand + "    .NonlinearMeasurementError ""1e-1""" + vbLf
	sCommand = sCommand + "    .NLAnisotropy ""False""" + vbLf
	sCommand = sCommand + "    .NLAStackingFactor ""1""" + vbLf
	sCommand = sCommand + "    .NLADirectionX ""1""" + vbLf
	sCommand = sCommand + "    .NLADirectionY ""0""" + vbLf
	sCommand = sCommand + "    .NLADirectionZ ""0""" + vbLf
	sCommand = sCommand + "    .Rho ""998.6""" + vbLf
	sCommand = sCommand + "    .ThermalType ""Normal""" + vbLf
	sCommand = sCommand + "    .ThermalConductivity ""0.5986""" + vbLf
	sCommand = sCommand + "    .SpecificHeat ""4184.2"", ""J/K/kg""" + vbLf
	sCommand = sCommand + "    .DynamicViscosity ""0.001005""" + vbLf
	sCommand = sCommand + "    .ThermalExpansionRateVolume ""207""" + vbLf
	sCommand = sCommand + "    .Emissivity ""0""" + vbLf
	sCommand = sCommand + "    .MetabolicRate ""0""" + vbLf
	sCommand = sCommand + "    .BloodFlow ""0""" + vbLf
	sCommand = sCommand + "    .VoxelConvection ""0""" + vbLf
	sCommand = sCommand + "    .MechanicsType ""Unused""" + vbLf
	sCommand = sCommand + "    .Colour ""0"", ""0"", ""1""" + vbLf
	sCommand = sCommand + "    .Wireframe ""False""" + vbLf
	sCommand = sCommand + "    .Reflection ""False""" + vbLf
	sCommand = sCommand + "    .Allowoutline ""True""" + vbLf
	sCommand = sCommand + "    .Transparentoutline ""False""" + vbLf
	sCommand = sCommand + "    .Transparency ""21""" + vbLf
	sCommand = sCommand + "    .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define a fluid domain: fluiddomain1"
	sCommand = ""
	sCommand = sCommand + "With FluidDomain" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""fluiddomain1""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .CavityMaterial ""Fluid/Water liquid (20C) (CHT)""" + vbLf
	sCommand = sCommand + "     .InvertNormal ""False""" + vbLf
	sCommand = sCommand + "     .AddFace ""Structure:Enclosure"", ""17""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define lid: inlet"
	sCommand = ""
	sCommand = sCommand + "With InteriorBoundary" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""inlet""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .AddFace ""Interior boundaries:inlet"", ""1""" + vbLf
	sCommand = sCommand + "     .Set ""BoundaryType"", ""open""" + vbLf
	sCommand = sCommand + "     .Set ""WallFlow"", ""no-slip"", ""False""" + vbLf
	sCommand = sCommand + "     .Set ""Emissivity"", ""1.0""" + vbLf
	sCommand = sCommand + "     .Set ""Temperature"", ""Ambient""" + vbLf
	sCommand = sCommand + "     .Set ""VolumeFlowRate"", ""V_in"", ""m3/h"", ""False""" + vbLf
	sCommand = sCommand + "     .Create ""Lid""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define lid: outlet"
	sCommand = ""
	sCommand = sCommand + "With InteriorBoundary" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""outlet""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .AddFace ""Interior boundaries:outlet"", ""1""" + vbLf
	sCommand = sCommand + "     .Set ""BoundaryType"", ""open""" + vbLf
	sCommand = sCommand + "     .Set ""WallFlow"", ""no-slip"", ""False""" + vbLf
	sCommand = sCommand + "     .Set ""Emissivity"", ""1.0""" + vbLf
	sCommand = sCommand + "     .Set ""Temperature"", ""Ambient""" + vbLf
	sCommand = sCommand + "     .Set ""Pressure"", ""0.0"", ""Pa"", ""True""" + vbLf
	sCommand = sCommand + "     .Create ""Lid""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define background"
	sCommand = ""
	sCommand = sCommand + "With Background" + vbLf
	sCommand = sCommand + "     .ResetBackground" + vbLf
	sCommand = sCommand + "     .XminSpace 2" + vbLf
	sCommand = sCommand + "     .XmaxSpace 2" + vbLf
	sCommand = sCommand + "     .YminSpace 1" + vbLf
	sCommand = sCommand + "     .YmaxSpace 2" + vbLf
	sCommand = sCommand + "     .ZminSpace 1" + vbLf
	sCommand = sCommand + "     .ZmaxSpace 6" + vbLf
	sCommand = sCommand + "     .ApplyInAllDirections ""False""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	sCommand = sCommand + vbLf
	sCommand = sCommand + "With Material" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Rho ""1.204""" + vbLf
	sCommand = sCommand + "     .ThermalType ""Normal""" + vbLf
	sCommand = sCommand + "     .ThermalConductivity ""0.026""" + vbLf
	sCommand = sCommand + "     .SpecificHeat ""1005"", ""J/K/kg""" + vbLf
	sCommand = sCommand + "     .DynamicViscosity ""1.84e-5""" + vbLf
	sCommand = sCommand + "     .UseEmissivity ""True""" + vbLf
	sCommand = sCommand + "     .Emissivity ""0.0""" + vbLf
	sCommand = sCommand + "     .MetabolicRate ""0.0""" + vbLf
	sCommand = sCommand + "     .VoxelConvection ""0.0""" + vbLf
	sCommand = sCommand + "     .BloodFlow ""0""" + vbLf
	sCommand = sCommand + "     .MechanicsType ""Unused""" + vbLf
	sCommand = sCommand + "     .IntrinsicCarrierDensity ""0""" + vbLf
	sCommand = sCommand + "     .FrqType ""all""" + vbLf
	sCommand = sCommand + "     .Type ""Normal""" + vbLf
	sCommand = sCommand + "     .MaterialUnit ""Frequency"", ""Hz""" + vbLf
	sCommand = sCommand + "     .MaterialUnit ""Geometry"", ""m""" + vbLf
	sCommand = sCommand + "     .MaterialUnit ""Time"", ""s""" + vbLf
	sCommand = sCommand + "     .MaterialUnit ""Temperature"", ""Kelvin""" + vbLf
	sCommand = sCommand + "     .ChangeBackgroundMaterial" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define thermal boundaries"
	sCommand = ""
	sCommand = sCommand + "With Boundary" + vbLf
	sCommand = sCommand + "     .ThermalBoundary ""All"", ""Open""" + vbLf
	sCommand = sCommand + "     .ThermalSymmetry ""X"", ""none""" + vbLf
	sCommand = sCommand + "     .ThermalSymmetry ""Y"", ""none""" + vbLf
	sCommand = sCommand + "     .ThermalSymmetry ""Z"", ""none""" + vbLf
	sCommand = sCommand + "     .ResetBoundaryValues" + vbLf
	sCommand = sCommand + "     .WallFlow ""All"", ""NoSlip""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "set CFD mesh type"
	sCommand = ""
	sCommand = sCommand + "Mesh.MeshType ""CFDNew""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "set mesh properties"
	sCommand = ""
	sCommand = sCommand + "With MeshSettings" + vbLf
	sCommand = sCommand + "     .SetMeshType ""CfdNew""" + vbLf
	sCommand = sCommand + "     .Set ""Version"", ""1%""" + vbLf
	sCommand = sCommand + "     'MAX CELL - GEOMETRY REFINEMENT" + vbLf
	sCommand = sCommand + "     .Set ""StepsPerBoxNear"", 16" + vbLf
	sCommand = sCommand + "     .Set ""StepsPerBoxFar"", 16" + vbLf
	sCommand = sCommand + "     .Set ""MaxStepNear"", 0" + vbLf
	sCommand = sCommand + "     .Set ""MaxStepFar"", 0" + vbLf
	sCommand = sCommand + "     .Set ""ModelBoxDescrNear"", ""maxedge""" + vbLf
	sCommand = sCommand + "     .Set ""ModelBoxDescrFar"", ""maxedge""" + vbLf
	sCommand = sCommand + "     .Set ""UseMaxStepAbsolute"", 0" + vbLf
	sCommand = sCommand + "     .Set ""GeometryRefinementSameAsNear"", ""1""" + vbLf
	sCommand = sCommand + "     'MIN CELL" + vbLf
	sCommand = sCommand + "     .Set ""UseRatioLimitGeometry"", 1" + vbLf
	sCommand = sCommand + "     .Set ""RatioLimitGeometry"", 10" + vbLf
	sCommand = sCommand + "     .Set ""MinStepGeometryX"", 0" + vbLf
	sCommand = sCommand + "     .Set ""MinStepGeometryY"", 0" + vbLf
	sCommand = sCommand + "     .Set ""MinStepGeometryZ"", 0" + vbLf
	sCommand = sCommand + "     .Set ""UseSameMinStepGeometryXYZ"", 1" + vbLf
	sCommand = sCommand + "End With" + vbLf
	sCommand = sCommand + "With MeshSettings" + vbLf
	sCommand = sCommand + "     .SetMeshType ""CfdNew""" + vbLf
	sCommand = sCommand + "     .Set ""Version"", ""1%""" + vbLf
	sCommand = sCommand + "     'OBJECT SETTINGS" + vbLf
	sCommand = sCommand + "     .Set  ""RefinementLevelNear"", 2" + vbLf
	sCommand = sCommand + "     .Set  ""RefinementLevelFar"", 1" + vbLf
	sCommand = sCommand + "     .Set  ""UseSameRefinementLevelForNearAndFar"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""ExtendRangePolicy"", ""RELATIVE""" + vbLf
	sCommand = sCommand + "     .Set  ""RelativeExtendNear"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""RelativeExtendFar"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""AbsoluteExtendNear"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""AbsoluteExtendFar"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""UseSameMaxStepNearXYZ"", ""True""" + vbLf
	sCommand = sCommand + "     .Set  ""MaxStepNearY"", 0" + vbLf
	sCommand = sCommand + "     .Set  ""MaxStepNearZ"", 0" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define heat source: Q1"
	sCommand = ""
	sCommand = sCommand + "With HeatSource" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Q1""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .Value ""Q""" + vbLf
	sCommand = sCommand + "     .ValueType ""Integral""" + vbLf
	sCommand = sCommand + "     .Face ""Structure:Chip1"", ""1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define heat source: Q2"
	sCommand = ""
	sCommand = sCommand + "With HeatSource" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Q2""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .Value ""Q""" + vbLf
	sCommand = sCommand + "     .ValueType ""Integral""" + vbLf
	sCommand = sCommand + "     .Face ""Structure:Chip2"", ""1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define heat source: Q3"
	sCommand = ""
	sCommand = sCommand + "With HeatSource" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Q3""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .Value ""Q""" + vbLf
	sCommand = sCommand + "     .ValueType ""Integral""" + vbLf
	sCommand = sCommand + "     .Face ""Structure:Chip3"", ""1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define heat source: Q4"
	sCommand = ""
	sCommand = sCommand + "With HeatSource" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Q4""" + vbLf
	sCommand = sCommand + "     .Folder """"" + vbLf
	sCommand = sCommand + "     .Enable ""True""" + vbLf
	sCommand = sCommand + "     .Value ""Q""" + vbLf
	sCommand = sCommand + "     .ValueType ""Integral""" + vbLf
	sCommand = sCommand + "     .Face ""Structure:Chip4"", ""1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define CHT solver parameters"
	sCommand = ""
	sCommand = sCommand + "With CHTSolver" + vbLf
	sCommand = sCommand + "     .SolverMode ""Steady-state""" + vbLf
	sCommand = sCommand + "     .FluidFlow ""True""" + vbLf
	sCommand = sCommand + "     .UseGravity ""True""" + vbLf
	sCommand = sCommand + "     .Gravity ""0"", ""0"", ""-9.81""" + vbLf
	sCommand = sCommand + "     .TurbulenceModelChoice ""Automatic""" + vbLf
	sCommand = sCommand + "     .ThermalConduction ""True""" + vbLf
	sCommand = sCommand + "     .AmbientTemperatureUnit ""20"", ""Celsius""" + vbLf
	sCommand = sCommand + "     .Radiation  ""False""" + vbLf
	sCommand = sCommand + "     .SolarRadiation  ""False""" + vbLf
	sCommand = sCommand + "     .AmbientRadiationTemperatureUnit ""20"", ""Celsius""" + vbLf
	sCommand = sCommand + "     .TransientSolverDuration ""10""" + vbLf
	sCommand = sCommand + "     .SetTimeStep ""Method"", ""Adaptive""" + vbLf
	sCommand = sCommand + "     .SetRadiationExplicit ""True""" + vbLf
	sCommand = sCommand + "     .LocalSolidTimeStep ""False""" + vbLf
	sCommand = sCommand + "     .VTKOutput ""False""" + vbLf
	sCommand = sCommand + "     .InitialSteadyStateRelax ""0""" + vbLf
	sCommand = sCommand + "     .TurbulenceLVELMaxViscosity ""0""" + vbLf
	sCommand = sCommand + "     .TurbulenceRelaxation ""0""" + vbLf
	sCommand = sCommand + "     .TurbulenceKEpsMaxViscosity ""0""" + vbLf
	sCommand = sCommand + "     .TurbulenceRampIt ""100""" + vbLf
	sCommand = sCommand + "     .UseMcalcRasterizer ""0""" + vbLf
	sCommand = sCommand + "     .NarrowChannelToAmbient ""0"", ""0"", ""0"", ""0"", ""0"", ""0""" + vbLf
	sCommand = sCommand + "     .UseMaxNumberOfThreads ""True""" + vbLf
	sCommand = sCommand + "     .MaxNumberOfThreads ""4""" + vbLf
	sCommand = sCommand + "     .MaximumNumberOfCPUDevices ""2""" + vbLf
	sCommand = sCommand + "     .UseDistributedComputing ""False""" + vbLf
	sCommand = sCommand + "     .HardwareAcceleration ""False""" + vbLf
	sCommand = sCommand + "     .MaximumNumberOfGPUs ""1""" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "activate global coordinates"
	sCommand = ""
	sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick faces"
	sCommand = ""
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Pipe_out"", ""4""" + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Enclosure"", ""20""" + vbLf
	sCommand = sCommand + vbLf
	sCommand = sCommand + "Pick.PickSolidEdgeChainFromId ""Structure:Enclosure"", ""32"", ""15""" + vbLf
	sCommand = sCommand + "Pick.PickSolidEdgeChainFromId ""Structure:Enclosure"", ""29"", ""15""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Enclosure"", ""14""" + vbLf
	sCommand = sCommand + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Interior boundaries:Baffle"", ""4""" + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Interior boundaries:Baffle"", ""6""" + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Interior boundaries:Baffle"", ""1""" + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Enclosure"", ""19""" + vbLf
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Pipe_in"", ""4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define time monitor 2d: Inner"
	sCommand = ""
	sCommand = sCommand + "With TimeMonitor2D" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Inner""" + vbLf
	sCommand = sCommand + "     .FieldType ""solid flux""" + vbLf
	sCommand = sCommand + "     .InvertOrientation ""False""" + vbLf
	sCommand = sCommand + "     .ReferenceTemperature ""Ambient""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_out"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""20""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""18""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""17""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""16""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""15""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""14""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""13""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Interior boundaries:Baffle"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Interior boundaries:Baffle"", ""6""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Interior boundaries:Baffle"", ""1""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""19""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_in"", ""4""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick faces"
	sCommand = ""
	sCommand = sCommand + "Pick.PickFaceFromId ""Structure:Pipe_out"", ""2""" + vbLf
	sCommand = sCommand + vbLf
	sCommand = sCommand + "Pick.PickSolidEdgeChainFromId ""Structure:Enclosure"", ""30"", ""9""" + vbLf
	sCommand = sCommand + "Pick.PickSolidEdgeChainFromId ""Structure:Enclosure"", ""31"", ""9""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Enclosure"", ""9""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Pipe_in"", ""2""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Chip1"", ""1""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Chip4"", ""1""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Chip3"", ""1""" + vbLf
	sCommand = sCommand + "Pick.PickFaceChainFromId ""Structure:Chip2"", ""4""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define time monitor 2d: Outer"
	sCommand = ""
	sCommand = sCommand + "With TimeMonitor2D" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Outer""" + vbLf
	sCommand = sCommand + "     .FieldType ""solid flux""" + vbLf
	sCommand = sCommand + "     .InvertOrientation ""False""" + vbLf
	sCommand = sCommand + "     .ReferenceTemperature ""Ambient""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_out"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""12""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""11""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""10""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""9""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""8""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Enclosure"", ""7""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_in"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_in"", ""3""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_in"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Pipe_in"", ""1""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""6""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""5""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""3""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip1"", ""1""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""6""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""5""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""3""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip4"", ""1""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""6""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""5""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""3""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip3"", ""1""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""6""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""5""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""4""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""3""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""2""" + vbLf
	sCommand = sCommand + "     .UsePickedFaceFromId ""solid$Structure:Chip2"", ""1""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick center point"
	sCommand = ""
	sCommand = sCommand + "Pick.PickCenterpointFromId ""Structure:Chip1"", ""2""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define time monitor 0d: Under chip"
	sCommand = ""
	sCommand = sCommand + "With TimeMonitor0D" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Under chip""" + vbLf
	sCommand = sCommand + "     .FieldType ""Temperature""" + vbLf
	sCommand = sCommand + "     .Component ""X""" + vbLf
	sCommand = sCommand + "     .UsePickedPoint ""True""" + vbLf
	sCommand = sCommand + "     .Position ""-1.05"", ""-0.7"", ""0.5""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick center point"
	sCommand = ""
	sCommand = sCommand + "Pick.PickCenterpointFromId ""Structure:Chip1"", ""2""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick center point"
	sCommand = ""
	sCommand = sCommand + "Pick.PickCenterpointFromId ""Structure:Enclosure"", ""7""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick mean point"
	sCommand = ""
	sCommand = sCommand + "Pick.MeanLastTwoPoints" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define time monitor 0d: Towards middle"
	sCommand = ""
	sCommand = sCommand + "With TimeMonitor0D" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Towards middle""" + vbLf
	sCommand = sCommand + "     .FieldType ""Temperature""" + vbLf
	sCommand = sCommand + "     .Component ""X""" + vbLf
	sCommand = sCommand + "     .UsePickedPoint ""True""" + vbLf
	sCommand = sCommand + "     .Position ""-0.525"", ""-0.35"", ""0.5""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "pick center point"
	sCommand = ""
	sCommand = sCommand + "Pick.PickCenterpointFromId ""Structure:Enclosure"", ""7""" + vbLf
	AddToHistory sCaption, sCommand

	sCaption = "define time monitor 0d: Middle"
	sCommand = ""
	sCommand = sCommand + "With TimeMonitor0D" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Name ""Middle""" + vbLf
	sCommand = sCommand + "     .FieldType ""Temperature""" + vbLf
	sCommand = sCommand + "     .Component ""X""" + vbLf
	sCommand = sCommand + "     .UsePickedPoint ""True""" + vbLf
	sCommand = sCommand + "     .Position ""0"", ""0"", ""0.5""" + vbLf
	sCommand = sCommand + "     .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory sCaption, sCommand

	' set view and activate bounding box
	ResetViewToStructure()
	Plot.DrawBox(True)
End Sub
