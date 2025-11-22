'#Language "WWB-COM"

' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 18-Jan-2022 ube: move macro location to Macros\Construct\Electronics Cooling, adapt picture path
' 09-Apr-2020 mha: Definition of dUEdgeLength in first version did not use absolute value. Can lead to errors in brick construction.
' 13-Jun-2018 mha: first version
'------------------------------------------------------------------------------------

Option Explicit

Sub Main

	Begin Dialog UserDialog 515, 490, "Create 2R model (classic solver with tet mesh)"
		Text          30,  21, 300,  14, "Name: ",                                 .Text1
		TextBox      340,  18, 155,  21, .Name
		Text          30,  49, 300,  14, "Dissipated power (W):",                 .Text2
		TextBox      340,  46, 155,  21, .DPW
		Text          30,  77, 300,  14, "Case emissivity:",                .Text3
		TextBox      340,  74, 155,  21, .UCE
		Text          30, 105, 300,  14, "Case convection coefficient (W / m^2 / K):",    .Text4
		TextBox      340, 102, 155,  21, .UCCC
		Text          30, 133, 300,  14, "Thermal resistance junction to case (K / W):",  .Text5
		TextBox      340, 130, 155,  21, .TRJTC
		Text          30, 161, 300,  14, "Thermal resistance junction to board (K / W):", .Text6
		TextBox      340, 158, 155,  21, .TRJTB
		Text          30, 189, 300,  14, "Thermal capacitance junction to case (J / K) *:", .Text7
		TextBox      340, 186, 155,  21, .TCJTC
		Text          30, 217, 300,  14, "Thermal capacitance junction to board (J / K) *:", .Text8
		TextBox      340, 214, 155,  21, .TCJTB
		Text          30, 245, 450, 120, "(*Only applicable to THt solver.)" & vbCrLf & _
										 "Select the case top surface and any point on solid touching the board," & vbCrLf & _
										 "then execute this macro.                ", .Text9
		Picture       20, 290, 475, 152, GetInstallPath + "\Library\Macros\Construct\Electronics Cooling\Picture2R.bmp", 16,                            .Picture1
		OKButton      30, 460,  90,  21
		CancelButton 130, 460,  90,  21
	End Dialog

	Dim dlg As UserDialog
	dlg.Name  = "My2R"
	dlg.DPW   = "0.0"
	dlg.UCE   = "0.0"
	dlg.UCCC  = "0.0"
	dlg.TRJTC = "0.0"
	dlg.TRJTB = "0.0"
	dlg.TCJTC = "0.0"
	dlg.TCJTB = "0.0"


	If (Dialog(dlg) = 0) Then Exit All

	Dim sGivenName  As String
	Dim sGivenDPW   As String
	Dim sGivenUCE   As String
	Dim sGivenUCCC  As String
	Dim sGivenTRJTC As String
	Dim sGivenTRJTB As String
	Dim sGivenTCJTC As String
	Dim sGivenTCJTB As String
	sGivenName  = dlg.Name
	sGivenDPW   = dlg.DPW
	sGivenUCE   = dlg.UCE
	sGivenUCCC  = dlg.UCCC
	sGivenTRJTC = dlg.TRJTC
	sGivenTRJTB = dlg.TRJTB
	sGivenTCJTC = dlg.TCJTC
	sGivenTCJTB = dlg.TCJTB

	If sGivenName = "" Then ' empty string has been entered for name...
		ReportInformationToWindow("You have entered an empty value as name. In this case the macro will not do anything...")
	Else
		' Check if entered values are expressions which can be evaluated.
		If sGivenDPW = "" Then ' empty string has been entered for dissipated power - set value to zero...
			sGivenDPW = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenDPW)		' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
        		ReportInformationToWindow("You have entered an expression for the dissipated power which cannot be evaluated at the moment." & vbCrLf & _
		    		                      "If you are using an existing parameter please check the spelling." & vbCrLf & _
               		                	  "If you wish to set up a new parameter please set this up before using it in this dialog.")
        		Exit All
   			End If
		If sGivenUCE = "" Then ' empty string has been entered for upper case emissivity - set value to zero...
			sGivenUCE = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenUCE)		' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
        		ReportInformationToWindow("You have entered an expression for the upper case emissivity which cannot be evaluated at the moment." & vbCrLf & _
		       			                  "If you are using an existing parameter please check the spelling." & vbCrLf & _
                    	   		          "If you wish to set up a new parameter please set this up before using it in this dialog.")
        		Exit All
   			End If
		If sGivenUCCC = "" Then ' empty string has been entered for upper case convection coefficient - set value to zero...
			sGivenUCCC = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenUCCC)	' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
        		ReportInformationToWindow("You have entered an expression for the upper case convection coefficient which cannot be evaluated at the moment." & vbCrLf & _
		   			                      "If you are using an existing parameter please check the spelling." & vbCrLf & _
               			                  "If you wish to set up a new parameter please set this up before using it in this dialog.")
        		Exit All
			End If
		If sGivenTRJTC = "" Then ' empty string has been entered for thermal resistance junction to case - set value to zero...
			sGivenTRJTC = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenTRJTC)	' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
	    		ReportInformationToWindow("You have entered an expression for the thermal resistance from junction to case which cannot be evaluated at the moment." & vbCrLf & _
		       			                  "If you are using an existing parameter please check the spelling." & vbCrLf & _
	               	    		          "If you wish to set up a new parameter please set this up before using it in this dialog.")
	    		Exit All
			End If
		If sGivenTRJTB = "" Then ' empty string has been entered for thermal resistance junction to board - set value to zero...
			sGivenTRJTB = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenTRJTB)	' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
       			ReportInformationToWindow("You have entered an expression for the thermal resistance from junction to board which cannot be evaluated at the moment." & vbCrLf & _
	        			                  "If you are using an existing parameter please check the spelling." & vbCrLf & _
                   				          "If you wish to set up a new parameter please set this up before using it in this dialog.")
       			Exit All
	   		End If
		If sGivenTCJTC = "" Then ' empty string has been entered for thermal capacitance junction to case - set value to zero...
			sGivenTCJTC = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenTCJTC)	' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
       			ReportInformationToWindow("You have entered an expression for the thermal capacitance from junction to case which cannot be evaluated at the moment." & vbCrLf & _
	        			                  "If you are using an existing parameter please check the spelling." & vbCrLf & _
                   				          "If you wish to set up a new parameter please set this up before using it in this dialog.")
       			Exit All
	   		End If
		If sGivenTCJTB = "" Then ' empty string has been entered for thermal capacitance junction to board - set value to zero...
			sGivenTCJTB = "0.0"
		End If
		On Error Resume Next
			evaluate(sGivenTCJTB)	' will cause an error if no such parameter has been set up yet...
			If Err.Number <> 0 Then
       			ReportInformationToWindow("You have entered an expression for the thermal capacitance from junction to board which cannot be evaluated at the moment." & vbCrLf & _
	        			                  "If you are using an existing parameter please check the spelling." & vbCrLf & _
                   				          "If you wish to set up a new parameter please set this up before using it in this dialog.")
       			Exit All
	   		End If
	End If


	' determine number of picked points and faces - one of each is needed for method to work...
	Dim lNumberOfPickedPoints As Long
	Dim lNumberOfPickedFaces  As Long
	Dim iNumberOfPickedFaces  As Long
	Dim sSelectedSolid        As String
	Dim lFaceId			      As Long

	lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints()
	lNumberOfPickedFaces  = Pick.GetNumberOfPickedFaces()

	If lNumberOfPickedPoints < 1 And lNumberOfPickedFaces < 1 Then
		ReportInformationToWindow("Please pick the face of the solid pointing away from the board and a point on the face of solid touching the board. At the moment no points or faces are picked.")
	ElseIf lNumberOfPickedPoints < 1 Then
		ReportInformationToWindow("Please pick a point on the face of solid touching the board. At the moment no points are picked.")
	ElseIf lNumberOfPickedFaces < 1 Then
		ReportInformationToWindow("Please pick the face of the solid pointing away from the board. At the moment no faces are picked.")
	Else
		' get name of solid via picked face
		sSelectedSolid = Pick.GetPickedFaceByIndex(iNumberOfPickedFaces-1, lFaceId)

		' Determine component name and solid name
		Dim sSolidName As String
		Dim sCompoName As String
		sSolidName = Right(sSelectedSolid, Len(sSelectedSolid) - InStrRev(sSelectedSolid, ":"))
		sCompoName = Right(Left(sSelectedSolid, InStrRev(sSelectedSolid, ":")-1), Len(Left(sSelectedSolid, InStrRev(sSelectedSolid, ":")-1)) - InStrRev(Left(sSelectedSolid, InStrRev(sSelectedSolid, ":")-1), ":"))

		' Check to see whether component of designated name already exists
		Dim bsGivenNameInTree     As Double
		bsGivenNameInTree = Resulttree.DoesTreeItemExist("Components\" & sCompoName & "\" & sGivenName)
		Dim Counter       As Integer
		Dim sGivenNameTmp As String
		Counter       = 0
		sGivenNameTmp = sGivenName & "_"
		While bsGivenNameInTree
			sGivenName = sGivenNameTmp & Cstr(Counter)
			bsGivenNameInTree = Resulttree.DoesTreeItemExist("Components\" & sCompoName & "\" & sGivenName)
		Wend

		' Declare variables
		Dim sNameOfLocalWCS  As String
		Dim bDoesVacuumExist As Double
		Dim bDoesPTCExist    As Double

		' Now comes the part which is actually added to the history list...
		Dim sAddToHistoryString As String
		sAddToHistoryString = "' Declare variables" & vbCrLf & _
		"Dim sGivenName            As String" & vbCrLf & _
		"Dim sGivenDPW             As String" & vbCrLf & _
		"Dim sGivenUCE             As String" & vbCrLf & _
		"Dim sGivenUCCC            As String" & vbCrLf & _
		"Dim sGivenTRJTC           As String" & vbCrLf & _
		"Dim sGivenTRJTB           As String" & vbCrLf & _
		"Dim sGivenTCJTC           As String" & vbCrLf & _
		"Dim sGivenTCJTB           As String" & vbCrLf & _
		"Dim lNumberOfPickedPoints As Long" & vbCrLf & _
		"Dim lNumberOfPickedFaces  As Long" & vbCrLf & _
		"Dim lRunningIndex         As Long" & vbCrLf & _
		"Dim bP1                   As Boolean" & vbCrLf & _
		"Dim bP1Convert            As Boolean" & vbCrLf & _
		"Dim dP1xGlobal            As Double" & vbCrLf & _
		"Dim dP1yGlobal            As Double" & vbCrLf & _
		"Dim dP1zGlobal            As Double" & vbCrLf & _
		"Dim dP1xLocal             As Double" & vbCrLf & _
		"Dim dP1yLocal             As Double" & vbCrLf & _
		"Dim dP1zLocal             As Double" & vbCrLf & _
		"Dim bP2                   As Boolean" & vbCrLf & _
		"Dim bP2Convert            As Boolean" & vbCrLf & _
		"Dim dP2xGlobal            As Double" & vbCrLf & _
		"Dim dP2yGlobal            As Double" & vbCrLf & _
		"Dim dP2zGlobal            As Double" & vbCrLf & _
		"Dim dP2xLocal             As Double" & vbCrLf & _
		"Dim dP2yLocal             As Double" & vbCrLf & _
		"Dim dP2zLocal             As Double" & vbCrLf & _
		"Dim dUEdgeLength          As Double" & vbCrLf & _
		"Dim dVEdgeLength          As Double" & vbCrLf & _
		"Dim dAreaOfPickedFace     As Double" & vbCrLf & _
		"Dim dP1ToP2zLocal         As Double" & vbCrLf & _
		"Dim bP3                   As Boolean" & vbCrLf & _
		"Dim bP3Convert            As Boolean" & vbCrLf & _
		"Dim dP3xGlobal            As Double" & vbCrLf & _
		"Dim dP3yGlobal            As Double" & vbCrLf & _
		"Dim dP3zGlobal            As Double" & vbCrLf & _
		"Dim dP3xLocal             As Double" & vbCrLf & _
		"Dim dP3yLocal             As Double" & vbCrLf & _
		"Dim dP3zLocal             As Double" & vbCrLf & _
		"Dim dDeltaDistance        As Double" & vbCrLf & _
		"Dim dOriginActiveLWCSX    As Double" & vbCrLf & _
		"Dim dOriginActiveLWCSY    As Double" & vbCrLf & _
		"Dim dOriginActiveLWCSZ    As Double" & vbCrLf & _
		"Dim dNormalActiveLWCSX    As Double" & vbCrLf & _
		"Dim dNormalActiveLWCSY    As Double" & vbCrLf & _
		"Dim dNormalActiveLWCSZ    As Double" & vbCrLf & _
		"Dim dUVecActiveLWCSX      As Double" & vbCrLf & _
		"Dim dUVecActiveLWCSY      As Double" & vbCrLf & _
		"Dim dUVecActiveLWCSZ      As Double" & vbCrLf & _
		"' Initialize variables" & vbCrLf & _
		"sGivenName  = """ & sGivenName & """" & vbCrLf & _
		"sGivenDPW   = """ & sGivenDPW & """" & vbCrLf & _
		"sGivenUCE   = """ & sGivenUCE & """" & vbCrLf & _
		"sGivenUCCC  = """ & sGivenUCCC & """" & vbCrLf & _
		"sGivenTRJTC = """ & sGivenTRJTC & """" & vbCrLf & _
		"sGivenTRJTB = """ & sGivenTRJTB & """" & vbCrLf & _
		"sGivenTCJTC = """ & sGivenTCJTC & """" & vbCrLf & _
		"sGivenTCJTB = """ & sGivenTCJTB & """" & vbCrLf & vbCrLf

		' Determine whether a local coordiante system is currently active.
		Dim bLWCSActive As Boolean
		bLWCSActive = IIf(WCS.IsWCSActive() = "local", True, False)

		' If local coordinate system is not already active, then activate it.
		If Not bLWCSActive Then
			sAddToHistoryString = sAddToHistoryString & "WCS.ActivateWCS(""local"")" & vbCrLf
		End If

		sAddToHistoryString = sAddToHistoryString & "' Save current position of local coordinate system." & vbCrLf & _
		"WCS.GetOrigin("""",  dOriginActiveLWCSX, dOriginActiveLWCSY, dOriginActiveLWCSZ)" & vbCrLf & _
		"WCS.GetNormal("""",  dNormalActiveLWCSX, dNormalActiveLWCSY, dNormalActiveLWCSZ)" & vbCrLf & _
		"WCS.GetUVector("""", dUVecActiveLWCSX,   dUVecActiveLWCSY,   dUVecActiveLWCSZ)" & vbCrLf & vbCrLf & _
		"lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints()" & vbCrLf & _
		"bP1 = Pick.GetPickpointCoordinatesByIndex(lNumberOfPickedPoints-1, dP1xGlobal, dP1yGlobal, dP1zGlobal)" & vbCrLf & _
		"' Check whether coordinates of picked points could be determined." & vbCrLf & _
		"If Not bP1 Then" & vbCrLf & _
		vbTab & "ReportInformationToWindow(""Unfortunately there was a problem with determining the coordinates of the picked points and the method cannot be used."")" & vbCrLf & _
		"Else" & vbCrLf

		' Check if vacuum / PEC already exists as materials. If not, then set them up
		bDoesVacuumExist = Material.Exists("Vacuum")
		If Not bDoesVacuumExist Then
		sAddToHistoryString = sAddToHistoryString & vbTab & "' Define material Vacuum" & vbCrLf & _
			vbTab & "With Material" & vbCrLf & _
				vbTab & vbTab & ".Reset" & vbCrLf & _
				vbTab & vbTab & ".Name ""Vacuum""" & vbCrLf & _
				vbTab & vbTab & ".Folder """"" & vbCrLf & _
				vbTab & vbTab & ".FrqType ""All""" & vbCrLf & _
				vbTab & vbTab & ".Type ""Normal""" & vbCrLf & _
				vbTab & vbTab & ".SetMaterialUnit ""Hz"", ""mm""" & vbCrLf & _
				vbTab & vbTab & ".Epsilon ""1.0""" & vbCrLf & _
				vbTab & vbTab & ".Mu ""1.0""" & vbCrLf & _
				vbTab & vbTab & ".Kappa ""0""" & vbCrLf & _
				vbTab & vbTab & ".TanD ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDFreq ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDGiven ""False""" & vbCrLf & _
				vbTab & vbTab & ".TanDModel ""ConstKappa""" & vbCrLf & _
				vbTab & vbTab & ".KappaM ""0""" & vbCrLf & _
				vbTab & vbTab & ".TanDM ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDMFreq ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDMGiven ""False""" & vbCrLf & _
				vbTab & vbTab & ".TanDMModel ""ConstKappa""" & vbCrLf & _
				vbTab & vbTab & ".DispModelEps ""None""" & vbCrLf & _
				vbTab & vbTab & ".DispModelMu ""None""" & vbCrLf & _
				vbTab & vbTab & ".DispersiveFittingSchemeEps ""General 1st""" & vbCrLf & _
				vbTab & vbTab & ".DispersiveFittingSchemeMu ""General 1st""" & vbCrLf & _
				vbTab & vbTab & ".UseGeneralDispersionEps ""False""" & vbCrLf & _
				vbTab & vbTab & ".UseGeneralDispersionMu ""False""" & vbCrLf & _
				vbTab & vbTab & ".Rho ""0""" & vbCrLf & _
				vbTab & vbTab & ".ThermalConductivity ""0""" & vbCrLf & _
				vbTab & vbTab & ".SetActiveMaterial ""All""" & vbCrLf & _
				vbTab & vbTab & ".Colour ""0.5"", ""0.82"", ""1""" & vbCrLf & _
				vbTab & vbTab & ".Wireframe ""False""" & vbCrLf & _
				vbTab & vbTab & ".Transparency ""0""" & vbCrLf & _
				vbTab & vbTab & ".Create" & vbCrLf & _
			vbTab & vbTab & "End With" & vbCrLf
		End If

		' Check if PTC already exists as material
		bDoesPTCExist = Material.Exists("PTC")
		If Not bDoesPTCExist Then
			sAddToHistoryString = sAddToHistoryString & vbTab & "' Define material PTC" & vbCrLf & _
			vbTab & "With Material" & vbCrLf & _
				vbTab & vbTab & ".Reset" & vbCrLf & _
				vbTab & vbTab & ".Name ""PTC""" & vbCrLf & _
				vbTab & vbTab & ".Folder """"" & vbCrLf & _
				vbTab & vbTab & ".Rho ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".ThermalType ""PTC""" & vbCrLf & _
				vbTab & vbTab & ".MechanicsType ""Unused""" & vbCrLf & _
				vbTab & vbTab & ".FrqType ""All""" & vbCrLf & _
				vbTab & vbTab & ".Type ""Normal""" & vbCrLf & _
				vbTab & vbTab & ".MaterialUnit ""Frequency"", ""Hz""" & vbCrLf & _
				vbTab & vbTab & ".MaterialUnit ""Geometry"", ""m""" & vbCrLf & _
				vbTab & vbTab & ".MaterialUnit ""Time"", ""s""" & vbCrLf & _
				vbTab & vbTab & ".MaterialUnit ""Temperature"", ""Kelvin""" & vbCrLf & _
				vbTab & vbTab & ".Epsilon ""1""" & vbCrLf & _
				vbTab & vbTab & ".Mu ""1""" & vbCrLf & _
				vbTab & vbTab & ".Sigma ""0""" & vbCrLf & _
				vbTab & vbTab & ".TanD ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDFreq ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDGiven ""False""" & vbCrLf & _
				vbTab & vbTab & ".TanDModel ""ConstTanD""" & vbCrLf & _
				vbTab & vbTab & ".EnableUserConstTanDModelOrderEps ""False""" & vbCrLf & _
				vbTab & vbTab & ".ConstTanDModelOrderEps ""1""" & vbCrLf & _
				vbTab & vbTab & ".SetElParametricConductivity ""False""" & vbCrLf & _
				vbTab & vbTab & ".ReferenceCoordSystem ""Global""" & vbCrLf & _
				vbTab & vbTab & ".CoordSystemType ""Cartesian""" & vbCrLf & _
				vbTab & vbTab & ".SigmaM ""0""" & vbCrLf & _
				vbTab & vbTab & ".TanDM ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDMFreq ""0.0""" & vbCrLf & _
				vbTab & vbTab & ".TanDMGiven ""False""" & vbCrLf & _
				vbTab & vbTab & ".TanDMModel ""ConstTanD""" & vbCrLf & _
				vbTab & vbTab & ".EnableUserConstTanDModelOrderMu ""False""" & vbCrLf & _
				vbTab & vbTab & ".ConstTanDModelOrderMu ""1""" & vbCrLf & _
				vbTab & vbTab & ".SetMagParametricConductivity ""False""" & vbCrLf & _
				vbTab & vbTab & ".DispModelEps ""None""" & vbCrLf & _
				vbTab & vbTab & ".DispModelMu ""None""" & vbCrLf & _
				vbTab & vbTab & ".DispersiveFittingSchemeEps ""Nth Order""" & vbCrLf & _
				vbTab & vbTab & ".MaximalOrderNthModelFitEps ""10""" & vbCrLf & _
				vbTab & vbTab & ".ErrorLimitNthModelFitEps ""0.1""" & vbCrLf & _
				vbTab & vbTab & ".UseOnlyDataInSimFreqRangeNthModelEps ""False""" & vbCrLf & _
				vbTab & vbTab & ".DispersiveFittingSchemeMu ""Nth Order""" & vbCrLf & _
				vbTab & vbTab & ".MaximalOrderNthModelFitMu ""10""" & vbCrLf & _
				vbTab & vbTab & ".ErrorLimitNthModelFitMu ""0.1""" & vbCrLf & _
				vbTab & vbTab & ".UseOnlyDataInSimFreqRangeNthModelMu ""False""" & vbCrLf & _
				vbTab & vbTab & ".UseGeneralDispersionEps ""False""" & vbCrLf & _
				vbTab & vbTab & ".UseGeneralDispersionMu ""False""" & vbCrLf & _
				vbTab & vbTab & ".NonlinearMeasurementError ""1e-1""" & vbCrLf & _
				vbTab & vbTab & ".NLAnisotropy ""False""" & vbCrLf & _
				vbTab & vbTab & ".NLAStackingFactor ""1""" & vbCrLf & _
				vbTab & vbTab & ".NLADirectionX ""1""" & vbCrLf & _
				vbTab & vbTab & ".NLADirectionY ""0""" & vbCrLf & _
				vbTab & vbTab & ".NLADirectionZ ""0""" & vbCrLf & _
				vbTab & vbTab & ".Colour ""0"", ""1"", ""1""" & vbCrLf & _
				vbTab & vbTab & ".Wireframe ""False""" & vbCrLf & _
				vbTab & vbTab & ".Reflection ""False""" & vbCrLf & _
				vbTab & vbTab & ".Allowoutline ""True""" & vbCrLf & _
				vbTab & vbTab & ".Transparentoutline ""False""" & vbCrLf & _
				vbTab & vbTab & ".Transparency ""0""" & vbCrLf & _
				vbTab & vbTab & ".Create" & vbCrLf & _
			vbTab & "End With" & vbCrLf & vbCrLf
		End If

		sAddToHistoryString = sAddToHistoryString & vbTab & "' Align working coordinate system with face" & vbCrLf & _
		vbTab & "WCS.AlignWCSWithSelected(""Face"")" & vbCrLf & _
		vbTab & "' Create dummy-solid for aligning the working coordinate system (take care of u/v-axis)" & vbCrLf & _
		vbTab & "Dim sNameDummySolid      As String" & vbCrLf & _
		vbTab & "Dim sComponentDummySolid As String" & vbCrLf & _
		vbTab & "sNameDummySolid      = Solid.GetNextFreeName()" & vbCrLf & _
		vbTab & "sComponentDummySolid = Resulttree.GetFirstChildName(""Components"")" & vbCrLf & _
		vbTab & "sComponentDummySolid = Right(sComponentDummySolid, Len(sComponentDummySolid) - Len(""Components\""))" & vbCrLf & _
		vbTab & "' Repick face in order to use CreateShapeFromFaces, delete older face picks first" & vbCrLf & _
		vbTab & "lNumberOfPickedFaces = Pick.GetNumberOfPickedFaces()" & vbCrLf & _
		vbTab & "For lRunningIndex = lNumberOfPickedFaces-1 To 0 STEP -1" & vbCrLf & _
			vbTab & vbTab & "Pick.DeleteFace(lRunningIndex)" & vbCrLf & _
		vbTab & "Next" & vbCrLf & _
		vbTab & "Pick.PickFaceFromId(""" & sSelectedSolid & """, """ & Cstr(lFaceId) & """)" & vbCrLf & _
		vbTab & "dAreaOfPickedFace = Pick.GetPickedFaceAreaByIndex(""0"")" & vbCrLf & _
		vbTab & "Solid.CreateShapeFromFaces(sNameDummySolid, sComponentDummySolid, ""Vacuum"")" & vbCrLf & _
		vbTab & "' Pick center of face and query coordinates" & vbCrLf & _
		vbTab & "Pick.PickCenterpointFromId(""" & sSelectedSolid & """, """ & Cstr(lFaceId) & """)" & vbCrLf & _
		vbTab & "lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints()" & vbCrLf & _
		vbTab & "bP2 = Pick.GetPickpointCoordinatesByIndex(lNumberOfPickedPoints-1, dP2xGlobal, dP2yGlobal, dP2zGlobal)" & vbCrLf

		' Find unused name for local WCS
		sNameOfLocalWCS = "Create2RModel"
		While WCS.DoesExist(sNameOfLocalWCS)
			sNameOfLocalWCS = sNameOfLocalWCS & "_01"
		Wend

		sAddToHistoryString = sAddToHistoryString & vbTab & "' Store local working coordinate system in order to convert coordinates" & vbCrLf & _
		vbTab & "WCS.Store(""" & sNameOfLocalWCS & """)" & vbCrLf & _
		vbTab & "bP1Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP1xLocal, dP1yLocal, dP1zLocal, dP1xGlobal, dP1yGlobal, dP1zGlobal)" & vbCrLf & _
		vbTab & "' Find point on edge of dummy solid / use picked point on board surface" & vbCrLf & _
		vbTab & "bP3Convert = WCS.GetGlobalPointFromWCS(""" & sNameOfLocalWCS & """, dP3xGlobal, dP3yGlobal, dP3zGlobal, dP1xLocal, dP1yLocal, ""0.0"")" & vbCrLf & _
		vbTab & "Pick.PickFaceFromPoint(sComponentDummySolid & "":"" & sNameDummySolid, dP2xGlobal, dP2yGlobal, dP2zGlobal)" & vbCrLf & _
		vbTab & "Pick.PickEdgeFromPoint(sComponentDummySolid & "":"" & sNameDummySolid, dP3xGlobal, dP3yGlobal, dP3zGlobal)" & vbCrLf & _
		vbTab & "' Align working coordinate system (this time with u, v aligned with the edges, origin in the middle of the edge)" & vbCrLf & _
		vbTab & "WCS.AlignWCSWithSelected(""EdgeAndFace"")" & vbCrLf & _
    	vbTab & "WCS.Store(""Create2RModel"")" & vbCrLf & _
		vbTab & "' determine length of edges" & vbCrLf & _
	    vbTab & "Pick.PickEndpointFromPoint(sComponentDummySolid & "":"" & sNameDummySolid, dP3xGlobal, dP3yGlobal, dP3zGlobal)" & vbCrLf & _
	    vbTab & "lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints()" & vbCrLf & _
	    vbTab & "bP3 = Pick.GetPickpointCoordinatesByIndex(lNumberOfPickedPoints-1, dP3xGlobal, dP3yGlobal, dP3zGlobal)" & vbCrLf & _
		vbTab & "bP3Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP3xLocal, dP3yLocal, dP3zLocal, dP3xGlobal, dP3yGlobal, dP3zGlobal)" & vbCrLf & _
	    vbTab & "dUEdgeLength = 2.0 * abs(dP3xLocal)" & vbCrLf & _
	    vbTab & "dVEdgeLength = dAreaOfPickedFace / dUEdgeLength" & vbCrLf & _
		vbTab & "' Unpick point on end of edge" & vbCrLf & _
		vbTab & "Pick.DeletePoint(lNumberOfPickedPoints-1)" & vbCrLf & _
		vbTab & "' convert coordinates of picked point to local working coordinate system." & vbCrLf & _
		vbTab & "bP1Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP1xLocal, dP1yLocal, dP1zLocal, dP1xGlobal, dP1yGlobal, dP1zGlobal)" & vbCrLf & _
		vbTab & "' convert coordinates of picked point to local working coordinate system." & vbCrLf & _
		vbTab & "bP1Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP1xLocal, dP1yLocal, dP1zLocal, dP1xGlobal, dP1yGlobal, dP1zGlobal)" & vbCrLf & _
		vbTab & "If Not bP1Convert Then" & vbCrLf & _
			vbTab & vbTab & "ReportInformationToWindow(""Unfortunately there was a problem with determining the coordinates of the picked points and the method cannot be used."")" & vbCrLf & _
		vbTab & "Else" & vbCrLf & _
			vbTab & vbTab & "' Align working coordinate system (with middle of face of dummy solid)" & vbCrLf & _
			vbTab & vbTab & "WCS.AlignWCSWithSelected(""Point"")" & vbCrLf & _
			vbTab & vbTab & "' check whether the distance between face and picked point goes beyond a certain minimum distance" & vbCrLf & _
			vbTab & vbTab & "If Abs(dP1zLocal) < 1.0e-10 Then" & vbCrLf & _
				vbTab & vbTab & vbTab & "ReportInformationToWindow(""Unfortunately the distance between the picked point and the picked Face is too small. Please check that picked face and point are truly opposite or choose a solid with a larger height."")" & vbCrLf & _
			vbTab & vbTab & "Else" & vbCrLf & _
				vbTab & vbTab & vbTab & "If dP1zLocal < 0.0 Then" & vbCrLf & _
				vbTab & vbTab & vbTab & vbTab & "WCS.RotateWCS(""u"", 180)" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Store coordinate system again..." & vbCrLf & _
				vbTab & vbTab & vbTab & "WCS.Store(""" & sNameOfLocalWCS & """)" & vbCrLf & _
				vbTab & vbTab & vbTab & "bP1Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP1xLocal, dP1yLocal, dP1zLocal, dP1xGlobal, dP1yGlobal, dP1zGlobal)" & vbCrLf & _
				vbTab & vbTab & vbTab & "Pick.PickCenterpointFromId(""" & sSelectedSolid & """, " & """" & Cstr(lFaceId) & """)" & vbCrLf & _
				vbTab & vbTab & vbTab & "lNumberOfPickedPoints = Pick.GetNumberOfPickedPoints()" & vbCrLf & _
				vbTab & vbTab & vbTab & "bP2 = Pick.GetPickpointCoordinatesByIndex(lNumberOfPickedPoints-1, dP2xGlobal, dP2yGlobal, dP2zGlobal)" & vbCrLf & _
				vbTab & vbTab & vbTab & "bP2Convert = WCS.GetWCSPointFromGlobal(""" & sNameOfLocalWCS & """, dP2xLocal, dP2yLocal, dP2zLocal, dP2xGlobal, dP2yGlobal, dP2zGlobal)" & vbCrLf & _
				vbTab & vbTab & vbTab & "dP1ToP2zLocal = dP2zLocal - dP1zLocal" & vbCrLf & _
				vbTab & vbTab & vbTab & "' create new component" & vbCrLf & _
				vbTab & vbTab & vbTab & "Component.New(""" & sCompoName & "/" & sGivenName & """)" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Create vacuum sheath" & vbCrLf & _
				vbTab & vbTab & vbTab & "' change material and name of solid used for construction" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.ChangeMaterial(""" & sSelectedSolid & """, ""Vacuum"")" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.ChangeComponent(""" & sSelectedSolid & """, """ & sCompoName & "/" & sGivenName & """)" & vbCrLf & _
				vbTab & vbTab & vbTab & "dDeltaDistance = 0.25 * (dUEdgeLength + dVEdgeLength) - Sqr(0.0625 * (dUEdgeLength + dVEdgeLength) * (dUEdgeLength + dVEdgeLength) - 0.0025 * dUEdgeLength * dVEdgeLength)" & vbCrLf & _
				vbTab & vbTab & vbTab & "With Brick" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Name (""JunctionToCase"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Component(""" & sCompoName & "/" & sGivenName & """)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Material (""PTC"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Xrange (-0.5*dUEdgeLength + dDeltaDistance, 0.5*dUEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Yrange (-0.5*dVEdgeLength + dDeltaDistance, 0.5*dVEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Zrange (0.0, -0.34*dP1ToP2zLocal)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "With Brick" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Name (""Junction"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Component(""" & sCompoName & "/" & sGivenName & """)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Material (""PTC"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Xrange (-0.5*dUEdgeLength + dDeltaDistance, 0.5*dUEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Yrange (-0.5*dVEdgeLength + dDeltaDistance, 0.5*dVEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Zrange (-0.34*dP1ToP2zLocal, -0.66*dP1ToP2zLocal)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "With Brick" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Name (""JunctionToBoard"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Component(""" & sCompoName & "/" & sGivenName & """)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Material (""PTC"")" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Xrange (-0.5*dUEdgeLength + dDeltaDistance, 0.5*dUEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Yrange (-0.5*dVEdgeLength + dDeltaDistance, 0.5*dVEdgeLength - dDeltaDistance)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Zrange (-0.66*dP1ToP2zLocal, -dP1ToP2zLocal)" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Boolean insert bricks" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.Insert(""" & sCompoName & "/" & sGivenName & ":" & sSolidName & """, """ & sCompoName & "/" & sGivenName & ":" & "JunctionToCase"")" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.Insert(""" & sCompoName & "/" & sGivenName & ":" & sSolidName & """, """ & sCompoName & "/" & sGivenName & ":" & "Junction"")" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.Insert(""" & sCompoName & "/" & sGivenName & ":" & sSolidName & """, """ & sCompoName & "/" & sGivenName & ":" & "JunctionToBoard"")" & vbCrLf & _
				vbTab & vbTab & vbTab & "' rename original solid" & vbCrLf & _
				vbTab & vbTab & vbTab & "Solid.Rename(""" & sCompoName & "/" & sGivenName & ":" & sSolidName & """, ""Vacuum Sheath"")" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Set up heat source" & vbCrLf & _
				vbTab & vbTab & vbTab & "If evaluate(sGivenDPW)  < 0.0 Then" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "sGivenDPW = ""0.0""" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The dissipated power value you have chosen for "" & sGivenName & "" is smaller than zero, therefore zero will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "With HeatSource" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"")" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Value sGivenDPW" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ValueType ""Integral""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Face """ & sCompoName & "/" & sGivenName & ":" & "Junction"", ""1""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Type ""PTC""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Define thermal surface property" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Check that values are in proper range" & vbCrLf & _
				vbTab & vbTab & vbTab & "If  evaluate(sGivenUCE)  < 0.0 Then" & vbCrLf & _
				vbTab & vbTab & vbTab & vbTab & "sGivenUCE = ""0.0""" & vbCrLf & _
				vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The emissivity value you have chosen for "" & sGivenName & "" is smaller than zero, therefore zero will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "If  evaluate(sGivenUCE)  > 1.0 Then" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "sGivenUCE = ""1.0""" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The emissivity value you have chosen for "" & sGivenName & "" is larger than one, therefore one will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "If  evaluate(sGivenUCCC) < 0.0 Then" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "sGivenUCCC = ""0.0""" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The convective heat transfer coefficient value you have chosen for "" & sGivenName & "" is smaller than zero, therefore zero will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "With ThermalSurfaceProperty" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"")" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".UseEmissivityValue ""False""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Emissivity sGivenUCE" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ConvectiveHeatTransferCoefficient sGivenUCCC, ""W/m^2/K""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".UseSurrogateHeatTransfer ""False""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".SurrogateHeatTransferCoefficient ""0.0"", ""W/m^2/K""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ReferenceTemperatureType ""Ambient""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".AddFace """ & sCompoName & "/" & sGivenName & ":" & "JunctionToCase"", Cstr(Pick.GetFaceIdFromPoint(""" & sCompoName & "/" & sGivenName & ":" & "JunctionToCase"", dP2xGlobal, dP2yGlobal, dP2zGlobal))" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Set up thermal resistance junction to case" & vbCrLf & _
				vbTab & vbTab & vbTab & "If  evaluate(sGivenTRJTC)  < 0.0 Then" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "sGivenTRJTC = ""0.0""" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The thermal resistance you have chosen for "" & sGivenName & "" from junction to case is smaller than zero, therefore zero will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "With ContactProperties" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"") & ""-Junction-Case""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ElResistance ""0.0""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ThermalResistance sGivenTRJTC" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ThermalCapacitance sGivenTCJTC" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".AddFace """ & sCompoName & "/" & sGivenName & ":" & "Junction"", ""1"", ""1""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".AddFace """ & sCompoName & "/" & sGivenName & ":" & "JunctionToCase"", ""1"", ""2""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Set up thermal resistance junction to board" & vbCrLf & _
				vbTab & vbTab & vbTab & "If  evaluate(sGivenTRJTB)  < 0.0 Then" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "sGivenTRJTB = ""0.0""" & vbCrLf & _
					vbTab & vbTab & vbTab & vbTab & "ReportInformationToWindow(""The thermal resistance you have chosen for "" & sGivenName & "" from junction to board is smaller than zero, therefore zero will be used instead."")" & vbCrLf & _
				vbTab & vbTab & vbTab & "End If" & vbCrLf & _
				vbTab & vbTab & vbTab & "With ContactProperties" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"") & ""-Junction-Board""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ElResistance ""0.0""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ThermalResistance sGivenTRJTB" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".ThermalCapacitance sGivenTCJTB" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".AddFace """ & sCompoName & "/" & sGivenName & ":" & "Junction"", ""1"", ""1""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".AddFace """ & sCompoName & "/" & sGivenName & ":" & "JunctionToBoard"", ""1"", ""2""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "' Set up temperature monitors" & vbCrLf & _
				vbTab & vbTab & vbTab & "WCS.GetGlobalPointFromWCS(""" & sNameOfLocalWCS & """, dP3xGlobal, dP3yGlobal, dP3zGlobal, 0.0, 0.0, Abs(dP1ToP2zLocal)/6.0)" & vbCrLf & _
				vbTab & vbTab & vbTab & "With TimeMonitor0D" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"") & ""-Temperature Case Node""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".FieldType ""Temperature""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Component ""X""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".UsePickedPoint ""False""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Position dP3xGlobal, dP3yGlobal, dP3zGlobal" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "WCS.GetGlobalPointFromWCS(""" & sNameOfLocalWCS & """, dP3xGlobal, dP3yGlobal, dP3zGlobal, 0.0, 0.0, 5*Abs(dP1ToP2zLocal)/6.0)" & vbCrLf & _
				vbTab & vbTab & vbTab & "With TimeMonitor0D" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"") & ""-Temperature Board Node""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".FieldType ""Temperature""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Component ""X""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".UsePickedPoint ""False""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Position dP3xGlobal, dP3yGlobal, dP3zGlobal" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
				vbTab & vbTab & vbTab & "WCS.GetGlobalPointFromWCS(""" & sNameOfLocalWCS & """, dP3xGlobal, dP3yGlobal, dP3zGlobal, 0.0, 0.0, 0.5*Abs(dP1ToP2zLocal))" & vbCrLf & _
				vbTab & vbTab & vbTab & "With TimeMonitor0D" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Reset" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Name Replace(""" & sCompoName & "/" & sGivenName & """, ""/"", ""-"") & ""-Temperature Junction Node""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".FieldType ""Temperature""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Component ""X""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".UsePickedPoint ""False""" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Position dP3xGlobal, dP3yGlobal, dP3zGlobal" & vbCrLf & _
				    vbTab & vbTab & vbTab & vbTab & ".Create" & vbCrLf & _
				vbTab & vbTab & vbTab & "End With" & vbCrLf & _
			vbTab & vbTab & "End If" & vbCrLf & _
		vbTab & "End If" & vbCrLf & vbCrLf & _
		vbTab & "' Put everything back the way it was - delete temporarily stored WCS" & vbCrLf & _
		vbTab & "WCS.Delete(""" & sNameOfLocalWCS & """)" & vbCrLf & _
		vbTab & "' Put everything back the way it was - delete dummy solid" & vbCrLf & _
		vbTab & "Solid.Delete(sComponentDummySolid & "":"" & sNameDummySolid)" & vbCrLf & _
		vbTab & "' Put everything back the way it was - alignment of local working coordinate system" & vbCrLf & _
		vbTab & "WCS.SetOrigin(dOriginActiveLWCSX, dOriginActiveLWCSY, dOriginActiveLWCSZ)" & vbCrLf & _
		vbTab & "WCS.SetNormal(dNormalActiveLWCSX, dNormalActiveLWCSY, dNormalActiveLWCSZ)" & vbCrLf & _
		vbTab & "WCS.SetUVector(dUVecActiveLWCSX,   dUVecActiveLWCSY,   dUVecActiveLWCSZ)" & vbCrLf

		If Not bLWCSActive Then
			sAddToHistoryString = sAddToHistoryString & vbTab & "' Put everything back the way it was - turn off local working coordinate system." & vbCrLf & _
			vbTab & "WCS.ActivateWCS(""Global"")" & vbCrLf
		End If

		sAddToHistoryString = sAddToHistoryString & vbTab & "End If"

		AddToHistory("(*) Create 2R model: " & sCompoName & ":" & sGivenName, sAddToHistoryString)

	End If
End Sub



