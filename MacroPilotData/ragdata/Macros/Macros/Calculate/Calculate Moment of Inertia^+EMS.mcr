Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

' ================================================================================================
' Macro: calculates moment of inertia
'
' Copyright 2019-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 06-May-2019 mha: first version
' ================================================================================================

' *** global variables
Dim sGlobalSolidArray_CST() As String
Dim dGlobalSolidMassArray() As Double
Dim dGlobalSolidCoMArrayX() As Double
Dim dGlobalSolidCoMArrayY() As Double
Dim dGlobalSolidCoMArrayZ() As Double
Dim iGlobalNumSolids_CST    As Integer
Dim iGlobalSolid_CST        As Integer
Dim dMomentum_xx            As Double
Dim dMomentum_xy            As Double
Dim dMomentum_xz            As Double
Dim dMomentum_yy            As Double
Dim dMomentum_yz            As Double
Dim dMomentum_zz            As Double
Dim dCenterOfMassX          As Double
Dim dCenterOfMassY          As Double
Dim dCenterOfMassZ          As Double
Dim dOverallMass            As Double



Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean) As Boolean
	Dim sTextSolids As String
	sTextSolids = GetScriptSetting("sTextSolids","")

	Begin Dialog UserDialog 640, 180, "Calculate Moment of Inertia",.DialogFunction ' %GRID:3,3,1,1

		' Groupbox - select solids
		GroupBox        9,   6, 349,  70, "Select solids",         .GBSelectSolids
		OptionGroup                                                .OGSelectSolids
		OptionButton   35,  25, 155,  14, "Use selection in tree", .OBSelectionInTree
		OptionButton   35,  46, 120,  14, "Select manually",       .OBSelectManually
		PushButton    220,  16,  90,  21, "Solids...",             .BrowseSolids
		Text          165,  37, 156,  27, sTextSolids,             .TextSolids

		' Groupbox - Specify density
		GroupBox        9,  85, 349,  57, "Specify density",          .GBSpecifyDensity
		CheckBox       30, 108, 180,  14, "Specify density (kg/m^3)", .CBSpecifyDensity
		TextBox       220, 105, 125,  21,                             .Density

		' Groupbox - Specify origin of rotational axis
		GroupBox      375,   6, 248, 137, "Specify origin of rotational axis",               .GBSpecifyOrigin
		CheckBox      395,  25, 188,  14, "Use CoM (Center of Mass)",                        .CBUseCenterOfMass
		Text          395,  53,  48,  15, "X " & "(" & CStr(Units.GetUnit("Length")) & "):", .TX
		TextBox       449,  49, 158,  21, .XCoorRotAx
		Text          395,  83,  48,  15, "Y " & "(" & CStr(Units.GetUnit("Length")) & "):", .TY
		TextBox       449,  79, 158,  21, .YCoorRotAx
		Text          395, 113,  48,  15, "Z " & "(" & CStr(Units.GetUnit("Length")) & "):", .TZ
		TextBox       449, 109, 158,  21, .ZCoorRotAx

		' OK and Cancel buttons
		OKButton       15, 150,  90,  21
		CancelButton  123, 150,  90,  21
	End Dialog

		' Initialize / retrieve script settings...
	Dim dlg               As UserDialog
	Dim lOrientation      As Long
	Dim lSpecifiedDensity As Long
	Dim dDensity          As Double
	Dim lUseCenterOfMass  As Long
	Dim dOriginXCoor      As Double
	Dim dOriginYCoor      As Double
	Dim dOriginZCoor      As Double
	Dim sSolidText        As String
	Dim sSelectedSolid    As String		' Name of selected solid (selected in tree)
	Dim bSelectedSolid    As Boolean	' True -> a solid is selected in the tree; False -> the selection in the tree is not a solid...
	Dim sSolidName        As String		' Only name of selected solid (without component etc.)
	Dim sCompoName        As String		' Only name of component of selected solid
	Dim sNewSolidName     As String		' Name of selected solid, re-formated to be used in Solid-Object methods

	' Make sure that a solid is selected
	bSelectedSolid = True
	sSelectedSolid = GetSelectedTreeItem()
	If InStr(sSelectedSolid, "\") = 0 Then
		' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
		bSelectedSolid = False
	Else
		If Left(sSelectedSolid, InStr(sSelectedSolid, "\")-1) <> "Components" Then
			' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
			bSelectedSolid = False
		ElseIf Resulttree.GetFirstChildName(GetSelectedTreeItem()) <> "" Then
			' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
			bSelectedSolid = False
		Else
			' Determine component name and solid name
			sSolidName = Right(sSelectedSolid, Len(sSelectedSolid) - InStrRev(sSelectedSolid, "\"))
			sCompoName = Right(sSelectedSolid, Len(sSelectedSolid) - Len("Components\"))
			sCompoName = Left(sCompoName, Len(sCompoName) - Len(sSolidName) - 1)
			sCompoName = Replace(sCompoName, "\", "/")
			sNewSolidName = sCompoName & ":" & sSolidName
		End If
	End If

	If bSelectedSolid Then
		iGlobalNumSolids_CST     = CInt(GetScriptSetting("SSnSolids", "1"))
		ReDim sGlobalSolidArray_CST(iGlobalNumSolids_CST-1)
		sGlobalSolidArray_CST(0) = sNewSolidName
		StoreScriptSetting("SSSolid" + CStr(iGlobalNumSolids_CST-1), sGlobalSolidArray_CST(iGlobalNumSolids_CST-1))
		StoreScriptSetting("SSTextSolids", sNewSolidName)
	Else
		iGlobalNumSolids_CST     = CInt(GetScriptSetting("SSnSolids", "0"))
		StoreScriptSetting("SSTextSolids", "")
	End If


	If (Not Dialog(dlg)) Then
		' The user left the dialog box without pressing Ok. Assigning False to the function will cause the framework to cancel the creation or modification without storing anything.
		Define = False
	Else
		' The user properly left the dialog box by pressing Ok. Assigning True to the function will cause the framework to complete the creation or modification and store the corresponding settings.
		Define = True
		' Store the script settings into the database for later reuse by either the define function (for modifications) or the evaluate function.

		StoreScriptSetting("SSCBSpecifyDensity", CStr(dlg.CBSpecifyDensity))
		StoreScriptSetting("SSDensity",          dlg.Density)
		StoreScriptSetting("SSOriginXCoor",      dlg.XCoorRotAx)
		StoreScriptSetting("SSOriginYCoor",      dlg.YCoorRotAx)
		StoreScriptSetting("SSOriginZCoor",      dlg.ZCoorRotAx)
		StoreScriptSetting("SSnSolids",          CStr(iGlobalNumSolids_CST))
		StoreScriptSetting("SSOGSelectSolids",   CStr(dlg.OGSelectSolids))
		StoreScriptSetting("SSCenterOfMass",     CStr(dlg.CBUseCenterOfMass))


		For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST STEP 1
			StoreScriptSetting("SSSolid" + CStr(iGlobalSolid_CST), sGlobalSolidArray_CST(iGlobalSolid_CST-1))
		Next iGlobalSolid_CST

		StoreTemplateSetting("TemplateType", "0D")
		StoreTemplateSetting("EvaluationType", "single_run")

		If (Not bNameChanged) Then
		    sName = "Calculate moment of inertia"
		    sName = NoForbiddenFilenameCharacters(sName)
		End If
	End If
End Function



' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Private Function DialogFunction(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim newname As String, sText As String

	Select Case iAction
		Case 1	' Dialog Box initialisation.
			sText = GetScriptSetting("SSTextSolids", "")
			DlgText "TextSolids",         sText
			DlgValue "CBUseCenterOfMass", CLng(GetScriptSetting("SSCenterOfMass", "0"))
			DlgValue "CBSpecifyDensity",  CLng(GetScriptSetting("SSCBSpecifyDensity", "0"))
			DlgValue "OGSelectSolids",    CLng(GetScriptSetting("SSOGSelectSolids", "0"))
			DlgText "XCoorRotAx",         CStr(GetScriptSetting("SSOriginXCoor", "0.0"))
			DlgText "YCoorRotAx",         CStr(GetScriptSetting("SSOriginYCoor", "0.0"))
			DlgText "ZCoorRotAx",         CStr(GetScriptSetting("SSOriginZCoor", "0.0"))
			DlgText "Density",            CStr(GetScriptSetting("SSDensity", "7850.0"))

			If ( DlgValue("OGSelectSolids") = 0 ) Then
				DlgEnable "BrowseSolids", False
			Else
				DlgEnable "BrowseSolids", True
			End If
			If ( DlgValue("OGSelectSolids") = 0 ) Then
				DlgText   "Density", DlgText("Density")
				DlgEnable "Density", False
			Else
				DlgText   "Density", DlgText("Density")
				DlgEnable "Density", True
			End If
			DlgText   "XCoorRotAx"       , GetScriptSetting("SSOriginXCoor", "0.0")
			DlgEnable "XCoorRotAx"       , True
			DlgText   "YCoorRotAx"       , GetScriptSetting("SSOriginYCoor", "0.0")
			DlgEnable "YCoorRotAx"       , True
			DlgText   "ZCoorRotAx"       , GetScriptSetting("SSOriginZCoor", "0.0")
			DlgEnable "ZCoorRotAx"       , True
			DlgEnable "CBUseCenterOfMass", True
		Case 2 ' Value changing or button pressed
			If ( sDlgItem = "OGSelectSolids" ) Then
				If ( DlgValue("OGSelectSolids") = 0 ) Then	' Use selection from tree
					DlgEnable "BrowseSolids", False
					Dim sSelectedSolid    As String		' Name of selected solid (selected in tree)
					Dim bSelectedSolid    As Boolean	' True -> a solid is selected in the tree; False -> the selection in the tree is not a solid...
					Dim sTreeSolidName    As String		' Only name of selected solid (without component etc.)
					Dim sCompoName        As String		' Only name of component of selected solid
					Dim sNewSolidName     As String		' Name of selected solid, re-formated to be used in Solid-Object methods
					' Make sure that a solid is selected
					bSelectedSolid = True
					sSelectedSolid = GetSelectedTreeItem()
					If InStr(sSelectedSolid, "\") = 0 Then
						' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
						bSelectedSolid = False
					Else
						If Left(sSelectedSolid, InStr(sSelectedSolid, "\")-1) <> "Components" Then
							' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
							bSelectedSolid = False
						ElseIf Resulttree.GetFirstChildName(GetSelectedTreeItem()) <> "" Then
							' ReportInformationToWindow("For this method to work properly it is necessary to select a solid.")
							bSelectedSolid = False
						Else
							' Determine component name and solid name
							sTreeSolidName = Right(sSelectedSolid, Len(sSelectedSolid) - InStrRev(sSelectedSolid, "\"))
							sCompoName     = Right(sSelectedSolid, Len(sSelectedSolid) - Len("Components\"))
							sCompoName     = Left(sCompoName, Len(sCompoName) - Len(sTreeSolidName) - 1)
							sCompoName     = Replace(sCompoName, "\", "/")
							sNewSolidName  = sCompoName & ":" & sTreeSolidName
						End If
					End If
					If bSelectedSolid Then
						StoreScriptSetting("SSnSolids", "1")
						iGlobalNumSolids_CST     = 1
						ReDim sGlobalSolidArray_CST(1)
						sGlobalSolidArray_CST(0) = sNewSolidName
						StoreScriptSetting("SSSolid" + CStr(iGlobalNumSolids_CST-1), sGlobalSolidArray_CST(iGlobalNumSolids_CST-1))
						StoreScriptSetting("SSTextSolids", sNewSolidName)
					Else
						iGlobalNumSolids_CST     = CInt(GetScriptSetting("SSnSolids", "0"))
						StoreScriptSetting("SSTextSolids", "")
					End If
					DlgText "TextSolids", GetScriptSetting("SSTextSolids", "")
				Else
					DlgEnable "BrowseSolids", True
					sText = ""
					StoreScriptSetting("SSTextSolids", sText)
					StoreScriptSetting("SSnSolids", "0")
					DlgText "TextSolids", sText
					iGlobalNumSolids_CST = CInt(GetScriptSetting("SSnSolids", "0"))
				End If
			ElseIf ( sDlgItem = "BrowseSolids" ) Then
				DialogFunction = True       ' Don't close the dialog box.
				SelectSolids_LIB sGlobalSolidArray_CST(), iGlobalNumSolids_CST
				sText = ""
				For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST STEP 1
					sText = sText + sGlobalSolidArray_CST(iGlobalSolid_CST-1)
					If (iGlobalSolid_CST < iGlobalNumSolids_CST) Then
						sText = sText + "; "
						If Len(sText) > 25 Then
							sText = sText + "..."
							Exit For
						End If
					End If
				Next iGlobalSolid_CST
				DlgText("TextSolids"), sText
				StoreScriptSetting("SSTextSolids", sText)
			ElseIf ( sDlgItem = "OK" ) Then
			' The user pressed the Ok button. Check the settings and display an error message if some required fields have been left blank.
				' Verify that solids have been selected.
				If (iGlobalNumSolids_CST = 0) Then
					MsgBox "No solid has been selected." & vbCrLf & "Please specify at least one solid.", vbExclamation
					DialogFunction = True	' There is an error in the settings -> Don't close the dialog box.
				End If
				' Check that all selected solids have a density that is greater zero
				If ( DlgValue("CBSpecifyDensity") = 0 ) Then
					Dim sSolidName    As String
					Dim sMaterialName As String
					Dim dSolidDensity As Double
					For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST STEP 1
						sSolidName    = sGlobalSolidArray_CST(iGlobalSolid_CST-1)
						sMaterialName = Solid.GetMaterialNameForShape(sSolidName)
						Material.GetRho(sMaterialName, dSolidDensity)
						If dSolidDensity <= 0.0 Then
							MsgBox "At least one of the selected solids has a density below or equal zero. Please specify a density or change the density value for the material." & vbCrLf & "Solid: " & sSolidName & vbCrLf & "Material: " & sMaterialName & vbCrLf & "Density: " & CStr(dSolidDensity), vbExclamation
							DialogFunction = True	' There is an error in the settings -> Don't close the dialog box.
							Exit For
						End If
					Next iGlobalSolid_CST
				Else
					If DlgText("Density") = "" Then
						MsgBox "You have enabled ""Specify density"" but not specified any density. Please enter a value in the corresponding textbox.", vbExclamation
						DialogFunction = True	' There is an error in the settings -> Don't close the dialog box.
					ElseIf CDbl(DlgText("Density")) <= 0.0 Then
						MsgBox "You have enabled ""Specify density"" but specified a density that is smaller than or equal to zero. Please enter a positive value in the corresponding textbox.", vbExclamation
						DialogFunction = True	' There is an error in the settings -> Don't close the dialog box.
					End If
				End If
				' Verify that all enabled textboxes are filled out with values.
				If ( DlgValue("CBUseCenterOfMass") = 0 ) Then
					If ( DlgText("XCoorRotAx") = "" ) Or ( DlgText("YCoorRotAx") = "" ) Or ( DlgText("ZCoorRotAx") = "" ) Then
						MsgBox "For at least one coordinate of the origin of the rotational axis no values are present. Please enter a value in the corresponding textbox.", vbExclamation
						DialogFunction = True	' There is an error in the settings -> Don't close the dialog box.
					End If
				End If
			ElseIf ( sDlgItem = "CBSpecifyDensity" ) Then		' If density is specify, enable correpsonding textbox
				If ( DlgValue("CBSpecifyDensity") = 0 ) Then
					DlgText   "Density", DlgText("Density")
					DlgEnable "Density", False
				Else
					DlgText   "Density", DlgText("Density")
					DlgEnable "Density", True
				End If
			ElseIf ( sDlgItem = "CBUseCenterOfMass" ) Then
				If ( DlgValue("CBUseCenterOfMass") = 0 ) Then
					DlgText   "XCoorRotAx", DlgText("XCoorRotAx")
					DlgEnable "XCoorRotAx", True
					DlgText   "YCoorRotAx", DlgText("YCoorRotAx")
					DlgEnable "YCoorRotAx", True
					DlgText   "ZCoorRotAx", DlgText("ZCoorRotAx")
					DlgEnable "ZCoorRotAx", True
				Else
					DlgText   "XCoorRotAx", DlgText("XCoorRotAx")
					DlgEnable "XCoorRotAx", False
					DlgText   "YCoorRotAx", DlgText("YCoorRotAx")
					DlgEnable "YCoorRotAx", False
					DlgText   "ZCoorRotAx", DlgText("ZCoorRotAx")
					DlgEnable "ZCoorRotAx", False
				End If
			End If
		Case 3	' TextBox or ComboBox text changed
			Dim dValueSet     As Double		' double value of string entered in textbox
			Dim sValueSet     As String		' string as entered in textbox
			If ( sDlgItem = "Density" ) Or ( sDlgItem = "XCoorRotAx" ) Or ( sDlgItem = "YCoorRotAx" ) Or ( sDlgItem = "ZCoorRotAx" ) Then
				' Check once to to see whether the entry is in fact a valid numerical value. If not, set value to empty.
				sValueSet   = DlgText(sDlgItem)
				On Error Resume Next
					evaluate(sValueSet)
				If Err.Number <> 0 Then
					DlgText sDlgItem, ""
				End If
			End If
	End Select

End Function



Function Evaluate0DResult(resultID As String, indexMultiple As Long, templateName As String, returnMultipleName As String) As Object
	' 0D result as function value...
	Dim oMomentOfInertia As Object
	Set oMomentOfInertia = Result0D("")

	If indexMultiple = 1 Then
		' Perform all relevant calcuations and save results to global variables
		' Retrieve script settings
		Dim lSpecifiedDensity As Long
		Dim dDensity          As Double
		Dim dOriginXCoor      As Double
		Dim dOriginYCoor      As Double
		Dim dOriginZCoor      As Double
		Dim lUseCenterOfMass  As Long
		Dim dSolidMass    	  As Double	' mass of individual solid
		Dim dSolidVolume      As Double	' volume of individual solid
		Dim dCoMXCoor         As Double	' center of mass x-coordinate (individual solid)
		Dim dCoMYCoor         As Double	' center of mass y-coordinate (individual solid)
		Dim dCoMZCoor         As Double	' center of mass z-coordinate (individual solid)
		Dim sSolidName        As String	' name of individual solid
		Dim dSolidDensity     As Double	' density of individual solid

		iGlobalNumSolids_CST = CInt(GetScriptSetting("SSnSolids", "0"))
		If (iGlobalNumSolids_CST > 0) Then
			ReDim sGlobalSolidArray_CST(iGlobalNumSolids_CST-1)
			ReDim dGlobalSolidMassArray(iGlobalNumSolids_CST-1)
			ReDim dGlobalSolidCoMArrayX(iGlobalNumSolids_CST-1)
			ReDim dGlobalSolidCoMArrayY(iGlobalNumSolids_CST-1)
			ReDim dGlobalSolidCoMArrayZ(iGlobalNumSolids_CST-1)
			For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST
				sGlobalSolidArray_CST(iGlobalSolid_CST-1) = GetScriptSetting("SSSolid" + CStr(iGlobalSolid_CST),"")
			Next
		End If

		lSpecifiedDensity = CLng(GetScriptSetting("SSCBSpecifyDensity", "0"))
		If lSpecifiedDensity <> 0 Then
				dDensity = CDbl(GetScriptSetting("SSDensity", "7850.0"))
		End If

		lUseCenterOfMass = CLng(GetScriptSetting("SSCenterOfMass", "0"))
		dOriginXCoor     = 0.0
		dOriginYCoor     = 0.0
		dOriginZCoor     = 0.0

		' Determine mass and center of mass
		dCenterOfMassX = 0.0
		dCenterOfMassY = 0.0
		dCenterOfMassZ = 0.0
		dOverallMass   = 0.0

		For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST STEP 1
			sSolidName = sGlobalSolidArray_CST(iGlobalSolid_CST-1)
			' Get mass
			If lSpecifiedDensity <> 0 Then
				dSolidVolume = Solid.GetVolume(sSolidName)
				dSolidVolume = dSolidVolume * Units.GetGeometryUnitToSI()^3
				dSolidMass   = dDensity * dSolidVolume
			Else
				dSolidMass   = Solid.GetMass(sSolidName)
			End If
			dGlobalSolidMassArray(iGlobalSolid_CST-1) = dSolidMass
			dOverallMass = dOverallMass + dSolidMass
			' Get coordinates of center of mass
			Solid.GetVolumeCenter(sSolidName, dCoMXCoor, dCoMYCoor, dCoMZCoor)
			dGlobalSolidCoMArrayX(iGlobalSolid_CST-1) = dCoMXCoor * Units.GetGeometryUnitToSI()
			dGlobalSolidCoMArrayY(iGlobalSolid_CST-1) = dCoMYCoor * Units.GetGeometryUnitToSI()
			dGlobalSolidCoMArrayZ(iGlobalSolid_CST-1) = dCoMZCoor * Units.GetGeometryUnitToSI()
			dCenterOfMassX = dCenterOfMassX + dGlobalSolidCoMArrayX(iGlobalSolid_CST-1) * dSolidMass
			dCenterOfMassY = dCenterOfMassY + dGlobalSolidCoMArrayY(iGlobalSolid_CST-1) * dSolidMass
			dCenterOfMassZ = dCenterOfMassZ + dGlobalSolidCoMArrayZ(iGlobalSolid_CST-1) * dSolidMass
		Next iGlobalSolid_CST
		dCenterOfMassX = dCenterOfMassX / dOverallMass
		dCenterOfMassY = dCenterOfMassY / dOverallMass
		dCenterOfMassZ = dCenterOfMassZ / dOverallMass

		If lUseCenterOfMass <> 0 Then
			dOriginXCoor = dCenterOfMassX
			dOriginYCoor = dCenterOfMassY
			dOriginZCoor = dCenterOfMassZ
		Else
			dOriginXCoor = CDbl(GetScriptSetting("SSOriginXCoor", "0.0"))
			dOriginYCoor = CDbl(GetScriptSetting("SSOriginYCoor", "0.0"))
			dOriginZCoor = CDbl(GetScriptSetting("SSOriginZCoor", "0.0"))
			dOriginXCoor = dOriginXCoor * Units.GetGeometryUnitToSI()
			dOriginYCoor = dOriginYCoor * Units.GetGeometryUnitToSI()
			dOriginZCoor = dOriginZCoor * Units.GetGeometryUnitToSI()
		End If

		Dim dInertiaX      As Double	' moment of inertia of individual solid along principal axis 1
		Dim dInertiaY      As Double	' moment of inertia of individual solid along principal axis 2
		Dim dInertiaZ      As Double	' moment of inertia of individual solid along principal axis 3
		Dim dInertiaX1     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaX2     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaX3     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaY1     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaY2     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaY3     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaZ1     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaZ2     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaZ3     As Double	' entry in matrix describing transformation to system of principle axes
		Dim dInertiaX1Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaX2Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaX3Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaY1Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaY2Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaY3Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaZ1Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaZ2Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaZ3Inv  As Double	' entry in inverse matrix describing transformation to system of principle axes
		Dim dInertiaDet    As Double	' Determinant of matrix
		Dim dInertiaXXGlob As Double	' entry in inertia tensor in global system
		Dim dInertiaXYGlob As Double	' entry in inertia tensor in global system
		Dim dInertiaXZGlob As Double	' entry in inertia tensor in global system
		Dim dInertiaYYGlob As Double	' entry in inertia tensor in global system
		Dim dInertiaYZGlob As Double	' entry in inertia tensor in global system
		Dim dInertiaZZGlob As Double	' entry in inertia tensor in global system
		Dim sMaterialName  As String	' material of individual solid
		Dim dRotAxDistSqXY As Double	' distance of center of mass to origin of rotational axis
		Dim dRotAxDistSqXZ As Double	' distance of center of mass to origin of rotational axis
		Dim dRotAxDistSqYZ As Double	' distance of center of mass to origin of rotational axis
		Dim dRotAxDistPtXY As Double	' distance of center of mass to origin of rotational axis
		Dim dRotAxDistPtXZ As Double	' distance of center of mass to origin of rotational axis
		Dim dRotAxDistPtYZ As Double	' distance of center of mass to origin of rotational axis

		' Initialize moment of inertia
		dMomentum_xx = 0.0
		dMomentum_xy = 0.0
		dMomentum_xz = 0.0
		dMomentum_yy = 0.0
		dMomentum_yz = 0.0
		dMomentum_zz = 0.0

		' Go through list of solids, get density, calculate moment of inertia
		For iGlobalSolid_CST = 1 To iGlobalNumSolids_CST STEP 1
			sSolidName = sGlobalSolidArray_CST(iGlobalSolid_CST-1)
			' Get Density
			If lSpecifiedDensity <> 0 Then
				dSolidDensity = dDensity
			Else
				sMaterialName = Solid.GetMaterialNameForShape(sSolidName)
				Material.GetRho(sMaterialName, dSolidDensity)
			End If
			' Get integral / inertia in project units
			Solid.GetInertia(sSolidName, 0, dInertiaX, dInertiaX1, dInertiaX2, dInertiaX3)
			Solid.GetInertia(sSolidName, 1, dInertiaY, dInertiaY1, dInertiaY2, dInertiaY3)
			Solid.GetInertia(sSolidName, 2, dInertiaZ, dInertiaZ1, dInertiaZ2, dInertiaZ3)
			' Take density into account and convert to SI units
			dInertiaX = dInertiaX * dSolidDensity * Units.GetGeometryUnitToSI()^5
			dInertiaY = dInertiaY * dSolidDensity * Units.GetGeometryUnitToSI()^5
			dInertiaZ = dInertiaZ * dSolidDensity * Units.GetGeometryUnitToSI()^5
			' Convert to global coordinate system
			' Determine inverse...
			dInertiaDet    = (dInertiaX1 * dInertiaY2 * dInertiaZ3 + dInertiaX2 * dInertiaY3 * dInertiaZ1 + dInertiaX3 * dInertiaY1 * dInertiaZ2) - (dInertiaX1 * dInertiaY3 * dInertiaZ2 + dInertiaX2 * dInertiaY1 * dInertiaZ3 + dInertiaX3 * dInertiaY2 * dInertiaZ1)
			dInertiaX1Inv  = dInertiaY2 * dInertiaZ3 - dInertiaZ2 * dInertiaY3
			dInertiaX2Inv  = dInertiaZ2 * dInertiaX3 - dInertiaX2 * dInertiaZ3
			dInertiaX3Inv  = dInertiaX2 * dInertiaY3 - dInertiaY2 * dInertiaX3
			dInertiaY1Inv  = dInertiaZ1 * dInertiaY3 - dInertiaY1 * dInertiaZ3
			dInertiaY2Inv  = dInertiaX1 * dInertiaZ3 - dInertiaZ1 * dInertiaX3
			dInertiaY3Inv  = dInertiaY1 * dInertiaX3 - dInertiaX1 * dInertiaY3
			dInertiaZ1Inv  = dInertiaY1 * dInertiaZ2 - dInertiaZ1 * dInertiaY2
			dInertiaZ2Inv  = dInertiaZ1 * dInertiaX2 - dInertiaX1 * dInertiaZ2
			dInertiaZ3Inv  = dInertiaX1 * dInertiaY2 - dInertiaY1 * dInertiaX2
			dInertiaX1Inv  = dInertiaX1Inv / dInertiaDet
			dInertiaX2Inv  = dInertiaX2Inv / dInertiaDet
			dInertiaX3Inv  = dInertiaX3Inv / dInertiaDet
			dInertiaY1Inv  = dInertiaY1Inv / dInertiaDet
			dInertiaY2Inv  = dInertiaY2Inv / dInertiaDet
			dInertiaY3Inv  = dInertiaY3Inv / dInertiaDet
			dInertiaZ1Inv  = dInertiaZ1Inv / dInertiaDet
			dInertiaZ2Inv  = dInertiaZ2Inv / dInertiaDet
			dInertiaZ3Inv  = dInertiaZ3Inv / dInertiaDet
			' Convert to global system (still anchored in CoM of individual solid)
			dInertiaXXGlob = dInertiaX * dInertiaX1 * dInertiaX1Inv + dInertiaY * dInertiaY1 * dInertiaX2Inv + dInertiaZ * dInertiaZ1 * dInertiaX3Inv
			dInertiaXYGlob = dInertiaX * dInertiaX2 * dInertiaX1Inv + dInertiaY * dInertiaY2 * dInertiaX2Inv + dInertiaZ * dInertiaZ2 * dInertiaX3Inv
			dInertiaXZGlob = dInertiaX * dInertiaX3 * dInertiaX1Inv + dInertiaY * dInertiaY3 * dInertiaX2Inv + dInertiaZ * dInertiaZ3 * dInertiaX3Inv
			dInertiaYYGlob = dInertiaX * dInertiaX2 * dInertiaY1Inv + dInertiaY * dInertiaY2 * dInertiaY2Inv + dInertiaZ * dInertiaZ2 * dInertiaY3Inv
			dInertiaYZGlob = dInertiaX * dInertiaX2 * dInertiaZ1Inv + dInertiaY * dInertiaY2 * dInertiaZ2Inv + dInertiaZ * dInertiaZ2 * dInertiaZ3Inv
			dInertiaZZGlob = dInertiaX * dInertiaX3 * dInertiaZ1Inv + dInertiaY * dInertiaY3 * dInertiaZ2Inv + dInertiaZ * dInertiaZ3 * dInertiaZ3Inv
			' Default of GetInertia is for center of mass - get matrix entries for origin of rotational axis
			dSolidMass     = dGlobalSolidMassArray(iGlobalSolid_CST-1)
			dCoMXCoor      = dGlobalSolidCoMArrayX(iGlobalSolid_CST-1)
			dCoMYCoor      = dGlobalSolidCoMArrayY(iGlobalSolid_CST-1)
			dCoMZCoor      = dGlobalSolidCoMArrayZ(iGlobalSolid_CST-1)
			dRotAxDistSqXY = (dCoMXCoor - dOriginXCoor)^2 + (dCoMYCoor - dOriginYCoor)^2
			dRotAxDistSqXZ = (dCoMXCoor - dOriginXCoor)^2 + (dCoMZCoor - dOriginZCoor)^2
			dRotAxDistSqYZ = (dCoMYCoor - dOriginYCoor)^2 + (dCoMZCoor - dOriginZCoor)^2
			dRotAxDistPtXY = (dOriginXCoor - dCoMXCoor) * (dOriginYCoor - dCoMYCoor)
			dRotAxDistPtXZ = (dOriginXCoor - dCoMXCoor) * (dOriginZCoor - dCoMZCoor)
			dRotAxDistPtYZ = (dOriginYCoor - dCoMYCoor) * (dOriginZCoor - dCoMZCoor)
			dInertiaXXGlob = dInertiaXXGlob + dSolidMass * dRotAxDistSqYZ
			dInertiaXYGlob = dInertiaXYGlob - dSolidMass * dRotAxDistPtXY
			dInertiaXZGlob = dInertiaXZGlob - dSolidMass * dRotAxDistPtXZ
			dInertiaYYGlob = dInertiaYYGlob + dSolidMass * dRotAxDistSqXZ
			dInertiaYZGlob = dInertiaYZGlob - dSolidMass * dRotAxDistPtYZ
			dInertiaZZGlob = dInertiaZZGlob + dSolidMass * dRotAxDistSqXY
			' Add to sum...
			dMomentum_xx = dMomentum_xx + dInertiaXXGlob
			dMomentum_xy = dMomentum_xy + dInertiaXYGlob
			dMomentum_xz = dMomentum_xz + dInertiaXZGlob
			dMomentum_yy = dMomentum_yy + dInertiaYYGlob
			dMomentum_yz = dMomentum_yz + dInertiaYZGlob
			dMomentum_zz = dMomentum_zz + dInertiaZZGlob
		Next iGlobalSolid_CST
		returnMultipleName = templateName & "\J_xx (SI)"
		oMomentOfInertia.SetData(dMomentum_xx)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 2 Then
		returnMultipleName = templateName & "\J_xy (SI)"
		oMomentOfInertia.SetData(dMomentum_xy)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 3 Then
		returnMultipleName = templateName & "\J_xz (SI)"
		oMomentOfInertia.SetData(dMomentum_xz)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 4 Then
		returnMultipleName = templateName & "\J_yy (SI)"
		oMomentOfInertia.SetData(dMomentum_yy)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 5 Then
		returnMultipleName = templateName & "\J_yz (SI)"
		oMomentOfInertia.SetData(dMomentum_yz)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 6 Then
		returnMultipleName = templateName & "\J_zz (SI)"
		oMomentOfInertia.SetData(dMomentum_zz)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 7 Then
		returnMultipleName = templateName & "\CoM_x (SI)"
		oMomentOfInertia.SetData(dCenterOfMassX)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 8 Then
		returnMultipleName = templateName & "\CoM_y (SI)"
		oMomentOfInertia.SetData(dCenterOfMassY)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 9 Then
		returnMultipleName = templateName & "\CoM_z (SI)"
		oMomentOfInertia.SetData(dCenterOfMassZ)
		Set Evaluate0DResult = oMomentOfInertia
	ElseIf indexMultiple = 10 Then
		returnMultipleName = templateName & "\Mass (SI)"
		oMomentOfInertia.SetData(dOverallMass)
		Set Evaluate0DResult = oMomentOfInertia
	Else
		returnMultipleName = ""
	End If

End Function


' -------------------------------------------------------------------------------------------------
' Main: This function serves as a main program for testing purposes.
'       You need to rename this function to "Main" for debugging the result template.
'
'		PLEASE NOTE that a result template file must not contain a main program for
'       proper execution by the framework. Therefore please ensure to rename this function
'       to e.g. "Main2" before the result template can be used by the framework.
' -------------------------------------------------------------------------------------------------

Sub Main
	' Activate the StoreScriptSetting / GetScriptSetting functionality. Clear the data in order to
	' provide well defined environment for testing.
	ActivateScriptSettings True
	ClearScriptSettings
	DS.ClearScriptSettings

	Dim sResultID As String

	If (Left(GetApplicationName, 2) = "DS") Then
		sResultID = DS.GetLastResultID()
	Else
		sResultID = GetLastResultID()
	End If

	' Now call the define method and check whether it is completed successfully

	If (Define("test", True, False)) Then
		' If the define method is executed properly, call the Evaluate0DResult method
		Dim stmpfile As String, dTemp As Double
		Dim ncount As Long, sTableName As String
		stmpfile = "Test_tmp.txt"
		Dim r0d As Object
		Evaluate0DResult(sResultID, 1, "Test 0D", "")
		ReportInformationToWindow("J_xx (SI): " & CStr(dMomentum_xx))
		ReportInformationToWindow("J_xy (SI): " & CStr(dMomentum_xy))
		ReportInformationToWindow("J_xz (SI): " & CStr(dMomentum_xz))
		ReportInformationToWindow("J_yy (SI): " & CStr(dMomentum_yy))
		ReportInformationToWindow("J_yz (SI): " & CStr(dMomentum_yz))
		ReportInformationToWindow("J_zz (SI): " & CStr(dMomentum_zz))
		ReportInformationToWindow("CoM_x (SI): " & CStr(dCenterOfMassX))
		ReportInformationToWindow("CoM_y (SI): " & CStr(dCenterOfMassY))
		ReportInformationToWindow("CoM_z (SI): " & CStr(dCenterOfMassZ))
		ReportInformationToWindow("Mass (SI): " & CStr(dOverallMass))
	End If

	 'Deactivate the StoreScriptSetting / GetScriptSetting functionality.
	ActivateScriptSettings False

End Sub
