' Create Full Tensor Material, a GUI frontend for the existing VBA commands

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "coordinate_systems.lib"

' ================================================================================================
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' --------------------------------------------------------------------------------------------------------------------------------------
' 23-Feb-2016 ube: Slight change wording of help text to describe term f
' 15-Jul-2015 fsr: Added WCS option for tensor alignment
' 18-Nov-2013 fsr: Fixed a problem with "<Main folder>" selection
' 18-Nov-2013 jma: Two vectors alignment for tensor formula material
' 01-Mar-2013 ube: removed help button, which was inactive
' 03-Aug-2012 fsr: use "eps_r-j*eps_i" instead of "eps_r+j*eps_i"; mcs->mcr; save/load dialog settings; accept parameters; help button
' 02-Feb-2012 fsr: improved 'Cancel' handling
' 09-Sep-2011 fsr: fixed a bug/typo: TensorAlignment(alignX,alignY,alignY) -> TensorAlignment(alignX,alignY,alignZ)
' 26-Oct-2010 fsr: initial version
' --------------------------------------------------------------------------------------------------------------------------------------

Sub Main ()

	Dim sMaterialFolderArray() As String, sMaterialNameArray() As String

	FillWCSArray

	Begin Dialog UserDialog 760,525,"Create tensor formula material",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,7,720,84,"General Settings",.GroupBox3
		Text 40,35,100,14,"Material folder:",.Text12
		DropListBox 150,28,230,192,sMaterialFolderArray(),.MaterialFolderDLB,1
		Text 40,63,100,14,"Material name:",.Text1
		DropListBox 150,56,230,192,sMaterialNameArray(),.MaterialNameDLB,1
		PushButton 390,56,110,21,"Load Settings",.RestoreDialogSettingsPB

		CheckBox 20,95,140,14,"Uniaxial material",.Uniaxial_on

		GroupBox 20,112,720,182,"Tensor formulas for eps_r",.GroupBox1
		Text 130,143,20,14,"-j",.Text3
		Text 610,217,20,14,"-j",.Text4
		Text 370,217,20,14,"-j",.Text5
		Text 130,217,20,14,"-j",.Text6
		Text 610,182,20,14,"-j",.Text7
		Text 370,182,20,14,"-j",.Text8
		Text 130,182,20,14,"-j",.Text9
		Text 610,143,20,14,"-j",.Text10
		Text 370,143,20,14,"-j",.Text11
		TextBox 40,136,90,21,.eps11r
		TextBox 150,136,90,21,.eps11i
		TextBox 280,136,90,21,.eps12r
		TextBox 390,136,90,21,.eps12i
		TextBox 520,136,90,21,.eps13r
		TextBox 630,136,90,21,.eps13i
		TextBox 40,175,90,21,.eps21r
		TextBox 150,175,90,21,.eps21i
		TextBox 280,175,90,21,.eps22r
		TextBox 390,175,90,21,.eps22i
		TextBox 520,175,90,21,.eps23r
		TextBox 630,175,90,21,.eps23i
		TextBox 40,210,90,21,.eps31r
		TextBox 150,210,90,21,.eps31i
		TextBox 280,210,90,21,.eps32r
		TextBox 390,210,90,21,.eps32i
		TextBox 520,210,90,21,.eps33r
		TextBox 630,210,90,21,.eps33i

		Text 40,238,120,14,"Alignment vectors:",.Text13_U
		Text 60,266,20,14,"U:",.Text15
		TextBox 80,259,30,21,.alignXeps_U
		TextBox 120,259,30,21,.alignYeps_U
		TextBox 160,259,30,21,.alignZeps_U

		Text 210,266,20,14,"V:",.Text13_V
		TextBox 230,259,30,21,.alignXeps_V
		TextBox 270,259,30,21,.alignYeps_V
		TextBox 310,259,30,21,.alignZeps_V

		Text 360,266,20,14,"W:",.Text13
		TextBox 390,259,30,21,.alignXeps
		TextBox 430,259,30,21,.alignYeps
		TextBox 470,259,30,21,.alignZeps

		Text 520,266,80,14,"From WCS:",.Text16
		DropListBox 610,259,110,170,aWCS(),.WCSEpsDLB

		GroupBox 20,301,720,182,"Tensor formulas for mu_r",.GroupBox2
		Text 130,336,20,14,"-j",.Text23
		Text 610,406,20,14,"-j",.Text24
		Text 370,406,20,14,"-j",.Text25
		Text 130,406,20,14,"-j",.Text26
		Text 610,371,20,14,"-j",.Text27
		Text 370,371,20,14,"-j",.Text28
		Text 130,371,20,14,"-j",.Text29
		Text 610,336,20,14,"-j",.Text210
		Text 370,336,20,14,"-j",.Text211
		TextBox 40,329,90,21,.mu11r
		TextBox 150,329,90,21,.mu11i
		TextBox 280,329,90,21,.mu12r
		TextBox 390,329,90,21,.mu12i
		TextBox 520,329,90,21,.mu13r
		TextBox 630,329,90,21,.mu13i
		TextBox 40,364,90,21,.mu21r
		TextBox 150,364,90,21,.mu21i
		TextBox 280,364,90,21,.mu22r
		TextBox 390,364,90,21,.mu22i
		TextBox 520,364,90,21,.mu23r
		TextBox 630,364,90,21,.mu23i
		TextBox 40,399,90,21,.mu31r
		TextBox 150,399,90,21,.mu31i
		TextBox 280,399,90,21,.mu32r
		TextBox 390,399,90,21,.mu32i
		TextBox 520,399,90,21,.mu33r
		TextBox 630,399,90,21,.mu33i

		Text 40,427,140,14,"Alignment vectors:",.Text14_U
		Text 60,455,20,14,"U:",.Text17
		TextBox 80,448,30,21,.alignXmu_U
		TextBox 120,448,30,21,.alignYmu_U
		TextBox 160,448,30,21,.alignZmu_U

		Text 210,455,20,14,"V:",.Text14_V
		TextBox 230,448,30,21,.alignXmu_V
		TextBox 270,448,30,21,.alignYmu_V
		TextBox 310,448,30,21,.alignZmu_V

		Text 370,455,20,14,"W:",.Text14
		TextBox 390,448,30,21,.alignXmu
		TextBox 430,448,30,21,.alignYmu
		TextBox 470,448,30,21,.alignZmu

		Text 520,455,80,14,"From WCS:",.Text18
		DropListBox 610,448,110,170,aWCS(),.WCSMuDLB

		Text 30,490,500,28,"Please use 'f' to describe frequency dependency. f will be replaced by the frequency in Hz, so entered formula is independent of project frequency unit.",.Text2
		OKButton 550,490,90,21
		CancelButton 650,490,90,21

		' PushButton 650,476,90,21,"Help",.HelpPB

	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim sMaterialFolderArray() As String, sMaterialNameArray() As String, MaterialFoldersAndNames As Variant

		Select Case Action%
			Case 1 ' Dialog box initialization

			MaterialFoldersAndNames = GetMaterialFoldersAndNames_LIB()
			sMaterialFolderArray = MaterialFoldersAndNames(0)
			sMaterialNameArray = MaterialFoldersAndNames(1)
			DlgListBoxArray("MaterialFolderDLB", sMaterialFolderArray)
			DlgListBoxArray("MaterialNameDLB", sMaterialNameArray)

			' Initialize dialog with some generic values
			DlgText("MaterialNameDLB", "FullTensorMaterial")
			DlgText("MaterialFolderDLB", "")

			DlgValue("Uniaxial_on", True)

			DlgText("eps11r", "1")
			DlgText("eps11i", "0")
			DlgText("eps12r", "0")
			DlgText("eps12i", "0")
			DlgText("eps13r", "0")
			DlgText("eps13i", "0")
			DlgText("eps21r", "0")
			DlgText("eps21i", "0")
			DlgText("eps22r", "1")
			DlgText("eps22i", "0")
			DlgText("eps23r", "0")
			DlgText("eps23i", "0")
			DlgText("eps31r", "0")
			DlgText("eps31i", "0")
			DlgText("eps32r", "0")
			DlgText("eps32i", "0")
			DlgText("eps33r", "1")
			DlgText("eps33i", "0")

			DlgText("mu11r", "1")
			DlgText("mu11i", "0")
			DlgText("mu12r", "0")
			DlgText("mu12i", "0")
			DlgText("mu13r", "0")
			DlgText("mu13i", "0")
			DlgText("mu21r", "0")
			DlgText("mu21i", "0")
			DlgText("mu22r", "1")
			DlgText("mu22i", "0")
			DlgText("mu23r", "0")
			DlgText("mu23i", "0")
			DlgText("mu31r", "0")
			DlgText("mu31i", "0")
			DlgText("mu32r", "0")
			DlgText("mu32i", "0")
			DlgText("mu33r", "1")
			DlgText("mu33i", "0")

			DlgText("alignXeps_U", "1")
			DlgText("alignYeps_U", "0")
			DlgText("alignZeps_U", "0")

			DlgText("alignXeps_V", "0")
			DlgText("alignYeps_V", "1")
			DlgText("alignZeps_V", "0")

			DlgText("alignXeps", "0")
			DlgText("alignYeps", "0")
			DlgText("alignZeps", "1")

			DlgText("alignXmu_U", "1")
			DlgText("alignYmu_U", "0")
			DlgText("alignZmu_U", "0")

			DlgText("alignXmu_V", "0")
			DlgText("alignYmu_V", "1")
			DlgText("alignZmu_V", "0")


			DlgText("alignXmu", "0")
			DlgText("alignYmu", "0")
			DlgText("alignZmu", "1")

			' Load dialog settings from previous macro run
			ReStoreAllDialogSettings_LIB(sIniFileName)
			DisableFields()
		Case 2 ' Value changing or button pressed
			Select Case DlgItem$
				Case "Cancel"
					Exit All
				Case "HelpPB"
					StartHelp("common_preloadedmacro_materials_full_tensor")
					DialogFunc = True
				Case "OK"
					Dim sHistoryString As String
					sHistoryString = GenerateHistoryStringForTensorMaterial()
					If (sHistoryString <> "ERROR") Then
						AddToHistory("add material: " + DlgText("MaterialFolderDLB") + "\" + DlgText("MaterialNameDLB"), sHistoryString)
					End If
					' Store dialog settings in ini file
					iniFileName = DlgText("MaterialFolderDLB")+"_"+DlgText("MaterialNameDLB")+".ini"
					iniFileName = Replace(iniFileName, "\", "")
					iniFileName = Replace(iniFileName, "/", "")
					iniFileName = Replace(iniFileName, "<Main folder>", "")
					StoreAllDialogSettings_LIB(GetProjectPath("Model3D")+iniFileName, "MaterialFolderDLB,MaterialNameDLB", "", "Dialog settings for 'Create Full Tensor Material' macro") ' exclude material name and folder
	    		Case "RestoreDialogSettingsPB"
	  				iniFileName = DlgText("MaterialFolderDLB")+"_"+DlgText("MaterialNameDLB")+".ini"
					iniFileName = Replace(iniFileName, "\", "")
					iniFileName = Replace(iniFileName, "/", "")
					iniFileName = Replace(iniFileName, "<Main folder>", "")
					'ReportInformationToWindow(iniFileName)
					ReStoreAllDialogSettings_LIB(GetProjectPath("Model3D")+iniFileName, "MaterialFolderDLB,MaterialNameDLB", "") ' exclude material name and folder
					DisableFields()
	    			DialogFunc = True
	    		Case "Uniaxial_on", "WCSEpsDLB", "WCSMuDLB"
	    			DisableFields()
			End Select
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function

Private Function DisableFields()

	If (DlgValue("Uniaxial_on") = 1) Then
		DlgEnable ("alignXeps_U",False)
		DlgEnable ("alignYeps_U",False)
		DlgEnable ("alignZeps_U",False)

		DlgEnable ("alignXeps_V",False)
		DlgEnable ("alignYeps_V",False)
		DlgEnable ("alignZeps_V",False)

		' only enable if global WCS is selected
		DlgEnable ("alignXeps",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignYeps",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignZeps",DlgValue("WCSEpsDLB") = 0)

		DlgEnable ("alignXmu_U",False)
		DlgEnable ("alignYmu_U",False)
		DlgEnable ("alignZmu_U",False)

		DlgEnable ("alignXmu_V",False)
		DlgEnable ("alignYmu_V",False)
		DlgEnable ("alignZmu_V",False)

		' only enable if global WCS is selected
		DlgEnable ("alignXmu",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignYmu",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignZmu",DlgValue("WCSMuDLB") = 0)
	Else
		' only enable if global WCS is selected
		DlgEnable ("alignXeps_U",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignYeps_U",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignZeps_U",DlgValue("WCSEpsDLB") = 0)

		' only enable if global WCS is selected
		DlgEnable ("alignXeps_V",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignYeps_V",DlgValue("WCSEpsDLB") = 0)
		DlgEnable ("alignZeps_V",DlgValue("WCSEpsDLB") = 0)

		DlgEnable ("alignXeps",False)
		DlgEnable ("alignYeps",False)
		DlgEnable ("alignZeps",False)

		' only enable if global WCS is selected
		DlgEnable ("alignXmu_U",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignYmu_U",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignZmu_U",DlgValue("WCSMuDLB") = 0)

		' only enable if global WCS is selected
		DlgEnable ("alignXmu_V",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignYmu_V",DlgValue("WCSMuDLB") = 0)
		DlgEnable ("alignZmu_V",DlgValue("WCSMuDLB") = 0)

		DlgEnable ("alignXmu",False)
		DlgEnable ("alignYmu",False)
		DlgEnable ("alignZmu",False)
	End If
End Function

Function GenerateHistoryStringForTensorMaterial() As String

	' Variable declaration
	Dim sHistoryString As String

	sHistoryString = ""
	' x/y/z vectors to describe the u/v/z vectors of the WCS
	sHistoryString = sHistoryString + "Dim dWCSUX As Double, dWCSUY As Double, dWCSUZ As Double" + vbNewLine
	sHistoryString = sHistoryString + "Dim dWCSVX As Double, dWCSVY As Double, dWCSVZ As Double" + vbNewLine
	sHistoryString = sHistoryString + "Dim dWCSWX As Double, dWCSWY As Double, dWCSWZ As Double" + vbNewLine

	' Read values from dialog, store in variables

	sHistoryString = sHistoryString + "With Material" + vbNewLine
	sHistoryString = sHistoryString + "	.Reset" + vbNewLine
	sHistoryString = sHistoryString + "	.Name " + Chr(34) + DlgText("MaterialNameDLB") + Chr(34) + vbNewLine
	If DlgText("MaterialFolderDLB") = "<Main folder>" Then
		sHistoryString = sHistoryString + "	.Folder " + Chr(34) + Chr(34) + vbNewLine
	Else
		sHistoryString = sHistoryString + "	.Folder " + Chr(34) + DlgText("MaterialFolderDLB") + Chr(34) + vbNewLine
	End If
	sHistoryString = sHistoryString + "	.Type ("+Chr(34) + "Tensor formula" + Chr(34) +")" + vbNewLine
	sHistoryString = sHistoryString + "	.Colour ("+Chr(34) + "0"+Chr(34) + ", "+Chr(34) + "1"+Chr(34) + ", "+Chr(34) + "1"+Chr(34) + ")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaFor("+Chr(34) + "epsilon_r"+Chr(34) + ")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,0,"+Chr(34)+DlgText("eps11r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,0,"+Chr(34)+DlgText("eps11i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,1,"+Chr(34)+DlgText("eps12r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,1,"+Chr(34)+DlgText("eps12i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,2,"+Chr(34)+DlgText("eps13r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,2,"+Chr(34)+DlgText("eps13i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,0,"+Chr(34)+DlgText("eps21r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,0,"+Chr(34)+DlgText("eps21i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,1,"+Chr(34)+DlgText("eps22r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,1,"+Chr(34)+DlgText("eps22i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,2,"+Chr(34)+DlgText("eps23r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,2,"+Chr(34)+DlgText("eps23i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,0,"+Chr(34)+DlgText("eps31r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,0,"+Chr(34)+DlgText("eps31i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,1,"+Chr(34)+DlgText("eps32r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,1,"+Chr(34)+DlgText("eps32i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,2,"+Chr(34)+DlgText("eps33r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,2,"+Chr(34)+DlgText("eps33i")+Chr(34)+")" + vbNewLine

	If (DlgValue("Uniaxial_on") = 1) Then
		If (DlgValue("WCSEpsDLB") = 0) Then
			sHistoryString = sHistoryString + "	.TensorAlignment("+Chr(34)+DlgText("alignXeps")+Chr(34)+","+Chr(34)+DlgText("alignYeps")+Chr(34)+","+Chr(34)+DlgText("alignZeps")+Chr(34)+")" + vbNewLine
		Else
			sHistoryString = sHistoryString + "	WCS.GetNormal(" + Chr(34) +  Replace(DlgText("WCSEpsDLB"), "local: ", "") + Chr(34) + ", dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
			sHistoryString = sHistoryString + "	.TensorAlignment(dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
		End If
	Else
		If (DlgValue("WCSEpsDLB") = 0) Then
			sHistoryString = sHistoryString + "	.TensorAlignment2("+Chr(34)+DlgText("alignXeps_U")+Chr(34)+","+Chr(34)+DlgText("alignYeps_U")+Chr(34)+","+Chr(34)+DlgText("alignZeps_U")+Chr(34)+","+Chr(34)+DlgText("alignXeps_V")+Chr(34)+","+Chr(34)+DlgText("alignYeps_V")+Chr(34)+","+Chr(34)+DlgText("alignZeps_V")+Chr(34)+")" + vbNewLine
		Else
			sHistoryString = sHistoryString + "	WCS.GetNormal(" + Chr(34) +  Replace(DlgText("WCSEpsDLB"), "local: ", "") + Chr(34) + ", dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
			sHistoryString = sHistoryString + "	WCS.GetUVector(" + Chr(34) +  Replace(DlgText("WCSEpsDLB"), "local: ", "") + Chr(34) + ", dWCSUX, dWCSUY, dWCSUZ)" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVX = dWCSWY*dWCSUZ - dWCSWZ*dWCSUY" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVY = dWCSWZ*dWCSUX - dWCSWX*dWCSUZ" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVZ = dWCSWX*dWCSUY - dWCSWY*dWCSUX" + vbNewLine
			sHistoryString = sHistoryString + "	.TensorAlignment2(dWCSUX, dWCSUY, dWCSUZ, dWCSVX, dWCSVY, dWCSVZ)" + vbNewLine
		End If
	End If

	sHistoryString = sHistoryString + "	.TensorFormulaFor("+Chr(34) + "mu_r"+Chr(34) + ")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,0,"+Chr(34)+DlgText("mu11r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,0,"+Chr(34)+DlgText("mu11i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,1,"+Chr(34)+DlgText("mu12r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,1,"+Chr(34)+DlgText("mu12i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(0,2,"+Chr(34)+DlgText("mu13r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(0,2,"+Chr(34)+DlgText("mu13i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,0,"+Chr(34)+DlgText("mu21r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,0,"+Chr(34)+DlgText("mu21i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,1,"+Chr(34)+DlgText("mu22r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,1,"+Chr(34)+DlgText("mu22i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(1,2,"+Chr(34)+DlgText("mu23r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(1,2,"+Chr(34)+DlgText("mu23i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,0,"+Chr(34)+DlgText("mu31r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,0,"+Chr(34)+DlgText("mu31i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,1,"+Chr(34)+DlgText("mu32r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,1,"+Chr(34)+DlgText("mu32i")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaReal(2,2,"+Chr(34)+DlgText("mu33r")+Chr(34)+")" + vbNewLine
	sHistoryString = sHistoryString + "	.TensorFormulaImag(2,2,"+Chr(34)+DlgText("mu33i")+Chr(34)+")" + vbNewLine

	If (DlgValue("Uniaxial_on") = 1) Then
		If (DlgValue("WCSMuDLB") = 0) Then
			sHistoryString = sHistoryString + "	.TensorAlignment("+Chr(34)+DlgText("alignXmu")+Chr(34)+","+Chr(34)+DlgText("alignYmu")+Chr(34)+","+Chr(34)+DlgText("alignZmu")+Chr(34)+")" + vbNewLine
		Else
			sHistoryString = sHistoryString + "	WCS.GetNormal(" + Chr(34) + Replace(DlgText("WCSMuDLB"), "local: ", "") + Chr(34) + ", dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
			sHistoryString = sHistoryString + "	.TensorAlignment(dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
		End If
	Else
		If (DlgValue("WCSMuDLB") = 0) Then
			sHistoryString = sHistoryString + "	.TensorAlignment2("+Chr(34)+DlgText("alignXmu_U")+Chr(34)+","+Chr(34)+DlgText("alignYmu_U")+Chr(34)+","+Chr(34)+DlgText("alignZmu_U")+Chr(34)+","+Chr(34)+DlgText("alignXmu_V")+Chr(34)+","+Chr(34)+DlgText("alignYmu_V")+Chr(34)+","+Chr(34)+DlgText("alignZmu_V")+Chr(34)+")" + vbNewLine
		Else
			sHistoryString = sHistoryString + "	WCS.GetNormal(" + Chr(34) +  Replace(DlgText("WCSMuDLB"), "local: ", "") + Chr(34) + ", dWCSWX, dWCSWY, dWCSWZ)" + vbNewLine
			sHistoryString = sHistoryString + "	WCS.GetUVector(" + Chr(34) +  Replace(DlgText("WCSMuDLB"), "local: ", "") + Chr(34) + ", dWCSUX, dWCSUY, dWCSUZ)" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVX = dWCSWY*dWCSUZ - dWCSWZ*dWCSUY" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVY = dWCSWZ*dWCSUX - dWCSWX*dWCSUZ" + vbNewLine
			sHistoryString = sHistoryString + "	dWCSVZ = dWCSWX*dWCSUY - dWCSWY*dWCSUX" + vbNewLine
			sHistoryString = sHistoryString + "	.TensorAlignment2(dWCSUX, dWCSUY, dWCSUZ, dWCSVX, dWCSVY, dWCSVZ)" + vbNewLine
		End If
	End If

	sHistoryString = sHistoryString + "	.Create" + vbNewLine
	sHistoryString = sHistoryString + "End With" + vbNewLine

	GenerateHistoryStringForTensorMaterial = sHistoryString

End Function
