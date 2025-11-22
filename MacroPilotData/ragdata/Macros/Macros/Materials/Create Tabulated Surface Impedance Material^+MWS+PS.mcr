'
' This macro defines a broadband 1D-surface impedance model for multilayer conductors or superconductors including surface roughness.
'
'
' Note: The standard "lossy metal" within CST MWS only produces accurate results
' for skindepths smaller than the metallic thickness. Especially at DC and low frequencies,
' the tabulated surface impedance will produce more accurate results.
'
' The model used here is exact for skindepth >> metal thickness and for metal thickness >> skindepth.
' For all other cases the model is a good approximation.
' The model realizes the "even" case where current flow on both sides of the metal is in the same direction. This makes sense for
' the traces of e.g. microstrip lines, which were the original motivation for this macro.
'
'--------------------------------------------------------------------------------------------------------------------------------------------------
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
'--------------------------------------------------------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 07-Mar-2022 ech: Corrected computation with Huray model and a problem with the frequency scale units. Set the default error limit for fitting to 0.01
' 11-Jan-2021 fsr: Corrected problem with second skin depth calculation
' 06-Jul-2018 fsr: Added correction factor to conductor thickness to accound for skin effect
' 12-Apr-2017 fsr: Fixed DC value for 2 layers (single side current, only half DC impedance); added DC correction factor for 3 layer
' 21-Jan-2014 fsr: Fixed a problem with period/comma in RMS and kappa values for some non-US locales
' 18-Nov-2013 fsr: Fixed a problem with "<Main folder>" selection
' 07-Jun-2012 fsr: Implemented Huray model (1st order for now) for surface roughness; joined GUIs for "Layer" and "Superconductor" configurations
' 04-Jun-2012 fsr: Dialog settings are now stored (restored) when a material is created (selected from drop down list)
' 14-Feb-2012 ube: help button activated and online help included
' 09-Feb-2012 fsr: Several minor improvements
' 03-Feb-2012 ube: "kappa" renamed to "conductivity" in dialog
' 25-Jan-2012 fsr: Clean up GlobalDataValues Mix_A...F at start to avoid conflicts with 'Mix Template Results', bugfixes and improvements;
'					material names and folders can now be selected from dropdown lists
' 13-Jan-2012 ube: help button commented out, until online help is ready
' 12-Jan-2012 fsr: Minor bugfixes and improvements, added help button
' 16-Dec-2011 fsr: Experimental: included to enforce Kramers-Kronig for complex kappa from surface roughness,
'					added "configuration" selection for different trace geometries; increased weight for DC value
' 12-Dec-2011 fsr: DC resistance now calculated in dependence of a width-to-height ratio of cross section
' 09-Dec-2011 fsr: Experimental: included option to enforce Kramers-Kronig for dispersive mu
' 08-Dec-2011 fsr: Combined macro with superconductor version; added a check to make sure that fmax>fmin;
'					mu_r can now be entered as a complex function of frequency (variable "F")
' 27-Sep-2011 mku: Included quadratic cross section
' 05-Sep-2011 mku: Corrected inclusion of surface roughness
' 15-Oct-2010 mku: Included surface roughness according to Hammerstad and Jensen paper
' 04-Oct-2010 mku: Deactivated: sCommand = sCommand + "     .Mu """ + "1" +"""" + vbLf
' 04-Oct-2010 mku: Modified to: sCommand = sCommand + "     .SetMaterialUnit """ + sFrqUnit + """, """ + sThicknessUnit + """" + vbLf
' 26-Mar-2009 mku: new version for multilayer conductors (1 and 3 layers are possible)
' 15-Mar-2009 mku: Corrected macro explanation
' 12-Mar-2009 mku: Corrected impedance formula
' 16-Jan-2009 fde: Fixed name problem
' 18-Dec-2008 ube: included In official release
' 18-Dec-2008 fde: Renamed some var to cst_corr
' 09-Dec-2008 fde: first version
'---------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'#include "vba_globals_all.lib"
'#include "infix_postfix.lib"
'#include "complex.lib"

Public sCommand As String
Public cst_materialfolder As String, cst_materialname As String
Public cst_samplepoints As Long
Public cst_min_frq As Double, cst_max_frq As Double
Public cst_mu_r As Double
Public cst_kappa1 As Double
Public cst_DCThickness1 As Double
Public cst_kappa2 As Double
Public cst_DCThickness2 As Double, cst_RFThickness2 As Double
Public sMu1Expr As String, sMu2Expr As String
Public sFrqUnit As String
Public sThicknessUnit As String
Public dThicknessUnit As Double
Public dFrqFactor As Double
Public cst_Zw1 As Complex, cst_Zw2 As Complex
Public cst_arg1_coeff As Complex, cst_arg2_coeff As Complex
Public cst_arg1 As Complex, cst_arg2 As Complex
Public cst_Rrf1 As Complex, cst_Rrf2 As Complex
Public cst_RMS1 As Double, cst_RMS2 As Double
Public cst_Zs As Complex
Public cst_Zin1 As Complex, cst_Zin2 As Complex, cst_Zin2e As Complex, cst_Zin2o As Complex
Public cst_Zin1_numerator As Complex, cst_Zin1_denominator As Complex
Public MaterialFolderArray() As String, MaterialNameArray() As String, MaterialFoldersAndNames As Variant
Public cst_kappa_n As Double
Public cst_Thickness As Double
Public cst_delta_L As Double
Public cst_delta_L0 As Double
Public cst_T As Double
Public cst_Tc As Double
Public cst_RMS As Double
Public cst_thin_layer As Boolean
Public dErrorLimit As Double

Public Const scriptTemplate = GetInstallPath + "\Library\Result Templates\General 1D\mix_templates.bas"
Public Const scriptFile = GetProjectPath("Result") + "mix_tmp.bas"
Public Const SupportedConfigurations = Array("Three layers (symmetric)", "Two layers", "One layer", "Superconductor")
Public Const AvailableRoughnessModels = Array("Hammerstad-Jensen", "Causal Huray")
Public sSupportedConfigurations() As String, sAvailableRoughnessModels() As String

Sub Main()

	Dim i As Long

	' Clean up "Mix_A...F" entries in case there were any left from "Mix Template Results"
	For i = 0 To 5
		DeleteGlobalDataValue("Mix_"+Chr(65+i))
	Next i

	FillArray(sSupportedConfigurations, SupportedConfigurations)
	FillArray(sAvailableRoughnessModels, AvailableRoughnessModels)

	' Some initial settings for both versions
	sThicknessUnit = Units.GetUnit("Length")
	dThicknessUnit = Units.GetGeometryUnitToSI
	sFrqUnit = Units.GetUnit("Frequency")
	dFrqFactor = Units.GetFrequencyUnitToSI
	cst_min_frq= Solver.GetFmin
    cst_max_frq= Solver.GetFmax
    If cst_max_frq <= cst_min_frq Then
    	MsgBox("Please check frequency interval settings. Fmax must be larger than Fmin.","Error")
    	Exit All
    End If

	Begin Dialog UserDialog 930,595,"Tabulated Surface Impedance (Broadband)",.DialogFunc ' %GRID:10,7,1,1

		' General settings
		GroupBox 10,7,910,91,"General Settings",.GroupBox1
		Text 30,35,100,14,"Material folder:",.Text7
		DropListBox 140,28,230,192,MaterialFolderArray(),.MaterialFolderDLB,1
		Text 30,63,100,14,"Material name:",.Text9
		DropListBox 140,56,230,192,MaterialNameArray(),.MaterialNameDLB,1
		PushButton 380,56,120,21,"Restore Settings",.RestoreDialogSettingsPB
		Text 540,35,200,14,"Number of frequency samples:",.Text8
		TextBox 740,28,50,21,.SamplePointsT
		TextBox 740,56,50,21,.ErrorLimitT
		Text 540,63,130,14,"Error limit for data fit:",.Text4
		CheckBox 800,35,110,14,"Log sampling",.LogSamplingCB

		GroupBox 10,105,510,224,"Special Settings",.GroupBox2
		Text 30,133,130,14,"Configuration:",.Text6
		DropListBox 210,126,290,192,sSupportedConfigurations(),.ConfigurationDLB
		Text 30,161,170,14,"Surface roughness model:",.Text1
		DropListBox 210,154,290,192,sAvailableRoughnessModels(),.RoughnessModelDLB
		CheckBox 50,182,230,14,"Enforce causality (experimental)",.EnforceCausalityCB
		Text 30,210,130,14,"For DC resistance:",.Text15
		Text 50,231,270,14,"Width-to-height ratio of total cross section:",.Text5
		TextBox 330,224,60,21,.AspectRatioT
		CheckBox 50,252,150,14,"Coated side walls",.SideWallCoatingCB

		GroupBox 10,336,510,224,"Cross Section",.GroupBox3
		Picture 40,357,450,189,"",0,.Picture1 ' Picture will be defined in DialogFunc_Layer

		' Settings specific to "layer" configuration
		GroupBox 530,105,390,224,"Outer Layer(s)",.GroupBox4
		Text 550,133,150,14,"Thickness1 ["+sThicknessUnit+"]:",.Thickness1Label
		TextBox 720,126,180,21,.DCThickness1T
		Text 550,161,150,14,"Conductivity1 [S/m]:",.Conductivity1Label
		TextBox 720,154,180,21,.kappa1T
		Text 550,189,160,14,"Mu_r1 (function of 'F'):",.Mu1Label
		TextBox 720,182,180,21,.mu_1T
		Text 550,217,130,14,"DeltaRMS1 [um]:",.RMS1Label
		TextBox 720,210,180,21,.rms1T
		Text 550,245,140,14,"Sphere radius [um]:",.SphereRadius1Label
		TextBox 720,238,180,21,.HuraySphereRadius1T
		Text 550,273,130,14,"Number of spheres:",.NumberOfSpheres1Label
		TextBox 720,266,180,21,.HurayNSpheres1T
		Text 550,301,160,14,"Hexagonal area [um^2]:",.HexArea1Label
		TextBox 720,294,180,21,.HurayHexArea1T

		GroupBox 530,336,390,224,"Inner Layer",.GroupBox5
		Text 550,364,150,14,"Thickness2 ["+sThicknessUnit+"]:",.Thickness2Label
		TextBox 720,357,180,21,.DCThickness2T
		Text 550,392,150,14,"Conductivity2 [S/m]:",.Conductivity2Label
		TextBox 720,385,180,21,.kappa2T
		Text 550,420,160,14,"Mu_r2 (function of 'F'):",.Mu2Label
		TextBox 720,413,180,21,.mu_2T
		Text 550,448,130,14,"DeltaRMS2 [um]:",.RMS2Label
		TextBox 720,441,180,21,.rms2T
		Text 550,476,140,14,"Sphere radius [um]:",.SphereRadius2Label
		TextBox 720,469,180,21,.HuraySphereRadius2T
		Text 550,504,130,14,"Number of spheres:",.NumberOfSpheres2Label
		TextBox 720,497,180,21,.HurayNSpheres2T
		Text 550,532,160,14,"Hexagonal area [um^2]:",.HexArea2Label
		TextBox 720,525,180,21,.HurayHexArea2T

		OKButton 530,567,90,21
		PushButton 630,567,90,21,"Apply",.ApplyPB
		PushButton 730,567,90,21,"Exit",.ClosePB
		PushButton 830,567,90,21,"Help",.HelpPB

	End Dialog
'
		Dim dlg As UserDialog
'
		If (Dialog(dlg,-1) = 0) Then Exit All

End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)

	' General variables
	Dim cst_i As Long, j As Long
	Dim cst_freq As Double
    Dim cst_freq_step As Double
    Dim cst_reactance As Double, cst_resistance As Double, cst_weighting As Double
    Dim bLogSampling As Boolean
	Dim iniFileName As String ' ini file to store dialog settings
	Dim sMaterialNameTemp As String, sMaterialFolderTemp As String
	Dim oFreq As Object

	' Variables for layer mode
    Dim dCurrentLayerThickness1 As Double, dCurrentLayerThickness2 As Double, dCurrentRatio As Double, dImpedanceRatio As Double
    Dim cst_corr1 As Double, cst_corr2 As Double
    Dim cst_kappa1rms As Complex, cst_kappa2rms As Complex
    Dim oKappa1RMS, oKappa2RMS As Object
    Dim cst_mu1 As Complex, cst_mu2 As Complex
    Dim cst_tmp_complex As Complex
    Dim oMu1 As Object, oMu2 As Object
    Dim dummy As String
	Dim inputFile As Long, outputFile As Long
	Dim bCoatedSideWalls As Boolean
	Dim bEnforceCausality As Boolean
	Dim dAspectRatio As Double, dArea1 As Double, dArea2 As Double, dCircumference As Double, dWidth As Double, dTotalHeight As Double
	Dim sRoughnessModel As String
	Dim nHurayModelOrder1 As Long, nNumberOfSpheres1() As Long, dHexArea1 As Double, dSphereRadius1() As Double, dSurfRoughLossCoefficient1() As Double, dCriticalFrequency1() As Double, oHurayFactor1 As Object, tmpHuray1 As Complex
	Dim nHurayModelOrder2 As Long, nNumberOfSpheres2() As Long, dHexArea2 As Double, dSphereRadius2() As Double, dSurfRoughLossCoefficient2() As Double, dCriticalFrequency2() As Double, oHurayFactor2 As Object, tmpHuray2 As Complex

	Dim dSkinDepth1 As Double, dSkinDepth2 As Double, dConductorThicknessWithSkinEffect As Double

	' Variables for superconductor mode
	Dim cst_corr As Complex
    Dim cst_kappa_s As Double
    Dim cst_fak As Double
    Dim cst_h As Double
    Dim cst_w1 As Double
    Dim cst_w2 As Double
    Dim cst_gamma As Complex
    Dim cst_Zs As Complex

	cst_min_frq= Solver.GetFmin ' needs to be re-read here in case it was modified during a previous run with logarithmic sampling
	'On Error GoTo InputError

	' ReportInformationToWindow(CStr(Action%)+":"+DlgItem$+":"+CStr(SuppValue%)) ' debug information
   	Select Case Action%
   	    Case 1 ' Dialog box initialization

			DlgText("ErrorLimitT","0.01")
			DlgText("SamplePointsT","21")
			DlgValue("LogSamplingCB", False)

			DlgText("DCThickness1T", "0.1")
			DlgText("mu_1T","1")
			DlgText("kappa1T","4.1e7")
			DlgText("rms1T","0")
			DlgText("HuraySphereRadius1T", "0.5")
			DlgText("HurayNSpheres1T", "70")
			DlgText("HurayHexArea1T", "100")

			DlgText("DCThickness2T", "3")
			DlgText("mu_2T","1")
			DlgText("kappa2T","5.8e7")
			DlgText("rms2T","0")
			DlgText("HuraySphereRadius2T", "0.5")
			DlgText("HurayNSpheres2T", "70")
			DlgText("HurayHexArea2T", "100")

			DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Materials\TabSI_TwoSidedCoating.BMP",0)
			DlgText("AspectRatioT","10")

			MaterialFoldersAndNames = FillMaterialArrays(DlgText("MaterialFolderDLB"))
			MaterialFolderArray = MaterialFoldersAndNames(0)
			MaterialNameArray = MaterialFoldersAndNames(1)
			DlgListBoxArray("MaterialFolderDLB", MaterialFolderArray)
			DlgListBoxArray("MaterialNameDLB", MaterialNameArray)
			DlgText("MaterialFolderDLB","")
			DlgText("MaterialNameDLB","TabulatedSurfaceImpedance")

    Case 2 ' Value changing or button pressed

		Select Case DlgItem$ '(2 open )

            Case "ApplyPB", "OK"
				DlgEnable("ApplyPB", False)
				DlgEnable("OK", False)

	            cst_materialname = NoForbiddenFilenameCharacters(DlgText("MaterialNameDLB"))
				' cst_materialfolder = NoForbiddenFilenameCharacters(DlgText "MaterialFolder") fsr: deactivated so user may use "/" or "\" for subfolders
				cst_materialfolder = Replace(DlgText("MaterialFolderDLB"),"\","/") ' material folder uses / instead of \
				If (cst_materialfolder = "<Main folder>") Then cst_materialfolder = ""
				' Remove last char if it is "/" or "\"
				If (cst_materialfolder<>"") Then
					If (InStr(Mid(cst_materialfolder,Len(cst_materialfolder),1),"/")=1) Then
						cst_materialfolder = Mid(cst_materialfolder, 1, Len(cst_materialfolder)-1)
					End If
				End If

  				cst_samplepoints= CInt(Evaluate(DlgText("SamplePointsT")))
				bLogSampling = CBool(DlgValue("LogSamplingCB"))
  				dErrorLimit = Evaluate(DlgText("ErrorLimitT"))

				' Set up frequency samples in a separate object
				Set oFreq = Result1DComplex("")
				If bLogSampling Then
					oFreq.AppendXY(cst_min_frq,cst_min_frq,0)
                	If cst_min_frq = 0 Then cst_min_frq = cst_max_frq/1e8
	                cst_freq_step = (Log(cst_max_frq) - Log(cst_min_frq))/(cst_samplepoints-1)
	                For cst_i = 2 To cst_samplepoints
   						cst_freq = Exp(Log(cst_min_frq)+(cst_i-1)*cst_freq_step)
        	        	oFreq.AppendXY(cst_freq,cst_freq,0)
            	    Next
				Else
	                cst_freq_step = (cst_max_frq - cst_min_frq)/(cst_samplepoints-1)
	               	For cst_i = 1 To cst_samplepoints
   						cst_freq = cst_min_frq+(cst_i-1)*cst_freq_step
        	        	oFreq.AppendXY(cst_freq,cst_freq,0)
            	    Next
				End If

                sCommand = ""
    			sCommand = sCommand + "With Material" + vbLf
    			sCommand = sCommand + "     .Reset" + vbLf
   				sCommand = sCommand + "     .Name " + Chr(34) + cst_materialname + Chr(34) + vbLf
   				sCommand = sCommand + "     .Folder " + Chr(34) + cst_materialfolder + Chr(34) + vbLf
   				sCommand = sCommand + "     .FrqType """ + "All" + """" + vbLf
    			sCommand = sCommand + "     .SetMaterialUnit """ + sFrqUnit + """, """ + sThicknessUnit + """" + vbLf
   				sCommand = sCommand + "     .Type """ + "Lossy Metal" +"""" + vbLf
				sCommand = sCommand + "     .SetTabulatedSurfaceImpedanceModel """ + "Opaque" +"""" + vbLf

				If (DlgText("ConfigurationDLB")<>"Superconductor") Then

		  		  	sMu1Expr = DlgText("mu_1T")
		  		  	sMu2Expr = DlgText("mu_2T")
	           		cst_DCThickness1 = Evaluate(DlgText "DCThickness1T")*dThicknessUnit
	              	If cst_DCThickness1 < 0 Then
						MsgBox("Thickness of coating must not be < 0.", "Check Settings")
		              	GoTo InputError
		            ElseIf ((cst_DCThickness1 = 0) And (DlgText("ConfigurationDLB")="Two layers")) Then
						MsgBox("Thickness of top layer must be > 0.", "Check Settings")
						GoTo InputError
		            End If
			  		cst_kappa1 = Evaluate(DlgText "kappa1T")
	  			  	cst_kappa2 = Evaluate(DlgText "kappa2T")
	  			  	If cst_kappa1*cst_kappa2 = 0 Then
	  			  		MsgBox("Please enter a valid conductivity.","Check Settings")
						GoTo InputError
		  			End If
	  			  	cst_RMS1 = Evaluate(DlgText("rms1T"))
	              	cst_DCThickness2 = Evaluate(DlgText("DCThickness2T"))*dThicknessUnit
	              	If cst_DCThickness2 <= 0 Then
						MsgBox("Thickness of inner layer must be > 0.", "Check Settings")
		              	GoTo InputError
		            End If
	  			  	cst_RMS2 = Evaluate(DlgText("rms2T"))
	  				dAspectRatio = Evaluate(DlgText("AspectRatioT"))
					bEnforceCausality = CBool(DlgValue("EnforceCausalityCB"))
					bCoatedSideWalls = CBool(DlgValue("SideWallCoatingCB"))

	                cst_corr1 = 1
	                cst_corr2 = 1

					sRoughnessModel = DlgText("RoughnessModelDLB")

					nHurayModelOrder1 = 1 ' hardcoded limit of "1" for now
					dHexArea1 = Evaluate(DlgText("HurayHexArea1T"))*1e-12
					ReDim nNumberOfSpheres1(nHurayModelOrder1-1)
					ReDim dSphereRadius1(nHurayModelOrder1-1)
					For j = 0 To nHurayModelOrder1-1
						nNumberOfSpheres1(j) = Evaluate(DlgText("HurayNSpheres1T"))
						If (nNumberOfSpheres1(j) <= 0) Then
							MsgBox("Number of spheres must be > 0.", "Check Settings")
			              	GoTo InputError
						End If
						dSphereRadius1(j) = Evaluate(DlgText("HuraySphereRadius1T"))*1e-6
						If (dSphereRadius1(j) <= 0) Then
							MsgBox("Sphere radius must be > 0.", "Check Settings")
			              	GoTo InputError
						End If
					Next j
					ReDim dSurfRoughLossCoefficient1(nHurayModelOrder1-1)
					ReDim dCriticalFrequency1(nHurayModelOrder1-1)

					nHurayModelOrder2 = 1 ' hardcoded limit of "1" for now
					dHexArea2 = Evaluate(DlgText("HurayHexArea2T"))*1e-12
					ReDim nNumberOfSpheres2(nHurayModelOrder2-1)
					ReDim dSphereRadius2(nHurayModelOrder2-1)
					For j = 0 To nHurayModelOrder2-1
						nNumberOfSpheres2(j) = Evaluate(DlgText("HurayNSpheres2T"))
						If (nNumberOfSpheres2(j) <= 0) Then
							MsgBox("Number of spheres must be > 0.", "Check Settings")
			              	GoTo InputError
						End If
						dSphereRadius2(j) = Evaluate(DlgText("HuraySphereRadius2T"))*1e-6
						If (dSphereRadius2(j) <= 0) Then
							MsgBox("Sphere radius must be > 0.", "Check Settings")
			              	GoTo InputError
						End If
					Next j
					ReDim dSurfRoughLossCoefficient2(nHurayModelOrder2-1)
					ReDim dCriticalFrequency2(nHurayModelOrder2-1)

	                Set oMu1 = Result1DComplex("")
	                Set oMu2 = Result1DComplex("")
					Set oKappa1RMS = Result1DComplex("")
					Set oKappa2RMS = Result1DComplex("")
					Set oHurayFactor1 = Result1DComplex("")
					Set oHurayFactor2 = Result1DComplex("")

	                ' add a zero term to make sure "F" is always in expression, then replace all occurrences with "F_cst_tmp"
	                sMu1Expr = CST_ReplaceString(sMu1Expr + "+0*F", "F")
	                sMu2Expr = CST_ReplaceString(sMu2Expr + "+0*F", "F")

					sMu1Expr = Replace(sMu1Expr," ","")	' remove all spaces from expression
					sMu2Expr = Replace(sMu2Expr," ","")	' remove all spaces from expression

					' Convert expressions from infix to postfix to VBA code
					sMu1Expr = InfixToPostfix(sMu1Expr)
					sMu1Expr = PostfixToVBA_1DC(sMu1Expr)
					sMu1Expr = sMu1Expr + vbNewLine + "Set EvaluateExpression = TmpResult_Final"
					sMu2Expr = InfixToPostfix(sMu2Expr)
					sMu2Expr = PostfixToVBA_1DC(sMu2Expr)
					sMu2Expr = sMu2Expr + vbNewLine + "Set EvaluateExpression = TmpResult_Final"

					' Save frequency abscissa as object
					StoreGlobalDataValue("Mix_F", "TRUE")
					oFreq.Save(GetProjectPath("Result") + "mix_tmpF1DC.sig")

					' Write the script file for mu1
					inputFile = FreeFile()
					Open scriptTemplate For Input As #inputFile
					outputFile = FreeFile()
					Open scriptFile For Output As #outputFile
					While Not EOF(inputFile)
						Line Input #inputFile, dummy
						If (InStr(dummy, "SET_APPLICATION_NAME") >0) Then
							dummy = "Public Const callingApp = " + Chr(34) + GetApplicationName() + Chr(34)
						End If
						If (InStr(dummy, "SET_COMPLEXITY_LEVEL") >0) Then
							dummy = "Public Const complexityLevel = " + Chr(34) + "1DC" + Chr(34)
						End If
						If (InStr(dummy, "SET_DEBUG_OUTPUT") >0) Then
							dummy = "Public Const DebugOutput = True"
						End If
						If (InStr(dummy, "EXPRESSION_TO_BE_REPLACED") > 0) Then
							' put in actual expression
							dummy = sMu1Expr
						End If
						Print #outputFile, dummy
					Wend
					Close #outputFile
					Close #inputFile
					' Shell "notepad " & scriptFile, 3
					' Run script and store result in mu1 object
					RunScript scriptFile
					Set oMu1 = Result1DComplex("mix_template_result1DC")
					If(bEnforceCausality And DlgEnable("mu_1T") And Not IsNumeric(DlgText("mu_1T"))) Then EnforceKramersKronig(DlgText("mu_1T") + "+0*F", oMu1) ' fsr: DlgText(mu1) can NOT be replaced with sMu1Expr here

					' Write the script file for mu2
					inputFile = FreeFile()
					Open scriptTemplate For Input As #inputFile
					outputFile = FreeFile()
					Open scriptFile For Output As #outputFile
					While Not EOF(inputFile)
						Line Input #inputFile, dummy
						If (InStr(dummy, "SET_APPLICATION_NAME") >0) Then
							dummy = "Public Const callingApp = " + Chr(34) + GetApplicationName() + Chr(34)
						End If
						If (InStr(dummy, "SET_COMPLEXITY_LEVEL") >0) Then
							dummy = "Public Const complexityLevel = " + Chr(34) + "1DC" + Chr(34)
						End If
						If (InStr(dummy, "SET_DEBUG_OUTPUT") >0) Then
							dummy = "Public Const DebugOutput = True"
						End If
						If (InStr(dummy, "EXPRESSION_TO_BE_REPLACED") > 0) Then
							' put in actual expression
							dummy = sMu2Expr
						End If
						Print #outputFile, dummy
					Wend
					Close #outputFile
					Close #inputFile
					' Shell "notepad " & scriptFile, 3
					' Run script and store result in mu2 object
					RunScript scriptFile
					Set oMu2 = Result1DComplex("mix_template_result1DC")
					If(bEnforceCausality And DlgEnable("mu_2T") And Not IsNumeric(DlgText("mu_2T"))) Then EnforceKramersKronig(DlgText("mu_2T") + "+0*F", oMu2) ' fsr: DlgText(mu2) can NOT be replaced with sMu2Expr here

					' Calculate dispersive kappa from surface roughness correction factor
					For cst_i = 1 To cst_samplepoints
						cst_freq = oFreq.GetX(cst_i-1)
						If cst_freq=0 Then
							oKappa1RMS.AppendXY(cst_freq, cst_kappa1,0)
							oKappa2RMS.AppendXY(cst_freq, cst_kappa2,0)
							oHurayFactor1.AppendXY(cst_freq,1,0)
							oHurayFactor2.AppendXY(cst_freq,1,0)
						ElseIf (sRoughnessModel = "Hammerstad-Jensen") Then
		   				   	cst_mu1 = AssignComplex(oMu1.GetYRe(cst_i-1), oMu1.GetYIm(cst_i-1))
							cst_mu2 = AssignComplex(oMu2.GetYRe(cst_i-1), oMu2.GetYIm(cst_i-1))
						'	If (absolute(cst_mu1)*absolute(cst_mu2) = 0) Then
							'	MsgBox("Error evaluating mu. Please check expression.")
						'	Exit Function
						'	End If
					        ' To include surface roughness: START
							' skin depth for top layer
							dSkinDepth1 = Sqr(2/(8*Pi^2*1E-7*absolute(cst_mu1)*cst_kappa1*cst_freq*dFrqFactor))
							' surface roughness correction factor according to Hammerstad/Jensen
							cst_corr1 = (1 + 2/Pi * Atn(1.4*(cst_RMS1*1E-6/dSkinDepth1)^2))^2
							oKappa1RMS.AppendXY(cst_freq, cst_kappa1/cst_corr1,0)

							' skin depth for bottom layer
							dSkinDepth2 = Sqr(2/(8*Pi^2*1E-7*absolute(cst_mu2)*cst_kappa2*cst_freq*dFrqFactor))
							' surface roughness correction factor according to Hammerstad/Jensen
							cst_corr2 = (1 + 2/Pi * Atn(1.4*(cst_RMS2*1E-6/dSkinDepth2)^2))^2
							oKappa2RMS.AppendXY(cst_freq, cst_kappa2/cst_corr2,0)
							' To include surface roughness: END
						ElseIf (sRoughnessModel = "Causal Huray") Then
		   				   	cst_mu1 = AssignComplex(oMu1.GetYRe(cst_i-1), 0)
							cst_mu2 = AssignComplex(oMu2.GetYRe(cst_i-1), 0)
							oKappa1RMS.AppendXY(cst_freq, cst_kappa1,0)
							oKappa2RMS.AppendXY(cst_freq, cst_kappa2,0)
							tmpHuray1=AssignComplex(1,0)
							For j = 0 To nHurayModelOrder1-1
								dSurfRoughLossCoefficient1(j) = 6*Pi*dSphereRadius1(j)^2*nNumberOfSpheres1(j)/dHexArea1
								dCriticalFrequency1(j) = 2/dSphereRadius1(j)^2/4e-7/Pi/cst_mu1.re/cst_kappa1
								'ReportInformationToWindow(dCriticalFrequency1(j))
								tmpHuray1 = plus(tmpHuray1, div(AssignComplex(dSurfRoughLossCoefficient1(j),0),plus(AssignComplex(1,0),multsc(AssignComplex(1, -1), 0.5 * Sqr(dCriticalFrequency1(j)/(2*Pi*cst_freq*dFrqFactor))))))
							Next
							oHurayFactor1.AppendXY(cst_freq, tmpHuray1.re, tmpHuray1.im)

							tmpHuray2=AssignComplex(1,0)
							For j = 0 To nHurayModelOrder2-1
								dSurfRoughLossCoefficient2(j) = 6*Pi*dSphereRadius2(j)^2*nNumberOfSpheres2(j)/dHexArea2
								dCriticalFrequency2(j) = 2/dSphereRadius2(j)^2/4e-7/Pi/cst_mu2.re/cst_kappa2
								'ReportInformationToWindow(dCriticalFrequency2(j))
								tmpHuray2 = plus(tmpHuray2, div(AssignComplex(dSurfRoughLossCoefficient2(j),0),plus(AssignComplex(1,0),multsc(AssignComplex(1, -1), 0.5 * Sqr(dCriticalFrequency2(j)/(2*Pi*cst_freq*dFrqFactor)))))) 
							Next
							oHurayFactor2.AppendXY(cst_freq, tmpHuray2.re, tmpHuray2.im)

						Else
							ReportError("Unknown roughness model.")
						End If
					Next

					If(bEnforceCausality And DlgEnable("RMS1T") And (Evaluate(cst_RMS1)>0)) Then EnforceKramersKronig("(1 + 2/Pi * Atn(1.4*("+CStr(Evaluate(cst_RMS1))+")*1E-6*Pi*F*("+DlgText("mu_1T")+")*"+CStr(Evaluate(cst_kappa1))+"))^2",oKappa1RMS)
					If(bEnforceCausality And DlgEnable("RMS2T") And (Evaluate(cst_RMS2)>0)) Then EnforceKramersKronig("(1 + 2/Pi * Atn(1.4*("+CStr(Evaluate(cst_RMS2))+")*1E-6*Pi*F*("+DlgText("mu_2T")+")*"+CStr(Evaluate(cst_kappa2))+"))^2",oKappa2RMS)

					For cst_i = 1 To cst_samplepoints
						cst_freq = oFreq.GetX(cst_i-1)

	   				   	cst_mu1 = AssignComplex(oMu1.GetYRe(cst_i-1), oMu1.GetYIm(cst_i-1))
						cst_mu2 = AssignComplex(oMu2.GetYRe(cst_i-1), oMu2.GetYIm(cst_i-1))
	   				   	cst_kappa1rms = AssignComplex(oKappa1RMS.GetYRe(cst_i-1), oKappa1RMS.GetYIm(cst_i-1))
						cst_kappa2rms = AssignComplex(oKappa2RMS.GetYRe(cst_i-1), oKappa2RMS.GetYIm(cst_i-1))

				    	If DlgText("ConfigurationDLB")="Two layers" Then
							dTotalHeight = cst_DCThickness2+cst_DCThickness1
						ElseIf (DlgText("ConfigurationDLB")="Three layers (symmetric)" Or DlgText("ConfigurationDLB")="One layer") Then
							dTotalHeight = cst_DCThickness2+2*cst_DCThickness1
						Else
							ReportError("Unknown layer configuration.")
						End If
						dWidth = dAspectRatio * dTotalHeight

						If cst_freq=0 Then ' if f=0: calculate R as inner and outer layer in parallel
					    	cst_weighting = 10 ' weighting factor for DC sample
							dArea2 = cst_DCThickness2*dWidth ' cross section of inner layer: thickness2*width
							If bCoatedSideWalls Then dArea2=dArea2-2*cst_DCThickness2*cst_DCThickness1 	' subtract side walls
							dArea1 = dTotalHeight*dWidth - dArea2 'full rect minus area2
							dCircumference = 2*(dAspectRatio+1)*dTotalHeight
							' DC resistance is a parallel circuit of the two layers with dArea1 and dArea2
							cst_Zs.re = (dCircumference)/(cst_kappa1*dArea1+cst_kappa2*dArea2)/IIf(DlgText("ConfigurationDLB")="Two layers", 2, 1)
					    	cst_Zs.im = 0
					    	' ReportInformationToWindow(cst_Zs.re)
						Else
							cst_weighting = 1

							' If coated on one side only -> double inner layer thickness to undo symmetry (true "open" on bottom side of conductor)
							If (DlgText("ConfigurationDLB")="Two layers") Then
								cst_RFThickness2 = 2*cst_DCThickness2
							Else
								cst_RFThickness2 = cst_DCThickness2
							End If
							' ReportInformationToWindow("RF Thickness at f=" & CStr(cst_freq) & ": " & CStr(cst_RFThickness2))

							' skin depth = Sqr(2/(kappa*mu*omega))
							dSkinDepth1 = Sqr(1/(absolute(mult(cst_mu1, cst_kappa1rms))*4e-7*Pi^2*cst_freq*dFrqFactor))
							dSkinDepth2 = Sqr(1/(absolute(mult(cst_mu2, cst_kappa2rms))*4e-7*Pi^2*cst_freq*dFrqFactor))
							' Rs = 1/(skin depth*conductivity)=sqrt(omega*mu/(2*conductivity))
		             		cst_Rrf1 = multsc(ComplexSqrt(div(cst_mu1,cst_kappa1rms)),Sqr(4*Pi^2*1e-7*cst_freq*dFrqFactor)) ' outer layer
		             		cst_Rrf2 = multsc(ComplexSqrt(div(cst_mu2,cst_kappa2rms)),Sqr(4*Pi^2*1e-7*cst_freq*dFrqFactor)) ' inner layer

		             		If sRoughnessModel = "Causal Huray" Then
		             			cst_Rrf1 = mult(cst_Rrf1, AssignComplex(oHurayFactor1.GetYRe(cst_i-1), oHurayFactor1.GetYIm(cst_i-1)))
		             			cst_Rrf2 = mult(cst_Rrf2, AssignComplex(oHurayFactor2.GetYRe(cst_i-1), oHurayFactor2.GetYIm(cst_i-1)))
		             		End If

	   	             		' arguments of coth and tanh: propagation const gamma times layer thickness
						    cst_arg1_coeff = multsc(ComplexSqrt(mult(AssignComplex(0,1),mult(cst_mu1,cst_kappa1rms))),Sqr(8*Pi^2*1e-7)*cst_DCThickness1) ' outer layer
		   				    cst_arg2_coeff = multsc(ComplexSqrt(mult(AssignComplex(0,1),mult(cst_mu2,cst_kappa2rms))),Sqr(8*Pi^2*1e-7)*cst_RFThickness2/2) ' inner layer Thickness/2 to reflect magnetic symmetry in center

		   				    ' Zs = (1+j)*Rs, intrinsic impedance
		   				    cst_Zw1 = mult(AssignComplex(1,1),cst_Rrf1) ' outer layer
		   				    cst_Zw2 = mult(AssignComplex(1,1),cst_Rrf2) ' inner layer

		   				    cst_arg1 = multsc(cst_arg1_coeff,Sqr(cst_freq*dFrqFactor)) ' outer layer
		   				    cst_arg2 = multsc(cst_arg2_coeff,Sqr(cst_freq*dFrqFactor)) ' inner layer
		   				    ' correction factor w/(w+t) applied to impedance below to account for full circumference at DC
						    cst_Zin2 = mult(cst_Zw2,zcoth(cst_arg2))

						    ' Calculate total resistance from staggered transformation
						    cst_Zin1_numerator = plus(cst_Zin2,mult(cst_Zw1,ztanh(cst_arg1)))
						    cst_Zin1_denominator = plus(cst_Zw1,mult(cst_Zin2,ztanh(cst_arg1)))
   		   				    ' correction factor w/(w+t) applied to impedance below to account for full circumference at DC
   		   				    If (cst_DCThickness1 > 0) Then
   		   				    	' outer and inner layer
   		   				    	If (dSkinDepth1 > 2 * cst_DCThickness1) Then ' field penetrates into inner layer
   		   				    		If (dSkinDepth2 > 2 * cst_DCThickness2) Then
   		   				    			dConductorThicknessWithSkinEffect = 2 * cst_DCThickness1 + cst_DCThickness2 ' full thickness
   		   				    		Else
										dConductorThicknessWithSkinEffect = cst_DCThickness1 + 2 * dSkinDepth2 ' inner conductor is affected by skin effect
   		   				    		End If
   		   				    	Else
									dConductorThicknessWithSkinEffect = 2 * dSkinDepth1 ' outer conductor is affected by skin effect, neglect field penetration into inner conductor
   		   				    	End If
   		   				    Else
   		   				    	' inner layer only
   		   				    	If (dSkinDepth2 > 2 * cst_DCThickness2) Then ' field penetrates into inner layer
									dConductorThicknessWithSkinEffect = cst_DCThickness2 ' inner conductor is affected by skin effect
   		   				    	Else
									dConductorThicknessWithSkinEffect = 2 * dSkinDepth2 ' outer conductor is affected by skin effect, neglect field penetration into inner conductor
   		   				    	End If
   		   				    End If
						    cst_Zin1 = multsc(mult(cst_Zw1,div(cst_Zin1_numerator,cst_Zin1_denominator)), (dWidth + dConductorThicknessWithSkinEffect)/dWidth)
							' ReportInformationToWindow("Impedance at f=" & CStr(cst_freq) & ": " & CStr(cst_Zin1.re))

						    cst_Zs = cst_Zin1

						End If
	'
					   cst_resistance = cst_Zs.re
					   cst_reactance = cst_Zs.im
	'
	                   sCommand = sCommand + "     .AddTabulatedSurfaceImpedanceFittingValue """ + CStr(cst_freq) + """,""" + CStr(cst_resistance) + """,""" + CStr(cst_reactance) + """,""" + CStr(cst_weighting) + """" + vbLf
	                   'ReportInformationToWindow(USFormat(cst_freq,"Scientific") + ":" + USFormat(cst_resistance,"Scientific") + " + j*" + USFormat(cst_reactance,"Scientific"))
	'
	   				Next

	   				cst_materialfolder = Replace(cst_materialfolder, "/", "\") ' Now we need "\" instead of "/" again for the rest of the tree
	   				If (DlgText("ConfigurationDLB")<>"One layer") Then
		   				oMu1.Save(GetProjectPath("Result")+"\1D Results\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"_Mu1")
		   				oMu1.AddToTree("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Mu (Coating)")
		   				SelectTreeItem("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Mu (Coating)")
		   				Resulttree.DeleteAt("never")
		   				oKappa1RMS.Save(GetProjectPath("Result")+"\1D Results\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"_Kappa1")
		   				oKappa1RMS.AddToTree("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Kappa (Coating)")
		   				SelectTreeItem("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Kappa (Coating)")
		   				Resulttree.DeleteAt("never")
		   			End If
	   				oMu2.Save(GetProjectPath("Result")+"\1D Results\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"_Mu2")
	   				oMu2.AddToTree("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Mu (Inner Layer)")
	   				SelectTreeItem("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Mu (Inner Layer)")
	   				Resulttree.DeleteAt("never")
	   				oKappa2RMS.Save(GetProjectPath("Result")+"\1D Results\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"_Kappa2")
	   				oKappa2RMS.AddToTree("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Kappa (Inner Layer)")
	   				SelectTreeItem("1D Results\Layered Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Dispersive\Kappa (Inner Layer)")
	   				Resulttree.DeleteAt("never")

	   				sCommand = sCommand + "     .ErrorLimitNthModelFitTabSI " +Chr(34) + CStr(dErrorLimit) + Chr(34) + vbLf
	       			sCommand = sCommand + "     .Create" + vbLf
	       			sCommand = sCommand + "End With" + vbLf

				Else ' Configuration is "Superconductor"

	              	cst_Thickness = Evaluate(DlgText "DCThickness2T")*dThicknessUnit
	              	cst_mu_r = Evaluate(DlgText "mu_2T")
	  			  	cst_T = Evaluate(DlgText("HuraySphereRadius2T")) ' temperature of superconductor, double use of dialog item
	  			  	cst_delta_L0 = Evaluate(DlgText("HurayNSpheres2T")) ' double use of Dialog item
	  			  	cst_Tc = Evaluate(DlgText("HurayHexArea2T")) ' critical temperature, double use of dialog item
	  			  	If (cst_Tc<=cst_T) Then
						MsgBox("Critical temperature must be larger than temperature.", "Check Settings")
		              	GoTo InputError
	  			  	End If
	  			  	cst_kappa_n = Evaluate(DlgText("kappa2T"))
	  			  	cst_RMS = Evaluate(DlgText("rms2T"))
	'
	                cst_corr.re = 1.0
	                cst_corr.im = 1.0
	'
					cst_delta_L = cst_delta_L0 / Sqr(1-(cst_T/cst_Tc)^4)*1E-9

	   			    For cst_i = 1 To cst_samplepoints
	   				   cst_freq = oFreq.GetX(cst_i-1)
	'
	   				   If cst_freq = 0 Then
					      cst_Zs.re = 0
					      cst_Zs.im = 0
					   Else
	'
	' To include surface roughness: START
	'
	                      dSkinDepth1 = Sqr(2/(8*Pi^2*1E-7*cst_mu_r*cst_kappa_n*cst_freq*dFrqFactor))
	'
	                      cst_corr.re = 1 + 2/Pi * Atn(((cst_RMS/cst_delta_L)^2)*(0.35+1.05*Exp(-0.5*(dSkinDepth1/cst_delta_L)^2)))
	                      cst_corr.im = 2 - Exp(-(cst_RMS/cst_delta_L)*(1+2.5*Exp(-0.5*(dSkinDepth1/cst_delta_L)^2)))
	'
	' To include surface roughness: END
	'
	                      cst_kappa_s = 1 / (8*Pi^2*1e-7*cst_mu_r*cst_delta_L^2*cst_freq*dFrqFactor)
	'
	                      cst_h = Sqr(cst_kappa_s^2+cst_kappa_n^2)
	                      cst_w1 = Sqr(0.5*(cst_h + cst_kappa_s))
	                      cst_w2 = Sqr(0.5*(cst_h - cst_kappa_s))
	                      cst_fak = Sqr(8*Pi^2*1e-7*cst_mu_r*cst_freq*dFrqFactor)
	'
	 				      cst_Zs.re = cst_fak/cst_h * cst_w2 * cst_corr.re
					      cst_Zs.im = cst_fak/cst_h * cst_w1 * cst_corr.im
	'
	                      If cst_thin_layer Then
	                         cst_gamma.re = cst_fak * cst_w1 * cst_Thickness * 0.5
	                         cst_gamma.im = cst_fak * cst_w2 * cst_Thickness * 0.5
	'
	                         cst_Zs = mult(cst_Zs,zcoth(cst_gamma))
	                      End If

					   End If
	'
					   cst_resistance = cst_Zs.re
					   cst_reactance = cst_Zs.im
	'
	                   sCommand = sCommand + "     .AddTabulatedSurfaceImpedanceValue """ + CStr(cst_freq) + """,""" + CStr(cst_resistance) + """,""" + CStr(cst_reactance) + """" + vbLf
	'
	   				Next
	   				sCommand = sCommand + "     .ErrorLimitNthModelFitTabSI " +Chr(34) + CStr(dErrorLimit) + Chr(34) + vbLf
	       			sCommand = sCommand + "     .Create" + vbLf
	       			sCommand = sCommand + "End With" + vbLf

				End If ' Decision between "layer" and "superconductor" configuration

				cst_materialfolder = Replace(cst_materialfolder, "/", "\") ' Now we need "\" instead of "/" again for the rest of the tree
   				AddToHistory("define Material: " + IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname, sCommand)
				SelectTreeItem("1D Results\Materials\"+IIf(cst_materialfolder<>"",cst_materialfolder+"\","")+cst_materialname+"\Surface Impedance")

				' Update dropdown boxes
				MaterialFoldersAndNames = FillMaterialArrays(DlgText("MaterialFolderDLB"))
				MaterialFolderArray = MaterialFoldersAndNames(0)
				MaterialNameArray = MaterialFoldersAndNames(1)
				sMaterialFolderTemp = DlgText("MaterialFolderDLB")
				sMaterialNameTemp = DlgText("MaterialNameDLB")
				DlgListBoxArray("MaterialFolderDLB", MaterialFolderArray)
				DlgListBoxArray("MaterialNameDLB", MaterialNameArray)
				DlgText("MaterialFolderDLB", sMaterialFolderTemp)
				DlgText("MaterialNameDLB", sMaterialNameTemp)

				' Store dialog settings
  				iniFileName = DlgText("MaterialFolderDLB")+"_"+DlgText("MaterialNameDLB")+".ini"
				iniFileName = Replace(iniFileName, "\", "")
				iniFileName = Replace(iniFileName, "/", "")
				iniFileName = Replace(iniFileName, "<Main folder>", "")
				StoreAllDialogSettings(GetProjectPath("Model3D")+iniFileName, "MaterialFolderDLB,MaterialNameDLB", "") ' exclude material name and folder

				DlgEnable("ApplyPB", True)
				DlgEnable("OK", True)
				DialogFunc% = DlgItem = "ApplyPB"

            Case "ThinLayerCB"
            	cst_thin_layer = CBool(SuppValue)
            	DlgEnable("ThicknessT", cst_thin_layer)

    		Case "RestoreDialogSettingsPB"
  				iniFileName = DlgText("MaterialFolderDLB")+"_"+DlgText("MaterialNameDLB")+".ini"
				iniFileName = Replace(iniFileName, "\", "")
				iniFileName = Replace(iniFileName, "/", "")
				iniFileName = Replace(iniFileName, "<Main folder>", "")
				'ReportInformationToWindow(iniFileName)
				ReStoreAllDialogSettings(GetProjectPath("Model3D")+iniFileName, "MaterialFolderDLB,MaterialNameDLB", "") ' exclude material name and folder
    			DialogFunc% = True

    		Case "ClosePB"
				Exit All

			Case "HelpPB"
				StartHelp "common_preloadedmacro_materials_tabulated_surface_impedance"
    			DialogFunc% = True

		End Select '(2 close)

		Case 3   ' Text box changed
		Case 4
			Select Case DlgItem
				Case "MaterialNameDLB"
					MaterialFoldersAndNames = FillMaterialArrays(DlgText("MaterialFolderDLB"))
					MaterialFolderArray = MaterialFoldersAndNames(0)
					MaterialNameArray = MaterialFoldersAndNames(1)
					sMaterialFolderTemp = DlgText("MaterialFolderDLB")
					sMaterialNameTemp = DlgText("MaterialNameDLB")
					DlgListBoxArray("MaterialFolderDLB", MaterialFolderArray)
					DlgListBoxArray("MaterialNameDLB", MaterialNameArray)
					DlgText("MaterialFolderDLB", sMaterialFolderTemp)
					DlgText("MaterialNameDLB", sMaterialNameTemp)
			End Select
		Case 5 ' Idle

       End Select '(1 close)

    UpdateDependentDialogSettings()

    Exit Function
	InputError:
		MsgBox("Error evaluating input values, please check your settings.","Error")
		DialogFunc% = True
		DlgEnable("ApplyPB", True)
		DlgEnable("OK", True)
		Exit Function

End Function

Sub UpdateDependentDialogSettings()

	' This function updates all the settings (values, visibility, availability, ...) of dialog items that depend on other dialog items

	' Superconductor only supports Hammerstad model for surface roughness
	' Aspect ratio for DC resistance calculation does not apply, either
	Select Case DlgText("ConfigurationDLB")
		Case "Superconductor"
			DlgText("RoughnessModelDLB", "Hammerstad-Jensen")
			DlgEnable("RoughnessModelDLB", False)
			DlgEnable("AspectRatioT", False)
		Case Else
			DlgEnable("RoughnessModelDLB", True)
			DlgEnable("AspectRatioT", True)
	End Select

	' Select picture
	Select Case DlgText("ConfigurationDLB")
		Case "Three layers (symmetric)"
			DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Materials\TabSI_TwoSidedCoating.BMP",0)
		Case "Two layers"
			DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Materials\TabSI_OneSidedCoating.BMP",0)
		Case "One layer", "Superconductor"
			DlgSetPicture("Picture1", GetInstallPath()+"\Library\Macros\Materials\TabSI_SingleLayer.BMP",0)
	End Select

	' Set up active/inactive
	Select Case DlgText("ConfigurationDLB")
		Case "Three layers (symmetric)", "Two layers"
			DlgEnable("DCThickness1T",True)
			DlgEnable("kappa1T",True)
			DlgEnable("mu_1T",True)
			DlgEnable("EnforceCausalityCB",True)
			DlgEnable("SideWallCoatingCB",True)
			Select Case DlgText("RoughnessModelDLB")
				Case "Hammerstad-Jensen"
					DlgEnable("rms1T",True)
					DlgEnable("HuraySphereRadius1T", False)
					DlgEnable("HurayNSpheres1T", False)
					DlgEnable("HurayHexArea1T", False)
					DlgEnable("rms2T",True)
					DlgEnable("HuraySphereRadius2T", False)
					DlgEnable("HurayNSpheres2T", False)
					DlgEnable("HurayHexArea2T", False)
				Case "Causal Huray"
					DlgEnable("rms1T",False)
					DlgEnable("HuraySphereRadius1T", True)
					DlgEnable("HurayNSpheres1T", True)
					DlgEnable("HurayHexArea1T", True)
					DlgEnable("rms2T",False)
					DlgEnable("HuraySphereRadius2T", True)
					DlgEnable("HurayNSpheres2T", True)
					DlgEnable("HurayHexArea2T", True)
			End Select
		Case "One layer"
			DlgEnable("DCThickness1T",False)
			DlgText("DCThickness1T", "0")
			DlgEnable("kappa1T",False)
			DlgEnable("mu_1T",False)
			DlgEnable("EnforceCausalityCB",True)
			DlgEnable("SideWallCoatingCB",False)
			DlgEnable("rms1T",False)
			DlgEnable("HuraySphereRadius1T", False)
			DlgEnable("HurayNSpheres1T", False)
			DlgEnable("HurayHexArea1T", False)
			Select Case DlgText("RoughnessModelDLB")
				Case "Hammerstad-Jensen"
					DlgEnable("rms2T",True)
					DlgEnable("HuraySphereRadius2T", False)
					DlgEnable("HurayNSpheres2T", False)
					DlgEnable("HurayHexArea2T", False)
				Case "Causal Huray"
					DlgEnable("rms2T",False)
					DlgEnable("HuraySphereRadius2T", True)
					DlgEnable("HurayNSpheres2T", True)
					DlgEnable("HurayHexArea2T", True)
			End Select
		Case "Superconductor"
			DlgEnable("DCThickness1T",False)
			DlgText("DCThickness1T", "0")
			DlgEnable("kappa1T",False)
			DlgEnable("mu_1T",False)
			DlgEnable("EnforceCausalityCB",False)
			DlgValue("EnforceCausalityCB",False)
			DlgEnable("SideWallCoatingCB",False)
			DlgEnable("rms1T",False)
			DlgEnable("HuraySphereRadius1T", False)
			DlgEnable("HurayNSpheres1T", False)
			DlgEnable("HurayHexArea1T", False)
			DlgEnable("rms2T",True)
			DlgEnable("HuraySphereRadius2T", True)
			DlgEnable("HurayNSpheres2T", True)
			DlgEnable("HurayHexArea2T", True)
		End Select

	' Set up labels
	Select Case DlgText("ConfigurationDLB")
		Case "Three layers (symmetric)", "Two layers", "One layer"
			DlgText("Thickness2Label", "Thickness2 ["+sThicknessUnit+"]:")
			DlgText("Conductivity2Label","Conductivity2 [S/m]:")
			DlgText("Mu2Label","Mu_r2 (Function of 'F'):")
			DlgText("RMS2Label","DeltaRMS2 [um]:")
			DlgText("SphereRadius2Label","Sphere radius [um]:")
			DlgText("NumberOfSpheres2Label","Number of spheres:")
			DlgText("HexArea2Label","Hexagonal area [um^2]:")
		Case "Superconductor"
			DlgText("Thickness2Label", "Thickness2 ["+sThicknessUnit+"]:")
			DlgText("Conductivity2Label","Conductivity2 [S/m]:")
			DlgText("Mu2Label","Mu_r2 (Function of 'F'):")
			DlgText("RMS2Label","DeltaRMS2 [um]:")
			DlgText("SphereRadius2Label","Temperature [K]:")
			DlgText("NumberOfSpheres2Label","delta_L (T=0K) [nm]:")
			DlgText("HexArea2Label","Critical temperature [K]:")
	End Select

End Sub

Sub EnforceKramersKronig(sFunctionInF As String, oResult As Object)

	' This function ensures that the Kramers-Kronig relationship is satisfied between real and imaginary part of oResult
	' The imaginary part of oResult will be replaced by the Hilbert transform of the real part of oResult
	' It is necessary to know the analytical expression that was used to create oResult, given as sFunctionInF

	Dim i As Long
	Dim X As Double
	Dim dIntegral, dDeltaF, dEpsilon, dIntegralLength As Double

	'ReportInformationToWindow(sFunctionInF)
	dDeltaF = .2 ' integral stepwidth
	dEpsilon = 1e-6 ' epsilon for Cauchy principal value
	dIntegralLength = 200 ' integral length for Hilbert transform

	sFunctionInF = CST_ReplaceString(sFunctionInF, "F")
	For i = 0 To oResult.GetN()-1
		' Hilbert transform by calculating Cauchy principal value
		' Integral split in two parts left and right of singularity, these parts are then added
		dIntegral = 0
		For X = oResult.GetX(i)-dEpsilon To oResult.GetX(i)-dIntegralLength/2 STEP -dDeltaF
				dIntegral = dIntegral + 1/Pi*Evaluate(Replace(sFunctionInF, "F_cst_tmp", "("+CSTr(X)+")"))*dDeltaF/(X-oResult.GetX(i))
		Next
		For X = oResult.GetX(i)+dEpsilon To oResult.GetX(i)+dIntegralLength/2 STEP dDeltaF
				dIntegral = dIntegral + 1/Pi*Evaluate(Replace(sFunctionInF, "F_cst_tmp", "("+CSTr(X)+")"))*dDeltaF/(X-oResult.GetX(i))
		Next
		oResult.SetYIm(i,dIntegral)
	Next

End Sub

Function FillMaterialArrays(Optional sLocatedInMaterialFolder As String, Optional sMaterialType As String) As Variant

	' This function parses all material entries in fills the folders as well as the materials it found
	' The user may restrict the materials listed to materials in a certain folder or of a certain type

	' Inputs:
	'	Optional sLocatedInMaterialFolder As String	:	Only materials in this folder will be included in sMaterialList;
	'													if no folder is given, this value defaults to "" and only materials in the main
	'													material folder will be listed. Use "CST_ALFOLDERS" to list all materials in all folders
	'	Optional sMaterialType As String			:	If this string is not empty, only materials of the given type will be listed;
	'													if no type is given, this value defaults to "" and all materials will be listed.
	' Output/Return values:
	' 	FillMaterialArrays As Variant				:	An array of two elements, first element is a string array containing the folder list,
	'													second element is a string array containing the list of material names

	Dim nMaterials As Long, i As Long, j As Long
	Dim sMaterialList() As String, sMaterialFolderList() As String
	Dim sSplitMaterialEntry() As String, sFoundMaterialName As String, sFoundMaterialFolder As String, sFoundMaterialType As String

	If ((sLocatedInMaterialFolder="<Main folder>") Or (sLocatedInMaterialFolder = "")) Then
		sLocatedInMaterialFolder = ""
	' Append "/" if none at the end
	ElseIf (Mid(sLocatedInMaterialFolder, Len(sLocatedInMaterialFolder),1) = "/") Then
		sLocatedInMaterialFolder = sLocatedInMaterialFolder
	Else
		sLocatedInMaterialFolder = sLocatedInMaterialFolder + "/"
	End If

	ReDim sMaterialFolderList(0)
	sMaterialFolderList(0) = "<Main folder>"
	ReDim sMaterialList(0)
	nMaterials = Material.GetNumberOfMaterials
	For i = 0 To nMaterials-1
		sSplitMaterialEntry = Split(Material.GetNameOfMaterialFromIndex(i),"/")
		sFoundMaterialName = sSplitMaterialEntry(UBound(sSplitMaterialEntry))
		sFoundMaterialType = Material.GetTypeOfMaterial(Material.GetNameOfMaterialFromIndex(i))
		If (((sMaterialType<>"") And (sFoundMaterialType<>sMaterialType)) _
			Or (sFoundMaterialName = "air_0")) Then
			' Do nothing if type is no match or if material is "air_0"
		Else
			' Determine material folder and go from there
			sFoundMaterialFolder = ""
			If (UBound(sSplitMaterialEntry) > 0) Then
				For j = 0 To UBound(sSplitMaterialEntry)-1
					sFoundMaterialFolder = sFoundMaterialFolder + sSplitMaterialEntry(j) + IIf(j=UBound(sSplitMaterialEntry),"","/")
				Next j
				' See if folder is already in list
				For j = 0 To UBound(sMaterialFolderList)
					If sFoundMaterialFolder = sMaterialFolderList(j) Then Exit For
				Next j
				If j = UBound(sMaterialFolderList)+1 Then ' loop above finished, a new folder was found
					ReDim Preserve sMaterialFolderList(UBound(sMaterialFolderList)+1)
					sMaterialFolderList(UBound(sMaterialFolderList)) = sFoundMaterialFolder
				End If
			End If

			' If the current folder matched sLocatedInMaterialFolder or if all folder are selected, add material to list
			If ((sLocatedInMaterialFolder = "CST_ALFOLDERS") _
				Or (sLocatedInMaterialFolder = sFoundMaterialFolder))Then
				ReDim Preserve sMaterialList(UBound(sMaterialList)+1)
				sMaterialList(UBound(sMaterialList)) = sFoundMaterialName
			End If

		End If
	Next

	FillMaterialArrays = Array(sMaterialFolderList,sMaterialList)

End Function

' --- The following functions are used to store and restore dialog settings ---

Function StoreAllDialogSettings(iniFileName As String, Optional sDenyList As String, Optional sAllowList As String)

	' Store values of each dialog item in a file
	' iniFileName: Full path to the file
	' sDenyList : comma separated list that can be used to exclude certain keys, use name of dialog item
	' sAllowList : comma separated list that can be used to restore only certain keys, use name of dialog item

	Dim i As Long
	Dim iniFile As Long
	Dim nDialogItems As Long
	Dim sDlgItem As String
	Dim sDenyListArray() As String, sAllowListArray() As String

	If (Len(sDenyList)>0 And Len(sAllowList)>0) Then
		ReportError("Store dialog settings: sDenyList and sAllowList are both non-empty.")
	End If

	nDialogItems = DlgCount()
	' Remove single spaces, just in case...
	sDenyList = Replace(sDenyList, " ", "")
	sAllowList = Replace(sAllowList, " ", "")
	sDenyListArray = Split(sDenyList, ",")
	sAllowListArray = Split(sAllowList, ",")

	If Dir(iniFileName) > "" Then
		Kill iniFileName
	End If
	iniFile = FreeFile
	Open iniFileName For Append As #iniFile
	Print #iniFile, "% Configuration/data file for tabulated surface impedance material, generated by 'Macros->Materials->Create Tabulated Surface Impedance Material'"
	Print #iniFile, ""
	For i = 0 To nDialogItems-1
		sDlgItem = DlgName(i)
		' Store sDlgItem if sDenyList AND sAllowList are empty, or if sDlgItem is not in sDenyList, or if sDlgItem is in sAllowList
		If ( (sDenyList = "") And (sAllowList = "") _
			Or ((sAllowList = "") And (FindListIndex(sDenyListArray,sDlgItem)=-1)) _
			Or ((sDenyList = "") And (FindListIndex(sAllowListArray,sDlgItem)>-1))) Then
				Select Case DlgType(i)
					Case "CheckBox", "ComboBox", "DropListBox", "ListBox", "MultiListBox", "OptionGroup"
						Print #iniFile,DlgName(i)+"="+CStr(DlgValue(i))
					Case "TextBox"
						Print #iniFile,DlgName(i)+"="+DlgText(i)
					Case "Text", "Picture", "GroupBox", "OKButton", "CancelButton", "PushButton", "OptionButton"
						' Do nothing
					Case Else
						MsgBox("Unknown DlgType: " + DlgType(i))
				End Select
		End If
	Next
	Close #iniFile

End Function

Function ReStoreAllDialogSettings(iniFileName As String, Optional sDenyList As String, Optional sAllowList As String)

	' Restore values of each dialog item in a file, if it exists
	' iniFileName: Full path to the file
	' sDenyList : comma separated list that can be used to exclude certain keys, use name of dialog item
	' sAllowList : comma separated list that can be used to restore only certain keys, use name of dialog item

	Dim i As Long
	Dim iniFile As Long
	Dim lineRead As String
	Dim sDlgItem As String, sValue As String
	Dim sDenyListArray() As String, sAllowListArray() As String
	Dim nDenyListLength As Long, nAllowListLength As Long
	Dim sToBeRestoredList As String ' list of dialog items to be restored

	If (Len(sDenyList)>0 And Len(sAllowList)>0) Then
		ReportError("Store dialog settings: sDenyList and sAllowList are both non-empty.")
	End If
	sValue = ""
	' Remove spaces, just in case...
	Do
		nDenyListLength = Len(sDenyList)
		sDenyList = Replace(sDenyList, " ", "")
	Loop While(nDenyListLength>Len(sDenyList))
	Do
		nAllowListLength = Len(sAllowList)
		sAllowList = Replace(sAllowList, " ", "")
	Loop While(nAllowListLength>Len(sAllowList))
	sDenyListArray = Split(sDenyList, ",")
	sAllowListArray = Split(sAllowList, ",")

	If (sAllowList<>"") Then
		' If allow list is used, this is the to-do list
		sToBeRestoredList = sAllowList + ","
	Else
		sToBeRestoredList = ""
		' If no allow list is used, to-do list consists of all storable entries, minus deny list
		For i = 0 To DlgCount()-1
			Select Case DlgName(i)
				Case "CheckBox", "ComboBox", "DropListBox", "ListBox", "MultiListBox", "OptionGroup", "TextBox"
					If (FindListIndex(sDenyListArray, DlgName(i))=-1) Then sToBeRestoredList = sToBeRestoredList + DlgName(i) + ","
				Case "Text", "Picture", "GroupBox", "OKButton", "CancelButton", "PushButton", "OptionButton"
					' Do nothing
			End Select
		Next
	End If

	iniFile = FreeFile
	If (Dir(iniFileName) = "") Then
		ReportWarningToWindow("Could not find data file, no dialog settings restored.")
		Exit Function
	End If
	' Else:
	Open iniFileName For Input As #iniFile
	While Not EOF(iniFile)
			Do
				Line Input #iniFile, lineRead
			Loop Until ((lineRead<>"") And (Left(lineRead,1)<>"%")) ' skip empty lines and comment lines
			sDlgItem = Split(lineRead, "=")(0)
			sValue = Split(lineRead, "=")(1)
			' Restore sDlgItem if sDenyList AND sAllowList are empty, or if sDlgItem is not in sDenyList, or if sDlgItem is in sAllowList
			If ( (sDenyList = "") And (sAllowList = "") _
				Or ((sAllowList = "") And (FindListIndex(sDenyListArray,sDlgItem)=-1)) _
				Or ((sDenyList = "") And (FindListIndex(sAllowListArray,sDlgItem)>-1))) Then
					On Error GoTo ReStoreError
						Select Case DlgType(sDlgItem)
							Case "CheckBox", "ComboBox", "DropListBox", "ListBox", "MultiListBox", "OptionGroup"
								DlgValue(sDlgItem, CLng(sValue))
								' Remove from to-do list
								sToBeRestoredList = Replace(sToBeRestoredList, sDlgItem + ",", "")
							Case "TextBox"
								DlgText(sDlgItem, sValue)
								' Remove from to-do list
								sToBeRestoredList = Replace(sToBeRestoredList, sDlgItem + ",", "")
							Case "Text", "Picture", "GroupBox", "OKButton", "CancelButton", "PushButton", "OptionButton"
								' Do nothing
							Case Else
								MsgBox("Unknown DlgType: " + DlgType(sDlgItem))
						End Select
					On Error GoTo 0
			End If
		BottomOfLoop:
	Wend
	Close #iniFile

	If (sToBeRestoredList<>"") Then
		ReportWarningToWindow("The following items could not be restored, because they were undefined in either the data file or the current dialog: "+Left(sToBeRestoredList, Len(sToBeRestoredList)-1)+". This can happen if the data file was saved with a different macro version. Please double check all dialog settings.")
	End If

	Exit Function

	ReStoreError:
		sToBeRestoredList = sToBeRestoredList + sDlgItem + ","
		GoTo BottomOfLoop

End Function

' --- End store/restore functions ---
