'#include "vba_globals_3d.lib"
'#include "vba_globals_all.lib"
'#include "coordinate_systems.lib"

' Create Maxwellian Particles
' This macro generates particle samples that are uniformly distributed in space and Maxwellian distributed in velocity space
' The samples are stored in a .pit file and can be loaded directly into CST PS as a particle interface
' Non-relativistic speeds are assumed, then gamma approx. 1, normed momentum approx. v/c
' ================================================================================================
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 18-Dec-2014 fsr: Fixed calculation of emission times
' 02-Jan-2014 fsr: bounding box (incl. by shape) is used to set up dialog for xmin...zmax; these settings are not stored anymore
' 27-Jun-2012 jwa: small modification to prevent memory overflow for large number of time steps
' 13-May-2011 fsr: a missing "\" prevented the macro from working for non-Admin users on Windows 7, fixed
' 21-Feb-2011 ube: adjusted online help link
' 26-Jan-2011 fsr: drift component is now stored relativistically (gamma>1); put in check for super high temperatures that would lead to relativistic speeds; minor improvements
' 30-Nov-2010 fsr: included new vba_globals_3d instead of vba_globals; bugfixes
' 19-Nov-2010 fsr: removed "simulation time step" option -> macro is now only compatible with 2011 and up, improved emission time step calculation
' 13-Nov-2010 fsr: minor improvements, sped up file writing (mainly over the network), added info file to result tree
' 18-Aug-2010 fsr: adjusted .pit output from current to macro charge - this was a change in 2011 without bwc! (macro still works for 2010 but files cannot be exchanged between 2010 and 2011)
' 19-Jul-2010 fsr: changed formatting so that a period is used as decimal separator, regardless of locale
' 03-Jun-2010 fsr: added option to create .pid file
' 24-May-2010 fsr: changed ini file name to include project name
' 23-Apr-2010 fsr: cosmetic changes
' 18-Mar-2010 fsr: replaced current input by time step input, added info comments about particle distribution to .pit file
' 19-Feb-2010 fsr: allowed for option to create particles in pairs,
'				moving in opposite directions to counter unphysical currents from creating a particle out of nothing
' 15-Jan-2010 fsr: added option to create particles inside selected solids
' 13-Jan-2010 fsr: added storing of parameters, use of expressions in text fields
' 12-Jan-2010 fsr: added cylindric and spherical shapes, super particle info, time controlled emission
' 07-Jan-2010 fsr: initial version

Const HelpFileName = "common_preloadedmacro_solver_maxwellian_particles"

Public Xmax As Double, Ymax As Double, Zmax As Double, Xmin As Double, Ymin As Double, Zmin As Double	' Volume limits for particles, cuboid
Public Rmin As Double, Rmax As Double																	' Volume limits for particles, cylinder
Public Volume As Double																					' Metric volume
Public VXdrift As Double, VYdrift As Double, VZdrift As Double											' Drift velocities for vx, vy, vz
Public VXSigma As Double, VYSigma As Double, VZSigma As Double											' Spread/Sigma for ux/uy/uz, the normed momentums
Public NSamples As Long, SPRatio As Double																' Number of samples (super particles) and s.p. ratio
Public USample1 As Double, USample2 As Double															' Two uniform temp samples
Public VXSamples() As Double, VYSamples() As Double, VZSamples() As Double								' Arrays for normal distribution samples
Public VMag1 As Double, VMag2 As Double																	' Temp variables
Public XYZSamples() As Double																			' Array for particle x/y/z positions
Public PMass As Double, ECurrent As Double, PTemp As Double												' Particle mass, emission current, and temperature of distribution function
Public ETime1 As Double, ETime2 As Double, ETimeStep As Double, NEmits As Long, ETime() As Double		' Emission start, stop, step width, number of steps, and vector containing time data
Public PCharge As Double, ZCharge As Integer															' Particle charge and ionization level
Public i As Long, j As Long, k As Long																	' Indexes for loops
Public tUnit As Double, lUnit As Double																	' Unit conversion variables
Public aSolidArray_CST() As String, nSolids_CST As Integer												' Array and count for selecting solids
Public outputFileName As String, infoFileName As String, tmpFileName As String
Public outputFile As Integer, infoFile As Integer, tmpFile As Integer
Public outputString As String
Public fileFormat As String
Public iniFileName As String
Public interfaceName As String
Public fixedSeed As Boolean			' Use a fixed seed for RND?
Public particlePairs As Boolean		' Create particles in same-kind pairs?
Public seedValue As Long

Public showFile As Boolean			' Open .pit file in notepad at the end?
Public importInterface As Boolean	' Import file into PS at the end?
Public cutoffMaxwellian As Boolean	' Truncate Maxwellian to avoid numerical issues caused by super fast particles?
Public cutoffFactor As Double

Public Const QElemental = 1.602176e-19
Public Const MElectron = 9.109381e-31
Public Const MProton = 1.672621e-27

Sub Main ()

	BeginHide

	nSolids_CST = 0

	iniFileName = GetProjectPath("Model3D")+"\CreatMaxwellianParticles_DialogSettings.ini"

	Dim PTypeArray(3) As String
	PTypeArray(1) = "Electrons"
	PTypeArray(2) = "Protons"
	PTypeArray(3) = "Other"

	Dim ShapeArray(4) As String
	ShapeArray(1) = "Cuboid"
	ShapeArray(2) = "Cylindric"
	ShapeArray(3) = "Spherical"
	ShapeArray(4) = "Solids in Cuboid:"

	Begin Dialog UserDialog 800,525,"Generate Maxwellian particles",.DialogFunc ' %GRID:10,7,1,1
		' Left Column
		GroupBox 10,7,380,504,"PDF settings",.GroupBox1
		GroupBox 400,7,390,154,"Boundary settings",.GroupBox2
		GroupBox 400,308,390,119,"Additional settings",.GroupBox3
		Text 30,32,120,21,"Particle type:",.Text1
		DropListBox 130,28,150,192,PTypeArray(),.PTypeDD
		Text 30,60,90,21,"Charge:",.Text2
		TextBox 130,56,30,21,.ZChargeT
		Text 161,60,10,21,"*",.Text3
		TextBox 170,56,110,21,.PChargeT
		Text 290,60,20,21,"C",.Text4
		Text 30,88,90,14,"Mass:",.Text5
		TextBox 130,84,150,21,.PMassT
		Text 290,88,20,21,"kg",.Text6
		Text 30,116,90,14,"Temperature:",.Text7
		TextBox 130,112,150,21,.PTempT
		Text 290,116,20,21,"eV",.Text8
		Text 200,252,10,21,"...",.Text25
		CheckBox 50,140,260,14,"Truncate distribution function above",.cutoffCB
		Text 72,158,20,14,"E=",.Text28
		TextBox 100,154,60,21,.cutoffFactorT
		Text 165,158,90,21,"* Temperature",.Text9
		Text 30,196,80,14,"File format:",.Text30
		OptionGroup .FileFormatRB
			OptionButton 170,196,50,14,"PIT",.OptionButton1
			OptionButton 230,196,50,14,"PID",.OptionButton2
		Text 30,224,130,14,"Emission current:",.Text14
		TextBox 170,217,110,21,.ECurrentT
		Text 290,224,90,21,"A",.Text15
		Text 30,252,90,14,"Emission time:",.Text16
		TextBox 130,245,60,21,.ETime1T
		TextBox 220,245,60,21,.ETime2T
		Text 290,252,90,21,Units.GetUnit("Time"),.Text17
		Text 30,280,180,14,"Number of emission steps:",.Text29
		TextBox 220,273,60,21,.NEmitsT
		CheckBox 30,308,260,14,"Create particles in pairs of same kind",.PPairCB
		Text 30,336,170,21,"Drift velocity in terms of c:",.Text10
		Text 30,364,90,21,"vx/c:",.Text11
		TextBox 130,364,150,21,.vxT
		Text 30,392,90,21,"vy/c:",.Text12
		TextBox 130,392,150,21,.vyT
		Text 30,420,90,21,"vz/c:",.Text13
		TextBox 130,420,150,21,.vzT
		Text 30,455,170,14,"Number of macro particles:",.Text18
		TextBox 210,448,70,21,.NSamplesT
		Text 30,483,130,14,"Macro particle ratio:",.Text20
		TextBox 210,476,70,21,.SPRatioT
		' Right column
		Text 420,32,50,14,"Shape:",.Text19
		DropListBox 480,28,150,192,ShapeArray(),.ShapeDD
		PushButton 640,28,90,21,"Solids",.SolidsB
		Text 430,60,40,21,"xmin:",.xminTT
		TextBox 480,56,90,21,.xminT
		Text 590,60,40,21,"xmax:",.xmaxTT
		TextBox 640,56,90,21,.xmaxT
		Text 430,88,40,14,"ymin:",.yminTT
		TextBox 480,84,90,21,.yminT
		Text 590,88,40,21,"ymax:",.ymaxTT
		TextBox 640,84,90,21,.ymaxT
		Text 430,116,40,21,"zmin:",.zminTT
		TextBox 480,112,90,21,.zminT
		Text 590,116,40,21,"zmax:",.zmaxTT
		TextBox 640,112,90,21,.zmaxT
		Text 420,140,340,14,"Values given in project units [" + Units.GetUnit("Length") + "].",.Text27
		GroupBox 400,168,390,133,"Particle info",.GroupBox4
		Text 420,193,140,14,"Phys. particle density:",.PdensityTT
		Text 570,193,120,14,"n/d",.PdensityT,1
		Text 705,193,50,21,"1/m^3",.Text21
		Text 420,221,150,14,"Macro particle density:",.SPdensityTT
		Text 570,221,120,14,"n/d",.SPdensityT,1
		Text 705,221,50,21,"1/m^3",.Text22
		Text 420,249,150,14,"Macro particle charge:",.SPchargeTT
		Text 570,249,120,14,"n/d",.SPchargeT,1
		Text 705,249,20,21,"C",.Text23
		Text 420,277,140,14,"Macro particle mass:",.SPmassTT
		Text 570,277,120,14,"n/d",.SPmassT,1
		Text 705,277,20,21,"kg",.Text24
		CheckBox 420,333,240,14,"Import into PS as particle interface",.importCB
		Text 440,351,100,14,"Interface name:",.Text26
		TextBox 550,347,180,21,.nameT
		CheckBox 420,371,220,14,"Show input file after generation",.showFileCB
		CheckBox 420,399,220,14,"Use fixed seed value for RND():",.fixedSeedCB
		TextBox 650,395,80,21,.seedValueT
		GroupBox 400,434,390,49,"Output:",.GroupBox5
		Text 410,455,370,14,"",.OutputT
		OKButton 400,490,90,21
		CancelButton 600,490,90,21
		PushButton 700,490,90,21,"Help",.Help
		PushButton 500,490,90,21,"Apply",.Apply
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All

	EndHide

End Sub

Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Select Case Action%
	Case 1 ' Dialog box initialization
		InitDialog()
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "PTypeDD"
				Select Case SuppValue
					Case 0
						DlgText "PChargeT", cstr(-QElemental)
						DlgText "PMassT", cstr(MElectron)
						DlgText "ZChargeT", "1"
						DlgEnable "ZChargeT", False
						DlgText "nameT", "Electrons"
					Case 1
						DlgText "PChargeT", cstr(QElemental)
						DlgText "PMassT", cstr(MProton)
						DlgText "ZChargeT", "1"
						DlgEnable "ZChargeT", False
						DlgText "nameT", "Protons"
					Case 2
						DlgText "PChargeT", cstr(QElemental)
						DlgText "PMassT", cstr(MProton)
						DlgEnable "ZChargeT", True
						DlgText "nameT", "MyParticles"
				End Select
				UpdateSPinfo()
			Case "cutoffCB"
					DlgEnable "cutoffFactorT",CBool(SuppValue)
			Case "fixedSeedCB"
					DlgEnable "seedValueT",CBool(SuppValue)
			Case "PPairCB"
					DlgEnable "vxT",Not CBool(DlgValue("PPairCB"))
					DlgEnable "vyT",Not CBool(DlgValue("PPairCB"))
					DlgEnable "vzT",Not CBool(DlgValue("PPairCB"))
					DlgText "vxT","0"
					DlgText "vyT","0"
					DlgText "vzT","0"
			Case "ShapeDD"	' What shape?
				Select Case SuppValue
					Case 0	' Cuboid
						DlgText "xminTT","xmin:"
						DlgText "xmaxTT","xmax:"
						DlgEnable "SolidsB", False
						DlgEnable "xminT",True
						DlgEnable "xmaxT",True
						DlgEnable "yminT",True
						DlgEnable "ymaxT",True
						DlgEnable "zminT",True
						DlgEnable "zmaxT",True
					Case 1	' Cylindric
						DlgText "xminTT","Rmin:"
						DlgText "xmaxTT","Rmax:"
						DlgEnable "SolidsB", False
						DlgEnable "xminT",True
						DlgEnable "xmaxT",True
						DlgEnable "yminT",False
						DlgEnable "ymaxT",False
						DlgEnable "zminT",True
						DlgEnable "zmaxT",True
					Case 2	' Spherical
						DlgText "xminTT","Rmin:"
						DlgText "xmaxTT","Rmax:"
						DlgEnable "SolidsB", False
						DlgEnable "xminT",True
						DlgEnable "xmaxT",True
						DlgEnable "yminT",False
						DlgEnable "ymaxT",False
						DlgEnable "zminT",False
						DlgEnable "zmaxT",False
					Case 3	' Selected solid, requires xmin-zmax settings, too!
						DlgText "xminTT","xmin:"
						DlgText "xmaxTT","xmax:"
						DlgEnable "SolidsB", True
						DlgEnable "xminT",True
						DlgEnable "xmaxT",True
						DlgEnable "yminT",True
						DlgEnable "ymaxT",True
						DlgEnable "zminT",True
						DlgEnable "zmaxT",True
				End Select
				UpdateSPinfo()
			Case "importCB"
						DlgEnable "nameT",CBool(SuppValue)
			Case "SolidsB"
				SelectSolids_LIB(aSolidArray_CST(), nSolids_CST)
				DialogFunc = True       ' Don't close the dialog box.
				UpdateBoundingBoxInfo()
				UpdateSPinfo()
			Case "FileFormatRB"
				DialogFunc = True
				Select Case SuppValue
					Case 0 ' PIT file
						fileFormat = "pit"
						DlgEnable "ECurrentT", False
						DlgEnable "ETime1T", True
						DlgEnable "ETime2T", True
						DlgEnable "NEmitsT", True
						DlgEnable "SPRatioT", True
					Case 1 ' PID file
						fileFormat = "pid"
						DlgEnable "ECurrentT", True
						DlgEnable "ETime1T", False
						DlgEnable "ETime2T", False
						DlgEnable "NEmitsT", False
						DlgEnable "SPRatioT", False
				End Select
			Case "Help"
				StartHelp HelpFileName
				DialogFunc = True
			Case "stopit"
				Exit All
			Case "Apply"
				RunApplyOK()
				DialogFunc = True
			Case "OK"
				' If RunApplyOK returns an error (False), keep dialog window alive
				DialogFunc = Not RunApplyOK()
		End Select
	Case 3
		UpdateSPinfo()
		DialogFunc = True
	End Select
End Function

Function InitDialog()

	Dim i As Long, j As Long
	Dim nDeletedSolids As Long

	DlgText "ZChargeT","1"
	DlgEnable "ZChargeT", False
	DlgText "PChargeT", cstr(-QElemental)
	DlgText "PMassT", cstr(MElectron)
	DlgText "PTempT", IIf(ReStoreDialogSetting("PTemp")<>"",ReStoreDialogSetting("PTemp"),"1")
	DlgValue "cutoffCB",CBool(IIf(ReStoreDialogSetting("cutoffCB")<>"",ReStoreDialogSetting("cutoffCB"),"True"))
	DlgText "cutoffFactorT",IIf(ReStoreDialogSetting("cutoffFactor")<>"",ReStoreDialogSetting("cutoffFactor"),"100")
	DlgEnable "cutoffFactorT",CBool(IIf(ReStoreDialogSetting("cutoffCB")<>"",ReStoreDialogSetting("cutoffCB"),"True"))
	DlgValue "PPairCB",CBool(IIf(ReStoreDialogSetting("PPairCB")<>"",ReStoreDialogSetting("PPairCB"),"True"))
	DlgEnable "vxT",Not CBool(DlgValue("PPairCB"))
	DlgEnable "vyT",Not CBool(DlgValue("PPairCB"))
	DlgEnable "vzT",Not CBool(DlgValue("PPairCB"))
	DlgText "vxT",IIf(Not CBool(DlgValue("PPairCB")) Or ReStoreDialogSetting("vx")<>"",ReStoreDialogSetting("vx"),"0")
	DlgText "vyT",IIf(Not CBool(DlgValue("PPairCB")) Or ReStoreDialogSetting("vy")<>"",ReStoreDialogSetting("vy"),"0")
	DlgText "vzT",IIf(Not CBool(DlgValue("PPairCB")) Or ReStoreDialogSetting("vz")<>"",ReStoreDialogSetting("vz"),"0")
	DlgValue "FileFormatRB", IIf(ReStoreDialogSetting("FileFormat")<>"1",0,1)
	DlgEnable "ECurrentT", IIf(ReStoreDialogSetting("FileFormat")="1",True,False)
	DlgEnable "ETime1T", IIf(ReStoreDialogSetting("FileFormat")="1",False,True)
	DlgEnable "ETime2T", IIf(ReStoreDialogSetting("FileFormat")="1",False,True)
	DlgEnable "NEmitsT", IIf(ReStoreDialogSetting("FileFormat")="1",False,True)
	DlgText "ECurrentT",IIf(ReStoreDialogSetting("ECurrent")<>"",ReStoreDialogSetting("ECurrent"),"1e-6")
	DlgText "ETime1T",IIf(ReStoreDialogSetting("ETime1")<>"",ReStoreDialogSetting("ETime1"),".01")
	DlgText "ETime2T",IIf(ReStoreDialogSetting("ETime2")<>"",ReStoreDialogSetting("ETime2"),".01")
	DlgText "NEmitsT",IIf(ReStoreDialogSetting("NEmits")<>"",ReStoreDialogSetting("NEmits"),"1")
	DlgText "NSamplesT",IIf(ReStoreDialogSetting("NSamples")<>"",ReStoreDialogSetting("NSamples"),"1000")
	DlgText "SPRatioT",IIf(ReStoreDialogSetting("SPRatio")<>"",ReStoreDialogSetting("SPRatio"),"1")
	DlgValue "fixedSeedCB",CBool(IIf(ReStoreDialogSetting("fixedSeedCB")<>"",ReStoreDialogSetting("fixedSeedCB"),"False"))
	DlgText "seedValueT",IIf(ReStoreDialogSetting("seedValue")<>"",ReStoreDialogSetting("seedValue"),"0")
	DlgEnable "seedValueT",CBool(IIf(ReStoreDialogSetting("fixedSeedCB")<>"",ReStoreDialogSetting("fixedSeedCB"),"False"))
	DlgValue "shapeDD", IIf(ReStoreDialogSetting("shapeDD")<>"",ReStoreDialogSetting("shapeDD"),0)
	Select Case DlgValue("shapeDD")
		Case 0	' Cuboid
			DlgText "xminTT","xmin:"
			DlgText "xmaxTT","xmax:"
			DlgEnable "SolidsB", False
			DlgEnable "xminT",True
			DlgEnable "xmaxT",True
			DlgEnable "yminT",True
			DlgEnable "ymaxT",True
			DlgEnable "zminT",True
			DlgEnable "zmaxT",True
		Case 1	' Cylindric
			DlgText "xminTT","Rmin:"
			DlgText "xmaxTT","Rmax:"
			DlgEnable "SolidsB", False
			DlgEnable "xminT",True
			DlgEnable "xmaxT",True
			DlgEnable "yminT",False
			DlgEnable "ymaxT",False
			DlgEnable "zminT",True
			DlgEnable "zmaxT",True
		Case 2	' Spherical
			DlgText "xminTT","Rmin:"
			DlgText "xmaxTT","Rmax:"
			DlgEnable "SolidsB", False
			DlgEnable "xminT",True
			DlgEnable "xmaxT",True
			DlgEnable "yminT",False
			DlgEnable "ymaxT",False
			DlgEnable "zminT",False
			DlgEnable "zmaxT",False
		Case 3	' Selected solid, requires xmin-zmax settings, too!
			DlgText "xminTT","xmin:"
			DlgText "xmaxTT","xmax:"
			DlgEnable "SolidsB", True
			DlgEnable "xminT",True
			DlgEnable "xmaxT",True
			DlgEnable "yminT",True
			DlgEnable "ymaxT",True
			DlgEnable "zminT",True
			DlgEnable "zmaxT",True
	End Select
	DlgValue "importCB",CBool(IIf(ReStoreDialogSetting("importCB")<>"",ReStoreDialogSetting("importCB"),"True"))
	DlgText "nameT","Electrons"
	DlgEnable "nameT",CBool(IIf(ReStoreDialogSetting("importCB")<>"",ReStoreDialogSetting("importCB"),"True"))
	DlgValue "showFileCB",CBool(IIf(ReStoreDialogSetting("showFileCB")<>"",ReStoreDialogSetting("showFileCB"),"False"))

	nSolids_CST = CInt(ReStoreDialogSetting("nSelectedSolids","0"))
	DeletedSolids = 0
	ReDim aSolidArray_CST(nSolids_CST)
	For i = 0 To nSolids_CST-1
		' Verify that the previously saved shapes still exist
		For j = 0 To Solid.GetNumberOfShapes
			If (Solid.GetNameOfShapeFromIndex(j) = ReStoreDialogSetting("SelectedSolid_"+CStr(i))) Then
				aSolidArray_CST(i-DeletedSolids) = ReStoreDialogSetting("SelectedSolid_"+CStr(i))
				Exit For
			End If
		Next
		If (j = Solid.GetNumberOfShapes + 1) Then ' the previously saved shape could not be found anymore
			DeletedSolids = DeletedSolids + 1
		End If
	Next
	nSolids_CST = nSolids_CST - DeletedSolids
	ReDim Preserve aSolidArray_CST(nSolids_CST)

	UpdateBoundingBoxInfo()

	With Units
		tUnit = .GetTimeUnitToSI
		lUnit = .GetGeometryUnitToSI
	End With
	UpdateSPinfo()

End Function

' RunApplyOK disables the buttons, reads in the dialog values and runs a check on the number
' of particles. If the file/interface was (not) created, True (False) is returned.
Function RunApplyOK() As Boolean

	DlgEnable "OK", False
	DlgEnable "Apply", False
	DlgEnable "Cancel", False
	DlgEnable "Help", False

	If Not ReadDialogValues() Then
		RunApplyOK = False
		GoTo ExitRunApplyOK
	End If

	If NSamples >= 1000000 Then
		Select Case MsgBox("A very large number of particles has been chosen." + vbNewLine _
							+ "The simulation will require large amounts of memory "+ vbNewLine _
							+ "and most likely be slow. Please consider fewer"+ vbNewLine _
							+ "particles with a higher macro particle ratio instead."+ vbNewLine _
							+ "Do you want to continue anyways?",vbYesNo,"Check Settings")
		Case 6 'Yes
			StoreAllSettings()
			RunApplyOK = GeneratePIFile() ' Forward any errors from GeneratePIFile
		Case 7 'No
			RunApplyOK = False
			GoTo ExitRunApplyOK
		End Select
	ElseIf NSamples >= 25000000 Then
		MsgBox("Too many samples! Please choose a number below 25,000,000")
		RunApplyOK = False
		GoTo ExitRunApplyOK
	Else
		StoreAllSettings()
		RunApplyOK = GeneratePIFile() ' Forward any errors from GeneratePIFile
	End If

	GoTo ExitRunApplyOK
	ExitRunApplyOK:
	DlgEnable "OK", True
	DlgEnable "Apply", True
	DlgEnable "Cancel", True
	DlgEnable "Help", True

End Function

Function ReadDialogValues() As Boolean

	showFile = DlgValue("showFileCB")
	importInterface = DlgValue("importCB")
	interfaceName=DlgText("nameT")
	cutoffMaxwellian = DlgValue("cutoffCB")
	cutoffFactor = Evaluate(DlgText("cutoffFactorT"))
	fixedSeed = DlgValue("fixedSeedCB")
	seedValue = Evaluate(DlgText("seedValueT"))
	ZCharge = Evaluate(DlgText("ZChargeT"))
	PCharge = Evaluate(DlgText("PChargeT"))
	PMass = Evaluate(DlgText("PMassT"))
	PTemp = Evaluate(DlgText("PTempT"))
	' Project units for time
	ETime1 = Evaluate(DlgText("ETime1T"))
	ETime2 = Evaluate(DlgText("ETime2T"))
	NEmits = Evaluate(DlgText("NEmitsT"))
	' Project units for positions
	Xmin = Evaluate(DlgText("xminT"))
	Xmax = Evaluate(DlgText("xmaxT"))
	Rmin = Evaluate(DlgText("xminT"))
	Rmax = Evaluate(DlgText("xmaxT"))
	Ymin = Evaluate(DlgText("yminT"))
	Ymax = Evaluate(DlgText("ymaxT"))
	Zmin = Evaluate(DlgText("zminT"))
	Zmax = Evaluate(DlgText("zmaxT"))
	particlePairs = DlgValue("PPairCB")
	VXdrift = Evaluate(DlgText("vxT"))
	VYdrift = Evaluate(DlgText("vyT"))
	VZdrift = Evaluate(DlgText("vzT"))
	If ((VXdrift^2 + VYdrift^2 + VZdrift^2) >= 1) Then
		MsgBox("Absolute drift velocity must be smaller than CLight, please check your settings.", "Error")
		ReadDialogValues = False
		Exit Function
	End If

	NSamples=Evaluate(DlgText("NSamplesT"))
	SPRatio = Evaluate(DlgText("SPRatioT"))
	' Quick check if temperature leads to relativistic speeds
	' Criterion: mean of Abs(v) is larger than c/100
	If (Sqr(8*PTemp*QElemental/PMass/Pi)>CLight/100) Then
		MsgBox("Average of absolute thermal velocity is larger than 1% of CLight. This may lead to inaccurate results.","Warning")
	End If
	VXSigma = Sqr(PTemp*QElemental/PMass)
	VYSigma = VXSigma
	VZSigma = VXSigma

	fileFormat = IIf(DlgValue("FileFormatRB")=0,"pit","pid")
	ECurrent = Evaluate(DlgText("ECurrentT"))

	If (NEmits > NSamples And Not particlePairs) Then
		ReportInformationToWindow("Number of emissions larger than number of particles. Limiting number of emissions to number of particles.")
		NEmits = NSamples
		ETimeStep = (ETime2-ETime1)/(NEmits-1)
	ElseIf (NEmits > NSamples And particlePairs) Then
		ReportInformationToWindow("Number of emissions larger than number of particles and particle pairs active. Limiting number of emissions to half the number of particles.")
		NEmits = Int((NSamples+1)/2)
		ETimeStep = (ETime2-ETime1)/(NEmits-1)
	ElseIf (NEmits > 1) Then
		ETimeStep = (ETime2-ETime1)/(NEmits-1)
	Else
		ETimeStep = 0
	End If

	' Do some checks on the input parameters
	Select Case DlgValue("ShapeDD")
		Case 0 ' Cuboid
				If(Xmax<Xmin Or Ymax<Ymin Or Zmax<Zmin) Then
					MsgBox("Check boundary settings: max value must not be smaller than min value!","Check Settings")
					ReadDialogValues = False
					Exit Function
				ElseIf (Xmax=Xmin And Ymax=Ymin And Zmax=Zmin) Then
					MsgBox("Check boundary settings: range must be >0 in at least one dimension.","Check Settings")
					ReadDialogValues = False
					Exit Function
				End If
		Case 1 ' Cylindric
				If(Rmax<Rmin Or Rmin <0 Or Zmax<Zmin) Then
					MsgBox("Check boundary settings: max value must not be smaller than min value, radii have to be positive.","Check Settings")
					ReadDialogValues = False
					Exit Function
				End If
		Case 2 ' Spherical
				If(Rmax<Rmin Or Rmin <0) Then
					MsgBox("Check boundary settings: Rmax must not be smaller than Rmin, radii have to be positive.","Check Settings")
					ReadDialogValues = False
					Exit Function
				End If
		Case 3 ' Selected solids in cuboid
				' No checks yet
	End Select

	' Success!
	ReadDialogValues = True

End Function

Function GeneratePIFile() As Boolean

	outputFileName = GetFilePath(interfaceName+"."+fileFormat,fileFormat,GetProjectPath("Model3D"),"Enter file name",3)
	If (outputFileName = "") Then ' user pressed "Cancel"
		GeneratePIFile = Fase
		Exit Function
	End If

	' If particles to be created in pairs, half the number of base samples
	If particlePairs Then
		NSamples = Int((NSamples+1)/2)
	End If

	GeneratePIFile = False	' File has not been created yet

	' When a fixed seed is used, the chain of pseudo rnds will always be the same for a given seedValue
	If fixedSeed Then
		Randomize seedValue
	End If

	' Create emission time data
	ReDim ETime(NSamples)
	For i = 1 To NSamples STEP 1
		ETime(i) = ETime1 + ETimeStep*Fix(NEmits/NSamples*(i-1)) ' Subtract a small number to make VBA round properly
		' If (ETime(i) <> ETime(i-1)) Then ReportInformationToWindow(ETime(i))
	Next i

	' Create particle positions - uniformly distributed within the volume
	ReDim XYZSamples(3*NSamples)
	k = 0
	Select Case DlgValue("ShapeDD")
		Case 0 ' Cuboid
				For i=0 To NSamples-1 STEP 1
					If (((NSamples-1-i) Mod 500) = 0) Then DlgText("OutputT","Generating particle data, step 1: "+USFormat(i/(NSamples-1)*100,"000.00")+"%")
					XYZSamples(3*i+1) = Xmin + (Xmax-Xmin)*Rnd()
					XYZSamples(3*i+2) = Ymin + (Ymax-Ymin)*Rnd()
					XYZSamples(3*i+3) = Zmin + (Zmax-Zmin)*Rnd()
				Next i
		Case 1 ' Cylindric
				For i=0 To NSamples-1 STEP 1
					If (((NSamples-1-i) Mod 500) = 0) Then DlgText("OutputT","Generating particle data, step 1: "+USFormat(i/(NSamples-1)*100,"000.00")+"%")
					XYZSamples(3*i+1) = 2*Rmax*Rnd()-Rmax
					XYZSamples(3*i+2) = 2*Rmax*Rnd()-Rmax
					XYZSamples(3*i+3) = Zmin + (Zmax-Zmin)*Rnd()
					' If Rnd outside cylindric shape, reject
					Mag1 = XYZSamples(3*i+1)^2 + XYZSamples(3*i+2)^2
					If (Mag1 >= Rmax^2 Or Mag1 <= Rmin^2) Then
						i = i-1
						k = k+1
						If (k = NSamples) Then
							Select Case	MsgBox(CStr(NSamples)+" consecutive samples have been rejected" + vbNewLine _
											+"and "+CStr(i+1)+" samples in total have been accepted." + vbNewLine _
											+"Something might be wrong with the volume settings." + vbNewLine _
											+"Do you want to continue?", vbYesNo,"Check Settings")
							Case 6 ' Yes
								k = 0
							Case 7 ' No
								GeneratePIFile = False
								Exit Function
							End Select
						End If
					Else
						k = 0
					End If
				Next i
		Case 2 ' Spherical
				For i=0 To NSamples-1 STEP 1
					If (((NSamples-1-i) Mod 500) = 0) Then DlgText("OutputT","Generating particle data, step 1: "+USFormat(i/(NSamples-1)*100,"000.00")+"%")
					XYZSamples(3*i+1) = 2*Rmax*Rnd()-Rmax
					XYZSamples(3*i+2) = 2*Rmax*Rnd()-Rmax
					XYZSamples(3*i+3) = 2*Rmax*Rnd()-Rmax
					' If Rnd outside cylindric shape, reject
					Mag1 = XYZSamples(3*i+1)^2 + XYZSamples(3*i+2)^2 + XYZSamples(3*i+3)^2
					If (Mag1 >= Rmax^2 Or Mag1 <= Rmin^2) Then
						i = i-1
						k = k+1
						If (k = NSamples) Then
							Select Case	MsgBox(Str(NSamples)+" consecutive samples have been rejected" + vbNewLine _
											+"and "+Str(i+1)+" samples in total have been accepted." + vbNewLine _
											+"Something might be wrong with the volume settings." + vbNewLine _
											+"Do you want to continue?", vbYesNo,"Check Settings")
							Case 6 ' Yes
								k = 0
							Case 7 ' No
								GeneratePIFile = False
								Exit Function
							End Select
						End If
					Else
						k = 0
					End If
				Next i
		Case 3 ' Selected solids in cuboid
				Dim isInside As Boolean
				For i=0 To NSamples-1 STEP 1
					If (((NSamples-1-i) Mod 500) = 0) Then DlgText("OutputT","Generating particle data, step 1: "+USFormat(i/(NSamples-1)*100,"000.00")+"%")
					isInside = False
					XYZSamples(3*i+1) = Xmin + (Xmax-Xmin)*Rnd()
					XYZSamples(3*i+2) = Ymin + (Ymax-Ymin)*Rnd()
					XYZSamples(3*i+3) = Zmin + (Zmax-Zmin)*Rnd()
					For j=1 To nSolids_CST STEP 1
						If (aSolidArray_CST(j-1) > "") And (Solid.IsPointInsideShape(XYZSamples(3*i+1), XYZSamples(3*i+2), XYZSamples(3*i+3), aSolidArray_CST(j-1))) Then
							isInside = True
							Exit For
						End If
					Next
					If Not isInside Then
						i = i-1
						k = k+1
						If (k = NSamples) Then
							Select Case	MsgBox(Str(NSamples)+" consecutive samples have been rejected" + vbNewLine _
											+"and "+Str(i+1)+" samples in total have been accepted." + vbNewLine _
											+"Something might be wrong with the volume settings." + vbNewLine _
											+"Do you want to continue?", vbYesNo,"Check Settings")
							Case 6 ' Yes
								k = 0
							Case 7 ' No
								GeneratePIFile = False
								Exit Function
							End Select
						End If
					Else
						k = 0
					End If
				Next i
	End Select


	' The normal distribution is created from a uniform distribution via a Box-Muller transform
	' This implies that always an even number of samples is transformed, so round up to next highest even number
	NSamplesEven=2*Round((NSamples+1)/2)
	ReDim VXSamples(NSamplesEven)
	ReDim VYSamples(NSamplesEven)
	ReDim VZSamples(NSamplesEven)

	For i=0 To NSamplesEven/2-1 STEP 1
		If (((NSamples/2-1-i) Mod 500) = 0) Then DlgText("OutputT","Generating particle data, step 2: "+USFormat(i/(NSamples/2-1)*100,"000.00")+"%")
		' 0<=Rnd()< 1, so it is necessary to use 1-Rnd() to avoid problems with Log(0)
		USample1=1-Rnd()
		USample2=1-Rnd()
		' Box-Muller transform
		VXSamples(2*i+1)=Sqr(-2*Log(USample1))*Cos(2*pi*USample2)*VXSigma
		VXSamples(2*i+2)=Sqr(-2*Log(USample1))*Sin(2*pi*USample2)*VXSigma

		USample1=1-Rnd()
		USample2=1-Rnd()
		' Box-Muller transform
		VYSamples(2*i+1)=Sqr(-2*Log(USample1))*Cos(2*pi*USample2)*VYSigma
		VYSamples(2*i+2)=Sqr(-2*Log(USample1))*Sin(2*pi*USample2)*VYSigma

		USample1=1-Rnd()
		USample2=1-Rnd()
		' Box-Muller transform
		VZSamples(2*i+1)=Sqr(-2*Log(USample1))*Cos(2*pi*USample2)*VZSigma
		VZSamples(2*i+2)=Sqr(-2*Log(USample1))*Sin(2*pi*USample2)*VZSigma

		' Calculate magnitude of velocity vectors. If one is too fast (Gaussian cutoff), reject both values
		' This is a little slower but should not happen too often and keeps the code short
		Mag1 = VXSamples(2*i+1)^2+VYSamples(2*i+1)^2+VZSamples(2*i+1)^2
		Mag2 = VXSamples(2*i+2)^2+VYSamples(2*i+2)^2+VZSamples(2*i+2)^2
		If (cutoffMaxwellian And (PMass*Mag1/2 > cutoffFactor*PTemp*QElemental Or PMass*Mag2/2 > cutoffFactor*PTemp*QElemental)) Then
			i=i-1
		End If
	Next i

	' Values in file are stored in SI units, non-relativistic speeds are assumed (relativistic drift speeds are ok)
	j=1
	outputString = ""
	outputString = outputString + "% Created with "+GetApplicationVersion + vbNewLine
	outputString = outputString + "% Number of samples/macro particles: "+IIf(particlePairs=True,CStr(2*NSamples),CStr(NSamples))+ vbNewLine
	outputString = outputString + "% 1 macro particle represents "+CStr(SPRatio)+" regular particle(s)."+ vbNewLine
	If (fileFormat="pit") Then
		outputString = outputString + "% Particles are created between "+USFormat(ETime1*tUnit,"Scientific")+" and "+USFormat(ETime2*tUnit,"Scientific")+" s in "+CStr(NEmits)+" step(s). "+ vbNewLine
	End If
	outputString = outputString + "%"+ vbNewLine
	outputString = outputString + "% Physical particle charge		: "+USFormat(ZCharge*PCharge,"Scientific")+" C" + vbNewLine
	outputString = outputString + "% Macro particle charge			: "+USFormat(ZCharge*PCharge*SPRatio,"Scientific")+" C" + vbNewLine
	outputString = outputString + "% Physical particle mass		: "+USFormat(PMass,"Scientific")+" kg" + vbNewLine
	outputString = outputString + "% Macro particle mass			: "+USFormat(PMass*SPRatio,"Scientific")+" kg" + vbNewLine
	outputString = outputString + "% Ionization degree (for ions)	: "+USFormat(ZCharge,"Scientific") + vbNewLine
	If (Volume > 0) Then
		outputString = outputString + "% Physical particle density		: "+USFormat(IIf(particlePairs=True,CStr(2*NSamples*SPRatio/Volume),CStr(2*NSamples*SPRatio/Volume)), "Scientific")+" m^-3" + vbNewLine
		outputString = outputString + "% Macro particle density		: "+USFormat(IIf(particlePairs=True,CStr(2*NSamples/Volume),CStr(2*NSamples/Volume)), "Scientific")+" m^-3" + vbNewLine
	End If
	outputString = outputString + "% Temp. of phys. particles		: "+USFormat(PTemp,"Scientific") +" eV" + vbNewLine
	outputString = outputString + "% Temp. of macro particles		: "+USFormat(SPRatio*PTemp,"Scientific") +" eV" + vbNewLine
	If (DlgValue("PTypeDD")=0 And Volume > 0) Then ' If electrons
		outputString = outputString + "% Phys. Debye length		: "+USFormat(IIf(particlePairs=True,CStr(Sqr(Eps0*PTemp/(2*NSamples*SPRatio*QElemental/Volume))),CStr(Sqr(Eps0*PTemp/(NSamples*SPRatio*QElemental/Volume)))),"Scientific") +" m" + vbNewLine
		outputString = outputString + "% Macro part. Debye length	: "+USFormat(IIf(particlePairs=True,CStr(Sqr(Eps0*PTemp*SPRatio/(2*NSamples*QElemental/Volume))),CStr(Sqr(Eps0*PTemp*SPRatio/(NSamples*QElemental/Volume)))),"Scientific") +" m" + vbNewLine
	End If
	outputString = outputString + "% Drift velocity in x/y/x	: "+USFormat(VXdrift,"Scientific")+"/"+USFormat(VYdrift,"Scientific")+"/"+USFormat(VZdrift,"Scientific")+" c" + vbNewLine

	infoFile = FreeFile
	infoFileName = GetProjectPath("Model3D")+interfaceName+"_ParticleInterfaceInfo.txt"
	If infoFileName<>"" Then
		Open infoFileName For Output As #infoFile
		Print #infoFile, outputString
		Close #infoFile
		Resulttree.UpdateTree
		Else
			MsgBox("Invalid file name, could not write info file.","Error")
	End If

	tmpFile = FreeFile
	tmpFileName = GetInstallPath()+"\CreateGaussianParticles.tmp" ' write in installation path as this is most likely a local folder
	If tmpFileName<>"" Then
		Open tmpFileName For Output As #tmpFile
		Else
			MsgBox("Invalid file name!","Error")
			Exit Function
	End If
	Print #tmpFile, outputString
	outputString = ""
	outputString = outputString + "%" + vbNewLine
	outputString = outputString + "% Format: pos_x  pos_y  pos_z  mom_x  mom_y  mom_z  mass  charge  "+IIf(fileFormat="pit","macro-charge  Time","") + vbNewLine	'% Line number: "+Str(j)
	outputString = outputString + "%" + vbNewLine

	' Dim time0 As Double, time1 As Double
	Dim flushsize As Long
	flushsize = 10 '4*IIf(particlePairs,1,2)	' a screw to turn... this value has been empirically found to be optimal, gain is very little
	' time0 = Timer()
	For i=0 To NSamples-1 STEP 1
		j+=1
		If (((NSamples-1-i) Mod 500) = 0) Then
			DlgText("OutputT","Writing interface file: "+USFormat(i/(NSamples-1)*100,"000.00")+"%")
		End If
		outputString = outputString + USFormat(XYZSamples(3*i+1)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(XYZSamples(3*i+2)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(XYZSamples(3*i+3)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(VXdrift/Sqr(1-VXdrift^2)+VXSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(VYdrift/Sqr(1-VYdrift^2)+VYSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(VZdrift/Sqr(1-VZdrift^2)+VZSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(PMass,"0.000000E+00")+" " _
						+USFormat(ZCharge*PCharge," 0.000000E+00;-0.000000E+00")+" " _
						+USFormat(IIf(fileFormat="pit",ZCharge*PCharge*SPRatio,ECurrent)," 0.000000E+00;-0.000000E+00")+" " _
						+IIf(fileFormat="pit",USFormat(ETime(i+1)*tUnit,"0.000000E+00"),"") + vbNewLine '+"          " _
						'+"% Line number: "+Str(j)
		If particlePairs Then
			j+=1
			outputString = outputString + USFormat(XYZSamples(3*i+1)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(XYZSamples(3*i+2)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(XYZSamples(3*i+3)*lUnit," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(-VXdrift/Sqr(1-VXdrift^2)-VXSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(-VYdrift/Sqr(1-VYdrift^2)-VYSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(-VZdrift/Sqr(1-VZdrift^2)-VZSamples(i+1)/CLight," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(PMass,"0.000000E+00")+" " _
							+USFormat(ZCharge*PCharge," 0.000000E+00;-0.000000E+00")+" " _
							+USFormat(IIf(fileFormat="pit",ZCharge*PCharge*SPRatio,ECurrent)," 0.000000E+00;-0.000000E+00")+" " _
							+IIf(fileFormat="pit",USFormat(ETime(i+1)*tUnit,"0.000000E+00"),"") + vbNewLine'+"          " _
							'+"% Line number: "+Str(j)
		End If
		If (((NSamples-1-i) Mod flushsize) = 0) Then ' flush outputString
			Print #tmpFile, outputString
			outputString = ""
		End If
	Next i
	' time1 = Timer()
	Close #tmpFile
	FileCopy(tmpFileName, outputFileName)
	Kill(tmpFileName)
	' MsgBox("Total write time: "+CStr(time1-time0))

	If importInterface Then
		DlgText("OutputT","Importing interface, this may take some time.")
		AddToHistory("define particle interface: "+interfaceName,"With ParticleInterface"+vbNewLine+ _
														"	.Reset"+vbNewLine+ _
	     												"	.Name "+Chr(34)+interfaceName+Chr(34)+vbNewLine+ _
		    											IIf(fileFormat="pit", _
															"	.Type "+Chr(34)+"Import ASCII TD"+Chr(34)+vbNewLine, _
															"	.Type "+Chr(34)+"Import ASCII DC"+Chr(34)+vbNewLine) + _
		    											"	.InterfaceFile "+Chr(34)+outputFileName+Chr(34)+vbNewLine+ _
		    											"	.UseLocalCopyOnly "+Chr(34)+"False"+Chr(34)+vbNewLine+ _
		    											"	.DirNew "+Chr(34)+"X"+Chr(34)+vbNewLine+ _
		    											"	.InvertOrientation "+Chr(34)+"False"+Chr(34)+vbNewLine+ _
		    											"	.XShift "+Chr(34)+"0.0"+Chr(34)+vbNewLine+ _
		    											"	.YShift "+Chr(34)+"0.0"+Chr(34)+vbNewLine+ _
		    											"	.ZShift "+Chr(34)+"0.0"+Chr(34)+vbNewLine+ _
				    									IIf(fileFormat="pit", _
		    												"	.PICEmissionModel "+Chr(34)+"TD"+Chr(34)+vbNewLine, _
															"	.PICEmissionModel "+Chr(34)+"DC"+Chr(34)+vbNewLine)+ _
		    											"	.Create"+vbNewLine+ _
		    											"End With")
	End If

	If showFile Then
		Shell "notepad " & outputFileName, 3
	End If

	DlgText("OutputT","Done!")
	GeneratePIFile = True

End Function

Function UpdateBoundingBoxInfo

	' Determine bounding box and initialize xmin...zmax accordingly
	Dim dXMinBound As Double, dXMaxBound As Double, dYMinBound As Double, dYMaxBound As Double, dZMinBound As Double, dZMaxBound As Double
	Dim dXMinBoundTmp As Double, dXMaxBoundTmp As Double, dYMinBoundTmp As Double, dYMaxBoundTmp As Double, dZMinBoundTmp As Double, dZMaxBoundTmp As Double
	If (nSolids_CST = 0) Then
		Boundary.GetCalculationBox(dXMinBound, dXMaxBound, dYMinBound, dYMaxBound, dZMinBound, dZMaxBound)
	Else
		Solid.GetLooseBoundingBoxOfShape(aSolidArray_CST(0), dXMinBound, dXMaxBound, dYMinBound, dYMaxBound, dZMinBound, dZMaxBound)
		For i = 1 To nSolids_CST-1
			Solid.GetLooseBoundingBoxOfShape(aSolidArray_CST(i), dXMinBoundTmp, dXMaxBoundTmp, dYMinBoundTmp, dYMaxBoundTmp, dZMinBoundTmp, dZMaxBoundTmp)
			dXMinBound = IIf(dXMinBoundTmp<dXMinBound, dXMinBoundTmp, dXMinBound)
			dXMaxBound = IIf(dXMaxBoundTmp>dXMaxBound, dXMaxBoundTmp, dXMaxBound)
			dYMinBound = IIf(dYMinBoundTmp<dYMinBound, dYMinBoundTmp, dYMinBound)
			dYMaxBound = IIf(dYMaxBoundTmp>dYMaxBound, dYMaxBoundTmp, dYMaxBound)
			dZMinBound = IIf(dZMinBoundTmp<dZMinBound, dZMinBoundTmp, dZMinBound)
			dZMaxBound = IIf(dZMaxBoundTmp>dZMaxBound, dZMaxBoundTmp, dZMaxBound)
		Next
	End If

	DlgText "xminT",IIf(ReStoreDialogSetting("xmin")<>"",ReStoreDialogSetting("xmin"),CStr(dXMinBound))
	DlgText "xmaxT",IIf(ReStoreDialogSetting("xmax")<>"",ReStoreDialogSetting("xmax"),CStr(dXMaxBound))
	DlgText "yminT",IIf(ReStoreDialogSetting("ymin")<>"",ReStoreDialogSetting("ymin"),CStr(dYMinBound))
	DlgText "ymaxT",IIf(ReStoreDialogSetting("ymax")<>"",ReStoreDialogSetting("ymax"),CStr(dYMaxBound))
	DlgText "zminT",IIf(ReStoreDialogSetting("zmin")<>"",ReStoreDialogSetting("zmin"),CStr(dZMinBound))
	DlgText "zmaxT",IIf(ReStoreDialogSetting("zmax")<>"",ReStoreDialogSetting("zmax"),CStr(dZMaxBound))

End Function

Function UpdateSPinfo()

	Xmin = Evaluate(DlgText("xminT"))
	Xmax = Evaluate(DlgText("xmaxT"))
	Rmin = Evaluate(DlgText("xminT"))
	Rmax = Evaluate(DlgText("xmaxT"))
	Ymin = Evaluate(DlgText("yminT"))
	Ymax = Evaluate(DlgText("ymaxT"))
	Zmin = Evaluate(DlgText("zminT"))
	Zmax = Evaluate(DlgText("zmaxT"))
	ZCharge = Evaluate(DlgText("ZChargeT"))
	PCharge = Evaluate(DlgText("PChargeT"))
	PMass = Evaluate(DlgText("PMassT"))
	NSamples=Evaluate(DlgText("NSamplesT"))
	SPRatio = Evaluate(DlgText("SPRatioT"))

	Volume = 0
	Select Case DlgValue("ShapeDD")
		Case 0
			Volume = (Xmax-Xmin)*(Ymax-Ymin)*(Zmax-Zmin)*lUnit^3
		Case 1
			Volume = pi*(Rmax^2-Rmin^2)*(Zmax-Zmin)*lUnit^3
		Case 2
			Volume = 4/3*pi*(Rmax^3-Rmin^3)*lUnit^3
		Case 3 ' Selected solids, assumes that all of the solids will be filled
			For i = 1 To nSolids_CST STEP 1
				If (aSolidArray_CST(i-1) > "") Then
					Volume = Volume + Solid.GetVolume(aSolidArray_CST(i-1))*lUnit^3
				End If
			Next i
	End Select
	If Volume > 0 Then
		DlgText "SPdensityT", USFormat(NSamples/Volume, "Scientific")
		DlgText "PdensityT", USFormat(NSamples*SPRatio/Volume, "Scientific")
	Else
		DlgText "SPdensityT", "n/d"
		DlgText "PdensityT", "n/d"
	End If
	DlgText "SPmassT", USFormat(SPRatio*PMass,"Scientific")
	DlgText "SPchargeT", USFormat(SPRatio*ZCharge*PCharge,"Scientific")
End Function

Function StoreAllSettings()

	Dim i As Long

	If Dir(iniFileName) > "" Then
		Kill iniFileName
	End If
	StoreDialogSetting("shapeDD", CStr(DlgValue("shapeDD")))
	' FSR 01/02/2014: Do not store xmin...zmax anymore now that bounding box is determined automatically
	'StoreDialogSetting("xmin", DlgText("xminT"))
	'StoreDialogSetting("xmax", DlgText("xmaxT"))
	'StoreDialogSetting("ymin", DlgText("yminT"))
	'StoreDialogSetting("ymax", DlgText("ymaxT"))
	'StoreDialogSetting("zmin", DlgText("zminT"))
	'StoreDialogSetting("zmax", DlgText("zmaxT"))
	StoreDialogSetting("PPairCB", CStr(DlgValue("PPairCB")))
	StoreDialogSetting("vx", DlgText("vxT"))
	StoreDialogSetting("vy", DlgText("vyT"))
	StoreDialogSetting("vz", DlgText("vzT"))
	StoreDialogSetting("FileFormat", CStr(DlgValue("FileFormatRB")))
	StoreDialogSetting("NSamples", DlgText("NSamplesT"))
	StoreDialogSetting("SPRatio", DlgText("SPRatioT"))
	StoreDialogSetting("seeValue", DlgText("seedValueT"))
	StoreDialogSetting("cutoffFactor", DlgText("cutoffFactorT"))
	StoreDialogSetting("ETime1", DlgText("ETime1T"))
	StoreDialogSetting("ETime2", DlgText("ETime2T"))
	StoreDialogSetting("NEmits", DlgText("NEmitsT"))
	StoreDialogSetting("ECurrent", DlgText("ECurrentT"))
	StoreDialogSetting("PTemp", DlgText("PTempT"))
	StoreDialogSetting("cutoffCB",CStr(DlgValue("cutoffCB")))
	StoreDialogSetting("importCB",CStr(DlgValue("importCB")))
	StoreDialogSetting("fixedSeedCB",CStr(DlgValue("fixedSeedCB")))
	StoreDialogSetting("showFileCB",CStr(DlgValue("showFileCB")))
	StoreDialogSetting("nSelectedSolids",CStr(nSolids_CST))
	For i = 0 To nSolids_CST-1
		StoreDialogSetting("SelectedSolid_"+CStr(i),aSolidArray_CST(i))
	Next

End Function


Function StoreDialogSetting(Key As String, Value As String)

	Dim iniFile As Long
	Dim temp$

	iniFile = FreeFile
	Open iniFileName For Append As #iniFile
	Print #iniFile,Key+"="+Value
	Close #iniFile

End Function

Function ReStoreDialogSetting$(Key$, Optional Default$)

	Dim iniFile As Long
	Dim keyFound As Boolean
	Dim lineRead As String
	Dim Value As String

	keyFound = False
	Value = Default

	iniFile = FreeFile
	If Dir(iniFileName) > "" Then
	Open iniFileName For Input As #iniFile
		While Not (keyFound Or EOF(iniFile))
			Line Input #iniFile, lineRead
			If StrComp(Split(lineRead,"=")(0),Key$,1)=0 Then
				Value = Split(lineRead,"=")(1)
				keyFound = True
			End If
		Wend
		Close #iniFile
	End If

	ReStoreDialogSetting$ = Value

End Function
