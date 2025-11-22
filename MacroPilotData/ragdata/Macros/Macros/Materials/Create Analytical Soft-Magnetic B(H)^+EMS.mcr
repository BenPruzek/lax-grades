'#Language "WWB-COM"

Option Explicit

' ================================================================================================
' Macro: Creates analytic B-H curve.
'
' Basic equations used:
' J(H, J_0, c, mu_i) = J_0 * (1 - (((mu_i - 1)*mu_0*H / (c*J_0) + 1)^(-c))
' B(H, J_0, c, mu_i) = mu_0*H + J(H, J_0, c, mu_i)
'
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 28-Feb-2018 cks: first version
' ================================================================================================

' public variables
Public M_s  As Double
Public mu_i As Double
Public dc   As Double
Public da   As Double

Sub Main
	Dim sMaterialName As String
	Dim dI As Double
	Dim H1 As Double
	Dim B1 As Double

    Dim sSaturatedM As String
    Dim sMueInitial As String
    Dim sTuningParamC As String

	Dim sCommand As String
	sCommand = ""

	Begin Dialog UserDialog 660, 553, "Create Analytical Soft-Magnetic B(H)", .DialogFunction ' %GRID:10,7,1,1
		Text          20,  20, 210,  14, "Name of material:",                 .Text1
		TextBox      240,  17, 200,  21,                                      .materialName
		Text          20,  50, 210,  14, "Initial permeability mue_i [1]:",   .Text2
		TextBox      240,  47, 200,  21,                                      .mueInitial
		Text          20,  80, 210,  14, "Saturation magnetization J_s [T]:", .Text3
		TextBox      240,  77, 200,  21,                                      .saturatedM
		Text          20, 110, 210,  14, "Tuning parameter c [1]:",           .Text4
		TextBox      240, 107, 200,  21,                                      .tuningC
		Text          20, 140, 210,  14, "Optional:",                         .Text6
		Text          40, 165, 190,  14, "Point on B(H)-Curve (H,B):",        .Text5
		TextBox      255, 162,  75,  21,                                      .pointH
		TextBox      350, 162,  75,  21,                                      .pointB
		PushButton   470, 162, 160,  21, "Calc Tuning parameter",             .CalcC
		OKButton     420, 518,  90,  21
		CancelButton 540, 518,  90,  21
		Picture       20, 200, 620, 299, GetInstallPath + "\Library\Macros\Materials\Analytical_BH_ParameterDescription.bmp",0,.Picture1
		Text         335, 165,  10,  14, ",",.Text7
		Text         240, 165,  10,  14, "(",.Text8
		Text         430, 165,  10,  14, ")",.Text9,1
	End Dialog
	Dim dlg As UserDialog

	dlg.materialName = "My SoftMagnetic Material"
    dlg.mueInitial = "5000"
    dlg.saturatedM = "0.5"
    dlg.tuningC = "2"

	If (Dialog(dlg) = 0) Then
		Exit All
	End If

	sMaterialName = dlg.materialName

    sSaturatedM = Replace$( Replace$( sMaterialName, " ", "" ), "-", "_" ) + "_saturatedM"
    sMueInitial = Replace$( Replace$( sMaterialName, " ", "" ), "-", "_" ) + "_mueInitial"
    sTuningParamC = Replace$( Replace$( sMaterialName, " ", "" ), "-", "_" ) + "_tuningParamC"
    StoreDoubleParameter(sSaturatedM, M_s)
    StoreDoubleParameter(sMueInitial, mu_i)
    StoreDoubleParameter(sTuningParamC, dc)

	sCommand = sCommand + "MakeSureParameterExists("""+sSaturatedM+""",""" + CStr(M_s) + """)" + vbLf
	sCommand = sCommand + "MakeSureParameterExists("""+sMueInitial+""",""" + CStr(mu_i) + """)" + vbLf
	sCommand = sCommand + "MakeSureParameterExists("""+sTuningParamC+""",""" + CStr(dc) + """)" + vbLf
	sCommand = sCommand + vbLf

	sCommand = sCommand + "Dim dFluxDensity As Double" + vbLf
	sCommand = sCommand + "Dim dFieldStrength As Double" + vbLf
	sCommand = sCommand + "Dim dMagnetization As Double" + vbLf
	sCommand = sCommand + "Dim da As Double" + vbLf
	sCommand = sCommand + "Dim db As Double" + vbLf
	sCommand = sCommand + "Dim dI As Double" + vbLf
	sCommand = sCommand + "Dim dtmp As Double" + vbLf
	sCommand = sCommand + vbLf

	sCommand = sCommand + "If "+sTuningParamC+" <= 0 Then" + vbLf
	sCommand = sCommand + "  ReportError(""The Tuning parameter c has to be positive!"")" + vbLf
	sCommand = sCommand + "End If" + vbLf
	sCommand = sCommand + "If "+sMueInitial+" <= 1 Then" + vbLf
	sCommand = sCommand + "  ReportError(""The Initial permeability mue_i has to be bigger than 1!"")" + vbLf
	sCommand = sCommand + "End If" + vbLf
	sCommand = sCommand + "If "+sSaturatedM+" <= 0 Then" + vbLf
	sCommand = sCommand + "  ReportError(""The Saturation magnetization J_s has to be positive!"")" + vbLf
	sCommand = sCommand + "End If" + vbLf
	sCommand = sCommand + vbLf

	sCommand = sCommand + "da = ("+sMueInitial+" - 1.0)*mue0" + vbLf
	sCommand = sCommand + "db = da / "+sSaturatedM+" / "+sTuningParamC + vbLf

	sCommand = sCommand + vbLf
	sCommand = sCommand + "With Material" + vbLf
	sCommand = sCommand + "  .Reset" + vbLf
	sCommand = sCommand + "  .Name """ + sMaterialName + """" + vbLf
	sCommand = sCommand + "  .Folder ""Analytical Soft-Magnetic B(H) (Macro)""" + vbLf
	sCommand = sCommand + "  .FrqType ""All""" + vbLf
	sCommand = sCommand + "  .Type ""Normal""" + vbLf
	sCommand = sCommand + "  .Mu "+ sMueInitial + vbLf
	sCommand = sCommand + "  .Sigma ""0""" + vbLf
	sCommand = sCommand + "  .NonlinearMeasurementError ""1e-3""" + vbLf
	sCommand = sCommand + "  .ResetHBList" + vbLf
	sCommand = sCommand + "  .SetNonlinearCurveType ""Soft-Magnetic-BH""" + vbLf
	sCommand = sCommand + "  .AddNonlinearCurveValue ""0.0"", ""0.0""" + vbLf

	sCommand = sCommand + "For dI = 0 To 6.001 STEP 0.025" + vbLf

	sCommand = sCommand + "   dFieldStrength = 10^dI" + vbLf

	sCommand = sCommand + "   dMagnetization = "+sSaturatedM+" * ( 1. - (db*dFieldStrength + 1.)^(- "+sTuningParamC+"))" + vbLf

	sCommand = sCommand + "   dFluxDensity = mue0*dFieldStrength + dMagnetization" + vbLf

	sCommand = sCommand + "   .AddNonlinearCurveValue CStr(dFieldStrength), CStr(dFluxDensity)" + vbLf

	sCommand = sCommand + "Next dI" + vbLf

	sCommand = sCommand + "  .GenerateNonlinearCurve" + vbLf
	sCommand = sCommand + "  .Create" + vbLf
	sCommand = sCommand + "End With" + vbLf
	AddToHistory "(*) define material: " + sMaterialName, sCommand
End Sub



Private Function DialogFunction(sDlgItem As String, iAction As Integer, lSuppValue As Long) As Boolean
	Dim H1 As Double
	Dim B1 As Double

	Dim numberOfBisectionSteps As Integer
	Dim bisectionCounter As Integer
	Dim cmin As Double
	Dim cmax As Double
	Dim tol As Double
	Dim fc As Double
	Dim fcmin As Double
	Dim fcmax As Double
	Dim tuningC As Double
	Dim db As Double
	Dim success As Boolean

	Select Case iAction
		Case 1 ' Dialog Box initialisation.

		    M_s = CDBL( DlgText("saturatedM") )
			mu_i = CDBL( DlgText("mueInitial") )
			dc = CDBL( DlgText("tuningC") )
			da = (mu_i - 1.0)*mue0

		Case 2 ' Value changing or button pressed

			M_s = CDBL( DlgText("saturatedM") )
			mu_i = CDBL( DlgText("mueInitial") )
			dc = CDBL( DlgText("tuningC") )
			da = (mu_i - 1.0)*mue0

			Select Case sDlgItem
				Case "CalcC"
					cmin = 0.001
					cmax = 1000
					tol = 1e-6
					numberOfBisectionSteps = 500
					success = False

					H1 = CDBL( DlgText("pointH") )
					B1 = CDBL( DlgText("pointB") )

					db = da / M_s / cmin
					fcmin = B1 - mue0*H1 - M_s * ( 1. - (db*H1 + 1.)^(- cmin))
					db = da / M_s / cmax
					fcmax = B1 - mue0*H1 - M_s * ( 1. - (db*H1 + 1.)^(- cmax))


					If ( (B1 / H1)  < (mu_i*mue0) ) And ( (B1 - mue0*H1) < M_s ) And ( Sgn(fcmax) <> Sgn(fcmin) )  Then
							Do While bisectionCounter <= numberOfBisectionSteps ' limit iterations to prevent infinite loop
								tuningC = (cmin + cmax)/2 ' new midpoint
								db = da / M_s / tuningC
								fc = B1 - mue0*H1 - M_s * ( 1. - (db*H1 + 1.)^(- tuningC))
								db = da / M_s / cmin
								fcmin = B1 - mue0*H1 - M_s * ( 1. - (db*H1 + 1.)^(- cmin))
								If (fc = 0) Or ((cmax - cmin)/2 < tol) Then
									success = True
									Exit Do
								End If
								bisectionCounter = bisectionCounter + 1 'increment step counter
								If Sgn(fc) = Sgn(fcmin) Then
									cmin = tuningC
								Else
									cmax = tuningC ' new interval
								End If
							Loop
					End If
					If Not success Then
						ReportWarning("Inconsistant B(H)-Point given. Tuning Parameter can not be calculated!")
					Else
						'ReportInformation(CStr(bisectionCounter) +  " bisections have been performed !")
						DlgText "tuningC", CStr( tuningC )
						dc = CDBL( DlgText("tuningC") )
					End If
					DialogFunction = True
                Case "OK"
					M_s = CDBL( DlgText("saturatedM") )
					mu_i = CDBL( DlgText("mueInitial") )
					dc = CDBL( DlgText("tuningC") )
					da = (mu_i - 1.0)*mue0
			End Select
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function
