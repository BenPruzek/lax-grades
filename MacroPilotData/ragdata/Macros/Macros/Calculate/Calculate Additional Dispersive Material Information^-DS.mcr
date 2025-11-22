' This macro calculates refractive index, extinction coefficient, wavelength, skin depth, and wave impedance over frequency;
' helpful to determine worst case skin depth / wavelength for dispersive materials.
' The macro also creates surface impedance materials that represent the behaviour of the volume material.
'
' ================================================================================================
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' -----------------------------------------------------------------------------------------------------------------------------------------------------
' 23-Aug-2018 ckr: Show subset of information in LF Solvers as well
' 07-Jul-2017 ckr: linux compatibility bug fixed
' 08-Dec-2015 ckr: all curves and materials are now created ones.
'				Before there was the problem that sometimes the created surface materials where not there when reopening the model
' 14-Nov-2014 ckr: fix bug: in some names of materials there was a problem
' 02-Nov-2014 ckr: include dialog to choose between list or fit data
' 24-Sep-2014 ckr: using list instead of fit data if existing, persistent folder and bug fixes
' 07-Apr-2014 ckr: initial version
' -----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main

Dim sHistoryString As String
sHistoryString = ""

Dim listFlag As Boolean


Begin Dialog UserDialog 410,98 ' %GRID:10,7,1,1
    CheckBox 40,14,320,14,"Use dispersive data list instead of fitted values",.CheckBox1
    OKButton 210,63,70,21
    CancelButton 310,63,70,21
End Dialog
Dim dlg As UserDialog
dlg.CheckBox1 = False

If (Dialog(dlg) = 0) Then
    Exit All
Else
    listFlag = dlg.CheckBox1
End If


AddToHistory "--- execute Macro: Calculate Additional Dispersive Material Information ---", "' New treefolder created with additional material information for each of the dispersive materials defined so far"


Dim i As Integer, j As Integer, k As Integer, counter As Integer

Dim numMaterials As Long
numMaterials = Material.GetNumberOfMaterials

Dim materialNames() As String
ReDim materialNames(numMaterials-1)

Dim index As Long
For index = 0 To numMaterials - 1
    materialNames(index) = Material.GetNameOfMaterialFromIndex( index )
Next

Dim materialName As String
Dim materialNameSI As String
Dim materialType As String
Dim treeItem As String
Dim fileName As String
Dim fileNames() As String
Dim flags() As Integer
Dim listFlagEps As Boolean, listFlagMu As Boolean
Dim nDataPtsEps As Integer, nDataPtsMu  As Integer, nDataPts  As Integer
Dim epsReObj As Object, epsImObj As Object
Dim muReObj As Object, muImObj As Object
Dim eps As Double, epsX As Double, epsY As Double, epsZ As Double
Dim mu As Double, muX As Double, muY As Double, muZ As Double
Dim lambdaObj As Object
Dim skindepthObj As Object
Dim reflectivityObj As Object
Dim waveimpedanceObj As Object
Dim nObj As Object
Dim kappaObj As Object
Dim sigmaObj As Object
Dim n As Double
Dim kappa As Double
Dim sigma As Double
Dim freq As Double
Dim lambda As Double
Dim skindepth As Double
Dim eps_r As Double
Dim eps_im As Double
Dim mu_r As Double
Dim mu_im As Double
Dim reflectivity As Double
Dim waveimpedanceRe As Double
Dim waveimpedanceIm As Double
Dim nStartIndex As Long
Dim dataIndex As Long
Dim sActiveSolver As String
Dim bIsLF As Boolean

bIsLF = False
sActiveSolver = GetSolverType()
If Split( sActiveSolver, " " )(0) = "LF" Then
	bIsLF = True
End If

For index = 0 To numMaterials - 1
' loop over all materials
    materialName = materialNames(index)
    materialType = Material.GetTypeOfMaterial( materialName )

    treeItem = Resulttree.GetFirstChildName( "1D Results\Materials\" + Replace( materialName, "/", "\" ) + "\Dispersive" )
    If materialType <> "Normal" Or treeItem = "" Then
        GoTo ContinueFor
    End If

    counter = 0
    ReDim fileNames(1, counter)
    While treeItem <> ""
        fileNames(0, counter) = Resulttree.GetFileFromTreeItem( treeItem )
        fileNames(1, counter) = Mid( fileNames(0, counter), Len( GetProjectPath ("Result")) + 1 )
        counter = counter + 1
        ReDim Preserve fileNames(1, counter)
        treeItem = Resulttree.GetNextItemName( treeItem )
    Wend

    listFlagEps = False
    listFlagMu = False
    ReDim flags(counter-1)
    For i = 0 To counter - 1
        fileName = fileNames(1, i)
        If InStr( fileName, "eps" ) > 0 Then
            If InStr( fileName, "re" ) > 0 Then
                If InStr( fileName, "_mre_" ) > 0 Then
                    If listFlag Then
                        flags(i) = 1
                        listFlagEps = True
                    End If
                ElseIf InStr( fileName, "_re." ) > 0 Then
                    flags(i) = 5
                End If
            End If
            If InStr( fileName, "im" ) > 0 Then
                If InStr( fileName, "_mim_" ) > 0 Then
                    If listFlag Then
                        flags(i) = 2
                    End If
                ElseIf InStr( fileName, "_im." ) > 0 Then
                    flags(i) = 6
                End If
            End If
        ElseIf InStr( fileName, "mu" ) > 0 Or InStr( fileName, "mue" ) > 0 Then
            If InStr( fileName, "re" ) > 0 Then
                If InStr( fileName, "_mre_" ) > 0 Then
                    If listFlag Then
                        flags(i) = 3
                        listFlagMu = True
                    End If
                ElseIf InStr( fileName, "_re." ) > 0 Then
                    flags(i) = 7
                End If
            End If
            If InStr( fileName, "im" ) > 0 Then
                If InStr( fileName, "_mim_" ) > 0 Then
                    If listFlag Then
                        flags(i) = 4
                    End If
                ElseIf InStr( fileName, "_im." ) > 0 Then
                    flags(i) = 8
                End If
            End If
        End If
    Next i

    For i = 0 To counter - 1
        If listFlagEps Then
            If flags(i) = 5 Then
                flags(i) = 0
            End If
            If flags(i) = 6 Then
                flags(i) = 0
            End If
        End If
        If listFlagMu Then
            If flags(i) = 7 Then
                flags(i) = 0
            End If
            If flags(i) = 8 Then
                flags(i) = 0
            End If
        End If
    Next


    nDataPtsEps = 0
    nDataPtsMu = 0
    nDataPts = 0
    For i = 0 To counter - 1

        fileName = fileNames(0, i)
        If flags(i) = 0 Then
            GoTo ContinueFor2
        End If
        If flags(i) = 1 Then
            Set epsReObj = Result1D( fileName )
            nDataPtsEps = epsReObj.GetN
        End If
        If flags(i) = 5 Then
            Set epsReObj = Result1D( fileName )
            nDataPtsEps = epsReObj.GetN
        End If
        If flags(i) = 2 Then
            Set epsImObj = Result1D( fileName )
        End If
        If flags(i) = 6 Then
            Set epsImObj = Result1D( fileName )
        End If
        If flags(i) = 3 Then
            Set muReObj = Result1D( fileName )
            nDataPtsMu = muReObj.GetN
        End If
        If flags(i) = 7 Then
            Set muReObj = Result1D( fileName )
            nDataPtsMu = muReObj.GetN
        End If
        If flags(i) = 4 Then
            Set muImObj = Result1D( fileName )
        End If
        If flags(i) = 8 Then
            Set muImObj = Result1D( fileName )
        End If
    ContinueFor2:
    Next i

   If (nDataPtsEps > 0) And (nDataPtsMu > 0) Then
        If (nDataPtsEps <= nDataPtsMu) Then
            muReObj.MakeCompatibleTo( epsReObj )
            muImObj.MakeCompatibleTo( epsImObj )
        End If
        If (nDataPtsMu <= nDataPtsEps) Then
            epsReObj.MakeCompatibleTo( muReObj )
            epsImObj.MakeCompatibleTo( muImObj )
        End If
        nDataPts = epsReObj.GetN
    End If

    If epsReObj Is Nothing Then
        ' indicates magnetic dispersion
        nDataPts = muReObj.GetN
        ' gets the maximum eps from the diagonal tensor
        Material.GetEpsilon( materialName, epsX, epsY, epsZ )
        eps = epsX
        eps = IIf( epsY > eps, epsY, eps )
        eps = IIf( epsZ > eps, epsZ, eps )
        ' initialize Result1D object and fill it with data
        Set epsReObj = Result1D("")
        Set epsImObj = Result1D("")
        epsReObj.Initialize nDataPts
        epsImObj.Initialize nDataPts
        For i = 0 To nDataPts - 1
            epsReObj.SetXY( i, muReObj.GetX(i), eps )
            epsImObj.SetXY( i, muReObj.GetX(i), 0.0 )
        Next i
    End If

    If muReObj Is Nothing Then
        ' indicates electric dispersion
        nDataPts = epsReObj.GetN
        ' gets the maximum mu from the diagonal tensor
        Material.GetMu( materialName, muX, muY, muZ )
        mu = muX
        mu = IIf( muY > mu, muY, mu )
        mu = IIf( muZ > mu, muZ, mu )
        ' initialize Result1D object and fill it with data
        Set muReObj = Result1D("")
        Set muImObj = Result1D("")
        muReObj.Initialize nDataPts
        muImObj.Initialize nDataPts
        For i = 0 To nDataPts - 1
            muReObj.SetXY( i, epsReObj.GetX(i), mu )
            muImObj.SetXY( i, epsReObj.GetX(i), 0.0 )
        Next i
    End If

    ' here the real work starts
    Set lambdaObj = Result1D("")
    Set skindepthObj = Result1D("")
    Set reflectivityObj = Result1D("")
    Set waveimpedanceObj = Result1DComplex("")
    Set nObj = Result1D("")
    Set kappaObj = Result1D("")
    Set sigmaObj = Result1D("")
    lambdaObj.Initialize nDataPts
    skindepthObj.Initialize nDataPts
    reflectivityObj.Initialize nDataPts
    waveimpedanceObj.Initialize nDataPts
    nObj.Initialize nDataPts
    kappaObj.Initialize nDataPts
    sigmaObj.Initialize nDataPts

    nStartIndex = IIf( epsReObj.GetX(0) = 0, 1, 0 ) ' omit first data point if DC
    For dataIndex = nStartIndex To nDataPts - 1
        freq = epsReObj.GetX(dataIndex)

        eps_r = epsReObj.GetY(dataIndex)
        eps_im = epsImObj.GetY(dataIndex)
        mu_r = muReObj.GetY(dataIndex)
        mu_im = muImObj.GetY(dataIndex)

        sigma = 2*Pi*freq*Units.GetFrequencyUnitToSI*eps0*eps_im
        sigmaObj.SetXY( dataIndex, freq, sigma )

        kappa = Sqr( Sqr(eps_r^2+eps_im^2) * Sqr(mu_r^2+mu_im^2) ) * Sin( (Atn2(-eps_im, eps_r) + Atn2(-mu_im, mu_r)) / 2 )

        skindepth = -( CLight / ( 2*Pi*freq*Units.GetFrequencyUnitToSI*kappa ) ) * Units.GetGeometrySIToUnit
        skindepthObj.SetXY( dataIndex, freq, skindepth )
        kappaObj.SetXY( dataIndex, freq, -kappa )

        n = Sqr( Sqr(eps_r^2+eps_im^2) * Sqr(mu_r^2+mu_im^2) ) * Cos( (Atn2(-eps_im, eps_r) + Atn2(-mu_im, mu_r)) / 2 )
        lambda = ( CLight / ( n*freq*Units.GetFrequencyUnitToSI ) ) * Units.GetGeometrySIToUnit
        lambdaObj.SetXY( dataIndex, freq, lambda )
        nObj.SetXY( dataIndex, freq, n )

        reflectivity = ( (n-1)^2 + kappa^2 ) / ( (n+1)^2 + kappa^2 )
        reflectivityObj.SetXY( dataIndex, freq, reflectivity )

        waveimpedanceRe = Sqr(mue0/eps0) * Sqr( Sqr( mu_r^2 + mu_im^2 ) / Sqr( eps_r^2 + eps_im^2 ) )*Cos( (Atn2(-mu_im, mu_r) - Atn2(-eps_im, eps_r)) / 2 )
        waveimpedanceIm = Sqr(mue0/eps0) * Sqr( Sqr( mu_r^2 + mu_im^2 ) / Sqr( eps_r^2 + eps_im^2 ) )*Sin( (Atn2(-mu_im, mu_r) - Atn2(-eps_im, eps_r)) / 2 )

        waveimpedanceObj.SetX( dataIndex, freq )
        waveimpedanceObj.SetYRe( dataIndex, waveimpedanceRe )
        waveimpedanceObj.SetYIm( dataIndex, waveimpedanceIm )

    Next dataIndex

    lambdaObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    lambdaObj.SetYLabelAndUnit "Wavelength" , Units.GetUnit("Length")

    skindepthObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    skindepthObj.SetYLabelAndUnit "Skindepth" , Units.GetUnit("Length")

    reflectivityObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    reflectivityObj.SetYLabelAndUnit "Reflectivity", "1"

    waveimpedanceObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    waveimpedanceObj.SetYLabelAndUnit "Waveimpedance", "Ohm"

    nObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    nObj.SetYLabelAndUnit "Refractive Index", "1"

    kappaObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    kappaObj.SetYLabelAndUnit "Extinction Coefficient", "1"

    sigmaObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    sigmaObj.SetYLabelAndUnit "Conductivity" , "(Ohm.m)^(-1)"

    waveimpedanceObj.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
    waveimpedanceObj.SetYLabelAndUnit "Complex Impedance" , "Ohm.m^(-2)"

    lambdaObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "wavelength_" + materialName + ".sig"
    skindepthObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "skindepth_" + materialName + ".sig"
    reflectivityObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "reflectivity_" + materialName + ".sig"
    nObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "refractiveIndex_" + materialName + ".sig"
    kappaObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "extinctionCoefficient_" + materialName + ".sig"
    sigmaObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "conductivity_" + materialName + ".sig"
	waveimpedanceObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + "waveImpedance_" + materialName + ".sig"

    lambdaObj.DeleteAt("never")
    skindepthObj.DeleteAt("never")
    reflectivityObj.DeleteAt("never")
    nObj.DeleteAt("never")
    kappaObj.DeleteAt("never")
    sigmaObj.DeleteAt("never")
	waveimpedanceObj.DeleteAt("never")

	If bIsLF Then
	    lambdaObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Wavelength"
    	skindepthObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Skindepth"
		sigmaObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Conductivity"
		waveimpedanceObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Surface impedance"
    Else
	    lambdaObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Wavelength"
	    skindepthObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Skindepth"
	    reflectivityObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Reflectivity"
	    nObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Refractive Index"
	    kappaObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Extinction Coefficient"
	    sigmaObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Conductivity"
		waveimpedanceObj.AddToTree "1D Results\Dispersive Materials Information (Macro)\" + Replace( materialName, "/", "\" ) + "\Waveimpedance"
	End If

	materialNameSI = Split(materialName,"/")(UBound(Split(materialName,"/")))

	sHistoryString = sHistoryString + "    With Material" + vbLf
	sHistoryString = sHistoryString + "        .Reset" + vbLf
	sHistoryString = sHistoryString + "        .Name """+ materialNameSI + """" + vbLf
	sHistoryString = sHistoryString + "        .Folder ""Surface Impedances of Dispersive Materials (Macro)""" + vbLf
	sHistoryString = sHistoryString + "        .FrqType ""all""" + vbLf
	sHistoryString = sHistoryString + "        .Type ""Lossy metal""" + vbLf
	sHistoryString = sHistoryString + "        .MaterialUnit ""Frequency"", Units.GetUnit(""Frequency"")" + vbLf
	sHistoryString = sHistoryString + "        .MaterialUnit ""Geometry"", Units.GetUnit(""Length"")" + vbLf
	sHistoryString = sHistoryString + "        .MaterialUnit ""Time"", Units.GetUnit(""Time"")" + vbLf
	sHistoryString = sHistoryString + "        .MaterialUnit ""Temperature"", ""Kelvin""" + vbLf
	sHistoryString = sHistoryString + "        .SetTabulatedSurfaceImpedanceModel ""Opaque""" + vbLf
	sHistoryString = sHistoryString + "        .DispersiveFittingSchemeTabSI ""Nth Order""" + vbLf
	sHistoryString = sHistoryString + "        .MaximalOrderNthModelFitTabSI ""10""" + vbLf
	sHistoryString = sHistoryString + "        .ErrorLimitNthModelFitTabSI ""0.01""" + vbLf
	sHistoryString = sHistoryString + "        .UseOnlyDataInSimFreqRangeNthModelTabSI ""True""" + vbLf

    For dataIndex = nStartIndex To nDataPts - 1
        freq = epsReObj.GetX(dataIndex)
        waveimpedanceRe = waveimpedanceObj.GetYRe( dataIndex )
        waveimpedanceIm = waveimpedanceObj.GetYIm( dataIndex )
		sHistoryString = sHistoryString + "        .AddTabulatedSurfaceImpedanceFittingValue """ +CStr(freq)+""", """+CStr(waveimpedanceRe)+""", """+CStr(waveimpedanceIm)+""", ""1.0""" + vbLf

	Next
	sHistoryString = sHistoryString + "        .Create" + vbLf
	sHistoryString = sHistoryString + "    End With" + vbLf

	If bIsLF = False Then
		AddToHistory "Define Material: " + materialNameSI + " (Surface Impedance)", sHistoryString
	End If

    Set epsReObj = Nothing
    Set muReObj = Nothing

    Set epsImObj = Nothing
    Set muImObj = Nothing
    sHistoryString = ""

ContinueFor:
Next index

AddToHistory "--- execute Macro: Calculate Additional Dispersive Material Information ---", ""

End Sub
