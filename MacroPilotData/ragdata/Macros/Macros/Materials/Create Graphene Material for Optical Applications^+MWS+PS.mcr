'#Language "WWB-COM"

' This macro generates a conductivity sheet material representing a graphene layer
' and additionally a equivalent permittivity as a function of the sheet thickness (used as fallback)

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' -----------------------------------------------------------------------------------------------------------------------------------------------------
' 12-Dec-2014 ckr: now fully parametric
' 09-Dec-2014 ckr: add surface conductivity to tree
' 07-Apr-2014 ckr: initial version
' -----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main

    ' Calculate a conductivity sheet material of graphene based on Falkovsky
    ' Journal of Physics: Conference Series 129 (2008) 012004 doi:10.1088/1742-6596/129/1/012004
    ' Returns a Result1DComplex object surfaceimpedance over frequency

    Dim sHistoryString As String
    Dim sMaterialName As String, sMaterialFolder As String, sTemperature As String, sChemicalPotential As String, sRelaxationTime As String, sSheetThickness As String
    Dim sMinFrequency As String, sMaxFrequency As String, sNumPoints As String
    Dim sParaNameTemperature As String, sParaNameChemPotential As String, sParaNameRelaxTime As String, sParaNameSheetThick As String


    Begin Dialog UserDialog 470,340,"Define Graphene Material" ' %GRID:10,7,1,1
        Text 20,20,220,14,"Material name:",.Text1
        TextBox 250,14,200,21,.MaterialNameT
        Text 20,50,220,14,"Material folder:",.Text2
        TextBox 250,44,200,21,.MaterialFolderT
        Text 20,80,220,14,"Temperature [K]:",.Text4
        TextBox 250,74,200,21,.TemperatureT
        Text 20,110,220,14,"Chemical potential [eV]:",.Text5
        TextBox 250,104,200,21,.ChemicalPotentialT
        Text 20,140,220,14,"Relaxation time [ps]:",.Text6
        TextBox 250,134,200,21,.RelaxationTimeT
        Text 20,170,220,14,"Thickness [nm] (for eps only):",.Text7
        TextBox 250,164,200,21,.SheetThicknessT
        Text 20,200,220,14,"Min Frequency [THz]:",.Text8
        TextBox 250,194,200,21,.MinFrequencyT
        
        Text 20,230,220,14,"Max Frequency [THz]:",.Text9
        TextBox 250,224,200,21,.MaxFrequencyT
        
        Text 20,260,220,14,"Number of Points [1]:",.Text10
        TextBox 250,254,200,21,.NumPointsT
        OKButton 260,300,90,21
        CancelButton 360,300,90,21
    End Dialog

    Dim dlg As UserDialog
    dlg.MaterialNameT = "Graphene"
    dlg.MaterialFolderT = ""
    dlg.TemperatureT = "293"
    dlg.ChemicalPotentialT = "0.0"
    dlg.RelaxationTimeT = "0.1"
    dlg.SheetThicknessT = "10"
    dlg.MinFrequencyT = "0.1"
    dlg.MaxFrequencyT = "1000"
    dlg.NumPointsT = "1000"

    If (Dialog(dlg) = 0) Then
        Exit All
    Else
        sMaterialName = dlg.MaterialNameT
        sMaterialFolder = dlg.MaterialFolderT
        sTemperature = dlg.TemperatureT
        sChemicalPotential = dlg.ChemicalPotentialT
        sRelaxationTime = dlg.RelaxationTimeT
        sSheetThickness = dlg.SheetThicknessT
        sMinFrequency = dlg.MinFrequencyT
        sMaxFrequency = dlg.MaxFrequencyT
        sNumPoints = dlg.NumPointsT
    End If

    'Dim dTStart As Double
    'dTStart = Timer()

    If sMaterialFolder <> "" Then
        sParaNameSheetThick = sMaterialFolder + "_" + sMaterialName + "_thickness"
        sParaNameTemperature = sMaterialFolder + "_" + sMaterialName + "_temperature"
        sParaNameChemPotential = sMaterialFolder + "_" + sMaterialName + "_chemPotential"
        sParaNameRelaxTime = sMaterialFolder + "_" + sMaterialName + "_relaxTime"
    Else
        sParaNameSheetThick = sMaterialName +"_thickness"
        sParaNameTemperature = sMaterialName + "_temperature"
        sParaNameChemPotential = sMaterialName + "_chemPotential"
        sParaNameRelaxTime = sMaterialName + "_relaxTime"        
    End If

    StoreDoubleParameter( sParaNameSheetThick, Evaluate(sSheetThickness) * 1e-9 * Units.GetGeometrySIToUnit )
    SetParameterDescription( sParaNameSheetThick, "Graphene: sheetThickness in Project Units" )

    StoreDoubleParameter( sParaNameTemperature, Evaluate(sTemperature) )
    SetParameterDescription( sParaNameTemperature, "Graphene: Temperature in K" )    

    StoreDoubleParameter( sParaNameChemPotential, Evaluate(sChemicalPotential) )
    SetParameterDescription( sParaNameChemPotential, "Graphene: chemicalPotential in eV" )    

    StoreDoubleParameter( sParaNameRelaxTime, Evaluate(sRelaxationTime) )
    SetParameterDescription( sParaNameRelaxTime, "Graphene: relaxationTime in ps" )    
    
    sHistoryString = ""
    sHistoryString = sHistoryString + "Dim dFMin As Double, dFMax As Double" + vbLf
    sHistoryString = sHistoryString + "Dim nSamples As Long" + vbLf
    sHistoryString = sHistoryString + "dFMin = Evaluate(" + sMinFrequency + "*1e12)" + vbLf
    sHistoryString = sHistoryString + "dFMax = Evaluate(" + sMaxFrequency + "*1e12)" + vbLf
    sHistoryString = sHistoryString + "nSamples = Evaluate(" + sNumPoints + ")" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "Dim dTemperature As Double, dChemicalPotential As Double, dRelaxationTime As Double" + vbLf
    sHistoryString = sHistoryString + "dTemperature = "+sParaNameTemperature+"" + vbLf
    sHistoryString = sHistoryString + "dChemicalPotential = "+sParaNameChemPotential+"" + vbLf
    sHistoryString = sHistoryString + "dRelaxationTime = "+sParaNameRelaxTime+"" + vbLf

    sHistoryString = sHistoryString + "" + vbLf    
    sHistoryString = sHistoryString + "MakeSureParameterExists ( """ + sParaNameSheetThick + """, "+sParaNameSheetThick+")" + vbLf
    sHistoryString = sHistoryString + "MakeSureParameterExists ( """ + sParaNameTemperature + """, "+sParaNameTemperature+")" + vbLf
    sHistoryString = sHistoryString + "MakeSureParameterExists ( """ + sParaNameChemPotential + """, "+sParaNameChemPotential+")" + vbLf
    sHistoryString = sHistoryString + "MakeSureParameterExists ( """ + sParaNameRelaxTime + """, "+sParaNameRelaxTime+")" + vbLf
    sHistoryString = sHistoryString + "" + vbLf  
    
    sHistoryString = sHistoryString + "Const dQElectron = 1.602176487e-19" + vbLf
    sHistoryString = sHistoryString + "Const dPlanckBar = 1.0545717253363e-34" + vbLf
    sHistoryString = sHistoryString + "Const dkBoltzmann = 1.3806504e-23" + vbLf
    sHistoryString = sHistoryString + "Dim dSigmaIntraRe As Double, dSigmaIntraIm As Double" + vbLf
    sHistoryString = sHistoryString + "Dim dSigmaInterRe As Double, dSigmaInterIm As Double" + vbLf
    sHistoryString = sHistoryString + "Dim dSigmaRe As Double, dSigmaIm As Double, dGeomUnit As Double, dTauInv As Double" + vbLf
    sHistoryString = sHistoryString + "Dim dSurfaceImpRe As Double, dSurfaceImpIm As Double" + vbLf
    sHistoryString = sHistoryString + "Dim dFrequency As Double, dOmega As Double, dDeltaFrequency As Double" + vbLf
    sHistoryString = sHistoryString + "Dim preFactor1 As Double, preFactor2 As Double, dd As Double" + vbLf
    sHistoryString = sHistoryString + "Dim i As Long, j As Long, k As Long" + vbLf
    sHistoryString = sHistoryString + "Dim surfaceConductivityObj As Object" + vbLf
    sHistoryString = sHistoryString + "" + vbLf

    sHistoryString = sHistoryString + "Set surfaceConductivityObj =Result1DComplex("""")" + vbLf

    sHistoryString = sHistoryString + "dGeomUnit = Evaluate(Units.GetGeometryUnitToSI)" + vbLf

    sHistoryString = sHistoryString + "' Falkovsky paper uses units 'K' for potential and omega, '1/K' for relaxation time" + vbLf
    sHistoryString = sHistoryString + "dChemicalPotential = dChemicalPotential*dQElectron / dkBoltzmann" + vbLf
    sHistoryString = sHistoryString + "dTauInv = dPlanckBar / (dkBoltzmann*dRelaxationTime*1e-12)" + vbLf
    sHistoryString = sHistoryString + "With Material" + vbLf
    sHistoryString = sHistoryString + "     .Reset" + vbLf
    sHistoryString = sHistoryString + "     .Name "+Chr(34)+sMaterialName+Chr(34)+ vbLf
    sHistoryString = sHistoryString + "     .Folder "+Chr(34)+sMaterialFolder+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .Type "+Chr(34)+"Lossy metal"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .MaterialUnit "+Chr(34)+"Frequency"+Chr(34)+", "+Chr(34)+"THz"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .SetTabulatedSurfaceImpedanceModel "+Chr(34)+"Transparent"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .DispersiveFittingSchemeTabSI "+Chr(34)+"Nth Order"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .MaximalOrderNthModelFitTabSI "+Chr(34)+"10"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .ErrorLimitNthModelFitTabSI "+Chr(34)+"0.001"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .UseOnlyDataInSimFreqRangeNthModelTabSI "+Chr(34)+"True"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     ' prepare integration for inter band contribution" + vbLf
    sHistoryString = sHistoryString + "     Dim w As Double, dw As Double, numerator As Double, denominator As Double, x As Double, y As Double, integration As Double" + vbLf
    sHistoryString = sHistoryString + "     Dim yfirst As Double, ylast As Double, GGw2 As Double, intAnalytic As Double" + vbLf
    sHistoryString = sHistoryString + "     Dim b As Double, c As Double, cosh1 As Double, cosh2 As Double" + vbLf
    sHistoryString = sHistoryString + "     Dim intDataPre(100000,1) As Double ' should be big enough" + vbLf
    sHistoryString = sHistoryString + "     Dim count As Long" + vbLf
    sHistoryString = sHistoryString + "     dd = 0." + vbLf
    sHistoryString = sHistoryString + "     w = 0." + vbLf
    sHistoryString = sHistoryString + "     dw = 2*Pi*0.1e12*dPlanckBar/dkBoltzmann" + vbLf
    sHistoryString = sHistoryString + "     count = 0" + vbLf
    sHistoryString = sHistoryString + "     While (dd < 1.-1e-14) Or (w < 0.5001 * 2*Pi*dFMax*dPlanckBar/dkBoltzmann) Or ( (count Mod 2) = 0 )" + vbLf
    sHistoryString = sHistoryString + "         dd = 1 / (Exp((-w-dChemicalPotential)/dTemperature)+1) - 1 / (Exp((w-dChemicalPotential)/dTemperature)+1)" + vbLf
    sHistoryString = sHistoryString + "         intDataPre(count,0) = w" + vbLf
    sHistoryString = sHistoryString + "         intDataPre(count,1) = dd" + vbLf
    sHistoryString = sHistoryString + "         w = w + dw" + vbLf
    sHistoryString = sHistoryString + "         count = count + 1" + vbLf
    sHistoryString = sHistoryString + "     Wend" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "     dd = (Exp(dChemicalPotential/(2*dTemperature))+Exp(-dChemicalPotential/(2*dTemperature)))/2" + vbLf
    sHistoryString = sHistoryString + "     preFactor1 = 2*dTemperature*dQElectron^2/(Pi*dPlanckBar)*Log(2*dd)" + vbLf
    sHistoryString = sHistoryString + "     preFactor2 = dQElectron^2/(4*dPlanckBar)" + vbLf
    sHistoryString = sHistoryString + "     dDeltaFrequency = (dFMax-dFMin)/(nSamples-1)" + vbLf
    sHistoryString = sHistoryString + "     For i = 0 To nSamples-1" + vbLf
    sHistoryString = sHistoryString + "         dFrequency = dFMin + i*dDeltaFrequency" + vbLf
    sHistoryString = sHistoryString + "         dOmega = 2*Pi*dFrequency*dPlanckBar/dkBoltzmann" + vbLf
    sHistoryString = sHistoryString + "         dSigmaIntraRe = preFactor1*dTauInv/(dOmega^2+dTauInv^2)" + vbLf
    sHistoryString = sHistoryString + "         dSigmaIntraIm = -preFactor1*dOmega/(dOmega^2+dTauInv^2)" + vbLf
    sHistoryString = sHistoryString + "         ' inter band contribution" + vbLf
    sHistoryString = sHistoryString + "         integration = 0" + vbLf
    sHistoryString = sHistoryString + "         GGw2 = ( 1 / (Exp((-(dOmega/2)-dChemicalPotential)/dTemperature)+1) - 1 / (Exp(((dOmega/2)-dChemicalPotential)/dTemperature)+1) )" + vbLf
    sHistoryString = sHistoryString + "         For j = 0 To count-1" + vbLf
    sHistoryString = sHistoryString + "             denominator = dOmega^2 - 4 * intDataPre(j,0)^2" + vbLf
    sHistoryString = sHistoryString + "             If Abs(denominator) > 1e-12 Then" + vbLf
    sHistoryString = sHistoryString + "                 numerator = intDataPre(j,1) - GGw2" + vbLf
    sHistoryString = sHistoryString + "                 y = numerator / denominator" + vbLf
    sHistoryString = sHistoryString + "                 If j=0 Then" + vbLf
    sHistoryString = sHistoryString + "                     yfirst = y" + vbLf
    sHistoryString = sHistoryString + "                 ElseIf j=count-1 Then" + vbLf
    sHistoryString = sHistoryString + "                     ylast = y" + vbLf
    sHistoryString = sHistoryString + "                 End If" + vbLf
    sHistoryString = sHistoryString + "             Else" + vbLf
    sHistoryString = sHistoryString + "                 ' analytic limit" + vbLf
    sHistoryString = sHistoryString + "                 b = Exp(dChemicalPotential / dTemperature)" + vbLf
    sHistoryString = sHistoryString + "                 c = Exp(dOmega/2/dTemperature)" + vbLf
    sHistoryString = sHistoryString + "                 cosh1 = ( b + 1/b ) / 2" + vbLf
    sHistoryString = sHistoryString + "                 cosh2 = ( c + 1/c ) / 2" + vbLf
    sHistoryString = sHistoryString + "                 y = -((1+cosh1*cosh2)/dTemperature) / (4*dOmega*(cosh1+cosh2)^2)" + vbLf
    sHistoryString = sHistoryString + "                 If j=0 Then" + vbLf
    sHistoryString = sHistoryString + "                     yfirst = y" + vbLf
    sHistoryString = sHistoryString + "                 ElseIf j=count-1 Then" + vbLf
    sHistoryString = sHistoryString + "                     ylast = y" + vbLf
    sHistoryString = sHistoryString + "                 End If" + vbLf
    sHistoryString = sHistoryString + "             End If" + vbLf
    sHistoryString = sHistoryString + "             ' numerical integration, simpson rule" + vbLf
    sHistoryString = sHistoryString + "             integration = integration + 2*(1+(j Mod 2))*y " + vbLf
    sHistoryString = sHistoryString + "         Next" + vbLf
    sHistoryString = sHistoryString + "         integration = (integration - yfirst - ylast) * (2*dw / 6)" + vbLf
    sHistoryString = sHistoryString + "         intAnalytic = (1 - GGw2) * ( Log(2*intDataPre(count-1,0) - dOmega) - Log(2*intDataPre(count-1,0) + dOmega) ) / (4*dOmega)" + vbLf
    sHistoryString = sHistoryString + "         dSigmaInterRe = preFactor2 * GGw2" + vbLf
    sHistoryString = sHistoryString + "         dSigmaInterIm = -preFactor2 * 4*dOmega* ( integration + intAnalytic ) / Pi" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "         dSigmaRe = dSigmaIntraRe + dSigmaInterRe" + vbLf
    sHistoryString = sHistoryString + "         dSigmaIm = dSigmaIntraIm + dSigmaInterIm" + vbLf
    sHistoryString = sHistoryString + "         surfaceConductivityObj.AppendXYDouble( dFrequency * Units.GetFrequencySIToUnit, dSigmaRe, dSigmaIm )" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "         dSurfaceImpRe = dSigmaRe/(dSigmaRe^2+dSigmaIm^2)" + vbLf
    sHistoryString = sHistoryString + "         dSurfaceImpIm = -dSigmaIm/(dSigmaRe^2+dSigmaIm^2)" + vbLf
    sHistoryString = sHistoryString + "         .AddTabulatedSurfaceImpedanceFittingValue CStr(dFrequency/1e12), CStr(dSurfaceImpRe), CStr(dSurfaceImpIm), CStr(1.0)" + vbLf
    sHistoryString = sHistoryString + "     Next" + vbLf
    sHistoryString = sHistoryString + "     .Colour "+Chr(34)+"0.2"+Chr(34)+", "+Chr(34)+"0.2"+Chr(34)+", "+Chr(34)+"0.2"+Chr(34) + vbLf
    sHistoryString = sHistoryString + "     .Create" + vbLf
    sHistoryString = sHistoryString + "End With" + vbLf
    sHistoryString = sHistoryString + "" + vbLf

    sHistoryString = sHistoryString + "surfaceConductivityObj.XLabel(""Frequency / "" + Units.GetFrequencyUnit )" + vbLf
    sHistoryString = sHistoryString + "surfaceConductivityObj.YLabel(""Conductivity / (Ohm)^(-1)"")" + vbLf
    sHistoryString = sHistoryString + "If """+sMaterialFolder+""" <> "" "" Then" + vbLf
    sHistoryString = sHistoryString + "    surfaceConductivityObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + ""surfaceConductivity_"+sMaterialFolder+"_"+sMaterialName+".sig""" + vbLf
    sHistoryString = sHistoryString + "Else" + vbLf
    sHistoryString = sHistoryString + "    surfaceConductivityObj.Save GetProjectBaseName() + GetProjectBaseNameSeparator() + ""surfaceConductivity_"+sMaterialName+".sig""" + vbLf
    sHistoryString = sHistoryString + "End If" + vbLf
    sHistoryString = sHistoryString + "surfaceConductivityObj.AddToTree ""1D Results\Dispersive Materials Information" +sMaterialFolder+"\"+sMaterialName+"\Surface Conductivity""" + vbLf

    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "With Material" + vbLf
    sHistoryString = sHistoryString + "    .Reset" + vbLf
    sHistoryString = sHistoryString + "    .Name "+Chr(34)+sMaterialName + "_Eps" + Chr(34) + vbLf
    sHistoryString = sHistoryString + "    .Folder "+Chr(34)+sMaterialFolder+Chr(34) + vbLf
    sHistoryString = sHistoryString + "    .Type ""Normal""" + vbLf
    sHistoryString = sHistoryString + "    .MaterialUnit ""Frequency"", ""THz""" + vbLf
    sHistoryString = sHistoryString + "    .DispersiveFittingFormatEps ""Real_Imag""" + vbLf
    sHistoryString = sHistoryString + "    .DispModelEps ""None""" + vbLf
    sHistoryString = sHistoryString + "    .DispModelMu ""None""" + vbLf
    sHistoryString = sHistoryString + "    .DispersiveFittingSchemeEps ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "    .MaximalOrderNthModelFitEps ""10""" + vbLf
    sHistoryString = sHistoryString + "    .ErrorLimitNthModelFitEps ""0.001""" + vbLf
    sHistoryString = sHistoryString + "    .UseOnlyDataInSimFreqRangeNthModelEps ""True""" + vbLf
    sHistoryString = sHistoryString + "    .DispersiveFittingSchemeEps ""Nth Order""" + vbLf
    sHistoryString = sHistoryString + "    .UseGeneralDispersionEps ""True""" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "    For i = 0 To nSamples-1" + vbLf
    sHistoryString = sHistoryString + "        dFrequency = surfaceConductivityObj.GetX(i) * Units.GetFrequencyUnitToSI" + vbLf
    sHistoryString = sHistoryString + "        dSigmaRe = surfaceConductivityObj.GetYRe(i)" + vbLf
    sHistoryString = sHistoryString + "        dSigmaIm = surfaceConductivityObj.GetYIm(i)" + vbLf
    sHistoryString = sHistoryString + "        .AddDispersionFittingValueEps Cstr(dFrequency/1e12), CStr(1 + dSigmaIm / (2*pi*dFrequency*Eps0*"+sParaNameSheetThick+"*dGeomUnit) ), CStr( dSigmaRe / (2*pi*dFrequency*Eps0*"+sParaNameSheetThick+"*dGeomUnit) ), ""1.0""" + vbLf
    sHistoryString = sHistoryString + "    Next i" + vbLf
    sHistoryString = sHistoryString + "" + vbLf
    sHistoryString = sHistoryString + "    .Colour ""0.25"", ""0.25"", ""0.25""" + vbLf
    sHistoryString = sHistoryString + "    .Create" + vbLf
    sHistoryString = sHistoryString + "End With"

    If sMaterialFolder <> "" Then
        AddToHistory("define material: " + sMaterialFolder + "/" + sMaterialName, sHistoryString)
    Else
        AddToHistory("define material: " + sMaterialName, sHistoryString)
    End If

    'MsgBox(CStr(Timer()-dTStart))

End Sub

