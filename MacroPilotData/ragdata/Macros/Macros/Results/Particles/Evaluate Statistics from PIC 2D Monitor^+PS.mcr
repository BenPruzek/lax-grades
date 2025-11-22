' Evaluate 2D PIC Monitor

' ================================================================================================
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 08-Apr-2014 fsr: Density, momentum, etc. in 2D plot are now shown an monitor position, not boundaries anymore
' 16-Aug-2013 fsr: Reduced memory footprint, improved performance
' 10-May-2010 fsr: Added skipping of empty frames, some GUI options
' 06-May-2010 fsr: GUI and small improvements
' 27-Apr-2010 fsr: Initial version

Sub Main()

	' Get monitor list and fill array
	Dim monitorList() As String
	ReDim monitorList(0)
	monitorList(0) = Split(Resulttree.GetFirstChildName("PIC 2D Monitors"),"\")(1)
	While Not (Resulttree.GetNextItemName("PIC 2D Monitors\"+monitorList(UBound(monitorList)))="")
		ReDim Preserve monitorList(UBound(monitorList)+1)
		monitorList(UBound(monitorList)) = Resulttree.GetNextItemName("PIC 2D Monitors\"+monitorList(UBound(monitorList)-1))
		monitorList(UBound(monitorList)) = Split(monitorList(UBound(monitorList)),"\")(1)
	Wend

	Begin Dialog UserDialog 890,245,"Evaluate 2D PIC Monitor",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 20,14,350,170,monitorList(),.MonitorListDBox
		OKButton 20,84,90,21
		CancelButton 130,84,90,21
		Text 30,49,330,14,"Please select monitor",.statusText
		CheckBox 240,88,120,14,"Abort",.CheckBox1
		TextBox 20,119,840,112,.messageBox,2
		Text 390,21,240,14,"Number of species in model:",.nSpeciesT
		Text 390,49,240,14,"Min. number of particles for histogram:",.histoNLimitT
		TextBox 650,14,60,21,.nSpeciesTT
		TextBox 650,42,60,21,.histoNLimitTT
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub


Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
			DlgEnable("CheckBox1",False)
			DlgText("nSpeciesTT","1")
			DlgText("histoNLimitTT","150")
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "OK"
				DialogFunc = True ' Prevent button press from closing the dialog box
				DlgEnable("Ok",False)
				DlgEnable("Cancel",False)
				DlgEnable("CheckBox1",True)
				Evaluate2DPICMonitor(DlgText("MonitorListDBox"), _
										DlgText("nSpeciesTT"), _
										DlgText("histoNLimitTT"))
				DlgEnable("Ok",True)
				DlgEnable("Cancel",True)
				DlgEnable("CheckBox1",False)
			Case "Cancel"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub Evaluate2DPICMonitor(monitorName As String, nSpecies As Integer, histoNLimit As Long)

	'--------------------- Initialization ----------------------
	'Dim monitorName As String
	Dim nFrames As Long, nFrameStep As Long
	Dim directionL As Long
	Dim dMonitorPosition As Double
	Dim directionS As String
	Dim frameNParticles() As Long
	Dim particlesOverTime As Object
	Dim particlesOverTimeS As String
	Dim particlesInFrameBySpecies() As Object
	Dim particlesInFrameBySpeciesS() As String
	Dim frameCurrent() As Double
	Dim currentOverTime As Object
	Dim currentOverTimeS As String
	Dim frameTime() As Double
	Dim timeOverFrame As Object
	Dim timeOverFrameS As String
	Dim frameTimeStep() As Double
	Dim frameCharge() As Double
	Dim chargeOverTime As Object
	Dim chargeOverTimeS As String
	Dim frameMacroCharge() As Double
	Dim macroChargeOverTime As Object
	Dim macroChargeOverTimeS As String
	Dim speciesMQ() As Double 	' Array to store mass and charge of a given species
	Dim nBins As Integer		' number of bins for histogram
	Dim bins() As Long
	Dim binWidth() As Double
	Dim edfPlot() As Object		' Result1D objects for EDF plots
	Dim edfPlotS() As String	' Names of plots
	Dim cumEDFPlot() As Object		' Result1D objects for cumulative EDF plots
	Dim cumEDFPlotS() As String	' Names of plots
	Dim histoXmax() As Double	' xmax of edfPlot
	Dim histoXmin() As Double	' xmin of edfPlot
	'Dim histoNLimit As Long		' Minimum number of particles to create a histogram
	Dim lowParticleN As Boolean ' Found a frame with low number of particles?
	Dim massList() As Double	' array to store mass values of particles in a given frame
	Dim momList() As Double		' array to store normed momentum components of particles in a given frame
	Dim momComponent As Integer	' 0=X,1=Y,2=Z,3=Abs
	Dim momComponentS() As String
	Dim chargeList() As Double	' array to store charge of particles in a given frame
	Dim energyList() As Double	' array to store energy of particles in a given frame (calculated value)
	Dim listFillLevel() As Long	' array to keep track of how full the lists are for each species
	Dim meanSum() As Double
	Dim energyMean() As Double
	Dim sigmaSum() As Double
	Dim energySigma() As Double
	Dim meanEnergyOverTime() As Object
	Dim meanEnergyOverTimeS() As String
	Dim sigmaEnergyOverTime() As Object
	Dim sigmaEnergyOverTimeS() As String

	Dim density2D() As Object
	Dim density2DS() As String
	Dim momentum2D() As Object
	Dim momentum2DS() As String
	Dim energy2D() As Object
	Dim energy2DS() As String
	Dim plot2Dnx As Integer
	Dim plot2Dny As Integer
	Dim plot2Dnz As Integer
	Dim particlePosition(2) As Double
	Dim tmpValue As Double
	Dim closestCellIndex As Long

	Dim i As Long, j As Long, k As Long, m As Long
	Dim outputString As String
	Dim ResultFolder1D As String
	Dim ResultFolder2D As String

	Dim lUnitS As String
	Dim lUnit As Double
	Dim tUnitS As String
	Dim tUnit As Double
	Dim Boundaries(5) As Double ' xmin,xmax,ymin,ymax,zmin,zmax

	Dim dStartTime As Double
	dStartTime = Timer()

	lUnitS = Units.GetUnit("Length")
	lUnit = Units.GetGeometryUnitToSI()
	tUnitS = Units.GetUnit("Time")
	tUnit = Units.GetTimeUnitToSI()
	Boundary.GetCalculationBox(Boundaries(0),Boundaries(1),Boundaries(2),Boundaries(3),Boundaries(4),Boundaries(5))

	Resulttree.EnableTreeUpdate(False)

	lowParticleN = False
	outputString = ""
	ReDim momComponentS(3)
	momComponentS(0)="X"
	momComponentS(1)="Y"
	momComponentS(2)="Z"
	momComponentS(3)="Abs"
	ResultFolder1D="1D Results\PIC 2D Monitor Results\"+monitorName+"\"
	ResultFolder2D="2D/3D Results\PIC 2D Monitor Results\"+monitorName+"\"
	SendToMessageBox("I: Loading PIC 2D monitor '"+monitorName+"' ... ")
	PIC2DMonitor.CreateMonitorData(monitorName)
	SendToMessageBox("Loaded, starting analysis."+vbNewLine)
	directionL=PIC2DMonitor.GetDirection()

	Set timeOverFrame = Result1D("")
	timeOverFrameS = "Time over frame number"
	Set particlesOverTime = Result1D("")
	particlesOverTimeS = "Number of particles over time"
	Set currentOverTime = Result1D("")
	currentOverTimeS = "Current over time"
	Set chargeOverTime = Result1D("")
	chargeOverTimeS = "Charge over time"
	Set macroChargeOverTime = Result1D("")
	macroChargeOverTimeS = "Macro charge over time"

	'-------------------- Data gathering -----------------------
	nFrames = PIC2DMonitor.GetNFrames
	'nFrameStep = Fix(nFrames/nFrameSamples)
	'If (nFrameStep = 0) Then nFrameStep = 1
	ReDim frameNParticles(nFrames-1) As Long
	ReDim frameCurrent(nFrames-1) As Double
	ReDim frameTime(nFrames-1) As Double
	ReDim frameTimeStep(nFrames-1) As Double
	ReDim frameCharge(nFrames-1) As Double
	ReDim frameMacroCharge(nFrames-1) As Double

	For i=0 To nFrames-1
		frameNParticles(i) = PIC2DMonitor.GetNParticles(i)
		frameCurrent(i) = PIC2DMonitor.GetCurrentPerFrame(i)
		frameTime(i) = PIC2DMonitor.GetTime(i)
		frameTimeStep(i) = PIC2DMonitor.GetTimeStep(i)
		frameCharge(i) = PIC2DMonitor.GetChargeTotal(i)
		frameMacroCharge(i) = PIC2DMonitor.GetChargeTotalMacro(i)
	Next

	'----------------------- Output ----------------------------
	For i=0 To nFrames-1
		timeOverFrame.AppendXY(i+1,frameTime(i))
		particlesOverTime.AppendXY(frameTime(i),frameNParticles(i))
		currentOverTime.AppendXY(frameTime(i),frameCurrent(i))
		chargeOverTime.AppendXY(frameTime(i),frameCharge(i))
		macroChargeOverTime.AppendXY(frameTime(i),frameMacroCharge(i))
	Next i
	timeOverFrame.Title(timeOverFrameS)
	timeOverFrame.SetXLabelAndUnit("Frame number","1")
	timeOverFrame.SetYLabelAndUnit("Time", "s")
	timeOverFrame.Save(monitorName+"_"+timeOverFrameS)
	timeOverFrame.AddToTree(ResultFolder1D+timeOverFrameS)

	particlesOverTime.Title(particlesOverTimeS)
	particlesOverTime.SetXLabelAndUnit("Time","s")
	particlesOverTime.SetYLabelAndUnit("Number of particles","1")
	particlesOverTime.Save(monitorName+"_"+particlesOverTimeS)
	particlesOverTime.AddToTree(ResultFolder1D+particlesOverTimeS)

	currentOverTime.Title(currentOverTimeS)
	currentOverTime.SetXLabelAndUnit("Time","s")
	currentOverTime.SetYLabelAndUnit("Current","A")
	currentOverTime.Save(monitorName+"_"+currentOverTimeS)
	currentOverTime.AddToTree(ResultFolder1D+currentOverTimeS)

	chargeOverTime.Title(chargeOverTimeS)
	chargeOverTime.SetXLabelAndUnit("Time","s")
	chargeOverTime.SetYLabelAndUnit("Charge","C")
	chargeOverTime.Save(monitorName+"_"+chargeOverTimeS)
	chargeOverTime.AddToTree(ResultFolder1D+chargeOverTimeS)

	macroChargeOverTime.Title(macroChargeOverTimeS)
	macroChargeOverTime.SetXLabelAndUnit("Time","s")
	macroChargeOverTime.SetYLabelAndUnit("Macro charge","C")
	macroChargeOverTime.Save(monitorName+"_"+macroChargeOverTimeS)
	macroChargeOverTime.AddToTree(ResultFolder1D+macroChargeOverTimeS)

	' This part finds the number of species... 3 nested loops, so it is slow!
	' Better to have the user enter the number of species (see below), he should know....

	'nSpecies = 1									' We have at least 1 species
	'ReDim speciesMQ(1,nSpecies-1)
	'speciesMQ(0,nSpecies-1) = PIC2DMonitor.GetMass(0,0) 	' Store mass of first particle in first frame
	'speciesMQ(1,nSpecies-1) = PIC2DMonitor.GetCharge(0,0) 	' Store charge of first particle in first frame
	'
	' Parse through particles in each frame, how many species are there?
	'For i=0 To nFrames-1 ' Consider one frame at a time
	'	For j=0 To frameNParticles(i)-1
	'		For k=0 To nSpecies-1
	'			' If the same species, both mass and charge have to match
	'			If (PIC2DMonitor.GetMass(i,j)=speciesMQ(0,k) And PIC2DMonitor.GetCharge(i,j)=speciesMQ(1,k)) Then
	'				Exit For
	'			ElseIf k=nSpecies-1 Then ' If not exited at end of k loop, a new species has been found
	'				nSpecies+=1
	'				ReDim Preserve speciesMQ(1,nSpecies-1)
	'				speciesMQ(0,nSpecies-1) = PIC2DMonitor.GetMass(i,j)
	'				speciesMQ(1,nSpecies-1) = PIC2DMonitor.GetCharge(i,j)
	'			End If
	'		Next k
	'	Next j
	'Next i
	' nSpecies is now known

	m = 0 ' Number of species found so far
	' If the number of species is known, things go faster
	ReDim speciesMQ(1,nSpecies-1)
	'speciesMQ(0,m-1) = PIC2DMonitor.GetMass(0,0) 	' Store mass of first particle in first frame
	'speciesMQ(1,m-1) = PIC2DMonitor.GetCharge(0,0) 	' Store charge of first particle in first frame
	'SendToMessageBox("I: Found species "+cstr(m)+": m="+Format(speciesMQ(0,m-1),"scientific")+" kg, q="+Format(speciesMQ(1,m-1),"scientific")+" C"+vbNewLine)
	For i=0 To nFrames-1 ' Consider one frame at a time
		For j=0 To frameNParticles(i)-1
			For k=0 To m
				' If the same species, both mass and charge have to match
				If Not (PIC2DMonitor.GetMass(i,j)=speciesMQ(0,k) And PIC2DMonitor.GetCharge(i,j)=speciesMQ(1,k)) Then
					' Found a new species!
					m+=1
					speciesMQ(0,m-1) = PIC2DMonitor.GetMass(i,j)
					speciesMQ(1,m-1) = PIC2DMonitor.GetCharge(i,j)
					SendToMessageBox("I: Found species "+cstr(m)+": m="+Format(speciesMQ(0,m-1),"scientific")+" kg, q="+Format(speciesMQ(1,m-1),"scientific")+" C"+vbNewLine)
				End If
				If (m=nSpecies) Then Exit For
			Next k
			If (m=nSpecies) Then Exit For
		Next j
		If (m=nSpecies) Then Exit For
	Next i

	'Dim speciesInfoS As String
	'speciesInfoS = ""
	'For i = 0 To nSpecies-1
	'	speciesInfoS = speciesInfoS+"Species "+cstr(i+1)+": m="+Format(speciesMQ(0,i),"scientific")+" kg, q="+Format(speciesMQ(1,i),"scientific")+" C"+vbNewLine
	'Next i
	'MsgBox(speciesInfoS)

	' Calculate edfPlot
	'nBins = 100			' Number of bins for edfPlot, calculated below depending upon number of particles
	momComponent = 3	' 0=X,1=Y,2=Z,3=Abs

	' Get mesh info
	plot2Dnx = Mesh.GetNx
	plot2Dny = Mesh.GetNy
	plot2Dnz = Mesh.GetNz

	' Set up Result1D objects and names for plots
	ReDim particlesInFrameBySpecies(nSpecies-1)
	ReDim particlesInFrameBySpeciesS(nSpecies-1)
	ReDim meanEnergyOverTime(nSpecies-1)
	ReDim meanEnergyOverTimeS(nSpecies-1)
	ReDim sigmaEnergyOverTime(nSpecies-1)
	ReDim sigmaEnergyOverTimeS(nSpecies-1)
	For k=0 To nSpecies-1
		Set meanEnergyOverTime(k) = Result1D("")
		meanEnergyOverTimeS(k) = "Mean particle energy ("+momComponentS(momComponent)+") over time, species "+cstr(k+1)
		Set sigmaEnergyOverTime(k) = Result1D("")
		sigmaEnergyOverTimeS(k) = "Particle energy ("+momComponentS(momComponent)+") standard deviation over time, species "+cstr(k+1)
		Set particlesInFrameBySpecies(k) = Result1D("")
		particlesInFrameBySpeciesS(k) = "Number of particles (species "+cstr(k+1)+") per frame"
	Next k

	ReDim edfPlot(nSpecies-1)
	ReDim edfPlotS(nSpecies-1)
	ReDim cumEDFPlot(nSpecies-1)
	ReDim cumEDFPlotS(nSpecies-1)
	ReDim density2D(nSpecies-1)
	ReDim density2DS(nSpecies-1)
	ReDim momentum2D(nSpecies-1)
	ReDim momentum2DS(nSpecies-1)
	ReDim energy2D(nSpecies-1)
	ReDim energy2DS(nSpecies-1)

	ReDim histoXmax(nSpecies-1)
	ReDim histoXmin(nSpecies-1)
	ReDim binWidth(nSpecies-1)

	For i=0 To nFrames-1 STEP 1	' Consider one frame at a time

		For k = 0 To nSpecies-1
			' Initialize arrays for current species
			Set edfPlot(k) = Result1D("")
			edfPlotS(k) = "EDF("+momComponentS(momComponent)+"), species "+cstr(k+1)+", frame "+cstr(i+1)
			Set cumEDFPlot(k) = Result1D("")
			cumEDFPlotS(k) = "Cumulative EDF("+momComponentS(momComponent)+"), species "+cstr(k+1)+", frame "+cstr(i+1)
			Set density2D(k) = Result3D("")
			density2DS(k) = ("Density, species "+cstr(k+1)+", frame "+cstr(i+1))
			density2D(k).Initialize(plot2Dnx, plot2Dny, plot2Dnz, "scalar")
			density2D(k).SetType("static charge")
			Set momentum2D(k) = Result3D("")
			momentum2DS(k) = ("Momentum, species "+cstr(k+1)+", frame "+cstr(i+1))
			momentum2D(k).Initialize(plot2Dnx, plot2Dny, plot2Dnz, "vector")
			momentum2D(k).SetType("dynamic current")
			Set energy2D(k) = Result3D("")
			energy2DS(k) = ("Energy("+momComponentS(momComponent)+"), species "+cstr(k+1)+", frame "+cstr(i+1))
			energy2D(k).Initialize(plot2Dnx, plot2Dny, plot2Dnz, "scalar")
			energy2D(k).SetType("static charge")
		Next k

		If (frameNParticles(i) >0) Then ' Only consider frames that contain particles

			' reset some frame related variables
			ReDim listFillLevel(nSpecies-1)
			ReDim meanSum(nSpecies-1)
			ReDim sigmaSum(nSpecies-1)
			ReDim energyMean(nSpecies-1)
			ReDim energySigma(nSpecies-1)

			' For each frame, read and store some particle values in lists
			ReDim massList(nSpecies-1,frameNParticles(i)-1)
			ReDim momList(nSpecies-1,3,frameNParticles(i)-1) ' dimension in center for x/y/z/abs component
			ReDim chargeList(nSpecies-1,frameNParticles(i)-1)
			ReDim energyList(nSpecies-1,frameNParticles(i)-1)

			' Parse through all particles in frame i
			For j=0 To frameNParticles(i)-1
				For k=0 To nSpecies-1

					' Arrays initialized, start calculations
					If (PIC2DMonitor.GetMass(i,j)=speciesMQ(0,k) And PIC2DMonitor.GetCharge(i,j)=speciesMQ(1,k)) Then
						' Particle j belongs to species k, store in the correct list and adjust fill level of that list
						massList(k,listFillLevel(k)) = PIC2DMonitor.GetMass(i,j)
						PIC2DMonitor.GetMomentumNormed(i,j,momList(k,0,listFillLevel(k)),momList(k,1,listFillLevel(k)),momList(k,2,listFillLevel(k)))
						momList(k,3,listFillLevel(k)) = PIC2DMonitor.GetMomentumNormedAbs(i,j)
						chargeList(k,listFillLevel(k)) = PIC2DMonitor.GetCharge(i,j)
						' E[eV] = (m-m0)*c^2/q = m0(1/sqrt(1+u^2)-1)*c^2/|q|
						energyList(k,listFillLevel(k)) = speciesMQ(0,k)*(Sqr(1+momList(k,momComponent,listFillLevel(k))^2)-1)*CLight^2/Abs(speciesMQ(1,k))
						meanSum(k)=meanSum(k)+energyList(k,listFillLevel(k))
						PIC2DMonitor.GetPosition(i,j,particlePosition(0),particlePosition(1),particlePosition(2))
						dMonitorPosition = PIC2DMonitor.GetWPosition() ' Get the monitor position along its normal
						Select Case directionL
							Case 0
								closestCellIndex = Mesh.GetClosestPtIndex(dMonitorPosition,particlePosition(1)/lUnit,particlePosition(2)/lUnit) ' project particles onto x plane
							Case 1
								closestCellIndex = Mesh.GetClosestPtIndex(particlePosition(0)/lUnit,dMonitorPosition,particlePosition(2)/lUnit) ' project particles onto y plane
							Case 2
								closestCellIndex = Mesh.GetClosestPtIndex(particlePosition(0)/lUnit,particlePosition(1)/lUnit,dMonitorPosition) ' project particles onto z plane
						End Select
						' Evaluate densities, histogram like
						tmpValue = density2D(k).GetXRe(closestCellIndex)
						density2D(k).SetXRe(closestCellIndex,tmpValue+1)
						' Evaluate momentum distribution, histogram like... v<<c, so normed momentum is approx v/c
						tmpValue = momentum2D(k).GetXRe(closestCellIndex)
						momentum2D(k).SetXRe(closestCellIndex,tmpValue+momList(k,0,listFillLevel(k))*CLight)
						tmpValue = momentum2D(k).GetYRe(closestCellIndex)
						momentum2D(k).SetYRe(closestCellIndex,tmpValue+momList(k,1,listFillLevel(k))*CLight)
						tmpValue = momentum2D(k).GetZRe(closestCellIndex)
						momentum2D(k).SetZRe(closestCellIndex,tmpValue+momList(k,2,listFillLevel(k))*CLight)
						' Evaluate energy distribution, histogram like
						tmpValue = energy2D(k).GetXRe(closestCellIndex)
						energy2D(k).SetXRe(closestCellIndex,tmpValue+speciesMQ(0,k)*(Sqr(1+momList(k,momComponent,listFillLevel(k))^2)-1)*CLight^2/Abs(speciesMQ(1,k)))
						listFillLevel(k)+=1
					End If
				Next k
				If ((j=0) Or (j Mod 100 = 0) Or (j=frameNParticles(i)-1)) Then DlgText("statusText","Frame "+cstr(i+1)+"/"+cstr(nFrames)+", Particle "+cstr(j+1)+"/"+cstr(frameNParticles(i)))
				If DlgValue("CheckBox1") = 1 Then Exit All
			Next j

			' Find xmax and xmin of histogram for each species
			For k=0 To nSpecies-1
				particlesInFrameBySpecies(k).AppendXY(i+1, listFillLevel(k))
				' Only run if a statistically large number of particles is present in frame (>0 is absolutely necessary!)
				If listFillLevel(k)>histoNLimit Then
					nBins = Round(Sqr(listFillLevel(k)))
					ReDim bins(nSpecies-1,nBins-1)
					histoXmax(k) = energyList(k,0)
					For m = 1 To listFillLevel(k)-1
			    		If energyList(k,m) > histoXmax(k) Then
			        		histoXmax(k) = energyList(k,m)
			    		End If
					Next m
					histoXmin(k) = energyList(k,0)
					For m = 1 To listFillLevel(k)-1
			    		If energyList(k,m) < histoXmin(k) Then
			        		histoXmin(k) = energyList(k,m)
			    		End If
					Next m
					' Calculate bin width
					binWidth(k) = (histoXmax(k)-histoXmin(k))/nBins
					'Fill bins
					For m = 0 To listFillLevel(k)-1
						' bins(species,bin number)
						bins(k,Fix(((energyList(k,m)-histoXmin(k))/binWidth(k))*0.9999))+=1 ' *0.9999 to avoid problems at histoXmax
					Next m
					For m = 0 To nBins-1
						edfPlot(k).AppendXY(histoXmin(k)+binWidth(k)*(m+1/2),bins(k,m)/listFillLevel(k))
					Next m
					cumEDFPlot(k).AppendXY(histoXmin(k)+binWidth(k)/2,bins(k,0)/listFillLevel(k))
					For m = 1 To nBins-1
						cumEDFPlot(k).AppendXY(histoXmin(k)+binWidth(k)*(m+1/2),cumEDFPlot(k).GetY(m-1)+bins(k,m)/listFillLevel(k))
					Next m

					' Normalize
					density2D(k).ScalarMult(1/listFillLevel(k)/lUnit^2)	' 2D monitor, so scale with lUnit^2 only, not lUnit^3
					energy2D(k).ScalarMult(1/listFillLevel(k))
					energyMean(k) = meanSum(k)/listFillLevel(k)
					meanEnergyOverTime(k).AppendXY(frameTime(i), energyMean(k))
					For m = 0 To listFillLevel(k)-1
						sigmaSum(k) = sigmaSum(k)+(energyList(k,m)-energyMean(k))^2
					Next
					energySigma(k) = Sqr(sigmaSum(k))/listFillLevel(k)
					sigmaEnergyOverTime(k).AppendXY(frameTime(i), energySigma(k))

					edfPlot(k).Title(edfPlotS(k))
					edfPlot(k).SetXLabelAndUnit("Energy","eV")
					edfPlot(k).SetYLabelAndUnit("Normalized EDF("+momComponentS(momComponent)+"), Value","1")
					edfPlot(k).Save(monitorName+"_"+edfPlotS(k))
					edfPlot(k).AddToTree(ResultFolder1D+"EDF\Species "+cstr(k+1)+": m="+Format(speciesMQ(0,k),"scientific")+" kg, q="+Format(speciesMQ(1,k),"scientific")+" C"+"\"+edfPlotS(k))
					cumEDFPlot(k).Title(cumEDFPlotS(k))
					cumEDFPlot(k).SetXLabelAndUnit("Energy","eV")
					cumEDFPlot(k).SetYLabelAndUnit("Cumulative Normalized EDF("+momComponentS(momComponent)+"), Value","1")
					cumEDFPlot(k).Save(monitorName+"_"+cumEDFPlotS(k))
					cumEDFPlot(k).AddToTree(ResultFolder1D+"Cumulative EDF\Species "+cstr(k+1)+": m="+Format(speciesMQ(0,k),"scientific")+" kg, q="+Format(speciesMQ(1,k),"scientific")+" C"+"\"+cumEDFPlotS(k))

					density2D(k).Save("^"+monitorName+"_"+density2DS(k))
					density2D(k).AddToTree(ResultFolder2D+"Density\species "+cstr(k+1)+"\"+density2DS(k),density2DS(k))
					momentum2D(k).Save("^"+monitorName+"_"+momentum2DS(k))
					momentum2D(k).AddToTree(ResultFolder2D+"Momentum\species "+cstr(k+1)+"\"+momentum2DS(k),momentum2DS(k))
					energy2D(k).Save("^"+monitorName+"_"+energy2DS(k))
					energy2D(k).AddToTree(ResultFolder2D+"Energy\species "+cstr(k+1)+"\"+energy2DS(k),energy2DS(k))
				Else
					SendToMessageBox("W: Fewer than "+cstr(histoNLimit)+" particles of species "+cstr(k+1)+" in frame "+cstr(i+1)+", skipping some statistics for this case."+vbNewLine)
					lowParticleN = True
				End If
			Next k
		Else	' Frame contains 0 particles
			SendToMessageBox("W: Frame "+cstr(i+1)+" contains no particles, skipping frame."+vbNewLine)
		End If
	Next i

	' Save and display plots
	For k=0 To nSpecies-1
		particlesInFrameBySpecies(k).Title(particlesInFrameBySpeciesS(k))
		particlesInFrameBySpecies(k).SetXLabelAndUnit("Frame number","1")
		particlesInFrameBySpecies(k).SetYLabelAndUnit("Number of particles","1")
		particlesInFrameBySpecies(k).Save(monitorName+"_"+particlesInFrameBySpeciesS(k))
		particlesInFrameBySpecies(k).AddToTree(ResultFolder1D+"Number of particles by species\"+particlesInFrameBySpeciesS(k))

		If (meanEnergyOverTime(k).GetN > 0) Then	' If all frames had a low number of particles, this plot does not exist
			meanEnergyOverTime(k).Title(meanEnergyOverTimeS(k))
			meanEnergyOverTime(k).SetXLabelAndUnit("Time","s")
			meanEnergyOverTime(k).SetYLabelAndUnit("Mean energy","eV")
			meanEnergyOverTime(k).Save(monitorName+"_"+meanEnergyOverTimeS(k))
			meanEnergyOverTime(k).AddToTree(ResultFolder1D+"EDF Mean\"+meanEnergyOverTimeS(k))
		End If

		If (sigmaEnergyOverTime(k).GetN > 0) Then	' If all frames had a low number of particles, this plot does not exist
			sigmaEnergyOverTime(k).Title(sigmaEnergyOverTimeS(k))
			sigmaEnergyOverTime(k).SetXLabelAndUnit("Time","s")
			sigmaEnergyOverTime(k).SetYLabelAndUnit("Standard deviation of particle energy","eV")
			sigmaEnergyOverTime(k).Save(monitorName+"_"+sigmaEnergyOverTimeS(k))
			sigmaEnergyOverTime(k).AddToTree(ResultFolder1D+"EDF Sigma\"+sigmaEnergyOverTimeS(k))
		End If
	Next k

	'----------------- Clean up and close ----------------------
	PIC2DMonitor.ClearMonitorData
	Resulttree.EnableTreeUpdate(True)
	Resulttree.UpdateTree
	SendToMessageBox("I: Done!")
	ReportInformationToWindow("Done in " + CStr(Timer()-dStartTime) + " secs.")

End Sub

Public Function ArrayMax(ByRef MyArray As Variant, Lower As Long, Upper As Long) As Double
	Dim i As Long
	ArrayMax = MyArray(Lower)
	For i = Lower+1 To Upper
	    If MyArray(i) > ArrayMax Then
	        ArrayMax = MyArray(i)
	    End If
	Next
End Function

Public Function ArrayMin(ByRef MyArray As Variant, Lower As Long, Upper As Long) As Double
	Dim i As Long
	ArrayMin = MyArray(Lower)
	For i = Lower+1 To Upper
	    If MyArray(i) < ArrayMin Then
	        ArrayMin = MyArray(i)
	    End If
	Next
End Function

Public Sub SendToMessageBox(message As String)
	DlgText("messageBox",DlgText("messageBox")+message)
End Sub
