' Convert PIC2D Data to .pit file
Option Explicit

' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 26-Mar-2012 mbk: added Time Shift for subsequent PIC simulation
' 20-Sep-2010 mbk: Current column changed to MacroCharge according to 2011 changes
' 21-May-2010 mbk: first version
'
' NOTE: Interaction only properly right, if time step width of Exporting 2D monitor is equal to timestep
'
' ================================================================================================

Sub Main ()

	' At the moment max 50 PIC 2D Monitors,
	' Extension could be having a variable Array and using ReDim,
	' But hen 2 loops are needed
	' Or a new VBA command GetNumberofPIC2D Monitors is needed
	Dim afield(50) As String
	Dim ssc As String
	Dim icount As Integer
	Dim monitor_name As String

	Dim pos_x As Double, pos_y As Double, pos_z As Double
	Dim mom_x As Double, mom_y As Double, mom_z As Double
	Dim mass As Double, pcharge As Double, macrocharge As Double

	Dim nFrames As Long
	Dim dTime As Double
	Dim timeshift As Double
	Dim iFrame As Long
	Dim nParticles As Long
	Dim iParticle As Long

	Dim pitfile_name As String

	ssc = Resulttree.GetFirstChildName ("PIC 2D Monitors") '33 charakters long
	icount = 0
	While ssc <> ""
		afield(icount)= Right(ssc, Len(ssc)-16)
		icount = icount + 1
		ssc=Resulttree.GetNextItemName (ssc)
	Wend
	If (icount = 0) Then
		MsgBox "No PIC 2D Monitors defined", vbOkOnly
		Exit All
	End If

	Begin Dialog UserDialog 390,140,"Convert PIC2D Monitors to .pit File" ' %GRID:10,7,1,1
		Text 20,14,300,14,"Name of Particle Monitor",.Text1
		DropListBox 190,14,190,192,afield(),.aField
		OKButton 10,112,120,21
		CancelButton 140,112,120,21
		Text 20,56,160,28,"Time Shift for subsequent PIC simulation",.Text2
		Text 20,84,160,14,"(negative = delay)",.Text3
		TextBox 190,63,90,21,.tshift
	End Dialog

	Dim dlg As UserDialog

	' default-settings
	dlg.aField = 0
	dlg.tshift = "0"


	If (Not Dialog(dlg)) Then Exit All

	monitor_name = afield(dlg.afield)
	pitfile_name = GetProjectPath("Root") + "\" + monitor_name + ".pit"
	timeshift    = Evaluate(dlg.tshift)*Units.GetTimeUnitToSI

	Open pitfile_name For Output As #1

	With PIC2DMonitor
		.CreateMonitorData(monitor_name)

		nFrames = .GetNFrames

		For iFrame = 0 To nFrames - 1
			nParticles       = .GetNParticles(iFrame)
			dTime            = .GetTime(iFrame)

				For iParticle = 0 To nParticles - 1
					.GetPosition(iFrame, iParticle, pos_x, pos_y, pos_z)
					.GetMomentumNormed(iFrame,iParticle,mom_x,mom_y,mom_z)

					mass        = .GetMass(iFrame,iParticle)
					pcharge     = .GetCharge(iFrame, iParticle)
					macrocharge = .GetChargeMacro(iFrame, iParticle)

					Print #1, Cstr(pos_x) +"  "+  _
				  			  Cstr(pos_y) +"  "+  _
				  			  Cstr(pos_z) +"  "+  _
				  			  Cstr(mom_x) +"  "+  _
							  Cstr(mom_y) +"  "+  _
							  Cstr(mom_z) +"  "+  _
							  Cstr(mass)  +"  "+  _
							  Cstr(pcharge) +"  "+  _
							  Cstr(macrocharge) +"  "+  _
							  Cstr(dTime-timeshift)

				Next iParticle
			Next iFrame

			.ClearMonitorData

	End With

	Close #1

	MsgBox "Particle Data successfully exported to:" + pitfile_name

End Sub
