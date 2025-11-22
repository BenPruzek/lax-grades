' TDR Cross-Probing
' This macro probes a given time from the TDR plot to the structure
' using the information contained in time domain power flow monitors

' ================================================================================================
' Copyright 2009-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------
' 07-Jul-2020 ube: use SelectModelView in addition to SelectTreeItem "Components" (to ensure, 2d3dplot windows become inactive)
' 06-Sep-2016 ebu: removed deprecated command
' 07-Jul-2014 rsj: Modify line 214 VectorPlot3D.GetNumberOfSamples to n, to avoid an error msg during opening the project
' 06-Jun-2014 ctc: Deactivate WCS before executing macro
' 06-Jan-2014 msc: Added Dialog to parameterize the wait between plotting samples to workaround plotting bug
' 25-Jun-2013 msc: Added Workaround for timing issue in plotting and reading maximum of powerflow monitor
' 14-May-2013 msc: Safety check added in case of power flow monitor with not enough samples.
' 09-Nov-2011 rsj: Cdbl inside Crosscorrelate to avoid some issue in string conversion.
' 03-Nov-2011 rsj: Changed result cache detection method for linux compatibility
' 05-May-2011 rsj: Added result cache
' 12-Jan-2010 msc: Bugfix in monitor drop down list
' 27-Nov-2009 ube: Help button added
' 24-Jul-2009 msc: Plot update included + Safety checks
' 01-Jul-2009 msc: initial version
'------------------------------------------------------------------------------

'#Language "WWB-COM"

Option Explicit

Public PFlowArray() As String
Public nPFReslts As Integer

Sub Main ()


	WCS.Store ("TempWCS")
	WCS.ActivateWCS ("global")

	'TDR Cross-Probing between TDR and structure
	If (Not CheckIfAnyTDPFlowDefined()) Then
		MsgBox("No time-domain power flow monitor defined. Please define one and run the macro again.")
		Exit All
	End If

	Call FillPFArray()

	Begin Dialog UserDialog 435,190,"TDR Cross-Probing",.DialogFunc ' %GRID:5,5,1,1
		GroupBox 15,10,405,140,"Settings",.GroupBox1
		TextBox 30,45,300,20,.Time
		Text 30,30,230,15,"TDR time (based on 50% at origin)",.Text1
		PushButton 25,160,190,20,"Cross-Probe to Structure",.CrossCorr
		PushButton 220,160,90,20,"Close",.Close
		Text 30,80,275,15,"Specify available Power Flow Result",.Text2
		DropListBox 30,95,375,125,PFlowArray(),.PFlowDrop
		Text 345,50,65,15,"time unit",.tUnit
		PushButton 320,160,90,20,"Help",.Help
		CheckBox 30,125,155,15,"Use result cache",.result_cache
		TextBox 305,120,45,20,.DebugWait
		Text 185,125,115,15,"Sample plot delay",.Text3
		Text 360,125,30,15,"/sec",.Text4
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

	WCS.Restore ("TempWCS")
	WCS.Delete ("TempWCS")

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim iMon As Integer
	iMon=Cint(DlgValue("PFlowDrop")) + 1 ' +1 because of 0 / 1 based indexing
	Dim File_exists As String


	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "Time", "Please enter time" ' Initiaization
		DlgText "tUnit", Units.GetUnit("Time")
		DlgText "DebugWait", "1"

		'RSJ: Initialize the Dialog box For result cache

		If PFlowArray(iMon) = "" Then
			MsgBox("No power flow monitor found. Please ensure a power flow result is present.")
			Exit Function
		End If

		'On Error Resume Next
		File_exists = Dir$(GetProjectPath("Result")+"TDR_CrossProbe_cache_"+Cstr(iMon)+".sig")
		If File_exists <> "" Then
			DlgValue("result_cache",1)
			DlgEnable("result_cache",True)
		Else
			DlgValue("result_cache",0)
			DlgEnable("result_cache",False)
		End If

	Case 2 ' Value changing or button pressed
		Rem %DialogFunc = True ' Prevent button press from closing the dialog box

	 	Select Case DlgItem$
	        Case "CrossCorr" 	'
	            DialogFunc = True 				'do not exit the dialog
				Dim dTime As Double
				Dim sTime As String
				Dim sDebugWaitTime As String
				sDebugWaitTime = DlgText("DebugWait")


				sTime = DlgText("Time")
				If (IsNumeric(sTime)=True) Then
					dTime = CDbl(sTime)
				Else
					MsgBox("Invalid time. Please enter a numeric time.")
					Exit Function
				End If

				If DlgValue("result_cache")=1 Then
					CrossCorrelate(dTime, PFlowArray(iMon),iMon,True, Cint(sDebugWaitTime))
					DlgValue("result_cache",1)
					DlgEnable("result_cache",True)
				Else
				    CrossCorrelate(dTime, PFlowArray(iMon),iMon,False,Cint(sDebugWaitTime))
				    DlgValue("result_cache",1)
				    DlgEnable("result_cache",True)
				End If
				SelectTreeItem "Components"
				SelectModelView
				Plot.Wireframe True
				Plot.Update

			'RSJ: Added for checking different result cache from power flow monitor index
			Case "PFlowDrop"
				'RSJ: Use the index of field monitor
				If PFlowArray(iMon) = "" Then
					MsgBox("No power flow monitor found. Please ensure a power flow result is present.")
					Exit Function
				End If
				File_exists = Dir$(GetProjectPath("Result")+"TDR_CrossProbe_cache_"+Cstr(iMon)+".sig")
				If File_exists <>"" Then
					DlgValue("result_cache",1)
					DlgEnable("result_cache",True)
				Else
					DlgValue("result_cache",0)
					DlgEnable("result_cache",False)
				End If

			Case "Help"
				StartHelp "common_preloadedmacro_transient_tdr_crossprobing"
				DialogFunc = True
        	Case "Close" 	'
				DialogFunc = False
		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : %DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


Sub CrossCorrelate(tTDR As Double, sPFName As String,mon_index As Integer,cst_use_cache As Boolean, iWaitSec As Integer)


	Dim sMonName As String, cst_line_input As String
	sMonName = "2D/3D Results\Power Flow\" + sPFName
	Pick.ClearAllPicks
	Dim ii As Integer
	ii=0

    Dim n As Integer, i As Integer
    Dim x() As Double, y() As Double, z() As Double, max As Double, t() As Double

	'RSJ: Add a simple text file for caching
 	If cst_use_cache=True Then
		'Load the result cache
		Open GetProjectPath("Result")+"TDR_CrossProbe_cache_"+Cstr(mon_index)+".sig" For Input As #2
		While Not EOF(2)
		    Line Input #2,cst_line_input
			If ii=0 Then
				n=Cint(cst_line_input)
				ReDim x(n), y(n),z(n), t(n)
				ii=1
			Else
				t(ii)=Cdbl(Split(cst_line_input," ")(0))
				x(ii)=Cdbl(Split(cst_line_input," ")(1))
				y(ii)=Cdbl(Split(cst_line_input," ")(2))
				z(ii)=Cdbl(Split(cst_line_input," ")(3))
				ii=ii+1
			End If
		Wend
		Close #2
 	Else
 		SelectTreeItem (sMonName)
 	    Plot3DPlotsOn2DPlane(False)   ' Deactivate 2D Cutting Plane Plot
    	VectorPlot3D.type "timesamplemax" ' schaltet die Abfrage auf Sample-weise
    	VectorPlot3D.type "withpicks"
 		Open GetProjectPath("Result")+"TDR_CrossProbe_cache_"+Cstr(mon_index)+".sig" For Output As #1
 		n =	Plot2D3D.GetNumberOfSamples
 		ReDim x(n), y(n),z(n), t(n)
 		Print #1, Cstr(n)
		'Build xyzt map
		For i = 1 To n
    		Plot2D3D.SetSample(i)
	    	Wait iWaitSec ' Workaround for timing issue with plotting and accessing maximum field
	    	max = GetFieldPlotMaximumPos(x(i), y(i), z(i))
	    	t(i) = Plot2D3D.GetTime
			'MsgBox Str$(max)+" at (" + Str$(x) + ", " + Str$(y) + ", " + Str$(z) + ")"
			'RSJ: write the array into a text file
			'format: time xi yi zi
			Print #1,Cstr(t(i))+" "+Cstr(x(i))+" "+Cstr(y(i))+" "+Cstr(z(i))
	    Next
		Close #1
 	End If

    Dim FieldTshift As Double
	FieldTshift = GetFieldTShift()

	i=0

	While ((t(i)-FieldTshift)<(tTDR)/2)
		i=i+1
			If (i>n) Then
			MsgBox("TDR Probing time outside of power flow monitor interval.")
			Exit All
		End If
	Wend

	' Constant velocity interpolation
	Dim x_TDR As Double, y_TDR As Double, z_TDR As Double
	Dim vx_TDR As Double, vy_TDR As Double, vz_TDR As Double

	vx_TDR = (x(i)-x(i-1))/(t(i)-t(i-1))
	vy_TDR = (y(i)-y(i-1))/(t(i)-t(i-1))
	vz_TDR = (z(i)-z(i-1))/(t(i)-t(i-1))


	x_TDR = x(i-1) + vx_TDR *(tTDR/2+FieldTshift-t(i-1))
	y_TDR = y(i-1) + vy_TDR *(tTDR/2+FieldTshift-t(i-1))
	z_TDR = z(i-1) + vz_TDR *(tTDR/2+FieldTshift-t(i-1))


	SelectTreeItem "Components"
	SelectModelView
	Pick.PickPointFromCoordinates(x_TDR, y_TDR, z_TDR)

End Sub


Sub DebugPlotAllMax

	Pick.ClearAllPicks
    VectorPlot3D.type "timesamplemax" ' schaltet die Abfrage auf Sample-weise
    VectorPlot3D.type "withpicks"
    Dim n As Integer, i As Integer
    Dim x() As Double, y() As Double, z() As Double, max As Double, t() As Double

	n =	Plot2D3D.GetNumberOfSamples

 	ReDim x(n), y(n),z(n), t(n)

    For i = 1 To n
    	Plot2D3D.SetSample(i)
	    max = GetFieldPlotMaximumPos(x(i), y(i), z(i))
	'   MsgBox Str$(max)+" at (" + Str$(x) + ", " + Str$(y) + ", " + Str$(z) + ")"
    Next

    SelectTreeItem "Components"
	SelectModelView

    For i = 1 To n
    	Pick.PickPointFromCoordinates(x(i),y(i),2)
    Next

    SelectTreeItem "Components"
End Sub

Function GetFieldTShift() As Double


	' case gaussian with fmin = 0
	Dim dBT As Double
	dBT = 3.5545485 ' constant BT: duration of excitation * bandwidth = constant = dBT

	Dim dFmin As Double
	Dim dFmax As Double
	dFmin = Solver.GetFmin*Units.GetFrequencyUnitToSI
	dFmax = Solver.GetFmax*Units.GetFrequencyUnitToSI

	Dim dTshift As Double
	dTshift = dBT/(dFmax-dFmin)*0.5*Units.GetTimeSIToUnit ' max Gauss Pulse = 50% Gauss Step

	Return dTshift
End Function

Function CheckIfAnyTDPFlowDefined As Boolean


	Dim nMon As Integer
	nMon =	Monitor.GetNumberOfMonitors()

	Dim nPFMon As Integer
	nPFMon = 0

	Dim i As Integer

	For i = 0 To nMon-1
		If (Monitor.GetMonitorTypeFromIndex(i)="Powerflow 3D") Then
			If (Monitor.GetMonitorDomainFromIndex(i)="Time") Then
				nPFMon = nPFMon + 1
			End If
		End If
	Next i

	Return CBool(nPFMon)
End Function

Sub FillPFArray

	Dim sNameIter As String
	Dim sNameFill As String

	nPFReslts = 0

	sNameIter = Resulttree.GetFirstChildName ("2D/3D Results\Power Flow")

	While sNameIter  <> ""			' get number of monitors first for redim

		sNameFill = Right(sNameIter, InStrRev(sNameIter, "\"))
		nPFReslts = nPFReslts + 1
		sNameIter = Resulttree.GetNextItemName(sNameIter)
	Wend

	ReDim PFlowArray(nPFReslts) As String

	sNameIter = Resulttree.GetFirstChildName ("2D/3D Results\Power Flow")

	Dim i As Integer

	For i = 1 To nPFReslts
		sNameFill = Right(sNameIter, Len(sNameIter)-InStrRev(sNameIter, "\"))
		PFlowArray(i) = sNameFill
		sNameIter = Resulttree.GetNextItemName(sNameIter)
	Next

End Sub
