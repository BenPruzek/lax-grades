'#Language "WWB-COM"

Option Explicit
Option Base 0

' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 19-Jul-2017 mwi: fixed off-by-one error in comments in trajectory export (trajectories are 
'                  numbered 0..N-1 when looking at the TrajectoryID quantity);
'                  added message about exported file name
' 08-Jun-2017 mwi: fixed off-by-one error in sample selector, added PIC2DMonitor support (needs
'                  CST Studio Suite 2018), added error handling if file is inaccessible
' 02-Jun-2017 mwi: initial implementation
' ================================================================================================

Enum MonitorType
	Particle2DMonitor
	PIC2DMonitor
	PICPositionMonitor
	Trajectories
End Enum

Type Monitor
	sName As String
	eType As MonitorType
End Type

Type Quantity
	sName As String
	sComp As String
End Type


Dim lstMonitors() As Monitor
Dim lstQuantities() As Quantity

Dim lstSamples() As String

Dim iSelectedSource As Long
Dim iSelectedSample As Long

Const sFileName = GetProjectPath("Root") + "\" + "ascii_export.csv"
Const sSeparator = ";"
Const sCommentChar = "# "


Sub Main

	Begin Dialog UserDialog 290,399,"Particle Reader Object ASCII export",.DialogFunction ' %GRID:10,7,1,1
		Text 10,7,170,14,"Data Source:",.Text1
		DropListBox 10,21,270,21,lstMonitors(),.lbSource
		Text 10,49,100,14,"Samples",.txtSamples
		DropListBox 10,63,270,21,lstSamples(),.lbSamples
		Text 10,98,260,14,"Quantities to be exported:",.Text2
		Text 60,350,220,14,"Multi-select using Ctrl+Click",.Text3,1
		MultiListBox 10,119,270,231,lstMonitors(),.lbQuantities

		OKButton 110,371,80,21
		CancelButton 200,371,80,21
	End Dialog

    Dim dlg As UserDialog

    If Dialog(dlg) Then
	    DoASCIIExport(sFileName, lstMonitors(iSelectedSource), dlg.lbSamples, dlg.lbQuantities)
    End If

End Sub


Sub GetAvailableMonitors(ByRef Monitors() As Monitor)
	Dim names() As String

	On Error GoTo Next1
	names = Particle2DMonitorReader.GetMonitorNames()
	AppendMonitors(Monitors, names, Particle2DMonitor)
Next1:
	On Error GoTo Next2
	names = PIC2DMonitorReader.GetMonitorNames()
	AppendMonitors(Monitors, names, PIC2DMonitor)
Next2:
	On Error GoTo Next3
	names = PICPositionMonitorReader.GetMonitorNames()
	AppendMonitors(Monitors, names, PICPositionMonitor)
Next3:
	On Error GoTo Next4
	If Resulttree.DoesTreeItemExist("2D/3D Results\Trajectories") Then
		ReDim names(0 To 0)
		names (0) = "Trajectories"
		AppendMonitors(Monitors, names, Trajectories)
	End If
Next4:
	On Error GoTo 0
End Sub


Sub AppendMonitors(ByRef lst() As Monitor, ByRef names() As String, eType As MonitorType)
	Dim OldSize As Long, AdditionalSize As Long

	On Error GoTo NothingToDo
	AdditionalSize = UBound(names) - LBound(names) + 1

	OldSize = 0

	On Error GoTo EmptyArray2
	OldSize = UBound(lst)
	On Error GoTo 0

EmptyArray2:

	ReDim Preserve lst(OldSize + AdditionalSize)

	Dim Index As Long
	For Index = LBound(names) To UBound(names)
		lst(Index+OldSize).sName = names(Index)
		lst(Index+OldSize).eType = eType
	Next

NothingToDo:
	On Error GoTo 0

End Sub


Sub UpdateDialog()

	Dim MonitorNames() As String
	MonitorNames = GetMonitorNamesList(lstMonitors)
	DlgListBoxArray("lbSource", MonitorNames)
	DlgValue("lbSource", iSelectedSource)

	Select Case lstMonitors(iSelectedSource).eType
		Case Particle2DMonitor
			Particle2DMonitorReader.SelectMonitor(lstMonitors(iSelectedSource).sName)
			PrepareQuantityList(lstQuantities, lstMonitors(iSelectedSource).eType)
			FillSampleSelector(lstSamples, Particle2DMonitorReader.GetNPlanes(), 0)
			DlgText("txtSamples", "Planes:")
		Case PIC2DMonitor
			PIC2DMonitorReader.SelectMonitor(lstMonitors(iSelectedSource).sName)
			PrepareQuantityList(lstQuantities, lstMonitors(iSelectedSource).eType)
			FillSampleSelector(lstSamples, PIC2DMonitorReader.GetNFrames(), 0)
			DlgText("txtSamples", "Frames:")
		Case PICPositionMonitor
			PICPositionMonitorReader.SelectMonitor(lstMonitors(iSelectedSource).sName)
			PrepareQuantityList(lstQuantities, lstMonitors(iSelectedSource).eType)
			FillSampleSelector(lstSamples, PICPositionMonitorReader.GetNFrames(), 0)
			DlgText("txtSamples", "Frames:")
		Case Trajectories
			ParticleTrajectoryReader.LoadTrajectoryData()
			PrepareQuantityList(lstQuantities, lstMonitors(iSelectedSource).eType)
			FillSampleSelector(lstSamples, ParticleTrajectoryReader.GetNTrajectories(), -1)
			DlgText("txtSamples", "Trajectories:")
	End Select

	Dim QuantityNames() As String
	QuantityNames = GetQuantityNamesList(lstQuantities)
	DlgListBoxArray("lbQuantities", QuantityNames)
	DlgListBoxArray("lbSamples", lstSamples)
	DlgValue("lbSamples", iSelectedSample)
End Sub


Function GetMonitorNamesList(ByRef Monitors() As Monitor) As String()
	Dim Result() As String
	ReDim Result(LBound(Monitors) To UBound(Monitors))
	Dim i As Long
	For i = LBound(Monitors) To UBound(Monitors)
		Result(i) = Monitors(i).sName
	Next
	GetMonitorNamesList = Result
End Function


Function GetQuantityNamesList(ByRef Quantities() As Quantity) As String()
	Dim Result() As String
	ReDim Result(LBound(Quantities) To UBound(Quantities))
	Dim i As Long
	For i = LBound(Quantities) To UBound(Quantities)
		Result(i) = GetQuantityLabel(Quantities(i))
	Next
	GetQuantityNamesList = Result
End Function


Function GetQuantityLabel(ByVal q As Quantity) As String
	GetQuantityLabel = q.sName + IIf(q.sComp = "", "", " [" + q.sComp + "]")
End Function


Sub FillSampleSelector(ByRef lst() As String, ByVal nEntries As Long, ByVal Offset As Long)
    ReDim lst(0 To nEntries)
	lst(0) = "--- all ---"
	Dim i As Long
	For i = 1 To nEntries
		lst(i) = Trim(Str(i + Offset))
	Next
End Sub


Sub DoASCIIExport(ByVal sFileName As String, ByVal Source As Monitor, ByVal iSelectedSample As Long, ByRef lstSelectedQuantities() As Integer)

	On Error GoTo FileError
	Open sFileName For Output As #1

	' write data for each sample
	If iSelectedSample = 0 Then
		Dim iSample As Long
		For iSample = 1 To UBound(lstSamples)
			DoASCIIExportPerSample(Source, iSample, lstSelectedQuantities)
		Next
	Else
		DoASCIIExportPerSample(Source, iSelectedSample, lstSelectedQuantities)
	End If

  	Close #1
    
    ReportInformationToWindow("The exported data hase been stored in: " + sFileName + ".")

	On Error GoTo 0
	Exit Sub

FileError:
	On Error GoTo 0
	MsgBox("Unable to write to " + sFileName, vbOkOnly + vbCritical)
End Sub


Sub WriteHeader(ByVal Source As Monitor, ByVal sSample As String, ByRef lstSelectedQuantities() As Integer)
	Print #1, sCommentChar;Source.sName
	Print #1, sCommentChar;sSample

	Dim iQuantity As Long
	Print #1, sCommentChar;
	For iQuantity = LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities)
		Print #1, GetQuantityLabel(lstQuantities(lstSelectedQuantities(iQuantity))) + sSeparator;
	Next
	Print #1
End Sub


Sub AppendQuantity(ByRef lst() As Quantity, ByVal sName As String, sComp As String)
	Dim OldSize As Long
	OldSize = -1

	On Error GoTo EmptyArray2
	OldSize = UBound(lst)
	On Error GoTo 0

EmptyArray2:

	ReDim Preserve lst(OldSize + 1)
	lst(UBound(lst)).sName = sName
	lst(UBound(lst)).sComp = sComp

	On Error GoTo 0
End Sub


Sub PrepareQuantityList(ByRef lstQuantities() As Quantity, ByVal eSourceType As MonitorType)
	Dim lstQuantityNames() As String
	Dim lstQuantityTypes() As Quantity

	Select Case eSourceType
		Case Particle2DMonitor
			lstQuantityNames = Particle2DMonitorReader.GetQuantityNames()
		Case PIC2DMonitor
			lstQuantityNames = PIC2DMonitorReader.GetQuantityNames()
		Case PICPositionMonitor
			lstQuantityNames = PICPositionMonitorReader.GetQuantityNames()
		Case Trajectories
			lstQuantityNames = ParticleTrajectoryReader.GetQuantityNames()
	End Select

	Dim iQuantity As Long

	For iQuantity = LBound(lstQuantityNames) To UBound(lstQuantityNames)
		Dim sName As String

		sName = lstQuantityNames(iQuantity)

		Dim c As Long

		Select Case eSourceType
			Case Particle2DMonitor
				c = Particle2DMonitorReader.GetNComponents(sName)
			Case PIC2DMonitor
				c = PIC2DMonitorReader.GetNComponents(sName)
			Case PICPositionMonitor
				c = PICPositionMonitorReader.GetNComponents(sName)
			Case Trajectories
				c = ParticleTrajectoryReader.GetNComponents(sName)
		End Select

		If c = 1 Then
			AppendQuantity(lstQuantityTypes, sName, "")
		ElseIf c = 3 Then
			AppendQuantity(lstQuantityTypes, sName, "X")
			AppendQuantity(lstQuantityTypes, sName, "Y")
			AppendQuantity(lstQuantityTypes, sName, "Z")
			AppendQuantity(lstQuantityTypes, sName, "ABS (XYZ)")
		End If
	Next

	lstQuantities = lstQuantityTypes
End Sub


Sub DoASCIIExportPerSample(ByVal Source As Monitor, ByVal iSelectedSample As Long, ByRef lstSelectedQuantities() As Integer)

	Dim nLines As Long
	Dim iQuantity As Long
	Dim q As Quantity

	Dim dValues() As Single
	Dim dColumn() As Single
	Dim iColumn() As Long
	Dim sSample As String

	Select Case Source.eType
		Case Particle2DMonitor
			Particle2DMonitorReader.SelectPlane(iSelectedSample-1)
			sSample = "Plane " + Trim(Str(iSelectedSample)) ' TODO: also print data from Particle2DMonitorReader.GetNormal()/GetPlaneDistance()
			WriteHeader(Source, sSample, lstSelectedQuantities)
			nLines = Particle2DMonitorReader.GetNParticles()

			If nLines > 0 Then
				ReDim dValues(LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities), nLines)

				For iQuantity = LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities)
					q = lstQuantities(lstSelectedQuantities(iQuantity))
					If q.sName = "EmissionID" Or q.sName = "ParticleID" Then
						iColumn = Particle2DMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumnInt(dValues, iColumn, iQuantity)
					Else
						dColumn = Particle2DMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumn(dValues, dColumn, iQuantity)
					End If
				Next
			End If


		Case PIC2DMonitor
			PIC2DMonitorReader.SelectFrame(iSelectedSample-1)
			sSample = "Frame " + Trim(Str(iSelectedSample)) ' TODO: also print data from PIC2DMonitorReader.GetFrameInfo()
			WriteHeader(Source, sSample, lstSelectedQuantities)
			nLines = PIC2DMonitorReader.GetNParticles()

			If nLines > 0 Then
				ReDim dValues(LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities), nLines)

				For iQuantity = LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities)
					q = lstQuantities(lstSelectedQuantities(iQuantity))
					If q.sName = "EmissionID" Or q.sName = "ParticleID" Then
						iColumn = PIC2DMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumnInt(dValues, iColumn, iQuantity)
					Else	
						dColumn = PIC2DMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumn(dValues, dColumn, iQuantity)
					End If
				Next
			End If


		Case PICPositionMonitor
			PICPositionMonitorReader.SelectFrame(iSelectedSample-1)
			sSample = "Frame " + Trim(Str(iSelectedSample)) ' TODO: also print data from PICPositionMonitorReader.GetFrameInfo()
			WriteHeader(Source, sSample, lstSelectedQuantities)
			nLines = PICPositionMonitorReader.GetNParticles()

			If nLines > 0 Then
				ReDim dValues(LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities), nLines)

				For iQuantity = LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities)
					q = lstQuantities(lstSelectedQuantities(iQuantity))
					If q.sName = "EmissionID" Or q.sName = "ParticleID" Then
						iColumn = PICPositionMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumnInt(dValues, iColumn, iQuantity)
					Else					
						dColumn = PICPositionMonitorReader.GetQuantityValues(q.sName, q.sComp)
						SetColumn(dValues, dColumn, iQuantity)
					End If
				Next
			End If


		Case Trajectories
			ParticleTrajectoryReader.SelectTrajectory(iSelectedSample-1)
			sSample = "Trajectory " + Trim(Str(iSelectedSample-1))
			WriteHeader(Source, sSample, lstSelectedQuantities)
			nLines = ParticleTrajectoryReader.GetNParticles()

			If nLines > 0 Then
				ReDim dValues(LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities), nLines)

				For iQuantity = LBound(lstSelectedQuantities) To UBound(lstSelectedQuantities)
					q = lstQuantities(lstSelectedQuantities(iQuantity))
					If q.sName = "EmissionID" Or q.sName = "ParticleID" Then
						iColumn = ParticleTrajectoryReader.GetQuantityValues(q.sName, q.sComp)
						SetColumnInt(dValues, iColumn, iQuantity)
					Else
						dColumn = ParticleTrajectoryReader.GetQuantityValues(q.sName, q.sComp)
						SetColumn(dValues, dColumn, iQuantity)
					End If
				Next
			End If

	End Select

	Dim iLine As Long
	For iLine = 0 To nLines-1
		For iQuantity = LBound(dValues,1) To UBound(dValues,1)
			Print #1, dValues(iQuantity, iLine); sSeparator;
		Next
		Print #1
	Next

	Print #1
End Sub


Sub SetColumn(ByRef dValues() As Single, ByRef dColumn() As Single, ByVal iQuantity As Long)
	Dim iLine As Long

	For iLine = LBound(dColumn) To UBound(dColumn)
		dValues(iQuantity, iLine) = dColumn(iLine)
	Next
End Sub

Sub SetColumnInt(ByRef dValues() As Single, ByRef iColumn() As Long, ByVal iQuantity As Long)
	Dim iLine As Long

	For iLine = LBound(iColumn) To UBound(iColumn)
		dValues(iQuantity, iLine) = CSng(iColumn(iLine))
	Next
End Sub


' This function is called whenever something changes in the dialog.
' It can be used to update dialog state etc.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		GetAvailableMonitors(lstMonitors)
		iSelectedSource = 0
		iSelectedSample = 0
		UpdateDialog()

	Case 2 ' Value changing or button pressed
		Select Case (DlgItem$)
			Case "lbSource"
				iSelectedSource = DlgValue("lbSource")
				iSelectedSample = 0
				UpdateDialog()
			Case "lbSample"
				iSelectedSample = DlgValue("lbSample")
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
