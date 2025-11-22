' *EMC / Renorm Probes+Voltages to EMC-frequency source
' !!! Do not change the line above !!!
' macro.835
' ================================================================================================
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 04-Aug-2017 ube: added hint to AC combine Task, which should make this macro obsolete
' 02-Aug-2017 ube: apd files removed, ascii file handing removed and replaced by result-1d objects
' 15-Dec-2009 ube: if SQL database exists ("Storage.sdb"), then extract it first and then compare
' 30-Jul-2009 ube: Split replaced by CSTSplit, since otherwise compeating with standrad VBA-Split function
' 26-Nov-2008 ube: converted to new cst file structure
' 21-Oct-2005 imu: Included into Online Help
' 20-Jul-2002 ube: first version
'------------------------------------
Option Explicit
'#include "vba_globals_all.lib"

Public ProjectBasepath As String

'-----------------------------------------------------------------------------------------------------------------------------
Sub Main
	
	Dim FrqUnit(4) As String 
	
	FrqUnit(0) = "Hz"
	FrqUnit(1) = "kHz"
	FrqUnit(2) = "MHz"
	FrqUnit(3) = "GHz"
	
	Begin Dialog UserDialog 460,273,"Renorm to EMC voltage source",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,70,420,84,"EMC-Source file name   (col1=freq / col2=scaling factor)",.GroupBox1
		TextBox 30,91,400,21,.emcfile
		GroupBox 20,7,420,49,"Only for discrete ports, defined as Power-Sources",.GroupBox2
		PushButton 340,119,80,21,"Browse...",.Browse 'open Window for filesearch
		CheckBox 40,28,260,14,"Norm to 1V voltage at discrete port",.CheckDiscrPortVoltage
		TextBox 310,28,40,21,.idport
		DropListBox 210,119,70,91,FrqUnit(),.frequ_scale
		Text 50,126,160,14,"Frq-Unit of external file:",.Text2
		GroupBox 20,161,420,49,"ResultTree-Label",.GroupBox3
		TextBox 30,182,400,21,.treelabel
		OKButton 20,224,90,21
		CancelButton 120,224,90,21
		PushButton 220,224,90,21,"Help",.Help
		Text 30,252,410,14,"Note: AC-Combine Task can import frq-dependent source data.",.Text1
	End Dialog

	Dim dlg As UserDialog
	If (Dialog(dlg) >= 0) Then Exit All
	
	' now extract sql database "Storage.sdb" into ascii files, this is necessary to find probe and voltages below
	UnpackDataToSigFiles (GetProjectBaseName)

	Dim nn As Long
	Dim ii As Long, ii2 As Long
	Dim sTreeLabel As String
	sTreeLabel = dlg.treelabel
	sTreeLabel = Right(sTreeLabel, Len(sTreeLabel)-InStrRev(sTreeLabel,"\"))

	Dim oScalePort  As Object
	Dim oScaleFile  As Object
	Dim oScaleTotal As Object

	Set oScalePort  = Result1D("")
	Set oScaleFile  = Result1D("")
	Set oScaleTotal = Result1D("")


	If dlg.CheckDiscrPortVoltage Then
	
		' ============================================
		' === Calculate port scaling factor 1/voltage
		' ============================================

		Dim oDiscPortVoltComplex As Object, spuc As String

		spuc = "puc" + dlg.idport + "(" + dlg.idport + ")"

		On Error Resume Next
		Set oDiscPortVoltComplex = Result1DComplex(spuc)
		On Error GoTo 0

		If oDiscPortVoltComplex Is Nothing Then
			MsgBox "Voltage at discrete port "+dlg.idport+" not found/monitored." + vbCrLf + vbCrLf + _
					"Please switch on Option ""Monitor Voltage and Current"""  + vbCrLf + vbCrLf + _
					"at discrete Port " + dlg.idport  + " and recalculate EM fields.",vbExclamation
			Exit All
		End If

		With oDiscPortVoltComplex.Magnitude
			For ii = 0 To .GetN-1	' Read all frequency points; index of first point is zero.
				If (.GetY(ii) <> 0) Then
					oScalePort.AppendXY .GetX(ii), 1/.GetY(ii)
				End If
			Next ii
		End With

		oScalePort.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oScalePort.Title "Frequency Dependent Scaling Factor"
		AddPlotToTree_LIB(oScalePort, "\Scaling Factor\1 / port voltage")

	End If
	
	Dim sline As String
	Dim string_item(10) As String
	Dim cst_nitems As Integer
	Dim dValue(2) As Double
	Dim iReal As Integer
	Dim dfrqscale As Double

	If (dlg.emcfile <> "") Then
		' ================================================
		' === Import scaling factor of frequency source
		' ================================================
		
		dfrqscale = 1000^dlg.frequ_scale / Units.GetFrequencyUnitToSI

		Open dlg.emcfile For Input As #12
			While Not EOF(12)
				Line Input #12, sline
				cst_nitems = CSTSplit (sline, string_item)
				
				iReal = 0
				For ii = 0 To cst_nitems-1
					On Error GoTo NoRealNumber
						dValue(iReal) = RealVal(string_item(ii))
						iReal = iReal+1
						If iReal = 2 Then
							Exit For
						End If
					NoRealNumber:
				Next ii
				
				' Disable Error Handling
				On Error GoTo 0
				
				oScaleFile.AppendXY dfrqscale*dValue(0), dValue(1)
			Wend
		Close #12
		
		oScaleFile.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oScaleFile.Title "Frequency Dependent Scaling Factor"
		AddPlotToTree_LIB(oScaleFile, "\Scaling Factor\File: " + ShortName(dlg.emcfile))

	End If
	
		
	' ================================================
	' === Calulcate total scaling factor
	' ================================================

	If oScaleFile.GetN = 0 Then
		If oScalePort.GetN = 0 Then
			MsgBox "Nothing todo. Macro stops.", vbInformation
			Exit All
		Else
			Set oScaleTotal = oScalePort
		End If
	Else
		If oScalePort.GetN = 0 Then
			Set oScaleTotal = oScaleFile
		Else
			' both scaling factors have to be multiplied
			' frequency points of file are taken with higher priority

			With oScaleFile
				For ii = 0 To .GetN-1	' Read all frequency points; index of first point is zero.
					ii2 = oScalePort.GetClosestIndexFromX(.GetX(ii))
					oScaleTotal.AppendXY .GetX(ii), .GetY(ii)*oScalePort.GetY(ii2)
				Next ii
			End With

			oScaleTotal.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
			oScaleTotal.Title "Frequency Dependent Scaling Factor"
			AddPlotToTree_LIB(oScaleTotal, "\Scaling Factor\Total Scale")

		End If	
	End If
	
	Dim myfilepath$, myfilepattern$, myfilepath_short$, myprobename$, myprobename_short$
			
	myfilepath    = GetProjectBaseName
	myfilepath_short = BaseName(GetProjectBaseName+".mod")

	' ==========================================
	' ================ Renorm all probes
	' ==========================================

	Dim oRescaled1DResult As Object
	Dim oRescaled1DResultdB As Object

	Dim sProbe As Object, sProbeComplex As Object
	myfilepattern = "*.prc"
	myprobename = FindFirstFile(myfilepath, myfilepattern, False) 'See if a Probe exist

	While (myprobename <> "")
		myprobename_short = myprobename
		RemoveLastChars(myprobename_short,4)

		Set sProbeComplex = Result1DComplex(myprobename)
		Set sProbe = sProbeComplex.Magnitude

		Set oRescaled1DResult   = Result1D("") ' this is resetting to zero length
		Set oRescaled1DResultdB = Result1D("") ' this is resetting to zero length

		With oScaleTotal
			For ii = 0 To .GetN	-1	' Read all frequency points; index of first point is zero.
				ii2 = sProbe.GetClosestIndexFromX(.GetX(ii))
				oRescaled1DResult.AppendXY   .GetX(ii), .GetY(ii)*sProbe.GetY(ii2)
				If (.GetY(ii)*sProbe.GetY(ii2))>0 Then
					oRescaled1DResultdB.AppendXY .GetX(ii), 20*Log(.GetY(ii)*sProbe.GetY(ii2))/Log(10)
				End If
			Next ii
		End With

		oRescaled1DResult.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oRescaled1DResult.Title "Renormed Probe Amplitudes linear"
		AddPlotToTree_LIB(oRescaled1DResult, dlg.treelabel + "\Probes linear\" + myprobename_short)

		oRescaled1DResultdB.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oRescaled1DResultdB.Title "Renormed Probe Amplitudes in dB"
		AddPlotToTree_LIB(oRescaled1DResultdB, dlg.treelabel + "\Probes in dB\" + myprobename_short)

		SelectTreeItem dlg.treelabel + "\Probes linear"

		myprobename = FindNextFile
	Wend
	
	' ==========================================
	' ================ Renorm all voltages
	' ==========================================

	myfilepattern = "*.vrc"
	myprobename = FindFirstFile(myfilepath, myfilepattern, False) 'See if a Probe exist 
	
	While (myprobename <> "")
		myprobename_short = myprobename
		RemoveLastChars(myprobename_short,4)

		Set sProbeComplex = Result1DComplex(myprobename)
		Set sProbe = sProbeComplex.Magnitude

		Set oRescaled1DResult   = Result1D("") ' this is resetting to zero length
		Set oRescaled1DResultdB = Result1D("") ' this is resetting to zero length

		With oScaleTotal
			For ii = 0 To .GetN	-1	' Read all frequency points; index of first point is zero.
				ii2 = sProbe.GetClosestIndexFromX(.GetX(ii))
				oRescaled1DResult.AppendXY   .GetX(ii), .GetY(ii)*sProbe.GetY(ii2)
				If (.GetY(ii)*sProbe.GetY(ii2))>0 Then
					oRescaled1DResultdB.AppendXY .GetX(ii), 20*Log(.GetY(ii)*sProbe.GetY(ii2))/Log(10)
				End If
			Next ii
		End With

		oRescaled1DResult.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oRescaled1DResult.Title "Renormed Voltage Amplitudes linear"
		AddPlotToTree_LIB(oRescaled1DResult, dlg.treelabel + "\Probes linear\" + myprobename_short)

		oRescaled1DResultdB.SetXLabelAndUnit( "Frequency" , Units.GetUnit("Frequency"))
		oRescaled1DResultdB.Title "Renormed Voltage Amplitudes in dB"
		AddPlotToTree_LIB(oRescaled1DResultdB, dlg.treelabel + "\Probes in dB\" + myprobename_short)

		SelectTreeItem dlg.treelabel + "\Voltages linear"
	
		myprobename = FindNextFile
	Wend
	
End Sub

Function DialogFunc%(DlgItem As String, Action As Integer, SuppValue As Integer)
    Dim file As String
    Dim basepath As String

    Debug.Print "Action=";Action
    Debug.Print DlgItem
    Debug.Print "SuppValue=";SuppValue

    Select Case Action
    Case 1 ' Dialog box initialization
            DlgEnable "idport",False
            DlgText "idport","1"
            DlgText "treelabel","1D Results\EMC-Source"

        'Beep
    Case 2 ' Value changing or button pressed
      Select Case DlgItem
		Case "Help"
			StartHelp "common_preloadedmacro_emc_renorm_probes+voltages_to_emc-frequency_source"
			DialogFunc = True
        Case "CheckDiscrPortVoltage"
        		If SuppValue = 0 Then
		            DlgEnable "idport",False
        		Else
		            DlgEnable "idport",True
        		End If
        Case "Browse" 
	          file = DlgText("emcfile")       
	          Debug.Print file	
	          ' Open dialog for File-selection
		      file = GetFilePath(ShortName(file),"inp;dat;txt",DirName(file),"Select EMC Source File",0)
			  If (file = "") Then
				DlgText "emcfile", file
		      ElseIf (LCase$(DirName(file)) = LCase$(ProjectBasepath)) Then 	'check for file-extensions
			    DlgText "emcfile", DirName(file)+"\"+ShortName(file)
			  ElseIf (file <> "") Then
				DlgText "emcfile", file
			  End If
				If (file <> "") Then
              		DlgText "treelabel","1D Results\EMC-Source "+ShortName(file)
			  	End If

			  'DlgText "subnametext", ShortName(file) 	'update Curve-entryname to dialog
	          DialogFunc = True				'do not exit the dialog
        Case Else
        	'
      End Select
    Case 3 ' Combo or text value changed
      'MsgBox "combo-box"
      DialogFunc = True 				'do not exit the dialog
	  'm = DlgText ("ComboBox1")
	  'MsgBox m
    Case 4 ' Focus changed
       Debug.Print "DlgFocus=""";DlgFocus();""""
    End Select
End Function
