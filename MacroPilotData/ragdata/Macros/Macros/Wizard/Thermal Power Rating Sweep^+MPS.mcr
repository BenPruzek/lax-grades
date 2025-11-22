Option Explicit
'-----------------------------------------------------------------------------
' This macro performs a sweep of the power scaling factor without rerunning 
' the MWS-Transient simulation for any new power level.
'
' As result the maximum temperature in degree celsius is plotted versus power scaling factor
'
' Also, this procedure can be automated for different frequencies, if monitor 
' (with automatic default naming !) are defined.
'
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ============================================================================
' History of Changes
' ------------------
' 22-Feb-2010 ube: 2010 result tree uses "Temperature [THs]" instad of earlier [Th]
' 26-Oct-2007 ube: converted to 2008
' 06-Jul-2007 ube: initial version
' ============================================================================
Sub Main

	Begin Dialog UserDialog 320,224,"Power Rating Sweep" ' %GRID:10,7,1,1
		GroupBox 20,7,290,70,"Monitor + Frequencies",.GroupBox1
		GroupBox 20,84,290,105,"Power Scaling Factor",.GroupBox2
		Text 40,112,90,14,"Min",.Text1
		Text 40,140,90,14,"Max",.Text2
		Text 40,168,70,14,"Step size",.Text3
		CheckBox 40,28,260,14,"Sweep over all available frequencies",.AllHFields
		TextBox 120,105,130,21,.PStart
		TextBox 120,133,130,21,.PStop
		TextBox 120,161,130,21,.PStepsize
		PushButton 30,196,90,21,"Start Sweep",.PushButton1
		CancelButton 130,196,90,21
		Text 40,54,110,14,"Excitation String",.Text4
		TextBox 160,49,80,21,.excit
	End Dialog
	Dim dlg As UserDialog

	dlg.AllHFields = 0
	dlg.excit = "[1]"
	dlg.PStart 		=  "800"
	dlg.PStop 		= "1600"
	dlg.PStepsize 	=  "200"

	If (Dialog(dlg) = 0) Then Exit All

	Dim sExcit As String, sFileExt As String, nh As Integer, ih As Integer, i0 As Integer, i1 As Integer, sfrq As String
	sExcit = dlg.excit

	Dim hname As String
	Dim harray(99) As String

	hname = Resulttree.GetFirstChildName ("2D/3D Results\H-Field")
	nh = 0

	If (dlg.AllHFields = 1) Then
		While hname <> ""
			If (Right(hname,Len(sExcit)) = sExcit) Then
				nh = nh + 1
				harray(nh) = hname
			End If
			hname=Resulttree.GetNextItemName (hname)
		Wend
	
		i0 = 1+InStr(sExcit,"[")
		i1 = InStr(sExcit,"]")
		sFileExt = Mid(sExcit,i0,i1-i0)
		If InStr(sFileExt,",") = 0 Then
			sFileExt = sFileExt + ",1"
		End If

		If (nh = 0) Then
			MsgBox "No H-Field found. Please check excitation string." + vbCrLf + "Exit all.",vbCritical
			Exit All
		End If
	End If

	Dim dpfac_min  As Double
	Dim dpfac_max  As Double
	Dim dpfac_step As Double
	Dim dfac As Double

	dpfac_min  = Evaluate(dlg.PStart)
	dpfac_max  = Evaluate(dlg.PStop)
	dpfac_step = Evaluate(dlg.PStepsize)

	Dim xcoord As Double, ycoord As Double, zcoord As Double
	Dim cst_max_temperature As Double

	If (dlg.AllHFields = 0) Then nh=1

	For ih = 1 To nh

		If (dlg.AllHFields = 1) Then
			hname = harray(ih)
			i0 = 1+InStr(hname,"(")
			i1 = InStr(hname,")")
			sfrq = Mid(hname,i0,i1-i0)
		Else
			sfrq = "result"
		End If
		' MsgBox sfrq

		Dim r1dtmp As Object
		Set r1dtmp = Result1D("")

		For dfac = dpfac_min To dpfac_max STEP dpfac_step

			With ThermalSourceParameter
				If (dlg.AllHFields = 1) Then
					If (SelectTreeItem("2D/3D Results\H-Field\h-field ("+sfrq+") "+sExcit)) Then
						.SurfaceSourceFieldName "h-field ("+sfrq+")_"+sFileExt
					Else
						MsgBox "Problem with surface sources, h-field does not exist:"+vbCrLf + _
								"2D/3D Results\H-Field\h-field ("+sfrq+") "+sExcit , vbCritical
					End If
					If (SelectTreeItem("2D/3D Results\Current Density\current ("+sfrq+") "+sExcit)) Then
						.VolumeSourceFieldName "current ("+sfrq+")_"+sFileExt
					Else
						If (SelectTreeItem("2D/3D Results\E-Field\e-field ("+sfrq+") "+sExcit)) Then
							.VolumeSourceFieldName "e-field ("+sfrq+")_"+sFileExt
						Else
							MsgBox "Problem with volume sources, neither current density nor e-field exists:"+vbCrLf + _
									"2D/3D Results\Current Density\current ("+sfrq+") "+sExcit   +vbCrLf + _
									"2D/3D Results\E-Field\e-field ("+sfrq+") "+sExcit , vbCritical
						End If
					End If
				End If
				.Factor dfac
				.AddSource
			End With

			ThermalSolver.Start

			SelectTreeItem "2D/3D Results\Temperature [THs]"
			Resulttree.UpdateTree
		    Resulttree.RefreshView
			Wait 0.2
			Plot.Update
			Wait 0.8

			cst_max_temperature = GetFieldPlotMaximumPos(xcoord, ycoord, zcoord)
			cst_max_temperature = cst_max_temperature - 273.15

			r1dtmp.AppendXY  dfac, cst_max_temperature

		Next dfac

		With r1dtmp
			.Title "Maximum Temperature"
			.SetYLabelAndUnit "Max Temperature" , "degC"
			.SetXLabelAndUnit "Power Scaling Factor","1"
			Wait 0.5
			.Save "PowerRating_"+sfrq+".sig"
			Wait 0.5
			.AddToTree "1D Results\Power Rating\"+sfrq
			Wait 0.5
			Resulttree.UpdateTree
		    Resulttree.RefreshView
		End With
	Next ih

	SelectTreeItem "1D Results\Power Rating"

	MsgBox "Power rating Sweep finished.",vbInformation

End Sub
