'#Language "WWB-COM"

' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 05-Aug-2011 ube: First version
' ================================================================================================
Option Explicit

Sub Main

	Dim minQuality As String
	Dim maxQuality As String

	minQuality = "0.0"
	maxQuality = "1.0"

	Begin Dialog UserDialog 400,112,"Plot Mesh in Quality Range" ' %GRID:10,7,1,1
		PushButton 290,14,100,21,"Plot",.PushButton1
		CancelButton 290,42,100,21
		GroupBox 20,14,250,84,"Mesh Quality Range",.GroupBox1
		TextBox 160,63,90,21,.maxQuality
		TextBox 160,35,90,21,.minQuality
		Text 34,37,110,21,"Minimum quality:",.Text1
		Text 34,65,110,21,"Maximum quality:",.Text2
	End Dialog
	Dim dlg As UserDialog

	dlg.minQuality = minQuality
	dlg.maxQuality = maxQuality

	If Dialog(dlg)=0 Then Exit All

	minQuality = dlg.minQuality
	maxQuality = dlg.maxQuality

	Dim dMinQuality As Double
	Dim dMaxQuality As Double

    dMinQuality = Evaluate(minQuality)
    dMaxQuality = Evaluate(maxQuality)

	Mesh.PlotElemQualityRange (dMinQuality, dMaxQuality)

End Sub
