'#Language "WWB-COM"

' This macro allows to generate a spatially varying material

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 11-Dec-2014 fsr: Added initial values for input text fields
' 05-Nov-2014 fsr: Initial version

Option Explicit

Public sModelChoices() As String, sModelProperties() As String
' array of choices. Format: Displayed name|Internal name|property1|property2|property3
Public Const SpatiallyVaryingMaterialModelsArray = Array("Constant|Constant|Value:||", _
														"Luneburg|Luneburg|Value at center:|Value on surface:|", _
														"Power Law|PowerLaw|Value at center:|Value on surface:|Model order:", _
														"Graded Index|GradedIndex|Value at center:|Gradient value:|")
Sub Main

	Dim i As Long

	ReDim sModelChoices(UBound(SpatiallyVaryingMaterialModelsArray))
	ReDim sModelProperties(2)
	For i = 0 To UBound(sModelChoices)
		sModelChoices(i) = Split(SpatiallyVaryingMaterialModelsArray(i), "|")(0)
	Next

	Begin Dialog UserDialog 360,294,"Spatially Varying Material (Optical)",.DialogFunction ' %GRID:10,7,1,1
		Text 20,14,100,14,"Material name:",.Text2
		TextBox 120,7,210,21,.MaterialNameT
		GroupBox 20,42,320,210,"Refractive Index n",.GroupBox1
		Text 40,70,50,14,"Model:",.Text1
		DropListBox 100,63,220,170,sModelChoices(),.RefractiveIndexModelDLB
		Text 40,98,220,14,"Property1:",.RefractiveIndexProperty1L
		TextBox 40,119,280,21,.RefractiveIndexProperty1T
		Text 40,147,220,14,"Property2:",.RefractiveIndexProperty2L
		TextBox 40,168,280,21,.RefractiveIndexProperty2T
		Text 40,196,220,14,"Property3:",.RefractiveIndexProperty3L
		TextBox 40,217,280,21,.RefractiveIndexProperty3T
		OKButton 140,259,90,21
		CancelButton 240,259,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.RefractiveIndexProperty1T = "1.5"
	dlg.RefractiveIndexProperty2T = "1"
	dlg.RefractiveIndexProperty3T = "1"

	If Dialog(dlg, -1) = 0 Then ' user pressed cancel
		Exit All
	End If


End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean

	Dim i As Long
	Dim nCurrentlySelectedModel As Long

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("MaterialNameT", "Spatially Varying Material")
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "OK"
				Dim sMaterialName As String
				Dim sRefractiveIndexModel As String, sRefractiveIndexProperty1 As String, sRefractiveIndexProperty2 As String, sRefractiveIndexProperty3 As String
				Dim sExtinctionCoeffModel As String, sExtinctionCoeffProperty1 As String, sExtinctionCoeffProperty2 As String, sExtinctionCoeffProperty3 As String

				sMaterialName = DlgText("MaterialNameT")
				sRefractiveIndexModel = Split(SpatiallyVaryingMaterialModelsArray(DlgValue("RefractiveIndexModelDLB")), "|")(1)
				sRefractiveIndexProperty1 = DlgText("RefractiveIndexProperty1T")
				sRefractiveIndexProperty2 = DlgText("RefractiveIndexProperty2T")
				sRefractiveIndexProperty3 = DlgText("RefractiveIndexProperty3T")
				sExtinctionCoeffModel = "" 'Split(SpatiallyVaryingMaterialModels(DlgValue("ExtinctionCoeffModelDLB)), "|")(1)
				sExtinctionCoeffProperty1 = "" 'DlgText("ExtinctionCoeffProperty1T")
				sExtinctionCoeffProperty2 = "" 'DlgText("ExtinctionCoeffProperty2T")
				sExtinctionCoeffProperty3 = "" 'DlgText("ExtinctionCoeffProperty3T")

				CreateSpatiallyVaryingMaterial_Optical(sMaterialName, sRefractiveIndexModel, sRefractiveIndexProperty1, sRefractiveIndexProperty2, sRefractiveIndexProperty3, _
																	sExtinctionCoeffModel, sExtinctionCoeffProperty1, sExtinctionCoeffProperty2, sExtinctionCoeffProperty3)
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select

	' Update dialog entries depending on DLB selection
	nCurrentlySelectedModel = DlgValue("RefractiveIndexModelDLB")
	For i = 0 To UBound(sModelProperties)
		sModelProperties(i) = Split(SpatiallyVaryingMaterialModelsArray(nCurrentlySelectedModel), "|")(i+2)
		DlgText("RefractiveIndexProperty"+CStr(i+1)+"L", sModelProperties(i))
		DlgEnable("RefractiveIndexProperty"+CStr(i+1)+"L", sModelProperties(i)<>"")
		DlgEnable("RefractiveIndexProperty"+CStr(i+1)+"T", sModelProperties(i)<>"")
	Next

End Function

Function CreateSpatiallyVaryingMaterial_Optical(sMaterialName As String, sRefractiveIndexModel As String, sRefractiveIndexProperty1 As String, sRefractiveIndexProperty2 As String, sRefractiveIndexProperty3 As String, _
													sExtinctionCoeffModel As String, sExtinctionCoeffProperty1 As String, sExtinctionCoeffProperty2 As String, sExtinctionCoeffProperty3 As String) As Integer

	' Creates a spatially varying material; currently, only case k=0 has been implemented

	Dim sHistoryString As String

	sHistoryString = ""
	sHistoryString = sHistoryString + "With Material" + vbNewLine
	sHistoryString = sHistoryString + ".Reset" + vbNewLine
	sHistoryString = sHistoryString + ".Name " + Chr(34) + sMaterialName + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + ".ResetSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + vbNewLine
	sHistoryString = sHistoryString + ".SpatiallyVaryingMaterialModel " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + sRefractiveIndexModel + Chr(34) + vbNewLine
	Select Case sRefractiveIndexModel
		Case "Constant"
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_constant" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty1)^2) + Chr(34) + vbNewLine
		Case "Luneburg"
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_center" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty1)^2) + Chr(34) + vbNewLine
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_surface" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty2)^2) + Chr(34) + vbNewLine
		Case "PowerLaw"
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_axis" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty1)^2) + Chr(34) + vbNewLine
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_cladding" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty2)^2) + Chr(34) + vbNewLine
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_profile" + Chr(34) + ", " + Chr(34) + sRefractiveIndexProperty3 + Chr(34) + vbNewLine
		Case "GradedIndex"
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_axis" + Chr(34) + ", " + Chr(34) + CStr(Evaluate(sRefractiveIndexProperty1)^2) + Chr(34) + vbNewLine
			sHistoryString = sHistoryString + ".AddSpatiallyVaryingMaterialParameter " + Chr(34) + "eps" + Chr(34) + ", " + Chr(34) + "value_gradient" + Chr(34) + ", " + Chr(34) + sRefractiveIndexProperty2 + Chr(34) + vbNewLine
		Case Else
			ReportError("Unknown material model: " + sRefractiveIndexModel)
	End Select
	sHistoryString = sHistoryString + ".Create" + vbNewLine
	sHistoryString = sHistoryString + "End With" + vbNewLine

	AddToHistory("define material: " + sMaterialName, sHistoryString)

End Function
