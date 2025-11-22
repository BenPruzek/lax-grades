' --------------------------------------------------------------------------------------------------------
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 24-Jul-2019 ube: typo in dialogue
' 05-Jan-2016 ube: first version
' ---------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main ()
	Begin Dialog UserDialog 450,119,"Enable Disable Pins" ' %GRID:10,7,1,1
		GroupBox 10,7,430,70,"",.GroupBox1
		Text 20,21,400,28,"For a large number of ports, block pins in schematic might slow down file opening. Here you can activate / deactivate pins.",.Text3
		OKButton 20,91,90,21
		CancelButton 120,91,90,21
		CheckBox 40,56,280,14,"Show Pins at Schematic Block",.CheckBox1
	End Dialog
	Dim dlg As UserDialog

	dlg.CheckBox1 = IIf(DS.SchematicBlockPinsEnabled,1,0)

	If (Dialog(dlg)<>0) Then
		DS.EnableSchematicBlockPins dlg.CheckBox1
	End If
End Sub
