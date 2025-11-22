' Wire meshing

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 07-Nov-2014 ube: First version
' ================================================================================================

Sub Main ()


	Dim DisableImprovedMeshAtWireSurfaceJunction  As Boolean
	
	With MeshSettings
  		.SetMeshType "Srf"
		DisableImprovedMeshAtWireSurfaceJunction = .Get("DisableImprovedMeshAtWireSurfaceJunction")
	End With

	Begin Dialog UserDialog 570,161,"Enable improved wire meshing for Surface Meshing" ' %GRID:10,7,1,1
		GroupBox 10,7,540,63,"INFO",.GroupBox4
		Text 30,25,510,42,"This script enables/disables an improved surface meshing of wires touching metallic objects. All the following simulations related to this project will use this setting.",.Text1
		
		CheckBox 10,84,290,21,"Disabled",.CheckBox1

		OKButton 30,133,100,21
		CancelButton 150,133,100,21
		
	End Dialog

	Dim dlg As UserDialog

	dlg.CheckBox1 = DisableImprovedMeshAtWireSurfaceJunction

	If (Dialog(dlg) = 0) Then Exit All

	'do the dialog


	Dim sCommand As String


	sCommand = ""
	sCommand = sCommand + "With MeshSettings " + vbLf
	sCommand = sCommand + ".SetMeshType ""Srf""" + vbLf

	If dlg.CheckBox1 Then

		sCommand = sCommand + ".Set ""DisableImprovedMeshAtWireSurfaceJunction"""+",True" + vbLf

	Else

		sCommand = sCommand + ".Set ""DisableImprovedMeshAtWireSurfaceJunction"""+",False" + vbLf

	End If

		sCommand = sCommand + "End With"
		AddToHistory "set mesh properties for wires (Surface)", sCommand


End Sub


