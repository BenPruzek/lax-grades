'#Language "WWB-COM"

' ==============================================================================================================================================================
' This macro loads a 3D field result into the project. The mesh needs to be identical with the mesh set up in the project.'
'
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ==============================================================================================================================================================
' History of Changes
' ------------------
' 24-Apr-2013 fsr: disabled FF renormalization before import: unnecessary and causes problems with plane wave results
' 18-Apr-2013 fsr: power flow monitors are now stored under "Power Flow" in the tree, instead of "Powerflow";
'					monitor names containing periods not cut off anymore when added to tree
' 29-Sep-2010 fsr: farfields with different orientations/positions were stored with the same file name when imported from the same file name, fixed
' 24-Sep-2010 fsr: imported farfields now stored under "farfields", nearfields in appropriate subfolder, instead of always "E-field"
' 22-Sep-2010 fsr: added farfield import capability
' 08-Sep-2010 fsr: initial version
' ==============================================================================================================================================================

Option Explicit
Sub Main

	Begin Dialog UserDialog 350,70,"Import 3D field data",.DialogFunc ' %GRID:10,7,1,1
		PushButton 20,42,90,21,"Nearfield",.nearfieldPB
		PushButton 130,42,90,21,"Farfield",.farfieldPB
		Text 10,14,340,14,"Please choose the field type you would like to import:",.Text1
		CancelButton 240,42,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "nearfieldPB"
				If (ImportNearfield = -1) Then ' error
					DialogFunc = True
				End If
			Case "farfieldPB"
				If (ImportFarfield = -1) Then ' error
					DialogFunc = True
				End If
			Case "Cancel"
				DialogFunc = False
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function ImportNearfield() As Integer

	ImportNearfield = -1 ' error

	Dim importedResult As Object
	Dim resultFullName As String, resultName As String
	Dim treePath As String, resultType As String

	treePath = "2D/3D Results\"

	Set importedResult = Result3D("")
	If (Mesh.GetMinimumEdgeLength() = 0) Then
		MsgBox("Please create a mesh first. The mesh needs to be identical with the mesh of the field to be imported.","Mesh check")
		ImportNearfield = -1
		Exit Function
	End If
	resultFullName = GetFilePath("*.m3d","m3d",GetProjectPath("Root"),"Select nearfield to load",0)
	If (resultFullName = "") Then
		ImportNearfield = -1
		Exit Function
	End If
	resultName=Split(resultFullName,"\")(UBound(Split(resultFullName,"\")))
	resultName = Split(resultName,".m3d")(0)
	importedResult.Load(resultFullName)
	If (importedResult.GetNx = Mesh.GetNx _
	                And importedResult.GetNy = Mesh.GetNy _
	                And importedResult.GetNz = Mesh.GetNz _
	                And importedResult.GetLength = Mesh.GetNp) Then
	    importedResult.Save("^imported_"+resultName)
	    resultType = importedResult.GetType()
	    resultType = Replace(resultType, "Powerflow", "Power Flow")
	    treePath = treePath + Mid(resultType,IIf(InStr(StrConv(resultType,vbLowerCase),"dynamic")>0,9,1))+"\"
	    importedResult.AddToTree(treePath+"imported_"+resultName,resultName)
	    ImportNearfield = 0
	    Exit Function
	Else
        MsgBox("File not imported, mesh needs to be identical.")
        ImportNearfield = -1
        Exit Function
	End If

End Function

Function ImportFarfield() As Integer

	Dim resultFullName As String, resultName As String, treeEntry As String, ResultPath As String
	Dim Position(2) As Double
	Dim EulerXAngles(2) As Double
	Dim FFAmplitude As Double
	Dim FFPhase As Double
	Dim fSTEP As Double

	ImportFarfield = -1 ' error

	ResultPath = GetProjectPath("Result")

	Position(0) = 0
	Position(1) = 0
	Position(2) = 0
	EulerXAngles(0) = 0
	EulerXAngles(1) = 0
	EulerXAngles(2) = 0
	FFAmplitude = 1
	FFPhase = 0
	fSTEP = 5

	resultFullName = GetFilePath("*.ffp;*.ffs","Farfield Files|*.ffp;*.ffs|All Files|*.*",GetProjectPath("Root"),"Select farfield to load",0)
	If (resultFullName = "") Then
		ImportFarfield = -1
		Exit Function
	Else
		resultName = "imported_"+Split(resultFullName,"\")(UBound(Split(resultFullName,"\")))
		treeEntry = Split(resultName,".ffp")(0)

	Begin Dialog UserDialog 550,210,"Farfield import settings" ' %GRID:10,7,1,1
		OKButton 350,182,90,21
		CancelButton 450,182,90,21
		Text 20,21,170,14,"Name in Navigation Tree:",.Text1
		TextBox 210,14,330,21,.treeEntryT
		Text 20,49,160,14,"Origin of farfield (x/y/z):",.Text2
		TextBox 210,42,90,21,.xPosT
		TextBox 330,42,90,21,.yPosT
		TextBox 450,42,90,21,.zPosT
		TextBox 210,70,90,21,.alphaT
		TextBox 330,70,90,21,.betaT
		TextBox 450,70,90,21,.gammaT
		Text 310,49,10,14,"/",.Text3
		Text 310,77,10,14,"/",.Text6
		Text 430,77,10,14,"/",.Text7
		Text 430,49,10,14,"/",.Text4
		Text 20,77,180,14,"Orientation in Euler-x angles:",.Text5
		Text 20,105,90,14,"Amplitude:",.Text8
		TextBox 210,98,90,21,.ampT
		Text 20,133,90,14,"Phase:",.Text9
		TextBox 210,126,90,21,.phaseT
		TextBox 210,154,90,21,.stepT
		Text 20,161,170,14,"Angle resolution (deg):",.Text10
	End Dialog

		Dim dlg As UserDialog

		dlg.treeEntryT = treeEntry
		dlg.xPosT = cstr(Position(0))
		dlg.yPosT = cstr(Position(1))
		dlg.zPosT = cstr(Position(2))
		dlg.alphaT = cstr(EulerXAngles(0))
		dlg.betaT = cstr(EulerXAngles(1))
		dlg.gammaT = cstr(EulerXAngles(2))
		dlg.ampT = cstr(FFAmplitude)
		dlg.phaseT = cstr(FFPhase)
		dlg.stepT = cstr(fSTEP)

		If (Dialog(dlg) = 0) Then ' User pressed Cancel
			ImportFarfield = -1
			Exit Function
		End If

		treeEntry = dlg.treeEntryT
		resultName = treeEntry+".ffp"
		Position(0) = Evaluate(dlg.xPosT)
		Position(1) = Evaluate(dlg.yPosT)
		Position(2) = Evaluate(dlg.zPosT)
		EulerXAngles(0) = Evaluate(dlg.alphaT)
		EulerXAngles(1) = Evaluate(dlg.betaT)
		EulerXAngles(2) = Evaluate(dlg.gammaT)
		FFAmplitude = Evaluate(dlg.ampT)
		FFPhase = Evaluate(dlg.phaseT)
		fSTEP = Evaluate(dlg.stepT)

		FarfieldArray.ClearAntennaItems
		FarfieldArray.AddAntennaItem(resultFullName, Position(0), Position(1), Position(2), EulerXAngles(0), EulerXAngles(1), EulerXAngles(2), FFAmplitude, FFPhase)
		FarfieldArray.SetNormalizeAntennas(False) ' simply import, do not renormalize; renormalization problematic in particular for results generated by plane wave
		FarfieldArray.ExecuteCombine(ResultPath + resultName, fSTEP)
		With Resulttree
			.Reset
			.Name "Farfields\"+treeEntry
			.File "^"+resultName
			.Type "Farfield"
			.Add
		End With
	    ImportFarfield = 0
		Exit Function

	End If

End Function
