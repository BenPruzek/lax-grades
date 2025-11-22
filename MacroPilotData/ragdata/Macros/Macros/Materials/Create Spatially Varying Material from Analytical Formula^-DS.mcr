'#Language "WWB-COM"

'#include "vba_globals_all.lib"
'#include "exports.lib"

' This macro allows the user to create a spatially distributed material from an analytical formula
'
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' -----------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 03-Apr-2024 yoo: Added thermal conductivity option
' 06-Mar-2020 ech: Fixed wrong file naming with correct "_" 
' 16-Jan-2019 fsr: Fixed wrong index for spherical coordinates
' 15-Jan-2019 yta: added analytical material function used to the navigation tree
' 21-Nov-2018 fsr: Fixed a potential problem with negative coordinate values
' 30-Oct-2017 fsr: Fixed calculation for spherical coordinates
' 14-Mar-2017 fsr: Initial version
' -----------------------------------------------------------------------------------------------------

Option Explicit

Const sCoordinateSystemVariables = Array(Array("$X", "$Y", "$Z"), Array("$R", "$T", "$Z"), Array("$R", "$T", "$P"))
Const sMaterialDataTypes = Array("eps", "mu", "sigma", "sigmam", "thermalcond")
Const sMaterialDataCalls = Array(".Epsilon", ".Mu", ".Sigma", ".SigmaM", ".ThermalConductivity")

Sub Main

	Dim MaterialDataTypes() As String
	Dim sFileName As Double
	Dim i As Long

	ReDim MaterialDataTypes(UBound(sMaterialDataTypes))
	For i = 0 To UBound(MaterialDataTypes)
		MaterialDataTypes(i) = sMaterialDataTypes(i)
	Next

	Begin Dialog UserDialog 800,301,"Define Material from Analytical Function",.DialogFunc ' %GRID:10,7,1,1
		Text 30,14,100,14,"Material name:",.Text4
		TextBox 180,14,180,21,.MaterialNameT
		Text 400,14,40,14,"Type:",.Text5
		DropListBox 470,14,300,70,MaterialDataTypes(),.MaterialDataTypeDLB
		OptionGroup .CoordinateSystemOG
			OptionButton 180,42,170,14,"Cartesian ($X/$Y/$Z)",.CartesianOB
			OptionButton 360,42,170,14,"Cylindrical ($R/$T/$Z)",.CylindricalOG
			OptionButton 540,42,170,14,"Spherical ($R/$T/$P)",.SphericalOB
		Text 30,42,130,14,"Coordinate system:",.Text1
		Text 30,133,90,14,"Range:",.Text2
		Text 30,77,90,14,"Function:",.Text3
		TextBox 30,98,740,21,.MaterialFunctionT
		Text 430,133,260,14,"Typical material data value (for meshing):",.Text6
		TextBox 690,126,80,21,.TypicalValueT
		Text 30,168,30,14,"X:",.Coordinate1L
		TextBox 70,161,90,21,.XMinT
		TextBox 200,161,90,21,.XMaxT
		TextBox 410,161,90,21,.XSamplesT
		' CheckBox 480,161,110,14,"YZ symmetry",.YZSymmetryCB
		Text 30,203,30,14,"Y:",.Coordinate2L
		TextBox 70,196,90,21,.YMinT
		TextBox 200,196,90,21,.YMaxT
		TextBox 410,196,90,21,.YSamplesT
		' CheckBox 480,196,110,14,"XZ symmetry",.XZSymmetryCB
		Text 30,238,30,14,"Z:",.Coordinate3L
		TextBox 70,231,90,21,.ZMinT
		TextBox 200,231,90,21,.ZMaxT
		TextBox 410,231,90,21,.ZSamplesT
		' CheckBox 480,231,110,14,"XY symmetry",.XYSymmetryCB
		OKButton 480,266,90,21
		PushButton 580,266,90,21,"Apply",.ApplyPB
		CancelButton 680,266,90,21
		Text 20,273,430,14,"",.StatusL
		Text 170,168,20,14,"...",.Text7
		Text 170,203,20,14,"...",.Text8
		Text 170,238,20,14,"...",.Text9
		Text 300,168,100,14,", # of samples:",.Text10
		Text 300,203,100,14,", # of samples:",.Text11
		Text 300,238,100,14,", # of samples:",.Text12

	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg, 0) = 0) Then
		Exit All
	End If


End Sub


Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

	Dim sFilePath As String

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgValue("CoordinateSystemOG", "0")
		DlgText("MaterialFunctionT", "3*($X^2 + $Y^2 + $Z^2)")
		DlgText("TypicalValueT", "1")
		DlgText("XMinT", "-10")
		DlgText("XMaxT", "10")
		DlgText("XSamplesT", "11") ' 3 is the minimum number of points
		DlgText("YMinT", "-10")
		DlgText("YmaxT", "10")
		DlgText("YSamplesT", "5") ' 3 is the minimum number of points
		DlgText("ZMinT", "-5")
		DlgText("ZMaxT", "5")
		DlgText("ZSamplesT", "3") ' 3 is the minimum number of points
		DlgText("MaterialNameT", "Spatially_Varying_Material_Analytical")
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "OK", "ApplyPB"
				If ((EValuate(DlgText("XSamplesT")) < 3) Or (EValuate(DlgText("YSamplesT")) < 3) Or (EValuate(DlgText("ZSamplesT")) < 3)) Then
					MsgBox("Please make sure to have at least 3 samples per dimension.", "Check Settings")
					DialogFunc = True
					Exit Function
				End If
				sFilePath = GetProjectPath("Temp") & DlgText("MaterialNameT") & "_" & DlgText("MaterialDataTypeDLB") & ".txt"
				'If (WriteAnalyticalMaterialDataFile(DlgValue("CoordinateSystemOG"), DlgText("MaterialFunctionT"), _
				'									Evaluate(DlgText("XMinT")), Evaluate(DlgText("XMaxT")), Evaluate(DlgText("XSamplesT")), _
				'									Evaluate(DlgText("YMinT")), Evaluate(DlgText("YMaxT")), Evaluate(DlgText("YSamplesT")), _
				'									Evaluate(DlgText("ZMinT")), Evaluate(DlgText("ZMaxT")), Evaluate(DlgText("ZSamplesT")), _
				'									sFilePath) = -1) Then
				'	' error, keep dialog open
				'	DialogFunc = True
				'	Exit Function
				'End If
				' FSC: alternative format - fixed grid, needed for 2017
				Dim streamNum As Integer, sMaterialName As String
				sMaterialName = DlgText("MaterialNameT")& "_" & DlgText("MaterialDataTypeDLB")&".txt"
				streamNum = FreeFile
				Open GetProjectPath("Model3D")+sMaterialName For Output As #streamNum
				Print #streamNum, "Analytical Material Function: " + DlgText("MaterialFunctionT")
				Close #streamNum
				If (WriteAnalyticalMaterialDataFileFixedGrid(DlgValue("CoordinateSystemOG"), DlgText("MaterialFunctionT"), _
																Evaluate(DlgText("XMinT")), Evaluate(DlgText("XMaxT")), Evaluate(DlgText("XSamplesT")), _
																Evaluate(DlgText("YMinT")), Evaluate(DlgText("YMaxT")), Evaluate(DlgText("YSamplesT")), _
																Evaluate(DlgText("ZMinT")), Evaluate(DlgText("ZMaxT")), Evaluate(DlgText("ZSamplesT")), _
																False, False, False, _
																sFilePath) = -1) Then
					' error, keep dialog open
					DialogFunc = True
					Exit Function
				End If

				' Convert from txt to md3
				Material.ConvertMaterialField(sFilePath, DlgText("MaterialNameT") & "_" & DlgText("MaterialDataTypeDLB"))
				' Create actual material
				CreateAnalyticalMaterial(DlgText("MaterialNameT"), DlgText("MaterialDataTypeDLB"), DlgText("TypicalValueT"))
				If (DlgItem = "ApplyPB") Then DialogFunc = True
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function WriteAnalyticalMaterialDataFile(nCoordinateSystem As Integer, sExpression As String, _
											dXMin As Double, dXMax As Double, nXSamples As Long, _
											dYMin As Double, dYMax As Double, nYSamples As Long, _
											dZMin As Double, dZMax As Double, nZSamples As Long, _
											sFileName As String) As Integer

	Dim i As Long, j As Long, k As Long
	Dim nFileID As Integer
	Dim sVar1 As String, sVar2 As String, sVar3 As String ' names, e.g. R/T/P
	Dim dVar1Val As Double, dVar2Val As Double, dVar3Val As Double ' values
	Dim dXValues() As Double, dYValues() As Double, dZValues() As Double
	Dim dDeltaX As Double, dDeltaY As Double, dDeltaZ As Double

	ReDim dXValues(nXSamples - 1)
	ReDim dYValues(nYSamples - 1)
	ReDim dZValues(nZSamples - 1)

	' Determine variable name by selected coordinate system type
	sVar1 = sCoordinateSystemVariables(nCoordinateSystem)(0)
	sVar2 = sCoordinateSystemVariables(nCoordinateSystem)(1)
	sVar3 = sCoordinateSystemVariables(nCoordinateSystem)(2)

	nFileID = OpenBufferedFile_LIB(sFileName, "Output")
	BufferedFileWriteLine_LIB(nFileID, "x	y	z	Value")
	dDeltaX = (dXMax - dXMin) / (nXSamples - 1)
	For i = 0 To nXSamples - 1
		dXValues(i) = dXMin + i * dDeltaX
	Next
	dDeltaY = (dYMax - dYMin) / (nYSamples - 1)
	For i = 0 To nYSamples - 1
		dYValues(i) = dYMin + i * dDeltaY
	Next
	dDeltaZ = (dZMax - dZMin) / (nZSamples - 1)
	For i = 0 To nZSamples - 1
		dZValues(i) = dZMin + i * dDeltaZ
	Next
	On Error GoTo ExitWithError
		DlgText("StatusL", "Creating data... 0%")
		For k = 0 To nZSamples - 1
			For j = 0 To nYSamples - 1
				For i = 0 To nXSamples - 1
					' Determine values of the variables depending upon selected coordinate system
					Select Case nCoordinateSystem
						Case 0 ' Cartesian, no change
							dVar1Val = dXValues(i)
							dVar2Val = dYValues(j)
							dVar3Val = dZValues(k)
						Case 1 ' Cylindric
							dVar1Val = Sqr(dXValues(i)^2 + dYValues(j)^2) ' R
							dVar2Val = Atn2(dYValues(j), dXValues(i)) ' Theta in rad
							dVar3Val = dZValues(k) ' z
						Case 2 ' Spherical
							dVar1Val = Sqr(dXValues(i)^2 + dYValues(j)^2 + dZValues(k) ^2) ' R
							dVar2Val = Atn2(dYValues(j), dXValues(i)) ' Theta in rad
							dVar3Val = Atn2(Sqr(dXValues(i)^2 + dYValues(j)^2), dZValues(k)) ' Phi in rad
					End Select
					BufferedFileWriteLine_LIB(nFileID, USFormat(dXValues(i), "0.00000E+00") & vbTab & USFormat(dYValues(j), "0.00000E+00") & vbTab & USFormat(dZValues(k), "0.00000E+00") & vbTab & USFormat(Evaluate(Replace(Replace(Replace(sExpression, sVar1, "(" & CStr(dVar1Val) & ")"), sVar2, "(" & CStr(dVar2Val) & ")"), sVar3, "(" & CStr(dVar3Val) & ")")), "0.00000E+00") & " ")
				Next
				DlgText("StatusL", "Creating data... " & CStr((k*(nYSamples)*(nXSamples)+j*(nXSamples))/((nXSamples)*(nYSamples)*(nZSamples)))*100 & "%")
			Next
		Next
		DlgText("StatusL", "Creating data... 100%")
	On Error GoTo 0

	CloseBufferedFile_LIB(nFileID)
	WriteAnalyticalMaterialDataFile = 0 ' all went well
	DlgText("StatusL", "")
	Exit Function

	ExitWithError:
		CloseBufferedFile_LIB(nFileID)
		WriteAnalyticalMaterialDataFile = -1 ' an error occured
		DlgText("StatusL", "")
		MsgBox("Error creating material data. Please check your settings. Typical reasons for this problem include syntax errors in the expression or division by zero.", "Error")
		Exit Function

End Function


Function CreateAnalyticalMaterial(sMaterialName As String, sMaterialDataType As String, sTypicalValue As String) As Integer

	Dim sHistoryCommand As String

	sHistoryCommand = ""
	AppendHistoryLine_LIB(sHistoryCommand, "With Material")
	AppendHistoryLine_LIB(sHistoryCommand, ".Reset")
	AppendHistoryLine_LIB(sHistoryCommand, ".Name", sMaterialName+"_"+sMaterialDataType)
	AppendHistoryLine_LIB(sHistoryCommand, ".Type", "Normal")
	AppendHistoryLine_LIB(sHistoryCommand, sMaterialDataCalls(FindListIndex(sMaterialDataTypes, sMaterialDataType)), sTypicalValue)
	AppendHistoryLine_LIB(sHistoryCommand, ".ResetSpaceMapBasedMaterial", sMaterialDataType)
	AppendHistoryLine_LIB(sHistoryCommand, ".SpaceMapBasedOperator", sMaterialDataType, "3DImport")
	AppendHistoryLine_LIB(sHistoryCommand, ".AddSpaceMapBasedMaterialStringParameter", sMaterialDataType, "map_filename", sMaterialName & "_" & sMaterialDataType & ".m3d")
	AppendHistoryLine_LIB(sHistoryCommand, ".Create")
	AppendHistoryLine_LIB(sHistoryCommand, "End With")
	AddToHistory("define material: " & sMaterialName, sHistoryCommand)

End Function

Function WriteAnalyticalMaterialDataFileFixedGrid(nCoordinateSystem As Integer, sExpression As String, _
													dXMin As Double, dXMax As Double, nXSamples As Long, _
													dYMin As Double, dYMax As Double, nYSamples As Long, _
													dZMin As Double, dZMax As Double, nZSamples As Long, _
													bXSymmetry As Boolean, bYSymmetry As Boolean, bZSymmetry As Double, _
													sFileName As String) As Integer

	' Same as WriteAnalyticalMaterialDataFileFixed, but uses a different intermediate format "FixedGrid"
	' Currently not used in this macro 3/14/2017 FSC

	Dim i As Long, j As Long, k As Long
	Dim nFileID As Integer
	Dim sVar1 As String, sVar2 As String, sVar3 As String ' names, e.g. R/T/P
	Dim dVar1Val As Double, dVar2Val As Double, dVar3Val As Double ' values
	Dim dXValues() As Double, dYValues() As Double, dZValues() As Double
	Dim dDeltaX As Double, dDeltaY As Double, dDeltaZ As Double

	ReDim dXValues(nXSamples - 1)
	ReDim dYValues(nYSamples - 1)
	ReDim dZValues(nZSamples - 1)

	' Determine variable name by selected coordinate system type
	sVar1 = sCoordinateSystemVariables(nCoordinateSystem)(0)
	sVar2 = sCoordinateSystemVariables(nCoordinateSystem)(1)
	sVar3 = sCoordinateSystemVariables(nCoordinateSystem)(2)

	nFileID = OpenBufferedFile_LIB(sFileName, "Output")
	BufferedFileWriteLine_LIB(nFileID, "# CST material field file")
	BufferedFileWriteLine_LIB(nFileID, "# Format: FixedGrid")
	BufferedFileWriteLine_LIB(nFileID, "# Version: 20150107")
	BufferedFileWriteLine_LIB(nFileID, "# LengthUnit: " & Units.GetUnit("Length"))
	BufferedFileWriteLine_LIB(nFileID, "# SamplePoint x:")
	dDeltaX = (dXMax - dXMin) / (nXSamples - 1)
	For i = 0 To nXSamples - 1
		dXValues(i) = dXMin + i * dDeltaX
		BufferedFileWrite_LIB(nFileID, CStr(dXValues(i)) & " ")
	Next
	BufferedFileWrite_LIB(nFileID, vbNewLine) ' complete line
	BufferedFileWriteLine_LIB(nFileID, "# SamplePoint y:")
	dDeltaY = (dYMax - dYMin) / (nYSamples - 1)
	For i = 0 To nYSamples - 1
		dYValues(i) = dYMin + i * dDeltaY
		BufferedFileWrite_LIB(nFileID, CStr(dYValues(i)) & " ")
	Next
	BufferedFileWrite_LIB(nFileID, vbNewLine) ' complete line
	BufferedFileWriteLine_LIB(nFileID, "# SamplePoint z:")
	dDeltaZ = (dZMax - dZMin) / (nZSamples - 1)
	For i = 0 To nZSamples - 1
		dZValues(i) = dZMin + i * dDeltaZ
		BufferedFileWrite_LIB(nFileID, CStr(dZValues(i)) & " ")
	Next
	BufferedFileWrite_LIB(nFileID, vbNewLine) ' complete line
	BufferedFileWriteLine_LIB(nFileID, "# Symmetry x: " & IIf(bXSymmetry, "True", "False"))
	BufferedFileWriteLine_LIB(nFileID, "# Symmetry y: " & IIf(bYSymmetry, "True", "False"))
	BufferedFileWriteLine_LIB(nFileID, "# Symmetry z: " & IIf(bZSymmetry, "True", "False"))
	BufferedFileWriteLine_LIB(nFileID, "# Data section:")

	On Error GoTo ExitWithError
		DlgText("StatusL", "Creating data... 0%")
		For k = 0 To nZSamples - 1
			For j = 0 To nYSamples - 1
				For i = 0 To nXSamples - 1
					' Determine values of the variables depending upon selected coordinate system
					Select Case nCoordinateSystem
						Case 0 ' Cartesian, no change
							dVar1Val = dXValues(i)
							dVar2Val = dYValues(j)
							dVar3Val = dZValues(k)
						Case 1 ' Cylindric
							dVar1Val = Sqr(dXValues(i)^2 + dYValues(j)^2) ' R
							dVar2Val = Atn2(dYValues(j), dXValues(i)) ' Theta in rad
							dVar3Val = dZValues(k) ' z
						Case 2 ' Spherical
							dVar1Val = Sqr(dXValues(i)^2 + dYValues(j)^2 + dZValues(k) ^2) ' R
							dVar2Val = Atn2(dYValues(j), dXValues(i)) ' Theta in rad
							dVar3Val = Atn2(Sqr(dXValues(i)^2 + dYValues(j)^2), dZValues(k)) ' Phi in rad
					End Select
					BufferedFileWrite_LIB(nFileID, USFormat(Evaluate(Replace(Replace(Replace(sExpression, sVar1, "(" & CStr(dVar1Val) & ")"), sVar2, "(" & CStr(dVar2Val) & ")"), sVar3, "(" & CStr(dVar3Val) & ")")), "0.00000E+00") & " ")
				Next
				BufferedFileWrite_LIB(nFileID, vbNewLine) ' complete line
				DlgText("StatusL", "Creating data... " & CStr((k*(nYSamples)*(nXSamples)+j*(nXSamples))/((nXSamples)*(nYSamples)*(nZSamples)))*100 & "%")
			Next
		Next
		DlgText("StatusL", "Creating data... 100%")
	On Error GoTo 0

	CloseBufferedFile_LIB(nFileID)
	WriteAnalyticalMaterialDataFileFixedGrid = 0 ' all went well
	DlgText("StatusL", "")
	Exit Function

	ExitWithError:
		CloseBufferedFile_LIB(nFileID)
		WriteAnalyticalMaterialDataFileFixedGrid = -1 ' an error occured
		DlgText("StatusL", "")
		MsgBox("Error creating material data. Please check your settings. Typical reasons for this problem include syntax errors in the expression or division by zero.", "Error")
		Exit Function

End Function
