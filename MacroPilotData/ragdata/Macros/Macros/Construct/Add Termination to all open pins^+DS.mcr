' This macro adds resistances and grounds to all open pins of a specified block, used within CST DS.
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2009-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 19-Feb-2025 iqa: Made applicable on reference pins and does not quit on the first error. (Fix CST-76797)
' 18-Jun-2020 fcu: Replace the link command with the net command
' 25-Apr-2020 ube: Replace spaces by underscore to avoid invalid parameter name
' 22-Jan-2019 iqa: Reduced distance between new blocks and the puin for a compact result
' 22-Jan-2019 iqa: Set Option Explicit and add missing declarations and fixed some code indentation (cl 661202)
' 22-Jan-2019 iqa: use GetTransformedPinLayout instead of GetPinLayout to position the new blocks on the right side of the main block if it was rotated or flipped. (Fix CST-56089: cl 661163)
' 03-Jan-2019 ube : parameter naming now more robust
' 05-Sep-2018 iqa: used the new implemented command CircuitProbe.SetBlockPin to connect R with the block
'                   to fix CST-54893 (Add Termination to all open pins fails in 2019 due to Link.GetName "not yet implemented")
' 08-Jan-2017 ytn: added option to use R/L/C as termination
' 30-Oct-2017 iqa: fix 45082, 49565: fix positioning and orientation of created components for top and bottom block pins
' 30-Oct-2017 iqa: fix if a pin is unconnected and its res or gnd was found (the found component will be overwritten with warning)
' 26-Oct-2017 fsr: Added option to go through all blocks; enabled parameterization for termination impedance
' 15-Oct-2015 gba: small speed up by calling block.setdoubleproperty before block.create
' 31-Oct-2011 gba: reorder block creation for better routing
' 25-Feb-2011 ube: change probe name, if already existing
' 15-Jun-2009 ube: new VBA command used "IsPortConnected"
' 28-Apr-2009 ube: First version
'-----------------------------------------------------------------------------------------------------------------------------

'#include "vba_globals_all.lib"

Option Explicit

Const LumpedNameArray = Array("R", "L", "C")
Dim SLumpedString As String		   ' Lumped element name
Dim SLumpedType As String		   ' Lumped element type
Dim SLumpedUnit As String		   ' Lumped element unit displayed
Dim BlockNameArray() As String     ' Array containing the names of all available blocks
Dim nBlocks As Long

Dim sTerminationParameterName As String

' the factor of distanses pin position <-> Res/Gnd
Private Const distPinRes = 3
Private Const distPinGnd = 5

Private Sub ConnectComponents(Comps() As Variant)
' -------------------------------------------------------------------------------------------------
' ConnectComponents: This function connects the components.
' -------------------------------------------------------------------------------------------------
	Dim numberOfComponents As Integer
	numberOfComponents = UBound(Comps)-LBound(Comps)+1

	Dim subnet() As Variant
	ReDim subnet(numberOfComponents - 1,2)

	Dim Comp As Variant
	Dim i As Integer
	i = 0
	For Each Comp In Comps
		subnet(i,0) = Comp(0)
		subnet(i,1) = Comp(1)

		If (Comp(0) = "B" Or Comp(0) = "BLOCK") Then
			Block.Reset
			Block.Name Comp(1)
			subnet(i,2) = Block.GetPinIndex(Comp(2))
		Else
			subnet(i,2) = CInt(Comp(2))
		End If
		i = i + 1
	Next Comp

	With Net
		.Reset
		.ConnectPins(subnet)
		.Apply
	End With
End Sub

Private Sub FillBlockNameArray()

' -------------------------------------------------------------------------------------------------
' FillBlockNameArray: This function fills the global array of block names
' -------------------------------------------------------------------------------------------------

	nBlocks = Block.StartBlockNameIteration

	ReDim BlockNameArray(nBlocks)

	Dim nIndex As Long
	BlockNameArray(0) = "++ALL BLOCKS++"
	For nIndex=1 To nBlocks
		BlockNameArray(nIndex) = Block.GetNextBlockName
	Next nIndex

End Sub

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
		DlgText("LumpedUnit", Units.GetUnit("Resistance"))
		DlgText("LumpedValue", "50")
		DlgValue("Addprobes", 1)
    Case 2 ' Value changing or button pressed'
     If DlgItem = "OK" Then
			If IsEmpty(DlgText("BlockSelectionDLB")) Then
			MsgBox ("Please select a block.")
				DialogFunc = True 'do not exit the dialog
		End If
    End If
	    Select Case DlgText("LumpedSelectionDLB")
    	Case "R"
    		DlgText("LumpedUnit", Units.GetUnit("Resistance"))
    	Case "L"
			DlgText("LumpedUnit", Units.GetUnit("Inductance"))
		Case "C"
			DlgText("LumpedUnit", Units.GetUnit("Capacitance"))
		End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Sub Main
	FillBlockNameArray

	If (nBlocks=0) Then
		MsgBox ("No blocks defined.", vbCritical, "Add external pins")
		Exit All
	End If

	Dim selectedBlock As String
	Dim nBlockStartIndex As Long, nBlockEndIndex As Long, nBlockIndex As Long
	Dim dImpedanceValue As Double
	Dim bParameterizeImpedance As Boolean, bAddProbes As Boolean

	' show dialog
	Begin Dialog UserDialog 380,154,"Add termination to open pins",.DialogFunc ' %GRID:10,7,1,1
		Text 10,14,90,14,"Select Block:",.Text2
		DropListBox 120,7,250,133,BlockNameArray(),.BlockSelectionDLB
		TextBox 120,35,150,21,.Lumpedvalue
		CheckBox 10,70,260,14,"Parameterize lumped element value",.ImpedanceParameterCB
		CheckBox 10,98,100,14,"Add probes",.Addprobes
		DropListBox 10,35,90,21,LumpedNameArray(),.LumpedSelectionDLB
		Text 280,42,80,21,"Unit",.LumpedUnit

		OKButton 180,119,90,21
		CancelButton 280,119,90,21

	End Dialog

	Dim dlg As UserDialog

	If Dialog(dlg)=0 Then Exit All

	selectedBlock = BlockNameArray(dlg.BlockSelectionDLB)

	
	If (dlg.BlockSelectionDLB = 0) Then
		' Go through all blocks
		nBlockStartIndex = 1
		nBlockEndIndex = nBlocks
	Else
		' only apply to selected block
		nBlockStartIndex = dlg.BlockSelectionDLB
		nBlockEndIndex = nBlockStartIndex
	End If

	Dim i As Long
	Dim offsetX As Long ' X Margin around main block
	Dim offsetY As Long ' Y Margin around main block
	Dim blockPosX As Long
	Dim blockPosY As Long
	Dim pinPosX As Long
	Dim pinPosY As Long
	Dim edge As String    ' The edge (side) of master input pin
	Dim edgeIndex As Long ' The edge index of the master input pin
	Dim nNumPins As Long
	Dim RAngel As Long    ' The rotation angel of the resistor.
	                      ' The rotation angel of the ground is always 90 degree contra clockwise to resistor rotation angel
	Dim pinName As String
	Dim validPinName As String
	Dim sRname As String
	Dim sGndName As String
	Dim dLumpedValue As Double
	Dim LumpedString As String

	bAddProbes = CBool(dlg.Addprobes)
	bParameterizeImpedance = CBool(dlg.ImpedanceParameterCB)
	dLumpedValue = Evaluate(dlg.LumpedValue)

	For nBlockIndex = nBlockStartIndex To nBlockEndIndex

		selectedBlock = BlockNameArray(nBlockIndex)

		With Block
			.Reset
			.Name (selectedBlock)
	        nNumPins  = .GetNumberOfPins
			blockPosX = .GetPositionX
			blockPosY = .GetPositionY
			'ReportInformation "Block position = " + CStr(blockPosX) + ", " + CStr(blockPosY)
		End With

		For i=0 To nNumPins-1
			With Block
				.Reset
				.Name (selectedBlock)
				pinPosX = .GetPinPositionX(i)
				pinPosY = .GetPinPositionY(i)
				pinName = .GetPinName(i)
				validPinName = Replace(pinName, "'", "-ref") ' "'" used for reference pins is forbidden for blocks names (for RES and GND created later) -> Replace
			End With
				If dlg.LumpedSelectionDLB = 0 Then
					LumpedString = "CircuitBasic\Resistor"
					SLumpedString = "R-"
					SLumpedType = "Resistance"
				ElseIf dlg.LumpedSelectionDLB = 1 Then
					LumpedString = "CircuitBasic\Inductor"
					SLumpedString = "L-"
					SLumpedType = "Inductance"
				ElseIf dlg.LumpedSelectionDLB = 2 Then
					LumpedString = "CircuitBasic\Capacitor"
					SLumpedString = "C-"
					SLumpedType = "Capacitance"
				End If

				sTerminationParameterName = "z" + selectedBlock + "_Termination_" + SLumpedType
				sTerminationParameterName = NoForbiddenFilenameCharacters(sTerminationParameterName)
				sTerminationParameterName = Replace (sTerminationParameterName, "-", "_")
				sTerminationParameterName = Replace (sTerminationParameterName, " ", "_")

				If bParameterizeImpedance Then
					DS.StoreParameter(sTerminationParameterName, dLumpedValue)
				End If

			If (Block.IsPinConnected(i)) Then
				' ReportInformation "Ignored pin "+CStr(i)+": nothing done for pins, which are connected already."
			ElseIf (Block.GetBusSize(i)>1) Then
				' ReportInformation "Ignored pin "+CStr(i)+": nothing done for bus pins."
			Else
				' Get the layout of the current pin
				Block.GetTransformedPinLayout(i, edge, edgeIndex)
				'ReportInformation "Pin(" + pinName + ") position = " + CStr(pinPosX) + ", " + CStr(pinPosY) + ", edge = " + edge + ", edgeIndex = " + CStr(edgeIndex)

				offsetX = 100
				offsetY = 100
				If edge = "LEFT" Then
					offsetX = -offsetX
					offsetY = 0
					RAngel = 180
				ElseIf edge = "RIGHT" Then
					offsetY = 0
					RAngel = 0
				ElseIf edge = "TOP" Then
					offsetX = 0
					offsetY = -offsetY
					RAngel = 270
				ElseIf edge = "BOTTOM" Then
					offsetX = 0
					RAngel = 90
				Else
					DS.ReportError "Error in block pin layout"
				End If


				With Block
					.Reset
					.type LumpedString
					sRname = SLumpedString+selectedBlock+"-"+validPinName
					.name sRname
					If .DoesExist Then
						DS.ReportWarning Mid(LumpedString, 14) + " " + sRname + " existed already and will be overwritten."
						.delete
						.type LumpedString
						.name sRname
					End If
					.Position(pinPosX+distPinRes*offsetX, pinPosY+distPinRes*offsetY)
					.Rotate (RAngel)
					If bParameterizeImpedance Then
						.SetDoubleProperty (SLumpedType, sTerminationParameterName )
					Else
						.SetDoubleProperty (SLumpedType, dLumpedValue )
					End If
					.Create
				End With

				With Block
					.Reset
					.type "CircuitBasic\Ground"
					sGndName = "zGND-"+selectedBlock+"-"+validPinName
					.name sGndName
					If .DoesExist Then
						ReportWarning "Ground " + sGndName + " existed already and will be overwritten."
						.delete
						.type "CircuitBasic\Ground"
						.name sGndName
					End If
					.Position(pinPosX+distPinGnd*offsetX, pinPosY+distPinGnd*offsetY)
					.Rotate (RAngel-90)
					.Create
				End With

				' Create Net
				ConnectComponents(Array(Array("B",sRname,"1"), Array("B",selectedBlock,pinName)))
				ConnectComponents(Array(Array("B",sRname,"2"), Array("B",sGndName,"1")))

				Dim ijk As Integer
				ijk=0

				If bAddProbes Then
					With CircuitProbe
						.Reset
						.Name "P-"+pinName
						.SetBlockPin("B", selectedBlock, pinName, True)
						While .DoesExist
							ijk = ijk+1
							.Name "P-"+pinName+"-"+Str(ijk)
						Wend
						.Create
					End With
				End If

			End If

		Next i

	Next nBlockIndex

End Sub
