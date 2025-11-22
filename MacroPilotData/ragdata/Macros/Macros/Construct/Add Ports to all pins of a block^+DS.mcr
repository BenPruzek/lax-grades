' This macro adds ports to all pins of a specified block, used within CST DS.
' A typical case is a MWS-model with many ports, which should be studied in the canvas for eg. transient or frequency studies.
'
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 19-Feb-2025 iqa: Made applicable on reference pins and does not quit on the first error. (Fix CST-76797)
' 18-Jun-2020 fcu: Replace the link command with the net command
' 22-Jan-2019 iqa: use GetTransformedPinLayout instead of GetPinLayout to position the new blocks on the right side of the main block if it was rotated or flipped. (Fix CST-56089: cl 661163)
' 30-Oct-2017 iqa: fix 45082, 49565: fix positioning and orientation of created components for top and bottom block pins
' 03-Feb-2012 ube: ignore those pins of the blockp, which are already connected to something else
' 15-Jul-2010 gba: ignore blocks without pins (e.g. reference blocks)
' 05-Oct-2009 gba: avoid duplicate port names and fix macro for blocks with non-numeric names, e.g. SPICE
' 07-Dec-2007 ube: included in 2008.02
' 22-Nov-2007 mel: First version
'-----------------------------------------------------------------------------------------------------------------------------

Dim BlockNameArray() As String     ' Array containing the names of all available blocks with at least one pin
Dim nBlocks As Long

Private Sub ConnectComponents(extPortName As String, blockName As String, blockPinName As String)
' -------------------------------------------------------------------------------------------------
' ConnectComponents: This function connects the external port with the block's pin.
' -------------------------------------------------------------------------------------------------
	ReportInformation "Connecting external port " + extPortName + " with block-pin " + blockName + "/" + blockPinName

	Dim subnet() As Variant
	ReDim subnet(1,2)

	subnet(0,0) = "P"
	subnet(0,1) = extPortName
	subnet(0,2) = 0

	subnet(1,0) = "B"
	subnet(1,1) = blockName
	With Block
		.Reset
		.Name blockName
		subnet(1,2) = .GetPinIndex(blockPinName)
	End With

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

	Dim nRawBlocks As Long
	Dim RawBlockNames() As String
	Dim nPins As Integer

	nRawBlocks = Block.StartBlockNameIteration
	ReDim RawBlockNames(nRawBlocks)
	ReDim BlockNameArray(nRawBlocks)

	For nIndex=0 To nRawBlocks-1
		RawBlockNames(nIndex) = Block.GetNextBlockName
	Next nIndex

	nBlocks = 0
	Dim nIndex2 As Integer
	For nIndex=0 To nRawBlocks-1
		With Block
			.Reset
			.Name RawBlockNames(nIndex)
			nPins = .GetNumberOfPins
		End With
		If nPins > 0 Then
			BlockNameArray(nIndex2) = RawBlockNames(nIndex)
			nIndex2 = nIndex2 + 1
			nBlocks = nBlocks + 1
		End If
	Next nIndex

	ReDim Preserve BlockNameArray(nBlocks)

End Sub

Private Function GetValidExtPortName(pinName As String) As String
	' Returns a valid external port name relying on the given pin name
	' 1. Prefer pin name if valid
	' 2. Replace possible forbidden charachters to '_'
	' 3. Search the schematic for same named external port and avoid them to ensure unique port numbers
	Dim nExtPorts As Long
	Dim validName As String
	Dim nameChanged As Boolean

	' '~' is a known used forbidden char in block pin names -> replace it here with '_'
	' If other forbidden chars make problem -> replace them either: ^\"\':/*?<>|°~´`
	validName = Replace(pinName, "~", "_")
	validName = Replace(pinName, "'", "-ref")
	nameChanged  = True
	With ExternalPort
		nExtPorts = .StartPortNameIteration
		While nameChanged
			nameChanged = False
			For nIndex=0 To nExtPorts-1
				If validName = .GetNextPortName Then
					validName = validName+"_"
					nameChanged = True
				End If
			Next nIndex
		Wend
	End With
	GetValidExtPortName = validName
End Function

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed'
     If DlgItem = "OK" Then
        If IsEmpty(DlgText(ComboBox1)) Then
			MsgBox ("Please select a block.")
			DialogFunc=True
		End If
    End If
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Sub Main
	FillBlockNameArray

	If (nBlocks=0) Then
		MsgBox ("No blocks defined.", vbCritical, "Add external ports")
		Exit All
	End If

	Dim selectedBlock As String

	If (nBlocks > 1) Then
		' show dialog
		Begin Dialog UserDialog 380,80,"Add external ports",.DialogFunc ' %GRID:10,7,1,1
			Text 10,7,40,14,"Select Block:",.Text2
			OKButton 90,42,90,21
			CancelButton 190,42,90,21
			DropListBox 120,7,250,192,BlockNameArray(),.DropListBox1
		End Dialog

		Dim dlg As UserDialog

		If Dialog(dlg)=0 Then Exit All

		selectedBlock = BlockNameArray(dlg.DropListBox1)

	Else
		selectedBlock = BlockNameArray(0)
	End If


	Dim nPins As Integer
	Dim posX As Long
	Dim posY As Long
	Dim edge As String    ' The edge (side) of master input pin
	Dim edgeIndex As Long ' The edge index of the master input pin
	Dim offsetX As Long   ' X offset between the block pins and the external ports
	Dim offsetY As Long   ' Y offset between the block pins and the external ports
	Dim extPortName As String
	Dim busSize As Integer

	With Block
		.Reset
		.Name (selectedBlock)
		nPins = .GetNumberOfPins
	End With

	For i=0 To nPins-1
		posX = Block.GetPinPositionX(i)
		posY = Block.GetPinPositionY(i)
		pinName = Block.GetPinName(i)
		offsetX = 100
		offsetY = 100
		If (Block.IsPinConnected(i)) Then
			' nothing done for pins, which are connected already
		Else
			busSize = Block.GetBusSize(i)

			' Get the layout of the current pin
			Block.GetTransformedPinLayout(i, edge, edgeIndex)
			'ReportInformation "Pin(" + pinName + ") position = " + CStr(pinPosX) + ", " + CStr(pinPosY) + ", edge = " + edge + ", edgeIndex = " + CStr(edgeIndex)

			If edge = "LEFT" Then
				offsetX = offsetX*-1
				offsetY = 0
			ElseIf edge = "RIGHT" Then
				offsetX = offsetX
				offsetY = 0
			ElseIf edge = "TOP" Then
				offsetX = 0
				offsetY = offsetY*-1
			ElseIf edge = "BOTTOM" Then
				offsetX = 0
				offsetY = offsetY
			Else
				ReportError "Error in block pin layout"
			End If

			With ExternalPort
				.Reset
				' Create external port
				extPortName = GetValidExtPortName(pinName)
				.Name extPortName
				.Position(posX+offsetX, posY+offsetY)
				.Create ' external port
				.SetNumberOfPorts busSize
			End With

			' Create net between external port and block pin
			ConnectComponents(extPortName, selectedBlock, pinName)
			
		End If

	Next i

End Sub
