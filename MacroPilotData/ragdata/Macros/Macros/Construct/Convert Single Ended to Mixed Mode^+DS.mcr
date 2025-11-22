' This macro adds ports to all pins of a specified block, used within CST DS.
' A typical case is a MWS-model with many ports, which should be studied in the canvas for eg. transient or frequency studies.
'
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 01-Sep-2021 iqa: Replaced deprecated 'ExternalPort.Number'-commands with the 'ExternalPort.Name'-command
' 12-Feb-2021 fcu: Connect pins and not ports, this macro is supposed to work on pins.
' 22-Jun-2020 fcu: Replace the link command with the net command
' 22-Jan-2019 iqa: use GetTransformedPinLayout instead of GetPinLayout to position the new blocks on the right side of the main block if it was rotated or flipped. (Fix CST-56089: cl 661163)
' 14-Nov-2017 iqa: add a dialog to pervent user from applying the macro on a block with bus pins.
'                   Note: The macro still works as before for blocks without bus pins.
' 30-Oct-2017 iqa: fix 45082, 49565: fix positioning and orientation of created components for top and bottom block pins
'                   Note: only case 0 (Single Ended) was fixed by positioning the external ports at the correct side.
' 26-Oct-2017 iqa: fix 49563: macro now can handle bus pins: 
'                   1. use block pin functions instead of block port functions (e.g. GetNumberOfPins instead of GetNumberOfPorts).
'                   2. skip bus pins from calculation and connection
' 19-May-2014 ctc: re-arrange photo placement - PR 33598
' 25-Apr-2014 ctc: include option to add transmission line delay - for TDR from S-parameter calculation
' 20-Aug-2013 ctc: add in mixed mode converter option
' 03-Feb-2012 ube: ignore those pins of the blockp, which are already connected to something else
' 15-Jul-2010 gba: ignore blocks without pins (e.g. reference blocks)
' 05-Oct-2009 gba: avoid duplicate port names and fix macro for blocks with non-numeric names, e.g. SPICE
' 07-Dec-2007 ube: included in 2008.02
' 22-Nov-2007 mel: First version
'-----------------------------------------------------------------------------------------------------------------------------
Type tPin
	BusSize As Integer
	Name As String
End Type

Dim BlockNameArray() As String     ' Array containing the names of all available blocks with at least one pin
Dim PinsArray() As tPin            ' Array containing the selected block's pins information
Dim nBlocks As Long
Dim nPins As Integer
Dim nBusPins As Integer
Dim edge As String    ' The edge (side) of a block's pin
Dim edgeIndex As Long ' The edge index of a block's pin
Dim posX As Long
Dim posY As Long
Dim blockPosX As Long
Dim blockPosY As Long
Dim Xoffset As Long
Dim DiffImp As Double
Dim CommImp As Double
Dim nR As Integer
Dim nmmPort As Integer
Dim i As Integer
Dim portNumber As Integer
Dim DiffPort As Boolean
Dim CommPort As Boolean
Dim NameJump As Integer
Dim PortJump As Boolean
Dim txdelay As Boolean
Dim txlength As Double
Dim TLshiftX As Double
Dim TLshiftY As Double
Dim TLname As Integer

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


Public Sub CreateMixedModeConverter()
					With Block
						.Reset
						.Type ("ModeConverter")
						.Name ("MOD" + Str(i+1))
						.Position(blockPosX-posX-TLshiftX, blockPosY-posY)
						'flip MM Conv if located at right hand side
						If i > (nPins/4-1) Then
							.FlipHorizontal
						End If

						.Create
					End With
End Sub

Public Sub CreateTLBlock()
					With Block
						.Reset
						.Type ("Transmissionline")
						.Name ("TL" + Str(i+1)+ "-" + Str(TLname))
						.Position(blockPosX-posX,blockPosY-posY-TLshiftY)
						.SetDoubleProperty ("Length", "20")
						If i > (nPins/4-1) Then
							.FlipHorizontal
						End If
						.Create
					End With
End Sub


Public Sub CommonModeLoad()
					If DiffPort = False Then
						nR = 1
					End If
				'Creating Resistor
					With Block
						.Reset
						.Type ("CircuitBasic\Resistor")
						.Name ("Res"+CStr(i+nR*nPins/2+1))
						.SetDoubleProperty ("Resistance", CommImp)
						If i > (nPins/4-1) Then
							.position(blockPosX-posX+250-TLshiftX, blockPosY-posY+50)
						Else
							.Position(blockPosX-posX-250-TLshiftX, blockPosY-posY+50)
						End If
						.Create
					End With

				'Creating Ground
					With Block
						.Reset
						.Type ("Ground")
						.Name ("GND"+CStr(i+nR*nPins/2+1))
						If i > (nPins/4-1) Then
							.position(blockPosX-posX+375-TLshiftX, blockPosY-posY+70)
						Else
							.Position(blockPosX-posX-375-TLshiftX, blockPosY-posY+70)
						End If
						.Create
					End With

				'Grounding Resistor
				If i > (nPins/4-1) Then
					ConnectComponents(Array(Array("B", "GND"+CStr(i+nR*nPins/2+1), "1"), Array("B", "Res"+CStr(i+nR*nPins/2+1), "2")))
				Else
					ConnectComponents(Array(Array("B", "GND"+CStr(i+nR*nPins/2+1), "1"), Array("B", "Res"+CStr(i+nR*nPins/2+1), "1")))
				End If

				'Linking Resistor to MM Conv
				If i > (nPins/4-1) Then
					ConnectComponents(Array(Array("B", "Res"+CStr(i+nR*nPins/2+1), "1"), Array("B", "MOD" + Str(i+1), "C")))
				Else
					ConnectComponents(Array(Array("B", "Res"+CStr(i+nR*nPins/2+1), "2"), Array("B", "MOD" + Str(i+1), "C")))				
				End If


End Sub

Public Sub CommonModePort()
	Dim aa As Integer
						If DiffPort = True Then
							nmmPort = 1
						End If
							' Create external port
							' Prefer port name if possible
							If IsValidPortNumber(pinName) Then
							  portNumber = pinName
							Else
							  portNumber = 1
							  aa = 1
							End If
							With ExternalPort
								.Reset
								.Name CStr(aa)
								' Ensure unique port numbers
								While .DoesExist()
									portNumber = portNumber + 1
									If PortJump = True Then
										aa = portNumber + nmmPort*nPins/2 + i - 1 + NameJump
									Else
										aa = portNumber + nmmPort*nPins/2 - 1
									End If

									.Reset
									.Name CStr(aa)
								Wend
							If i > (nPins/4-1) Then
								.Position(blockPosX-posX+250-TLshiftX, blockPosY-posY+50)
							Else
								.Position(blockPosX-posX-250-TLshiftX, blockPosY-posY+50)
							End If
								.Create
								.SetFixedImpedance True
								.SetImpedance CommImp
							End With

							' Create net
							ConnectComponents(Array(Array("P",CStr(aa), "0"), Array("B", "MOD" + Str(i+1), "C")))
End Sub

Public Sub DifferentialPort()
							' Create external port
							' Prefer port name if possible
							If IsValidPortNumber(pinName) Then
							  portNumber = pinName
							Else
							  portNumber = 1
							End If
							With ExternalPort
								.Reset
								.Name CStr(portNumber)
								' Ensure unique port numbers
								While .DoesExist()
									If PortJump = True Then
										portNumber = portNumber + i + NameJump
									Else
										portNumber = portNumber + 1
									End If

									.Reset
									.Name CStr(portNumber)
								Wend
							If i > (nPins/4-1) Then
								.Position(blockPosX-posX+250-TLshiftX, blockPosY-posY-30)
							Else
								.Position(blockPosX-posX-250-TLshiftX, blockPosY-posY-30)
							End If
								.Create
								.SetFixedImpedance True
								.SetImpedance DiffImp
							End With

							' Create net
							ConnectComponents(Array(Array("P", CStr(portNumber), "0"), Array("B", "MOD" + Str(i+1), "D")))			
End Sub

Public Sub DifferentialLoad()
				'Creating Resistor
					With Block
						.Reset
						.Type ("CircuitBasic\Resistor")
						.Name ("Res"+CStr(i+1))
						.SetDoubleProperty ("Resistance", DiffImp)
						If i > (nPins/4-1) Then
							.position(blockPosX-posX+250-TLshiftX, blockPosY-posY-30)
						Else
							.Position(blockPosX-posX-250-TLshiftX, blockPosY-posY-30)
						End If
						.Create
					End With

				'Creating Ground
					With Block
						.Reset
						.Type ("Ground")
						.Name ("GND"+CStr(i+1))
						If i > (nPins/4-1) Then
							.position(blockPosX-posX+375-TLshiftX, blockPosY-posY+10)
						Else
							.Position(blockPosX-posX-375-TLshiftX, blockPosY-posY+10)
						End If
						.Create
					End With

				'Grounding Resistor
				If i > (nPins/4-1) Then
					ConnectComponents(Array(Array("B", "GND"+CStr(i+1), "1"), Array("B", "Res"+CStr(i+1), "2")))			
				Else
					ConnectComponents(Array(Array("B", "GND"+CStr(i+1), "1"), Array("B", "Res"+CStr(i+1), "1")))							
				End If

				'Linking Resistor to MM Conv

				If i > (nPins/4-1) Then
					ConnectComponents(Array(Array("B", "Res"+CStr(i+1), "1"), Array("B", "MOD" + Str(i+1), "D" )))
				Else
					ConnectComponents(Array(Array("B", "Res"+CStr(i+1), "2"), Array("B", "MOD" + Str(i+1), "D" )))				
				End If
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

Private Function IsValidPortNumber(pinName As String) As Boolean
	If IsNumeric(pinName) Then
		If Int(Val(pinName)) = Val(pinName) Then
			If Val(pinName) >= 1 Then
				IsValidPortNumber = True
			End If
		End If
	End If
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
		MsgBox ("No blocks defined.", vbCritical, "Convert Single Ended to Mixed Mode")
		Exit All
	End If

	Dim selectedBlock As String

	If (nBlocks > 1) Then
		' show dialog
		Begin Dialog UserDialog 380,80,"Convert Single Ended to Mixed Mode",.DialogFunc ' %GRID:10,7,1,1
			Text 10,10,96,17,"Select Block:",.Text2
			OKButton 90,42,90,21
			CancelButton 190,42,90,21
			DropListBox 100,7,250,184,BlockNameArray(),.DropListBox1
		End Dialog

		Dim dlg As UserDialog

		If Dialog(dlg)=0 Then Exit All

		selectedBlock = BlockNameArray(dlg.DropListBox1)

	Else
		selectedBlock = BlockNameArray(0)
	End If

	nBusPins = 0

	With Block
		.Reset
		.Name (selectedBlock)
		nPins = .GetNumberOfPins
		blockPosX = .GetPositionX
		blockPosY = .GetPositionY
		ReDim PinsArray(nPins)
		For i=0 To nPins-1
			PinsArray(i).BusSize = .GetBusSize(i)
			If PinsArray(i).BusSize > 1 Then
				nBusPins = nBusPins + 1
			End If
			PinsArray(i).Name = .GetPinName(i)
		Next i

	End With

	If nBusPins > 0 Then
		MsgBox ("Script can not be applied on a block with bus pins. Please select another block or resolve buses.")
		Exit All
	End If

'-------------------------------------------------------------------------------
	Begin Dialog UserDialog 1190,399,"Convert Single Ended to Mixed Mode: Mixed Converter Configuration" ' %GRID:10,7,1,1
		Text 70,84,230,14,"(Single Ended Port Configuration)",.Text2
		Text 440,84,240,14,"(Single Ended Port Configuration)",.Text3
		Text 810,84,230,14,"(Single Ended Port Configuration)",.Text4
		CheckBox 60,252,190,14,"Differential Port",.CheckBox1
		CheckBox 420,252,160,14,"Common Mode Port",.CheckBox2
		TextBox 80,273,90,21,.TextBox1
		Text 180,273,200,21,"Ohm (Port or Load Impedance)",.Text5
		TextBox 440,273,90,21,.TextBox2
		Text 550,273,200,14,"Ohm (Port or Load Impedance)",.Text6
		OKButton 60,371,90,21
		CancelButton 160,371,90,21
		OptionGroup .Group1
			OptionButton 50,21,210,14,"Single Ended",.OptionButton1
			OptionButton 50,63,350,14,"Port 1 --> 3, Port 2 --> 4, Port 5 --> 7, Port 6 --> 8, ...",.OptionButton2
			OptionButton 420,63,350,14,"Port 1 --> 2, Port 3 --> 4, Port 5 --> 6, Port 7 --> 8, ...",.OptionButton3
			OptionButton 790,63,350,14,"Port 1 --> 7, Port 2 --> 8, Port 3 --> 9, Port 4 --> 10, ...",.OptionButton4
		Text 70,42,180,16,"Mixed Mode Configuration",.Text1
		Picture 70,105,300,112,GetInstallPath + "\Library\Macros\Construct\Option1.bmp",0,.Picture1
		Picture 440,105,300,105,GetInstallPath + "\Library\Macros\Construct\Option2.bmp",0,.Picture2
		Picture 810,105,300,105,GetInstallPath + "\Library\Macros\Construct\Option3.bmp",0,.Picture3
		Text 20,224,380,21,"*p/s: Insertion Loss == S2,1 after Mixed Mode Conversion",.Text7,2
		Text 400,224,400,14,"*p/s: Insertion Loss == S2,1 after Mixed Mode Conversion",.Text8,2
		Text 800,224,390,21,"*p/s: Insertion Loss == S2,1 after Mixed Mode Conversion",.Text9,2
		CheckBox 60,315,390,14,"Include Tx Line (adding Delay for TDR from S-parameter)",.txdelay
		TextBox 80,336,90,21,.TextBox3
		Text 180,336,90,14,"mm",.Text10
	End Dialog
	Dim dlg1 As UserDialog

	dlg1.Group1 = 0
	dlg1.CheckBox1 = True
	dlg1.txdelay = False
	dlg1.TextBox1 = "100"
	dlg1.TextBox2 = "25"
	dlg1.TextBox3 = "20"

	If Dialog(dlg1)=0 Then Exit All

	Dim ConfigType As Integer



	posX = 400
	posY = (nPins/4-1)*150
	DiffPort = dlg1.Checkbox1
	CommPort = dlg1.Checkbox2
	DiffImp = CDbl(dlg1.TextBox1)
	CommImp = CDbl(dlg1.TextBox2)
	txdelay = dlg1.txdelay
	txlength = cdbl(dlg1.Textbox3)

	ConfigType = dlg1.group1

	If txdelay = True Then
		TLshiftX = 300
		TLshiftY = 75
	Else
		TLshift = 0
	End If


	Select Case ConfigType

		Case 0
				For i=0 To nPins-1
					pinName = PinsArray(i).Name
					If PinsArray(i).BusSize > 1 Then
						ReportInformation "Bus pin " + pinName + " was skipped."
					Else
						' Get the layout of the current pin
						Block.GetTransformedPinLayout(i, edge, edgeIndex)
						posX = Block.GetPinPositionX(i)
						posY = Block.GetPinPositionY(i)
						'ReportInformation "Pin(" + pinName + ") position = " + CStr(pinPosX) + ", " + CStr(pinPosY) + ", edge = " + edge + ", edgeIndex = " + CStr(edgeIndex)
						offsetX = 100
						offsetY = 100
						If (Block.IsPinConnected(i)) Then
							' nothing done for pins, which are connected already
						Else
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

							' Create external port
							' Prefer port name if possible
							If IsValidPortNumber(pinName) Then
							  portNumber = pinName
							Else
							  portNumber = 1
							End If
							With ExternalPort
								.Reset
								.Name CStr(portNumber)
								' Ensure unique port numbers
								While .DoesExist()
									portNumber = portNumber + 1
									.Reset
									.Name CStr(portNumber)
								Wend
								.Position(posX+offsetX, posY+offsetY)
								.Create
							End With

							' Create net
							ConnectComponents(Array(Array("P",CStr(portNumber), "0"), Array("B",selectedBlock, pinName)))
						End If ' .IsPinConnected(i)
					End If ' BusSize(i) > 1
				Next i


		Case 1
			If nPins Mod 4 <> 0 Then
				MsgBox ("Number of pins is incorrect!! Differential Pair Signal should have multiple of 4 pins!!", vbCritical)
				Exit All
			End If

			For i = 0 To nPins/2-1
				'loading Mixed Mode Converter
				CreateMixedModeConverter

'-------------------------------------------------------------------------------- major difference between case 1, 2 and 3
				'linking MM Conv to DS Canvas
				If txdelay = False Then
					If PinsArray(i*2).BusSize > 1 Then
						ReportInformation "Bus pin " + PinsArray(i*2).Name + " was skipped."
					Else
						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B",selectedBlock, PinsArray(i*2).Name)))
					End If ' BusSize(i*2) > 1

					If PinsArray(i*2+1).BusSize > 1 Then
						ReportInformation "Bus pin " + PinsArray(i*2+1).Name + " was skipped."
					Else
							'Linking MM Conv to MWS Block
							ConnectComponents(Array(Array("B","MOD" + Str(i+1), "N"), Array("B",selectedBlock, PinsArray(i*2+1).Name)))
					End If ' BusSize(i*2+1) > 1
				Else
							TLname = 1
							CreateTLBlock
							TLshiftY = -TLshiftY
							ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B","TL" + Str(i+1) + "-" + Str(TLname),"1")))

							If PinsArray(i*2).BusSize > 1 Then
								ReportInformation "Bus pin " + PinsArray(i*2).Name + " was skipped."
							Else
								ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i*2).Name)))
							End If ' BusSize(i*2) > 1

							TLname = 2
							CreateTLBlock
							TLshiftY = -TLshiftY

							ConnectComponents(Array(Array("B","MOD" + Str(i+1), "N"), Array("B","TL" + Str(i+1) + "-" + Str(TLname),"1")))

							If PinsArray(i*2+1).BusSize > 1 Then
								ReportInformation "Bus pin " + PinsArray(i*2+1).Name + " was skipped."
							Else
								ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i*2+1).Name)))
							End If ' BusSize(i*2+1) > 1
							TLname = 1
				End If


'--------------------------------------------------------------------------------
				If DiffPort = False Then
					DifferentialLoad
				Else
					DifferentialPort
				End If

				If CommPort = False Then
					CommonModeLoad
				Else
					CommonModePort
				End If


				'determine PosX and PosY location --> 250 = size of MM Conv
				posY = posY - 300

				If (i+1) = nPins/4 Then
					posX = -posX
					posY = (nPins/4-1)*150
					TLshiftX = -TLshiftX
				End If

			Next i


		Case 2
			'If no. of port is not even, display error message, then quit
			If nPins Mod 4 <> 0 Then
				MsgBox ("Number of pins is incorrect!! Differential Pair Signal should have multiple of 4 pins!!", vbCritical)
				Exit All
			End If

			Dim nIncrement As Integer
			For i = 0 To nPins/2-1
				'loading Mixed Mode Converter
				CreateMixedModeConverter

				If txdelay = False Then
'-------------------------------------------------------------------------------- major difference between case 1, 2 and 3
				If i Mod 2 = 0 And i <> 0 Then
					nIncrement = nIncrement + 2
				End If
				If PinsArray(i+nIncrement).BusSize > 1 Then
					ReportInformation "Bus pin " + PinsArray(i+nIncrement).Name + " was skipped."
				Else
					'linking MM Conv to DS Canvas
					ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B",selectedBlock,PinsArray(i+nIncrement).Name)))
				End If ' BusSize(i+nIncrement) > 1

				If PinsArray(i+2+nIncrement).BusSize > 1 Then
					ReportInformation "Bus pin " + PinsArray(i+2+nIncrement).Name + " was skipped."
				Else
					'Linking MM Conv to MWS Block
					ConnectComponents(Array(Array("B","MOD" + Str(i+1), "N"), Array("B",selectedBlock,PinsArray(i+2+nIncrement).Name)))
				End If ' BusSize(i+2+nIncrement) > 1
'--------------------------------------------------------------------------------
				Else
						TLname = 1
						CreateTLBlock
						TLshiftY = -TLshiftY
						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B","TL" + Str(i+1) + "-" + Str(TLname),"1")))

						If i Mod 2 = 0 And i <> 0 Then
							nIncrement = nIncrement + 2
						End If
							
						If PinsArray(i+nIncrement).BusSize > 1 Then
							ReportInformation "Bus pin " + PinsArray(i+nIncrement).Name + " was skipped."
						Else
							ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i+nIncrement).Name)))
						End If

						TLname = 2
						CreateTLBlock
						TLshiftY = -TLshiftY

						ConnectComponents(Array(Array("B", "MOD" + Str(i+1), "N"), Array("B", "TL" + Str(i+1) + "-" + Str(TLname), "1")))

						If PinsArray(i+2+nIncrement).BusSize > 1 Then
							ReportInformation "Bus pin " + PinsArray(i+2+nIncrement).Name + " was skipped."
						Else
							ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i+2+nIncrement).Name)))
						End If
						TLname = 1
				End If
				

				If DiffPort = False Then
					DifferentialLoad
				Else
					DifferentialPort
				End If

				If CommPort = False Then
					CommonModeLoad
				Else
					CommonModePort
				End If


				'determine PosX and PosY location --> 250 = size of MM Conv
				posY = posY - 300

				If (i+1) = nPins/4 Then
					posX = -posX
					posY = (nPins/4-1)*150
					TLshiftX = -TLshiftX
				End If

			Next i


		Case 3
			'If no. of pins is not even, display error message, then quit
			If nPins Mod 4 <> 0 Then
				MsgBox ("Number of pins is incorrect!! Differential Pair Signal should have multiple of 4 pins!!", vbCritical)
				Exit All
			End If


			PortJump = True

			For i = 0 To nPins/2-1
				'loading Mixed Mode Converter
				CreateMixedModeConverter

				If txdelay = False Then
'-------------------------------------------------------------------------------- major difference between case 1, 2 and 3
					If PinsArray(i*2).BusSize > 1 Then
						ReportInformation "Bus pin " + PinsArray(i*2).Name + " was skipped."
					Else
						'linking MM Conv to DS Canvas
						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B",selectedBlock,PinsArray(i*2).Name)))
					End If

					If PinsArray(i*2+1).BusSize > 1 Then
						ReportInformation "Bus pin " + PinsArray(i*2+1).Name + " was skipped."
					Else
						'Linking MM Conv to MWS Block
						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "N"), Array("B",selectedBlock,PinsArray(i*2+1).Name)))
					End If
'--------------------------------------------------------------------------------

			Else
						TLname = 1
						CreateTLBlock
						TLshiftY = -TLshiftY
						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "P"), Array("B","TL" + Str(i+1) + "-" + Str(TLname),"1")))

						If PinsArray(i*2).BusSize > 1 Then
							ReportInformation "Bus pin " + PinsArray(i*2).Name + " was skipped."
						Else
							ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i*2).Name)))
						End If


						TLname = 2
						CreateTLBlock
						TLshiftY = -TLshiftY

						ConnectComponents(Array(Array("B","MOD" + Str(i+1), "N"), Array("B","TL" + Str(i+1) + "-" + Str(TLname),"1")))

						If PinsArray(i*2+1).BusSize > 1 Then
							ReportInformation "Bus pin " + PinsArray(i*2+1).Name + " was skipped."
						Else
							ConnectComponents(Array(Array("B","TL" + Str(i+1) + "-" + Str(TLname), "2"), Array("B",selectedBlock,PinsArray(i*2+1).Name)))
						End If

						TLname = 1
				End If

				If DiffPort = False Then
					DifferentialLoad
				Else
					DifferentialPort
				End If

				If CommPort = False Then
					CommonModeLoad
				Else
					CommonModePort
				End If


				'determine PosX and PosY location --> 250 = size of MM Conv
				posY = posY - 300
				NameJump = NameJump + 1

				If (i+1) = nPins/4 Then
					posX = -posX
					posY = (nPins/4-1)*150
					NameJump = -NameJump+1
					TLshiftX = -TLshiftX
				End If

			Next i


	End Select





End Sub
