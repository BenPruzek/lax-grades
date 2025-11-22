'#Language "WWB-COM"
Option Explicit

'-----------------------------------------------------------------------------
' Copyright 2018-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------
' 30-Jul-2022 ube: code handling of blocks without anchorpoints  (empty anchorpoint array)
' 19-Jan-2022 dta: increased size of dialogue
' 24-Nov-2021 ube: possibility to create copies instead of clones, for some use cases clone blocks have limitations
' 21-Apr-2020 ube: add command ReleaseSnappedAnchorPoints to be able to use different anchorpoint in a second snap
' 19-Jul-2019 fcu: updated .GetTypeName() to work correctly after fix of CST-52569
' 17-Jul-2018 ube: First version
'-----------------------------------------------------------------------------

Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean

	If (Action%=1) Or (Action%=2) Then
		If DlgItem="OK" Then
			If DlgValue("Component") = 0 Or DlgValue("Platform") = 0 Or DlgValue("Component") = DlgValue("Platform") Then
				MsgBox "Please select 2 different blocks for ""Component"" and ""Platform"".",vbExclamation, "Block selection"
				DialogFunction = True
			End If
		End If
	End If

End Function

Sub Main

	' fill block array, only put in blocks containing anchorpoints

	Dim nblocks As Integer, i1 As Integer, nBlocksWithAnchor As Integer

	With Block
		nblocks = .StartBlockNameIteration  ' Resets the block iterator and returns the number of blocks.

		If nblocks = 0 Then
			MsgBox("No parametric blocks defined!","Error")
			Exit Sub
		End If

		Dim aNames() As String, sTempName As String
		ReDim aNames(nblocks)

		nBlocksWithAnchor = 0

		aNames(0) = "Please select a block"
		For i1=1 To nblocks
			With Block
				.Enable3DCommands(False)
				sTempName = .GetNextBlockName
				.Name sTempName
				If (InStr(.GetTypeName, "STUDIO")>0  And  InStr(.GetTypeName, "CST DESIGN STUDIO block")=0 And  InStr(.GetTypeName, "PCB")=0)  Then
					If Not IsEmpty(.GetAnchorPointsList(sTempName)) Then
						If  UBound(.GetAnchorPointsList(sTempName)) > -1 Then
							nBlocksWithAnchor = nBlocksWithAnchor + 1
							aNames(nBlocksWithAnchor) = sTempName
						End If
					End If
				End If
			End With
		Next i1
	End With


	If nBlocksWithAnchor < 2 Then
		MsgBox ("This macros requires at least 2 blocks with defined anchor points.","Error")
		Exit Sub
	End If

	ReDim Preserve aNames(nBlocksWithAnchor)

	Begin Dialog UserDialog 700,266,"Add Block to all anchor points of a platform",.DialogFunction ' %GRID:10,7,1,1

		GroupBox 20,7,660,84,"Select Component and Platform",.GroupBox1
		Text 50,33,270,14,"Component, to be placed multiple times:",.Text1
		DropListBox 360,28,290,21,aNames(),.Component
		Text 50,63,270,14,"Platform, containing multiple anchor points:",.Text2
		DropListBox 360,60,290,28,aNames(),.Platform

		OKButton 30,238,90,21
		CancelButton 130,238,90,21

		GroupBox 20,101,660,98,"Snapping Properties",.GroupBox3
		CheckBox 50,126,160,14,"Flip z-vector",.Flip
		Text 50,150,190,14,"Rotation angle:",.Text4
		Text 50,174,390,14,"Additional longitudinal distance between anchor points:",.Text5
		TextBox 410,147,180,21,.Angle
		TextBox 410,171,180,21,.Distance
		OptionGroup .GroupCopyClone
			OptionButton 40,210,130,14,"Create Copies",.OptionButton1
			OptionButton 190,210,130,14,"Create Clones",.OptionButton2
	End Dialog
	Dim dlg As UserDialog

	dlg.GroupCopyClone = 1

	dlg.Flip = 0
	dlg.Angle = "0.0"
	dlg.Distance = "0.0"

	If (Not Dialog(dlg)) Then
		Exit All
	End If

	Dim anchorNames() As String, sAnchorComponent As String
	Dim sBlockComponent As String, sBlockPlatform As String
	Dim sBlockFileName As String

	sBlockComponent = aNames(dlg.Component)
	sBlockPlatform  = aNames(dlg.Platform)

	With Block
		.Reset
		.Name sBlockComponent
		sBlockFileName = .GetFile
	End With

	anchorNames = Block.GetAnchorPointsList(sBlockComponent)

	If UBound(anchorNames) = 0 Then
		sAnchorComponent = anchorNames(0)
	Else
	Begin Dialog UserDialog 600,210,"Component anchor point" ' %GRID:10,7,1,1
		Text 30,14,340,14,"Please select anchor point, to be used for placement",.Text1
		OKButton 40,182,90,21
		CancelButton 140,182,90,21
		ListBox 40,35,550,140,anchorNames(),.AnchorListBox1
	End Dialog
		Dim dlg2 As UserDialog
		If (Not Dialog(dlg2)) Then
			Exit All
		End If
		sAnchorComponent = anchorNames(dlg2.AnchorListBox1)
	End If

	anchorNames = Block.GetAnchorPointsList(sBlockPlatform)

	Dim sNewBlock As String, sAnchorTemp As String

	Dim N As Integer
	For N = 0 To UBound(anchorNames)
		sAnchorTemp = anchorNames(N)
		ReportInformation("Snap at anchor: " + sAnchorTemp)
		sNewBlock = sBlockComponent
		If N > 0 Then
			sNewBlock = sBlockComponent + "_" + Cstr(N)
			With Block
				.Reset
				.Name (sNewBlock)
				If .DoesExist Then
					If .GetTypeName <> IIf(dlg.GroupCopyClone,"Clone","CST MICROWAVE STUDIO block") Then
						MsgBox "Requested copies/clones already exist, but with DIFFERENT block-type. Please first manually delete those blocks and try again."+vbCrLf+vbCrLf+"Exit Macro",vbExclamation
						Exit All
					End If
				End If
				If dlg.GroupCopyClone Then
					.Type ("Clone")
					.SetClonedBlock (sBlockComponent)
				Else
					.Type ("CSTMWS")
					.SetFile sBlockFileName
				End If
				If Not .DoesExist Then .Create
			End With
		End If


		Dim bFlip As Boolean
		bFlip = IIf(dlg.Flip,1,0)
		
		' dlg.Angle  and dlg.Distance  can be parametric expressions, therefore the expression strings 
		' are used as input to the Snap function and not the values

		Block.ReleaseSnappedAnchorPoints ( sBlockPlatform, sNewBlock )

		Block.SnapAnchorPoints ( sBlockPlatform, sAnchorTemp, sNewBlock, sAnchorComponent, dlg.Angle, dlg.Distance, bFlip )
	Next

End Sub
