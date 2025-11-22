
' ==========================================================================
' This VBA code calculates the envelope for a 1D result entry for all existing RunIDs
' ==========================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 17-Oct-2022 ube: little update
' 29-Aug-2022 ube: first version
' ==========================================================================

Option Explicit

Sub Main

	Dim sTreeName As String
	sTreeName = GetSelectedTreeItem

	If (Left(sTreeName,11) <> "1D Results\") And (Left(sTreeName,18) <> "Tables\1D Results\") Then
		MsgBox "Please select a single 1D result entry in the tree before running this macro. (a)"
		Exit All
	End If

	If Resulttree.GetFirstChildName(sTreeName) <> "" Then
		' it is a folder
		MsgBox "Please select a single 1D result entry in the tree before running this macro. (b)"
		Exit All
	End If

	Begin Dialog UserDialog 690,133,"Calculate Envelope from All RunIDs" ' %GRID:10,7,1,1
		Text 20,14,210,14,"1D Result to be processed",.Text1
		TextBox 20,35,640,21,.sTreeName
		OKButton 20,105,90,21
		CancelButton 120,105,90,21
		OptionGroup .complex
			OptionButton 30,70,50,14,"Re",.OptionButton1
			OptionButton 100,70,50,14,"Im",.OptionButton2
			OptionButton 160,70,60,14,"Mag",.OptionButton3
			OptionButton 230,70,90,14,"MagdB10",.OptionButton4
			OptionButton 330,70,100,14,"MagdB20",.OptionButton5
			OptionButton 440,70,50,14,"Ph",.OptionButton6
	End Dialog
	Dim dlg As UserDialog
	dlg.complex = 0
	dlg.sTreeName = sTreeName
	If (Dialog(dlg) = 0) Then Exit All

	Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	Dim n As Long

	sTreeName = dlg.sTreeName
	nResults = Resulttree.GetTreeResults(sTreeName,"0D/1D recursive","filetype0D1D",paths,types,files,info)

	If nResults <> 1 Then
		MsgBox "Please enter a tree entry, containing a single 1D result entry. (a)"
		Exit All
	End If

	If types(0) <> "XYSIGNAL" Then
		MsgBox "Please enter a tree entry, containing a single 1D result entry. (b)"
		Exit All
	End If

	Dim result As Object, o1DRunID As Object, omax As Object, omin As Object
	Dim i As Long, j As Long

	Dim bminXvaluesDiasagree As Boolean
	Dim bmaxXvaluesDiasagree As Boolean
	bminXvaluesDiasagree = False
	bmaxXvaluesDiasagree = False

	Dim IDs As Variant
	IDs = Resulttree.GetResultIDsFromTreeItem(sTreeName)

	If IsEmpty(IDs) Then
		ReportInformationToWindow("No parametric data available.")
	Else
		For n = 0 To UBound(IDs)
			Set result = Resulttree.GetResultFromTreeItem(sTreeName, IDs(n))
			Select Case result.GetResultObjectType()
			Case "1DC"
				'transfer to real1d with selected property re/im/dB/...
				Select Case dlg.complex
				Case 0  ' Re
					Set o1DRunID = result.Real
				Case 1  ' Im
					Set o1DRunID = result.Imaginary
				Case 2  ' Mag
					Set o1DRunID = result.Magnitude
				Case 3  ' MagdB10
					Set o1DRunID = result.Magnitude
					With o1DRunID
						For i = 0 To .GetN-1
							If .GetY(i) > 0 Then .SetY(i, 10*Log(.GetY(i))/Log(10))
						Next
					End With
				Case 4  ' MagdB20
					Set o1DRunID = result.Magnitude
					With o1DRunID
						For i = 0 To .GetN-1
							If .GetY(i) > 0 Then .SetY(i, 20*Log(.GetY(i))/Log(10))
						Next
					End With
				Case 5  ' Ph
					Set o1DRunID = result.Phase
				End Select
			Case "1D"
				'transfer to real1d with selected property re/im/dB/...
				Set o1DRunID = result
			Case Else
				MsgBox "neither 1D nor 1DC, exit all"
				Exit All
			End Select

			If (n = 0) Then
				' copy first RunId as reference into omin and omax
				Set omin = o1DRunID.Copy
				Set omax = o1DRunID.Copy
			Else
				' for 2nd, 3rd, etc RunID sompare with min and max curve to get the final envelope
				' Note check x-values to agree across different RunIDs

				With o1DRunID
					If .GetN <> omin.GetN Or .GetN <> omax.GetN Then
						MsgBox "runids with different amount of data, exit all"
						Exit All
					Else
						If .GetX(0) <> omin.GetX(0) Or .GetX(0) <> omax.GetX(0) Or .GetN < 1 Then
							MsgBox "runids - first data point different x-value, exit all"
							Exit All
						Else
							If .GetX(.GetN-1) <> omin.GetX(omin.GetN-1) Or .GetX(.GetN-1) <> omax.GetX(omin.GetN-1) Then
								MsgBox "runids - last data point different x-value, exit all"
								Exit All
							Else
								' ===== FINALLY now we are confident, that x-axis are identical - to compare min max y-values ===========
								For i = 0 To .GetN-1
									If .GetX(i) = omin.GetX(i) Then
										If .GetY(i) < omin.GetY(i) Then
											omin.SetY(i, .GetY(i))
										End If
									Else
										bminXvaluesDiasagree = True
									End If
									If .GetX(i) = omax.GetX(i) Then
										If .GetY(i) > omax.GetY(i) Then
											omax.SetY(i, .GetY(i))
										End If
									Else
										bmaxXvaluesDiasagree = True
									End If
								Next i
							End If
						End If
					End If
				End With
			End If
		Next n
	End If

	If bminXvaluesDiasagree Or bmaxXvaluesDiasagree Then
		MsgBox "some x-axis values disagreed between runids"
	End If

	With omin
		.Save("min-envelope.sig")
		.AddToTree("1D Results\Envelope\min")
	End With
	With omax
		.Save("max-envelope.sig")
		.AddToTree("1D Results\Envelope\max")
	End With

	SelectTreeItem "1D Results\Envelope\"

End Sub
