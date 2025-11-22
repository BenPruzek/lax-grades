'#Language "WWB-COM"


' ================================================================================================
' This macro allows to align the Farfield Origin to a chosen Anchor rpoint
' ================================================================================================
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' -------------------
' 31-Aug-2020 dta: more robust retrieving of anchor points in presence of tree folder/subfolder
' 11-Dec-2017 ube: increased dialogue width
' 20-Jan-2017 dta: adapted to support anchor points.
' 20-Dec-2016 ses: initial version
' ================================================================================================

Option Explicit


Sub Main
	Dim anchortemp As String, anchorparenttemp As String
	Dim AnchorList() As String, AnchorFolderList() As String, AnchorSubFolderList() As String
	Dim NumAnchorPoints As Integer
	Dim AnchorName As String, AnchorPath As String, AnchorNewPath As String
	Dim IsAnchor As Boolean

	Dim NumFolderAnchorTree As Integer,NumSubFolderAnchorTree As Integer, iii As Integer

	NumAnchorPoints = 0
	NumFolderAnchorTree=0
	NumSubFolderAnchorTree=0

	anchortemp = Resulttree.GetFirstChildName("Anchor Points")

	If anchortemp = "" Then
		MsgBox("Error: Align to stored Anchor Point not possible. No stored Anchor Point found.", 16,)
		End
	Else
		Do
		AnchorName=Mid$(anchortemp,InStrRev(anchortemp,"\")+1)                       'anchor stored with no subfolder

		If AnchorPoint.DoesExist(AnchorName) Then     'check if anchor exists and is not stored in any subfolder

			NumAnchorPoints = NumAnchorPoints + 1
			ReDim Preserve AnchorList(NumAnchorPoints)
			AnchorList(NumAnchorPoints-1) = AnchorName
		 	anchortemp = Resulttree.GetNextItemName(anchortemp)

		Else
			NumFolderAnchorTree = NumFolderAnchorTree + 1
			ReDim Preserve AnchorFolderList(NumFolderAnchorTree)
			AnchorFolderList(NumFolderAnchorTree-1) = AnchorName
		 	anchortemp = Resulttree.GetNextItemName(anchortemp)
		End If

		Loop While anchortemp <> ""

	End If

	For iii=0 To NumFolderAnchorTree-1                                 'Anchor tree folder loop

		anchortemp = Resulttree.GetFirstChildName("Anchor Points\"+AnchorFolderList(iii))        'get to subfolder

		Do
			AnchorName=Mid$(anchortemp,InStrRev(anchortemp,"\")+1)                       'anchor stored with no subfolder
			AnchorPath=Replace(Right(anchortemp,Len(anchortemp)-InStr(anchortemp,"\")),"\","/")
			AnchorPath=Left(AnchorPath,InStrRev(AnchorPath,"/")-1)
			AnchorNewPath=AnchorPath+":"+AnchorName

			If AnchorPoint.DoesExist(AnchorNewPath) Then     'check if anchor exists and is not stored in any subfolder

				NumAnchorPoints = NumAnchorPoints + 1
				ReDim Preserve AnchorList(NumAnchorPoints)
				AnchorList(NumAnchorPoints-1) = AnchorNewPath
				anchortemp = Resulttree.GetNextItemName(anchortemp)

			Else
				NumSubFolderAnchorTree = NumSubFolderAnchorTree + 1
				ReDim Preserve AnchorSubFolderList(NumSubFolderAnchorTree)
				AnchorSubFolderList(NumSubFolderAnchorTree-1) = AnchorPath+"\"+AnchorName

				Do
				 	anchortemp = Resulttree.GetNextItemName(anchortemp)

					If anchortemp <> "" Then

					 	AnchorName=Mid$(anchortemp,InStrRev(anchortemp,"\")+1)                       'anchor stored with no subfolder
						AnchorPath=Replace(Right(anchortemp,Len(anchortemp)-InStr(anchortemp,"\")),"\","/")
						AnchorPath=Left(AnchorPath,InStrRev(AnchorPath,"/")-1)
						AnchorNewPath=AnchorPath+":"+AnchorName

						If Not(AnchorPoint.DoesExist(AnchorNewPath))  Then
							NumSubFolderAnchorTree = NumSubFolderAnchorTree + 1
							ReDim Preserve AnchorSubFolderList(NumSubFolderAnchorTree)
							AnchorSubFolderList(NumSubFolderAnchorTree-1) = AnchorPath+"\"+AnchorName

						ElseIf AnchorPoint.DoesExist(AnchorNewPath) Then

							NumAnchorPoints = NumAnchorPoints + 1
							ReDim Preserve AnchorList(NumAnchorPoints)
							AnchorList(NumAnchorPoints-1) = AnchorNewPath
						 	anchortemp = Resulttree.GetNextItemName(anchortemp)
						End If
					End If

				Loop While anchortemp <> ""

			End If

		Loop While anchortemp <> ""
	Next iii

	For iii=0 To NumSubFolderAnchorTree-1                                 'Anchor tree folder loop

		anchortemp = Resulttree.GetFirstChildName("Anchor Points\"+AnchorSubFolderList(iii))        'access child tree item

		Do
			If anchortemp <> "" Then
			 	AnchorName=Mid$(anchortemp,InStrRev(anchortemp,"\")+1)                       'anchor stored with no subfolder
				AnchorPath=Replace(Right(anchortemp,Len(anchortemp)-InStr(anchortemp,"\")),"\","/")
				AnchorPath=Left(AnchorPath,InStrRev(AnchorPath,"/")-1)
				AnchorNewPath=AnchorPath+":"+AnchorName

				NumAnchorPoints = NumAnchorPoints + 1
				ReDim Preserve AnchorList(NumAnchorPoints)
				AnchorList(NumAnchorPoints-1) = AnchorNewPath
			 	anchortemp = Resulttree.GetNextItemName(anchortemp)
			End If
		Loop While anchortemp <> ""

	Next iii

	Begin Dialog UserDialog 550,301,"Select Stored Anchor Points" ' %GRID:10,7,1,1
		ListBox 20,28,510,238,AnchorList(),.AnchorListbox
		OKButton 10,273,90,21
		CancelButton 110,273,90,21
		Text 20,7,190,14,"Anchor Point List:",.Text1
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg)) Then
		Dim x As Double,y As Double,z As Double
		AnchorPoint.Restore(AnchorList(dlg.AnchorListbox))
		WCS.Store(Replace(AnchorList(dlg.AnchorListbox),":","__"))
		WCS.GetOrigin(Replace(AnchorList(dlg.AnchorListbox),":","__"),x,y,z)
		FarfieldPlot.Origin("free")
		FarfieldPlot.Userorigin(x,y,z)
		ReportInformationToWindow("Farfield origin was set to "+Cstr(x)+","+Cstr(y)+","+Cstr(z))
	End If


End Sub
