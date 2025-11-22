'#Language "WWB-COM"

' Bend selected solids at UV-plane -macro.
'
' Slices the selected components by UV-plane and bends the parts on the negative W-axis side by user-defined angle
' towards the positive U-axis.
'
' Instructions for use:
'
' 1) Select all components that you want to bend.
' 2) Place WCS in the place where components should be bent, so that axes are aligned with the object.
'    U-axis normal to the top most face of the selected solids, V- and W-axis tangential. 

'-------------------------------------------------------------------------------------------------
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-------------------------------------------------------------------------------------------------
' 01-Oct-2015 iha: corrected wcs handling
' 28-May-2015 iha: Integrated the bending by face rotation
' 11-Aug-2014 iha: Several improvements, e.g. arbitrary angle bending
' 02-May-2014 iha: First checked-in version
'-------------------------------------------------------------------------------------------------


Option Explicit
Dim maxSize%                                ' Present max size allocated for dynamic arrays
Const NEWNAME$ = "CSTShapes"                ' Temporary name appended to new rotated segments created
Const PositionUncertainity# = 1e-6          ' Uncertainity in the position of a point (Used while detecting if a face lies on the UV plane i.e W = 0 )
Const UVTranslate# = 0.05                   ' Translation in the W direction for the second UV plane slicing (Optional Feature for increasing speed for complex models). Can
                                            ' be made variable and dependent on the structure
Sub Main

	Dim Test As String
	Dim DlgResult As Boolean, BooleanAdd As Boolean, DebugFlag As Boolean
	Dim BendingAngle As String, BendingRadius As String, NSeg As String, SegmentLength As String, FileName As String
	Dim angle As Double, radius As Double, nsegments As Long

	Begin Dialog UserDialog 500,378,"Bend selected objects at UV Plane",.dialogfunc ' %GRID:10,7,1,1
		GroupBox 10,77,480,133,"Bending properties",.Settings
		Text 30,105,110,14,"Bending angle:",.text1
		TextBox 170,98,90,21,.angle
		Text 30,133,150,14,"Number of segments:",.textsegments
		Text 30,161,150,14,"Segment length:",.textlength
		TextBox 170,126,90,21,.nsegments
		TextBox 170,154,90,21,.length
		Text 30,189,150,14,"Bending radius:",.text_radius
		TextBox 170,182,90,21,.radius
		GroupBox 10,217,480,119,"Debug settings",.GroupBox1

		CheckBox 30,245,140,14,"Use Boolean add",.booleanadd

		PushButton 30,308,90,21,"Browse",.Browse
		TextBox 130,308,350,21,.filename
		CheckBox 30,280,290,14,"Debug (macro is written to file and not run)",.debugflag

		OKButton 10,350,90,21
		CancelButton 120,350,90,21
		GroupBox 10,7,480,63,"Bending type",.GroupBox2
		OptionGroup .options
			OptionButton 30,35,180,14,"Slice and copy"
			OptionButton 260,35,160,14,"Slice and rotate"
		PushButton 400,350,90,21,"Help",.Help
	End Dialog
	Dim dlg As UserDialog
	dlg.angle = "90"
	dlg.booleanadd = True
	dlg.nsegments = "0"
	dlg.length = "0"
	dlg.radius = "0"

	DlgResult = Dialog(dlg)
	If DlgResult = False Then
		End
	End If

	BendingAngle = dlg.angle
	angle = CDbl(BendingAngle)
	BendingRadius = dlg.radius
	radius = CDbl(BendingRadius)
	BooleanAdd = CBool(dlg.booleanadd)
	DebugFlag = CBool(dlg.debugflag)
	FileName = dlg.filename
	NSeg = dlg.nsegments
	nsegments = CLng(dlg.nsegments)
	SegmentLength = dlg.length

	If DebugFlag Then
		If FileName = "" Then
			MsgBox("Empty filename", vbOkOnly, "Error")
			End
		End If
	End If

	If dlg.options = 0 Then
		SliceAndCopy(BendingAngle, SegmentLength, NSeg, BendingRadius, nsegments, BooleanAdd, DebugFlag, FileName)
	ElseIf dlg.options = 1 Then
		SliceAndRotate(angle, radius, BooleanAdd, DebugFlag, FileName)
	End If

End Sub

Sub SliceAndCopy(BendingAngle As String, SegmentLength As String, NSeg As String, BendingRadius As String, nsegments As Long, BooleanAdd As Boolean, DebugFlag As Boolean, FileName As String)

	Dim SelectedTreeItems(1000) As String, SelectedItems(1000) As String, SelectedComponents(1000) As String, SelectedSolids(1000) As String
	Dim SelectedSolids1(1000) As String, SelectedSolids2(1000) As String, SelectedSolids3(1000) As String
	Dim SelectedSolids4(1000) As String, SelectedSolids5(1000) As String
	Dim SelectedSolids6(1000) As String, SelectedSolids7(1000) As String
	Dim SelectedSolidsNew(1000) As String, SelectedSolidsOrig(1000) As String
	Dim CommandContents As String, CommandName As String, Item As String, sParent As String
	Dim n As Long, m As Long, k As Long, n_selected_items As Long, index As Long
	Dim zmin(1000) As Double, zmax(1000) As Double, xmin(1000) As Double, xmax(1000) As Double, ymin(1000) As Double, ymax(1000) As Double
	Dim x As Double, y As Double, z As Double, u As Double, v As Double, w As Double, radius As Double, length As Double
	Dim DlgResult As Boolean
	Dim RotateAngle As String, SliceAngle As String
	Dim p As Long
	Dim xcoord As String, ycoord As String

	Debug.Clear

	' Go through selected tree items, and check which are solids, which are components or groups.
	n = 0
	SelectedTreeItems(n) = GetSelectedTreeItem
	While SelectedTreeItems(n) <> ""
		n = n+1
		SelectedTreeItems(n) = GetNextSelectedTreeItem
	Wend
	m = n-1
	n_selected_items = m

	k = 0
	n = 0
	Item = SelectedTreeItems(n)

	For n = 0 To m
		Item = SelectedTreeItems(n)

		If Not HasChildren(Item) Then
			SelectedItems(k) = Item
			'Debug.Print "k = "+CStr(k)+", "+SelectedItems(k)
			k = k+1
		Else
			' Selected item is a component
			While StrComp( Item, "", 0 )

				If ( CStr(InStr(Item, SelectedTreeItems(n))) = "0" ) Then
					GoTo Finish
				Else

					While HasChildren( Item ) = True
						Item = Resulttree.GetFirstChildName ( Item )
					Wend

					SelectedItems(k) = Item
					k = k+1

					sParent = Item
   					Item = Resulttree.GetNextItemName( Item )

	   				While (Item = "")

						index = InStrRev( sParent, "\" )

						If (index = 0) Then
							GoTo Finish
						Else
							sParent = Left$( sParent, index-1 )
							Item = Resulttree.GetNextItemName( sParent )

						End If

	  	  			Wend
	  	  		End If
			Wend

			Finish:

		End If

	Next n

	CommandContents = ""
	' Disable tree update to speed up operations.
	CommandContents = CommandContents + "Resulttree.EnableTreeUpdate(False)"+vbLf

	' Go through the solids, and store their names and component names, and min. and max. coordinates.
	m = k-1
	For n = 0 To m
		SelectedSolids(n) = Right(SelectedItems(n),Len(SelectedItems(n))-InStrRev(SelectedItems(n),"\"))
		SelectedComponents(n) = Left(SelectedItems(n),InStrRev(SelectedItems(n),"\")-1)
		SelectedComponents(n) = Right(SelectedComponents(n),Len(SelectedComponents(n))-InStr(SelectedComponents(n),"\"))
		SelectedComponents(n) = Replace(SelectedComponents(n),"\","/")
	Next n

	WCS.Store("Temporary")
	CommandContents = CommandContents + "WCS.Store(""Temporary"")"+vbLf
	For n = 0 To m
		zmax(n) = -1e20
		zmin(n) = 1e20
		xmax(n) = -1e20
		xmin(n) = 1e20
		ymax(n) = -1e20
		ymin(n) = 1e20
		For k = 1 To Solid.GetNumberOfPoints(SelectedComponents(n)+":"+SelectedSolids(n))
			Solid.GetPointCoordinates(SelectedComponents(n)+":"+SelectedSolids(n),CStr(k),x,y,z)
			WCS.GetWCSPointFromGlobal("Temporary",u,v,w,x,y,z)
			If zmax(n) < w Then
				zmax(n) = w
			End If
			If zmin(n) > w Then
				zmin(n) = w
			End If
			If xmax(n) < u Then
				xmax(n) = u
			End If
			If xmin(n) > u Then
				xmin(n) = u
			End If
			If ymax(n) < v Then
				ymax(n) = v
			End If
			If ymin(n) > v Then
				ymin(n) = v
			End	If

		Next k
	Next n


	' Check which solids are completely on the negative w-axis side and can be just transformed,
	' which solids are completely on the positive w-axis side and can be left alone,
	' and which ones need to be sliced.
	k = 0
	If nsegments > 0 Then
		xcoord = SegmentLength+"*(cosd("+BendingAngle+")*"+CStr(nsegments)
		ycoord = "-"+SegmentLength+"*(sind("+BendingAngle+")*"+CStr(nsegments)
		For p = 1 To nsegments
			xcoord = xcoord + " - cosd("+Cstr(p)+"*"+BendingAngle+"/"+CStr(nsegments+1)+")"
			ycoord = ycoord + " - sind("+Cstr(p)+"*"+BendingAngle+"/"+CStr(nsegments+1)+")"
		Next p
	End If
	xcoord = xcoord + ")"
	ycoord = ycoord + ")"
	For n = 0 To m
		If zmax(n) <= 0.0 Then
			' These solids need to be transformed.
			'Debug.Print "Transform: "+SelectedComponents(n)+":"+SelectedSolids(n)
			CommandContents = CommandContents+"With Transform"+vbLf
			CommandContents = CommandContents+"		.Reset"+vbLf
			CommandContents = CommandContents+"    .Name """+SelectedComponents(n)+":"+SelectedSolids(n)+""""+vbLf
			CommandContents = CommandContents+"    .Origin ""Free"""+vbLf
			CommandContents = CommandContents+"    .Center ""0"", ""0"", ""0"""+vbLf
			CommandContents = CommandContents+"    .Angle ""0"", ""-"+CStr(BendingAngle)+""", ""0"""+vbLf
			CommandContents = CommandContents+"    .MultipleObjects ""False"""+vbLf
			CommandContents = CommandContents+"    .GroupObjects ""False"""+vbLf
			CommandContents = CommandContents+"    .Repetitions ""1"""+vbLf
			CommandContents = CommandContents+"    .MultipleSelection ""True"""+vbLf
			CommandContents = CommandContents+"    .RotateAdvanced"+vbLf
			CommandContents = CommandContents+"End With"+vbLf
			If nsegments > 0 Then
				CommandContents = CommandContents+"With Transform"+vbLf
				CommandContents = CommandContents+"		.Reset"+vbLf
    			CommandContents = CommandContents+"		.Name """+SelectedComponents(n)+":"+SelectedSolids(n)+""""+vbLf
    			CommandContents = CommandContents+"    	.Vector """+ycoord+""", ""0.0"", """+xcoord+""""+vbLf
     			CommandContents = CommandContents+"    	.UsePickedPoints ""False"""+vbLf
     			CommandContents = CommandContents+"    	.InvertPickedPoints ""False"""+vbLf
     			CommandContents = CommandContents+"    	.MultipleObjects ""False"""+vbLf
     			CommandContents = CommandContents+"    	.GroupObjects ""False"""+vbLf
     			CommandContents = CommandContents+"    	.Repetitions ""1"""+vbLf
     			CommandContents = CommandContents+"    	.MultipleSelection ""False"""+vbLf
     			CommandContents = CommandContents+"    	.Transform ""Shape"", ""Translate"""+vbLf
				CommandContents = CommandContents+"End With"+vbLf
			End If
		ElseIf zmin(n) > 0.0 Then
			' Do nothing.
			'Debug.Print "Do nothing: "+SelectedComponents(n)+":"+SelectedSolids(n)
		Else
		' These solids must be sliced.
		'Debug.Print "Slice: "+SelectedComponents(n)+":"+SelectedSolids(n)
		SelectedComponents(k) = SelectedComponents(n)
		SelectedSolids(k) = SelectedSolids(n)
		xmin(k) = xmin(n)
		xmax(k) = xmax(n)
		ymin(k) = ymin(n)
		ymax(k) = ymax(n)
		zmin(k) = zmin(n)
		zmax(k) = zmax(n)
		k = k + 1
	End If
	Next n
	m = k - 1

	' Rename solids at first to get unique name.
	For n = 0 To m
		CommandContents = CommandContents+"Solid.Rename """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids(n)+"_temp"""+vbLf
		SelectedSolidsOrig(n) = SelectedSolids(n)
		SelectedSolids(n) = SelectedSolids(n)+"_temp" ' Original part
		SelectedSolids2(n) = SelectedSolids(n)+"_1" ' Sliced and rotated part from original
		SelectedSolids3(n) = SelectedSolids(n)+"_2"
		SelectedSolids4(n) = SelectedSolids(n)+"_2_1"
		SelectedSolids5(n) = SelectedSolids(n)+"_2_1_1"
		'SelectedSolids6(n) = SelectedSolids(n)+"_1_2"
	Next n

	For p = 0 To nsegments

	' Slice the selected solids at current WCS.
	For n = 0 To m
		'Debug.Print "Add to list: "+SelectedSolids(n)+""", """+SelectedComponents(n)
		CommandContents = CommandContents + "Solid.SliceShape """+SelectedSolids(n)+""", """+SelectedComponents(n)+"""" + vbLf
	Next n

	If nsegments = 0 Then
		RotateAngle = BendingAngle
		SliceAngle = BendingAngle+"*0.5"
	ElseIf p = 0  Then
		RotateAngle = BendingAngle+"*0.5/"+NSeg
		SliceAngle = BendingAngle+"*0.25/"+NSeg
	ElseIf p = nsegments Then
		RotateAngle = BendingAngle+"*0.5/"+NSeg
		SliceAngle = BendingAngle+"*0.25/"+NSeg
	Else
		RotateAngle = BendingAngle+"/"+NSeg
		SliceAngle = BendingAngle+"*0.5/"+NSeg
	End If

	' Rotate the other half by user-defined angle.
	For n = 0 To m
		CommandContents = CommandContents+"With Transform"+vbLf
		CommandContents = CommandContents+"    .Reset"+vbLf
		CommandContents = CommandContents+"    .Name """+SelectedComponents(n)+":"+SelectedSolids2(n)+""""+vbLf
		CommandContents = CommandContents+"    .Origin ""Free"""+vbLf
		CommandContents = CommandContents+"    .Center ""0"", ""0"", ""0"""+vbLf
		CommandContents = CommandContents+"    .Angle ""0"", ""-"+RotateAngle+""", ""0"""+vbLf
		CommandContents = CommandContents+"    .MultipleObjects ""False"""+vbLf
		CommandContents = CommandContents+"    .GroupObjects ""False"""+vbLf
		CommandContents = CommandContents+"    .Repetitions ""1"""+vbLf
		CommandContents = CommandContents+"    .MultipleSelection ""True"""+vbLf
		CommandContents = CommandContents+"    .RotateAdvanced"+vbLf
		CommandContents = CommandContents+"End With"+vbLf
	Next n

	' Rotate WCS around v-axis by user-defined angle and slice all components
	CommandContents = CommandContents + "WCS.RotateWCS ""v"", """+SliceAngle+"""" + vbLf
	For n = 0 To m
		CommandContents = CommandContents + "Solid.SliceShape """+SelectedSolids(n)+""", """+SelectedComponents(n)+"""" + vbLf
	Next n
	CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+SliceAngle+"""" + vbLf

	' Mirror and copy the sliced part of the original component.
	For n = 0 To m
		CommandContents = CommandContents+"With Transform"+vbLf
		CommandContents = CommandContents+"    .Reset"+vbLf
		CommandContents = CommandContents+"    .Name """+SelectedComponents(n)+":"+SelectedSolids3(n)+""""+vbLf
		CommandContents = CommandContents+"    .Origin ""Free"""+vbLf
		CommandContents = CommandContents+"    .Center ""0"", ""0"", ""0"""+vbLf
		CommandContents = CommandContents+"    .PlaneNormal ""0"", ""0"", ""-1"""+vbLf
		CommandContents = CommandContents+"    .MultipleObjects ""True"""+vbLf
		CommandContents = CommandContents+"    .GroupObjects ""False"""+vbLf
		CommandContents = CommandContents+"    .Repetitions ""1"""+vbLf
		CommandContents = CommandContents+"    .MultipleSelection ""True"""+vbLf
		CommandContents = CommandContents+"    .Transform ""Shape"", ""Mirror"""+vbLf
		CommandContents = CommandContents+"End With"+vbLf
	Next n

	CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+SliceAngle+"""" + vbLf

	' Mirror and copy the sliced part of the original component.
	For n = 0 To m
		CommandContents = CommandContents+"With Transform"+vbLf
		CommandContents = CommandContents+"    .Reset"+vbLf
		CommandContents = CommandContents+"    .Name """+SelectedComponents(n)+":"+SelectedSolids4(n)+""""+vbLf
		CommandContents = CommandContents+"    .Origin ""Free"""+vbLf
		CommandContents = CommandContents+"    .Center ""0"", ""0"", ""0"""+vbLf
		CommandContents = CommandContents+"    .PlaneNormal ""0"", ""0"", ""-1"""+vbLf
		CommandContents = CommandContents+"    .MultipleObjects ""True"""+vbLf
		CommandContents = CommandContents+"    .GroupObjects ""False"""+vbLf
		CommandContents = CommandContents+"    .Repetitions ""1"""+vbLf
		CommandContents = CommandContents+"    .MultipleSelection ""True"""+vbLf
		CommandContents = CommandContents+"    .Transform ""Shape"", ""Mirror"""+vbLf
		CommandContents = CommandContents+"End With"+vbLf
	Next n

	CommandContents = CommandContents + "WCS.RotateWCS ""v"", """+SliceAngle+"""" + vbLf

	If BooleanAdd Then
		' Boolean add all new solids together.
		For n = 0 To m
			CommandContents = CommandContents + "Solid.Add """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids2(n)+""""+vbLf
			CommandContents = CommandContents + "Solid.Add """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids3(n)+""""+vbLf
			CommandContents = CommandContents + "Solid.Add """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids4(n)+""""+vbLf
			CommandContents = CommandContents + "Solid.Add """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids5(n)+""""+vbLf
			'CommandContents = CommandContents + "Solid.Add """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolids6(n)+""""+vbLf

		Next n
	End If

	If nsegments > 0 Then
		If nsegments = 1 Then
			CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+RotateAngle+"""" + vbLf
			CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", ""-"+SegmentLength+"""" + vbLf
		ElseIf p = 0 Then
			CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+RotateAngle+"""" + vbLf
			CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", ""-"+SegmentLength+"""" + vbLf
			CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+RotateAngle+"""" + vbLf
		ElseIf p = 1 Then
			CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+SliceAngle+"""" + vbLf
			CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", ""-"+SegmentLength+"""" + vbLf
		Else
			CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-"+RotateAngle+"""" + vbLf
			CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", ""-"+SegmentLength+"""" + vbLf
		End If
	End If

	'Debug.Print CommandContents
	If BooleanAdd = False And nsegments > 0 Then
		' Rename solids at first to get unique name.
		For n = 0 To m
			SelectedSolidsOrig(n) = SelectedSolids2(n)
			SelectedSolids(n) = SelectedSolids2(n) ' Original part
			SelectedSolids2(n) = SelectedSolids(n)+"_1" ' Sliced and rotated part from original
			SelectedSolids3(n) = SelectedSolids(n)+"_2"
			SelectedSolids4(n) = SelectedSolids(n)+"_2_1"
			SelectedSolids5(n) = SelectedSolids(n)+"_2_1_1"
			'SelectedSolids6(n) = SelectedSolids(n)+"_1_2"
	'		Debug.Print "Round "+CSTr(p)
	'		Debug.Print SelectedSolids(n)
	'		Debug.Print SelectedSolids2(n)
	'		Debug.Print SelectedSolids3(n)
	'		Debug.Print SelectedSolids4(n)
	'		Debug.Print SelectedSolids5(n)
	'		Debug.Print "=============="
		Next n
	End If

	'Debug.Print CommandContents
	'Debug.Print "=========="

	Next p ' End of segments

	If BooleanAdd Then
		' Rename solids to original names.
		For n = 0 To m
			CommandContents = CommandContents+"Solid.Rename """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+":"+SelectedSolidsOrig(n)+""""+vbLf
		Next n
	End If

	CommandContents = CommandContents + "WCS.Restore(""Temporary"")"+vbLf
	CommandContents = CommandContents + "WCS.Delete(""Temporary"")"+vbLf


	' Enable tree update again, also updates the tree.
	CommandContents = CommandContents + "Resulttree.EnableTreeUpdate(True)"+vbLf

	If DebugFlag Then
		Open FileName For Output As #111
		Write #111, CommandContents
		Close #111
	Else
		' Add commands to history list.
		CommandName = "Macro: Bend selected objects at UV-plane:"
		For n = 0 To n_selected_items
			CommandName = CommandName+" "+Right(SelectedTreeItems(n),Len(SelectedTreeItems(n))-InStr(SelectedTreeItems(n),"\"))
		Next n
		AddToHistory(CommandName, CommandContents)
	End If

End Sub

Function HasChildren( Item As String ) As Boolean

	Dim Name As String
	Dim sChild As String

	Name = Item
	sChild = Resulttree.GetFirstChildName ( Name )
	If sChild = "" Then
		HasChildren = False
	Else
		HasChildren = True
	End If

End Function

Function DialogFunc%(Item As String, Action As Integer, value As Integer)

	Dim extension As String
	Dim bCritical As Boolean
	Dim Filename As String
	bCritical = False
	Dim angle As Double, length As Double, radius As Double
	Dim nsegments As Double

	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changed or button pressed
		Select Case Item
		Case "OK"

		Case "Cancel"

		Case "Help"
			MsgBox("See the PDF-file in the CST STUDIO SUITE installation directory in the sub-directory ..\Library\Macros\Construct\Miscellaneous.", vbOkOnly, "Help")
			DialogFunc = True

		Case "Browse"
			Filename = GetFilePath(, , ,"Select file" ,2)
			DlgText "filename", Filename
			DialogFunc = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
		'Debug.Print "Item = "+ Item+", Action = "+CStr(Action)+", value = "+CStr(value)+ ", segment = "+DlgText("length")+", radius = "+DlgText("radius")
		angle = CDbl(DlgText("angle"))
		nsegments = CLng(DlgText("nsegments"))
		radius = CDbl(DlgText("radius"))
		length = CDbl(DlgText("length"))

		If nsegments > 1 Then
			angle = angle/nsegments
		End If

		Select Case Item
		Case "angle"
			If nsegments > 0 And radius > 0 Then
				length = 2*radius*sind(angle/2)
				DlgText "length", CStr(length)
			End If
		Case "nsegments"
			If nsegments > 0 And radius > 0 Then
				length = 2*radius*sind(angle/2)
				DlgText "length", CStr(length)
			End If
		Case "length"
			radius = length/(2*sind(angle/2))
			DlgText "radius", CStr(radius)
		Case "radius"
			length = 2*radius*sind(angle/2)
			DlgText "length", CStr(length)
		End Select
		DialogFunc = True
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function

Sub SliceAndRotate(angle As Double, radius As Double, BooleanAdd As Boolean, DebugFlag As Boolean, Filename As String)

	Dim axisVector#(5)    ' axisVector(5) stores the values of the end points of the rotation edge defined by the user ; the order is u1,v1,w1,u2,v2,w2 for index 0 to5
                                            ' angle is the rotation angle
	Dim rotateObjInPosW As Boolean          ' True/False for rotation in positive/negative W coordinate side

	Dim hStringMoveWCS$, hStringRestoreWCS$
	Dim hStringRotateObj$, hStringRotateFace$, hStringWCSTransform$                                 ' History strings for various operations
	Dim hStringSeparateUVSlice$, hStringSeparateShape$
	Dim hStringMergeUVSlice$, hStringMergeShape$, hStringMergeRotatedSegment$, hStringMergeAll$
    Dim hStringSeparateUVSlice2$, hStringMergeUVSlice2$, hStringWCSTransform2$, hStringWCSTransform2Back$

	Dim selectedObj$()                      ' Array for storing user selected shapes
	Dim numSelectedObj%

	Dim slicePosObj$(), sliceNegObj$(), newSliceObj$()     ' Arrays for storing shapes which get split on UV plane slicing
	Dim numSlicePosObj%, numSliceNegObj% , numNewSliceObj%

	Dim noSlicePosObj$(), noSliceNegObj$()                 ' Arrays for storing shapes which do not get split on UV plane slicing
	Dim numNoSlicePosObj%, numNoSliceNegObj%

	' Store original WCS.
	Dim hStringWCS As String
	hStringWCS = "WCS.Store ""temp_orig""" + vbLf

	' If radius is non-zero, move WCS first.
	radius = radius*0.5
	If radius > 0 Then
		'WCS.MoveWCS "local", radius, "0.0", "0.0"
		hStringMoveWCS = "WCS.MoveWCS ""local"", """+CStr(radius)+""", ""0.0"", ""0.0"""+vbLf
	End If


' Initial Settings
	Dim hStringStoreWCS$
	Resulttree.EnableTreeUpdate(False)
    maxSize = Solid.GetNumberOFShapes
    Pick.ClearAllPicks
    WCS.ActivateWCS("local")
    WCS.Store("temp_local1")
	hStringStoreWCS = "WCS.Store ""temp_local1""" + vbLf

' Assigning values to userInputs
	axisVector(0) =  0
	axisVector(1) = 10
	axisVector(2) =  0
	axisVector(3) =  0
	axisVector(4) = -10
	axisVector(5) =  0
'	angle = 180
	rotateObjInPosW = False


' Finding all Selected Shapes for bending
	numSelectedObj = 0
	ReDim selectedObj(maxSize)
	getSelectedObjects(selectedObj,numSelectedObj)
	ReDim Preserve selectedObj(numSelectedObj)


' Peform UV slicing
    ' Initialising strings and arrays
		hStringSeparateUVSlice = ""
		hStringMergeUVSlice = ""
        numSlicePosObj = 0
        ReDim slicePosObj(numSelectedObj)
		numSliceNegObj = 0
        ReDim sliceNegObj(numSelectedObj)
        numNoSlicePosObj = 0
        ReDim noSlicePosObj(numSelectedObj)
        numNoSliceNegObj = 0
        ReDim noSliceNegObj(numSelectedObj)

    ' Function for performing UV plane slicing. Slices all the shapes in selectedObj and returns split and unsplit shapes. Also returns the history strings for separation and merge operations
	   performUVSlicing( slicePosObj, numSlicePosObj, sliceNegObj, numSliceNegObj, noSlicePosObj, numNoSlicePosObj, noSliceNegObj, numNoSliceNegObj, selectedObj, numSelectedObj, hStringSeparateUVSlice, hStringMergeUVSlice )


' Performing 2nd UV slicing for improving speeds in complex structures with shapes having a lot of faces
	hStringSeparateUVSlice2 = ""
	hStringMergeUVSlice2 = ""

	If( rotateObjInPosW ) Then
	    'WCS.MoveWCS( "local",0,0,-UVTranslate)   ' Since slicing performed at an offset from the current UV plane position
		'WCS.RotateWCS( "v" , 180)
	    hStringWCSTransform2 = "WCS.MoveWCS " + Chr(34) + "local" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + Cstr(-UVTranslate) + Chr(34) + vbLf
	    hStringWCSTransform2 = hStringWCSTransform2 + "WCS.RotateWCS " + Chr(34) + "v" + Chr(34) + "," + Chr(34) + "180" + Chr(34) + vbLf
		performUVSlicing2(sliceNegObj, numSliceNegObj, hStringSeparateUVSlice2, hStringMergeUVSlice2 )
	    'WCS.RotateWCS( "v" , 180)
	    'WCS.MoveWCS( "local",0,0,UVTranslate)
	    hStringWCSTransform2Back = "WCS.RotateWCS " + Chr(34) + "v" + Chr(34) + "," + Chr(34) + "180" + Chr(34) + vbLf
	    hStringWCSTransform2Back = hStringWCSTransform2Back + "WCS.MoveWCS " + Chr(34) + "local" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + Cstr(UVTranslate) + Chr(34) + vbLf
	Else
	    'WCS.MoveWCS( "local",0,0,UVTranslate)
	    'WCS.RotateWCS( "v" , 180)
	    hStringWCSTransform2 = "WCS.MoveWCS " + Chr(34) + "local" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + Cstr(UVTranslate) + Chr(34) + vbLf
	    hStringWCSTransform2 = hStringWCSTransform2 + "WCS.RotateWCS " + Chr(34) + "v" + Chr(34) + "," + Chr(34) + "180" + Chr(34) + vbLf
		performUVSlicing2(slicePosObj, numSlicePosObj, hStringSeparateUVSlice2, hStringMergeUVSlice2 )
	    'WCS.RotateWCS( "v" , 180)
	    'WCS.MoveWCS( "local",0,0,-UVTranslate)
	    hStringWCSTransform2Back = "WCS.RotateWCS " + Chr(34) + "v" + Chr(34) + "," + Chr(34) + "180" + Chr(34) + vbLf
	    hStringWCSTransform2Back = hStringWCSTransform2Back + "WCS.MoveWCS " + Chr(34) + "local" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + Cstr(-UVTranslate) + Chr(34) + vbLf
	End If



' Rotation of objects and faces

	' Transforming WCS for performing rotation operations
	  hStringWCSTransform = ""
	  getWCSTransformStrings( axisVector, hStringWCSTransform )

	' Initialising arrays and strings for rotation
      numNewSliceObj = 0
      ReDim newSliceObj(maxSize)
      hStringSeparateShape = ""
	  hStringMergeShape = ""
	  hStringMergeRotatedSegment = ""
	  hStringRotateObj = ""
	  hStringRotateFace = ""

	If( rotateObjInPosW ) Then

		' Function calls for rotating objects
	 	getRotateObjStrings( noSlicePosObj, numNoSlicePosObj, hStringRotateObj, angle )
	    getRotateObjStrings( slicePosObj, numSlicePosObj, hStringRotateObj, angle )

	   ' Function calls for rotating faces
	    performShapeSeparation( newSliceObj, numNewSliceObj, sliceNegObj, numSliceNegObj,  hStringSeparateShape, hStringMergeShape)
	    performFaceRotation( sliceNegObj, numSliceNegObj, hStringRotateFace, angle, hStringMergeRotatedSegment )
	    performFaceRotation( newSliceObj, numNewSliceObj, hStringRotateFace, angle, hStringMergeRotatedSegment )

	 Else

	 	' Function calls for rotating objects
	    getRotateObjStrings( noSliceNegObj, numNoSliceNegObj, hStringRotateObj, angle)
	    getRotateObjStrings( sliceNegObj, numSliceNegObj, hStringRotateObj, angle )

	   ' Function calls for rotating faces
	    performShapeSeparation( newSliceObj, numNewSliceObj, slicePosObj, numSlicePosObj, hStringSeparateShape, hStringMergeShape)
	    performFaceRotation( slicePosObj, numSlicePosObj, hStringRotateFace, angle, hStringMergeRotatedSegment )
	    performFaceRotation( newSliceObj, numNewSliceObj, hStringRotateFace, angle, hStringMergeRotatedSegment )

	 End If

	' Restore WCS if needed.
	'WCS.Restore "temp_orig"
	hStringRestoreWCS = "WCS.Restore ""temp_orig"""+vbLf
	hStringRestoreWCS = hStringRestoreWCS + "WCS.Delete ""temp_orig"""+vbLf
	hStringRestoreWCS = hStringRestoreWCS + "WCS.Delete ""temp_local1"""+vbLf
	'hStringRestoreWCS = "WCS.Restore ""temp_local1"""+vbLf
	'hStringRestoreWCS = hStringRestoreWCS + "WCS.Delete ""temp_local1"""+vbLf
	'If radius > 0 Then
		''WCS.MoveWCS "local", -radius, "0.0", "0.0"
		'hStringRestoreWCS = hStringRestoreWCS + "WCS.MoveWCS ""local"", """+CStr(-radius)+""", ""0.0"", ""0.0"""+vbLf
	'End If

	If BooleanAdd Then
		hStringMergeAll = hStringMergeRotatedSegment + hStringMergeShape + hStringMergeUVSlice2 + hStringMergeUVSlice    ' Combining all merge strings
	Else
		hStringMergeAll = ""
	End If

' Writing strings to History List
	Dim sHistory As String
	sHistory = hStringWCS+hStringMoveWCS+hStringStoreWCS+hStringSeparateUVSlice+hStringWCSTransform2+hStringSeparateUVSlice2+hStringWCSTransform2Back+hStringSeparateShape+hStringWCSTransform+hStringRotateObj+hStringRotateFace+hStringMergeAll+hStringRestoreWCS
	If DebugFlag Then
		Open Filename For Output As #111
		Write #111, sHistory
		Close #111
	Else
		AddToHistory("Macro: Bend selected objects at UV-plane" , sHistory)
	End If



End Sub


' This function finds the correct faces for the given shapes and returns history list strings for creating and merging rotated segments
Function performFaceRotation( allShapes$(), numAllShapes%, hRotate$, angle# , hMerge$)
	Dim count%, count2%, num%
	Dim faceId As Long
    Dim currentShape$
    Dim x#,y#,z#

	For count = 0 To (numAllShapes-1)
        currentShape = allShapes(count)
		Pick.PickFaceChainFromId(currentShape,Solid.GetAnyFaceIdFromSolid(currentShape))
        num = Pick.GetNumberOfPickedFaces

        For count2 = 0 To num-1
          Pick.getPickedFaceFromIndex( 1,faceId )
          Pick.PickFaceCenterpointFromIndex(0)
          Pick.getPickpointCoordinates((count2 + 1),x,y,z)
          WCS.getWCSPointFromGlobal("temp_local1",x,y,z,x,y,z)
          If(Abs(z) < PositionUncertainity) Then
          getFaceRotateStrings( currentShape, faceId, hRotate, angle, hMerge )
          End If
        Next
       Pick.ClearAllPicks
     Next
End Function


' This function returns history list strings for rotating a given face and merging(boolean add) the rotated segments
Function getFaceRotateStrings( solidName$, faceId As Long, hString$, angle# , hString2$ )

    Dim tempSt$(1), temp$, temp0$, dummyAxisVector#(5)
    separateComponentName( solidName,tempSt )

    dummyAxisVector(0) = -10
    dummyAxisVector(1) = 0
    dummyAxisVector(2) = 0
    dummyAxisVector(3) =  10
    dummyAxisVector(4) = 0
    dummyAxisVector(5) = 0

	' Pick Edge History String
	hString = hString +  "Pick.AddEdge " + Chr(34) + Cstr(dummyAxisVector(0)) + Chr(34) + "," + Chr(34) + Cstr(dummyAxisVector(1)) + Chr(34) + "," + Chr(34) + Cstr(dummyAxisVector(2)) + Chr(34) + ","
	hString = hString + Chr(34) + Cstr(dummyAxisVector(3)) + Chr(34) + "," + Chr(34) + Cstr(dummyAxisVector(4)) + Chr(34) + "," + Chr(34) + Cstr(dummyAxisVector(5)) + Chr(34) + vbLf

    ' Pick Face History String
	hString = hString + "Pick.PickFaceFromId " + Chr(34) + solidName + Chr(34) + "," + Chr(34) + Cstr(faceId) + Chr(34) +vbLf

    ' Rotate Face Strings

     temp = ""
     temp = temp + ".Component " + Chr(34) + tempSt(0) + Chr(34) + vbLf
     temp = temp + ".Material " + Chr(34) + Solid.GetMaterialNameForShape(solidName) + Chr(34) + vbLf
     temp = temp + ".Mode " + Chr(34) + "Picks" + Chr(34) + vbLf
     temp = temp + ".Angle "  + Chr(34) + Cstr(angle) + Chr(34) + vbLf
     temp = temp + ".Height " + Chr(34) + "0.0" + Chr(34) + vbLf
     temp = temp + ".RadiusRatio " + Chr(34) + "1.0" + Chr(34) + vbLf
     temp = temp + ".Nsteps " + Chr(34) + "0" + Chr(34) + vbLf

     temp = temp + ".SplitClosedEdges " + Chr(34) + "True" + Chr(34) + vbLf
     temp = temp + ".SegmentedProfile " + Chr(34) + "False" + Chr(34) + vbLf
     temp = temp + ".DeleteBaseFaceSolid " + Chr(34) + "False" + Chr(34) + vbLf
     temp = temp + ".ClearPickedFace " + Chr(34) + "True" + Chr(34) + vbLf
     temp = temp + ".SimplifySolid " + Chr(34) + "True" + Chr(34) + vbLf
     temp = temp + ".UseAdvancedSegmentedRotation " + Chr(34) + "True" + Chr(34) + vbLf
     temp = temp + ".Create" + vbLf
     temp = temp + "End With" + vbLf

     temp0 = ""
     temp0 = temp0 + "With Rotate" + vbLf
     temp0 = temp0 + ".Reset" + vbLf
     temp0 = temp0 + ".Name "

     hString = hString + temp0 + Chr(34) + tempSt(1) + NEWNAME + Cstr(faceId) + Chr(34) + vbLf + temp

   ' Merge Strings
     hString2 = hString2 + "Solid.Add " + Chr(34) + solidName + Chr(34) + "," + Chr(34) + solidName + NEWNAME + Cstr(faceId) + Chr(34) + vbLf

End Function


' This function returns history list strings for separating and merging the given shapes. Also returns the new shapes created
Function performShapeSeparation( newShapes$(), numNewShapes%, allShapes$(), numAllShapes%, hSeparate$, hMerge$ )
  Dim count%, numTotalObj%, numPresentObj%, count2%
  Dim tempA$(1), parentString$, newString$
  numTotalObj = Solid.GetNumberOfShapes

  For count = 0 To (numAllShapes-1)

      parentString = allShapes(count)
      separateComponentName( parentString,tempA )
      Solid.SplitShape( tempA(1), tempA(0) )
      numPresentObj = Solid.GetNumberOfShapes
      If ( numPresentObj > numTotalObj ) Then
        hSeparate = hSeparate + "Solid.SplitShape " + Chr(34) + tempA(1) + Chr(34) + "," + Chr(34) + tempA(0) + Chr(34) + vbLf
        For count2 = numTotalObj To (numPresentObj - 1)
        	newString = Solid.GetNameOfShapeFromIndex(count2)
        	addString( newString, newShapes, numNewShapes)
            hMerge = hMerge + "Solid.Add " + Chr(34) + parentString + Chr(34) + "," + Chr(34) + newString + Chr(34) + vbLf
        Next
        numTotalObj = numPresentObj
      End If
  Next

End Function


' This function returns the history list strings for rotating the shapes
Function getRotateObjStrings( rotateObj$(), numRotateObj% , hString$, angle# )

   Dim count%, temp$, temp0$, tempSt$(1)

   temp = ".Origin " + Chr(34) + "Free" + Chr(34) + vbLf
   temp = temp + ".Center " + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + vbLf
   temp = temp + ".Angle "  + Chr(34) + Cstr(angle) + Chr(34) + "," + Chr(34) + "0" + Chr(34) + "," + Chr(34) + "0" + Chr(34) + vbLf
   temp = temp + ".MultipleObjects " + Chr(34) + "False" + Chr(34) + vbLf
   temp = temp + ".GroupObjects " + Chr(34) + "False" + Chr(34) + vbLf
   temp = temp + ".Repetitions " + Chr(34) + "1" + Chr(34) + vbLf
   temp = temp + ".MultipleSelection " + Chr(34) + "False" + Chr(34) + vbLf
   temp = temp + ".Transform " + Chr(34) + "Shape" + Chr(34)+ "," +  Chr(34) + "Rotate" + Chr(34)  + vbLf
   temp = temp + "End With" + vbLf

   temp0 = "With Transform" + vbLf + ".Reset" + vbLf + ".Name "

  For count = 0 To numRotateObj-1
  hString = hString +  temp0 + Chr(34) + rotateObj(count) + Chr(34) + vbLf + temp
  Next

End Function


' This function changes the WCS for rotation operations
Function getWCSTransformStrings(axisVector#(), hString$)
hString = hString + "WCS.ActivateWCS " + Chr(34) + "local" + Chr(34) +vbLf
hString = hString +  "Pick.AddEdge " + Chr(34) + Cstr(axisVector(0)) + Chr(34) + "," + Chr(34) + Cstr(axisVector(1)) + Chr(34) + "," + Chr(34) + Cstr(axisVector(2)) + Chr(34) + ","
hString = hString + Chr(34) + Cstr(axisVector(3)) + Chr(34) + "," + Chr(34) + Cstr(axisVector(4)) + Chr(34) + "," + Chr(34) + Cstr(axisVector(5)) + Chr(34) + vbLf
hString = hString + "WCS.AlignWCSWithSelected " + Chr(34)  + "EdgeCenter" + Chr(34) + vbLf
hString = hString + "Pick.ClearAllPicks " + vbLf
End Function

' This function performs UV slicing2 for improvign speed for complex structures with shapes having a lot of faces
 Function performUVSlicing2( sliceObj$(), numSliceObj#, hStringSeparate$, hStringMerge$)

 	Dim count%, currentShape$, tempA$(1), numTotalShapes%, numPresentShapes%, newShape$
    numTotalShapes = Solid.GetNumberOFShapes

	  For count = 0 To ( numSliceObj - 1 )
		currentShape = sliceObj(count)
	    separateComponentName( currentShape,tempA )
	    Solid.SliceShape( tempA(1), tempA(0) )
	    numPresentShapes = Solid.GetNumberOFShapes
	    If( numPresentShapes > numTotalShapes ) Then
	    	newShape = Solid.GetNameOfShapeFromIndex(numPresentShapes - 1)
	        hStringSeparate = hStringSeparate + "Solid.SliceShape " + Chr(34) + tempA(1) + Chr(34) + "," + Chr(34) + tempA(0) + Chr(34) + vbLf
	        hStringMerge = hStringMerge + "Solid.Add " + Chr(34) + currentShape + Chr(34) + "," + Chr(34) + newShape + Chr(34) + vbLf
	        numTotalShapes = numPresentShapes
	    End If
	  Next

 End Function


' Function for performing UV plane slicing. Slices all the shapes In selectedObj and returns split And unsplit shapes. Also returns the history strings for separation and merge operations
 Function performUVSlicing( slicePosObj$(), numSlicePosObj%, sliceNegObj$(), numSliceNegObj%, noSlicePosObj$(), numNoSlicePosObj%, noSliceNegObj$(), numNoSliceNegObj%, selectedObj$(), numSelectedObj%, hStringSeparateUVSlice$, hStringMergeUVSlice$ )

 	Dim count%, currentShape$, tempA$(1), numTotalShapes%, numPresentShapes%, newShape$
    numTotalShapes = Solid.GetNumberOFShapes

	  For count = 0 To ( numSelectedObj - 1 )

		currentShape = selectedObj(count)
	    separateComponentName( currentShape,tempA )
	    Solid.SliceShape( tempA(1), tempA(0) )
	    numPresentShapes = Solid.GetNumberOFShapes
	    If( numPresentShapes > numTotalShapes ) Then
	    	newShape = Solid.GetNameOfShapeFromIndex(numPresentShapes - 1)
	        hStringSeparateUVSlice = hStringSeparateUVSlice + "Solid.SliceShape " + Chr(34) + tempA(1) + Chr(34) + "," + Chr(34) + tempA(0) + Chr(34) + vbLf
	        hStringMergeUVSlice = hStringMergeUVSlice + "Solid.Add " + Chr(34) + currentShape + Chr(34) + "," + Chr(34) + newShape + Chr(34) + vbLf
	        slicePosObj(numSlicePosObj) = currentShape
	        numSlicePosObj = numSlicePosObj + 1
	        sliceNegObj(numSliceNegObj) = newShape
	        numSliceNegObj = numSliceNegObj + 1
	        numTotalShapes = numPresentShapes
	    Else
			If ( onPosWDirection(currentShape) ) Then
	           noSlicePosObj(numNoSlicePosObj) = currentShape
	           numNoSlicePosObj = numNoSlicePosObj + 1
			Else
	           noSliceNegObj(numNoSliceNegObj) = currentShape
	           numNoSliceNegObj = numNoSliceNegObj + 1
			End If
	    End If

	  Next

 End Function


' This function returns whether an unsplit shape(on UV slicing) lies on positive/negative W axis
	 Function onPosWDirection(solidName$) As Boolean

	   Dim x#,y#,z#
	   Pick.PickFaceFromID( solidName, Solid.GetAnyFaceIdFromSolid(solidName) )
	   Pick.PickFaceCenterpointFromIndex(0)
	   Pick.getPickpointCoordinates(1,x,y,z)
	   WCS.getWCSPointFromGlobal("temp_local1",x,y,z,x,y,z)
	       If(z>0) Then
	          onPosWDirection = True
	       Else
	          onPosWDirection = False
	       End If
	   Pick.ClearAllPicks

	 End Function


' This function separates the component and shape names
	Function separateComponentName( inString$,tempString$() )
	Dim num%
	num = InStrRev(inString, ":")
	tempString(0) = Left(inString,num-1)
	tempString(1) = Mid(inString,num+1)
	End Function


' This function returns all the selected shapes
Function getSelectedObjects(selectedObj$(), count%)

Dim tempString$, shapeName$, dummyTempString$
tempString = getSelectedTreeItem

While(Not(tempString=""))

	If( Left( tempString,10 ) = "Components" ) Then
         If (Resulttree.getFirstChildName(tempString) = "") Then
             shapeName = getShapeNameFromNTString(tempString)
             If (isValidShape(shapeName)) Then
                 selectedObj(count) = shapeName
                 count = count + 1
             End If
          Else
          	 dummyTempString = tempString
             getAllObjectsInNT(dummyTempString, selectedObj, count)
          End If
    End If

   tempString = getNextSelectedTreeItem
Wend

End Function


' This is a recursive function to get all valid shapes under parentString

Function getAllObjectsInNT( parentString$ , foundStrings$(), count% )

   Dim childString As String, dummyChildString$
   Dim shapeName$

   childString = Resulttree.GetFirstChildName(parentString)

   While( Not (childString = "" ) )

      If ( Resulttree.GetFirstChildName(childString) = "" ) Then

         shapeName = getShapeNameFromNTString(childString)
         If (isValidShape(shapeName)) Then
             foundStrings(count) = shapeName
             count = count + 1
         End If

      Else
       	 dummyChildString = childString
         getAllObjectsInNT( dummyChildString, foundStrings, count)
      End If

   childString = Resulttree.GetNextItemName(childString)
   Wend

  End Function


' This function returns the ShapeName from a string in 'Navigation Tree' string format
' shapeName should be passed by Value and not by Reference
Function getShapeNameFromNTString(ByVal shapeName$) As String

Dim num%
shapeName = Replace(shapeName,"\","/")
num = InStrRev(shapeName, "/")
shapeName = Left( shapeName,num-1 ) + Replace(shapeName,"/",":",num)
shapeName = Mid(shapeName,12)
getShapeNameFromNTString = shapeName

End Function


' Function returns if a shapeName corresponds to a valid shape
Function isValidShape(shapeName$) As Boolean
Dim valid As Boolean
If ( Solid.GetAnyFaceIdFromSolid(shapeName) = -1 ) Then
	valid = False
Else
    valid = True
End If
isValidShape = valid
End Function


' This function adds a new string to the input array and changes array size if required
 Function addString( newString$, foundStrings$(), count%)

 	 If ( count > maxSize) Then
    	maxSize = 3*maxSize
    	ReDim Preserve foundStrings(maxSize)
     End If
        foundStrings(count) = newString
        count = count + 1

 End Function
