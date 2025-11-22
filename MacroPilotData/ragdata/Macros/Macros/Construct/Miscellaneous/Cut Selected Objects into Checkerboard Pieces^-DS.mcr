'#Language "WWB-COM"

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 02-May-2014 ube: First version
' ================================================================================================
Option Explicit

Sub Main

	Dim SelectedTreeItems(1000) As String, SelectedItems(1000) As String
	Dim SelectedComponents(1000) As String, SelectedSolids(1000) As String
	Dim umin(1000) As Double, vmin(1000) As Double, wmin(1000) As Double
	Dim umax(1000) As Double, vmax(1000) As Double, wmax(1000) As Double
	Dim umin_global As Double, vmin_global As Double, wmin_global As Double
	Dim umax_global As Double, vmax_global As Double, wmax_global As Double
	Dim x As Double, y As Double, z As Double, u As Double, v As Double, w As Double, u0 As Double, v0 As Double, w0 As Double
	Dim CommandContents As String, CommandName As String, Item As String, sParent As String
	Dim n_ustep As Long, n_vstep As Long, n_wstep As Long
	Dim ustep As Double, vstep As Double, wstep As Double
	Dim ustart As Double, ustop As Double, vstart As Double, vstop As Double, wstart As Double, wstop As Double
	Dim n As Long, m As Long, k As Long, n_selected_items As Long, Index As Long

	ustep = 1.0
	vstep = 1.0
	wstep = 1.0

	Begin Dialog UserDialog 230,190,"Checkerboard Macro",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 5,5,220,110,"Give the step sizes for cutting",.GroupBox1
		'Text 15,21,290,21,"Give the step sizes for cutting:",.Text1
		Text 20,26,80,21,"U-axis:",.Text2
		TextBox 80,25,80,21,.ustep
		Text 20,56,80,21,"V-axis:",.Text3
		TextBox 80,55,80,21,.vstep
		Text 20,86,80,21,"W-axis:",.Text4
		TextBox 80,85,80,21,.wstep

		GroupBox 5,120,220,40,"2D or 3D cutting",.GroupBox2
		CheckBox 20,140,140,14,"2D (v- and w-axis)",.twod
		CheckBox 170,140,40,14,"3D",.threed

		OKButton 10,164,100,21
		CancelButton 120,164,100,21

	End Dialog
	Dim dlg As UserDialog
	dlg.ustep = CStr(ustep)
	dlg.vstep = CStr(vstep)
	dlg.wstep = CStr(wstep)
	dlg.twod = True

	Dialog dlg

	ustep = CDbl(dlg.ustep)
	vstep = CDbl(dlg.vstep)
	wstep = CDbl(dlg.wstep)

	If dlg.twod And dlg.threed Then
		MsgBox ("Select only 2D or 3D cutting","Error")
		End
	End If
	If dlg.twod Then
		' Do nothing.
	ElseIf dlg.threed Then
		' Do nothing.
	Else
		MsgBox ("Select either 2D or 3D cutting","Error")
		End
	End If

	' Go through selected tree items, and check which are solids, which are components or groups.
	' Get all sub-components and add them to selected items.
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
			SelectedItems(k) = Item ' Selected item is a shape.
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
					SelectedItems(k) = Item ' Add solid to selected items.
					k = k+1
					sParent = Item
   					Item = Resulttree.GetNextItemName( Item )
	   				While (Item = "")
						Index = InStrRev( sParent, "\" )
						If (Index = 0) Then
							GoTo Finish
						Else
							sParent = Left$( sParent, Index-1 )
							Item = Resulttree.GetNextItemName( sParent )
						End If
	  	  			Wend
	  	  		End If
			Wend
			Finish:
		End If
	Next n

	' Go through the solids, and store their names and component names.
	WCS.Store("Checkerboard_macro_temporary_wcs")
	umax_global = -1e200
	vmax_global = -1e200
	wmax_global = -1e200
	umin_global = 1e200
	vmin_global = 1e200
	wmin_global = 1e200
	m = k-1
	For n = 0 To m
		SelectedSolids(n) = Right(SelectedItems(n),Len(SelectedItems(n))-InStrRev(SelectedItems(n),"\"))
		SelectedComponents(n) = Left(SelectedItems(n),InStrRev(SelectedItems(n),"\")-1)
		SelectedComponents(n) = Right(SelectedComponents(n),Len(SelectedComponents(n))-InStr(SelectedComponents(n),"\"))
		SelectedComponents(n) = Replace(SelectedComponents(n),"\","/")
		' Go through the points and get the minimum and maximum coordinates
		umin(n) = 1e200
		vmin(n) = 1e200
		wmin(n) = 1e200
		umax(n) = -1e200
		vmax(n) = -1e200
		wmax(n) = -1e200
		For k = 1 To Solid.GetNumberOfPoints(SelectedComponents(n)+":"+SelectedSolids(n))
			Solid.GetPointCoordinates(SelectedComponents(n)+":"+SelectedSolids(n),CStr(k),x,y,z)
			WCS.GetWCSPointFromGlobal("Checkerboard_macro_temporary_wcs",u,v,w,x,y,z)
			If wmax(n) < w Then
				wmax(n) = w
			End If
			If wmin(n) > w Then
				wmin(n) = w
			End If
			If umax(n) < u Then
				umax(n) = u
			End If
			If umin(n) > u Then
				umin(n) = u
			End If
			If vmax(n) < v Then
				vmax(n) = v
			End If
			If vmin(n) > v Then
				vmin(n) = v
			End	If
		Next k
		' Also store the global min. and max. coordinates
		If wmax_global < wmax(n) Then
			wmax_global = wmax(n)
		End If
		If wmin_global > wmin(n) Then
			wmin_global = wmin(n)
		End If
		If umax_global < umax(n) Then
			umax_global = umax(n)
		End If
		If umin_global > umin(n) Then
			umin_global = umin(n)
		End If
		If vmax_global < vmax(n) Then
			vmax_global = vmax(n)
		End If
		If vmin_global > vmin(n) Then
			vmin_global = vmin(n)
		End	If
	Next n

	'Debug.Print CStr(umin_global)+" "+CStr(vmin_global)+" "+CStr(wmin_global)
	'Debug.Print CStr(umax_global)+" "+CStr(vmax_global)+" "+CStr(wmax_global)

	' Compute the starting and end points.
	If Abs(umin_global) - Abs(Round(umin_global/ustep)*ustep) > 0 Then
		u0 = Sgn(umin_global)*(Abs(Round(umin_global/ustep))*ustep + ustep)
	Else
		u0 = Sgn(umin_global)*(Abs(Round(umin_global/ustep))*ustep)
	End If
	If Abs(vmax_global) - Abs(Round(vmax_global/vstep)*vstep) > 0 Then
		v0 = Sgn(vmin_global)*(Abs(Round(vmin_global/vstep))*vstep + vstep)
	Else
		v0 = Sgn(vmin_global)*(Abs(Round(vmin_global/vstep))*vstep)
	End If
	If Abs(wmin_global) - Abs(Round(wmin_global/wstep)*wstep) > 0 Then
		w0 = Sgn(wmin_global)*(Abs(Round(wmin_global/wstep))*wstep + wstep)
	Else
		w0 = Sgn(wmin_global)*(Abs(Round(wmin_global/wstep))*wstep)
	End If


	' Go through the solids and create new components for them, move solids into these components.
	CommandContents = ""
	For n = 0 To m
		CommandContents = CommandContents + "Component.New """+SelectedComponents(n)+"/"+SelectedSolids(n)+""""+vbLf
		CommandContents = CommandContents + "Solid.ChangeComponent """+SelectedComponents(n)+":"+SelectedSolids(n)+""", """+SelectedComponents(n)+"/"+SelectedSolids(n)+""""+vbLf
	Next n

	Debug.Print "Starting slicing..."

	'Debug.Print CStr(u0)+" "+CStr(v0)+" "+CStr(w0)
	'Debug.Print CStr(umax_global)+" "+CStr(vmax_global)+" "+CStr(wmax_global)

	' Go through the components and make the slicing commands to be inserted into history.
	If dlg.threed Then
		' u-axis
		CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""90.0"""+vbLf
		CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(u0)+""""+vbLf
		w = u0
		While w < umax_global
			For n = 0 To m
				If umin(n) < w And umax(n) > w Then
					CommandContents = CommandContents + "Solid.SliceComponent """+SelectedComponents(n)+"/"+SelectedSolids(n)+""""+vbLf
				End If
			Next n
			w = w + ustep
			CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(vstep)+""""+vbLf
		Wend
		CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(-w)+""""+vbLf
		CommandContents = CommandContents + "WCS.RotateWCS ""v"", ""-90.0"""+vbLf
	End If

	' v-axis
	CommandContents = CommandContents + "WCS.RotateWCS ""u"", ""-90.0"""+vbLf
	CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(v0)+""""+vbLf
	w = v0
	While w < vmax_global
		For n = 0 To m
			If vmin(n) < w And vmax(n) > w Then
				CommandContents = CommandContents + "Solid.SliceComponent """+SelectedComponents(n)+"/"+SelectedSolids(n)+""""+vbLf
			End If
		Next n
		w = w + vstep
		CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(vstep)+""""+vbLf
	Wend
	CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(-w)+""""+vbLf
	CommandContents = CommandContents + "WCS.RotateWCS ""u"", ""90.0"""+vbLf

	' w-axis
	CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(w0)+""""+vbLf
	w = w0
	While w < wmax_global
		For n = 0 To m
			If wmin(n) < w And wmax(n) > w Then
				CommandContents = CommandContents + "Solid.SliceComponent """+SelectedComponents(n)+"/"+SelectedSolids(n)+""""+vbLf
			End If
		Next n
		w = w + wstep
		CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(wstep)+""""+vbLf
	Wend
	CommandContents = CommandContents + "WCS.MoveWCS ""local"", ""0.0"", ""0.0"", """+CStr(-w)+""""+vbLf


	' Add commands to the history list.
	CommandName = "Checkerboard macro:"
	For n = 0 To n_selected_items
		CommandName = CommandName+" "+Right(SelectedTreeItems(n),Len(SelectedTreeItems(n))-InStr(SelectedTreeItems(n),"\"))
	Next n

	AddToHistory(CommandName, CommandContents)

	WCS.Delete("Checkerboard_macro_temporary_wcs")

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

	bCritical = False

	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changed or button pressed
		Select Case Item
		Case "OK"

		Case "Cancel"
			End
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
