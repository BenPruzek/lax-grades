Option Explicit
'#include "vba_globals_all.lib"

' ================================================================================================
' Macro to visualize average volume
'
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 07-Jul-2020 ube: use SelectModelView in addition to SelectTreeItem "Components" (to ensure, 2d3dplot windows become inactive)
' 12-Jul-2011 ube: .Load command added, which updates settings according to selected SAR result.
' 03-Dec-2007 ube: always switch to global WCS (to draw picked edges correctly)
' 13-Jul-2007 ube: draw edges for cube
' 12-Jul-2007 ube: first version
' ================================================================================================
Sub Main

	Dim sarname As String, asarname(99) As String, nh As Integer, s1 As String

	sarname = Resulttree.GetFirstChildName ("2D/3D Results\SAR")
	nh = 0

	While sarname <> ""
		nh = nh + 1
		s1 = sarname
		RemoveFirstChars(s1,18)
		asarname(nh) = s1
		sarname = Resulttree.GetNextItemName (sarname)
	Wend

	If (nh = 0) Then
		MsgBox "No SAR result found. Please calculate SAR first." + vbCrLf + "Exit all.",vbCritical
		Exit All
	End If

	Begin Dialog UserDialog 370,112,"Mark averaging volume" ' %GRID:10,7,1,1
		GroupBox 10,7,350,63,"Select SAR Result",.GroupBox1
		DropListBox 20,28,330,192,asarname(),.asarname
		PushButton 10,84,100,21,"Mark volume",.PushButton1
		PushButton 120,84,90,21,"Exit",.PushButton2
	End Dialog
	Dim dlg As UserDialog
	dlg.asarname = 0
	If (Dialog(dlg) = 2) Then Exit All

	sarname = asarname(1+dlg.asarname)

	With SAR
		.Reset
		.SetLabel sarname
		.Load

		If Dir$(GetProjectPath("Result") + sarname + ".sar")="" Then
			MsgBox "Result file for specified SAR result not found.", vbExclamation
		Else

			SelectTreeItem "Components"
			SelectModelView
			Plot.Wireframe True
			Pick.ClearAllPicks
			WCS.ActivateWCS "global"

			Dim xc As Double
			Dim yc As Double
			Dim zc As Double

			xc = .GetValue("max sar x")
			yc = .GetValue("max sar y")
			zc = .GetValue("max sar z")

			Dim xmin As Double
			Dim ymin As Double
			Dim zmin As Double
			Dim xmax As Double
			Dim ymax As Double
			Dim zmax As Double

			xmin = .GetValue("avg vol min x")
			ymin = .GetValue("avg vol min y")
			zmin = .GetValue("avg vol min z")
			xmax = .GetValue("avg vol max x")
			ymax = .GetValue("avg vol max y")
			zmax = .GetValue("avg vol max z")

			Pick.PickPointFromCoordinates xc,yc,zc

			Pick.PickPointFromCoordinates xmin,ymin,zmin
			Pick.PickPointFromCoordinates xmax,ymin,zmin
			Pick.PickPointFromCoordinates xmin,ymax,zmin
			Pick.PickPointFromCoordinates xmax,ymax,zmin
			Pick.PickPointFromCoordinates xmin,ymin,zmax
			Pick.PickPointFromCoordinates xmax,ymin,zmax
			Pick.PickPointFromCoordinates xmin,ymax,zmax
			Pick.PickPointFromCoordinates xmax,ymax,zmax

			Pick.AddEdge xmin, ymin, zmin, xmax, ymin, zmin
			Pick.AddEdge xmin, ymax, zmin, xmax, ymax, zmin
			Pick.AddEdge xmin, ymin, zmax, xmax, ymin, zmax
			Pick.AddEdge xmin, ymax, zmax, xmax, ymax, zmax

			Pick.AddEdge xmin, ymin, zmin, xmin, ymax, zmin
			Pick.AddEdge xmax, ymin, zmin, xmax, ymax, zmin
			Pick.AddEdge xmin, ymin, zmax, xmin, ymax, zmax
			Pick.AddEdge xmax, ymin, zmax, xmax, ymax, zmax

			Pick.AddEdge xmin, ymin, zmin, xmin, ymin, zmax
			Pick.AddEdge xmin, ymax, zmin, xmin, ymax, zmax
			Pick.AddEdge xmax, ymin, zmin, xmax, ymin, zmax
			Pick.AddEdge xmax, ymax, zmin, xmax, ymax, zmax

			Plot.Update

		End If

	End With

End Sub
