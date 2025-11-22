' *Graphics / 3D Plot Max-Field xyz-Position
'
' macro.966
'
' Finds the location of the maximum field quantity and plots a pick-point
'
' ================================================================================================
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 07-Jul-2020 ube: re-add SelectTreeItem "Components" in addition to SelectModelView
' 15-Jun-2020 thn: use SelectModelView instead of SelectTreeItem "Components"
' 06-Sep-2017 lwe: switch to component view instead of mesh view (CST-50517)
' 15-May-2017 ube: correct behaviour if local WCS is active
' 29-Mar-2006 fhi: Msg if no 3D plot is selected
' 28-Mar-2006 fhi: Initial version
' ================================================================================================
'
' select a 3d-fieldplot first, then execute macro
'----------------------------------------------
Option Explicit

Sub Main
	Dim x As Double, y As Double, z As Double, mmax As Double

	screenupdating False'True
	On Error GoTo no_3d_field_active
	mmax = GetFieldPlotMaximumPos ( x,  y,  z) 			'Absolute Koordinaten des Maximums
	
	' change view (from 3d fieldplot into the geom. view)
	SelectTreeItem "Components"
	SelectModelView
	
	If WCS.IsWCSActive = "local" Then
		WCS.ActivateWCS("global")
		Pick.PickPointFromCoordinates  x,y,z	' plot a pick-point
		WCS.ActivateWCS("local")
	Else
		Pick.PickPointFromCoordinates  x,y,z	' plot a pick-point
	End If
 	
 	Plot.Wireframe True
 	Exit All
 	
 	no_3d_field_active:
 	MsgBox " Please select a 3D field result first!",,"Information"
End Sub
