'#Language "WWB-COM"

' ================================================================================================
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ============================================================================================================================
' 22-Sep-2015 ube: first version
' ============================================================================================================================

Option Explicit

Public Const sxyz = Array("X","Y","Z")

Sub Main
	Dim i As Long, j As Long
	Dim via_mat_name() As String
	j = -1
	With Material
		For i=0 To .GetNumberOfMaterials-1
			If InStr(.GetNameOfMaterialFromIndex(i),"VIA") > 0 Then
				j=j+1
				ReDim Preserve via_mat_name(j)
				via_mat_name(j) = .GetNameOfMaterialFromIndex(i)
			End If
		Next
	End With

	If j=-1 Then
		MsgBox "No Material name contains the string VIA, Exit Macro."
		Exit All
	End If

	Begin Dialog UserDialog 720,182,"Define current monitors around vias" ' %GRID:10,7,1,1
		DropListBox 30,28,650,128,via_mat_name(),.via_mat_name
		OKButton 30,154,90,21
		CancelButton 130,154,90,21
		Text 20,7,340,14,"Choose Material of VIA-Layer, to be monitored",.Text1
		Text 30,70,180,14,"Normal direction of PCB:",.Text2
		DropListBox 270,67,140,121,sxyz(),.sxyz
		Text 30,105,220,14,"Relative Radius-Factor  (RRF):",.Text3
		Text 40,130,430,14,"(Current Monitor will have RRF times the radius of each single via.)",.Text4
		TextBox 270,102,140,21,.RF
	End Dialog
	Dim dlg As UserDialog

	dlg.RF = "2.0"
	dlg.sxyz = 2

	If (Dialog(dlg)=0) Then Exit All

	Dim solidname As String, compname As String, sfullname As String, i1 As Long

	Dim sCommand As String

'	ScreenUpdating False  '  GetLooseBoundingBoxOfShape fails if ScreenUpdating is set to false
	LockTree True
	SetLock True

	With Solid
		For i=0 To .GetNumberOfShapes-1
			If .GetMaterialNameForShape(.GetNameOfShapeFromIndex (i)) = via_mat_name(dlg.via_mat_name) Then
				sfullname = .GetNameOfShapeFromIndex(i)
				i1 = InStr(sfullname, ":")
				solidname = Mid(sfullname,i1+1)
				compname = Left(sfullname,i1-1)
				sCommand = "Solid.SplitShape """ + solidname + """, """ + compname + """"
				AddToHistory "Separate Shape " + solidname, sCommand
			End If
		Next
	End With

	Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, z1 As Double, z2 As Double
	Dim xctr As Double, yctr As Double, zctr As Double, radius As Double, rx As Double, ry As Double, rz As Double

	Dim imoni As Long
	imoni = 0
	With Solid
		For i=0 To .GetNumberOfShapes-1
			If .GetMaterialNameForShape(.GetNameOfShapeFromIndex (i)) = via_mat_name(dlg.via_mat_name) Then
				sfullname = .GetNameOfShapeFromIndex(i)
				i1 = InStr(sfullname, ":")
				solidname = Mid(sfullname,i1+1)
				compname = Left(sfullname,i1-1)
				If .GetLooseBoundingBoxOfShape("solid$"+.GetNameOfShapeFromIndex(i), x1,x2,y1,y2,z1,z2) Then
					xctr = (x1+x2)/2
					yctr = (y1+y2)/2
					zctr = (z1+z2)/2
					rx   = (x2-x1)/2
					ry   = (y2-y1)/2
					rz   = (z2-z1)/2
					Select Case dlg.sxyz
					Case 0  ' x
						MsgBox "normal x not yet implemented, exit"
						Exit All
						radius = IIf(ry>rz,ry,rz)
					Case 1  ' y
						MsgBox "normal y not yet implemented, exit"
						Exit All
						radius = IIf(rx>rz,rx,rz)
					Case 2  ' z
						radius = IIf(ry>rx,ry,rx)
					End Select

					sCommand = "WCS.ActivateWCS ""local""" + vbCrLf

					sCommand = sCommand + "With WCS" + vbCrLf
					sCommand = sCommand + " .SetNormal 0, 0, 1" + vbCrLf
					sCommand = sCommand + " .SetOrigin 0, 0, " + Cstr(zctr) + vbCrLf
					sCommand = sCommand + " .SetUVector 1, 0, 0" + vbCrLf
					sCommand = sCommand + "End With" + vbCrLf

					sCommand = sCommand + "With Circle" + vbCrLf
					sCommand = sCommand + " .Reset" + vbCrLf
					sCommand = sCommand + " .Name ""circle1""" + vbCrLf
					sCommand = sCommand + " .Curve ""Curve-via-monitors""" + vbCrLf
					sCommand = sCommand + " .Radius " + Cstr(radius*Evaluate(dlg.RF)) + vbCrLf
					sCommand = sCommand + " .Xcenter " + Cstr(xctr) + vbCrLf
					sCommand = sCommand + " .Ycenter " + Cstr(yctr) + vbCrLf
					sCommand = sCommand + " .Segments ""0""" + vbCrLf
					sCommand = sCommand + " .Create" + vbCrLf
					sCommand = sCommand + "End With" + vbCrLf

					imoni = imoni + 1

					sCommand = sCommand + "With CurrentMonitor" + vbCrLf
					sCommand = sCommand + " .Reset" + vbCrLf
					sCommand = sCommand + " .Name ""current-" + Cstr(imoni) + " - " + solidname + """" + vbCrLf
					sCommand = sCommand + " .Curve ""Curve-via-monitors:circle1""" + vbCrLf
					sCommand = sCommand + " .InvertOrientation ""False""" + vbCrLf
					sCommand = sCommand + " .Add" + vbCrLf
					sCommand = sCommand + "End With" + vbCrLf

					sCommand = sCommand + "WCS.ActivateWCS ""global""" + vbCrLf

					' MsgBox sCommand

					AddToHistory "define current monitor " + Cstr(imoni) + " - " + solidname, sCommand

				End If
			End If
		Next
	End With

	ScreenUpdating True
	LockTree False
	SetLock False

	MsgBox "Current Monitors successfully defined.", vbInformation

End Sub
