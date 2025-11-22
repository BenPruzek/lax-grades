'#Language "WWB-COM"

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 10-Jun-2024    : Speed up probe creation by disabling tree updates
' 09-Jun-2021 rsj: Cylindrical coordinate is set as default option now. Adapt the dialog func logic.
' 09-Jun-2021 hcm: farfield probe cylindrical coordinate definition using cartesian (vertical) and spherical phi (horizontal) probe type
' 20-Jul-2018 reu: changed casting operators that prevented the input of decimal numbers in the History List
' 15-Jan-2018 mha: added (*) in caption of history list in order to preserve probes when creating simulation projects
' 09-May-2016 ctc: modify wrong z-coordinate formula in E-field probe option, under spherical coordinate
' 27-Oct-2015 ctc: adapted macro for 2016 version
' 21-Oct-2015 ctc: include option for Spherical Coordinate System
' 05-Nov-2014 ctc: include options to define H-field, H-field (farfield), E-field.
' 23-Apr-2014 ctc: initial version

Option Explicit

Dim sHeader As String
Dim sHistEntry As String
Dim dtheta As Double
Dim dphi As Double
Dim radius As Double
Dim k As Double
Dim radiusC As Double
Dim SphSample As Double
Dim angleRes As Double
Dim MinDist As Double
Dim MaxDist As Double
Dim zStep As Double

Sub Main

	Dim PlaneNormal$()
	ReDim Preserve PlaneNormal$(0)
				   PlaneNormal$(0) = "x"
	ReDim Preserve PlaneNormal$(1)
				   PlaneNormal$(1) = "y"
	ReDim Preserve PlaneNormal$(2)
				   PlaneNormal$(2) = "z"
	Dim PlaneNormalC$()
	ReDim Preserve PlaneNormalC$(0)
				   PlaneNormalC$(0) = "x"
	ReDim Preserve PlaneNormalC$(1)
				   PlaneNormalC$(1) = "y"
	ReDim Preserve PlaneNormalC$(2)
				   PlaneNormalC$(2) = "z"


	Begin Dialog UserDialog 350,553,"Generating Field Probes in 3D Volume",.DialogFunc ' %GRID:10,7,1,1
		OKButton 70,511,90,21
		CancelButton 180,511,90,21
		GroupBox 30,4,280,77,"Probe Type",.GroupBox1
		CheckBox 50,25,100,14,"E-field",.CheckBox1
		CheckBox 50,46,90,14,"H-field",.CheckBox2
		CheckBox 160,25,130,14,"E-field (farfield)",.CheckBox3
		CheckBox 160,46,140,14,"H-field (farfield)",.CheckBox4


		GroupBox 30,91,280,413,"Coordinate System",.GroupBox2
		OptionGroup .Group1
			OptionButton 50,112,160,14,"Spherical Coordinate",.OptionButton1
			OptionButton 50,196,160,14,"Cartesian Coordinate",.OptionButton2
			OptionButton 50,357,160,14,"Cylindrical Coordinate",.OptionButton3
		Text 70,133,90,14,"Radius (mm)",.Text4
		Text 70,154,90,14,"Theta (delta)",.Text5
		Text 70,175,90,14,"Phi (delta)",.Text6
		TextBox 160,133,90,21,.radius
		TextBox 160,154,90,21,.theta
		TextBox 160,175,90,21,.phi
		DropListBox 70,217,60,170,PlaneNormal(),.Plane
		Text 70,245,70,14,"Xmin (mm)",.Text1
		Text 150,245,90,14,"Xmax (mm)",.Text2
		Text 230,245,60,14,"samples",.Text3
		TextBox 70,259,70,21,.xmin
		TextBox 150,259,70,21,.xmax
		TextBox 230,259,60,21,.xsample
		Text 70,280,90,14,"Ymin (mm)",.Text7
		Text 150,280,90,14,"Ymax (mm)",.Text8
		Text 230,280,60,14,"samples",.Text9
		TextBox 70,294,70,21,.ymin
		TextBox 150,294,70,21,.ymax
		TextBox 230,294,60,21,.ysample
		Text 70,315,90,14,"Zmin (mm)",.Text10
		Text 150,315,70,14,"Zmax (mm)",.Text11
		Text 230,315,70,14,"samples",.Text12
		TextBox 70,329,70,21,.zmin
		TextBox 150,329,70,21,.zmax
		TextBox 230,329,60,21,.zsample
		Text 140,217,90,14,"Plane",.Text13
		Text 70,392,90,14,"Zmin (mm)",.Text15
		Text 70,455,90,14,"Radius (mm)",.Text16
		Text 70,476,90,14,"Angle Res.",.Text17
		Text 70,434,90,14,"Step Sample",.Text18
		TextBox 160,392,90,21,.minDistance
		TextBox 160,413,90,21,.maxDistance
		TextBox 160,434,90,21,.SphSample
		TextBox 160,455,90,21,.radiusC
		TextBox 160,476,90,21,.angleRes
		Text 70,413,90,14,"Zmax (mm)",.Text19
		DropListBox 70,371,80,170,PlaneNormalC(),.PlaneC
		Text 160,371,70,14,"Axis",.Text14
	End Dialog
	Dim dlg As UserDialog

	dlg.radius = "3000"
	dlg.theta = "45"
	dlg.phi = "45"
	dlg.checkbox3 = 1
	dlg.xmin = "-100"
	dlg.xmax = "100"
	dlg.xsample = "10"
	dlg.ymin = "-100"
	dlg.ymax = "100"
	dlg.ysample = "10"
	dlg.zmin = "-100"
	dlg.zmax = "100"
	dlg.zsample = "10"
	dlg.Group1 = 2
	dlg.minDistance = "-500"
	dlg.maxDistance = "500"
	dlg.angleRes = "30"
	dlg.SphSample = "11"
	dlg.RadiusC = "500"
	dlg.PlaneC = 2



	If Dialog(dlg)=0 Then Exit All

	Dim Efar As Boolean
	Dim Enear As Boolean
	Dim Hfar As Boolean
	Dim Hnear As Boolean
	Dim CoordinateSys As Double



	k = CDbl(Units.GetGeometryUnitToSI)/0.001 'scale distance to mm
	Enear = CBool(dlg.checkbox1)
	Hnear = CBool(dlg.checkbox2)
	Efar = CBool(dlg.checkbox3)
	Hfar = CBool(dlg.checkbox4)
	CoordinateSys = dlg.Group1 '0 = Spherical, 1 = Cartesian, 2 = Spherical
	dtheta = CDbl(dlg.theta)
	dphi = CDbl(dlg.phi)
	radius = CDbl(dlg.radius)
	SphSample = CDbl(dlg.SphSample)
	angleRes = CDbl(dlg.angleRes)
	zStep = (CDbl(dlg.MaxDistance)-CDbl(dlg.MinDistance))/(SphSample-1)
	
	sHistEntry = "SetLock True" + vbLf + _
			"ScreenUpdating False" + vbLf + _
			"SetNoMouseSelection True" + vbLf + _
			"LockTree True" + vbLf + _
			"ResultTree.EnableTreeUpdate False" + vbLf

	If CoordinateSys = 0 Then
			sHistEntry 	=	sHistEntry + vbLf + _
							"Dim xx As Double" + vbLf + _
							"Dim yy As Double" + vbLf + _
							"Dim zz As Double" + vbLf + _
							"Dim radius As Double" + vbLf + _
							"Dim dtheta As Double" + vbLf + _
							"Dim dphi As Double" + vbLf + _
							"Dim n As Double" + vbLf + _
							"Dim m As Double" + vbLf + _
							"Dim IDD As Integer" + vbLf + _
							"radius = " + CStr(radius) + vbLf + _
							"dtheta = " + CStr(dtheta) + vbLf + _
							"dphi = " + CStr(dphi) + vbLf

			If Enear = True Then
				SphericalField1
				sHistEntry	= sHistEntry + vbLf + _
									"			.Reset" + vbLf + _
						     		"			.Field ""Efield""" + vbLf
				CartesianField2

			End If

			If Hnear = True Then
				SphericalField1
				sHistEntry =	sHistEntry + vbLf + _
									"			.Reset" + vbLf + _
					    			"			.Field ""Hfield""" + vbLf
				CartesianField2

			End If


			If Efar = True Then
				SphericalField1
				sHistEntry	= sHistEntry + vbLf + _
						     		"			.Reset" + vbLf + _
					    			"			.Field ""EFarfield""" + vbLf

					SphericalField2

			End If

			If Hfar = True Then
				SphericalField1
				sHistEntry =	sHistEntry + vbLf + _
						     		"			.Reset" + vbLf + _
									"			.Field ""HFarfield""" + vbLf

					SphericalField2

			End If

			sHeader = "(*) Define Field Probes in Spherical Coordinate"
	
	ElseIf CoordinateSys = 1 Then 'cartesian coordinate

			sHistEntry 	=	sHistEntry + vbLf + _
							"Dim xmin As Double" + vbLf + _
							"Dim ymin As Double" + vbLf + _
							"Dim zmin As Double" + vbLf + _
							"Dim xmax As Double" + vbLf + _
							"Dim ymax As Double" + vbLf + _
							"Dim zmax As Double" + vbLf + _
							"Dim stepX As Double" + vbLf + _
							"Dim stepY As Double" + vbLf + _
							"Dim stepZ As Double" + vbLf + _
							"Dim xsample As Double" + vbLf + _
							"Dim ysample As Double" + vbLf + _
							"Dim zsample As Double" + vbLf + _
							"Dim n As Double" + vbLf + _
							"Dim m As Double" + vbLf + _
							"Dim LocX As Double" + vbLf + _
							"Dim LocY As Double" + vbLf + _
							"Dim LocZ As Double" + vbLf + _
							"Dim IDD As Integer" + vbLf + _
							"xmin = " + CStr(dlg.xmin) + vbLf + _
							"xmax = " + CStr(dlg.xmax) + vbLf + _
							"xsample = " + CStr(dlg.xsample) + vbLf + _
							"ymin = " + CStr(dlg.ymin) + vbLf + _
							"ymax = " + CStr(dlg.ymax) + vbLf + _
							"ysample = " + CStr(dlg.ysample) + vbLf + _
							"zmin = " + CStr(dlg.zmin) + vbLf + _
							"zmax = " + CStr(dlg.zmax) + vbLf + _
							"zsample = " + CStr(dlg.zsample) + vbLf + _
							"If xsample=1 Then" + vbLf + _
							"	   stepX = 0" + vbLf + _
							"Else" + vbLf + _
							"	   stepX = (Xmax-Xmin)/(xsample-1)" + vbLf + _
							"End If" + vbLf + _
							"If ysample=1 Then" + vbLf + _
							"	   stepY = 0" + vbLf + _
							"Else" + vbLf + _
							"	   stepY = (Ymax-Ymin)/(ysample-1)" + vbLf + _
							"End If" + vbLf + _
							"If zsample=1 Then" + vbLf + _
							"	   stepZ = 0" + vbLf + _
							"Else" + vbLf + _
							"	   stepZ = (Zmax-Zmin)/(zsample-1)" + vbLf + _
							"End If" + vbLf + _
							"LocX = xmin" + vbLf + _
							"LocY = ymin" + vbLf + _
							"LocZ = zmin" + vbLf

			If Enear = True Then
				If dlg.Plane = 0 Then

					Cartesian_Xplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efield""" + vbLf
					Cartesian_Xplane2

				ElseIf dlg.Plane = 1 Then

					Cartesian_Yplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efield""" + vbLf
					Cartesian_Yplane2

				ElseIf dlg.Plane = 2 Then

					Cartesian_Zplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efield""" + vbLf
 					Cartesian_Zplane2

				End If
				sHistEntry =	sHistEntry + vbLf + _
     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
			    			"			.Create" + vbLf + _
     						"		End With" + vbLf + _
							"	Next m" + vbLf + _
							"Next n" + vbLf

			End If

			If Hnear = True Then
				If dlg.Plane = 0 Then

					Cartesian_Xplane1
					sHistEntry =	sHistEntry + vbLf + _
							"			.Reset" + vbLf + _
				  			"			.Field ""Hfield""" + vbLf
					Cartesian_Xplane2

				ElseIf dlg.Plane = 1 Then

					Cartesian_Yplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Hfield""" + vbLf
					Cartesian_Yplane2

				ElseIf dlg.Plane = 2 Then

					Cartesian_Zplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Hfield""" + vbLf
 					Cartesian_Zplane2

				End If
				sHistEntry =	sHistEntry + vbLf + _
     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
			    			"			.Create" + vbLf + _
     						"		End With" + vbLf + _
							"	Next m" + vbLf + _
							"Next n" + vbLf
			End If

			If Efar = True Then
				If dlg.Plane = 0 Then

					Cartesian_Xplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efarfield""" + vbLf
					Cartesian_Xplane2

				ElseIf dlg.Plane = 1 Then

					Cartesian_Yplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efarfield""" + vbLf
					Cartesian_Yplane2

				ElseIf dlg.Plane = 2 Then

					Cartesian_Zplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Efarfield""" + vbLf
 					Cartesian_Zplane2

				End If
				sHistEntry =	sHistEntry + vbLf + _
     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
			    			"			.Create" + vbLf + _
     						"		End With" + vbLf + _
							"	Next m" + vbLf + _
							"Next n" + vbLf
			End If

			If Hfar = True Then
				If dlg.Plane = 0 Then

					Cartesian_Xplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Hfarfield""" + vbLf
					Cartesian_Xplane2

				ElseIf dlg.Plane = 1 Then

					Cartesian_Yplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Hfarfield""" + vbLf
					Cartesian_Yplane2

				ElseIf dlg.Plane = 2 Then

					Cartesian_Zplane1
					sHistEntry =	sHistEntry + vbLf + _
     						"			.Reset" + vbLf + _
							"			.Field ""Hfarfield""" + vbLf
 					Cartesian_Zplane2

 				End If

				sHistEntry =	sHistEntry + vbLf + _
     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
			    			"			.Create" + vbLf + _
     						"		End With" + vbLf + _
							"	Next m" + vbLf + _
							"Next n" + vbLf
			End If
			
			sHeader =  "(*) Define Field Probes in Cartesian Coordinate"
			
	ElseIf CoordinateSys = 2 Then

			sHistEntry 	=	sHistEntry + vbLf + _
							"Dim xx As Double" + vbLf + _
							"Dim yy As Double" + vbLf + _
							"Dim zz As Double" + vbLf + _
							"Dim radius As Double" + vbLf + _
							"Dim SphSample As Double" + vbLf + _
							"Dim angleRes As double" + vbLf + _
							"Dim MinDist As Double" + vbLf + _
							"Dim MaxDist As Double" + vbLf + _
							"Dim zStep As Double" + vbLf + _
							"Dim n As Double" + vbLf + _
							"Dim m As Double" + vbLf + _
							"Dim LocX As Double" + vbLf + _
							"Dim LocY As Double" + vbLf + _
							"Dim LocZ As Double" + vbLf + _
							"Dim IDD As Integer" + vbLf + _
							"Dim LocR As Double" + vbLf + _
							"Dim LocT As Double" + vbLf + _
							"Dim LocP As Double" + vbLf + _
							"radius = " + CStr(dlg.radiusC) + vbLf + _
							"SphSample = " + CStr(dlg.SphSample) + vbLf + _
							"angleRes = " + CStr(dlg.angleRes) + vbLf + _
							"MinDist = " + CStr(dlg.MinDistance) + vbLf + _
							"MaxDist = " + CStr(dlg.MaxDistance) + vbLf + _
							"zStep = " + CStr(zStep) + vbLf

			'dlg.PlaneC ==> Axis Selection

			If Enear = True Then
				If dlg.PlaneC = 0 Then
					CylindricalFieldX
				ElseIf dlg.PlaneC = 1 Then
					CylindricalFieldY
				ElseIf dlg.PlaneC = 2 Then
					CylindricalFieldZ
				End If
					sHistEntry	= sHistEntry + vbLf + _
										"		With Probe" + vbLf + _
							     		"			.Reset" + vbLf + _
						    			"			.Field ""Efield""" + vbLf

					CylindricalField

			End If

			If Hnear = True Then
				If dlg.PlaneC = 0 Then
					CylindricalFieldX
				ElseIf dlg.PlaneC = 1 Then
					CylindricalFieldY
				ElseIf dlg.PlaneC = 2 Then
					CylindricalFieldZ
				End If
					sHistEntry =	sHistEntry + vbLf + _
										"		With Probe" + vbLf + _
							     		"			.Reset" + vbLf + _
						    			"			.Field ""Hfield""" + vbLf
					CylindricalField

			End If

			If Efar = True Then
				If dlg.PlaneC = 0 Then
					CylindricalFieldX
				ElseIf dlg.PlaneC = 1 Then
					CylindricalFieldY
				ElseIf dlg.PlaneC = 2 Then
					CylindricalFieldZ
				End If

					CylindricalEFieldbisphiz

			End If

			If Hfar = True Then
				If dlg.PlaneC = 0 Then
					CylindricalFieldX
				ElseIf dlg.PlaneC = 1 Then
					CylindricalFieldY
				ElseIf dlg.PlaneC = 2 Then
					CylindricalFieldZ
				End If

					CylindricalHFieldbisphiz

			End If

			sHeader = "(*) Define Field Probes in Cylindrical Coordinate"
	End If

	sHistEntry = sHistEntry + vbLf + _
	"ResultTree.EnableTreeUpdate True" + vbLf + _
	"LockTree False" + vbLf + _
	"SetNoMouseSelection False" + vbLf + _
	"ScreenUpdating True" + vbLf + _
	"SetLock False" + vbLf
	
	addToHistory(sHeader, sHistEntry)

End Sub

Public Sub CylindricalFieldX()
					sHistEntry =	sHistEntry + vbLf + _
									"For n = 0 To SphSample-1 " + vbLf + _
									"	For m = 0 To 360-" + CStr(angleRes) + " STEP " + CStr(angleRes) + vbLf + _
									"        	If InStr(CStr(radius*sind(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 180 Then" + vbLf + _
									"					LocY = 0" + vbLf + _
									"				Else" + vbLf + _
		                            "					LocY = CDbl(Left(CStr(radius*sind(m)), InStr(CStr(radius*sind(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									" 				LocY = radius*sind(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(radius*cosd(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 90 Or m = 270 Then" + vbLf + _
									"					LocZ = 0" + vbLf + _
									"				else" + vbLf + _
		                            "					LocZ = CDbl(Left(CStr(radius*cosd(m)), InStr(CStr(radius*cosd(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"				LocZ = radius*cosd(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(minDist + n*zStep)," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocX = CDbl(Left(CStr(minDist + n*zStep), InStr(CStr(minDist + n*zStep), " + """.""" + ") + 2))" + vbLf + _
									"			else" + vbLf + _
									"				LocX = minDist + n*zStep" + vbLf + _
		             				"			End If" + vbLf + _
									"			If Probe.GetFirst = 0 Then" + vbLf + _
									"				IDD = 0" + vbLf + _
									"			Else" + vbLf + _
									"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
									"			End If" + vbLf + _
									"        	If LocX=0 and LocY>0 Then" + vbLf + _
									"				LocP=90" + vbLf + _
									"				ElseIf LocX=0 and LocY<0 then" + vbLf + _
									"				LocP=270" + vbLf + _
									"				ElseIf LocY=0 And LocX<0 Then" + vbLf + _
									"				LocP = 180" + vbLf + _
									"				ElseIf LocY=0 And LocX=0 Then" + vbLf + _
									"				LocP = 90" + vbLf + _
									"				ElseIf LocY<0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"				ElseIf LocY>0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"			Else" + vbLf + _
									"        		If InStr(CStr(Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "					LocP = CDbl(Left(CStr(Atnd(LocY/LocX)), InStr(CStr(Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"				Else" + vbLf + _
									"					LocP = Atnd(LocY/LocX)" + vbLf + _
									"				End If" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocR = CDbl(Left(CStr(sqr(LocX^2+LocY^2+LocZ^2)), InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocR=sqr(LocX^2+LocY^2+LocZ^2)" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(Acosd(LocZ/LocR))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocT = CDbl(Left(CStr(Acosd(LocZ/LocR)), InStr(CStr(Acosd(LocZ/LocR)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocT=Acosd(LocZ/LocR)" + vbLf + _
									"			End If" + vbLf



End Sub

Public Sub CylindricalFieldY()
					sHistEntry =	sHistEntry + vbLf + _
									"For n = 0 To SphSample-1 " + vbLf + _
									"	For m = 0 To 360-" + CStr(angleRes) + " STEP " + CStr(angleRes) + vbLf + _
									"        	If InStr(CStr(radius*sind(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 180 Then" + vbLf + _
									"					LocZ = 0" + vbLf + _
									"				Else" + vbLf + _
		                            "					LocZ = CDbl(Left(CStr(radius*sind(m)), InStr(CStr(radius*sind(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									" 				LocZ = radius*sind(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(radius*cosd(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 90 Or m = 270 Then" + vbLf + _
									"					LocX = 0" + vbLf + _
									"				else" + vbLf + _
		                            "					LocX = CDbl(Left(CStr(radius*cosd(m)), InStr(CStr(radius*cosd(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"				LocX = radius*cosd(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(minDist + n*zStep)," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocY = CDbl(Left(CStr(minDist + n*zStep), InStr(CStr(minDist + n*zStep), " + """.""" + ") + 2))" + vbLf + _
									"			else" + vbLf + _
									"				LocY = minDist + n*zStep" + vbLf + _
		             				"			End If" + vbLf + _
									"			If Probe.GetFirst = 0 Then" + vbLf + _
									"				IDD = 0" + vbLf + _
									"			Else" + vbLf + _
									"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
									"			End If" + vbLf + _
									"        	If LocX=0 and LocY>0 Then" + vbLf + _
									"				LocP=90" + vbLf + _
									"				ElseIf LocX=0 and LocY<0 then" + vbLf + _
									"				LocP=270" + vbLf + _
									"				ElseIf LocY=0 And LocX<0 Then" + vbLf + _
									"				LocP = 180" + vbLf + _
									"				ElseIf LocY=0 And LocX=0 Then" + vbLf + _
									"				LocP = 90" + vbLf + _
									"				ElseIf LocY<0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"				ElseIf LocY>0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"			Else" + vbLf + _
									"        		If InStr(CStr(Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "					LocP = CDbl(Left(CStr(Atnd(LocY/LocX)), InStr(CStr(Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"				Else" + vbLf + _
									"					LocP = Atnd(LocY/LocX)" + vbLf + _
									"				End If" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocR = CDbl(Left(CStr(sqr(LocX^2+LocY^2+LocZ^2)), InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocR=sqr(LocX^2+LocY^2+LocZ^2)" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(Acosd(LocZ/LocR))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocT = CDbl(Left(CStr(Acosd(LocZ/LocR)), InStr(CStr(Acosd(LocZ/LocR)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocT=Acosd(LocZ/LocR)" + vbLf + _
									"			End If" + vbLf
End Sub

Public Sub CylindricalFieldZ()
					sHistEntry =	sHistEntry + vbLf + _
									"For n = 0 To SphSample-1 " + vbLf + _
									"	For m = 0 To 360-" + CStr(angleRes) + " STEP " + CStr(angleRes) + vbLf + _
									"        	If InStr(CStr(radius*sind(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 180 Then" + vbLf + _
									"					LocY = 0" + vbLf + _
									"				Else" + vbLf + _
		                            "					LocY = CDbl(Left(CStr(radius*sind(m)), InStr(CStr(radius*sind(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									" 				LocY = radius*sind(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(radius*cosd(m))," + """.""" + ") <> 0 Then" + vbLf + _
									"				If m = 90 Or m = 270 Then" + vbLf + _
									"					LocX = 0" + vbLf + _
									"				else" + vbLf + _
		                            "					LocX = CDbl(Left(CStr(radius*cosd(m)), InStr(CStr(radius*cosd(m)), " + """.""" + ") + 2))" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"				LocX = radius*cosd(m)" + vbLf + _
		             				"			End If" + vbLf + _
								    "        	If InStr(CStr(minDist + n*zStep)," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocZ = CDbl(Left(CStr(minDist + n*zStep), InStr(CStr(minDist + n*zStep), " + """.""" + ") + 2))" + vbLf + _
									"			else" + vbLf + _
									"				LocZ = minDist + n*zStep" + vbLf + _
		             				"			End If" + vbLf + _
									"			If Probe.GetFirst = 0 Then" + vbLf + _
									"				IDD = 0" + vbLf + _
									"			Else" + vbLf + _
									"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
									"			End If" + vbLf + _
									"        	If LocX=0 and LocY>0 Then" + vbLf + _
									"				LocP=90" + vbLf + _
									"				ElseIf LocX=0 and LocY<0 then" + vbLf + _
									"				LocP=270" + vbLf + _
									"				ElseIf LocY=0 And LocX<0 Then" + vbLf + _
									"				LocP = 180" + vbLf + _
									"				ElseIf LocY=0 And LocX=0 Then" + vbLf + _
									"				LocP = 90" + vbLf + _
									"				ElseIf LocY<0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"				ElseIf LocY>0 And LocX<0 Then" + vbLf + _
									"        			If InStr(CStr(180+Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "						LocP = CDbl(Left(CStr(180+Atnd(LocY/LocX)), InStr(CStr(180+Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"					Else" + vbLf + _
									"						LocP = 180+Atnd(LocY/LocX)" + vbLf + _
									"					End If" + vbLf + _
									"			Else" + vbLf + _
									"        		If InStr(CStr(Atnd(LocY/LocX))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "					LocP = CDbl(Left(CStr(Atnd(LocY/LocX)), InStr(CStr(Atnd(LocY/LocX)), " + """.""" + ") + 2))" + vbLf + _
									"				Else" + vbLf + _
									"					LocP = Atnd(LocY/LocX)" + vbLf + _
									"				End If" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocR = CDbl(Left(CStr(sqr(LocX^2+LocY^2+LocZ^2)), InStr(CStr(sqr(LocX^2+LocY^2+LocZ^2)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocR=sqr(LocX^2+LocY^2+LocZ^2)" + vbLf + _
									"			End If" + vbLf + _
									"        	If InStr(CStr(Acosd(LocZ/LocR))," + """.""" + ") <> 0 Then" + vbLf + _
		                            "				LocT = CDbl(Left(CStr(Acosd(LocZ/LocR)), InStr(CStr(Acosd(LocZ/LocR)), " + """.""" + ") + 2))" + vbLf + _
									"			Else" + vbLf + _
									"				LocT=Acosd(LocZ/LocR)" + vbLf + _
									"			End If" + vbLf
					End Sub

Public Sub CylindricalField()
						sHistEntry =	sHistEntry + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.AutoLabel 1" + vbLf + _
									"			.Orientation ""All""" + vbLf + _
		     						"			.SetPosition1 LocX/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocY/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocZ/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf


End Sub


Public Sub CylindricalFieldbis()
						sHistEntry =	sHistEntry + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.AutoLabel 1" + vbLf + _
									"			.Orientation ""All""" + vbLf + _
		     						"			.SetPosition1 LocT/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocP/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocR/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Spherical""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf

End Sub

Public Sub CylindricalEFieldbisphiz()
					sHistEntry	= sHistEntry + vbLf + _
								    "		With Probe" + vbLf + _
							     	"			.Reset" + vbLf + _
						    		"			.Field ""EFarfield""" + vbLf + _
									"			.caption ""E_Field (Farfield) ( "" + cstr(cint(LocT)) + "" "" + Cstr(cint(LocP)) + "") (Hor)"" + "" ("" + cstr(radius)+ "")""" + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.Orientation ""phi""" + vbLf + _
		     						"			.SetPosition1 LocT/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocP/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocR/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Spherical""" + vbLf + _
					    			"			.Create" + vbLf + _
						            "			.Origin ""zero""" + vbLf + _
		     						"		End With" + vbLf + _
									"			IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
								    "		With Probe" + vbLf + _
							     	"			.Reset" + vbLf + _
						    		"			.Field ""EFarfield""" + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.caption ""E_Field (Farfield) ( "" + cstr(cint(LocT)) + "" "" + Cstr(cint(LocP)) + "") (Ver)"" + "" ("" + cstr(radius)+ "")""" + vbLf + _
									"			.Orientation ""z""" + vbLf + _
		     						"			.SetPosition1 LocX/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocY/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocZ/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf

End Sub
Public Sub CylindricalHFieldbisphiz()
					sHistEntry	= sHistEntry + vbLf + _
								    "		With Probe" + vbLf + _
							     	"			.Reset" + vbLf + _
						    		"			.Field ""HFarfield""" + vbLf + _
									"			.caption ""H_Field (Farfield) ( "" + cstr(cint(LocT)) + "" "" + Cstr(cint(LocP)) + "") (Hor)"" + "" ("" + cstr(radius)+ "")""" + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.Orientation ""phi""" + vbLf + _
		     						"			.SetPosition1 LocT/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocP/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocR/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Spherical""" + vbLf + _
					    			"			.Create" + vbLf + _
						            "			.Origin ""zero""" + vbLf + _
		     						"		End With" + vbLf + _
									"			IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
								    "		With Probe" + vbLf + _
							     	"			.Reset" + vbLf + _
						    		"			.Field ""HFarfield""" + vbLf + _
									"			.ID IDD" + vbLf + _
									"			.caption ""H_Field (Farfield) ( "" + cstr(cint(LocT)) + "" "" + Cstr(cint(LocP)) + "") (Ver)"" + "" ("" + cstr(radius)+ "")""" + vbLf + _
									"			.Orientation ""z""" + vbLf + _
		     						"			.SetPosition1 LocX/" + CStr(k) + vbLf + _
									"			.SetPosition2 LocY/" + CStr(k) + vbLf + _
									"			.SetPosition3 LocZ/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf

End Sub



Public Sub SphericalField1()
					sHistEntry =	sHistEntry + vbLf + _
									"For n = 0 To 180 STEP " + Cstr(dtheta) + vbLf + _
									"	For m = 0 To 360-" + CStr(dphi) + " STEP " + CStr(dphi) + vbLf + _
									"			If Probe.GetFirst = 0 Then" + vbLf + _
									"				IDD = 0" + vbLf + _
									"			Else" + vbLf + _
									"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
									"			End If" + vbLf + _
									"		With Probe" + vbLf

End Sub

Public Sub SphericalField2()
				sHistEntry	= sHistEntry + vbLf + _
		     						"			.ID IDD" + vbLf + _
									"			.AutoLabel 1" + vbLf + _
									"			.Orientation ""All""" + vbLf + _
		     						"			.SetPosition1 n" + vbLf + _
									"			.SetPosition2 m" + vbLf + _
									"			.SetPosition3 radius/" + CStr(k) + vbLf + _
		     						"			.SetCoordinateSystemType ""Spherical""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"       If n = 0 or n = 180 Then" + vbLf + _
                                    "     		Exit For" + vbLf + _
                               		"		End If" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf
End Sub

Public Sub CartesianField2()
				sHistEntry	= sHistEntry + vbLf + _
		     						"			.ID IDD" + vbLf + _
									"			.AutoLabel 1" + vbLf + _
									"			.Orientation ""All""" + vbLf + _
									"			If InStr(CStr(radius*sind(n)*cosd(m)),"+"""."""+")<>0 Then" + vbLf + _
									"				If n=180 Or m=90 Or m=270 Then" + vbLf + _
									"					.SetPosition1 0" + vbLf + _
									"				else" + vbLf + _
									"					.SetPosition1 Left(CStr(radius*sind(n)*cosd(m)), InStr(CStr(radius*sind(n)*cosd(m)), " + """.""" + ") + 2)" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"					.SetPosition1 radius*sind(n)*cosd(m)/" + CStr(k) + vbLf + _
									"			End If" + vbLf + _
									"			If InStr(CStr(radius*sind(n)*sind(m)),"+"""."""+")<>0 Then" + vbLf + _
									"				If n=180 Or m=180 Then" + vbLf + _
									"					.SetPosition2 0" + vbLf + _
									"				else" + vbLf + _
									"					.SetPosition2 Left(CStr(radius*sind(n)*sind(m)), InStr(CStr(radius*sind(n)*sind(m)), " + """.""" + ") + 2)" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"					.SetPosition2 radius*sind(n)*sind(m)/" + CStr(k) + vbLf + _
									"			End If" + vbLf + _
									"			If InStr(CStr(radius*cosd(n)),"+"""."""+")<>0 Then" + vbLf + _
									"				If n=90 Or n=270 Then" + vbLf + _
									"					.SetPosition3 0" + vbLf + _
									"				else" + vbLf + _
									"					.SetPosition3 Left(CStr(radius*cosd(n)), InStr(CStr(radius*cosd(n)), " + """.""" + ") + 2)" + vbLf + _
									"				End If" + vbLf + _
									"			else" + vbLf + _
									"				.SetPosition3 radius*cosd(n)/" + CStr(k) + vbLf + _
									"			End If" + vbLf + _
		     						"			.SetCoordinateSystemType ""Cartesian""" + vbLf + _
					    			"			.Create" + vbLf + _
		     						"		End With" + vbLf + _
									"       If n = 0 or n = 180 Then" + vbLf + _
                                    "     		Exit For" + vbLf + _
                               		"		End If" + vbLf + _
									"	Next m" + vbLf + _
									"Next n" + vbLf


End Sub




Public Sub Cartesian_Xplane1()
						sHistEntry =	sHistEntry + vbLf + _
							"For n = 0 To ysample-1" + vbLf + _
							"	For m = 0 To zsample-1" + vbLf + _
						    "        	If InStr(CStr(m*stepZ+zmin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocZ = CDbl(Left(CStr(m*stepZ+zmin), InStr(CStr(m*stepZ+zmin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							" 				LocZ = m*stepZ + zmin" + vbLf + _
             				"			End If" + vbLf + _
						    "        	If InStr(CStr(n*stepY+ymin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocY = CDbl(Left(CStr(n*stepY+ymin), InStr(CStr(n*stepY+ymin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							"				LocY = n*stepY + ymin" + vbLf + _
             				"			End If" + vbLf + _
							"			If Probe.GetFirst = 0 Then" + vbLf + _
							"				IDD = 0" + vbLf + _
							"			Else" + vbLf + _
							"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
							"			End If" + vbLf + _
							"		With Probe" + vbLf
End Sub

Public Sub Cartesian_Yplane1()
						sHistEntry =	sHistEntry + vbLf + _
							"For n = 0 To xsample-1" + vbLf + _
							"	For m = 0 To zsample-1" + vbLf + _
						    "        	If InStr(CStr(m*stepZ+zmin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocZ = CDbl(Left(CStr(m*stepZ+zmin), InStr(CStr(m*stepZ+zmin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							"				LocZ = m*stepZ+zmin" + vbLf + _
             				"			End If" + vbLf + _
						    "        	If InStr(CStr(n*stepX+xmin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocX = CDbl(Left(CStr(n*stepX+xmin), InStr(CStr(n*stepX+xmin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							"				LocX = n*stepX+xmin" + vbLf + _
             				"			End If" + vbLf + _
							"			If Probe.GetFirst = 0 Then" + vbLf + _
							"				IDD = 0" + vbLf + _
							"			Else" + vbLf + _
							"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
							"			End If" + vbLf + _
							"		With Probe" + vbLf
End Sub

Public Sub Cartesian_Zplane1()
						sHistEntry =	sHistEntry + vbLf + _
							"For n = 0 To xsample-1" + vbLf + _
							"	For m = 0 To ysample-1" + vbLf + _
						    "        	If InStr(CStr(m*StepY+ymin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocY = CDbl(Left(CStr(m*StepY+ymin), InStr(CStr(m*StepY+ymin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							" 				LocY = m*StepY+ymin" + vbLf + _
             				"			End If" + vbLf + _
						    "        	If InStr(CStr(n*StepX+xmin)," + """.""" + ") <> 0 Then" + vbLf + _
                            "				LocX = CDbl(Left(CStr(n*StepX+xmin), InStr(CStr(n*StepX+xmin), " + """.""" + ") + 2))" + vbLf + _
							"			else" + vbLf + _
							"				LocX = n*StepX+xmin" + vbLf + _
             				"			End If" + vbLf + _
							"			If Probe.GetFirst = 0 Then" + vbLf + _
							"				IDD = 0" + vbLf + _
							"			Else" + vbLf + _
							"				IDD = 1 + CInt(Probe.GetLastAddedID)" + vbLf + _
							"			End If" + vbLf + _
							"		With Probe" + vbLf
End Sub

Public Sub Cartesian_Xplane2()
						sHistEntry =	sHistEntry + vbLf + _
	     					"			.Orientation ""All""" + vbLf + _
     						"			.SetPosition1 xmin/" + CStr(k) + vbLf + _
							"			.SetPosition2 LocY/" + CStr(k) + vbLf + _
							"			.SetPosition3 LocZ/" + CStr(k) + vbLf
End Sub

Public Sub Cartesian_Yplane2()
						sHistEntry =	sHistEntry + vbLf + _
	     					"			.Orientation ""All""" + vbLf + _
     						"			.SetPosition1 LocX/" + CStr(k) + vbLf + _
							"			.SetPosition2 ymin/" + CStr(k) + vbLf + _
							"			.SetPosition3 LocZ/" + CStr(k) + vbLf
End Sub

Public Sub Cartesian_Zplane2()
						sHistEntry =	sHistEntry + vbLf + _
	     					"			.Orientation ""All""" + vbLf + _
     						"			.SetPosition1 LocX/" + CStr(k) + vbLf + _
							"			.SetPosition2 LocY/" + CStr(k) + vbLf + _
							"			.SetPosition3 zmin/" + CStr(k) + vbLf
End Sub




Rem See DialogFunc help topic for more information.

Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable ("Text1", False)
		DlgEnable ("Text2", False)
		DlgEnable ("Text3", False)
		DlgEnable ("Text4", False)
		DlgEnable ("Text5", False)
		DlgEnable ("Text6", False)
		DlgEnable ("radius", False)
		DlgEnable ("theta", False)
		DlgEnable ("phi", False)
		DlgEnable ("Text7", False)
		DlgEnable ("Text8", False)
		DlgEnable ("Text9", False)
		DlgEnable ("Text10", False)
		DlgEnable ("Text11", False)
		DlgEnable ("Text12", False)
		DlgEnable ("Text13", False)
		DlgEnable("Plane",False)
		DlgEnable("xmin",False)
		DlgEnable("xmax",False)
		DlgEnable("xsample",False)
		DlgEnable("ymin", False)
		DlgEnable("ymax",False)
		DlgEnable("ysample",False)
		DlgEnable("zmin",False)
		DlgEnable("zmax",False)
		DlgEnable("zsample",False)
		DlgEnable("Text14",False)
		DlgEnable("Text15",False)
		DlgEnable("Text16",False)
		DlgEnable("Text17",False)
		DlgEnable("Text18",False)
		DlgEnable("Text19", False)
		DlgEnable("radiusC",True)
		DlgEnable("minDistance",True)
		DlgEnable ("maxDistance", True)
		DlgEnable("angleRes",True)
		DlgEnable("SphSample",True)
		DlgEnable("PlaneC", True)

	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem

		Case "Group1"

			If DlgValue("Group1")=0 Then
				DlgEnable ("Text1", False)
				DlgEnable ("Text2", False)
				DlgEnable ("Text3", False)
				DlgEnable ("Text4", True)
				DlgEnable ("Text5", True)
				DlgEnable ("Text6", True)
				DlgEnable("radius", True)
				DlgEnable("theta", True)
				DlgEnable("phi", True)
				DlgEnable ("Text7", False)
				DlgEnable ("Text8", False)
				DlgEnable ("Text9", False)
				DlgEnable ("Text10", False)
				DlgEnable ("Text11", False)
				DlgEnable ("Text12", False)
				DlgEnable ("Text13", False)
				DlgEnable("Plane",False)
				DlgEnable("xmin",False)
				DlgEnable("xmax",False)
				DlgEnable("xsample",False)
				DlgEnable("ymin", False)
				DlgEnable("ymax",False)
				DlgEnable("ysample",False)
				DlgEnable("zmin",False)
				DlgEnable("zmax",False)
				DlgEnable("zsample",False)
				DlgEnable("Text14",False)
				DlgEnable("Text15",False)
				DlgEnable("Text16",False)
				DlgEnable("Text17",False)
				DlgEnable("Text18",False)
				DlgEnable("Text19", False)
				DlgEnable("radiusC",False)
				DlgEnable("minDistance",False)
				DlgEnable ("maxDistance", False)
				DlgEnable("angleRes",False)
				DlgEnable("SphSample",False)
				DlgEnable("PlaneC", False)

			ElseIf DlgValue("Group1")=1 Then
				DlgEnable ("Text1", True)
				DlgEnable ("Text2", True)
				DlgEnable ("Text3", True)
				DlgEnable ("Text4", False)
				DlgEnable ("Text5", False)
				DlgEnable ("Text6", False)
				DlgEnable ("Text7", True)
				DlgEnable ("Text8", True)
				DlgEnable ("Text9", True)
				DlgEnable ("Text10", True)
				DlgEnable ("Text11", True)
				DlgEnable ("Text12", True)
				DlgEnable ("Text13", True)
				DlgEnable("xmax",False)
				DlgEnable("xsample",False)
				DlgEnable("Plane",True)
				DlgEnable("xmin",True)
				DlgEnable("ymin", True)
				DlgEnable("ymax",True)
				DlgEnable("ysample",True)
				DlgEnable("zmin",True)
				DlgEnable("zmax",True)
				DlgEnable("zsample",True)
				DlgEnable("radius",False)
				DlgEnable("theta",False)
				DlgEnable("phi",False)
				DlgEnable("Text14",False)
				DlgEnable("Text15",False)
				DlgEnable("Text16",False)
				DlgEnable("Text17",False)
				DlgEnable("Text18",False)
				DlgEnable("Text19", False)
				DlgEnable("radiusC",False)
				DlgEnable("minDistance",False)
				DlgEnable ("maxDistance", False)
				DlgEnable("angleRes",False)
				DlgEnable("SphSample",False)
				DlgEnable("PlaneC", False)
			ElseIf DlgValue("Group1")=2 Then
				DlgEnable ("Text1", False)
				DlgEnable ("Text2", False)
				DlgEnable ("Text3", False)
				DlgEnable ("Text4", False)
				DlgEnable ("Text5", False)
				DlgEnable ("Text6", False)
				DlgEnable ("Text7", False)
				DlgEnable ("Text8", False)
				DlgEnable ("Text9", False)
				DlgEnable ("Text10", False)
				DlgEnable ("Text11", False)
				DlgEnable ("Text12", False)
				DlgEnable ("Text13", False)
				DlgEnable("Plane",False)
				DlgEnable("xmin",False)
				DlgEnable("xmax",False)
				DlgEnable("xsample",False)
				DlgEnable("ymin", False)
				DlgEnable("ymax",False)
				DlgEnable("ysample",False)
				DlgEnable("zmin",False)
				DlgEnable("zmax",False)
				DlgEnable("zsample",False)
				DlgEnable("radius",False)
				DlgEnable("theta",False)
				DlgEnable("phi",False)
				DlgEnable("Text14",True)
				DlgEnable("Text15",True)
				DlgEnable("Text16",True)
				DlgEnable("Text17",True)
				DlgEnable("Text18",True)
				DlgEnable("Text19", True)
				DlgEnable("radiusC",True)
				DlgEnable("minDistance",True)
				DlgEnable ("maxDistance", True)
				DlgEnable("angleRes",True)
				DlgEnable("SphSample",True)
				DlgEnable("PlaneC", True)

			End If

		Case "Plane"
			If DlgValue("Plane")=0 Then
				DlgEnable("xmax",False)
				DlgEnable("xsample",False)
				DlgEnable("ymax",True)
				DlgEnable("zmax",True)
				DlgEnable("ysample",True)
				DlgEnable("zsample",True)
			ElseIf DlgValue("Plane")=1 Then
				DlgEnable("ymax",False)
				DlgEnable("ysample",False)
				DlgEnable("xmax",True)
				DlgEnable("zmax",True)
				DlgEnable("xsample",True)
				DlgEnable("zsample",True)
			ElseIf DlgValue("Plane")=2 Then
				DlgEnable("zmax",False)
				DlgEnable("zsample",False)
				DlgEnable("xmax",True)
				DlgEnable("ymax",True)
				DlgEnable("ysample",True)
				DlgEnable("xsample",True)
			End If

		Case "PlaneC"
			If DlgValue("PlaneC")=0 Then
				DlgText ("Text15", "Xmin(mm)")
				DlgText ("Text19", "Xmax(mm)")
			ElseIf DlgValue("PlaneC")=1 Then
				DlgText ("Text15", "Ymin(mm)")
				DlgText ("Text19", "Ymax(mm)")
			ElseIf DlgValue("PlaneC")=2 Then
				DlgText ("Text15", "Zmin(mm)")
				DlgText ("text19", "Zmax(mm)")
			End If

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
