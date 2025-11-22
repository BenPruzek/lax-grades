'-------------------------------------------------
' *Graphics / 3D Plot Grid on selected major plane
' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 13-Aug-2014 ube: First version
' ================================================================================================
'
'
' Plot grids on selected 3D field plot
'
' ube: 23-Dec-2014  Help button + online help
' yta: 07-Aug-2014  Default step width changed to 1/10th of bounding box
' yta: 24-Jul-2014  Added auto-detect cutplane
' yta: 14-Jul-2014  Added step width input option and bounding box size
' yta: 10-Jul-2014  Initial version
'
' select a 3d-fieldplot first, then execute macro
'--------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_Results_DrawGrid_2DPlot"

'Global variable

Global boxsize As Double, aAxisT As String
Dim x1 As Double, y1 As Double, z1 As Double, mmax As Double
Dim x2 As Double, y2 As Double, z2 As Double, mmin As Double



Sub Main


boxsize = 0.2 'size of the box signifying the origin as a fraction of the resolution
VectorPlot3D.Type "WithPicks"
Plot.update



    mmax = GetFieldPlotMaximumPos ( x1,  y1,  z1)
    mmin = GetFieldPlotMinimumPos ( x2, y2, z2 )

    If (x1 = x2) And (y1 <> y2) And (z1 <> z2) Then
    	aAxisT = "X"
    ElseIf (y1 = y2) And (z1 <> z2) Then
    	aAxisT = "Y"
    ElseIf z1 = z2 Then
    	aAxisT = "Z"
    Else
    	MsgBox("Please select X, Y or Z as cutplane normal")
    	Exit All
    End If



	Begin Dialog UserDialog 290,245,"Draw Grid on Field Plot",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,270,231,"Grid Settings",.SettingsGB
		Text 30,105,140,14,"Cut-Plane Normal Axis",.AxisT
		TextBox 190,98,50,21,.aAxisT
		Text 30,154,140,14,"Horizontal step width",.Text1
		Text 30,182,140,14,"Vertical step width",.Text2
		TextBox 200,147,60,21,.hresT
		TextBox 200,175,60,21,.vresT
		OKButton 20,210,80,21
		PushButton 105,210,80,21,"Exit",.ExitPB
		OptionGroup .DStepOrLine
			OptionButton 40,126,90,14,"Delta step",.Deltastep
			OptionButton 170,126,80,14,"Lines",.Numoflines
		Text 30,35,30,14,"xmin",.Text3
		Text 30,56,30,14,"ymin",.Text5
		Text 30,77,30,14,"zmin",.Text7
		Text 150,35,40,14,"xmax",.Text4
		Text 150,56,40,14,"ymax",.Text6
		Text 150,77,40,14,"zmax",.Text8
		TextBox 60,28,80,21,.xminT
		TextBox 60,49,80,21,.yminT
		TextBox 60,70,80,21,.zminT
		TextBox 190,28,80,21,.xmaxT
		TextBox 190,49,80,21,.ymaxT
		TextBox 190,70,80,21,.zmaxT
		Text 170,154,30,14,"d1",.hT
		Text 170,182,30,14,"d2",.vT
		PushButton 190,210,80,21,"Help",.Help

	End Dialog

    Dim dlg As UserDialog
    Dialog dlg

End Sub

Private Function Dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
    Dim hres As Double, vres As Double

	Plot3DPlotsOn2DPlane True
	WCS.AlignWCSWithGlobalCoordinates
    Boundary.GetCalculationBox (xmin, xmax, ymin, ymax, zmin, zmax)

    Select Case Action%
    	Case 1 'Dialog box initialization
			DlgText("hresT", "10")
			DlgText("vresT", "10")
        	DlgText("xminT", Format(xmin, "Scientific"))
        	DlgText("xmaxT", Format(xmax, "Scientific"))
        	DlgText("yminT", Format(ymin, "Scientific"))
        	DlgText("ymaxT", Format(ymax, "Scientific"))
        	DlgText("zminT", Format(zmin, "Scientific"))
        	DlgText("zmaxT", Format(zmax, "Scientific"))
        	DlgEnable("xminT", False)
        	DlgEnable("xmaxT", False)
        	DlgEnable("yminT", False)
        	DlgEnable("ymaxT", False)
        	DlgEnable("zminT", False)
        	DlgEnable("zmaxT", False)
        	DlgText("aAxisT", "  "+aAxisT)
        	DlgEnable("aAxisT", False)
        	If aAxisT = "X" Then
    			DlgText("hT", ": dy")
				DlgText("vT", ": dz")
				DlgText("hresT",cstr((ymax-ymin)/10))		'horizontal resolution is 1/10th of y width
				DlgText("vresT",cstr((zmax-zmin)/10))		'vertical resolution is 1/10th of z width
			ElseIf aAxisT = "Y" Then
				DlgText("hT", ": dz")
				DlgText("vT", ": dx")
				DlgText("hresT",cstr((zmax-zmin)/10))
				DlgText("vresT",cstr((xmax-xmin)/10))
			ElseIf aAxisT = "Z" Then
				DlgText("hT", ": dx")
				DlgText("vT", ": dy")
				DlgText("hresT",cstr((xmax-xmin)/10))
				DlgText("vresT",cstr((ymax-ymin)/10))
			End If

    	Case 2 'value changed or button pressed
        	Select Case DlgItem$
        		Case "ExitPB"
               		Exit All

                Case "Help"
					StartHelp HelpFileName
                    Dialogfunc = True

        		Case "DStepOrLine"
        			If DlgValue("DStepOrLine") = 0 Then
                		DlgText("Text1","Horizontal Step width")
                		DlgText("text2", "Vertizal Step width")
                		DlgVisible("hT", True)
                		DlgVisible("vT", True)
                		If  aAxisT = "X" Then
						    DlgText("hresT", cstr((ymax-ymin)/10))
						    DlgText("vresT", cstr((zmax-zmin)/10))
						ElseIf aAxisT = "Y" Then
							DlgText("hresT",cstr((zmax-zmin)/10))
						    DlgText("vresT",cstr((xmax-xmin)/10))
						ElseIf aAxisT = "Z" Then
							DlgText("hresT",cstr((xmax-xmin)/10))
						    DlgText("vresT",cstr((ymax-ymin)/10))
                		End If

                	Else
                		DlgText("Text1", "Total Horizontal Lines")
                		DlgText("Text2", "Total Vertical Lines")
                		DlgVisible("hT", False)
                		DlgVisible("vT", False)
                		DlgText("hresT", "10")
                		DlgText("vresT", "10")
                	End If


				Case "OK"
					hres = Evaluate(DlgText("hresT"))
    				vres = Evaluate(DlgText("vresT"))

						If aAxisT = "X" Then
							If DlgValue("DStepOrLine") = 1 Then
								DlgText("hresT", "10")
								DlgText("vresT", "10")
								hres = (ymax-ymin)/hres
								vres = (zmax-zmin)/vres
							End If
                    		DrawGridYZ(ymin, zmin, ymax, zmax, x1, vres, hres)
                    	End If

                		If aAxisT = "Y" Then
                			If DlgValue("DStepOrLine") = 1 Then
                				DlgText("hresT", "10")
								DlgText("vresT", "10")
                				hres = (zmax-zmin)/hres
                				vres = (xmax-xmin)/vres
							End If
                    		DrawGridZX(zmin, xmin, zmax, xmax, y1, vres, hres)
                    	End If

                		If aAxisT = "Z" Then
                			If DlgValue("DStepOrLine") = 1 Then
                				DlgText("hresT", "10")
								DlgText("vresT", "10")
                				hres = (xmax-xmin)/hres
                				vres = (ymax-ymin)/vres
							End If
               				DrawGridXY(xmin, ymin, xmax, ymax, z1, vres, hres)
               			End If

            End Select

    End Select

End Function

Private Function DrawGridXY(u1 As Double, v1 As Double, u2 As Double, v2 As Double , w As Double , vres As Double, hres As Double)
'draw bounding box in XY plane
    Dim i As Integer, j As Integer, umin As Double, vmin As Double
    umin = u1
    vmin = v1

'Draw the bounding box
	Pick.addedge u1, v1, w, u2, v1, w
	Pick.addedge u1, v2, w, u2, v2, w
	Pick.addedge u1, v1, w, u1, v2, w
	Pick.addedge u2, v1, w, u2, v2, w


'Draw 0-0 box
   Pick.addedge -hres*boxsize, -vres*boxsize, w, hres*boxsize, -vres*boxsize, w
   Pick.addedge -hres*boxsize, vres*boxsize, w, hres*boxsize, vres*boxsize, w
   Pick.addedge -hres*boxsize, -vres*boxsize, w, -hres*boxsize, vres*boxsize, w
   Pick.addedge hres*boxsize, -vres*boxsize, w, hres*boxsize, vres*boxsize, w

'draw vertical lines right
    For i = 0 To Fix(u2/hres)
        umin = hres*i
        Pick.addedge umin, v1, w, umin, v2, w
    Next i

'draw vertical lines left
    For j = 0 To Fix(Abs(u1)/hres)
        umin = -hres*j
        Pick.addedge umin, v1, w, umin, v2, w
    Next j

'draw horizontal lines up
    For i = 0 To Fix(v2/vres)
        vmin = vres*i
        Pick.addedge u1, vmin, w, u2, vmin, w
    Next i

'draw horizontal lines down
    For j = 0 To Fix(Abs(v1)/vres)
        vmin = -vres*j
        Pick.addedge u1, vmin, w, u2, vmin, w
    Next j

End Function

Private Function DrawGridZX(w1 As Double, u1 As Double, w2 As Double, u2 As Double , v As Double , vres As Double, hres As Double)
'draw bounding box in ZX plane
    Dim i As Integer, j As Integer, wmin As Double, umin As Double
    wmin = w1
    umin = u1

'Draw the bounding box
	Pick.addedge u1, v, w1, u1, v, w2
	Pick.addedge u2, v, w1, u2, v, w2
	Pick.addedge u1, v, w1, u2, v, w1
	Pick.addedge u1, v, w2, u2, v, w2

'Draw 0-0 box
   Pick.addedge -vres*boxsize, v, -hres*boxsize, -vres*boxsize, v, hres*boxsize
   Pick.addedge vres*boxsize, v, -hres*boxsize, vres*boxsize, v, hres*boxsize
   Pick.addedge -vres*boxsize, v, -hres*boxsize, vres*boxsize, v, -hres*boxsize
   Pick.addedge -vres*boxsize, v, hres*boxsize, vres*boxsize, v, hres*boxsize

'draw vertical lines right
    For i = 0 To Fix(w2/hres)
        wmin = hres*i
        Pick.addedge u1, v, wmin, u2, v, wmin
    Next i

'draw vertical lines left
    For j = 0 To Fix(Abs(w1)/hres)
        wmin = -hres*j
        Pick.addedge u1, v, wmin, u2, v, wmin
    Next j

'draw horizontal lines up
    For i = 0 To Fix(u2/vres)
        umin = vres*i
        Pick.addedge umin, v, w1, umin, v, w2
    Next i

'draw horizontal lines down
    For j = 0 To Fix(Abs(u1)/vres)
        umin = -vres*j
        Pick.addedge umin, v, w1, umin, v, w2
    Next j

End Function

Private Function DrawGridYZ(v1 As Double, w1 As Double, v2 As Double, w2 As Double , u As Double , vres As Double, hres As Double)
'draw bounding box in YZ plane
    Dim i As Integer, j As Integer, vmin As Double, wmin As Double, hor As Double, ver As Double
    vmin = v1
    wmin = w1

'Draw the bounding box
	Pick.addedge u, v1, w1, u, v2, w1
	Pick.addedge u, v1, w2, u, v2, w2
	Pick.addedge u, v1, w1, u, v1, w2
	Pick.addedge u, v2, w1, u, v2, w2

'Draw 0-0 box
	Pick.addedge u, -hres*boxsize, -vres*boxsize, u, hres*boxsize, -vres*boxsize
	Pick.addedge u, -hres*boxsize, vres*boxsize, u, hres*boxsize, vres*boxsize
	Pick.addedge u, -hres*boxsize, -vres*boxsize, u, -hres*boxsize, vres*boxsize
	Pick.addedge u, hres*boxsize, -vres*boxsize, u, hres*boxsize, vres*boxsize

'draw vertical lines right
    For i = 0 To Fix(v2/hres)
        vmin = hres*i
        Pick.addedge u, vmin, w1, u, vmin, w2
    Next i

'draw vertical lines left
    For j = 0 To Fix(Abs(v1)/hres)
        vmin = -hres*j
        Pick.addedge u, vmin, w1, u, vmin, w2
    Next j

'draw horizontal lines up
    For i = 0 To Fix(w2/vres)
        wmin = vres*i
        Pick.addedge u, v1, wmin, u, v2, wmin
    Next i

'draw horizontal lines down
    For j = 0 To Fix(Abs(w1)/vres)
        wmin = -vres*j
        Pick.addedge u, v1, wmin, u, v2, wmin
    Next j

End Function
