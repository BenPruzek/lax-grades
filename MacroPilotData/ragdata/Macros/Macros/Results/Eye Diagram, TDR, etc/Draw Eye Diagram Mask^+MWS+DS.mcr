'#Language "WWB-COM"

' ================================================================================================
' Macro: Draw Eye Diagram Mask
'
' Copyright 2009-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 10-Apr-2019 pgl: fixed rendering of hatching lines
' 18-Dec-2009 ube: bugfix for negative levels
' 10-Dec-2009 ube: first version
' ================================================================================================

Option Explicit

Const NLines=20

Dim m(6) As Double, x(6) As Double, y(6) As Double

Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		If (DlgItem = "OK") Then
		    ' The user pressed the Ok button. Check the settings and display an error message if some required
		    ' fields have been left blank.

		    Dim bSettingsOK As Boolean
		    bSettingsOK = True

		    If (Evaluate(DlgText("Twidth")) < 0) Then bSettingsOK = False
		    If (Evaluate(DlgText("Trise")) < 0)  Then bSettingsOK = False
		    If (Evaluate(DlgText("High")) < Evaluate(DlgText("Low")))  Then bSettingsOK = False

			If (Not bSettingsOK) Then
				MsgBox "Please check and complete your settings. Twidth and Trise must be non-negative. High Level must be larger then Low Level.", vbCritical
				DialogFunction = True						' There is an error in the settings -> Don't close the dialog box.
			End If
		End If
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function

Sub Main

	Dim sPictureFile As String
	sPictureFile = GetInstallPath + "\Library\Macros\Results\Eye Diagram, TDR, etc\eyemask.bmp"

	Begin Dialog UserDialog 420,371,"Add Eye Mask",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 10,203,400,105,"",.GroupBox1
		Text 230,224,80,14,"High Level:",.Text1
		Text 230,252,90,14,"Low Level:",.Text2
		Text 20,224,90,14,"Time center:",.Text3
		Text 20,252,90,14,"Time width:",.Text4
		Text 20,280,90,14,"Time rise:",.Text5
		OKButton 20,343,90,21
		CancelButton 120,343,90,21
		TextBox 110,217,90,21,.Tcenter
		TextBox 110,245,90,21,.Twidth
		TextBox 110,273,90,21,.Trise
		TextBox 310,217,90,21,.High
		TextBox 310,245,90,21,.Low
		Picture 35,14,350,175,sPictureFile,0,.Picture1
		CheckBox 20,315,360,14,"Delete existing eye mask(s)",.Delete
	End Dialog
	Dim dlg As UserDialog

	dlg.High   = GetSetting("CST STUDIO SUITE", "EyeMask", "High", "0.9")
	dlg.Low    = GetSetting("CST STUDIO SUITE", "EyeMask", "Low", "0.1")
	dlg.Tcenter= GetSetting("CST STUDIO SUITE", "EyeMask", "Tcenter", "0.5")
	dlg.Twidth = GetSetting("CST STUDIO SUITE", "EyeMask", "Twidth", "0.2")
	dlg.Trise  = GetSetting("CST STUDIO SUITE", "EyeMask", "Trise", "0.05")
	dlg.Delete  = GetSetting("CST STUDIO SUITE", "EyeMask", "Delete", 1)

	If (Dialog(dlg) = 0) Then Exit All

    SaveSetting  "CST STUDIO SUITE", "EyeMask", "High", dlg.High
    SaveSetting  "CST STUDIO SUITE", "EyeMask", "Low", dlg.Low
    SaveSetting  "CST STUDIO SUITE", "EyeMask", "Tcenter", dlg.Tcenter
    SaveSetting  "CST STUDIO SUITE", "EyeMask", "Twidth", dlg.Twidth
    SaveSetting  "CST STUDIO SUITE", "EyeMask", "Trise", dlg.Trise
    SaveSetting  "CST STUDIO SUITE", "EyeMask", "Delete", dlg.Delete

	Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
	Dim tctr As Double, twidth As Double, trise As Double
	Dim a0 As Double, a1 As Double, a2 As Double

	tctr = Evaluate(dlg.Tcenter)
	twidth = Evaluate(dlg.Twidth)
	trise = Evaluate(dlg.Trise)
	a0 = Evaluate(dlg.Low)
	a2 = Evaluate(dlg.High)
	a1 = 0.5 * (a2+a0)
	t1 = tctr - 0.5 * twidth
	t2 = tctr + 0.5 * twidth
	t0 = t1 - trise
	t3 = t2 + trise

	If (dlg.Delete) Then
		Plot1D.DeleteAllBackGroundShapes
	End If

	Plot1D.AddThickBackGroundLine(t0, a1, t1, a2)
	Plot1D.AddThickBackGroundLine(t1, a2, t2, a2)
	Plot1D.AddThickBackGroundLine(t2, a2, t3, a1)
	Plot1D.AddThickBackGroundLine(t3, a1, t2, a0)
	Plot1D.AddThickBackGroundLine(t2, a0, t1, a0)
	Plot1D.AddThickBackGroundLine(t1, a0, t0, a1)

	If (trise = 0.0) And (twidth = 0.0) Then
		Plot1D.Plot
		Exit All
	End If

	If (trise = 0.0) Then trise=1e-35

	x(1)=t0: y(1)=a1: m(1)=(a2-a1)/trise
	x(2)=t1: y(2)=a2: m(2)=0
	x(3)=t2: y(3)=a2: m(3)=(a1-a2)/trise
	x(4)=t3: y(4)=a1: m(4)=(a1-a0)/trise
	x(5)=t2: y(5)=a0: m(5)=0
	x(6)=t1: y(6)=a0: m(6)=(a0-a1)/trise

	Dim xx As Double, yy As Double, ii As Integer, nfound As Integer
	Dim xnew(6) As Double, ynew(6) As Double

	Dim dt As Double
	dt = (t3-t0)/(NLines)

	If (t1 = t2) Then
		m(0)=(a2-a0)
	Else
		m(0)=(a2-a0)/(t2-t1)
	End If

	If m(0) > m(1) Then
		' start at point 1 (t0/a1)
		x(0)=x(1)
		y(0)=y(1)
	Else
		' start at point 2 (t1/a2)
		x(0)=x(2)
		y(0)=y(2)
	End If

	Dim a2compare As Double, a0compare As Double, t0compare As Double, t1compare As Double, t2compare As Double, t3compare As Double

	a2compare = a2 + (a2-a0) * 0.0000001
	a0compare = a0 - (a2-a0) * 0.0000001
	t0compare = t0 - (t3-t0) * 0.0000001
	t1compare = t1 - (t3-t0) * 0.0000001
	t2compare = t2 + (t3-t0) * 0.0000001
	t3compare = t3 + (t3-t0) * 0.0000001

	' now draw other thin lines, a few calculations are necessary to calc crossing points

	Do

		x(0) = x(0)+dt

		nfound = 0
		For ii=1 To 6
			xnew(ii)=0 : ynew(ii)=0
		Next

		For ii=1 To 6
			CalcLineCrossing(ii,xx,yy)

			If ((xx<t1compare) Or (xx>t2compare)) And (ii=2 Or ii=5) Then
				' skip this point
			ElseIf (yy>a2compare) Or (yy<a0compare) Or (xx<t0compare) Or (xx>t3compare) Then
				' skip this point
			Else
				nfound = nfound + 1
				xnew(nfound)=xx
				ynew(nfound)=yy
			End If
		Next

		If (nfound >= 2) Then
			If (nfound > 2) Then
				Dim xabove As Double, yabove As Double, xbelow As Double, ybelow As Double

				xabove = 0.0
				yabove = -1e99
				xbelow = 0.0
				ybelow = +1e99
				
				For ii=1 To nfound
					If ynew(ii)<ybelow Then
						ybelow = ynew(ii)
						xbelow = xnew(ii)
					End If
					If ynew(ii)>yabove Then
						yabove = ynew(ii)
						xabove = xnew(ii)
					End If
				Next
				xnew(1)=xabove
				ynew(1)=yabove
				xnew(2)=xbelow
				ynew(2)=ybelow
			End If

			' Wait 0.00001

			For ii=1 To 10
				' silly dummy loop (otherwise timing problems)
			Next

			Plot1D.AddThinBackGroundLine(xnew(1), ynew(1), xnew(2), ynew(2))

		End If

	Loop Until (nfound < 2)

	Plot1D.Plot

End Sub
Sub CalcLineCrossing(i As Integer, xx As Double ,yy As Double)

	If (m(i)=m(0)) Then
		xx = -1.2345e32
		yy = -1.2345e32
	Else
		xx=(y(0)-m(0)*x(0)-y(i)+m(i)*x(i))/(m(i)-m(0))
		yy=y(0)+m(0)*(xx-x(0))
	End If

End Sub
