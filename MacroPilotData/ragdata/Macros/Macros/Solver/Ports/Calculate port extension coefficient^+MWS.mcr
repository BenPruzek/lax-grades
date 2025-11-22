' Calculate\Calculate port extension coefficient
' !!! Do not change the line above !!!
'--------------------------------------------------------------------------------------------
' The macro calculates port extension koefficient
'--------------------------------------------------------------------------------------------
' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 09-Oct-2017 kkt: removed CPW and CPWG option button,  corrected macro location in AddHist
' 10-Oct-2016 ube: removed Help button, since no Online Help page is existing
' 30-Sep-2011 ube: one file location still was incorrect
' 05-Aug-2011 ube: moved to folder Library\Macros\Solver\Ports\
' 14-Jan-2011 vso: added brief macro description below the line cross-section figure
' 15-Dec-2010 jei: pick in pictures, increased maximal number of points
' 28-Oct-2010 vso,jei: debugging
' 13-Oct-2010 jei: port construction, line type detection
' 10-Sep-2010 jei: load dimensions from picks
' 03-Sep-2010 jei: closing button, construct feature for MS
' 27-Aug-2010 jei: corrected stripline height definition, added MS 0.1GHz data
' 21-Aug-2010 jei: added stripline 1% error and tested interpolation algorithm
' 08-Aug-2010 jei: version 1 - only for microstrip and 1%error
'--------------------------------------------------------------------------------------------

Option Explicit

Public lineType$
' microstrip
Public MS100#()  ' data 0.1GHz
Public MS1000#()  ' data 1GHz
Public MS2000#()  ' data 2GHz
Public MS5000#()  ' data 5GHz
Public MS10000#()  ' data 10GHz
Public MS15000#()  ' data 15GHz
Public MS20000#()  ' data 20GHz
Public MSfreq#() ' !! list of available frequencies

' slotline
Public SL#()  ' data for all frequencies

' next port number
Public iPort%


Sub Main ()
	Debug.Clear

' data load
	MS100=LoadData("Calculate port extension coefficient_MS_0.1GHz.txt")
	MS1000=LoadData("Calculate port extension coefficient_MS_1GHz.txt")
	MS2000=LoadData("Calculate port extension coefficient_MS_2GHz.txt")
	MS5000=LoadData("Calculate port extension coefficient_MS_5GHz.txt")
	MS10000=LoadData("Calculate port extension coefficient_MS_10GHz.txt")
	MS15000=LoadData("Calculate port extension coefficient_MS_15GHz.txt")
	MS20000=LoadData("Calculate port extension coefficient_MS_20GHz.txt")

	ReDim MSfreq(7)
	MSfreq(1)=0.1
	MSfreq(2)=1
	MSfreq(3)=2
	MSfreq(4)=5
	MSfreq(5)=10
	MSfreq(6)=15
	MSfreq(7)=20

	SL=LoadData("Calculate port extension coefficient_SL.txt")

' dialog
	Begin Dialog UserDialog 870,380,"Calculate port extension coefficient",.dlgFcn ' %GRID:5,5,0,1

		GroupBox 25,10,275,135,"Type",.GB_type
		OptionGroup .options
			OptionButton 45,32,90,15,"Microstrip",.OB_MS
			OptionButton 45,60,90,15,"Strip Line",.OB_SL
		'	OptionButton 45,90,180,15,"Coplanar waveguide",.OB_CPW
		'	OptionButton 45,120,250,15,"Coplanar waveguide with ground",.OB_CPWG

		GroupBox 25,160,275,115,"Dimensions",.GB_dim
		Text 45,187,55,15,"w [mm]",.T_dim1
		Text 45,217,55,15,"h [mm]",.T_dim2
		Text 45,247,55,15,"Text1",.T_dim3

		TextBox 110,185,165,20,.TB_dim1
		TextBox 110,215,165,20,.TB_dim2
		TextBox 110,245,165,20,.TB_dim3

		GroupBox 25,285,275,55,"Material Properties",.GB_mat
		Text 50,313,130,15,"Er",.T_er
		TextBox 110,310,165,20,.TB_er

		Text 45,355,260,15,"Frequency range",.T_f

		Picture 330,15,519,180,"Picture1",0,.P_type

		GroupBox 330,285,515,55,"Extension Coefficient",.GB_k
		Text 355,313,30,15,"k =",.T_k
		TextBox 390,310,165,20,.TB_k
		Text 580,313,215,15,"k Freq Range ",.T_kRange

		PushButton 330,350,90,20,"Calculate",.PB_calculate
		PushButton 655,350,90,20,"Close",.PB_close
		PushButton 425,350,225,20,"Construct port from picked face",.PB_construct

		CancelButton 545,90,90,20,.CB ' visible - false for closing by X

		PushButton 750,350,90,20,"Help",.PB_help
		Text 335,205,510,15,"Please pick the metal face as depicted above before launching the macro! ",.PictureCaption
		Text 335,230,510,45,"Macro helps to set up the waveguide port size for planar transmission lines. The size of the port is extended by factor k in order to get line impedance with error smaller that 1%. The extension coefficient can be however adjusted manually.",.Text2

	End Dialog

' for closing by "x"
	Dim dlg As UserDialog
	If Dialog(dlg)=1 Then
		' nothing
	End If

End Sub







' DIALOG FUNCTION

Function DlgFcn%(DlgItem$, Action%, SuppValue%)
' dialog function

' universal variables
	Dim sline$
	Dim i1%

' other variables
	Dim Er#
	Dim w#
	Dim h#
	Dim fmin#
	Dim fmax#
	Dim k#
	Dim B#

	Dim t#()


	Select Case Action%
	Case 1 ' Dialog box initialization
	' buttons
	'	DlgEnable("OB_CPW",False)
	'	DlgEnable("OB_CPWG",False)
		DlgVisible("PB_help",False)
		DlgVisible("CB",False)
		DlgEnable("PB_construct",False)

		DlgText("T_kRange","")
	' f range
		DlgText("T_f","Frequency range: " & cstr(Solver.getfmin) & " to " & cstr(Solver.getfmax) & " " & cstr(Units.GetUnit("Frequency")))

	' next port number
		If Pick.GetNumberOfPickedFaces=1 Then
			i1=0
			Port.StartPortNumberIteration
			While i1<>-1
				iPort=i1
				i1=Port.GetNextPortNumber
			Wend
			iPort=iPort+1
			' autodetect
			SelectType(DetectType())
		Else
			SelectType(0) ' MS is default
		End If



	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "options"
			SelectType(SuppValue)
		Case "PB_close"
			Exit All

		Case "PB_calculate"
			DlgFcn=True ' do not close the dialog
			If IsInputOK(False) = True Then
				Select Case lineType
				Case "MS"
					DlgText("TB_k",cstr(MS_CalculateK() ))
				Case "SL"
					DlgText("TB_k",cstr(SL_CalculateK() ))
				Case "CPW"
					cdbl(DlgText("TB_dim3"))

				Case "CPWG"
				End Select
			End If

		Case "PB_construct"
			DlgFcn=True ' do not close the dialog
			DlgEnable("PB_construct")
			If IsInputOK(False) = True Then

				Select Case lineType
				Case "MS"
					If IsKset = False Then
						k=MS_CalculateK()
						DlgText("TB_k",cstr(k))
					End If
					MS_ConstructPort()

				Case "SL"
					If IsKset = False Then
						k=SL_CalculateK()
						DlgText("TB_k",cstr(k))
					End If
					SL_ConstructPort()

				Case "CPW"
				Case "CPWG"
				End Select
			End If


		End Select

	Case 3 ' TextBox or ComboBox text changed

	Case 4 ' Focus changed

	Case 5 ' Idle

	End Select
End Function









' HELP FUNCTIONS

Function ErrSub(num#,Optional id$ = "")
	Select Case id
	Case "sub"
		MsgBox("Macro is unable to recognize the substrate. Please fill in the parameters manualy.")
	End Select

	Debug.Print "Error" & num

End Function


Function LoadData(fName$)
' returns data matrix
	Dim sline$
	Dim dataRows%
	Dim A#()
	Dim iRow%

	Open GetInstallPath + "\Library\Macros\Solver\Ports\" + fName For Input As #1
	Line Input #1,sline ' headder
	dataRows=cint(Left(sline,InStr(sline,"_")-1))
	ReDim A(dataRows,3)

	Line Input #1,sline ' headder

    For iRow=1 To dataRows
      	Line Input #1,sline
      	A(iRow,1)=cdbl(Mid(sline,1,13))
   	  	A(iRow,2)=cdbl(Mid(sline,18,13))
   		A(iRow,3)=cdbl(Mid(sline,34,14))
    Next

    LoadData=A

	Close #1

End Function

Function Is0(num#) As Boolean
	Dim limit#
	limit=1e-12
	If Abs(num#)<limit Then
		Is0=True
	Else
		Is0=False
	End If

End Function

Function Round2(num#,pos%)
' returns rounded value
	Round2=Int(num*10^pos)
	Round2=Round2*10^(-pos)
End Function

Function GetArea(xA#,yA#,xB#,yB#,xc#,yc#)
' returns area of triangle
	GetArea = Abs((xB*yA-xA*yB)+(xc*yB-xB*yc)+(xA*yc-xc*yA))/2

End Function

Function IsInRange(Er#,var1#,Optional var2#) As Boolean
' check if all input variables are in range
' MS: (Er,w/h)
' SL: (Er,2*w/B)


	Dim x_min#
	Dim x_max#
	Dim sline$
	IsInRange=True
	sline="Some variables are out of range. Please change the input to satisfy:" & vbCrLf
	' Er check
	If Not IsInLim(Er,1,x_min,x_max) Then
		IsInRange=False
		sline=sline & vbCrLf & cstr(x_min) & " < Er < " & cstr(x_max)
	End If

	Select Case lineType
	Case "MS"
		' wh limits
		If Not IsInLim(var1,2,x_min,x_max) Then
			IsInRange=False
			sline=sline & vbCrLf & cstr(x_min) & " < W/h < " & cstr(x_max) & "     (W/h=" & cstr(Round2(var1,2)) & ")"
		End If
	Case "SL"
		' wh limits
		' var1 is W/h
		If Not IsInLim(var1,2,x_min,x_max) Then
			IsInRange=False
			sline=sline & vbCrLf & cstr(x_min/2) & " < W/B < " & cstr(x_max/2)& "     (W/B=" & cstr(Round2(var1/2,2)) & ")"
		End If
	Case "CPW"
	Case "CPWG"

	End Select

	If IsInRange=False Then
		MsgBox(sline)
	End If



End Function

Function IsInLim(x#,column#,ByRef x_min#, ByRef x_max#) As Boolean
' simple check
	Dim x_i#
	Dim i%

	x_min=1e6
	x_max=0

	For i=1 To UBound(MS100,1)
		x_i=MS100(i,column)
		If x_i>x_max Then
			x_max=x_i
		End If
		If x_i<x_min Then
			x_min=x_i
		End If
	Next

	If x>x_max Or x<x_min Then
		IsInLim=False
	Else
		IsInLim=True
	End If

End Function

Function GetT(shName$, faceId As Long, cPt#(), ByRef t#(), Optional ByRef tMinAx%)
' returns thickness of shape
' cPt is center point
' tMinAx is axis of minimal t except zero..

	Dim endIt As Boolean
	Dim i1%
	Dim i2%

	Dim iPt#()
	ReDim iPt(3)

	Dim checkFaceID As Long
	Dim ti#

	endIt=False
	i1=0

	' pick the right point
	While endIt=False
		If Solid.GetPointCoordinates( shName,i1,iPt(1),iPt(2),iPt(3)) Then
			For i2=1 To 3
				t(i2)=2*Abs(cPt(i2)-iPt(i2))
				If Not Is0(t(i2)) Then
					iPt(i2)=iPt(i2)+0.1*(cPt(i2)-(iPt(i2)))	 ' for further use
				End If
			Next

			For i2=1 To 3
				If Is0(t(i2)) Then

					Debug.Print iPt(1) & "  " & iPt(2) & "  " & iPt(3)

					Pick.clearAllPicks
					Pick.PickFaceFromPoint(shName,iPt(1),iPt(2),iPt(3))
					Pick.GetPickedFaceFromIndex (1, checkFaceID)
					Debug.Print faceId
					Debug.Print checkFaceID
					If checkFaceID=faceId Then
						endIt=True
					End If

				End If
			Next
		End If

		i1=i1+1
		If i1>10000 Then
			MsgBox("Thickness not found.")
			Pick.clearAllPicks
			Exit Function
		End If
	Wend

' axis of minimal t except 0
	ti=1e10
	For i1=1 To 3
		If ti>t(i1) And Not Is0(t(i1)) Then
			ti=t(i1)
			tMinAx=i1
		End If
	Next

	Pick.clearAllPicks
	' pick original face
	Pick.PickFaceFromId ( shName,  faceId )


End Function
Function AddHist$(M$())
' M is range add matrix


	AddHist="' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient" & vbCrLf & _
			"" & vbCrLf & _
			"" & vbCrLf & _
			"With Port" & vbCrLf & _
			"  .Reset" & vbCrLf & _
			"  .PortNumber " & Chr(34) & cstr (iPort) & Chr(34) & vbCrLf & _
			"  .NumberOfModes " & Chr(34) & "1"  & Chr(34) & vbCrLf & _
			"  .AdjustPolarization False" & vbCrLf & _
			"  .PolarizationAngle " & Chr(34) & "0.0" & Chr(34) & vbCrLf & _
			"  .ReferencePlaneDistance " & Chr(34) & "0" & Chr(34) & vbCrLf & _
			"  .TextSize " & Chr(34) & "50" & Chr(34) & vbCrLf & _
			"  .Coordinates " & Chr(34) & "Picks" & Chr(34) & vbCrLf & _
			"  .Orientation " & Chr(34) & "Positive" & Chr(34) & vbCrLf & _
			"  .PortOnBound " & Chr(34) & "True" & Chr(34) & vbCrLf & _
			"  .ClipPickedPortToBound " & Chr(34) & "False" & Chr(34) &  vbCrLf & _
			"  .XrangeAdd " & Chr(34) & M(1,1) & Chr(34) & ", " & Chr(34) & M(1,2) & Chr(34) & vbCrLf & _
			"  .YrangeAdd " & Chr(34) & M(2,1) & Chr(34) & ", " & Chr(34) & M(2,2) & Chr(34) & vbCrLf & _
			"  .ZrangeAdd " & Chr(34) & M(3,1) & Chr(34) & ", " & Chr(34) & M(3,2) & Chr(34) & vbCrLf & _
			"  .Shield "& Chr(34) & "PEC" & Chr(34) & vbCrLf & _
			"  .SingleEnded " & Chr(34) & "False" & Chr(34) & vbCrLf & _
			"  .Create" & vbCrLf & _
			"End With" & vbCrLf


	'StoreDoubleParameter ( kName ,k )

	AddToHistory("define port:" & cstr(iPort),AddHist)
	RebuildForParametricChange

End Function

Function DetectType%()
' returns lineTypeID (0/1/2/3) or -1

	Dim i1%
	Dim iName$

	Dim stripName$
	Dim stripFaceID As Long

	Dim cPt#()
	ReDim cPt(3)
	Dim posPt#()
	ReDim posPt(3)
	Dim negPt#()
	ReDim negPt(3)

	Dim nPos%
	Dim nNeg%

	Dim posSubName$
	Dim negSubName$

	Dim t#()
	ReDim t(3)

	Dim tMinAx%

	DetectType=-1
	If Not Pick.GetNumberOfPickedFaces=1 Then
		Exit Function
	End If

' stripName and faceId
	stripName=Pick.GetPickedFaceFromIndex (1, stripFaceID)

' thickness of picked face in all directions
	Pick.PickCenterpointFromId (stripName, stripFaceID)
	Pick.GetPickPointCoordinates (1, cPt(1), cPt(2),cPt(3) )

	GetT(stripName, stripFaceID, cPt, t, tMinAx)

' points for check
	posPt=cPt
	posPt(tMinAx)=posPt(tMinAx)+0.501*t(tMinAx)' t * (small number)
	negPt=cPt
	negPt(tMinAx)=negPt(tMinAx)-0.501*t(tMinAx)   ' t * (small number)

' check all shapes
	nPos=0
	nNeg=0

	For i1 = 0 To (Solid.GetNumberOfShapes-1)
		iName=Solid.GetNameOfShapeFromIndex(i1)
		If iName <> stripName Then
			' positive direction
			If Solid.IsPointInsideShape ( posPt(1), posPt(2), posPt(3), iName) Then
				nPos=nPos+1
				posSubName=iName
			End If

			' negative direction
			If Solid.IsPointInsideShape ( negPt(1), negPt(2), negPt(3), iName) Then
				nNeg=nNeg+1
				negSubName=iName
			End If
		End If
	Next

' check nPos nNeg and materials if nessesary
	If (nPos=1 And nNeg=0) Or (nPos=0 And nNeg=1) Then
		DetectType=0
		Exit Function
	End If

	If nPos=1 And nNeg=1 Then
		If posSubName=negSubName Then
			DetectType=1 ' SL
			Exit Function
		ElseIf Solid.GetMaterialNameForShape ( posSubName ) = Solid.GetMaterialNameForShape ( negSubName ) Then
			DetectType=1 ' SL
			Exit Function
		End If
	End If


End Function

Function SelectType(lineTypeID%)
'MS is default
	If lineTypeID = -1 Then lineTypeID = 0
	DlgValue("options",lineTypeID)

	Dim t#()
	Dim h#
	Dim B#
	Dim Bneg#
	Dim w#
	Dim Er#
	Dim posDir As Boolean


	DlgEnable("PB_construct",False)
	DlgText("TB_k","")
	DlgText("T_kRange","")
	DlgText("TB_Er","0")
	DlgText("TB_dim1","0")
	DlgText("TB_dim2","0")


	Select Case lineTypeID
           	Case 0 'MS is selected
           		lineType="MS"
           		DlgText("T_dim1","W [mm]" )
				DlgText("T_dim2","h [mm]" )
				DlgText("T_dim3","-" )
				DlgEnable("TB_dim3",False)
		        DlgSetPicture("P_type",GetInstallPath + "\Library\Macros\Solver\Ports\Calculate port extension coefficient_MS.bmp",0)

		        If Pick.GetNumberOfPickedFaces=1 Then
					If MS_GetDimFromPicks(t,posDir,h,w,Er) = True Then
						DlgText("TB_Er",cstr(Er))
						DlgText("TB_dim1",cstr(w))
						DlgText("TB_dim2",cstr(h))
						If IsInRange(Er,w/h) = True Then
							DlgEnable("PB_construct",True)
						End If
					End If
				End If

           	Case 1 'SL is selected
           		lineType="SL"
           		DlgText("T_dim1","W [mm]" )
				DlgText("T_dim2","B [mm]" )
				DlgText("T_dim3","-" )
				DlgEnable("TB_dim3",False)
		        DlgSetPicture("P_type",GetInstallPath + "\Library\Macros\Solver\Ports\Calculate port extension coefficient_SL.bmp",0)

		        If Pick.GetNumberOfPickedFaces=1 Then
					If SL_GetDimFromPicks(t,B,w,Er,Bneg) = True Then
						DlgText("TB_Er",cstr(Er))
						DlgText("TB_dim1",cstr(w))
						If B > Bneg Then
							DlgText("TB_dim2",cstr(B))
						Else
							DlgText("TB_dim2",cstr(Bneg))
						End If

						If IsInRange(Er,2*w/cdbl(DlgText("TB_dim2"))) = True Then
							DlgEnable("PB_construct",True)
						End If
					End If
				End If
           	Case 2 'CPW is selected
           		DlgEnable("PB_construct",False)

           	Case 3 'CPWG is selected
				DlgEnable("PB_construct",False)


           	End Select

End Function

Function IsInputOK(tryDim3) As Boolean

	On Error GoTo Err:
		cdbl(DlgText("TB_Er"))
		cdbl(DlgText("TB_dim1"))
		cdbl(DlgText("TB_dim2"))
		If tryDim3 = True Then
			cdbl(DlgText("TB_dim3"))
		End If

	GoTo OK:

	Err:
		IsInputOK=False
		MsgBox("Please fill in all parameters.")
		Exit Function

	OK:
		IsInputOK=True
		On Error GoTo 0

End Function

Function IsKset() As Boolean
	On Error GoTo Err1:
		cdbl(DlgText("TB_k"))
	GoTo OK1:

	Err1:
		IsKset=False
		Exit Function
	OK1:
		IsKset=True
 		On Error GoTo 0
End Function










' COEFFICIENT COUNTING FUNCTIONS

Function GetK_int(data#(),Er#,wh#)
' returns interpolated k for specified er and wh and freq~data
' uses getKestim()
	Dim i%
	Dim er_low#
	Dim er_high#
	Dim wh_low#
	Dim wh_high#

	Dim A#
	Dim B#
	Dim k1#
	Dim k2#

' k in corners
	Dim k_ll#
	Dim k_hl#
	Dim k_lh#
	Dim k_hh#
' in center
	Dim k_c#
	Dim er_c#
	Dim wh_c


' lowest and highest in data tables
	er_low=0
	er_high=1e6
	wh_low=0
	wh_high=1e6

' get closest higher and lower value of er,w_h in data table
	For i=1 To UBound(data,1)
		' find er
		If data(i,1)<=Er And data(i,1)> er_low Then
			er_low=data(i,1)
		End If
		If data(i,1)>=Er And data(i,1)< er_high Then
			er_high=data(i,1)
		End If

		' find w_h
		If data(i,2)<=wh And data(i,2)> wh_low Then
			wh_low=data(i,2)
		End If
		If data(i,2)>=wh And data(i,2)< wh_high Then
			wh_high=data(i,2)
		End If
	Next



' interpolation
	If er_low=er_high And wh_low=wh_high Then
		' exact k is in data table
		GetK_int=FindK(data,er_low,wh_low)
	ElseIf er_low=er_high Then
		' k is on line er
		k1=FindK(data,er_low,wh_low)
		k2=FindK(data,er_low,wh_high)
		' k=a*wh + b
		A=(k2-k1)/(wh_high-wh_low)
		B=k1-A*wh_low

		GetK_int=A*wh + B
	ElseIf wh_low=wh_high Then
		' k is on line w_h
		k1=FindK(data,er_low,wh_low)
		k2=FindK(data,er_high,wh_low)
		' k=a*er + b
		A=(k2-k1)/(er_high-er_low)
		B=k1-A*er_low

		GetK_int=A*Er + B
	Else
		' k is in plane
		' get k in 4 corners


		k_ll=FindK(data,er_low,wh_low)
		k_hl=FindK(data,er_high,wh_low)
		k_lh=FindK(data,er_low,wh_high)
		k_hh=FindK(data,er_high,wh_high)

		' get k in center (higher k is considered)


		If (k_hh+k_ll)/2 > (k_hl+k_lh)/2 Then
			k_c =(k_hh+k_ll)/2
		Else
			k_c =(k_hl+k_lh)/2
		End If

		' get er and w_h in center

		er_c=(er_low+er_high)/2
		wh_c=(wh_low+wh_high)/2

		' get k from 4 triangles in space (4 corner points, 1 center point)
		GetK_int=GetKestim(Er, wh,  er_c, wh_c, k_c, er_low,wh_low,k_ll,   er_high,wh_low,k_hl)
		If GetK_int=-1 Then
			GetK_int=GetKestim(Er,wh,  er_c,wh_c,k_c,  er_high,wh_low,k_hl,   er_high,wh_high,k_hh)
			If GetK_int=-1 Then
				GetK_int=GetKestim(Er,wh,  er_c,wh_c,k_c,  er_high,wh_high,k_hh,   er_low,wh_high,k_lh)
				If GetK_int=-1 Then
					GetK_int=GetKestim(Er,wh,  er_c,wh_c,k_c,  er_low,wh_high,k_lh,   er_low,wh_low,k_ll)
				End If
			End If
		End If
	End If


End Function


Function GetKestim(xP#,yP#,xA#,yA#,zA#,xB#,yB#,zB#,xc#,yc#,zc)
' returns interpolated value of k or -1 if P is outside triangle ABC
' uses getArea()

' x ~ Er
' y ~ w_h
' z ~ k

' P - desired xy point, where k should be found
' A,B,C - points of triangle in space ~ plane for interpolation

	Dim PAB#
	Dim PCB#
	Dim PAC#
	Dim CAB#

	Dim A#
	Dim B#
	Dim C#
	Dim D#

	PAB=GetArea(xP,yP,xA,yA,xB,yB)
	PCB=GetArea(xP,yP,xc,yc,xB,yB)
	PAC=GetArea(xP,yP,xA,yA,xc,yc)
	CAB=GetArea(xc,yc,xA,yA,xB,yB)

' Area PAB+Area PBC +Area PAC=Area ABC
	If (PAB+PCB+PAC)-CAB < CAB/1e6 Then
		' is inside triangle

		' plane analytical description
		A = yA *(zB - zc) + yB *(zc - zA) + yc* (zA - zB)
		B = zA *(xB - xc) + zB *(xc - xA) + zc* (xA - xB)
		C = xA *(yB - yc) + xB *(yc - yA) + xc* (yA - yB)
		D = xA *(yB* zc - yc* zB) + xB* (yc* zA - yA* zc) + xc*(yA* zB - yB* zA)
		' !!! -D in A*x + B*y + C*z - D=0
		GetKestim=(-A*xP-B*yP+D)/C

	Else
		' is not inside triangle
		GetKestim=-1
	End If


End Function

Function FindK#(data#(),Er#,wh#)
' simple k find if I know exact er and wh in data table
' -1 error

	Dim i%
	For i=1 To UBound(data,1)
		If Is0(data(i,1)-Er) And Is0(data(i,2)-wh) Then
			FindK=data(i,3)
		End If
	Next
End Function











' MICROSTRIP FUNCTIONS

Function MS_GetKfRange#(Er#,wh#,fMin#,fMax#)
' returns k in freq range
' uses getK_int()
	Dim ifr%
	Dim fi#
	Dim f_low#
	Dim f_high#
	Dim data#()

	Dim ki#
	Dim kMin#
	Dim kMax#

	f_high=1e6
	f_low=0

' get f_low and f_high
	For ifr=1 To UBound(MSfreq)
		fi=MSfreq(ifr)

		If fi<=fMin And fi> f_low Then
			f_low=fi
		End If
		If fi>=fMax And fi< f_high Then
			f_high=fi
		End If
	Next

' find highest k from f_low to f_high
	kMax=0
	kMin=1e6
	For ifr=1 To UBound(MSfreq)
		fi=MSfreq(ifr)
		'Debug.Print "fi=" & fi

		If fi>= f_low And fi<= f_high Then
			data=Eval("MS" & cstr(fi*1000))
			ki=GetK_int(data,Er,wh)
			If ki > kMax Then
				kMax=ki
			End If

			If ki < kMin Then
				kMin=ki
			End If

		End If
	Next

' inhomogenous port
	'If kMax/kMin > 1.5 Then
		'cstr(kMin) & "  " & cstr(kMax) & "  kMax/kMin =" & cstr(kMax/kMin) & vbCrLf & _
	'	MsgBox("To improve accuracy consider using Solver -> Transient solver -> Inhomogenous port accuracy enhancement.")
	'End If
	DlgText("T_kRange","k varies in the range: " & cstr(Round2(kMin,2)) & " - " & cstr(Round2(kMax,2)) )

	MS_GetKfRange=kMax

End Function

Function MS_CalculateK#()
' returns rounded final value
	Dim Er#
	Dim w#
	Dim h#
	Dim fMin#
	Dim fMax#
	Dim k#

	Er=cdbl(DlgText("TB_er"))
	w=cdbl(DlgText("TB_dim1"))
	h=cdbl(DlgText("TB_dim2"))
	fmin=Solver.getfmin*Units.GetFrequencyUnitToSI/1e9
	fmax=Solver.getfmax*Units.GetFrequencyUnitToSI/1e9

	If Not Is0(h) Then
		If IsInRange(Er,w/h,0) Then
			k=MS_GetKfRange(Er,w/h,fmin,fmax)
			k=Round2(k,2)
		End If
	End If

	MS_CalculateK=k
End Function

Function MS_ConstructPort$()
' universal variables
	Dim i1%
	Dim i2%

' concrete variables
	Dim kString$

	Dim t#()	' thickness
	Dim h#  ' substrate height

	' is inside n shapes (x/y/z   ,   positive/ negative direction)
	Dim posDir As Boolean

	' Range Add Matix
	Dim M$()
	ReDim M(3,2)

	MS_GetDimFromPicks(t, posDir,h)

	kString=DlgText("TB_k")

' Range Add Matix
	For i1 = 1 To 3
		If Is0(t(i1)) Then ' thickness =0
			M(i1,1)="0"
			M(i1,2)="0"
		Else ' t(i1) <> 0
			For i2 = 1 To 3
				If t(i2) <> 0  And t(i1) < t(i2) Then ' t(i2) is thicker
					M(i2,1)=cstr(h) & "*" & kString
					M(i2,2)=cstr(h) & "*" & kString
					If  posDir = True Then ' positive direction
						M(i1,1)=cstr(h) & "*" & kString
						M(i1,2)=cstr(h)
					Else ' negative direction
						M(i1,1)=cstr(h)
						M(i1,2)=cstr(h) & "*" & kString
					End If
				End If
			Next i2
		End If
	Next i1

' history add

	AddHist(M)

End Function

Function MS_GetDimFromPicks(ByRef t#(), ByRef posDir As Boolean,  Optional ByRef h#,Optional ByRef w#, Optional ByRef Er#) As Boolean

	MS_GetDimFromPicks=True

' universal variables
	Dim endIt As Boolean
	Dim iName$

	Dim i1%
	Dim i2%
	Dim i3%

	Dim ti#   ' for cycle
	Dim iPt#()
	ReDim iPt(3)


' concrete veriables

	' thickness
	ReDim t(3)

	Dim tMinAx%

	' possible substrate height
	Dim hs#()
	ReDim hs(3)

	' center point
	Dim cPt#()
	ReDim cPt(3)

	Dim stripFaceID As Long
	Dim stripName$
	Dim subName$
	Dim subMatName$ ' material


	' material
	Dim EpsX#
	Dim EpsY#
	Dim EpsZ#


' stripName and faceId
	stripName=Pick.GetPickedFaceFromIndex (1, stripFaceID)

' thickness of picked face in all directions
	Pick.PickCenterpointFromId (stripName, stripFaceID)
	Pick.GetPickPointCoordinates (1, cPt(1), cPt(2),cPt(3) )
	Pick.ClearAllPicks

	GetT(stripName, stripFaceID, cPt, t,tMinAx)

	subName=MS_GetSubName(t,cPt,stripName,posDir)
	If subName="" Then
		MS_GetDimFromPicks=False
		Pick.PickFaceFromId ( stripName,  stripFaceID )
		Exit Function
	End If

	Debug.Print posDir

' get substrate height
	endIt = False
	i1=0
	While endIt=False
		endIt=False
		' calculate height in all directions
 		If Solid.GetPointCoordinates( subName,i1,iPt(1),iPt(2),iPt(3)) Then
			For i2=1 To 3
				hs(i2)=cPt(i2)-iPt(i2) ' positive or negative!!
			Next
		End If

		' check if the height is possible
		For i2= 1 To 3
			If i2=tMinAx And Not Is0(Abs(hs(i2))-t(i2)/2) Then
				If posDir=True Then
					If hs(i2) < 0 Then
						endIt=True
					End If
				Else ' negative dir
					If hs(i2) > 0 Then
						endIt=True
					End If
				End If
			End If
		Next
		i1=i1+1
		If i1>1000 Then
			ErrSub(6,"sub")
			MS_GetDimFromPicks=False
			Pick.PickFaceFromId ( stripName,  stripFaceID )
			Exit Function
		End If
	Wend


	' substrate height is in strip smallestt thickness direction
	h=Abs(hs(tMinAx))-t(tMinAx)/2

' substrate Er
	subMatName=Solid.GetMaterialNameForShape ( subName)

	Material.GetEpsilon ( subMatName, EpsX, EpsY, EpsZ )
	Er=EpsX

' strip width
	w=0
	For i1=1 To 3
		If t(i1)>w Then
			w=t(i1)
		End If
	Next

' pick original face
	Pick.PickFaceFromId ( stripName,  stripFaceID )

End Function

Function MS_GetSubName$(t#(), centerPt#(),stripName$, ByRef posDir As Boolean)
	MS_GetSubName=""

' universal
	Dim i1%
	Dim iName$

	Dim ti#

	Dim EpsX#
	Dim EpsY#
	Dim EpsZ#

' concrete
	Dim tMinAx% ' axis of minimal t
	Dim PtPos#()
	Dim PtNeg#()

	Dim nPos% ' number of shapes
	Dim nNeg%

	Dim posSubName$
	Dim negSubName$

	Dim posSubMatName$
	Dim negSubMatName$

	Dim negEpsXYZ#
	Dim posEpsXYZ#

' axis of minimal t except 0
	ti=1e10
	For i1=1 To 3
		If ti>t(i1) And Not Is0(t(i1)) Then
			ti=t(i1)
			tMinAx=i1
		End If
	Next

' points for check
	PtPos=centerPt
	PtPos(tMinAx)=PtPos(tMinAx)+0.501*t(tMinAx)' t * (small number)
	PtNeg=centerPt
	PtNeg(tMinAx)=PtNeg(tMinAx)-0.501*t(tMinAx)   ' t * (small number)

' check all shapes
	nPos=0
	nNeg=0

	For i1 = 0 To (Solid.GetNumberOfShapes-1)
		iName=Solid.GetNameOfShapeFromIndex(i1)
		If iName <> stripName Then
			' positiv direction
			If Solid.IsPointInsideShape ( PtPos(1), PtPos(2), PtPos(3), iName) Then
				nPos=nPos+1
				posSubName=iName
			End If

			' negative direction
			If Solid.IsPointInsideShape ( PtNeg(1), PtNeg(2), PtNeg(3), iName) Then
				nNeg=nNeg+1
				negSubName=iName
			End If
		End If
	Next

' is there more than one shape?
	If nPos>1 Or nNeg >1 Then
			ErrSub(2,"sub")

		Exit Function
	End If

' check number of shapes
	If nPos<>nNeg Then
		If nPos>nNeg Then
			MS_GetSubName=posSubName
			posDir=True
		Else
			MS_GetSubName=negSubName
			posDir=False
		End If

	Else ' check possible substrates  Er

		posSubMatName=Solid.GetMaterialNameForShape ( posSubName)
		negSubMatName=Solid.GetMaterialNameForShape ( negSubName)

		Material.GetEpsilon ( posSubMatName, EpsX, EpsY, EpsZ )
		posEpsXYZ=EpsX+EpsY+EpsZ

		Material.GetEpsilon ( negSubMatName, EpsX, EpsY, EpsZ )
		negEpsXYZ=EpsX+EpsY+EpsZ

		If posEpsXYZ = 3 And negEpsXYZ <> 3 Then
			MS_GetSubName=negSubName
			posDir=False
		ElseIf posEpsXYZ <> 3 And negEpsXYZ = 3 Then
			MS_GetSubName=posSubName
			posDir=True
		Else
			ErrSub(3,"sub")
		End If

	End If

End Function









' STRIPLINE FUNCTIONS

Function SL_CalculateK#()
	Dim Er#
	Dim w#
	Dim B#
	Dim k#

	Er=cdbl(DlgText("TB_er"))
	w=cdbl(DlgText("TB_dim1"))
	B=cdbl(DlgText("TB_dim2"))
	If Not Is0(B) Then
		If IsInRange(Er,2*w/B,0) Then
			k=GetK_int(SL,Er,2*w/B)
			k=Round2(k,2)
		End If
	End If
	SL_CalculateK=k
End Function

Function SL_GetDimFromPicks(ByRef t#(), Optional ByRef B#, Optional ByRef w#, Optional ByRef Er#,Optional ByRef Bneg) As Boolean

	SL_GetDimFromPicks=True

' universal
	Dim ti#
	Dim i1%
	Dim i2%
	Dim iName$
	Dim endIt As Boolean


' concrete
	Dim stripFaceID As Long
	Dim stripName$
	Dim subMatName$

	Dim posSubName$
	Dim negSubName$

	Dim tMinAx%

	Dim cPt#()
	ReDim cPt(3)

	Dim posPt#()
	Dim negPt#()
	Dim nPos%
	Dim nNeg%

	ReDim t(3)

	Dim iPt#()
	ReDim iPt(3)

	Dim hsP#()
	ReDim hsP(3)
	Dim hsN#()
	ReDim hsN(3)

	Dim EpsX#
	Dim EpsY#
	Dim EpsZ#



' stripName and faceId
	stripName=Pick.GetPickedFaceFromIndex (1, stripFaceID)

' thickness of picked face in all directions
	Pick.PickCenterpointFromId (stripName, stripFaceID)
	Pick.GetPickPointCoordinates (1, cPt(1), cPt(2),cPt(3) )
	Pick.ClearAllPicks

	GetT(stripName,stripFaceID, cPt, t,tMinAx)

	posSubName=""
	negSubName=""

' points for check
	posPt=cPt
	posPt(tMinAx)=posPt(tMinAx)+0.501*t(tMinAx)' t * (small number)
	negPt=cPt
	negPt(tMinAx)=negPt(tMinAx)-0.501*t(tMinAx)   ' t * (small number)

' check all shapes
	nPos=0
	nNeg=0

	For i1 = 0 To (Solid.GetNumberOfShapes-1)
		iName=Solid.GetNameOfShapeFromIndex(i1)
		If iName <> stripName Then
			' positive direction
			If Solid.IsPointInsideShape ( posPt(1), posPt(2), posPt(3), iName) Then
				nPos=nPos+1
				posSubName=iName
			End If

			' negative direction
			If Solid.IsPointInsideShape ( negPt(1), negPt(2), negPt(3), iName) Then
				nNeg=nNeg+1
				negSubName=iName
			End If
		End If
	Next

' check number of shapes ...
	If nPos=1 And nNeg=1 Then
		If Solid.GetMaterialNameForShape ( posSubName ) <> Solid.GetMaterialNameForShape ( negSubName ) Then
			ErrSub(4,"sub")
			SL_GetDimFromPicks=False
			Pick.PickFaceFromId ( stripName,  stripFaceID )
			Exit Function
		End If
	Else
		ErrSub(5,"sub")
		SL_GetDimFromPicks=False
		Pick.PickFaceFromId ( stripName,  stripFaceID )
		Exit Function
	End If


' height positive
	endIt = False
	i1=0
	While endIt=False
		endIt=True
		' calculate height in all directions
 		If Solid.GetPointCoordinates( posSubName,i1,iPt(1),iPt(2),iPt(3)) Then
			For i2=1 To 3
				hsP(i2)=Abs(cPt(i2)-iPt(i2))
			Next
		End If

		' check if the height is possible
		For i2= 1 To 3
			If Is0(t(i2)) Then
				If Is0(hsP(i2)) Then
					endIt=False
				End If
			ElseIf hsP(i2) < 0.501* t(i2) Then ' wrong point
				endIt=False
			End If
		Next
		i1=i1+1
		If i1>1000 Then
			ErrSub(6,"sub")
			SL_GetDimFromPicks=False
			Pick.PickFaceFromId ( stripName,  stripFaceID )
			Exit Function
		End If
	Wend


' height negative
	If posSubName <> negSubName Then
		endIt = False
		i1=0
		While endIt=False
			endIt=True
			' calculate height in all directions
	 		If Solid.GetPointCoordinates( negSubName,i1,iPt(1),iPt(2),iPt(3)) Then
				For i2=1 To 3
					hsN(i2)=Abs(cPt(i2)-iPt(i2))
				Next
			End If

			' check if the height is possible
			For i2= 1 To 3
				If Is0(t(i2)) Then
					If Is0(hsN(i2)) Then
						endIt=False
					End If
				ElseIf hsN(i2) < 0.501* t(i2) Then ' wrong point
					endIt=False
				End If
			Next
			i1=i1+1
			If i1>1000 Then
				ErrSub(7,"sub")
				SL_GetDimFromPicks=False
				Pick.PickFaceFromId ( stripName,  stripFaceID )
				Exit Function
			End If
		Wend

	' substrate height is in strip smallestt thickness direction



	  	Bneg=hsN(tMinAx)*2
	  	B=hsP(tMinAx)*2

	Else
	' substrate height is in strip smallestt thickness direction
		Bneg=0
		B=hsP(tMinAx)*2
	End If







' substrate Er
	subMatName=Solid.GetMaterialNameForShape ( posSubName)

	Material.GetEpsilon ( subMatName, EpsX, EpsY, EpsZ )
	Er=EpsX

' strip width
	w=0
	For i1=1 To 3
		If t(i1)>w Then
			w=t(i1)
		End If
	Next

' pick original face
	Pick.PickFaceFromId ( stripName,  stripFaceID )


End Function


Function SL_ConstructPort()

' universal variables
	Dim i1%

' concrete variables
	Dim stripName$
	Dim stripFaceID As Long

	Dim cPt#()
	ReDim cPt(3)
	Dim t#()
	ReDim t(3)

	Dim tMinAx%

	Dim kString$

	Dim B#  ' substrate height
	Dim Bneg#
	Dim Bmax#

	Dim Er#
	Dim w#


	' Range Add Matix
	Dim M$()
	ReDim M(3,2)

	kString=DlgText("TB_k")

' stripName and faceId
	stripName=Pick.GetPickedFaceFromIndex (1, stripFaceID)

' thickness of picked face in all directions
	Pick.PickCenterpointFromId (stripName, stripFaceID)
	Pick.GetPickPointCoordinates (1, cPt(1), cPt(2),cPt(3) )

	GetT(stripName,stripFaceID, cPt, t,tMinAx)

' substrate height
	SL_GetDimFromPicks(t,B,w,Er,Bneg)

	If Is0(Bneg) Then
		Bneg=B
	End If

	Bmax=cdbl(DlgText("TB_dim2"))

' Range Add Matix

	For i1 = 1 To 3
		If Is0(t(i1)) Then ' thickness =0
			M(i1,1)="0"
			M(i1,2)="0"
		ElseIf i1=tMinAx Then
			M(i1,1)=cstr(Bneg/2 - t(i1)/2)
			M(i1,2)=cstr(B/2 - t(i1)/2)
		Else
			M(i1,1)=cstr(Bmax/2) & "*" & kString
			M(i1,2)=cstr(Bmax/2) & "*" & kString
		End If
	Next

' pick original face
	Pick.PickFaceFromId ( stripName,  stripFaceID )

' history add
	AddHist(M)

End Function


