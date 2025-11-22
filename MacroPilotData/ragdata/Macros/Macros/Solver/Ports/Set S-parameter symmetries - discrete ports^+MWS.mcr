' *Solver / Set S-parameter symmetries - discrete ports
' !!! Do not change the line above !!!

' macro.959
' ================================================================================================
' Copyright 2004-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 15-Jul-2007 imu: adapted for version 2008 (AddToHistory)
' 06-May-2005 ube: tiny change in message box, if ports are not subsequentially numbered
' 04-May-2005 imu: allow non-sequentially numbered ports, possibly waveguide ports available
' 10-Feb-2004 imu: First version
'--------------------------------------------------------------------------------------------
' Set automatically S-parameter symmetries for discrete ports
'
' Limitation: Ports only along z in the current version (x and y symmetries)

' TO do:
'   Check what happens if port is placed strangely in the mesh
'   CheckSymmMat
'   Verify for more examples of uneven number of ports on x/y

Option Explicit
Public Cst_Acc As Double

Sub Main

	Cst_Acc = 1.e-4
	Dim xp() As Double, yp() As Double	' Ports x and y grid indices; x, y would be better, but ...
	Dim x() As Double, y() As Double	' different x and y indices
	Dim nodx() As Long, nody() As Long, nodxy() As Long
	Dim portno() As Long, porttree As String, porttreeno As String

	Dim np As Integer	' Number of ports
	Dim nx As Long, ny As Long
	Dim iii As Long, jjj As Long, kkk As Long
	Dim ddd As Long, ix As Long, iy As Long, iz As Long
	Dim stmp As String

	Dim IsSymm As Boolean
	Dim IsSymmX As Boolean, IsSymmY As Boolean, IsSymmD As Boolean
	Dim IsWriteMWS As Boolean

	Dim AdjMat() As Integer

	If Not Solver.ArePortsSubsequentlyNamed Then
		MsgBox "Note, that ports are not subsequently numbered." + vbCrLf + "The symmetry flags will be set, but later Touchstone Export might be denied."
		'Exit All
	End If

	np = Solver.GetNumberOfPorts
	ReDim portno(1 To np)
	kkk = 0
	porttree = ResultTree.GetFirstChildName ( "Ports")
	For iii = 1 To np
		stmp = Right$(porttree, Len(porttree)-Len("Ports\port"))
		If Port.GetType (stmp) = "Discrete" Then
			kkk = kkk+1
			portno(kkk) = Val(stmp)
		End If
		porttree = ResultTree.GetNextItemName (porttree)
	Next iii

	np = kkk	' Number of discrete ports; their numbers stored in "portno"

	ReDim xp(1 To np):	ReDim yp(1 To np)
	ReDim x(1 To np):	ReDim y(1 To np)
	ReDim nodx(1 To np):	ReDim nody(1 To np)
	ReDim nodxy(1 To np, 1 To np)

	BeginHide

	Begin Dialog UserDialog 240,217,"Enter discrete port symmetries" ' %GRID:10,7,1,1
		GroupBox 10,14,220,35,"X-Symmetry (Oy symmetry axis)",.GroupBox1
		OptionGroup .Xsymm
			OptionButton 20,28,90,14,"Yes",.OptionButton1
			OptionButton 110,28,90,14,"No",.OptionButton2
		GroupBox 10,63,220,35,"Y-Symmetry (Ox symmetry axis)",.GroupBox2
		OptionGroup .Ysymm
			OptionButton 20,77,90,14,"Yes",.OptionButton3
			OptionButton 110,77,90,14,"No",.OptionButton4
		GroupBox 10,112,220,35,"Diag. Symmetry (  /  )",.GroupBox3
		OptionGroup .Dsymm
			OptionButton 20,126,90,14,"Yes",.OptionButton5
			OptionButton 110,126,90,14,"No",.OptionButton6
		OKButton 10,161,90,21
		CancelButton 130,161,90,21
		Text 10,196,210,14,"Works only for discrete ports !",.Text2
	End Dialog
	Dim dlg As UserDialog
	dlg.Dsymm = 1

	If (Dialog(dlg) = 0) Then Exit All

	If dlg.Xsymm = 0 Then IsSymmX = True: Else IsSymmX = False
	If dlg.Ysymm = 0 Then IsSymmY = True: Else IsSymmY = False
	If dlg.Dsymm = 0 Then IsSymmD = True: Else IsSymmD = False

	Assign IsSymmX: Assign IsSymmY: Assign IsSymmD:

	EndHide


	For iii = 1 To np
		DiscretePort.GetElementLocation portno(iii), ddd, ix, iy, iz
		xp(iii) = Mesh.GetX (ix)
		yp(iii) = Mesh.GetY (iy)
	Next iii
		'ShowVector(np, xp): ShowVector(np, yp)

	SortXY(np, xp, yp, nx, ny, x, y, nodx, nody) ', nodxy)
		'ShowVector(np, nodx): ShowVector(np, nody)

	IsSymm = CheckPortSymm(nx, ny, x, y, IsSymmX, IsSymmY, IsSymmD)

	Dim matyx() As Long	' 2D Matrix with columns corresponding to x values, rows to y values
	ReDim matyx(ny, nx)
	For iii = 1 To np
		matyx( nody(iii), nodx(iii) ) = iii
	Next iii

	' Display matrix
	'ShowMatrix(ny, nx, matyx)

	IsSymm = CheckMatSymm(nx, ny, matyx)

	ReDim AdjMat(np*(np+1)/2,np)
	BuildAdjacencyMat(AdjMat, np)		' ShowMatrix(np*(np+1)/2,np, AdjMat)
	DetermineSymmetricSPars (IsSymmX, IsSymmY, IsSymmD, AdjMat, nx, ny, np, matyx,portno)

End Sub

' ================================================================
' Function for Sorting
' ================================================================

' ----------------------------------------------------------------

Function SortXY(np As Integer, xp() As Double, yp() As Double, nx As Long, ny As Long, _
					x() As Double, y() As Double, nodx() As Long, nody() As Long)
'					x() As Double, y() As Double, nodx() As Long, nody() As Long, nodxy( ) As Long)
	Dim iii As Long, jjj As Long, tmp As Double
	Dim crt As Long

	Dim xs() As Double, ys() As Double	' Sorted ports x and y indices
	ReDim xs(1 To np):	ReDim ys(1 To np)

	' First initialize sorted x,y    and node indices
	For iii = 1 To np
		xs(iii) = xp(iii)
		ys(iii) = yp(iii)
		'nodx(iii) = iii: nody(iii) = iii
	Next iii

	' Now sort xp and yp, put result in xs, ys
	crt = 1
	While (crt < np)
		For iii = crt+1 To np
			If xp(iii) < xs(crt) Then	' Permute
				tmp = xs(crt): xs(crt) = xs(iii): xs(iii) = tmp
			End If
			If yp(iii) < ys(crt) Then	' Permute
				tmp = ys(crt): ys(crt) = ys(iii): ys(iii) = tmp
			End If
		Next iii
		crt = crt+1
	Wend
'ShowVector(np, xs): ShowVector(np, ys)


	' Finally, find out how many different values of xs there are; put result in x, their number in nx
	x(1) = xs(1): crt = 2
	For iii = 2 To np
		If Abs(xs(iii) - xs(iii-1)) > Cst_Acc * Abs(xs(iii)) Then
			x(crt) = xs(iii)
			crt = crt+1
		End If
	Next iii
	nx = crt -1

	' Similarly, calculate number of different y values
	y(1) = ys(1): crt = 2
	For iii = 2 To np
		If Abs(ys(iii) - ys(iii-1)) > Cst_Acc*Abs(ys(iii)) Then
			y(crt) = ys(iii)
			crt = crt+1
		End If
	Next iii
	ny = crt -1

	' Nodx, nody will contain the index in x and y of the ports' coordinates
	For iii = 1 To np
		For jjj = 1 To nx
			If xp(iii) = x(jjj) Then nodx(iii) = jjj
		Next jjj
		For jjj = 1 To ny
			If yp(iii) = y(jjj) Then nody(iii) = jjj
		Next jjj
	Next iii

End Function

' ================================================================
' Functions for building adjacency matrices
' ================================================================

' ----------------------------------------------------------------
Function BuildAdjacencyMat(AdjMat() As Integer, np As Integer)
	Dim iii As Integer, jjj As Integer, crt As Integer
	crt = 1

	' First the connections of type 12, 13, ...
	For iii = 1 To np-1
		For jjj = iii+1 To np
			AdjMat(crt, iii) = 1
			AdjMat(crt, jjj) = 1
			crt = crt+1
		Next jjj
	Next iii

	' Now the diagonal ones 11, 22, ...
	For iii = 1 To np
		AdjMat(crt, iii) = 1
		crt = crt+1
	Next iii
End Function

Function BuildDMat(nip As Integer, np As Integer, npairs As Integer, AdjMat() As Integer, pairs() As Integer, DMat() As Integer)
	Dim iii As Integer, jjj As Integer

	' Initalize DMat with the AdjMat values
	For iii = 1 To nip
		For jjj = 1 To np
			DMat(iii, jjj) = AdjMat(iii, jjj)
		Next jjj
	Next iii

	For iii = 1 To nip
		For jjj = 1 To npairs
			DMat(iii, pairs(jjj, 1)) = AdjMat(iii, pairs(jjj, 2))
			DMat(iii, pairs(jjj, 2)) = AdjMat(iii, pairs(jjj, 1))
		Next jjj
	Next iii

End Function


' ================================================================
' Main function to determine the symmetric S parameters
' ================================================================

' ----------------------------------------------------------------

Function DetermineSymmetricSPars (IsSymmX As Boolean, IsSymmY As Boolean, IsSymmD As Boolean, _
							AdjMat() As Integer, nx As Long, ny As Long, np As Integer, matyx() As Long, portno() As Long)
	Dim iii As Integer, jjj As Integer, crt As Integer
	Dim tmp1 As Integer, tmp2 As Integer, tmp As Integer
	Dim tmpstr As String

	Dim npairs As Integer, pairs() As Integer
	Dim nip As Integer	' Number of "independent" pairs

	Dim vecYN() As Integer	' Holds following information: has a pair, such as 13, already been considered?
	Dim matSymS () As String	' Holds strings of the type "1,3", necessary in the MWS def. of symmetries
	Dim vecNoSymEntries() As Integer	' Holds number of nonzero entries on the rows of matSymS
	Dim vecPairsStr() As String
	Dim DMat() As Integer


	nip = np*(np+1) / 2
	ReDim vecYN(1 To nip)
	ReDim vecNoSymEntries(1 To nip)
	ReDim matSymS(1 To nip, 1 To 10*nip)		' Columns generously reserved!
	ReDim vecPairsStr(1 To nip)

	' Put names of possible port pairs in vecPairsStr and in first column of matSymS
	crt = 1
	For iii = 1 To np-1
		For jjj = iii+1 To np
			vecPairsStr(crt) = CStr(jjj) + "," + CStr(iii)
			matSymS(crt, 1) = vecPairsStr(crt)
			crt = crt+1
		Next jjj
	Next iii

	For iii = 1 To np
		vecPairsStr(crt) = CStr(iii)+ "," + CStr(iii)
		matSymS(crt, 1) = vecPairsStr(crt)
		crt = crt+1
	Next iii

	' Deal with the various symmetries
	npairs = 3  *(Fix(np/2) +1)	' max 3 symmetries considered
	ReDim pairs(1 To npairs, 1 To 2)

	Dim IsSymm(1 To 3)As Integer
	If IsSymmX Then IsSymm(1) = 1: If IsSymmY Then IsSymm(2) = 1: If IsSymmD Then IsSymm(3) = 1

	' Construct pairs of symmetric ports
	Dim kkk As Integer
	For kkk = 1 To 3
	  If IsSymm(kkk) = 1 Then
		npairs = 0:	FindPairsOfSymmPorts(kkk, ny , nx , matyx, pairs, npairs )	':ShowMatrix(npairs,2, pairs)

		' Now check the Spars symmetry. First, construct a temp. matrix based on AdjMat
		' This trick seems to work only for each symm. separately.
		ReDim DMat (1 To nip, 1 To np)
		BuildDMat(nip, np, npairs, AdjMat, pairs, DMat)							':ShowMatrix(nip,npairs, DMat)
		ConstructSymmetryPairs(nip, npairs, DMat, matSymS, vecPairsStr, np)		':ShowMatrix(nip,nip, matSymS)
	  End If
	Next kkk

	' We're almost done. Just eliminate the doubles now
	EliminateDoublePairs(np, nip, matSymS, vecNoSymEntries )
	' .... and translate the port numbers into the possibly not sequentially numbered
	RenumberPortsInMATSYMS(nip, matSymS,portno)

	' MWS command + write logfile
	WriteSymToMWS(nip, IsSymmX, IsSymmY, IsSymmD, vecNoSymEntries, matSymS)

End Function

' ----------------------------------------------------------------
Function FindPairsOfSymmPorts(symm As Integer, ny As Long, nx As Long, matyx() As Long, pairs() As Integer, npairs As Integer)
	' Construct pairs of symmetric ports

	Dim crt As Integer, iii As Integer, jjj As Integer, tmp As Integer, tmp1 As Integer, tmp2 As Integer

	' For y symmetry,  use nny = ny/2, nnx = nx
	' For x symmetry,  use nny = ny, nnx = nx/2
	Dim nnx As Integer, nny As Integer
	Select Case symm
		Case 2		' "y"
			nny = ny/2: nnx = nx
		Case 1		' "x"
			nny = ny: nnx = nx/2
		Case 3		' "d"
			nny = ny: nnx = nx
	End Select

	crt = npairs + 1
	For iii = 1 To nny
		For jjj = 1 To nnx
			If matyx(iii, jjj) <> 0 Then
				tmp1 = matyx(iii, jjj):
				Select Case symm
					Case 2	'"y"
						tmp2 = matyx(ny-iii+1, jjj)
					Case 1	'"x"
						tmp2 = matyx(iii, nx-jjj+1)
					Case 3	'"d"
						tmp2 = matyx(jjj, iii)
				End Select
				If (tmp1 > tmp2) Then tmp=tmp2: tmp2=tmp1: tmp1 = tmp:
				pairs(crt, 1) = tmp1	' matyx(iii, jjj)
				pairs(crt, 2) = tmp2	' matyx(ny-iii+1, jjj)
				crt = crt+1
			End If
		Next jjj
	Next iii
	npairs = crt-1


End Function

' ----------------------------------------------------------------
Function ConstructSymmetryPairs(nip As Integer, npairs As Integer, DMat() As Integer, matSymS() As String, vecPairsStr() As String, np As Integer)
	Dim iii As Integer, jjj As Integer, IsId As Boolean, crt As Integer
	Dim tmpstr As String, vn(2) As Integer

	' Construct symmetry pairs
	For iii = 1 To nip
		tmpstr = ""
		crt = 0
		For jjj = 1 To np
			If DMat(iii, jjj) <> 0 Then	vn(crt) = jjj: crt = crt+1
		Next jjj
		If (crt = 1) Then vn(1) = vn(0)
		If (vn(0) < vn(1)) Then	jjj = vn(0): vn(0) = vn(1): vn(1) = jjj	' Permute them
		tmpstr = CStr(vn(0)) + "," + CStr(vn(1))

		crt = 2
		While matSymS(iii, crt) <> ""
			crt = crt+1
		Wend
		matSymS(iii, crt) = tmpstr
	Next iii
	'ShowMatrix(nip,nip, matSymS)

End Function

' ----------------------------------------------------------------
Function EliminateDoublePairs(np As Integer, nip As Integer, matSymS() As String, vecNoSymEntries() As Integer)
	Dim iii As Integer, crt As Integer
	Dim tmpstr As String
	Dim tmp As Integer, tmp1 As Integer, tmp2 As Integer

	Dim vecYN() As Integer

'MsgBox "Matrix before line doubles elimination"
'ShowMatrix(nip,nip, matSymS)

	' Concatenate lines which contain the same entries
	For iii = 1 To nip
		crt = 1
		While matSymS(iii, crt) <> ""
			tmp = GetIndexInVector(np, matSymS(iii, crt))
			If tmp <> iii Then ConcatenateLines(iii, tmp, matSymS)
			crt = crt+1
		Wend
	Next iii
'MsgBox "Matrix after line concatenation"
'ShowMatrix(nip,nip, matSymS)

	' Now eliminate doubles on each line
	For iii = 1 To nip		' Rows in matrix matSymS
		ReDim vecYN(1 To nip)
		crt = 1
		While matSymS(iii, crt) <> ""
			tmp = GetIndexInVector(np, matSymS(iii, crt))
			If vecYN(tmp) = 0 Then
				vecYN(tmp) = 1
				crt = crt+1
			Else	' Eliminate this entry from the matrix, move whole row to the left
				If crt = 1 Then
					'matSymS(iii, crt) = ""
				Else
					tmp = crt
					While matSymS(iii, tmp) <> ""
						matSymS(iii, tmp) = matSymS(iii, tmp+1)
						tmp = tmp+1
					Wend
				End If
			End If
		Wend
		'If crt = 1 Then	matSymS(iii,crt) = ""
		vecNoSymEntries(iii) = crt-1
	Next iii

	For iii = 1 To nip
		If vecNoSymEntries(iii) = 1 Then
			matSymS(iii, 1) = ""
			vecNoSymEntries(iii) = 0
		End If
	Next iii
'ShowMatrix(nip,nip, matSymS)
'ShowVector(nip, vecNoSymEntries)

End Function

' ----------------------------------------------------------------
Function GetIndexInVector(np As Integer, tmpstr As String) As Integer
	Dim i As Integer, j As Integer
	Dim tmp As Integer

	'tmpstr is something like  "2,1"

	j = Val(Left$(tmpstr, InStr(tmpstr, ",")-1) )
	i = Val(Right$(tmpstr, Len(tmpstr)-InStr(tmpstr, ",")) )
	If (i <> j) Then
		tmp = (i-1)*np - i*(i-1)/2 + j-i		' Index in vector
	Else
		tmp = np*(np-1)/2+i
	End If
	GetIndexInVector = tmp
End Function

' ----------------------------------------------------------------
Function ConcatenateLines(i As Integer, j As Integer, M() As String)
	' Concatenate Line j to line i,  delete line j
	Dim crti As Integer, crtj As Integer

	crti = 1
	While M(i, crti) <> ""
		crti = crti+1
	Wend

	crtj = 1
	While M(j, crtj) <> ""
		M(i, crti) = M(j, crtj): crti = crti+1
		M(j, crtj) = "": crtj = crtj+1
	Wend
End Function
' ================================================================
' Functions for checking symmetries
' For now, partially used
' ================================================================

' ----------------------------------------------------------------

Function CheckOneSymm(nx As Long, x() As Double) As Boolean
	Dim iii As Long
	Dim xmed As Double, ymed As Double, tmp As Double

	CheckOneSymm = True

	iii = Fix(nx/2)
	If iii*2 = nx Then
		xmed = (x(iii)+x(iii+1))/2:
	Else
		xmed = x(iii+1)
	End If

	For iii = 1 To Fix(nx/2)
		tmp = 2*xmed -x(iii) - x(nx-iii+1)
		If (xmed <> 0) Then
			If Abs(tmp) > Abs(xmed)*Cst_Acc Then
				CheckOneSymm = False
				Exit For
			End If
		Else
			If Abs(tmp) > Cst_Acc Then
				CheckOneSymm = False
				Exit For
			End If
		End If
	Next iii
End Function

' ----------------------------------------------------------------
Function CheckPortSymm(nx As Long, ny As Long, x() As Double, y() As Double, IsSymmX As Boolean, IsSymmY As Boolean, IsSymmD As Boolean) As Boolean
	Dim IsSymm As Boolean

	If IsSymmX = True Then
		IsSymm = CheckOneSymm(nx, x)
		If IsSymm = False Then
			MsgBox "Discrete port arrangement does not appear to be symmetric with respect to Ox!"
			Exit All
		End If
	End If

	If IsSymmY = True Then
		IsSymm = CheckOneSymm(ny, y)
		If IsSymm = False Then
			MsgBox "Discrete port arrangement does not appear to be symmetric with respect to Oy!"
			Exit All
		End If
	End If

	If IsSymmD = True Then		' TO DO: check it!
		IsSymm = CheckOneSymm(ny, y)
		If IsSymm = False Then
			MsgBox "Discrete port arrangement does not appear to be symmetric with respect to the diagonal!"
			Exit All
		End If
	End If

End Function

' ----------------------------------------------------------------
Function CheckMatSymm(nx As Long, ny As Long, matxy() As Long) As Boolean
' Not used for the moment
	CheckMatSymm = True
End Function

' ================================================================
' Functions for displaying matrices and vectors (for debugging)
' ================================================================

' ----------------------------------------------------------------
Function ShowMatrix(nx As Variant, ny As Variant, matr As Variant)
	Dim tmpstr As String, iii As Long, jjj As Long
	tmpstr = ""
	For iii = 1 To nx
		For jjj = 1 To ny
			tmpstr = tmpstr + CStr(matr(iii, jjj)) + " "
		Next jjj
		tmpstr = tmpstr + vbCrLf
	Next iii
	MsgBox tmpstr

End Function

' ----------------------------------------------------------------
Function ShowVector(nx As Variant, matr As Variant)
	Dim tmpstr As String, iii As Long, jjj As Long
	tmpstr = ""
	For iii = 1 To nx
		tmpstr = tmpstr + CStr(matr(iii)) + " "
		tmpstr = tmpstr + vbCrLf
	Next iii
	MsgBox tmpstr

End Function

Function WriteSymToMWS(nip As Integer, IsSymmX As Boolean, IsSymmY As Boolean, IsSymmD As Boolean, _
						vecNoSymEntries() As Integer, matSymS() As String)
		' MWS command + write logfile
	Dim tmpstr1 As String, tmpstr As String, tmpstr_header As String
	Dim iii As Integer, crt As Integer

	If Left(GetApplicationVersion, 9) = "Version 4" Then
		Open GetProjectBaseName+ "~SPar_Symm_Macro.log" For Output As #1
	Else
		Open GetProjectBaseName + "^SPar_Symm_Macro.log" For Output As #1
	End If
	tmpstr=""
	If IsSymmX Then tmpstr = tmpstr+"X=True   " : Else tmpstr=tmpstr+"X=False   "
	If IsSymmY Then tmpstr = tmpstr+"Y=True   " : Else tmpstr=tmpstr+"Y=False   "
	If IsSymmD Then tmpstr = tmpstr+"D=True   " : Else tmpstr=tmpstr+"D=False   "
	Print #1, "Symmetries:  " + tmpstr + vbCrLf
	Print #1, "Identified S-parameter symmetries"
	tmpstr_header = "define solver s-parameter symmetries" '+ vbCrLf+vbCrLf+"With Solver" + vbCrLf
	tmpstr = "With Solver" + vbCrLf

	With Solver
    	.ResetSParaSymm
    						tmpstr = tmpstr + " .ResetSParaSymm" + vbCrLf
    	For iii = 1 To nip
    		If vecNoSymEntries (iii) <> 0 Then
				.DefSParaSymm
							tmpstr = tmpstr + " .DefSParaSymm" + vbCrLf

    			crt = 1
    			While matSymS(iii, crt) <> ""
					.SPara matSymS(iii, crt)
							tmpstr1 = tmpstr1 +	matSymS(iii, crt) + "  "
							tmpstr = tmpstr + "   .SPara """ + matSymS(iii, crt) + """" + vbCrLf
					crt = crt+1
				Wend
							Print #1, tmpstr1
							tmpstr1 = ""
			End If
		Next iii
		.SparaSymmetry "True"
							tmpstr = tmpstr + ".SparaSymmetry ""True""" + vbCrLf
	End With
							tmpstr = tmpstr + "End With" + vbCrLf

	Print #1, vbCrLf + vbCrLf + "MWS .mod-file command" + vbCrLf
	Print #1, tmpstr

	Close #1

	' Write in .mod file
		Save 	'save the mod-File
 		Wait 5

 		AddToHistory (tmpstr_header, tmpstr)

		'Open GetProjectBaseName+ ".mod" For Append As #1
		'Kill GetProjectBaseName+ ".sat"
		'Print #1, tmpstr
		'Close #1

		MsgBox "Port symmetries have been successfully generated." '+vbCrLf+"Please be patient while MWS rebuilds the history list."+ _
	   		'vbCrLf+"This operation might take some time.",vbExclamation

		' Openfile GetProjectBaseName+ ".mod"		'reopen the mod-file

End Function

Function RenumberPortsInMATSYMS(nip As Integer, matSymS() As String, portno() As Long)
	Dim iii As Long, crt As Long
	Dim smat As String, sno1 As String, sno2 As String
	Dim no1 As Long, no2 As Long

	For iii = 1 To nip
		crt = 1
    	While matSymS(iii, crt) <> ""
			smat = matSymS(iii, crt)
			sno1 = Left$(smat, InStr(smat, ",")-1): no1 = Val(sno1)
			sno2 = Right$(smat, Len(smat)- InStr(smat, ",")): no2 = Val(sno2)
			matSymS(iii, crt) = Str(portno(no1)) + "," + Str(portno(no2))
			crt = crt+1
		Wend
	Next iii
End Function
