'#Language "WWB-COM"
' ================================================================================================
' This macro creates Design of Experiments points for CST parametersweep object
' Users need to open the optimizer and select the variables to use, and define their limits, as if they were going to perfom an optimization
' This macro reads the Optimizer settings in the optimizer and use them as input for DOE
' Main references:
'   Joseph, V. R, Hung, Y, "ORTHOGONAL-MAXIMIN LATIN HYPERCUBE DESIGNS" Statistica Sinica 18(2008), 171-186
'   Morris, M. D. and Mitchell, T. J. (1995). Exploratory designs for computer experiments. J. Statist. Plann. Inference 43, 381-402
'	R. Iman, M. J. Shortencarier "A FORTRAN 77 Program and User’s Guide for the Generation of Latin Hypercube and Random Sampies For Use With Computer Modeis"
'	T. Wong, W. Luk, P. Heng "Sampling with Hammersley and Halton Points"
' ================================================================================================
' Copyright 2022-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 09-Jan-2025 wxu: minor dialog improvements, include warning with more instructions
' 14-Jul-2023 wxu: fixed a small issue when all rho = 0
' 04-Jul-2023 wxu: added optimzed option for LHS, based on Joseph and Hung's paper
' 06-Jun-2023 wxu: minor changes on the text box size and macro title
' 02-Jun-2023 wxu: Implemented "Random", "Center" options for LHS, and "Standard", "Halton" options for Hammersley method
' 30-May-2023 wxu: implemented Hammersley sampling method
' 23-May-2023 wxu: initial version: only include Latin Hypercube sampling (LHS) method
' ================================================================================================

Option Explicit
Public Const bDebug = False
Public n As Integer, k As Integer 'n: number of levels/points, k: number of variables
Public w As Double, Imax_total As Long, max_time As Long
Public LL As Variant, HH As Variant 'Lower and upper limits of each variable
Public DOEPoints() As Double
Public sParaNames() As String
Public ListP() As String, ListMin() As String, ListMax() As String, ListSettings1() As Variant, ListSettings2() As Variant

Sub Main

	ListSettings1 = Array("Optimal","Random", "Center")
	ListSettings2 = Array("Standard", "Halton")

	Begin Dialog UserDialog 380,420,"Create DOE Points",.DiagFuncDoE ' %GRID:10,7,1,1
		Text 10,7,360,28,"Use Optimizer to select parameters and set their ranges",.tInfo
		GroupBox 10,42,370,154,"Review parameter settings",.GroupBox1

		Text 30,56,80,14,"Parameter",.Text3,2
		Text 190,56,50,14,"Min",.Text4,2
		Text 310,56,40,14,"Max",.Text5,2

		MultiListBox 20,77,130,91,ListP(),.PList
		MultiListBox 180,77,80,91,ListMin(),.MultiListBox2
		MultiListBox 290,77,80,91,ListMax(),.MultiListBox3
		CheckBox 20,175,250,14," Plot selected (at least 2) parameters",.cbPlot

		GroupBox 10,210,370,112,"DOE Settings",.GroupBox2
		Text 30,231,140,14,"Number of samples",.Text2
		TextBox 200,231,140,21,.tbNumOfSamples
		OptionGroup .DOEOptions
			OptionButton 20,266,140,14,"Latin Hypercube",.LHS
			OptionButton 20,294,140,14,"Hammersley",.Hammersley

		PushButton 340,266,30,21,"...",.pbOpt

		DropListBox 200,266,130,14,ListSettings1(),.dSettings1
		DropListBox 200,294,130,14,ListSettings2(),.dSettings2

		CheckBox 10,336,300,14," Remove previous sequences",.cbDelSeq
		CheckBox 10,357,300,14," Remove previous plots",.cbDelPlots

		PushButton 10,385,90,21,"Preview",.pbPreview
		PushButton 140,382,100,28,"Create",.pbCreate
		CancelButton 280,385,90,21
	End Dialog
	Dim dlg As UserDialog

	If Not Dialog(dlg) Then
		Exit All
	End If

End Sub
Private Function DiagFuncDoE(DlgItem$, Action%, SuppValue?) As Boolean
Dim ii As Integer, jj As Integer
	Select Case Action%
	Case 1 ' Dialog box initialization

		k = Optimizer.GetNumberOfVaryingParameters
		If k = 0 Then
			'DlgText("tInfo","No parameter has been selected" + vbNewLine + "Please use Optimizer to select parameters and set ranges")
			Dim warning_info As String
			warning_info = "No parameter has been selected" + vbNewLine + "Please use CST Optimizer to select parameters and set their ranges, then re-run this macro."

			MsgBox (warning_info,vbCritical,"No parameter(s) selected")
			Exit All
			'DlgEnable("cbPlot",False)
			'DlgEnable("PList", True)
			'DlgEnable("MultiListBox2", False)
			'DlgEnable("MultiListBox3", False)
			'DlgEnable("tbNumOfSamples",False)
			'DlgEnable("DOEOptions",False)
			'DlgEnable("pbOpt",False)
			'DlgEnable("dSettings1",False)
			'DlgEnable("dSettings2",False)
			'DlgEnable("cbDelSeq",False)
			'DlgEnable("cbDelPlots",False)
			'DlgEnable("pbPreview", False)
			'DlgEnable("pbCreate", False)

		Else
			ReDim LL(k-1), HH(k-1), sParaNames(k-1)
			ReDim ListP(k-1), ListMin(k-1), ListMax(k-1)

			For jj = 0 To k-1
				LL(jj) = Optimizer.GetParameterMinOfVaryingParameter(jj)
				HH(jj) = Optimizer.GetParameterMaxOfVaryingParameter(jj)
				sParaNames(jj) = Optimizer.GetNameOfVaryingParameter(jj)
				ListP(jj) = sParaNames(jj)
				ListMin(jj) = cstr(LL(jj))
				ListMax(jj) = cstr(HH(jj))
			Next

			DlgValue("dSettings1",0)
			DlgEnable("dSettings2",False)
			DlgValue("cbPlot",False)
			DlgEnable("cbPlot",False)
			DlgValue("cbDelSeq",True)
			DlgValue("cbDelPlots",True)
			DlgListBoxArray ("PList",ListP)
			DlgListBoxArray ("MultiListBox2",ListMin)
			DlgListBoxArray ("MultiListBox3",ListMax)
			DlgEnable("PList", True)
			DlgEnable("MultiListBox2", False)
			DlgEnable("MultiListBox3", False)
			DlgText("tbNumOfSamples","25")
			DlgValue("DOEOptions",0)
			DlgEnable("pbPreview", False)
			'LHD optimization user parameters
				w = 0.5
				n = Cint(DlgText("tbNumOfSamples"))
				Imax_total = 10 * n * (n-1)/2 * k 'recommended by Morris & Michell
				'Imax_total = 1000
				max_time = 5
		End If
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "PList"
				Select Case UBound(DlgValue("Plist"))
					Case -1, 0
						DlgValue("cbPlot",False)
						DlgEnable("cbPlot",False)
					Case Else
						DlgValue("cbPlot",True)
						DlgEnable("cbPlot",True)
				End Select
			Case "tbNumOfSamples"
				DiagFuncDoE = True

			Case "DOEOptions"
				If DlgValue("DOEOptions") = 0 Then
					DlgEnable("dSettings1",True)
					DlgEnable("dSettings2",False)
					If DlgValue("dSettings1") = 0 Then DlgEnable("pbOpt",True)
				Else
					DlgEnable("dSettings1",False)
					DlgEnable("dSettings2",True)
					DlgEnable("pbOpt",False)

				End If
			Case "dSettings1"
					If DlgValue("dSettings1") = 0 Then
						DlgEnable("pbPreview", False)
						DlgEnable("pbOpt",True)
					Else
						DlgEnable("pbPreview", True)
						DlgEnable("pbOpt",False)
					End If
			Case "pbPreview"
				n = Cint(DlgText("tbNumOfSamples"))
				ReDim DOEPoints(n-1, k-1)

				If DlgValue("DOEOptions") = 0 Then
					DOEPoints = LHS(n,k, ListSettings1(DlgValue("dSettings1")))
				Else
					DOEPoints = Hammersley(n,k,ListSettings2(DlgValue("dSettings2")))
				End If

				For jj = 0 To k-1
					For ii = 0 To n-1
						DOEPoints(ii,jj) = LL(jj) + (HH(jj)-LL(jj))*DOEPoints(ii,jj)
					Next
				Next

				Display2DNumArray(DOEPoints)

				DiagFuncDoE = True
			Case "pbOpt"
				'need to revise to remember dialog settings...-----------------


				Begin Dialog UserDialog 370,252,"Latin Hypercube Optimization Settings",.DiagLHDOpt ' %GRID:10,7,1,1
					Text 20,14,340,35,"The optimization will try to find an orthogonal-Maximin Latin Hypercube Design",.Text2
					GroupBox 20,56,340,70,"Weighting factors for",.GroupBox2
					TextBox 190,70,120,21,.tbW
					TextBox 190,98,120,21,.tbOrth
					Text 70,77,90,14,"Maximin",.Text1
					Text 70,98,90,14,"Orthogonality",.Text3

					GroupBox 20,133,340,84,"Limit Optimization by",.GroupBox1
					CheckBox 50,154,120,21," Max Iterations",.ckIter
					CheckBox 50,182,120,14," Max Time",.ckTime
					TextBox 190,182,120,21,.tbTime
					TextBox 190,154,120,21,.tbIter
					Text 320,184,30,14,"min",.tUnit
					OKButton 70,224,100,21
					CancelButton 190,224,100,21
				End Dialog
				Dim dlg1 As UserDialog

				If Dialog(dlg1) Then
					w = CDbl(dlg1.tbW)
					Imax_total = CLng(dlg1.tbIter)
					max_time = CLng(dlg1.tbTime)
				End If

			 DiagFuncDoE = True ' Prevent button press from closing the dialog box

			Case "pbCreate"

				CreateDOEPoints()

			 DiagFuncDoE = False ' Close dialog
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DiagFuncLHS = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

'Controls the advanced optimization settings
Private Function DiagLHDOpt(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText("tbW",cstr(w))
		DlgText("tbOrth", cstr(1-w))
		DlgEnable("tbOrth",False)
		DlgText("tbIter",cstr(Imax_total))
		DlgText("tbTime", cstr(max_time))
		DlgEnable("ckIter",False)
		DlgValue("ckIter",1)
		DlgValue("ckTime",1)
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "ckTime"
				If DlgValue("ckTime") = 0 Then
					DlgText("tbTime","0")
					DlgEnable("tbTime",False)
				Else
					DlgEnable("tbTime",True)
					DlgText("tbTime", cstr(max_time))
				End If
		End Select
		Rem DiagLHDOpt = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem
			Case "tbW"
				w = Cdbl(DlgText("tbW"))
				If w >1 Then w = 1.0
				DlgText("tbOrth", cstr(1-w))
		End Select
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : LHDOpt = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


Sub CreateDOEPoints()

Dim sSeq As String, sPrefix As String, sPoint As String, ii As Integer, jj As Integer

If DlgValue("cbDelSeq") Then ParameterSweep.DeleteAllSequences 'remove all previous sequences to prevent duplicated names later

If DlgValue("cbDelPlots") And Resulttree.DoesTreeItemExist("1D Results\DOEPlots") Then

	Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long

	nResults = Resulttree.GetTreeResults("1D Results\DOEPlots","0D/1D recursive","",paths,types,files,info)

	For ii = 0 To nResults-1
		With Resulttree
			.Name paths(ii)
			.Delete
		End With
		' ReportInformationToWindow(paths(ii) + " is removed" )
	Next
End If

n = Cint(DlgText("tbNumOfSamples"))
'ReDim DOEPoints(n-1, k-1)

If DlgValue("DOEOptions") = 0 Then
	DOEPoints = LHS(n,k, ListSettings1(DlgValue("dSettings1")))
	sPrefix = "DOE-LHS-"
	If ListSettings1(DlgValue("dSettings1")) = "Random" Then
		sPrefix = sPrefix + "RND"
	ElseIf ListSettings1(DlgValue("dSettings1")) = "Center" Then
		sPrefix = sPrefix + "CTR"
	Else
		sPrefix = sPrefix + "OPT"
	End If
Else
	DOEPoints = Hammersley(n,k,ListSettings2(DlgValue("dSettings2")))
	sPrefix = "DOE-Hammersley-"
	If ListSettings2(DlgValue("dSettings2")) = "Standard" Then
		sPrefix = sPrefix + "STD"
	Else
		sPrefix = sPrefix + "HTN"
	End If

End If

sPrefix = sPrefix + cstr(Int(Rnd()*1000))

For jj = 0 To k-1
	For ii = 0 To n-1
		DOEPoints(ii,jj) = LL(jj) + (HH(jj)-LL(jj))*DOEPoints(ii,jj)
	Next
Next

For ii = 0 To n-1
	sSeq =  sPrefix + "-" + cstr(ii)
	ParameterSweep.AddSequence(sSeq)
	sPoint = ""
	For jj = 0 To k-1
		ParameterSweep.AddParameter_ArbitraryPoints(sSeq,sParaNames(jj),cstr(DOEPoints(ii,jj)))
	Next
Next

'Plot DOE points in a XY graph
If DlgValue("cbPlot") Then
	Dim iSelectedParaIndex() As Integer, nSelecedtParaToPlot As Integer, kk As Integer

	nSelecedtParaToPlot = UBound(DlgValue("PList"))
	ReDim iSelectedParaIndex( nSelecedtParaToPlot )

	For ii = 0 To nSelecedtParaToPlot
		iSelectedParaIndex(ii) = DlgValue("PList")(ii)
	Next

	For kk = 0 To nSelecedtParaToPlot -1
		For jj = kk + 1 To nSelecedtParaToPlot
			Dim DOEPlot As Object, xLabel As String, yLabel As String
			Set DOEPlot = Result1D("")

			xLabel = ListP(iSelectedParaIndex(kk))
			yLabel = ListP(iSelectedParaIndex(jj))
			For ii = 0 To n-1
				DOEPlot.Appendxy(DOEPoints(ii,iSelectedParaIndex(kk)), DOEPoints(ii,iSelectedParaIndex(jj)))
			Next
			DOEPlot.xlabel( xLabel)
			DOEPlot.ylabel( yLabel)
			DOEPlot.Save(xLabel + "-" + yLabel + ".sig")
			DOEPlot.AddToTree("1D Results\DOEPlots\" + xLabel + "-" + yLabel)
			SelectTreeItem("1D Results\DOEPlots\" + xLabel + "-" + yLabel)
			With Plot1D
				.SetMarkerStyle(0, "marksonly", "circles", 6 )
				.XRange(LL(kk),HH(kk))
				.YRange(LL(jj),HH(jj))
				.XTicksDistance((HH(kk)-LL(kk))/n)
				.YTicksDistance((HH(jj)-LL(jj))/n)
				.Plot
			End With
		Next
	Next
End If

MsgBox ("Please open parameter sweep dialog to inspect newly created DOE sequences, which" + _
		"starts with:   " + sPrefix, vbOkOnly)

End Sub

Function LHS(n As Integer,k As Integer,  sOption As String) As Double() ' n - number of levels/points, k - number of variables/dimensions

	Dim ii As Integer, jj As Integer
	Dim aTmp() As Integer, aN() As Integer, aSample As Double
	Dim aInterval() As Double, SamplePoints() As Double, LHSDOEPoints() As Double

	ReDim  aTmp(n-1), aInterval(n-1,k-1), aN(n-1, k-1), SamplePoints(n-1,k-1), LHSDOEPoints(n-1,k-1)

	For jj = 0 To k -1
		For ii = 0 To n-1
			aInterval (ii,jj) = 1/n 	'for future non-uniform interval enhancements
		Next
	Next

	If sOption = "Random" Then
		aSample = Rnd()
	ElseIf sOption = "Center" Then
		aSample = 0.5
	ElseIf sOption = "Optimal" Then 'optimal design start from centered option
		aSample = 0.5
	Else
		aSample = Rnd() 'placeholder for future options
	End If

	For jj = 0 To k-1
		For ii = 0 To n-1
			SamplePoints(ii,jj) = (ii + aSample )* aInterval(ii,jj)
		Next
	Next

	If bDebug Then	Display2DNumArray(SamplePoints)

	For ii = 0 To n-1
		'aTmp(ii) = n-1 - ii
		aTmp(ii) = ii
	Next

' standard LHS
	For jj = 0 To k-1
		If jj <>0 Then Shuffle(aTmp) 'skip the first column
		For ii = 0 To n-1
			aN(ii,jj) = aTmp(ii)
		Next
	Next

'	optimization process starts below
	If sOption = "Optimal" Then
		aN = OptimizeLHS(aN)
	End If

	If bDebug Then	Display2DNumArray(aN)

	For jj = 0 To k-1
		For ii = 0 To n-1
			LHSDOEPoints(ii,jj) = SamplePoints(aN(ii,jj),jj)
		Next
	Next

	If bDebug Then	Display2DNumArray(LHSDOEPoints)
	LHS = LHSDOEPoints
End Function


Function OptimizeLHS(LHD As Variant) As Integer() 'LHD is a nxk matrix
'This function provides ORTHOGONAL-MAXIMIN optimzed LHS matrix

Dim D_current As Variant, D_try As Variant, D_best As Variant
Dim Psai_D_try As Double, Psai_D_current As Double, Psai_best As Double
Dim phi_p As Double, rho_p As Double, phi_L As Double, phi_U As Double, davg As Double
Dim t As Double, Imax_loop As Long, FAC As Double
Dim Improved As Boolean, p As Integer, alpha As Integer

Dim i_select As Integer, j_select As Integer, source As Integer, i_target As Integer

Dim dist() As Double, rho () As Double
Dim ii As Long, jj As Long, ss As Long, tt As Long
Dim Prob_col() As Double, Prob_row() As Double

ReDim dist(n-1,n-1), rho (k-1, k-1)
ReDim Prob_col(k-1), Prob_row(n-1)

' ---- additional phi_p based optimization parameters
p = 15
Imax_loop = 100

t = CalculateT0(p)
FAC = 0.9

'----------- calculate lower and upper bound of phi_p ------------------------
davg = (n+1)*k/3.0  'Lemma 1 in Joseph and Hung's paper, when distrance is L1
phi_U = 0
For ii = 1 To n-1
'	phi_U += (n-ii)*(1/(ii*k/n))^p
	phi_U += (n-ii)*(1/(ii*k))^p
Next
phi_U = phi_U^(1/p)

'calculate phi lower bound
If Int(davg) = davg Then
	phi_L = 0
Else
	phi_L = (Int(davg + 1) - davg)/(Int(davg))^p + (davg -Int(davg))/(Int(davg +1))^p
	phi_L = (n*(n-1)/2*(phi_L))^(1/p)
End If
'--------- phi_p bound calculation ends --------------------------------------

D_current = LHD
D_try = LHD
D_best = LHD
rho = Calc_Corr_Matrix(LHD)
dist = Calc_Dist_Matrix(LHD)

Psai_D_current = Calc_Psai(dist,rho, p, w, phi_p,rho_p, phi_L, phi_U)
Psai_best = Psai_D_current
DlgText("tinfo","LHS optimization starts, this may take a while...")
DlgText("tinfo","Initial Psai = " + Format(Psai_best, "0.0000") + " phi_p/rho_p: " + Format(phi_p,"0.0000") + "/" + Format(rho_p,"0.0000"))

If bDebug Then ReportInformation("Lower and upper limit of phi_p: " + Format(phi_L,"0.0000") + "/" + Format(phi_U,"0.0000"))

If bDebug Then Display2DNumArray(D_best)

Dim i_total As Long, t_start As Double, t_end As Double  'total iteration and timecounter
i_total = 0
t_start = Timer

Do
	Improved = False
	ii = 1

	While (ii < Imax_loop) And (i_total < Imax_total)
		D_try = D_current

		rho = Calc_Corr_Matrix(D_try)
		dist = Calc_Dist_Matrix(D_try)
		Psai_D_current = Calc_Psai(dist, rho, p, w, phi_p, rho_p, phi_L, phi_U)

		'calculate probability based on the correlation and distance
		alpha = 10
		Prob_col = Calc_Corr_Prob(rho, alpha)
		Prob_row = Calc_Dist_Prob(dist,p, alpha)

		'select column and row based on probablities
		j_select = Weighted_RandomChoice(Prob_col)
		i_select = Weighted_RandomChoice(Prob_row)
		'switch selected row with a random row within the selected column
		source = D_try(i_select,j_select)

		i_target = Int(n*Rnd())
		While i_target = i_select
			i_target = Int(n*Rnd())
		Wend

		D_try(i_select,j_select) = D_try(i_target,j_select)
		D_try(i_target, j_select) = source

		'update rho and dist matrix
		Update_Corr_Matrix(D_try, rho, j_select)
		Update_Dist_Matrix(D_try, dist, i_select, i_target)
		Psai_D_try = Calc_Psai(dist, rho, p, w, phi_p, rho_p, phi_L, phi_U)

		If 	Psai_D_try = -1 Then
			Display2DNumArray(D_try)
			Exit All
		End If

		If Psai_D_try < Psai_D_current Or (Rnd() < Exp(-(Psai_D_try - Psai_D_current)/t)) Then
			D_current = D_try
			Improved = True
		End If

		If Psai_D_try < Psai_best Then 'better design found
			D_best = D_try
			Psai_best = Psai_D_try
			ii = 1
			DlgText("tinfo", "Better design found: " + vbNewLine + "Psai = " + Format(Psai_best, "0.0000") + "  phi_p/rho_p= " + Format(phi_p,"0.0000") + "/" + Format(rho_p,"0.0000"))
			ReportInformation("Better design found: Psai = " + Format(Psai_best, "0.0000") + " phi_p/rho_p: " + Format(phi_p,"0.0000") + "/" + Format(rho_p,"0.0000"))
			If bDebug Then Display2DNumArray(D_best)
		Else
			ii += 1
		End If
		i_total +=1
	Wend
	DlgText("tinfo", "Working hard...iteration " + cstr(i_total) +" of " + cstr(Imax_total) + ", " + Format((i_total/Imax_total),"Percent" )+ " done...")

	If Improved Then t = t*FAC
	If bDebug Then ReportInformation("Current t= " + Format(t, "0.000E+00"))

	t_end = Timer
Loop While (Improved = True And i_total < Imax_total) And ((t_end - t_start) < max_time*60 )

DlgText("tinfo", "Optimization finished! Optimized objective function value =" + Format(Psai_best, "0.0000"))
rho = Calc_Corr_Matrix(D_best)
dist = Calc_Dist_Matrix(D_best)
Psai_best = Calc_Psai(dist, rho, p, w, phi_p, rho_p, phi_L, phi_U)

ReportInformation("Optimization finished! Optimized objective function value =" + Format(Psai_best, "0.0000") + " phi_p/rho_p: " + Format(phi_p,"0.0000") + "/" + Format(rho_p,"0.0000"))
ReportInformation("Total number of iterations: " + cstr(i_total) + ", total elapsed time: " + cstr(Int((t_end-t_start))) + "s")

If bDebug Then Display2DNumArray(D_best)

OptimizeLHS = D_best

End Function

Sub Update_Dist_Matrix(D() As Integer, dist() As Double, i_select As Integer, i_target As Integer)
Dim row_s() As Integer, row_t() As Integer, ii As Long
ReDim row_s(k-1), row_t(k-1)

row_s = row(D, i_select)

For ii = 0 To n-1
	row_t = row(D, ii)
	If ii = i_select Then
		dist(i_select, ii) = 0
	Else
		dist(i_select,ii) = d_st(row_s, row_t,1)
		dist(ii,i_select) = dist(i_select,ii)
	End If
Next

row_s = row(D, i_target)
For ii = 0 To n-1
	row_t = row(D, ii)
	If ii = i_target Then
		dist(i_target, ii) = 0
	Else
		dist(i_target,ii) = d_st(row_s, row_t,1)
		dist(ii,i_target) = dist(i_target,ii)
	End If
Next

End Sub

Sub Update_Corr_Matrix(D() As Integer, rho() As Double, j_select As Integer)
Dim col_L1() As Integer, col_L2() As Integer, jj As Long
ReDim col_L1(n-1), col_L2(n-1)

col_L1 = col(D, j_select)
For jj = 0 To k - 1
	If jj = j_select Then
		rho(jj,jj) = 1.0 'self correlation = 1
	Else
		col_L2 = col(D,jj)
		rho(j_select,jj) = Rho_ij(col_L1,col_L2)
		rho(jj,j_select) = rho(j_select,jj)
	End If
Next

End Sub

Function Calc_Dist_Prob(dist, p As Integer, alpha As Integer) As Double()

' this function calculates the probablity of the rows based on its distance, if the distance is small, the probability is higher
' the probability values will be used later to pick which row to shuffle

Dim ii As Long, jj As Long
Dim phi_row() As Double, P_row() As Double, sum_phi_row As Double

ReDim phi_row(n-1), P_row(n-1)

sum_phi_row = 0
For ii = 0 To n-1
	phi_row(ii) = 0
	For jj = 0 To n-1
		If ii <> jj Then phi_row(ii) += (1/dist(ii,jj))^p
	Next
	phi_row(ii) = (phi_row(ii))^(alpha/p)
	sum_phi_row += phi_row(ii)
Next

For ii = 0 To n - 1
	P_row(ii) = phi_row(ii)/sum_phi_row
Next

Calc_Dist_Prob = P_row

End Function

Function Calc_Corr_Prob(rho() As Double, alpha As Integer) As Double()

' this function calculate each variable/column's pick probability based on their correlation to
' other variables, later this probability values will be used to determine which column to pick for shuffling.

Dim ii As Long, jj As Long
Dim rho_col() As Double, P_col() As Double, sum_rho_col As Double
ReDim rho_col(k-1), P_col(k-1)

sum_rho_col = 0
For ii = 1 To k-1
	rho_col(ii) = 0
	For jj = 0 To k-1
		If ii <> jj Then rho_col(ii) += rho(ii,jj)^2
	Next
	rho_col(ii) = (Sqr(rho_col(ii)/(k-1)))^alpha
	sum_rho_col += rho_col(ii)
Next

If sum_rho_col = 0.0 Then 'if all rows are orthoganal already, sum of rho = 0, in this case all columns should have equal probabilities
  For ii = 1 To k-1
  	P_col(ii) = 1/(k-1)
  Next
Else
	For ii = 1 To k-1
		P_col(ii) = rho_col(ii)/sum_rho_col 'probability of each rhoL
	Next
End If

Calc_Corr_Prob = P_col

End Function


Function Calc_Psai(dist() As Double,rho() As Double, p As Integer, w As Double, _
				   phi_p As Double, rho_p As Double, phi_L As Double, phi_U As Double) As Double 'objective function psai -- weighted objective of maxmin and orthogonality

'this function calculates the objective function

Dim ii As Long, jj As Long, ss As Long, tt As Long

phi_p = 0
rho_p = 0

For ss = 0 To n -2
	For tt = ss + 1 To n-1
		phi_p += (1/dist(ss,tt))^p
	Next
Next

phi_p = phi_p^(1/p)

For ii = 0 To k-2
	For jj = ii + 1 To k-1
		rho_p += rho(ii,jj)^2
	Next
Next

rho_p = rho_p/(k*(k-1)/2)

If (phi_p - phi_L)/(phi_U - phi_L) < 0 Then
	ReportInformation("Negative Maximin criteria detected, please contact wxu1@3ds.com")
	ReportInformation("phi_p :" + Format(phi_p,"0.0000"))
	ReportInformation("phi_L :" + Format(phi_L,"0.0000"))
	ReportInformation("phi_U :" + Format(phi_U,"0.0000"))
	Calc_Psai = -1
	Exit Function
End If

Calc_Psai = w * (phi_p - phi_L)/(phi_U - phi_L) + (1 - w) * rho_p
End Function

Function Calc_Dist_Matrix(D() As Integer) As Double()

'this function calculate distance between each levels, first index is the 'From" row, second index is the "To" row

Dim ss As Long, tt As Long, dist() As Double, row_s() As Integer, row_t() As Integer
ReDim row_s(k-1), row_t(k-1)
ReDim dist(n-1,n-1)

For ss = 0 To n - 2
	row_s = row(D,ss)
	dist(ss,ss) = 0.0 'self distance = 0
	For tt = ss + 1 To n -1
		row_t = row(D,tt)
		dist(ss,tt) = d_st(row_s,row_t,1)
		dist(tt,ss) = dist(ss,tt) 'symmetry
	Next
Next

Calc_Dist_Matrix = dist

End Function

Function Calc_Corr_Matrix(D() As Integer) As Double()

' this function calculates linear correlations btw two variables

Dim ii As Long, jj As Long, rho() As Double
Dim col_L1() As Integer, col_L2() As Integer

ReDim rho(k-1, k-1), col_L1(n-1), col_L2(n-1)

For ii = 0 To k -2
	col_L1 = col(D, ii)
	rho(ii,ii) = 1.0
	For jj = ii + 1 To k - 1
		col_L2 = col(D,jj)
		rho(ii,jj) = Rho_ij(col_L1,col_L2)
		rho(jj,ii) = rho(ii,jj)
	Next
Next

Calc_Corr_Matrix = rho

End Function

Function CalculateT0(p As Integer) As Double
	Dim davg As Double, D As Double, Delta As Double, sum1 As Double, sum2 As Double, Cn2 As Long
	Dim ii As Long

	davg = (n+1)*k/3.0 'Lemma 1 in Joseph and Hung's paper, when distrance is L1

	Delta = 1	' use integer distance, instead of "real distance" to simplify code
	Cn2 = n*(n-1)/2
	sum1 = 0

	'Following Morris and Michell's paper: a hyperthotical distribution of distance 50% - 150% of average distance
	For ii = 1 To Cn2 -1
		D = davg/2 + davg/(Cn2-1)*ii
		sum1 += (1/D)^p
	Next
	sum2 = sum1
	sum1 = (sum1 + (1/2/davg)^p)^(1/p)
	sum2 = (sum2 + (1/(0.5*davg - Delta))^p)^(1/p)

	CalculateT0 = (sum2-sum1)*99.5 'ln(1/0.99) ~= 99.5

End Function

Function col(matrix As Variant, index As Integer) As Variant
Dim ii As Long, tmp() As Integer

ReDim tmp(LBound(matrix,1) To UBound(matrix,1))

For ii = LBound(matrix,1) To UBound(matrix,1)
	tmp(ii) = matrix(ii,index)
Next

col = tmp
End Function
Function row(matrix As Variant, index As Integer) As Variant
Dim jj As Long, tmp() As Integer

ReDim tmp(LBound(matrix,2) To UBound(matrix,2))
For jj = LBound(matrix,2) To UBound(matrix,2)
	tmp(jj) = matrix(index,jj)
Next

row = tmp
End Function

Sub Shuffle(arr)
	Dim ii As Integer, jj As Integer, n As Integer, tmp As Double

	For ii = 0 To UBound(arr)
		jj = Int(ii * Rnd())
		tmp = arr(ii)
		arr(ii) = arr(jj)
		arr(jj) = tmp
	Next
End Sub

Function Rho_ij(X As Variant, Y As Variant) As Double
'This function returns Linear correlation of X and Y vectors

Dim num As Long, ii As Long
Dim xavg As Double, yavg As Double, sumXsq As Double, sumYsq As Double, sumXY As Double

num = UBound(X)
If num <> UBound(Y) Or num <> n-1 Then
	ReportInformation("Correlation calculation error: the length of the ij vector is not correct")
	Rho_ij = 1
	Exit Function
End If

xavg = (n-1)/2
yavg = xavg

sumXsq = 0
sumYsq = 0
sumXY = 0
For ii = 0 To n -1
	sumXY += (X(ii)-xavg)*(Y(ii)-yavg)
	sumXsq += (X(ii)-xavg)^2
	sumYsq += (Y(ii)-yavg)^2
Next

Rho_ij = sumXY/Sqr(sumXsq*sumYsq)

End Function
Function d_st(X As Variant, Y As Variant, nType As Integer) As Double
' This function returns either L1 or L2 distance between vectors X and Y
Dim num As Long, jj As Long
Dim dsum As Double

num = UBound(X)
If num <> UBound(Y) Or num <> k -1 Then
	ReportInformation("Distance calculation error: the length of s,t vectors is not correct")
	d_st = 0
	Exit Function
End If

dsum = 0

If nType = 1 Then 'L1 distance
	For jj = 0 To k - 1
		dsum += Abs(X(jj) - Y(jj))
	Next
ElseIf nType = 2 Then 'L2 distance
	For jj = 0 To k - 1
		dsum += (X(jj)-Y(jj))^2
	Next
	dsum = Sqr(dsum)
End If

'd_st = dsum/n
d_st = dsum

End Function
Function Weighted_RandomChoice(p As Variant) As Long

'This function returns an integer value between 0 and ubound(P) based on the probability values given by P
Dim kk As Long, ii As Long, binNo As Long
Dim rangeL() As Double, rangeH() As Double, tmp As Double

kk = UBound(p)

ReDim rangeL(kk), rangeH(kk)

rangeL(0) = 0.0
rangeH(0) = p(0)

For ii = 1 To kk
	rangeL(ii) = rangeH(ii-1)
	rangeH(ii) = rangeL(ii) + p(ii)
Next

tmp = Rnd()
For ii = 0 To kk
	If tmp >= rangeL(ii) And tmp< rangeH(ii) Then
		binNo = ii
		Exit For
	End If
Next

Weighted_RandomChoice = binNo

End Function

Function Hammersley(n As Integer, k As Integer, sOption As String) As Double() ' same as LHS, n - 'number of levels/points, k -' number of variables/dimensions

	Dim bDebug As Boolean
	Dim ii As Integer, jj As Integer
	Dim SamplePoints() As Double, HammersleyDOEPoints() As Double
	ReDim HammersleyDOEPoints (n-1, k-1) As Double

	bDebug = False

	If sOption = "Standard" Then
		For ii = 0 To n -1
			HammersleyDOEPoints(ii,0) = (ii +0.5)/ n  'avoid always lands on 0,0.. based on paper of Wong
			For jj = 1 To k-1
				HammersleyDOEPoints(ii,jj) = phi(ii,jj)
			Next
		Next
	ElseIf sOption = "Halton" Then
		For ii = 0 To n -1
			For jj = 0 To k-1
				HammersleyDOEPoints(ii,jj) = phi(ii,jj+1)
			Next
		Next
	End If
	If bDebug Then	Display2DNumArray(HammersleyDOEPoints)

Hammersley = HammersleyDOEPoints
End Function
Function Prime(D As Integer) As Integer 'return the d-th prime number
Dim p As Variant

p = Array( 2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, _
		  73, 79, 83, 89, 97, 101, 103, 107, 109, 113,  127, 131, 137, 139, 149, 151, 157, 163, 167, 173, _
		 179, 181,191, 193, 197, 199, 211, 223, 227, 229, 233, 239, 241, 251, 257,263, 269, 271, 277, 281, _
         283, 293, 307, 311, 313, 317, 331,337,347, 349, 353, 359, 367, 373, 379, 383, 389, 397, 401, 409, _
  		 419, 421, 431, 433, 439, 443, 449, 457, 461, 463, 467, 479, 487, 491, 499, 503, 509, 521, 523, 541 )

If D > 0 And D <100 Then
	Prime = p(D -1)
Else
	ReportInformation ("The dimension (# of variables) should be greater than 0 but not exceeds maximum number of 100, using first prime number - 2")
	Prime = p(0)
End If

End Function
Function phi(k As Integer, D As Integer) As Double
Dim p As Integer, pp As Long, kk As Integer, a As Double, tmp As Double

	p = Prime(D)
	pp = p
	kk = k
	a = 0.0
	tmp = 0

	While kk > 0
		a = kk Mod p
		tmp = tmp + a/pp
		kk = Int(kk/p)
		pp = pp * p
	Wend
phi = tmp
End Function


Sub Display2DNumArray(arr)
Dim ii As Integer, jj As Integer, sDisplay As String, sNum As String

sDisplay = "n" + Chr(9)

If UBound(arr,2) = UBound(sParaNames) Then
	For jj = 0 To UBound(arr,2)
		sDisplay = sDisplay + sParaNames(jj) + Chr(9)
	Next
End If

sDisplay = sDisplay + vbNewLine + vbNewLine

For ii = 0 To UBound(arr,1)
	sDisplay = sDisplay + cstr(ii + 1) + Chr(9)
	For jj = 0 To UBound(arr,2)
		sNum = IIf ((arr(ii,jj) - Int(arr(ii,jj)) <> 0), Format(arr(ii,jj),"0.000"), cstr(arr(ii,jj)))
		sDisplay = sDisplay + sNum +  Chr(9)
	Next
	sDisplay = sDisplay + vbNewLine
Next
	ReportInformation(sDisplay)
End Sub

Sub Display1DNumArray(arr)
Dim ii As Integer, jj As Integer, sDisplay As String, sNum As String
sDisplay = ""

	For ii = 0 To UBound(arr)
		'sDisplay = sDisplay  + cstr(arr(ii)) +  " "
		sNum = IIf ((arr(ii) - Int(arr(ii)) <> 0), Format(arr(ii),"0.000"), cstr(arr(ii)))
		sDisplay = sDisplay + sNum +  Chr(9)
	Next ii
	ReportInformation(sDisplay)
End Sub


