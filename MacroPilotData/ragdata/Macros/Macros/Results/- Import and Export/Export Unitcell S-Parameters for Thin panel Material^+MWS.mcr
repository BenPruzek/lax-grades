'#Language "WWB-COM"

' this macro writes unitcell S-Parameters into a txt file, which can be imported by Thin panel material and then be used with A-Solver

' --------------------------------------------------------------------------------------------------------
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 14-Oct-2024 ech: modified the sign of TM reflection coefficient according to the local coordinate system for the Refl / Trans coefficient compared to the one of the S para
' 03-Jan-2022 fhi: fixed the problem if only a single frequency is available, automatically adjust fmin and fmax to fit the given F-Solver frequ range
' 18-Oct-2021 fhi: modified check for the existence of a param-sweep
' 04-Oct-2021 fhi: extended to broadband S-Parameter export, corrected missing first theta-angle
' 30-Dec-2015 mki: first version
' ---------------------------------------------------------------------------------------------------------

'#include "vba_globals_all.lib"

'Code structure:
'Sub Main
'Sub fillVArraywithParValues
'ReturnRunIDs
'getFreq
'DialogFunc - for getFreq
'getOtherParameter
'DialogFunc1 - for getOtherParameter
'GetThetaRunIDs
'GetThetaRunIDsOtherParsGiven
'Match
'Sub doExport


Public MyVar() As Variant 'array of arrays
Public dParName() As String
Public dParVal() As Double
Public lPara As Long
Public dParaCount As Double 'index of dParName and dParVal
Public names1 As Variant, values1 As Variant
Public direct_exp As Boolean

Public sPals() As String
Public selVal As String

Option Explicit

Sub Main
Dim fmi As Double, fma As Double, NrF As Integer
'Dim lPara As Long
Dim lPara2 As Long
Dim iP As Integer, iThetaIdx As Integer
Dim sThetaCheck As String
lPara = ParameterSweep.GetNumberOfVaryingParameters

fillVArraywithParValues
'Now we have a (variant) array of arrays, with length equal to number of parameters and each element
'containing all values of a particular variable empty element where the value of a paramter didnt change

If direct_exp Then
	getFreq (fmi, fma, NrF)
	'get run ids
	Dim lThetaRunIDs() As Variant
	lThetaRunIDs = GetThetaRunIDs			'MsgBox lThetaRunIDs(0)
	'do export
	'MsgBox cstr(dfreq2)
	doExport(lThetaRunIDs,fmi, fma, NrF)	'continue with export
Else
	ReDim dParName(UBound(MyVar)-1)'varying parameters other than theta
	ReDim dParVal(UBound(MyVar)-1)
	Dim i3 As Integer
	dParaCount = 0
	For i3=0 To UBound(MyVar)
		If names1(i3)<>"theta" Then
			getOtherParameter(i3)
			dParaCount += 1
		End If
	Next
	'display freq dialogue
	getFreq (fmi, fma, NrF)
	'get run ids
	Dim lThetaRunIDsOtherParsGiven() As Variant
	lThetaRunIDsOtherParsGiven = GetThetaRunIDsOtherParsGiven 'find out what is needed as input to this function, approach 1: nothing, populate names and values globally
	'do export
	doExport(lThetaRunIDsOtherParsGiven,fmi, fma, NrF)

End If

End Sub


Sub fillVArraywithParValues

'fill variant Arraqy of Arrays with all available parameter values subroutine to automatically get all parameters

Dim TempVar() As Double 'tempArray for allocation to variant elements
Dim TVarSizeCount As Integer
Dim CompVal As Double 'to store previous instance of values for comaprison with current

Dim ii As Long, i2 As Integer, itemp As Integer
Dim found As Boolean
direct_exp = True 'flag to export directly in case no other varying variables

Dim IDs As Variant
IDs = ReturnRunIDs
Dim exists0 As Boolean
exists0 = GetParameterCombination(IDs(1), names1, values1 )
If Not exists0 Then
	MsgBox "No parameters defined"
	Exit All
End If

ReDim MyVar(UBound(values1)) 'length equal to number of parameters in file

For i2=0 To UBound(names1)'or values1, doesnt matter
	TVarSizeCount = 0 'initial array containing parameter values is size zero
	found = False

	For ii = 1 To UBound(IDs)'for each parameter, iterate over all runids
		exists0 = GetParameterCombination(IDs(ii), names1, values1 )
		If ii = 1 Then	'initialize
			ReDim TempVar(TVarSizeCount)
			TempVar(TVarSizeCount)=values1(i2)
			TVarSizeCount += 1
		Else 'compare with prev value to see if changing
			If values1(i2) <> CompVal Then 'new value discovered, but memory of only one
				If names1(i2) <> "theta" Then
					direct_exp = False
				End If
				'now check over all temp array to see if it was stored before
				For itemp = 0 To UBound(TempVar)
					If values1(i2)=TempVar(itemp) Then
						found = True
					End If
				Next
					If Not found Then
						ReDim Preserve TempVar(TVarSizeCount)
						TempVar(TVarSizeCount)=values1(i2)
						TVarSizeCount += 1
						found = False
					End If
			End If
		End If
		CompVal=values1(i2)
	Next
	MyVar(i2) = TempVar
Next
End Sub

Function ReturnRunIDs As Variant
Dim iTE As Long
Dim iTM As Long

On Error Resume Next
With FloquetPort	'first call resumes in an error, second trial works ok
    .Port ("Zmax")
    .GetModeNumberByName (iTE, "TE(0,0)") 'mode1
    .GetModeNumberByName (iTM, "TM(0,0)") 'mode2
End With
With FloquetPort
    .Port ("Zmax")
    .GetModeNumberByName (iTE, "TE(0,0)") 'mode1
    .GetModeNumberByName (iTM, "TM(0,0)") 'mode2
End With
On Error GoTo 0

Dim TreeItem As String
TreeItem = "1D Results\S-Parameters\SZmax("&CStr(iTM)&"),Zmax("&CStr(iTM)&")" 'R_TM
'get an array of existing result ids for this tree item
Dim IDs As Variant
IDs = Resulttree.GetResultIDsFromTreeItem(TreeItem)
ReturnRunIDs = IDs
End Function

'Function to Display MAIN-Dialogue for Frequency Input
Sub getFreq (dfreqmin As Double, dfreqmax As Double, dNrOfFrequ As Integer)
	Dim sdfreq As String, sdfreqmin As String, sdfreqmax As String,sNrOfFrequ As String
 	Dim cst_result	As Integer

	BeginHide
		Dim bRedefine As Boolean
	 Begin Dialog UserDialog 410,140,"Frequency Range, to be exported",.DialogFunc ' %GRID:10,5,1,1
		OKButton 20,115,90,20
		CancelButton 120,115,90,20
		GroupBox 20,5,370,105,"",.GroupBox1
		Text 30,30,110,15,"Min. frequency:",.Text1
		'TextBox 160,21,110,21,.dfreq
		Text 30,55,120,15,"Max. frequency:",.Text2
		Text 30,80,150,20,"# of frequency steps:",.Text3
		TextBox 200,25,110,20,.fmin
		TextBox 200,50,110,20,.fmax
		TextBox 200,75,110,20,.NrOfFrequ
		Text 320,30,50,15,Units.GetUnit("Frequency"),.Text4
		Text 320,55,50,15,Units.GetUnit("Frequency"),.Text5
		'PushButton 220,95,90,20,"Help",.Help
	End Dialog
	 	Dim dlg As UserDialog
	 	dlg.fmin  = CStr(Solver.GetFmin)
	 	dlg.fmax  = CStr(Solver.GetFmax)
	 	dlg.NrOfFrequ = "10"
	 	cst_result = Dialog(dlg)
     	assign "cst_result"        '  writes e.g. "cst_result = -1/0/1"     into history list
     	If (cst_result = 0) Then Exit All
		sdfreqmin  = dlg.fmin
		sdfreqmax  = dlg.fmax
		sNrOfFrequ = dlg.NrOfFrequ
		dfreqmin  = Evaluate(sdfreqmin)
		dfreqmax = Evaluate(sdfreqmax)
		If dfreqmin > dfreqmax Then
			ReportInformationToWindow ( "Frequencies reset due to wrong input data."   )
			sdfreqmin  = CStr(Solver.GetFmin)
			sdfreqmax  = CStr(Solver.GetFmax)
			dfreqmin  = Evaluate(sdfreqmin)
			dfreqmax = Evaluate(sdfreqmax)
		End If
		dNrOfFrequ = Abs(Evaluate(sNrOfFrequ))
		If dNrOfFrequ = 0 Then
			ReportInformationToWindow ( "Number of Frequencies set back to 1."   )
			dNrOfFrequ = 1
		End If
		If dNrOfFrequ > 1001 Then
			ReportInformationToWindow ( "Number of Frequencies exceeds 1001."   )
			dNrOfFrequ = 1
		End If
	 	assign "dfreqmin"
		assign "dfreqmax"
		assign "dNrOfFrequ"
	EndHide

 	cst_result = Evaluate(cst_result)
 	If (cst_result =0) Then Exit All   ' if cancel/help is clicked, exit all
 	If (cst_result =1) Then Exit All

End Sub

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		'Select Case Item
		'Case "Help"
		'	StartHelp "special_struct_coated_material"
		'	DialogFunc = True
		'End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function

Function getOtherParameter(ByVal iP2 As Integer) As Boolean

dParName(dParaCount) = names1(iP2)
ReDim sPals(UBound(MyVar(iP2)))
Dim i4 As Integer
For i4=0 To UBound(sPals)
	sPals(i4) = CStr(MyVar(iP2)(i4))
Next
Dim Val1 As Double
		Dim bRedefine As Boolean
	Begin Dialog UserDialog1 410,105,"Specify Value of Other Parameters",.DialogFunc1 ' %GRID:10,7,1,1
			'OKButton 20,85,90,20
		PushButton 10,77,90,21,"Ok",.instigate
		CancelButton 110,77,90,21
		GroupBox 10,7,390,56,"",.GroupBox1
		Text 20,28,180,21,dParName(dParaCount),.Text1
		DropListBox 240,28,130,121,sPals(),.Val
	End Dialog
	 	Dim dlg1 As UserDialog1
		If Dialog(dlg1)=0   Then Exit All
End Function

Function DialogFunc1%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "instigate"
			'MsgBox "User pressed ok
			dParVal(dParaCount)= CDbl(DlgText("Val"))
			'MsgBox Str$(dParVal(dParaCount))
			DialogFunc1 = False
		Case "Help"
			StartHelp "special_struct_coated_material"
			DialogFunc1 = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function


Function GetThetaRunIDs() As Variant()

Dim IDs As Variant
IDs = ReturnRunIDs

Dim thetarunids As Variant
ReDim thetarunids(UBound(IDs))	'-1)  !!!!
Dim N As Long 'number of run IDs
For N = 0 To UBound(IDs)	'-1
	thetarunids(N) = IDs(N) 	'+1)
	'MsgBox IDs(N)	'+1)
Next
GetThetaRunIDs = thetarunids

'MsgBox "getIDs" + cstr( N)
End Function

Function GetThetaRunIDsOtherParsGiven() As Variant()
'we have dParName() and dParVal() filled
'store only runids where values match as in dParValues
'otherwise issue warning that no info for this value of parameter is stored and abort

Dim vTempRID As Variant
Dim count As Long
count = 0

Dim IDs As Variant
IDs = ReturnRunIDs 'IDs contains all run ids, goes from 1 to UBound(IDs), 0=current
Dim iteratorID As Integer, iteratorVar As Integer, iteratorAll As Integer

Dim exists1 As Boolean

For iteratorID = 1 To UBound(IDs)
	exists1 = GetParameterCombination(IDs(iteratorID), names1, values1 )
		If Match(0,UBound(names1),UBound(dParName)+1) Then 'UBound gives upper index, not size
			ReDim Preserve vTempRID(count)
			vTempRID(count)=IDs(iteratorID)
			count += 1
		End If
Next
GetThetaRunIDsOtherParsGiven = vTempRID
End Function

'recursive function
Function Match(ByVal lower_bound As Integer, ByVal upper_bound As Integer, ByVal vPara_size As Integer)As Boolean
Dim abc As Integer
For abc = lower_bound To upper_bound
If vPara_size <> 0 Then
	If (dParName(vPara_size-1) = names1(abc) And dParVal(vPara_size-1) = values1(abc))Then 'comparing from backwards
		vPara_size -= 1
		If(Match(lower_bound,upper_bound,vPara_size))Then
			Match = True
			Exit Function
		End If
	End If
Else
	Match = True
	Exit For
	Exit Function
End If
Next
End Function

Sub doExport(ByVal runIDs() As Variant, dfreqmin As Double, dfreqmax As Double,NrOfFrequ As Integer)
Dim sFileName As String, sFileName1 As String
'sFileName = GetProjectPath("Root")
sFileName1 = getName
sFileName = GetProjectPath("Root")+"\"+sFileName1+".txt"
Open sFileName For Output As #1
Dim iter As Integer, dfreqsingle As Double
If Not direct_exp Then
	For iter = 0 To UBound(dParName)
		Print #1,  "# " + CStr(dParName(iter)) + " = " + CStr(dParVal(iter))
	Next
End If
Print #1,  "#-----------------------------------------------------------------------------------------------------------------------------------------------------"
Print #1,  "# Frequency " + " Inc. Angle " + " Re(R_TM) " + " Im(R_TM) " + " Re(R_TE) " + " Im(R_TE) " + " Re(T_TM) " + " Im(T_TM) " + " Re(T_TE) " + " Im(T_TE) "
Print #1,  "#"
Print #1,  "#-----------------------------------------------------------------------------------------------------------------------------------------------------"

Dim iTE As Long
Dim iTM As Long
With FloquetPort
    .Port ("Zmax")
    .GetModeNumberByName (iTE, "TE(0,0)") 'mode1
    .GetModeNumberByName (iTM, "TM(0,0)") 'mode2
End With
Dim TreeItem(0 To 3) As String
TreeItem(0) = "1D Results\S-Parameters\SZmax("&CStr(iTM)&"),Zmax("&CStr(iTM)&")" 'R_TM
TreeItem(1) = "1D Results\S-Parameters\SZmax("&CStr(iTE)&"),Zmax("&CStr(iTE)&")" 'R_TE
TreeItem(2) = "1D Results\S-Parameters\SZmin("&CStr(iTM)&"),Zmax("&CStr(iTM)&")" 'T_TM
TreeItem(3) = "1D Results\S-Parameters\SZmin("&CStr(iTE)&"),Zmax("&CStr(iTE)&")" 'T_TE

	Dim iterator As Long

   'For iterator = 0 To UBound(runIDs) ' runid=0 contains no parametric info
	For iterator = 1 To UBound(runIDs)

		Dim spara(4) As Object
		Set spara(0) = Resulttree.GetResultFromTreeItem(TreeItem(0), runIDs(iterator))
		Set spara(1) = Resulttree.GetResultFromTreeItem(TreeItem(1), runIDs(iterator))
		Set spara(2) = Resulttree.GetResultFromTreeItem(TreeItem(2), runIDs(iterator))
		Set spara(3) = Resulttree.GetResultFromTreeItem(TreeItem(3), runIDs(iterator))

		Dim names As Variant, values As Variant, exists As Boolean
		exists = GetParameterCombination(runIDs(iterator), names, values )

		'	Find index of theta in case more than one variables changing
		Dim index As Integer 'index of theta
		Dim M As Long
		'check if pararmetric data is available at all
	   'If values = "" Then  ' modified check_for_parametric statement
		If Not exists Then
			ReportInformationToWindow ("No parametric data available! Please run a parameter sweep for ""theta"" first!")
			Exit All
		End If
		For M = 0 To UBound( values )
				If names(M)="theta" Then
					index = M
				End If
		Next

		Dim iiitmp As Long, iiitmp1 As Long
		Dim dXtmp1 As Double, dXtmp2 As Double
		Dim dfreq As Double, df As Double
		Dim dXValues() As Double, i As Integer
		Dim FSolver_fmin As Double, FSolver_fmax As Double


		If spara(0).getn() > 1 Then	'check GUI min and max frequencies and adjust if necessary according to available F-solver data
			FSolver_fmin = (spara(0).getx(0))
			FSolver_fmax = (spara(0).getx(spara(0).getn()-1))
			If dfreqmin < FSolver_fmin Then
				dfreqmin = FSolver_fmin
			End If
			If dfreqmax > FSolver_fmax Then
				dfreqmax = FSolver_fmax
			End If
		End If

		'frequency loop
		If (dfreqmax -dfreqmin) > 0 And NrOfFrequ > 0 Then
			df = (dfreqmax -dfreqmin)/NrOfFrequ
		Else
			df=0
			NrOfFrequ=0
		End If

		dfreq = dfreqmin

		If spara(0).getn() = 1 Then  ' case: only one single frequency available, checked by using the length of data in the object
			df=0
			NrOfFrequ=0
		End If

 		For i = 1 To NrOfFrequ+1
			iiitmp = 0
			If (dfreq <> spara(0).GetX(0)) Then
				' Go through all values, calculate delta between target and previous value, as well as between target and current value
				' If product of the two deltas is 0 or negative, the target was met exactly or lies between the current and the previous value
				dXtmp1 = dfreq-spara(0).GetX(0)
				dXValues = spara(0).GetArray("x")
				For iiitmp = 1 To spara(0).GetN()-1
					dXtmp2 = dfreq - dXValues(iiitmp)
					If (dXtmp1*dXtmp2 <= 0) Then
						Exit For
					Else
						dXtmp1 = dXtmp2
					End If
				Next
			End If
			'values needed, with or without interpolation:
			'spara(0).GetYRe = Re(R_TM)
			'spara(0).GetYIm = Im(R_TM)
			'spara(1).GetYRe = Re(R_TE)
			'spara(1).GetYIm = Im(R_TE)
			'spara(2).GetYRe = Re(T_TM)
			'spara(2).GetYIm = Im(T_TM)
			'spara(3).GetYRe = Re(T_TE)
			'spara(3).GetYIm = Im(T_TE)
			Dim dRe_R_TM As Double,dIm_R_TM As Double,dRe_R_TE As Double,dIm_R_TE As Double
			Dim dRe_T_TM As Double,dIm_T_TM As Double,dRe_T_TE As Double,dIm_T_TE As Double

			If spara(0).getn() = 1 Then
				iiitmp = 0 'iiitmp-1	reset the index
			End If


			If ( dfreq <> spara(0).GetX(iiitmp) )  And (spara(0).getn() > 1) Then
			'If ( dfreq <> spara(0).GetX(iiitmp) ) Then	' modified the interpolation check, enter only if more that one frequency is available
				' interpolate between iiitmp-1 and iiitmp
				dRe_R_TM = dInterpolate_linear_y(dfreq, spara(0).GetX(iiitmp-1),spara(0).GetX(iiitmp),spara(0).GetYRe(iiitmp-1),spara(0).GetYRe(iiitmp))
				dIm_R_TM = dInterpolate_linear_y(dfreq, spara(0).GetX(iiitmp-1),spara(0).GetX(iiitmp),spara(0).GetYIm(iiitmp-1),spara(0).GetYIm(iiitmp))
				dRe_R_TE = dInterpolate_linear_y(dfreq, spara(1).GetX(iiitmp-1),spara(1).GetX(iiitmp),spara(1).GetYRe(iiitmp-1),spara(1).GetYRe(iiitmp))
				dIm_R_TE = dInterpolate_linear_y(dfreq, spara(1).GetX(iiitmp-1),spara(1).GetX(iiitmp),spara(1).GetYIm(iiitmp-1),spara(1).GetYIm(iiitmp))
				dRe_T_TM = dInterpolate_linear_y(dfreq, spara(2).GetX(iiitmp-1),spara(2).GetX(iiitmp),spara(2).GetYRe(iiitmp-1),spara(2).GetYRe(iiitmp))
				dIm_T_TM = dInterpolate_linear_y(dfreq, spara(2).GetX(iiitmp-1),spara(2).GetX(iiitmp),spara(2).GetYIm(iiitmp-1),spara(2).GetYIm(iiitmp))
				dRe_T_TE = dInterpolate_linear_y(dfreq, spara(3).GetX(iiitmp-1),spara(3).GetX(iiitmp),spara(3).GetYRe(iiitmp-1),spara(3).GetYRe(iiitmp))
				dIm_T_TE = dInterpolate_linear_y(dfreq, spara(3).GetX(iiitmp-1),spara(3).GetX(iiitmp),spara(3).GetYIm(iiitmp-1),spara(3).GetYIm(iiitmp))

			Else
				' no interpolation necessary
				dRe_R_TM = spara(0).GetYRe(iiitmp)
				dIm_R_TM = spara(0).GetYIm(iiitmp)
				dRe_R_TE = spara(1).GetYRe(iiitmp)
				dIm_R_TE = spara(1).GetYIm(iiitmp)
				dRe_T_TM = spara(2).GetYRe(iiitmp)
				dIm_T_TM = spara(2).GetYIm(iiitmp)
				dRe_T_TE = spara(3).GetYRe(iiitmp)
				dIm_T_TE = spara(3).GetYIm(iiitmp)

			End If

			' need to invert the TM reflection coefficient to match the general Refl / Trans. local coordinate system compared to the S-para one
			dRe_R_TM = -dRe_R_TM
			dIm_R_TM = -dIm_R_TM
			
			If spara(0).getn() = 1 Then 'only one frequency available, assign it for output
				dfreq = spara(0).getx(0)
				dfreqsingle= dfreq
			End If

			Print #1, CStr(dfreq)+" "+CStr(values(index))+" "+CStr(dRe_R_TM)+" "+CStr(dIm_R_TM)+" "+CStr(dRe_R_TE)+" "+CStr(dIm_R_TE)+" "+CStr(dRe_T_TM)+" "+CStr(dIm_T_TM)+" "+CStr(dRe_T_TE)+" "+CStr(dIm_T_TE)
			'Print #1, CStr(dfreq)+" "+CStr(values(index))+"  0 0 0 0    0 0 0 0" 	'perfect absorber
			 'Print #1, CStr(dfreq)+" "+CStr(values(index))+"  -1 0 -1 0    0 0 0 0" 	'perfect reflector PEC
			 'Print #1, CStr(dfreq)+" "+CStr(values(index))+"  1 0 1 0   0 0 0 0" 	'perfect reflector PMC
			'Print #1, CStr(dfreq)+" "+CStr(values(index))+"  0 0 0 0    1 0 1 0" 	'perfect transmitter


			dfreq = dfreqmin + i*df

		Next 'frequ loop
	Next
	Close #1

	If spara(0).getn() = 1 Then 'only one frequency available, assign it for output
		ReportInformationToWindow ( "Output for a single Frequency at " + Str$(dfreqsingle)+ " "+Units.GetUnit("Frequency"))
	Else
		ReportInformationToWindow ( "Frequency_min: " + Str$(dfreqmin)+ " "+Units.GetUnit("Frequency") + "    Frequency_max: " + Str$(dfreqmax)+ " "+Units.GetUnit("Frequency")+ "    Delta-Frequency: "+Str$(df)+ " "+Units.GetUnit("Frequency")  + "    Nr_of_Frequencies: "+Str$(NrOfFrequ+1))
	End If

	ReportInformationToWindow ( "Number of RunIDs found:        " + Str$(UBound(runIDs)+1 ))
	ReportInformationToWindow ( "Data successfully exported to: " + sFileName )

End Sub

'Function to Display Dialogue for File Name Selection
Function getName As String
	Dim sName As String
 	Dim cst_result	As Integer

	BeginHide
		Dim bRedefine As Boolean
	Begin Dialog UserDialog 420,119,"Enter File Name",.DialogFunc ' %GRID:10,7,1,1
		OKButton 10,91,90,21
		CancelButton 110,91,90,21
		GroupBox 10,7,400,77,"",.GroupBox1
		Text 20,28,80,21,"File Name:",.Text1
		TextBox 100,21,280,21,.dNameHandle
		Text 30,56,360,14,"(File will be exported to same location as cst-file)",.Text2
	End Dialog
	 	Dim dlg As UserDialog
	 	dlg.dNameHandle = "TXRX_Table_for_Thin_Panel"
	 	cst_result = Dialog(dlg)
     	assign "cst_result"        '  writes e.g. "cst_result = -1/0/1"     into history list
     	If (cst_result = 0) Then Exit All
		sName = dlg.dNameHandle
	 	assign "sName"
	EndHide

 	cst_result = Evaluate(cst_result)
 	If (cst_result =0) Then Exit All   ' if cancel/help is clicked, exit all
 	If (cst_result =1) Then Exit All

 	getName = sName
End Function

Function DialogFunc2%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		'Select Case Item
		'Case "Help"
		'	StartHelp "special_struct_coated_material"
		'	DialogFunc = True
		'End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
