' *Filter Analysis / Tchebychev-Factors
' !!!
' macro.951
'
' ================================================================================================
' Copyright 2001-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 22-Jan-2013 fhi: added k/BW and Qe
' 10-Jun-2009 fhi: refresh_view (readme file)
' 01-Dec-2008 fhi: added center frequency to get rel.couplingBW
' 21-Oct-2005 imu: Included into Online Help
' 05-Oct-2005 fhi: DesEnv2006: cosmetic
' 14-Feb-2002 fhi: correction of Editor-Resultsfilename
' 03-Dec-2001 fhi: Make output-name unique by adding projectbasename 
' 23-Nov-2001 fhi: Added comments to Output such as Bandwidth, Order and VSWR
' 01-Nov-2001 reh: initial version
Option Explicit

Const HelpFileName = "common_preloadedmacro_filter_analysis_tchebychev-factors"

Sub Main
	
	Dim iii As Integer, iiim1 As Integer, outstr As String, outstr1 As String, StoreDir As String, outstr3 As String
	'--- passband ripple (z.B. 0.01 dB)
	Dim Lar As Double
	'--- bandwidth of the filter; reference Lar
	Dim bw As Double
	'--- return loss (e.g. -20dB)
	Dim RTloss As Double
	Dim norder As Integer
	Dim gis() As Double
	Dim kvals() As Double
	'Dim insertionloss As Boolean
	Dim tds() As Double
	Dim vswr As Double
	
	Begin Dialog UserDialog 310,224,"Tchebychev Filter Coefficients",.DialogFunc ' %GRID:10,7,1,1
		TextBox 140,21,110,21,.ordnung
		Text 10,24,90,14,"Filter Order",.Text1
		TextBox 140,42,110,21,.bandwidth
		Text 8,46,130,14,"Bandwidth in MHz",.Text2
		Text 11,4,143,14,"Tchebyscheff Filter",.Text3
		GroupBox 0,98,280,84,"Given Value",.GroupBox1
		CheckBox 20,119,250,14,"Insertion Loss (Passband ripple) dB",.CheckIloss
		CheckBox 20,161,140,14,"Return Loss dB",.CheckRloss
		CheckBox 20,140,140,14,"Passband VSWR",.passbandVswr
		TextBox 170,154,100,21,.Lossvalue
		OKButton 10,196,90,21
		CancelButton 110,196,90,21
		PushButton 210,196,90,21,"Help",.Help
		TextBox 140,63,110,21,.f_o
		Text 10,68,120,14,"Center Frequency",.Text4
	End Dialog
	Dim dlg As UserDialog
	
	Dim  projectname As String
 
 	projectname = GetProjectPath("Model3D")
	
	With dlg
		.ordnung = "4"
        .bandwidth = "36"
        .CheckIloss = 1
        .CheckRloss = 0 
        .passbandVswr = 0             			
        .Lossvalue = "0.01"
        .f_o = "1000"
	End With
	
	If (Dialog(dlg) = 0) Then Exit All
	
	StoreDir = ""
	
	
	If dlg.CheckILoss Then
		Lar = Abs(RealVal(dlg.Lossvalue))
		RTloss = -Abs(Iloss(Lar))
		vswr = sparm2vswr(10^(-Abs(RTloss)/20))
	End If
	If dlg.CheckRloss Then
		RTloss = -Abs(RealVal(dlg.Lossvalue))
		Lar = Abs(Iloss(RTloss))
		vswr = sparm2vswr(10^(-Abs(RTloss)/20))
	End If
	If dlg.passbandVswr Then 
	     vswr= Abs(RealVal(dlg.Lossvalue))
	     Lar = Abs(20*Log(Sqr(1-vswr2sparm(RealVal(dlg.Lossvalue))^2))/Log(10))
		 RTloss = -Abs(Iloss(Lar))
		 
	End If
	
	
	norder = CInt(dlg.ordnung)
	If norder < 1 Then
		MsgBox "Specify an order bigger than 0"
		Exit Sub
	End If
	'--- resonance freq. and bandwidth in MHz
	bw = RealVal(dlg.bandwidth)

	'--- compute the g values -----------------------------------------------------------
	ReDim gis(1 To norder+1)
	gis(1) = 2*ak(1,norder)/gamma(Lar,norder)
	For iii = 2 To norder
		iiim1 = iii-1
		gis(iii) = 4*ak(iiim1, norder)*ak(iii, norder)/bk(iiim1, norder, Lar)/gis(iiim1)
	Next iii
	If Abs(Int(CDbl(norder)/2.0)-CDbl(norder)/2.0) < 0.1 Then
		gis(norder+1) = coth(beta(Lar)/4)^2
	Else
		gis(norder+1) = 1.0
	End If
	'------------------------------------------------------------------------------------
	
	'--- compute the coupling bandwidths
	ReDim kvals(1 To norder+1)
	kvals(1) = bw/gis(1)
	For iii = 2 To norder
		iiim1 = iii-1
		kvals(iii) = bw/Sqr(gis(iiim1)*gis(iii))
	Next iii
	kvals(norder+1) = bw/gis(norder)/gis(norder+1)
	
	outstr = "k_E   = " + Format(kvals(1),"###.##") + _
	 			vbTab + "("+ Format(kvals(1)/Abs(RealVal(dlg.f_o)),"###0.#######")+")"+  _
				vbTab + "("+ Format((kvals(1)/bw),"###0.#######")+")"+  _
					vbTab + "("+ Format((Abs(RealVal(dlg.f_o))/kvals(1)),"###0.#######")+")"+ vbCrLf

	For iii = 2 To norder
		outstr = outstr + "k" + CStr(iii-1) + "_" + CStr(iii) + "  = " + Format(kvals(iii),"###.##") +  _
		               vbTab + "("+ Format(kvals(iii)/Abs(RealVal(dlg.f_o)),"###0.#######")+")"+ _
						 vbTab + "("+ Format((kvals(iii)/bw),"###0.#######")+")"+ vbCrLf
	Next iii 
	outstr1 = outstr + "k_out = " + Format(kvals(norder+1),"###.##") +vbTab + "("+ Format(kvals(norder+1)/Abs(RealVal(dlg.f_o)),"###0.#######")+")" +  _
								vbTab + "("+ Format((kvals(norder+1)/bw),"###0.#######")+")"+vbCrLf
	
	'--- compute the group delay time for the 
	ReDim tds(1 To norder+1)
	tds(1) = 636.6/kvals(1)
	If norder > 1 Then
		tds(2) = (636.6/kvals(2))^2/tds(1)
	End If
	If norder > 2 Then
		tds(3) = (636.6/kvals(3))^2/tds(2)+tds(1)
	End If
	For iii = 4 To norder+1
		tds(iii) = (636.6/kvals(iii))^2/(tds(iii-1)-tds(iii-3))+tds(iii-2)
	Next iii
	
	' --- write output string
	outstr3 = "Group Delay Time" + vbCrLf + "----------------" + vbCrLf
	For iii = 1 To norder+1
		outstr3 = outstr3 + "t_d" + CStr(iii) + " = " + Format(tds(iii),"##0.###") + " ns" + vbCrLf
	Next iii
	
	'--- write g values to file -----------------------------------------------
	outstr=""
	For iii = 1 To norder+1
		outstr = outstr+"g"+CStr(iii)+" = "+Format(gis(iii),"0.0000")+vbCrLf
	Next iii
	
	Open StoreDir + projectname + "Tchebychev_g_K_td.txt" For Output As #1
	
	Print #1, "Tchebychev Filter"
	Print #1, "==================="
	Print #1, " "
	Print #1, "Order            = " + CStr(norder)
	Print #1, "Bandwidth        = " + CStr(bw) + " MHz"
	Print #1, "Center Frequency = " + CStr(dlg.f_o) + " MHz"
	Print #1, "Passband ripple  = " + Format(Lar,"0.######") + " dB" + _
	              "   (" + Format(vswr,"#.######") + " VSWR)"
	Print #1, "Return loss      = " + Format(RTloss,"##.0###") + " dB"
	Print #1, ""
	Print #1, "Normed g values:"
	Print #1, "-------------------------------------------"
	Print #1, outstr
	Print #1, ""
	Print #1, "Coupling Coefficients k"
	Print #1, "k       (MHz)      (k/fo)        (k/BW)          Qe"
	Print #1, "------------------------------------------------------------"
	Print #1, outstr1
	Print #1, ""
	Print #1, outstr3

	Close #1
	'--------------------------------------------------------------------------
	Start(StoreDir+ projectname+"Tchebychev_g_K_td.txt")

	'ResultTree.RefreshView
	Resulttree.updatetree


		
End Sub
'---------------------------------------
Function sinh(x As Double) As Double
	sinh = 0.5 * (Exp(x)-Exp(-x))
End Function
'---------------------------------------
Function cosh(x As Double) As Double
	cosh = 0.5 * (Exp(x)+Exp(-x))
End Function
'---------------------------------------
Function tanh(x As Double) As Double
	tanh = sinh(x)/cosh(x)
End Function
'---------------------------------------
Function coth(x As Double) As Double
	coth = 1/tanh(x)
End Function
'---------------------------------------
Function Iloss(Rloss As Double) As Double
	Iloss = 20*Log(Sqr(1-(10^(-Abs(Rloss)/20))^2))/Log(10)
End Function
'---------------------------------------
Function vswr2sparm(vswr As Double) As Double
	vswr2sparm=(vswr-1)/(vswr+1)
End Function
'---------------------------------------
Function sparm2vswr(sparm As Double) As Double
	sparm2vswr=(sparm+1)/(1-sparm)
End Function
'---------------------------------------
Function ak(k As Integer, norder As Integer) As Double
	ak = Sin((2*k-1)*Pi/2/norder)
End Function
'---------------------------------------
Function bk(k As Integer, norder As Integer, Lar As Double) As Double
	bk = gamma(Lar,norder)^2 + (Sin(k*Pi/norder))^2
End Function
'---------------------------------------
Function beta(Lar As Double) As Double
	Dim tval As Double
	tval = coth(Lar/17.37)
	If tval <0 Then
		beta = 1
	Else
		beta = Log(tval)
	End If
End Function
'---------------------------------------
Function gamma(Lar As Double, norder As Integer) As Double
	gamma = sinh(beta(Lar)/2.0/CDbl(norder))
End Function
'---------------------------------------
Sub Start (lib_filename As String)

        On Error GoTo Win95
        WINNT:
                Shell "cmd /c " + Quote(lib_filename)
                Exit Sub
        Win95:
                Shell "start " + Quote(lib_filename)
                Exit Sub

End Sub
'---------------------------------------
Function Quote (lib_Text As String) As String

        Quote = Chr$(34) + lib_Text + Chr$(34)

End Function
'----------------------------------------
Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

            Select Case Action
                Case 1 ' Dialog box initialization
                	Beep
                Case 2 ' Value changing or button pressed
                        Select Case Item
 							   Case "Help"
									StartHelp HelpFileName
									DialogFunc = True
                               Case "CheckIloss"
                                		DlgValue("CheckRloss", 0)
                                		DlgValue("passbandVswr",0)
                                		DlgText("Lossvalue","0.01")
                                	    DialogFunc% = True
                                Case "CheckRloss"
                                        DlgValue("CheckIloss",0)
                                        	DlgValue("passbandVswr",0)

                                        DlgText("Lossvalue","-20.0")
                                       
                                        DialogFunc% = True
                                Case "passbandVswr"
                                 		
                                 			DlgValue("CheckRloss", 0)
											 DlgValue("CheckIloss",0)

                                        DlgText("Lossvalue","1.05")
                                        DialogFunc% = True

                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select

End Function
'---------------------------------------------------------------
Function RealVal_old(lib_Text As Variant) As Double

        If (CDbl("0.5") > 1) Then
                RealVal_old = CDbl(Replace(lib_Text, ".", ","))
        Else
                On Error Resume Next
                        RealVal_old = CDbl(lib_Text)
                On Error GoTo 0
        End If

End Function
Function RealVal(lib_Text As Variant) As Double
	RealVal=evaluate(lib_Text)
End Function

Function g_k_BaseName (lib_path As String) As String

        Dim lib_dircount As Integer, lib_extcount As Integer, lib_filename As String

        lib_dircount = InStrRev(lib_path, "\")
        lib_filename = Mid$(lib_path, lib_dircount+1)
        lib_extcount = InStrRev(lib_filename, ".")
        g_k_BaseName = Left$(lib_filename, IIf(lib_extcount > 0, lib_extcount-1, 999))

End Function
