'#Language "WWB-COM"

' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 27-Sep-2011 rsj: First version
' ================================================================================================
Option Explicit

' List of changes
' fsr: 29-Jul-2015 : Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' rsj: 25-Feb-2015 : Add message box for an empty file path and few GUI modifications to avoid the crash as well.
' rsj: 09-Jan-2012 : Start index from 10% of response signal and improving the PDA plot for better agreement with eye diagram plot
' rsj: 28-Dec-2011 : Implement ExportToASCII function to avoid error message inside PCBS and CBLS and move all the results to the /Temp/DS folder
' rsj: 25-Jul-2011 : Add Xtalk functionality
' rsj: 20-Jul-2011 : Add s-parameter input and some changes in input data and plotting
' gro: First version

Public UI As Double,sTrise As Double,sThold As Double 'Store the unit bit interval T_r+T_hold
Dim sPath As String,sTimeUnit As String
Dim sSparaList() As String, SelectedArray() As String, TempSparaList () As String, temp() As String
Dim ii As Integer, ll As Integer
Dim FileInputResponse As String, FileXtalkList () As String
Dim hlpcnt As Integer


Sub Main

	ReDim SelectedArray (1)
	ReDim TempSparaList (1)

    'UI=0.5     'Default value, corresponds to 2Gbps

'While True
	Begin Dialog UserDialog 480,350,"Peak Distortion Analysis",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,14,460,154,"Pulse Response ",.GroupBox1
		Text 30,63,60,14,"Filepath",.Filepath
		PushButton 70,301,190,21,"Calculate PDA",.CalcPDA
		PushButton 360,56,90,21,"Browse...",.Browse
		PushButton 300,301,90,21,"Close",.Close
		TextBox 130,56,210,21,.PulseFile
		TextBox 30,140,310,21,.SParaTask
		GroupBox 10,182,460,112,"Data Rate settings",.GroupBox3
		Text 30,203,140,28,"Bit Unit Interval"+vbNewLine+"(Trise+Thold)",.Text1
		TextBox 150,203,110,21,.u_i
		OptionGroup .Group1
			OptionButton 30,35,220,14,"From External Time Signal File",.OptionButton1
			OptionButton 30,119,220,14,"From S-Parameter Results",.OptionButton2
		Text 30,238,90,14,"Trise",.sTrise
		Text 30,266,90,14,"Thold",.sThold
		TextBox 150,231,110,21,.cst_trise
		TextBox 150,259,110,21,.cst_thold
		Text 270,210,90,14,"["+Units.GetUnit("Time")+"]",.Text2
		Text 270,238,90,14,"["+Units.GetUnit("Time")+"]",.Text3
		Text 270,266,90,14,"["+Units.GetUnit("Time")+"]",.Text4
		TextBox 130,84,90,21,.timeaxis_unit
		Text 30,91,90,14,"Time axis unit",.Text5
		PushButton 360,140,90,21,"XTalk List",.CrossTalkList
		Text 310,238,150,14,"freq =",.Freq
		Text 310,266,150,14,"fclk =",.Fclk
		Text 20,329,140,14,"",.cst_Progress
	End Dialog

	Dim dlg As UserDialog
		If (Dialog(dlg) = 0) Then Exit All

'Wend

End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
	    Dim dlg As UserDialog
		Dim cst_SparaList() As String
		Dim success As Boolean

	    Select Case Action%
	    Case 1 ' Dialog box initialization
			DlgEnable "PulseFile", 0 ' Field not editable
			DlgText "PulseFile", "Please chose file" ' Initiaization
			DlgText "SParaTask", "Tasks\SPara1\S-Parameters\S2,1"
			sPath=DlgText "SParaTask"
			DlgValue "Group1",1
			DlgText "cst_trise","30.0"
			DlgText "cst_thold","70.0"
			sTrise=Evaluate(DlgText("cst_trise"))
			sThold=Evaluate(DlgText("cst_thold"))
			DlgText "u_i","100.0"
			DlgEnable "u_i", False
			UI=100
			DlgEnable "cst_trise",True
			DlgEnable "cst_thold",True
			DlgEnable "SParaTask",True
			DlgText "timeaxis_unit",Units.GetUnit("Time")
			DlgEnable "timeaxis_unit", False
			DlgEnable "CrossTalkList",True
			DlgEnable "Freq",False
			DlgEnable "Fclk",False
			DlgText "Freq","freq = "+Format(Cstr(0.35/(sTrise*Units.GetFrequencyUnitToSI*Units.GetTimeUnitToSI)),"0.000")+" "+Units.GetUnit("Frequency")
			DlgText "Fclk","fclk = "+Format(Cstr(0.35/((2*sTrise+2*sThold)*Units.GetFrequencyUnitToSI*Units.GetTimeUnitToSI)),"0.000")+" "+Units.GetUnit("Frequency")
			DlgText "cst_Progress",""
	    Case 2 ' Value changing or button pressed

	      	Select Case DlgItem$
		        Case "Browse" 	'
				'RSJ: Improved browse capability for several input datas
				DataInput
				'shows only the response file
				DlgText "PulseFile",FileInputResponse
	            DialogFunc = True 				'do not exit the dialog
	        Case "CalcPDA" 	'
	        	If DlgValue("Group1")=1 Then
					DlgText "cst_Progress","Calculating..."
	        		CalculatePulseRespondFromSPara(DlgValue("Group1"))
	        	Else
	        		DlgText "cst_Progress","Calculating..."
	        	End If
				success=CalculatePDAFromSParameter(getprojectpath("TempDS")+"PDA_responses.txt",DlgValue("Group1"))

				If success=True Then
					DS.SelectTreeItem("Results\PDA - WorstEye")
					DlgText "cst_Progress","Done!"
				Else
					DlgText "cst_Progress","Calculation Failed!"
				End If
				DialogFunc = True
			Case "Group1"
				If DlgValue("Group1")=1 Then
  				   DlgEnable "SParaTask",True
				   DlgEnable "cst_trise",True
				   DlgEnable "cst_thold",True
				   DlgEnable "u_i",False
				   DlgEnable "Browse",False
				   DlgEnable "timeaxis_unit",False
				   DlgEnable "CrossTalkList",True
				Else
				   DlgEnable "SParaTask",False
				   DlgEnable "cst_trise",False
				   DlgEnable "cst_thold",False
				   DlgEnable "u_i",True
				   DlgEnable "Browse",True
				   DlgEnable "timeaxis_unit",True
				   DlgEnable "CrossTalkList",False
				End If
			Case "CrossTalkList"
				SParaList
				DialogCrossTalkList
				DialogFunc = True
	        Case "Close" 	'
				DialogFunc = False
				Exit All
			End Select

		Case 3 ' TextBox or ComboBox text changed
				UI = Evaluate(DlgText "u_i") ' Update UI value
				sPath=DlgText("SParaTask")
				sTrise=Evaluate(DlgText("cst_trise"))
				sThold=Evaluate(DlgText("cst_thold"))
				sTimeUnit=DlgText("timeaxis_unit")
				'use factor 1 to calculate the max freq
				DlgText "Freq","freq = "+Format(Cstr(1/(sTrise*Units.GetFrequencyUnitToSI*Units.GetTimeUnitToSI)),"0.000")+" "+Units.GetUnit("Frequency")
				'2 bits is 1 periode of sinus signal
				DlgText "Fclk","fclk = "+Format(Cstr(1/((2*sTrise+2*sThold)*Units.GetFrequencyUnitToSI*Units.GetTimeUnitToSI)),"0.000")+" "+Units.GetUnit("Frequency")
	End Select
	End Function

'This is the original from Gerardo. Never called in this macro
Sub CalculatePDA(sFname As String)

	Dim rImpulseResponse As Object
 	Set rImpulseResponse = DS.result1d("") ' DSResults1D, otherwise addtotree adds into MWS

 	rImpulseResponse.LoadPlainFile(sFname) ' PlainFile, since it could be from another project.

	'rsj: Scale the time signal unit
	Dim ScaleFactor As Double
	Select Case sTimeUnit
		Case "ns" 'Nanosecond
			ScaleFactor=Units.GetTimesitounit*1e-9
		Case "s" 'Second
			ScaleFactor=Units.GetTimesitounit
		Case "us" 'Microsecond
			ScaleFactor=Units.gettimesitounit*1e-6
		Case "ps" 'Picosecond
			ScaleFactor=Units.gettimesitounit*1e-12
		Case "ms" 'Microsecond
			ScaleFactor=Units.gettimesitounit*1e-3
		Case "fs" 'Picosecond
			ScaleFactor=Units.gettimesitounit*1e-15
	End Select

	'rsj:Now scale the time unit with proper factor
	Dim cst_index As Long
	If ScaleFactor<>1 Then
		For cst_index=0 To rImpulseResponse.getn-1 Step 1
			rImpulseResponse.SetX(cst_index,rImpulseResponse.getx(cst_index)*ScaleFactor)
		Next
	End If

	'rsj:switch to worsteye1 and worsteye0 to get a better curve plot
 	Dim rWorstEye1 As Object
 	Set rWorstEye1 = DS.result1d("")
	Dim rWorstEye0 As Object
	Set rWorstEye0 = DS.result1d("")
	' Variables For the Loop
	Dim sum_neg_isi As Double
	Dim sum_pos_isi As Double
    Dim t_mine As Double
	Dim ind As Integer
    Dim cnt As Integer
    Dim nn As Long  'rsj:change to long to avoid overflow
    Dim cnt2 As Integer
    Dim k As Integer

    Dim cursor As Integer
    Dim cursor_tmp As Integer
	Dim K_Tr As Integer

	Dim nap As Double
	Dim na As Double
	Dim cnt3 As Double
	'Dim UI As Double         'Store the unit bit interval T_r+T_hold, now user-defined
	Dim UIs As Double         'Total number of UIs


	Dim delta_t As Double    'Time step ~1/10 of rise time

	Dim t_pda As Double
    Dim tmax As Double

	Dim w_c_11 As Double
    Dim w_c_00 As Double

    'Initiallize some variables and accumulators
    sum_neg_isi=0
    sum_pos_isi=0
    ind=0
    nap=1
    na=0
    cnt3=0

    k=36            'ToDo: Make the K a variable, it is the number of samples per UI!
    K_Tr=12         'ToDO: Samples per T_rise
   ' UI= 0.36       'User-defined
	cursor=k*3      'Initial starting point for the cursor, recalculated later

	nn=rImpulseResponse.GetN             ' Get number of elements
	tmax=rImpulseResponse.GetX(nn-1)     ' Get tmax simulation time

	UIs=Round((tmax/UI)-0.5)                        ' It rounds to the maximum number of UIs
	rImpulseResponse.ResampleTo(0, UI*UIs, k*UIs+1) ' Resample to get a uniform time step
	nn=rImpulseResponse.GetN                        ' Get number of elements after resampling


	Dim w_c_p1(2048) As Integer    ' These variables store the worst case bit patterns, 01011...
    Dim w_c_p0(2048) As Integer

	cursor_tmp=rImpulseResponse.GetGlobalMaximum   'Points to the index of maximum in pulse response
	cursor=((Round((cursor_tmp/k)-0.5))-1)*k       ' Moves the cursor to the second nearest UI start point


 For cnt=cursor To cursor+3*(k)           'covers Three UIs, cnt tracks the cursor position
       For cnt2=ind To nn-1 Step k       ' This loop goes over the various UI bins in the pulse response
            If (cnt2 <> cnt) Then         ' The UI bin containing the cursor [y(t)] is treated differently
                If (rImpulseResponse.GetY(cnt2) > 0) Then
                    sum_pos_isi=sum_pos_isi+rImpulseResponse.GetY(cnt2) 'Imp_response(cnt2);
                    cnt3=cnt3+1
                    If (cnt=cursor_tmp) Then  'The worst case bit patterns are determined when cursor is at maximum
						w_c_p1(nap)=0
						w_c_p0(nap)=1
                    End If
                Else
                    sum_neg_isi=sum_neg_isi+rImpulseResponse.GetY(cnt2) 'Imp_response(cnt2);
					If (cnt=cursor_tmp) Then  'The worst case bit patterns are determined when cursor is at maximum
						w_c_p1(nap)=1
                        w_c_p0(nap)=0
                    End If
                End If
            Else
                If (cnt=cursor_tmp) Then       'The worst case bit patterns are determined when cursor is at maximum
					w_c_p1(nap)=1
                    w_c_p0(nap)=0
                End If
            End If
            nap=nap+1
        Next

        ind=ind+1
        If (ind = k) Then
        	ind=0              ' Reset ind after each UI calcuation
        End If
		t_pda=rImpulseResponse.GetX(cnt-cursor)    ' So signal is shifted to first three UIs in time
        w_c_11=rImpulseResponse.GetY(cnt)+sum_neg_isi
        w_c_00=0+sum_pos_isi
        sum_neg_isi=0
        sum_pos_isi=0
        nap=1
		'Add again the For loop for the Xtalk

		rWorstEye1.AppendXY(t_pda,w_c_11)
        rWorstEye0.AppendXY(t_pda,w_c_00)
 Next
    '	sum_neg_isi =rImpulseResponse.GetX(0)

	rWorstEye1.Save("WorstEye1")
	rWorstEye0.Save("WorstEye0")
	rWorstEye1.AddToTree("Results\PDA - WorstEye\WorstEye1") ' Creates a folder under DS/Results and stores it there
	rWorstEye0.AddToTree("Results\PDA - WorstEye\WorstEye0") ' Creates a folder under DS/Results and stores it there

    ' /// END OF THE MAIN PART ////


	' Work with the worst case bit patterns
	' first invert

	Dim w_c_bit_pat_1(2048) As Integer
    Dim w_c_bit_pat_0(2048) As Integer
	Dim i As Integer


	'Invert and write worst bit patterns
	For i=1 To UIs
		w_c_bit_pat_1(i)=w_c_p1(UIs+1-i)
        w_c_bit_pat_0(i)=w_c_p0(UIs+1-i)
	Next
	'w_c_bit_pat_1(2)=1


	Dim time_scale As Double
	time_scale=1e-9 'ASCII file for use in DS needs Time scale in seconds
	'Generate time signal

	cnt2=0  ' Initiallize counter

	'set and open file for output
	Open "WC_1.txt" For Output As #i

	'Generate time signal
	Dim Last As Integer   ' Contains the "Last" (n-1) bit info to manage the transition when generating the time signal below

	' Make first bit equal to last so time signal does not start with a transition
	If w_c_bit_pat_1(1)=1 Then
		Last =1
	Else
		Last =0
	End If

  	For cnt=1 To UIs
		If (Last=0) And w_c_bit_pat_1(cnt)=1 Then   '0 to 1
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*(cnt3/(K_Tr))
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
					cnt2=cnt2+1
				End If
			Next
		ElseIf (Last=0) And w_c_bit_pat_1(cnt)=0 Then   '0 to 0
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_1(cnt)=1 Then   '1 to 1
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_1(cnt)=0 Then   '1 to 0
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*((K_Tr-cnt3)/K_Tr)
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
					cnt2=cnt2+1
				End If
			Next
        End If
		Last = w_c_bit_pat_1(cnt)
	Next

	Close #i  'Close file

	cnt2=0 'Initiallize counter

	' Worst case 0
	Open "WC_0.txt" For Output As #i
	'Generate time signal

	' Make first bit equal to last so time signal does not start with a transition
	If w_c_bit_pat_0(1)=1 Then
		Last =1
	Else
		Last =0
	End If

  	For cnt=1 To UIs
		If (Last=0) And w_c_bit_pat_0(cnt)=1 Then   '0 to 1
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*(cnt3/(K_Tr))
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
					cnt2=cnt2+1
				End If
			Next
		ElseIf (Last=0) And w_c_bit_pat_0(cnt)=0 Then   '0 to 0
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_0(cnt)=1 Then   '1 to 1
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_0(cnt)=0 Then   '1 to 0
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*((K_Tr-cnt3)/K_Tr)
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
					cnt2=cnt2+1
				End If
			Next
        End If
		Last = w_c_bit_pat_0(cnt)
	Next
	Close #i  'Close file

	Open "WC_bits.txt" For Output As #i
	'Write worst bit patterns, first colum is 1's, second column is 0's
	For cnt=1 To UIs
		Print #i, w_c_bit_pat_1(cnt); w_c_bit_pat_0(cnt)
	Next
	Close #i


	'Thi section generates a pulse response for use with DS.
	'We only need one bit to generate the pulse, so reset the rest to zero
	For i=1 To UIs
		w_c_bit_pat_1(i)=0
	Next
	w_c_bit_pat_1(2)=1    'Second bit is set to 1, so Tdelay is one UI

	cnt2=0  ' Initiallize counter

	'set and open file for output
	Open "In_pulse.txt" For Output As #i

	' Make first bit equal to last so time signal does not start with a transition
	If w_c_bit_pat_1(1)=1 Then
		Last =1
	Else
		Last =0
	End If

  	For cnt=1 To UIs
		If (Last=0) And w_c_bit_pat_1(cnt)=1 Then   '0 to 1
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*(cnt3/(K_Tr))
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
					cnt2=cnt2+1
				End If
			Next
		ElseIf (Last=0) And w_c_bit_pat_1(cnt)=0 Then   '0 to 0
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_1(cnt)=1 Then   '1 to 1
			For cnt3=0 To (k-1)
				Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1
				cnt2=cnt2+1
			Next
		ElseIf (Last=1) And w_c_bit_pat_1(cnt)=0 Then   '1 to 0
			For cnt3=0 To (k-1)
				If cnt3 <= K_Tr Then
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 1*((K_Tr-cnt3)/K_Tr)
					cnt2=cnt2+1
				Else
					Print #i, time_scale*rImpulseResponse.GetX(cnt2); 0
					cnt2=cnt2+1
				End If
			Next
        End If
		Last = w_c_bit_pat_1(cnt)
	Next
	Close #i  'Close file


End Sub

Sub ReadASCIIData (sfilepath As String,ASCIIData As Object)
	Set ASCIIData = DS.Result1d("")
	Dim inline As String

	'rsj: Scale the time signal unit
	Dim ScaleFactor As Double
	Select Case sTimeUnit
		Case "ns" 'Nanosecond
			ScaleFactor=Units.GetTimesitounit*1e-9
		Case "s" 'Second
			ScaleFactor=Units.GetTimesitounit
		Case "us" 'Microsecond
			ScaleFactor=Units.gettimesitounit*1e-6
		Case "ps" 'Picosecond
			ScaleFactor=Units.gettimesitounit*1e-12
		Case "ms" 'Microsecond
			ScaleFactor=Units.gettimesitounit*1e-3
		Case "fs" 'Picosecond
			ScaleFactor=Units.gettimesitounit*1e-15
	End Select

	'rsj:Now scale the time unit with proper factor
	Open sfilepath For Input As #1
	While Not EOF(1)
		Line Input #1, inline
		ASCIIData.appendxy(Cdbl(Split(inline)(0))*ScaleFactor,Cdbl(Split(inline)(1)))
	Wend
	Close #1

	'Now plot the input data!
	inline=Mid(sfilepath,InStrRev(sfilepath,"\")+1,InStrRev(sfilepath,".")-InStrRev(sfilepath,"\")-1)
	ASCIIData.SetXLabelAndUnit("Time" , Units.GetUnit("Time"))
	ASCIIData.Save(inline)
	ASCIIData.AddToTree("Results\PDA - TimeSignal\"+inline)

End Sub


Sub DataInput

	If hlpcnt=0 Then
		ReDim Preserve FileXtalkList(1)
		FileInputResponse=""
		FileXtalkList(0)=""
	End If

	Begin Dialog UserDialog 580,259,"Input File",.DialogFunction3 ' %GRID:10,7,1,1
		GroupBox 20,14,530,70,"Response file path",.GroupBox1
		GroupBox 20,91,530,133,"XTalks file path",.GroupBox2
		OKButton 20,231,90,21,.Ok
		CancelButton 130,231,90,21
		TextBox 40,42,400,21,.PulseFile
		PushButton 450,42,90,21,"Browse..",.Browse
		PushButton 450,133,90,21,"Add",.AddFile
		PushButton 450,168,90,21,"Remove",.RemoveFile
		ListBox 40,119,400,91,FileXtalkList(),.Xtalkfilename,1
	End Dialog
	Dim dlg_Inputfile As UserDialog
	If (Dialog(dlg_Inputfile) = 0) Then 'Do Nothing
	End If


End Sub


Rem See DialogFunc help topic for more information.
Private Function DialogFunction3(DlgItem$, Action%, SuppValue?) As Boolean
	Dim sFile As String
	Dim sRootPath As String
	Dim temp() As String
	Dim ss As Integer, mm As Integer
	mm=1

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "PulseFile",FileInputResponse

	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "Browse"
			sRootPath=GetProjectPath("Root")
			sFile  = GetFilePath("", "sig;txt", sRootPath, "Browse pulse response", 0)' Let user browse for file
 	 		If (sFile <> "") Then
                DlgText "PulseFile", sFile ' store filename in field. TODO cut path
                FileInputResponse=sFile     ' Response file is stored with index 0
                hlpcnt=hlpcnt+1
	        End If
			DialogFunction3=True
		Case "AddFile"
			sRootPath=GetProjectPath("Root")
			sFile  = GetFilePath("", "sig;txt", sRootPath, "Browse pulse response", 0)
			If (sFile <> "") Then
				ReDim Preserve FileXtalkList(hlpcnt)
				FileXtalkList(hlpcnt)=sFile
				hlpcnt=hlpcnt+1
			End If
			DlgListBoxArray "Xtalkfilename", FileXtalkList()
			DialogFunction3=True
		Case "RemoveFile"
			sRootPath=DlgText "Xtalkfilename"
			For ss=1 To UBound(FileXtalkList) STEP 1
				If FileXtalkList(ss) <> sRootPath Then
				   ReDim Preserve temp(mm)
				   temp (mm) = FileXtalkList (ss)
				   mm=mm+1
				Else
				   'Do nothing
				End If
			Next
			ReDim FileXtalkList(mm-1)
			FileXtalkList=temp
			DlgListBoxArray "Xtalkfilename", FileXtalkList()
			DialogFunction3=True
		Case "Ok"
			If FileInputResponse = "" Then
				MsgBox "Please specify the response file input"
				DialogFunction3=True
			End If
		End Select
		Rem DialogFunction3 = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction3 = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


Sub CalculatePulseRespondFromSPara (flag As Boolean)
	Dim ocom As Object
	Dim oam() As Object
	Dim oph() As Object
	Dim o2new() As Object
	Dim oooo As Object
	Dim counter As Integer
	Dim temp As Integer
	Dim cst_sPath As String
	Dim TotXtalk As Object
	Set TotXtalk = DS.result1d("")

	If  SelectedArray(1)<>"" Then
		counter=UBound(SelectedArray)
	Else
		counter=0 'No xtalk s-para is selected or available
	End If

	Set oooo = DS.result1d("")
	Dim TimeSignalFromSPara As Object
	Set TimeSignalFromSPara= DS.result1d("")

	'Define the input file
	Dim i1 As Object
	Dim iam As Object, iph As Object
	Dim dt As Double
	Set iam = DS.result1d("")
	Set iph = DS.result1d("")
	Set i1 = DS.result1d("")

	'setup the input signal. 1 bit signal
	i1.initialize(1)
	i1.appendxy(0.0,0.0)
	i1.appendxy(sTrise,1)
	i1.appendxy(sThold+sTrise,1)
	i1.appendxy(sThold+2*sTrise,0)

	'This part is only used to window the time signal, avoid unnecessary zero signal
	Dim tmax As Double
	Dim IndexWindow As Long

	If flag=True Then
		UI=sTrise+sThold
	End If

	For temp=0 To counter STEP 1        'Here, zero index is used to store Pulse response signal,non-zero index for crosstalk signal
		ReDim Preserve oam(temp+1)
		ReDim Preserve oph(temp+1)
		ReDim Preserve o2new(temp+1)
		Set o2new(temp)= DS.result1d("")
		If temp=0 Then
			Set ocom = DS.Result1DComplex(DSResultTree.GetFileFromTreeItem(sPath))
			Set oam(temp) = ocom.Magnitude
			Set oph(temp) = ocom.Phase
		Else
			cst_sPath=Left(sPath,InStrRev(sPath,"\"))+SelectedArray(temp)
			Set ocom = DS.Result1DComplex(DSResultTree.GetFileFromTreeItem(cst_sPath))
			Set oam(temp) = ocom.Magnitude
			Set oph(temp) = ocom.Phase
		End If

		'IFFT uses 1024 number of samples
		calculateifft(oam(temp),oph(temp),TimeSignalFromSPara)

		If TimeSignalFromSPara.getx(TimeSignalFromSPara.getn-1)-TimeSignalFromSPara.getx(0)>TimeSignalFromSPara.getn Then
			'upsampling For better resolution, happens when the Time signal Is quite Long
			TimeSignalFromSPara.ResampleTo(TimeSignalFromSPara.getx(0),TimeSignalFromSPara.getx(TimeSignalFromSPara.getn-1),2*Round(TimeSignalFromSPara.getx(TimeSignalFromSPara.getn-1)-TimeSignalFromSPara.getx(0)))
		Else
			'Fix the samples to 5000.
			TimeSignalFromSPara.ResampleTo(TimeSignalFromSPara.getx(0),TimeSignalFromSPara.getx(TimeSignalFromSPara.getn-1),5000)
		End If
		i1.resampleto(TimeSignalFromSPara.getx(0),TimeSignalFromSPara.getx(TimeSignalFromSPara.getn-1),TimeSignalFromSPara.getn)

	    'calculation the convolution input signal and transfer function time signal
		CalculateCONV(i1,TimeSignalFromSPara,o2new(temp))

		'It happens sometimes, that the time signal response is flipped 180°
		'here i just multiply it again with -1 -> Hope it works for many cases

		If Abs(o2new(temp).gety(o2new(temp).GetGlobalMaximum))-Abs(o2new(temp).gety(o2new(temp).GetGlobalMinimum))<0 Then
			o2new(temp).ScalarMult(-1.0)
		End If


		If temp=0 Then
			'TODO LIST!!
			'Cut the tail for zero time signal
			'Perform this step only once, to make sure that all the time signal will have the same samples and length
			'Use 6 times signal length based on the echo time (12x transmission)-> should be sufficient
			tmax= 6 * o2new(temp).getx(o2new(temp).GetGlobalMaximum)
			IndexWindow=o2new(temp).GetClosestIndexFromX(tmax)
			If tmax<o2new(temp).getx(o2new(temp).getn-1) Then
				o2new(temp).resampleto(o2new(temp).getx(0),o2new(temp).getx(IndexWindow),o2new(temp).getn)
				i1.Makecompatibleto(o2new(temp))
			Else
				'no resample. Original data is taken
			End If

			'
			i1.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
			'i1.save(getprojectpath("ResultDSTask")+"PDA_excitation.txt")
			ExportToASCII(i1,getprojectpath("TempDS")+"PDA_excitation.txt")
			i1.addtotree("Results\PDA - TimeSignal\Excitation")

			'o2new(temp).save(getprojectpath("ResultDSTask")+"PDA_responses.txt")          '!-> Stupid VBA doesnt work for this save command in PCBS and CBL Studio!
			ExportToASCII(o2new(temp),getprojectpath("TempDS")+"PDA_responses.txt")  'Replace with manual write file
			oooo.loadplainfile(getprojectpath("TempDS")+"PDA_responses.txt")                        'Need this trick, otherwise result is not added in DS Tree, stupid vba!
			oooo.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
			oooo.addtotree("Results\PDA - TimeSignal\Responses")
		Else
			o2new(temp).resampleto(o2new(temp).getx(0),o2new(temp).getx(IndexWindow),o2new(temp).getn)
			'o2new(temp).resampleto(o2new(temp).getx(0),o2new(temp).getx(o2new(temp).getn-1),2*o2new(temp).getn)
			'o2new(temp).save(getprojectpath("ResultDSTask")+"PDA_responses_"+SelectedArray(temp)+".txt")
			ExportToASCII(o2new(temp),getprojectpath("TempDS")+"PDA_responses_"+SelectedArray(temp)+".txt")
			oooo.loadplainfile(getprojectpath("TempDS")+"PDA_responses_"+SelectedArray(temp)+".txt")'Need this trick, otherwise result is not added in DS Tree, stupid vba!
			oooo.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
			oooo.addtotree("Results\PDA - TimeSignal\Responses "+SelectedArray(temp))
			'If temp=1 Then
			'   TotXtalk.initialize(o2new(temp).getn)
			'   Set TotXtalk=o2new(temp).Copy  'Need the x range
			'   TotXtalk.ScalarMult(0.0) 'Initialize
			'End If
			'TotXtalk.Add(o2new(temp))
		End If
	Next

	'Now before exit, store the total xtalk signal
	'If counter<>0 Then TotXtalk.save(getprojectpath("ResultDSTask")+"PDA_totxtalk.txt")
End Sub

Function CalculatePDAFromSParameter (sFname As String,flag As Boolean) As Boolean
	'RSJ: This part is new and consider all the crosstalks
	'     Slightly different algorithm as the one from Gerardo --> CalculatePDA

	'Initialization for ISI and XTalk
	Dim rImpulseResponse As Object
 	Set rImpulseResponse = DS.result1d("") ' DSResults1D, otherwise addtotree adds into MWS

	Dim rImpulseResponseXT() As Object
	Dim NumXtalk As Integer
	Dim cnt3 As Integer

 	' RSJ: Response file from s-para or external ASCII file
	If flag = True Then
		UI=sTrise+sThold
		rImpulseResponse.LoadPlainFile(sFname) ' PlainFile, since it could be from another project.
		If SelectedArray(1)="" Then
			'No XTalk is neither selected nor available
			NumXtalk=0
		Else
			NumXtalk=UBound(SelectedArray)
			ReDim rImpulseResponseXT(NumXtalk)
			For cnt3=1 To NumXtalk Step 1      '0 Index is not included as it is not xtalk
				Set rImpulseResponseXT(cnt3) = DS.result1d("")
				rImpulseResponseXT(cnt3).LoadPlainFile(getprojectpath("TempDS")+"PDA_responses_"+SelectedArray(cnt3)+".txt")
			Next
		End If
	Else
		If FileInputResponse="" Then
			MsgBox "Please specify the input file"
			CalculatePDAFromSParameter=False
			Exit Function
		End If

		ReadASCIIData(FileInputResponse,rImpulseResponse)
		If FileXtalkList(1)="" Then
			'No XTalk is neither selected nor available
			NumXtalk=0
		Else
			NumXtalk=UBound(FileXtalkList)
			ReDim rImpulseResponseXT(NumXtalk)
			For cnt3=1 To NumXtalk Step 1     	 '0 Index is not included
				Set rImpulseResponseXT(cnt3) = DS.result1d("")
				ReadASCIIData(FileXtalkList(cnt3),rImpulseResponseXT(cnt3))
			Next
		End If
	End If

	Dim nsamples As Long
	Dim iii As Long,kkk As Long,jjj As Long
	Dim index_per_UI As Long
	Dim index_maxy As Long
	Dim starting_index As Long
	Dim stop_index As Long

	jjj=0
	nsamples=rImpulseResponse.getn
	index_per_UI=rImpulseResponse.GetClosestIndexFromX(UI)
	index_maxy=rImpulseResponse.GetGlobalMaximum

	'28.12.2012: rsj: Add msg box in case the response signal is overlapping with excitation signal
	'09.01.2012: rsj: Removed as it is not necessary since the UI is now starting from 10% signal

'	Dim msg As String
'	msg="Channel length is too short. This might lead to inaccurate result."+vbNewLine+"Please extend the channel length using ideal transmission line block."
'	If Round(index_maxy/index_per_UI) < 4 Then  'Set to 4*UI for a better results.
'		If MsgBox (msg+vbNewLine+vbNewLine+"Press Ok to continue or Cancel to abort.",vbOkCancel,"Peak Distortion Analysis")=vbOK Then
'			'Do nothing
'		Else
'			Exit All
'		End If
'	End If


	' 3 UI is used, which is started around Max Value
	'starting_index=Round(index_maxy/index_per_UI)*index_per_UI-index_per_UI
	'stop_index=starting_index+2*index_per_UI

	'3UI with starting index is determined from the 10% of signal response, not from MaxValue
	starting_index=GetClosestIndexFromY(0.1,rImpulseResponse)
	stop_index=starting_index+2*index_per_UI
	If stop_index>nsamples Then stop_index=nsamples-1

	Dim ISI0 As Double
	Dim ISI1 As Double
	Dim XT0 As Double
	Dim XT1 As Double

	ISI0=0.0
	ISI1=0.0
	XT0=0.0
	XT1=0.0

	'ISI Results
	Dim rWorstEye1 As Object
 	Set rWorstEye1 = DS.result1d("")
	Dim rWorstEye0 As Object
	Set rWorstEye0 = DS.result1d("")

	'XTALK Results
	Dim rWorstEye2 As Object
	Set rWorstEye2 = DS.result1d("")
	Dim rWorstEye3 As Object
	Set rWorstEye3 = DS.result1d("")

	For iii=starting_index To stop_index Step 1
		For kkk=1 To nsamples-1 Step index_per_UI
			If kkk+jjj>nsamples-1 Then
				Exit For
			End If
			'ISI PART
			If (rImpulseResponse.GetY(kkk+jjj) > 0) Then
				ISI0=ISI0+rImpulseResponse.GetY(kkk+jjj)
			Else
				ISI1=ISI1+rImpulseResponse.gety(kkk+jjj)
			End If

			'XTALK PART, "for" loop each aggressor
			For cnt3=1 To NumXtalk Step 1
				If rImpulseResponseXT(cnt3).gety(kkk+jjj)>0 Then
					XT0=XT0+rImpulseResponseXT(cnt3).GetY(kkk+jjj)
				Else
					XT1=XT1+rImpulseResponseXT(cnt3).gety(kkk+jjj)
				End If
			Next
		Next

		rWorstEye0.appendxy(jjj*(UI/index_per_UI),ISI0-rImpulseResponse.GetY(iii))
		rWorstEye1.appendxy(jjj*(UI/index_per_UI),ISI1+rImpulseResponse.GetY(iii))
		rWorstEye2.appendxy(jjj*(UI/index_per_UI),ISI0-rImpulseResponse.GetY(iii)+XT0)
    	rWorstEye3.appendxy(jjj*(UI/index_per_UI),ISI1+rImpulseResponse.GetY(iii)+XT1)

		ISI1=0.0
		ISI0=0.0
		XT0=0.0
		XT1=0.0
		jjj=jjj+1
	Next

	'Just cut the tail. The PDA xrange will be kept within the UI Range
	'Fix the number of samples for all PDA results. 1001 Points should be sufficient

'	rWorstEye0.ResampleTo(0,2*UI,rWorstEye0.getn)
'	rWorstEye1.ResampleTo(0,2*UI,rWorstEye1.getn)
'	rWorstEye2.ResampleTo(0,2*UI,rWorstEye2.getn)
'   rWorstEye3.ResampleTo(0,2*UI,rWorstEye3.getn)


	'rsj: 09.01.2012 : This part is removed and being replaced in the next line

'	rWorstEye1.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
'	rWorstEye0.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
'	rWorstEye1.Save(getprojectpath("TempDS")+"WorstEye1")
'	rWorstEye0.Save(getprojectpath("TempDS")+"WorstEye0")
'	rWorstEye1.AddToTree("Results\PDA - WorstEye\WorstEyeISI1") ' Creates a folder under DS/Results and stores it there
'	rWorstEye0.AddToTree("Results\PDA - WorstEye\WorstEyeISI0") ' Creates a folder under DS/Results and stores it there


'	rWorstEye3.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
'	rWorstEye2.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
'	rWorstEye3.Save(getprojectpath("TempDS")+"WorstEyeXT1")
'	rWorstEye2.Save(getprojectpath("TempDS")+"WorstEyeXT0")
'	rWorstEye3.AddToTree("Results\PDA - WorstEye\WorstEyeXT1") ' Creates a folder under DS/Results and stores it there
'	rWorstEye2.AddToTree("Results\PDA - WorstEye\WorstEyeXT0") ' Creates a folder under DS/Results and stores it there

	'ISI and Xtalks height function
	Dim EyeHeightISI As Object
	Set EyeHeightISI = DS.result1d("")
	Dim EyeHeightXT As Object
	Set EyeHeightXT = DS.result1d("")

	Set EyeHeightXT=rWorstEye3.copy
	EyeHeightXT.Subtract(rWorstEye2)
	Set EyeHeightISI=rWorstEye1.copy
	EyeHeightISI.Subtract(rWorstEye0)

	EyeHeightXT.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	EyeHeightISI.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	EyeHeightXT.save (getprojectpath("TempDS")+"EyeHeightXT.txt")
	EyeHeightISI.save (getprojectpath("TempDS")+"EyeHeightISI.txt")
	EyeHeightISI.AddToTree ("Results\PDA - Eye Height Function\Eye Height ISI")
	EyeHeightXT.AddToTree ("Results\PDA - Eye Height Function\Eye Height XT")


	'rsj: 09.01.2012 : Improve the worsteye plot for better visualization. Cut out the unnecessary part.

	Dim index_Null1 As Double
	Dim index_Null2 As Double
	Dim eye_width_ISI As Double
	Dim eye_width_XT As Double
	Dim shiftx As Double
	Dim temp0 As Object
	Set temp0 = DS.Result1d("")
	Dim temp1 As Object
	Set temp1 = DS.Result1d("")
	Dim temp2 As Object
	Set temp2 = DS.Result1d("")
	Dim temp3 As Object
	Set temp3 = DS.Result1d("")

	Dim ii As Long

	'ISI Part
	For ii=0 To EyeHeightISI.GetClosestIndexFromX(UI)-1
		If EyeHeightISI.gety(ii)<0.0 And EyeHeightISI.gety(ii+1)>0.0 Then
			index_Null1=ii
			Exit For
		End If
	Next
	For ii=EyeHeightISI.GetClosestIndexFromX(UI) To EyeHeightISI.GetClosestIndexFromX(2*UI)-1
		If EyeHeightISI.gety(ii)>0.0 And EyeHeightISI.gety(ii+1)<0.0 Then
			index_Null2=ii+1
			Exit For
		End If
	Next

	' Before plotting it, center the PDA plot for better visualization as in the new eye diagram plot
	eye_width_ISI=EyeHeightISI.getx(index_Null2)- EyeHeightISI.getx(index_Null1)
	shiftx=(2*UI-eye_width_ISI)/2-EyeHeightISI.getx(index_Null1)
	'Cut the tail and force to have the same samples
	rWorstEye0.ResampleTo(rWorstEye0.getx(index_Null1),rWorstEye0.getx(index_Null2),rWorstEye0.getn)
	rWorstEye1.ResampleTo(rWorstEye1.getx(index_Null1),rWorstEye1.getx(index_Null2),rWorstEye0.getn)
	'Center the PDA Plot
	For ii=0 To rWorstEye0.getn-1 Step 1
		temp0.appendxy(rWorstEye0.getx(ii)+shiftx,rWorstEye0.gety(ii))
		temp1.appendxy(rWorstEye1.getx(ii)+shiftx,rWorstEye1.gety(ii))
	Next
	Set rWorstEye0=temp0.copy
	Set rWorstEye1=temp1.copy


	rWorstEye1.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	rWorstEye0.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	rWorstEye1.Save(getprojectpath("TempDS")+"WorstEye1")
	rWorstEye0.Save(getprojectpath("TempDS")+"WorstEye0")
	rWorstEye1.AddToTree("Results\PDA - WorstEye\WorstEyeISI1") ' Creates a folder under DS/Results and stores it there
	rWorstEye0.AddToTree("Results\PDA - WorstEye\WorstEyeISI0") ' Creates a folder under DS/Results and stores it there

	'XT Part
	For ii=0 To EyeHeightXT.GetClosestIndexFromX(UI)-1
		If EyeHeightXT.gety(ii)<0.0 And EyeHeightXT.gety(ii+1)>0.0 Then
			index_Null1=ii
			Exit For
		End If
	Next
	For ii=EyeHeightXT.GetClosestIndexFromX(UI) To EyeHeightXT.GetClosestIndexFromX(2*UI)-1
		If EyeHeightXT.gety(ii)>0.0 And EyeHeightXT.gety(ii+1)<0.0 Then
			index_Null2=ii+1
			Exit For
		End If
	Next

	' Before plotting it, center the PDA plot for better visualization as in the new eye diagram plot
	eye_width_XT=EyeHeightXT.getx(index_Null2)- EyeHeightXT.getx(index_Null1)
	shiftx=(2*UI-eye_width_XT)/2-EyeHeightXT.getx(index_Null1)
	'Cut the tail and force to have the same samples
	rWorstEye2.ResampleTo(rWorstEye2.getx(index_Null1),rWorstEye2.getx(index_Null2),rWorstEye2.getn)
	rWorstEye3.ResampleTo(rWorstEye3.getx(index_Null1),rWorstEye3.getx(index_Null2),rWorstEye2.getn)

	'Center the PDA Plot
	For ii=0 To rWorstEye2.getn-1 Step 1
		temp2.appendxy(rWorstEye2.getx(ii)+shiftx,rWorstEye2.gety(ii))
		temp3.appendxy(rWorstEye3.getx(ii)+shiftx,rWorstEye3.gety(ii))
	Next
	Set rWorstEye2=temp2.copy
	Set rWorstEye3=temp3.copy

	rWorstEye3.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	rWorstEye2.SetXLabelAndUnit ("Time" , Units.GetUnit("Time"))
	rWorstEye3.Save(getprojectpath("TempDS")+"WorstEyeXT1")
	rWorstEye2.Save(getprojectpath("TempDS")+"WorstEyeXT0")
	rWorstEye3.AddToTree("Results\PDA - WorstEye\WorstEyeXT1") ' Creates a folder under DS/Results and stores it there
	rWorstEye2.AddToTree("Results\PDA - WorstEye\WorstEyeXT0") ' Creates a folder under DS/Results and stores it there

	CalculatePDAFromSParameter=True
End Function


Function SParaList As String
	Dim sTreeName As String
	Dim InputPort As Integer
	Dim TempInputPort As Integer
	Dim OutputPort As Integer
	Dim TempOutputPort As Integer
	Dim NumPort As Integer
	Dim schild1 As String
	Dim snextitem As String
	Dim Sparastring As String
	Dim Tempstring As String
	NumPort=0

	Sparastring=Right(sPath,Len(sPath)-(InStrRev(sPath,"\")+1))  '+1 to remove the S char
	InputPort=Cint(Right(Sparastring,Len(Sparastring)-InStrRev(Sparastring,",")))
	OutputPort=Cint(Left(Sparastring,Len(Sparastring)-InStrRev(Sparastring,",")))

	sTreeName=Left(sPath,InStrRev(sPath,"\")-1)
	schild1=DSResultTree.GetFirstChildName(sTreeName)

	While schild1 <> ""
		Tempstring=Right(schild1,Len(schild1)-(InStrRev(schild1,"\")+1))
		TempOutputPort=Cint(Left(Tempstring,Len(Tempstring)-InStrRev(Tempstring,",")))
		TempInputPort=Cint(Right(Tempstring,Len(Tempstring)-InStrRev(Tempstring,",")))

		'Consider only relevant Xtalk and return loss
		If TempOutputPort=OutputPort And TempInputPort<>InputPort Then
			NumPort=NumPort+1
			ReDim Preserve sSparaList(NumPort)
			sSparaList(NumPort)=Right(schild1,Len(schild1)-(InStrRev(schild1,"\")))
		End If
		snextitem=DSResultTree.GetNextItemName(schild1)
		schild1=snextitem
	Wend

End Function

Function DialogCrossTalkList ()

	If ii=0 Then
		ReDim Preserve SelectedArray (1)
		TempSparaList=sSparaList
	End If

	Begin Dialog UserDialog 370,203,"S-Parameter for Crosstalk",.DialogFunction2 ' %GRID:10,7,1,1
		OKButton 80,168,90,21,.OK
		CancelButton 210,168,90,21
		PushButton 160,63,50,21,"->",.Add
		PushButton 160,98,50,21,"<-",.Remove
		ComboBox 30,14,110,140,TempSparaList(),.OriginalSPara
		ComboBox 230,14,110,140,SelectedArray(),.SelectedSPara
	End Dialog
	Dim dlg_Crosstalk As UserDialog
	If (Dialog(dlg_Crosstalk) = 0) Then 'Do Nothing
	End If

End Function


Rem See DialogFunc help topic for more information.
Private Function DialogFunction2(DlgItem$, Action%, SuppValue?) As Boolean
	Dim jj As Integer,kk As Integer
	kk = 1
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
		Case "Add"
		ADD:
			If DlgValue("OriginalSPara")<>"" Then
				ReDim Preserve SelectedArray(ii+1)
				SelectedArray(UBound(SelectedArray))=TempSparaList(DlgValue("OriginalSPara")+1)
				For jj=1 To UBound(TempSparaList) STEP 1
					If jj<>DlgValue("OriginalSPara")+1 And UBound(TempSparaList)<>1 Then
						ReDim Preserve temp(kk)
						temp(kk)=TempSparaList(jj)
						kk=kk+1
					End If
					If UBound(TempSparaList)=1 Then  'Special handling for the last entry
						ReDim temp(1)
						temp(0)=""
					End If
				Next
				DlgListBoxArray "SelectedSPara",SelectedArray()
				ReDim TempSparaList(UBound(temp))
				TempSparaList=temp
				DlgListBoxArray "OriginalSPara",TempSparaList()
				ii=ii+1
			Else
			 	'do nothing
			End If
			DialogFunction2=True
		Case "Remove"
		REMOVE:
			If DlgValue("SelectedSPara")<>"" Then
				ReDim Preserve TempSparaList(UBound(TempSparaList)+1)
				TempSparaList(UBound(TempSparaList))=SelectedArray(DlgValue("SelectedSPara")+1)
				For ll=1 To UBound(SelectedArray) STEP 1
					If ll<>DlgValue("SelectedSPara")+1 And UBound(SelectedArray)<>1 Then
						ReDim Preserve temp(kk)
						temp(kk)=SelectedArray(ll)
						kk=kk+1
					End If
					If UBound(SelectedArray)=1 Then
						ReDim temp(1)
						temp(0)=""
					End If
				Next
				DlgListBoxArray "OriginalSPara",TempSparaList()
				ReDim SelectedArray(UBound(temp))
				SelectedArray=temp
				DlgListBoxArray "SelectedSPara",SelectedArray()
				ll=ll+1
			Else
				'do nothing
			End If
			DialogFunction2=True
		'Handling of double click --> avoid it
		Case "OK"
			If DlgValue("SelectedSPara")=-1 And DlgValue("OriginalSPara")=-1 Then
				'Do nothing, the real "ok" button is pressed
			Else
			  'If DlgValue("SelectedSPara")<>"" And DlgValue("SelectedSPara")<>-1 Then
			  '   GoTo REMOVE  'right combobox's entry is double clicked
			  '  Else
			  ' 	   GoTo ADD     'left combobox's entry is double clicked
			  ' End If
			DialogFunction2=True 'Do nothing, disable the double click. Can be complicated. The above logic doesn't work if 2 comboxes are selected.
			End If
		End Select
		Rem DialogFunction2 = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction2 = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub ExportToASCII(temp As Object,filepath As String)

	Dim counter As Long

	Open filepath For Output As #2
	'Open "C:\Users\RichardSjiariel\Desktop\test.txt" For Output As #2

	For counter=0 To temp.getn-1 STEP 1
		Print #2, Cstr(temp.getx(counter))+"   "+Cstr(temp.gety(counter))
	Next

	Close #2

End Sub

Function GetClosestIndexFromY(YValue As Double, temp As Object) As Integer
	Dim counter As Long
	Dim Yindex As Long
	For counter=0 To temp.getn-1
		If temp.gety(counter) > YValue Then
			Exit For
		End If
	Next
	GetClosestIndexFromY=counter-1
End Function
