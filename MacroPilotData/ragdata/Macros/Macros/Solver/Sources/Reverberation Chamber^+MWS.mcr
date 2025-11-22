' Run superposition solver for Reverberation chamber application
Option Explicit

' ================================================================================================
' Copyright 2024-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------------------------------------------------------------
' 30-Oct-2024 ech: Initial version
' ------------------------------------------------------------------------------------------------------------------------------------------------------


Const HelpFileName = "common_preloadedmacro_solver_reverberation_chamber"
Public sSampleStepSelector() As String
Public Const SampleStepSelectorArray = Array("Freq. samples", _
											 "Freq. step")

Sub Main ()

	Dim i As Long
	ReDim sSampleStepSelector(UBound(SampleStepSelectorArray))

	For i = 0 To UBound(sSampleStepSelector)
		sSampleStepSelector(i) = SampleStepSelectorArray(i)
	Next

	Begin Dialog UserDialog 560,357,"Reverberation Chamber Near Field Source",.DialogFunc ' %GRID:10,7,1,1
		'
		OKButton 80,329,90,21
		PushButton 400,329,90,21,"&Help",.HelpB
		GroupBox 10,14,280,175,"Working volume",.GroupBox1
		GroupBox 310,14,240,175,"Frequency settings",.GroupBox2
		GroupBox 10,196,240,126,"General Settings",.GroupBox3
		TextBox 80,56,80,21,.WVx
		TextBox 190,56,80,21,.WVx2
		TextBox 80,84,80,21,.WVy
		TextBox 190,84,80,21,.WVy2
		TextBox 80,112,80,21,.WVz
		TextBox 190,112,80,21,.WVz2
		TextBox 190,147,80,21,.Spacer
		TextBox 460,28,80,21,.FreqMin
		TextBox 460,56,80,21,.FreqMax
		TextBox 460,84,80,21,.FreqSamplesStep
		CheckBox 330,112,160,14,"Use log sampling",.UseLogSamples
		TextBox 330,154,150,21,.FreqFilePath
		PushButton 480,154,60,21,"Browse",.FreqFilePathB
		TextBox 150,217,80,21,.MCIters
		TextBox 150,245,80,21,.Waves
		TextBox 150,273,80,21,.Seed
		CheckBox 30,301,170,14,"Write near field source",.WriteFieldSource
		Text 30,63,20,14,"x:",.Text1
		Text 100,35,50,14,"min",.Text6
		Text 20,147,150,28,"spacer (multiple of lambda_min / 15)",.Text12
		Text 210,35,50,14,"max",.Text11
		Text 30,91,20,14,"y:",.Text2
		Text 30,119,20,14,"z:",.Text3
		Text 330,35,70,14,"freq. min.:",.Text4
		Text 330,63,70,14,"freq. max.:",.Text5
		Text 330,133,120,14,"External freq. file",.Text7
		Text 30,224,110,14,"Monte Carlo iters:",.Text8
		Text 30,252,110,14,"Num waves:",.Text9
		Text 30,280,110,14,"Seed:",.Text10
		CancelButton 240,329,90,21
		DropListBox 330,84,120,70,sSampleStepSelector(),.SampleStepSelectorDLB
	End Dialog
		'

	'Dialog
	Dim dlg As UserDialog

	' initialization
	dlg.WVx = CStr(0)
	dlg.WVy = CStr(0)
	dlg.WVz = CStr(0)
	dlg.WVx2 = CStr(1000)
	dlg.WVy2 = CStr(1000)
	dlg.WVz2 = CStr(1000)
	dlg.Spacer = CStr(10)
	dlg.FreqMin = CStr(.1)
	dlg.FreqMax = CStr(1)
	dlg.FreqSamplesStep = CStr(10)
	dlg.UseLogSamples = False
	dlg.FreqFilePath = ""
	dlg.MCIters = CStr(10)
	dlg.Waves = CStr(100)
	dlg.Seed = CStr(-1)
	dlg.WriteFieldSource = True

	' open dialog
	Dim iDialogResponse As Integer
	iDialogResponse = Dialog(dlg)

    If (iDialogResponse >= 0) Then Exit All

	If iDialogResponse = -1 Then	'OK button

		Dim nSelectedSampleStep As Integer
		nSelectedSampleStep = dlg.SampleStepSelectorDLB

		Dim LinFreqSamples As Integer
		Dim LinFreqStep As Double

		If(nSelectedSampleStep = 0) Then
			LinFreqSamples = Cint(dlg.FreqSamplesStep)
			LinFreqStep = 0
		ElseIf (nSelectedSampleStep = 1) Then
			LinFreqSamples = 0
			LinFreqStep = CDbl(dlg.FreqSamplesStep)
		End If

		Dim use_lin_freq_sample_step As Boolean
		use_lin_freq_sample_step = (nSelectedSampleStep = 0)

		With FieldSource
			.CreateReverberationChamberFieldImport(dlg.FreqMin, dlg.FreqMax, _
												   dlg.WVx, dlg.WVy, dlg.WVz, dlg.WVx2, dlg.WVy2, dlg.WVz2, _
												   dlg.Spacer, _
												   use_lin_freq_sample_step, CStr(LinFreqSamples), CStr(LinFreqStep), dlg.UseLogSamples, _
												   dlg.FreqFilePath, _
												   dlg.MCIters, dlg.Seed, dlg.Waves, dlg.WriteFieldSource )
		End With
	End If

End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
    Debug.Print "Action=";Action%
    Select Case Action%
	Case 1 ' Initialization
		DlgText("SampleStepSelectorDLB",SampleStepSelectorArray(0))

    Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "HelpB"
			StartHelp HelpFileName
			DialogFunc = True
		Case "FreqFilePathB"
			Dim FilePath As String
			FilePath = GetFilePath("*.dat;*txt","Data Files|*.dat|Text Files|*.txt|All Files|*.*",GetProjectPath("Root"),"Select external frequency file To load",0)
			If (FilePath <> "") Then
				DlgText("FreqFilePath", FilePath)
			End If
			DialogFunc = True
		Case "OK"

			Dim WVx, WVy, WVz As Double
			Dim WVx2, WVy2, WVz2 As Double
			Dim Spacer As Double
			Dim FreqMin, FreqMax As Double
			Dim LinFreqSamples As Integer
			Dim LinFreqStep As Double
			Dim UseLogSamples As Boolean
			Dim FreqFilePath As String
			Dim MCIters, Waves, Seed As Integer
			Dim WriteFieldSource As Boolean

			WVx = CDbl(Evaluate(DlgText("WVx")))
			WVy = CDbl(Evaluate(DlgText("WVy")))
			WVz = CDbl(Evaluate(DlgText("WVz")))
			WVx2 = CDbl(Evaluate(DlgText("WVx2")))
			WVy2 = CDbl(Evaluate(DlgText("WVy2")))
			WVz2 = CDbl(Evaluate(DlgText("WVz2")))

			Spacer = CDbl(Evaluate(DlgText("Spacer")))

			FreqMin = CDbl(Evaluate(DlgText("FreqMin")))
			FreqMax = CDbl(Evaluate(DlgText("FreqMax")))

			Dim sSelectedSampleStep As String
			Dim nSelectedSampleStep As Integer
			sSelectedSampleStep = DlgText("SampleStepSelectorDLB")

			If(sSelectedSampleStep = SampleStepSelectorArray(0)) Then
				nSelectedSampleStep = 0
				LinFreqSamples = Cint(Evaluate(DlgText("FreqSamplesStep")))
				LinFreqStep = 0
			ElseIf (sSelectedSampleStep = SampleStepSelectorArray(1)) Then
				nSelectedSampleStep = 1
				LinFreqSamples = 0
				LinFreqStep = CDbl(Evaluate(DlgText("FreqSamplesStep")))
			End If

			UseLogSamples = CBool(DlgValue("UseLogSamples"))
			FreqFilePath = DlgText("FreqFilePath")

			MCIters = Cint(Evaluate(DlgText("MCIters")))
			Waves = Cint(Evaluate(DlgText("Waves")))
			Seed = Cint(Evaluate(DlgText("Seed")))
			WriteFieldSource = CBool(DlgValue("WriteFieldSource"))

			Dim bOk As Boolean
			bOk	= True

			' check if everything is ok
			If WVx2 <= WVx Then
				bOk	= False
				MsgBox "Please provide a positive x working volume size", vbCritical & vbOkOnly, "Error"
			End If

			If WVy2 <= WVy Then
				bOk	= False
				MsgBox "Please provide a positive y working volume size", vbCritical & vbOkOnly, "Error"
			End If

			If WVz2 <= WVz Then
				bOk	= False
				MsgBox "Please provide a positive z working volume size", vbCritical & vbOkOnly, "Error"
			End If

			If Spacer <= 0 Then
				bOk	= False
				MsgBox "Please provide a positive number of spacer layer between working volume and field source", vbCritical & vbOkOnly, "Error"
			End If

			If FreqMin <= 0 Then
				bOk	= False
				MsgBox "Please provide a positive minimum freqency", vbCritical & vbOkOnly, "Error"
			End If

			If FreqMax <= FreqMin Then
				bOk	= False
				MsgBox "Please provide a maximum frequency greater than the minimum freqency", vbCritical & vbOkOnly, "Error"
			End If

			If(nSelectedSampleStep = 0) Then
				If LinFreqSamples < 0 Then
					bOk	= False
					MsgBox "Please provide a non negative number of linear frequency samples", vbCritical & vbOkOnly, "Error"
				End If
			ElseIf (nSelectedSampleStep = 1) Then
				If LinFreqStep <= 0 Then
					bOk	= False
					MsgBox "Please provide a positive linear frequency step", vbCritical & vbOkOnly, "Error"
				End If
			End If

			If  (nSelectedSampleStep = 0) And (LinFreqSamples = 0) And (UseLogSamples = False) And (Len(FreqFilePath)=0) Then
				bOk	= False
				MsgBox "Please define a positive number of frequency samples (among custom, linear and logarithmic distributed)", vbCritical & vbOkOnly, "Error"
			End If

			If MCIters <= 0 Then
				bOk	= False
				MsgBox "Please provide a positive number of monte Carlo iterations", vbCritical & vbOkOnly, "Error"
			End If

			If Waves <= 0 Then
				bOk	= False
				MsgBox "Please provide a positive number of waves", vbCritical & vbOkOnly, "Error"
			End If

			If bOk = False Then
				DialogFunc = True
				Exit Function
			End If
	    End Select
    End Select
End Function
