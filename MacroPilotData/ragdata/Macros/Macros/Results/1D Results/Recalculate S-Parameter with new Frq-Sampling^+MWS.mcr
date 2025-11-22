' *Results / Recalculate S-Parameter with new Frq-Sampling
' !!! Do not change the line above !!!

' ==========================================================================
' This VBA code demonstrates how to recalculate the sparameters based on the
' time signals of an existing result data set.
' The recalculation is based on a DFT and therefore allows to specify a
' number of frequency samples which is independent from the setting used
' for the original simulation run.

' NOTE: This macro only works for frequency independent port impedances.
' The results are inaccurate when the "full deembedding" option was used
' for inhomogeneous ports. Furthermore, the results may be completely
' wrong when non TEM like modes (e.g. TE / TM) are present at the ports and
' the ports have a different cutoff frequency.

' Copyright 2005-2023 Dassault Systemes Deutschland GmbH
' ==========================================================================
' History of Changes
' --------------------------------------------------------------------------
' 26-Jul-2007 imu,ube: works now also for discrete ports
' 21-Oct-2005 imu: Included into Online Help, previously no Help was present !
' 13-Oct-2005 ube: Included into Online Help
' 11-Oct-2005 ube: Removed limitation of TEM modes only
' ==========================================================================

Const HelpFileName = "common_preloadedmacro_results_recalculate_s-parameter_with_new_frequency_sampling"


Sub Main

' Specify the new number of frequency samples

	Begin Dialog UserDialog 310,91, "Recalculate S-Parameters", .DialogFunc ' %GRID:10,7,1,1
		Text 30,14,130,14,"Number of samples:",.Text1
		TextBox 170,7,80,21,.NSamples
		Text 30,35,160,14,"(only for TEM-like modes)",.Text2
		OKButton 10,63,90,21
		CancelButton 110,63,90,21
		PushButton 210,63,90,21,"Help",.Help
	End Dialog
	Dim dlg As UserDialog
	dlg.NSamples = "1001"
	If (Dialog(dlg) = 0) Then Exit All

	Dim nSamples As Long
	nSamples = CLng(dlg.NSamples)

' The following loop determines which ports have been used in this model
' and stores them in any array nPortNumber for easier access. The total
' number of ports is stored in nPorts

	Port.StartPortNumberIteration

	Dim nPorts As Long
	Dim nPortNumber() As Long

	nPorts = 0

	While (1)
		Dim nPort As Long

		nPort = Port.GetNextPortNumber()
		If nPort = -1 Then Exit While

		nPorts = nPorts+1
		ReDim Preserve nPortNumber(nPorts)

		nPortNumber(nPorts-1) = nPort
	Wend

' Now determine the frequency range used for the simulation

	Dim dFmin As Double, dFmax As Double
	dFmin = Solver.GetFmin()
	dFmax = Solver.GetFmax()

' Once the actual port numbers are stored in an array, this array may now be
' used to iterate over all ports as input

	Dim nInPort As Long, nOutPort As Long, nInMode As Long, nOutMode As Long
	Dim nIn As Long, nOut As Long, noutportmodes As Long, ninportmodes As Long

	For nIn = 0 To nPorts-1

		' For each of the input ports, loop over all modes in the port

		nInPort = nPortNumber(nIn)

		ninportmodes = Port.GetNumberOfModes(nInPort)
		If ninportmodes = 0 Then ninportmodes = 1   ' discrete port

		For nInMode = 1 To ninportmodes

			' Now use the port number array again in order to loop over all ports as output

			For nOut = 0 To nPorts-1

				' For each of the output ports, loop over all modes in the port

				nOutPort = nPortNumber(nOut)

				noutportmodes = Port.GetNumberOfModes(nOutPort)

				If noutportmodes = 0 Then noutportmodes=1 ' discrete port

				For nOutMode = 1 To noutportmodes

				' Make sure that the mode is either a TEM or QTEM mode (otherwise frequency dependent impedance factor
				' needs to be considered, which isn't done here)

				If Port.GetType(nOutPort) = "Waveguide" Then

					If Port.GetModeType(nOutPort, nOutMode) <> "TEM" And Port.GetModeType(nOutPort, nOutMode) <> "QTEM" Then
						ReportError "Invalid mode type for resampling"
					End If

				End If

				' Determine the strings for the input and output naming convention

					Dim sInput As String, sOutput As String
					sInput = CStr(nInPort) + "(" + CStr(nInMode) + ")"
					sOutput = CStr(nOutPort) + "(" + CStr(nOutMode) + ")" + CStr(nInPort) + "(" + CStr(nInMode) + ")"

					Dim Sinp As Object, Soutp As Object

					' In the following, the corresponding time signals are being read into Result1D objects.
					' This operation may fail, when the corresponding ports have not been excited in the
					' simulation. Therefore, an error handler is set in order to catch this type of error.

					On Error GoTo Failed

					' Now read the input and output time signals

					Set Sinp  = Result1D("i" + sInput)
					Set Soutp = Result1D("o" + sOutput)
					
					Dim SinpC as Object
					Dim SoutpC as Object
					Set SinpC = Result1DComplex("")
					Set SoutpC = Result1DComplex("")

					SinpC.Initialize(Sinp.GetN)
					SoutpC.Initialize(Soutp.GetN)

					Dim ii As Integer
					For ii=0 To Sinp.GetN-1
						SinpC.SetX(ii, Sinp.GetX(ii))
						SinpC.SetYRe(ii, Sinp.GetY(ii))
					Next ii
					
					For ii=0 To Soutp.GetN-1
						SoutpC.SetX(ii, Soutp.GetX(ii))
						SoutpC.SetYRe(ii, Soutp.GetY(ii))
					Next ii
					
					
					Dim AmInp As Object, PhInp As Object, AmOutp As Object, PhOutp As Object

					SetIntegrationMethod "trapezoidal"
					CalculateFourierComplex(SinpC,  "time", SInpC, "frequency", "-1", "1.0", dFmin, dFmax,nSamples)
					CalculateFourierComplex(SoutpC, "time", SOutpC, "frequency", "-1", "1.0", dFmin, dFmax,nSamples)

					' Divide the output spectrum by the input spectrum in order to get the sparameters
					SoutpC.ComponentDiv(SinpC)
					SoutpC.SetLogarithmicFactor(20.0)
					SoutpC.SetXLabelAndUnit( "Frequency" ,  Units.GetUnit("Frequency"))
					SoutpC.YLabel ""
					SoutpC.Title "S-Parameters (Resampled)" 

					SoutpC.Save("^Sc" + sOutput + "_resampled.sig")
					
					SoutpC.AddToTree("1D Results\Resampled S-Parameters\S" + sOutput)

				Failed:

				Next nOutMode
			Next nOut

		Next nInMode
	Next nIn

	SelectTreeItem "1D Results\Resampled S-Parameters\|S| linear"

End Sub
'--------------------------------------------------------------------------------------------
Function DialogFunc%(Item As String, Action As Integer, Value As Integer)
	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "Help"
			StartHelp HelpFileName
			DialogFunc = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
