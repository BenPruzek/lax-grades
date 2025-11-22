
' This macro imports S parameters from Touchstone files and adds them as 1D results to the navigation tree

' ---------------------------------------------------------------------------------------------------------------------------------
' Version history:
'
' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 28-Dec-2020 fsr: Removed obsolete "Design\" prefix for plots added to DS tree
' 19-Aug-2019 fsr: Added reference impedance plot; minor performance improvements
' 30-Sep-2016 fsr: Improved error feedback (slightly), blank lines in file are now supported
' 04-Jan-2016 fsr: For compatibility with Czech (and possibly other) locales, switch to US locale for execution
' 24-Sep-2013 fsr: Special case: for s2p files, S21 comes before S12 in the file; fixed.
' 11-Sep-2012 chn,fsr: last data point was copy of previous point, fixed
' 19-Sep-2011 fsr: Some GUI improvements (Abort button, progress feedback, ...)
' 02-Sep-2011 fsr: Initial version with GUI
' ---------------------------------------------------------------------------------------------------------------------------------

'#include "complex.lib"

Const bDebug = False 'Debug flag
Dim GlobalDataFileName As String

Sub Main

	GlobalDataFileName = GetFilePath("*.s*p", "TOUCHSTONE files|*.s*p", , "Please select Touchstone file to be imported", 0)
	If GlobalDataFileName = "" Then Exit All

	Begin Dialog UserDialog 400,98,"Importing TOUCHSTONE file",.DialogFunc ' %GRID:10,7,1,1
		Text 20,14,370,14,"Importing '"+Split(GlobalDataFileName,"\")(UBound(Split(GlobalDataFileName,"\")))+"'",.FileNameT
		Text 20,42,370,14,"",.OutputT
		CheckBox 20,70,60,14,"Abort",.AbortCB
		OKButton 290,70,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable("OK", False)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		If (ImportTouchstoneFile(GlobalDataFileName)) = 0 Then
			'ReportInformationToWindow("Finished Touchstone import.")
			DlgText("OutputT", "Done!")
		Else
			DlgText("OutputT", "Error during import")
		End If
		DlgEnable("OK", True)
		DlgEnable("AbortCB", False)
	Case 6 ' Function key
	End Select
End Function

Function ImportTouchstoneFile(DataFileName As String) As Integer

	Dim DataFile As Long
	Dim SParameterMatrix() As Complex, SParameterMatrixMagDeg() As Double
	Dim tmpLine As String
	Dim i As Long, j As Long, k As Long, lastBreak As Long, entryNumber As Long
	Dim n As Long
	Dim SParaMag As Double, SParaDegree As Double, SParaFreq As Double
	Dim SParaPlot() As Object, ImpedancePlot As Object
	Dim DataFormat As String, dFreqScale As Double, sFreqText As String, dReferenceImpedance As Double
	Dim nTempIndex As Long
	Dim iCurrentLocale As Integer
	Dim sReferenceImpedanceTreePath As String

	ImportTouchstoneFile = -1 ' 0 = All OK, -1 = unfinished/error

	' Store current locale and switch to US
	iCurrentLocale = GetLocale
	SetLocale(&H409)
	On Error GoTo ExitImport

	n = CLng(Mid(Split(DataFileName,".")(UBound(Split(DataFileName,"."))),2,Len(Split(DataFileName,".")(UBound(Split(DataFileName,"."))))-2)) ' size of matrix
	ReDim SParameterMatrix(n-1,n-1)
	ReDim SParameterMatrixMagDeg(n-1,2*(n-1)+1)

	ReDim SParaPlot(n-1,n-1)
	For i = 1 To n
		For j = 1 To n
			If Left(GetApplicationName, 2) = "DS" Then
				Set SParaPlot(i-1,j-1) = DS.Result1DComplex("")
			Else
				Set SParaPlot(i-1,j-1) = Result1DComplex("")
			End If
		Next
	Next

	'ReportInformationToWindow("Starting Touchstone import. This might take some time for large files.")

	DataFile = FreeFile
	Open DataFileName For Input As #DataFile
		' Skip comment files if any
		Do
			Do ' skip any empty lines
				Line Input #DataFile, tmpLine
			Loop While tmpLine = ""
			If Mid(tmpLine,1,1)="#" Then
				' Determine data format
				If InStr(1,UCase(tmpLine),"MA")>0 Then
					DataFormat = "MA"
				ElseIf InStr(1,UCase(tmpLine),"RI")>0 Then
					DataFormat = "RI"
				ElseIf InStr(1,UCase(tmpLine),"DB")>0 Then
					DataFormat = "DB"
				Else
					ReportError("Cannot determine data format (MA/RI/DB).")
				End If
				' Determine frequency format
				If InStr(1, UCase(tmpLine), "PHZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI/1e15
					sFreqText = "PHz"
				ElseIf InStr(1, UCase(tmpLine), "THZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI/1e12
					sFreqText = "THz"
				ElseIf InStr(1, UCase(tmpLine), "GHZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI/1e9
					sFreqText = "GHz"
				ElseIf InStr(1, UCase(tmpLine), "MHZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI/1e6
					sFreqText = "MHz"
				ElseIf InStr(1, UCase(tmpLine), "KHZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI/1e3
					sFreqText = "kHz"
				ElseIf InStr(1, UCase(tmpLine), "HZ")>0 Then
					FreqScale = Units.GetFrequencyUnitToSI
					sFreqText = "Hz"
				Else
					FreqScale = 1
					sFreqText = ""
				End If
				' Determine reference impedance
				nTempIndex = InStr(1, UCase(tmpLine), "R ")
				If nTempIndex > 0 Then
					dReferenceImpedance = Evaluate(Mid(tmpLine, nTempIndex + 2, Max(InStr(nTempIndex + 2, tmpLine, " "), Len(tmpLine)) - nTempIndex))
				Else
					dReferenceImpedance = -1
				End If
			End If
		Loop Until (Mid(tmpLine,1,1)<>"!" And Mid(tmpLine,1,1)<>"#")
		' Now parse the rest of the file
		k = 1 ' char position in tmpLine
		While Not EOF(DataFile)
			' Wind forward until first number is found
			lastBreak = k
			While Not IsNumeric(Mid(tmpLine, lastBreak, k-lastBreak+1))
				k = k + 1
				If k > Len(tmpLine) Then
					If EOF(DataFile) Then
						GoTo EOFReached
					Else
						' Get Next line, reset k
						k = 1
						lastBreak = 1
						Do ' skip any empty lines
							Line Input #DataFile, tmpLine
						Loop While tmpLine = ""
					End If
				End If
			Wend
			' Beginning of number has been found, forward until end of numerical entry
			lastBreak = k
			While (IsNumeric(Mid(tmpLine, lastBreak, k-lastBreak+1)+"0") And k <= Len(tmpLine)) ' Add a 0 to detect cases like "3e..."
				k = k+1
			Wend
			' When number is not numeric anymore, save frequency
			SParaFreq = Evaluate(Mid(tmpLine, lastBreak, k-lastBreak))
			' Step to next line if number was terminated by end of line
			If k > Len(tmpLine) Then
				' Get Next line, reset k
				k = 1
				Do ' skip any empty lines
					Line Input #DataFile, tmpLine
				Loop While tmpLine = ""
			End If
			entryNumber = 0
			i = 0 ' i represents the number of rows for the matrix
			' Now search the rest of the entry
			While (i < n)
				' Wind forward until first number is found
				lastBreak = k
				While (Not IsNumeric(Mid(tmpLine, lastBreak, k-lastBreak+1)) And Mid(tmpLine,k,1)<>"-")
					k = k + 1
					If k > Len(tmpLine) Then
						' Get Next line, reset k
						k = 1
						lastBreak = 1
						Do ' skip any empty lines
							Line Input #DataFile, tmpLine
						Loop While tmpLine = ""
					End If
				Wend
				' Now forward until number is not numeric anymore -> entry found
				lastBreak = k
				While (IsNumeric(Mid(tmpLine, lastBreak, k-lastBreak+1)+"0") And k <= Len(tmpLine)) ' Add a 0 to detect cases like "3e..."
					k = k+1
				Wend
				' When number is not numeric anymore, save entry
				entryNumber = entryNumber + 1
				SParameterMatrixMagDeg(i,entryNumber-1) = Evaluate(Mid(tmpLine, lastBreak, k-lastBreak))
				' Step to next line if number was terminated by end of line and if i<n-1
				If ((k > Len(tmpLine)) And (i<n-1)) Then
					' Get Next line, reset k
					k = 1
					Do ' skip any empty lines
						Line Input #DataFile, tmpLine
					Loop While tmpLine = ""
				End If
				If (entryNumber = 2*n) Then ' next row in the matrix
					i = i+1
					entryNumber = 0
				End If
			Wend

			If DataFormat = "MA" Then
				For i = 0 To n-1
					For j = 0 To n-1
						SParameterMatrix(i,j) = AssignComplex(SParameterMatrixMagDeg(i,2*j)*CosD(SParameterMatrixMagDeg(i,2*j+1)), SParameterMatrixMagDeg(i,2*j)*SinD(SParameterMatrixMagDeg(i,2*j+1)))
						SParaPlot(i,j).AppendXYDouble(SParaFreq/FreqScale, SParameterMatrix(i,j).re, SParameterMatrix(i,j).im)
					Next
				Next
			ElseIf DataFormat = "RI" Then
				For i = 0 To n-1
					For j = 0 To n-1
						SParameterMatrix(i,j) = AssignComplex(SParameterMatrixMagDeg(i,2*j), SParameterMatrixMagDeg(i,2*j+1))
						SParaPlot(i,j).AppendXYDouble(SParaFreq/FreqScale, SParameterMatrix(i,j).re, SParameterMatrix(i,j).im)
					Next
				Next
			ElseIf DataFormat = "DB" Then
				For i = 0 To n-1
					For j = 0 To n-1
						SParameterMatrix(i,j) = AssignComplex(10^(SParameterMatrixMagDeg(i,2*j)/20)*CosD(SParameterMatrixMagDeg(i,2*j+1)), 10^(SParameterMatrixMagDeg(i,2*j)/20)*SinD(SParameterMatrixMagDeg(i,2*j+1)))
						SParaPlot(i,j).AppendXYDouble(SParaFreq/FreqScale, SParameterMatrix(i,j).re, SParameterMatrix(i,j).im)
					Next
				Next
			End If

			If (SParaPlot(0,0).GetN() Mod 10 = 1) Then DlgText("OutputT", "Frequency: "+CStr(SParaFreq) + " " + sFreqText)

			'ReportInformationToWindow(CStr(SParaFreq/FreqScale)+": "+CStr(SParameterMatrix(n-1,n-1).re)+"+j*"+CStr(SParameterMatrix(n-1,n-1).im))

			If CBool(DlgValue("AbortCB")) = True Then
				DlgValue("AbortCB", IIf(MsgBox("Plot frequency samples read so far?", vbYesNo, "Abort")=vbYes, False, True))
				GoTo EOFReached
			End If
		Wend

	EOFReached:
		Close DataFile

		If Left(GetApplicationName, 2) = "DS" Then
			ResultTree.EnableTreeUpdate(False)
			If (dReferenceImpedance > -1) Then
				If dReferenceImpedance < 1e-9 Then ReportInformation("TOUCHSTONE Import: Reference impedance is zero, S-Parameters were not renormalized.")
				Set ImpedancePlot = SParaPlot(0,0).Copy()
				For i = 0 To ImpedancePlot.GetN()-1
					ImpedancePlot.SetYRe(i, dReferenceImpedance)
					ImpedancePlot.SetYIm(i, 0)
				Next
				ImpedancePlot.Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_Z")
				sReferenceImpedancePath = "Results\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\Reference Impedance"
				ImpedancePlot.AddToTree(sReferenceImpedancePath)
			End If
			For i = 1 To n
				If (sReferenceImpedancePath <> "") Then SParaPlot(i-1,i-1).SetReferenceImpedanceLink(sReferenceImpedancePath)
				For j = 1 To n
					If (SParaPlot(i-1,j-1).GetN() < 1) Then MsgBox("TS import: No data found for S" & CStr(i) & "," & CStr(j) & ". Please check file format.", "Error")
					If CBool(DlgValue("AbortCB")) = True Then
						ResultTree.EnableTreeUpdate(True)
						If Not (i=1 And j =1) Then DS.SelectTreeItem("Results\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_"))
						Exit All
					End If
					DlgText("OutputT", "Data read, adding plots to tree... " + Cstr((i-1)*n+j) + "/" + Cstr(n^2))
					If (n > 2) Then
						SParaPlot(i-1,j-1).Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_S"+CStr(i)+CStr(j))
						SParaPlot(i-1,j-1).AddToTree("Results\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\S"+CStr(i)+CStr(j))
					Else ' special case for n = 2: S21 comes first, then S12 (see Touchstone standard)
						SParaPlot(j-1,i-1).Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_S"+CStr(i)+CStr(j))
						SParaPlot(j-1,i-1).AddToTree("Results\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\S"+CStr(i)+CStr(j))
					End If
				Next
			Next
			ResultTree.EnableTreeUpdate(True)
			DS.SelectTreeItem("Results\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_"))
		Else
			ResultTree.EnableTreeUpdate(False)
			If (dReferenceImpedance > -1) Then
				If dReferenceImpedance < 1e-9 Then ReportInformation("TOUCHSTONE Import: Reference impedance is zero, S-Parameters were not renormalized.")
				Set ImpedancePlot = SParaPlot(0,0).Copy()
				For i = 0 To ImpedancePlot.GetN()-1
					ImpedancePlot.SetYRe(i, dReferenceImpedance)
					ImpedancePlot.SetYIm(i, 0)
				Next
				ImpedancePlot.Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_S")
				sReferenceImpedancePath = "1D Results\TOUCHSTONE Imports\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\Reference Impedance"
				ImpedancePlot.AddToTree(sReferenceImpedancePath)
			End If
			For i = 1 To n
				If (sReferenceImpedancePath <> "") Then SParaPlot(i-1,i-1).SetReferenceImpedanceLink(sReferenceImpedancePath)
				For j = 1 To n
					If (SParaPlot(i-1,j-1).GetN() < 1) Then MsgBox("TS import: No data found for S" & CStr(i) & "," & CStr(j) & ". Please check file format.", "Error")
					If CBool(DlgValue("AbortCB")) = True Then
						Resulttree.EnableTreeUpdate(True)
						If Not (i=1 And j =1) Then SelectTreeItem("1D Results\TOUCHSTONE Imports\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_"))
						Exit All
					End If
					DlgText("OutputT", "Data read, adding plots to tree... " + Cstr((i-1)*n+j) + "/" + Cstr(n^2))
					If (n > 2) Then
						SParaPlot(i-1,j-1).Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_S"+CStr(i)+CStr(j))
						SParaPlot(i-1,j-1).AddToTree("1D Results\TOUCHSTONE Imports\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\S"+CStr(i)+CStr(j))
					Else ' special case for n = 2: S21 comes first, then S12 (see Touchstone standard)
						SParaPlot(j-1,i-1).Save(GetProjectPath("Result")+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"_S"+CStr(i)+CStr(j))
						SParaPlot(j-1,i-1).AddToTree("1D Results\TOUCHSTONE Imports\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_")+"\S"+CStr(i)+CStr(j))
					End If
				Next
			Next
			Resulttree.EnableTreeUpdate(True)
			SelectTreeItem("1D Results\TOUCHSTONE Imports\"+Replace(Split(DataFileName,"\")(UBound(Split(DataFileName,"\"))),".","_"))
		End If

	ImportTouchstoneFile = 0 ' All OK

	ExitImport:
		SetLocale(iCurrentLocale)

End Function
