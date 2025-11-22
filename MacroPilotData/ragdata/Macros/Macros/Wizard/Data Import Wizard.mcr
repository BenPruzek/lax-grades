'#Language "WWB-COM"

' This wizard allows to import 1D/1DC data stored in ASCII files

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ----------------
' 06-Apr-2022 fsr: Fixed scaling of 2nd column; allowed dB magnitude for polar import
' 03-Aug-2020 fsr: Improved error message when auto format detection fails; switch to US locale for format detection to avoid locale issues
' 05-Aug-2019 fsr: Trim leading and trailing spaces from input lines before parsing
' 26-Jul-2019 fsr: Performance improvements
' 13-Mar-2018 fsr: Fixed potential endless loop if input line does not contain any numerical data
' 25-Oct-2017 fsr: Fixed MagDB import for Schematic; allowed user to specify dB factor; axis scaling now applied to multiple imports with periodic shift
' 22-Jul-2016 fsr: Set DeleteAt("never") for all cases (3D, DS, R1D and R1DC)
' 28-Mar-2016 fsr: Disabled "Next" button if no consistent numerical data are found
' 16-Jun-2015 fsr: Replaced obsolete polar plot with new method; replaced VBA bubble sort by Result1DComplex.SortByX
' 17-Nov-2014 fsr: Dialog would update display continuously in some cases, fixed;
' 11-Nov-2014 fsr: Fixed a problem with y scaling buttons for 1D imports without imaginary part
' 11-Aug-2014 tgl,fsr: little fix
' 18-Apr-2014 fsr: Fixed import for multiple columns at the same time
' 16-Jun-2013 fsr: Added functionality for PS/EMS/CS; now using file name as preset for import name; plots can now get sorted w.r.t. x values
' 12-Jun-2013 tgl,fsr: added .sig extension to fix save/load problems of files
' 12-Dec-2011 fsr: changed a problem with AM header recognition for plots from chart trace tool; curve properties (Re/Im/Mag...) can now be changed after import
' 26-Sep-2011 fsr: forward to Touchstone import macro if file extension is s*p
' 11-May-2011 fsr: more line break issues
' 04-May-2011 fsr: fixed a problem with recognizing vbLf as line break
' 21-Apr-2011 fsr: import as 1d polar plot now possible
' 15-Apr-2011 fsr: added support for AM 2.5+ header format (incl. auto detect), CST ASCII Export format; performance improvements
' 17-Feb-2011 fsr: Added clipboard option, 1D results are now persistent in MWS
' 16-Feb-2011 fsr: Added new data format options: "fixed length" and "auto detect". Some changes to the GUI. Mag in dB now allowed.
' 10-Feb-2011 fsr: Initial version
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Dim DataFileName As String, DataContents As String
Dim DataString(4) As String, ShowColumnsArray() As String, SelectedShowColumns As Integer
Dim HeaderAndComments As String
Dim DataArray() As String, XDataColumn() As Double, Y1DataColumn() As Double, Y2DataColumn() As Double
Dim SepString As String, FixedLength As Integer, ComChars As String, NHeaderLines As Long, ColRowFormat As String, FileFormat As Long
Dim DataGroupList() As String
Dim ColRowList() As Long, ColRowLabelList() As String
Dim DataPlotType() As String
Dim LastPlotCreated As String

Const ResultTypeArray = Array("1D", "1DC")
Const PresetsArray = Array("None (customized)", "CST ASCII Export", "Antenna Magus (up to v2.5)", "Antenna Magus (v2.5 and up)")
Const SepArray = Array("Tab", "Comma", "Space", "Colon", "Semicolon", "<Custom>")
Const SepArrayMAP = Array(vbTab, ",", " ", ":", ";") ' Maps the displayed char names to the corresponding characters

Sub Main

	Begin Dialog UserDialog 1010,574,"CST data import wizard",.ImportWizardDialogFunction ' %GRID:10,7,1,1
		GroupBox 20,7,970,147,"General",.GroupBox1
		Text 30,35,110,14,"Data file location:",.Text4
		TextBox 150,28,430,21,.DataFileLocationT
		PushButton 590,28,90,21,"Change File",.ChangeFilePB
		PushButton 690,28,90,21,"Clipboard",.ReadClipboardPB
		Text 490,70,120,14,"Data arranged in:",.Text1
		OptionGroup .ColRowOG
			OptionButton 620,70,80,14,"Columns",.ColOB
			OptionButton 710,70,60,14,"Rows",.RowOB
		DropListBox 300,91,90,192,SepArray(),.SeparatorDLB,1
		Text 490,126,150,14,"Number of lines to skip:",.Text5
		TextBox 650,119,50,21,.NHeaderLinesT
		PushButton 700,112,20,14,"+",.NHeaderUPPB
		PushButton 700,133,20,14,"-",.NHeaderDOWNPB
		Text 490,98,140,14,"Comment characters:",.Text6
		TextBox 650,91,50,21,.CommentCharsT

		GroupBox 20,154,970,77,"Discarded header and comment lines",.GroupBox2
		TextBox 30,175,950,49,.headerDisplay,2

		GroupBox 20,231,970,308,"Data, reorganized into columns",.GroupBox3
		TextBox 30,252,190,280,.dataDisplay1,2
		TextBox 220,252,190,280,.dataDisplay2,2
		TextBox 410,252,190,280,.dataDisplay3,2
		TextBox 600,252,190,280,.dataDisplay4,2
		TextBox 790,252,190,280,.dataDisplay5,2
		PushButton 800,546,90,21,"Next >>",.NextPB
		PushButton 900,546,90,21,"Exit",.ClosePB
		Text 30,70,80,14,"Data format:",.Text3
		OptionGroup .FileFormatOG
			OptionButton 120,70,100,14,"Auto detect",.AutoOB
			OptionButton 120,98,160,14,"Character separated:",.SeparatorFormat
			OptionButton 120,126,170,14,"Fixed data item length:",.FixedLength
		TextBox 300,119,90,21,.FixedLengthT
		PushButton 390,112,20,14,"+",.FixedLengthUPPB
		PushButton 390,133,20,14,"-",.FixedLengthDOWNPB
		Text 30,546,100,14,"Show columns:",.Text2
		DropListBox 130,539,90,192,ShowColumnsArray(),.ShowColumnsDLB

	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = -1) Then

	Else

	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function ImportWizardDialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		InitDialog()
	Case 2 ' Value changing or button pressed
		Rem ImportWizardDialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
			Case "ChangeFilePB"
				ImportWizardDialogFunction = True
				InitDialog()
			Case "ReadClipboardPB"
				ImportWizardDialogFunction = True
				If(MsgBox(Clipboard, vbOkCancel, "Confirm import") = vbOK) Then
					DlgText("DataFileLocationT","Data imported from clipboard.")
					DataContents = Clipboard()
					DlgText("NHeaderLinesT",CStr(AutoDetectHeaderLines(DataContents, DlgValue("FileFormatOG"), IIf(DlgValue("FileFormatOG") = 1, DlgText("SeparatorDLB"), DlgText("FixedLengthT")))))
					UpdateDialog()
				End If
			Case "NextPB"
				ImportWizardDialogFunction = True
				OpenImportDialog()
			Case "CancelPB"
				ImportWizardDialogFunction = False
				Exit All
			Case "ShowColumnsDLB"
				SelectedShowColumns = DlgValue("ShowColumnsDLB")
				ImportWizardDialogFunction = True
				UpdateDialog()
			Case "ColRowOG", "FileFormatOG"
				ImportWizardDialogFunction = True
				UpdateDialog()
			Case "NHeaderUPPB"
				ImportWizardDialogFunction = True
				DlgText("NHeaderLinesT", DlgText("NHeaderLinesT")+1)
				UpdateDialog()
			Case "NHeaderDOWNPB"
				ImportWizardDialogFunction = True
				If (DlgText("NHeaderLinesT")>"0") Then DlgText("NHeaderLinesT", DlgText("NHeaderLinesT")-1)
				UpdateDialog()
			Case "FixedLengthUPPB"
				ImportWizardDialogFunction = True
				DlgText("FixedLengthT", DlgText("FixedLengthT")+1)
				UpdateDialog()
			Case "FixedLengthDOWNPB"
				ImportWizardDialogFunction = True
				If (DlgText("FixedLengthT")>"0") Then DlgText("FixedLengthT", DlgText("FixedLengthT")-1)
				UpdateDialog()
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Wait 0.1 : ImportWizardDialogFunction = True ' Continue getting idle actions
		Select Case DlgFocus()	' This should be the item that was changed last
			Case "SeparatorDLB"
				If (SepString<>MapString(DlgText("SeparatorDLB"), SepArray, SepArrayMAP)) Then
					DlgText("NHeaderLinesT",CStr(AutoDetectHeaderLines(DataContents, DlgValue("FileFormatOG"), DlgText("SeparatorDLB"))))
					UpdateDialog()
				End If
			Case "CommentCharsT"
				If (DlgText("CommentCharsT")<>ComChars) Then
					UpdateDialog()
				End If
			Case "NHeaderLinesT"
				If (DlgText("NHeaderLinesT")<>NHeaderLines) Then
					UpdateDialog()
				End If
			Case "FixedLengthT"
				If (DlgText("FixedLengthT")<>"") Then
					If(CLng(DlgText("FixedLengthT"))<>FixedLength) Then
						UpdateDialog()
					End If
				End If
		End Select
	Case 6 ' Function key
	End Select
End Function

Function InitDialog() As Integer
	Dim i As Long, minHeader As Long, tmpHeader As Long, minHeaderIndex As Long
	Dim tmpFileName As String, tmpFileExtension As String

	ReDim ShowColumnsArray(0)
	ShowColumnsArray(0) = "1-5"
	DlgListBoxArray("ShowColumnsDLB", ShowColumnsArray)
	DlgValue("ShowColumnsDLB",0)
	DlgValue("SeparatorDLB", 0)
	DlgText("FixedLengthT", "1")
	DlgText("NHeaderLinesT", "0")

	tmpFileName = GetFilePath("*.*", "All files|*.*|CSV files|*.csv|TSV files|*.tsv|TXT files|*.txt", "", "Please select data file", 0+4)
	If tmpFileName = "" Then Exit All ' User pressed cancel in file selection dialog
	tmpFileExtension = Split(tmpFileName, ".")(UBound(Split(tmpFileName,".")))
	' Check if file has TOUCHSTONE extension
	If (Left(tmpFileExtension,1) = "s" And Right(tmpFileExtension,1) = "p" And IsNumeric(Mid(tmpFileExtension,2,Len(tmpFileExtension)-2))) Then
		If (MsgBox("The file extension indicates that the file is in Touchstone format,"+vbNewLine _
					+"which is not supported by this wizard. Would you like to start the Touchstone import?", vbYesNo, "Touchstone Format Check") = vbYes) Then
			RunScript(GetInstallPath+"\Library\Macros\Results\- Import and Export\Import Touchstone File^+MWS+DS.mcr")
			Exit All
		End If
	End If

	If ((Dir(tmpFileName) = "") And (DataFileName = "")) Then
		'MsgBox("Please select a valid file.", "File error")
		' Disable all dialog items except for "Close" and "Change" (file)
		For i = 0 To DlgCount-1
			DlgEnable(i, False)
		Next
		DlgEnable("ClosePB", True)
		DlgEnable("ChangeFilePB",True)
		DlgEnable("ReadClipboardPB", True)
		InitDialog = 0 ' Error
	ElseIf (tmpFileName = "") Then
		' Do nothing, simply keep old existing file
	Else
		DataFileName = tmpFileName
		DlgText("DataFileLocationT",DataFileName)
		DataContents = ReadDataFromFile(DataFileName)
		' Make an educated guess on what the separator could be:
		If(Right(DataFileName, 4) = ".tsv") Then
			DlgValue("FileFormatOG", 1)
			DlgText("SeparatorDLB", "Tab")
		ElseIf(Right(DataFileName, 4) = ".csv") Then
			DlgValue("FileFormatOG", 1)
			DlgText("SeparatorDLB", "Comma")
		Else ' Try auto
			DlgValue("FileFormatOG", 0)
			AutoDetectHeaderLines(DataContents, 0, "")
		End If
		DlgText("NHeaderLinesT",CStr(AutoDetectHeaderLines(DataContents, DlgValue("FileFormatOG"), IIf(DlgValue("FileFormatOG") = 1, DlgText("SeparatorDLB"), DlgText("FixedLengthT")))))
		UpdateDialog
		InitDialog = 1 ' All OK
	End If

End Function

Function UpdateDialog() As Integer
	Dim i As Long, j As Long, jmod5 As Long
	Dim sTempString(4) As String
	Dim jmin As Long, jmax As Long
	Dim iCurrentLocale As Long

	' Enable all dialog items (in case they were disabled before due to missing file)
	For i = 0 To DlgCount-1
		DlgEnable(i, True)
	Next

	DlgEnable("SeparatorDLB", DlgValue("FileFormatOG") = 1)
	DlgEnable("FixedLengthT", DlgValue("FileFormatOG") = 2)

	'DlgText("dataFileLocationT", DataFileName)
	DlgText("headerDisplay", HeaderAndComments)
	DlgText("dataDisplay1", "Updating...")
	For i = 2 To UBound(DataString)+1
		DlgText("dataDisplay"+CStr(i), "")
	Next
	' Store the currently used separator in a global variable
	SepString = MapString(DlgText("SeparatorDLB"), SepArray, SepArrayMAP)

	FixedLength = CLng(DlgText("FixedLengthT"))

	DataArray = FillDataArrayFromString(DataContents, _
										IIf(DlgValue("ColRowOG")=0, "COL", "ROW"), _
										DlgValue("FileFormatOG"), _
										IIf(DlgValue("FileFormatOG") = 1, SepString, CLng(DlgText("FixedLengthT"))), _
										DlgText("NHeaderLinesT"), _
										DlgText("CommentCharsT"))

	ReDim ShowColumnsArray(Fix((UBound(DataArray,1)+1)/5))
	For i = 0 To UBound(ShowColumnsArray)
		ShowColumnsArray(i) = CStr(1+5*i) + "-" + CStr(5+5*i)
	Next
	DlgListBoxArray("ShowColumnsDLB", ShowColumnsArray)
	DlgValue("ShowColumnsDLB", IIf(SelectedShowColumns<=UBound(ShowColumnsArray), SelectedShowColumns, 0))

	For i = 0 To UBound(DataString)
		DataString(i) = ""
	Next i

	' Switch to US locale temporaliy to format String, then switch back
	iCurrentLocale = GetLocale
	SetLocale(&H409) ' &H409 = US

	If (Not IsNumeric(DataArray(0,0))) Then
		DataString(0) = "Inconsistent or no numerical data found. Please check your settings."
		DlgEnable("NextPB", False)
	Else
		DlgEnable("NextPB", True)
		jmin = 0+5*DlgValue("ShowColumnsDLB")
		jmax = Min_LIB(Array(UBound(DataArray,1), 4+5*DlgValue("ShowColumnsDLB")))
		For j = jmin To jmax
			sTempString(j Mod 5) = ""
		Next
		For i = 0 To UBound(DataArray,2)
			For j = jmin To jmax
				jmod5 = j Mod 5
				' Appending to strings becomes very slow for long strings
				If IsNumeric(DataArray(j,i)) Then
					sTempString(jmod5) = sTempString(jmod5) + Format(DataArray(j, i),"+0.0000000000e+000;-0.0000000000e+000;+0.0000000000e+000")  + vbNewLine
				Else
					sTempString(jmod5) = sTempString(jmod5) + "NaN"  + vbNewLine
				End If
			Next
			If (i Mod 100 = 0) Then DlgText("dataDisplay1", "Updating... "+vbNewLine+CStr(i)+"/"+Cstr(UBound(DataArray,2)))
			If (i Mod 1000 = 0) Then
				For j = jmin To jmax
					jmod5 = j Mod 5
					DataString(jmod5) = DataString(jmod5) + sTempString(jmod5)
					sTempString(jmod5) = ""
				Next
			End If
		Next
		' flush buffer
		For j = jmin To jmax
			jmod5 = j Mod 5
			DataString(jmod5) = DataString(jmod5) + sTempString(jmod5)
			' Now remove all commas, they can conflict with VBA code
			DataString(jmod5) = Replace(DataString(jmod5), ",", "")

			sTempString(jmod5) = ""
		Next

		' FSR: Alternative code below might be a little bit faster but leaves MWS non-responsive and does not allow progress feedback
		'Dim tmpArray() As String
		'ReDim tmpArray(UBound(DataArray,2))
		'For j = jmin To jmax
		'	For i = 0 To UBound(DataArray,2)
		'		tmpArray(i) = USFormat(DataArray(j, i),"+0.0000000000e+000; -0.0000000000e+000;+0.0000000000e+000")
		'	Next
		'	DataString(j Mod 5) = Join(tmpArray,vbNewLine)
		'Next
		' FSR End alternative code
	End If

	' Restore original locale setting
	SetLocale(iCurrentLocale)

	DlgText("headerDisplay", HeaderAndComments) ' Call again in case any comments have been found, etc.
	For i = 0 To UBound(DataString)
		DlgText("dataDisplay" + CStr(i+1), DataString(i))
	Next i

End Function

Function ReadDataFromFile(DataFileName As String) As String
	Dim DataFile As Long, DataStringLength As Long
	DataFile = FreeFile()
	Open DataFileName For Input As #DataFile
	ReadDataFromFile = Input(LOF(DataFile), DataFile)
	Close DataFile
	' Replace different line breaks with vbNewLine, order is important
	ReadDataFromFile = Replace(ReadDataFromFile, vbCrLf, vbLf)
	ReadDataFromFile = Replace(ReadDataFromFile, vbLf, vbCr)
	ReadDataFromFile = Replace(ReadDataFromFile, vbCr, vbNewLine)
	' In case of multiple line breaks: Remove double line breaks
	DataStringLength = -1
	While (Len(ReadDataFromFile)<>DataStringLength)
		DataStringLength = Len(ReadDataFromFile)
		ReadDataFromFile = Replace(ReadDataFromFile, vbNewLine+vbNewLine, vbNewLine) ' remove double line breaks
		ReadDataFromFile = Replace(ReadDataFromFile, "  ", " ") ' remove multiple spaces
	Wend

End Function

Function AutoDetectHeaderLines(LocalDataContents As String, LocalFileFormat As Integer, ByVal SeparatorOrLength As String) As Long
	Dim DataContentLines() As String
	Dim tmpLine As String, tmpLineLength
	Dim Separator As String
	Dim EntryLength As Long
	Dim i As Long, LinePosition As Long

	AutoDetectHeaderLines = 0
 	HeaderAndComments = ""
	DataContentLines = Split(LocalDataContents, vbNewLine)

	LinePosition = 0
	tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))

	Select Case LocalFileFormat
		Case 0 ' Auto detect
			While (LinePosition < UBound(DataContentLines))
				i = 0
				While (i<=Len(tmpLine) And Not IsNumeric(Mid(tmpLine, 1, i)))
					i = i + 1
				Wend
				If (i = Len(tmpLine)+1) Then ' no numeric entry found, line is header
					AutoDetectHeaderLines =  AutoDetectHeaderLines + 1
					HeaderAndComments = HeaderAndComments + tmpLine  + vbNewLine
					LinePosition = LinePosition + 1
					tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))
				Else ' numeric entry found, leave loop
					Exit While
				End If
			Wend
		Case 1 ' char separated
			' If Separator is not 'ByVal' in the function declaration, the line below will mess up SepArray!
			Separator = MapString(SeparatorOrLength, SepArray, SepArrayMAP)
			While (LinePosition < UBound(DataContentLines)) And Not IsNumeric(Split(tmpLine, Separator)(0))
				AutoDetectHeaderLines =  AutoDetectHeaderLines + 1
				HeaderAndComments = HeaderAndComments + tmpLine  + vbNewLine
		 		LinePosition = LinePosition + 1
				tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))
			Wend
		Case 2 ' fixed length
			EntryLength = CLng(SeparatorOrLength)
			While (LinePosition < UBound(DataContentLines)) And Not IsNumeric(Mid(tmpLine, EntryLength))
				AutoDetectHeaderLines =  AutoDetectHeaderLines + 1
				HeaderAndComments = HeaderAndComments + tmpLine  + vbNewLine
		 		LinePosition = LinePosition + 1
				tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))
			Wend
	End Select

End Function

Function MapString(ByVal QueryString As String, ByVal localSepArray As Variant, ByVal localSepArrayMAP As Variant) As String
	' Return value in DataArrayMAP at position where QueryString matches entry in DataArray.
	' If QueryString is not in DataArray, return QueryString
	Dim i As Long

	MapString = QueryString
	For i = 0 To UBound(localSepArray)-1
		If (QueryString = CStr(localSepArray(i))) Then
			MapString = CStr(localSepArrayMAP(i))
			Exit For
		End If
	Next

End Function


Function FillDataArrayFromString(LocalDataContents As String, _
								COLorROW As String, _
								LocalFileFormat As Integer, _
								ByVal SeparatorOrLength As String, _
								ByVal SkipLines As Long, _
								ByVal CommentChars As String) As Variant

	Dim DataContentLines() As String, LinePosition As Long
	Dim localArray() As String, localArray2() As String
	Dim DataFile As Long
	Dim tmpLine As String, nTmpLineLength As Long
	Dim NCols As Long
	Dim nValidEntries As Long
	Dim nTempLinePosition As Long
	Dim i As Long, j As Long, k As Long, lastBreak As Long
	Dim iCurrentLocale As Long

	' Switch to US locale temporaliy to format String, then switch back
	iCurrentLocale = GetLocale
	SetLocale(&H409) ' &H409 = US

	ComChars = CommentChars
	FileFormat = LocalFileFormat
	NHeaderLines = SkipLines
	ColRowFormat = COLorROW
	HeaderAndComments = ""

	DataContentLines = Split(LocalDataContents, vbNewLine)

	LinePosition = 0
	tmpLine = IIf(DataContentLines(LinePosition) = "", "void", Trim(DataContentLines(LinePosition)))
	nTmpLineLength = Len(tmpLine)
	i = 0

	' Skip header and comment lines
	While (((i<SkipLines) Or ((CommentChars <> "") And (InStr(CommentChars,Mid(tmpLine, 1, 1))>0))) And (LinePosition < UBound(DataContentLines)))
		HeaderAndComments = HeaderAndComments + tmpLine + vbNewLine
		LinePosition = LinePosition + 1
		tmpLine = IIf(DataContentLines(LinePosition) = "", "void", Trim(DataContentLines(LinePosition)))
		nTmpLineLength = Len(tmpLine) ' store in a separate variable to avoid multiple calls as this value will be required multiple times in the loop(s) below
		i = i+1
	Wend
	ReDim localArray(0,0)
	' Check if last line has been reached (data is all header and comment)
	If (LinePosition < UBound(DataContentLines)) Then
		Select Case FileFormat
			Case 0 ' auto detect power algo
				NCols = 0
				lastBreak = 0
				i = 1
				' Go through algorithm once for one line just to dectect number of columns
				' Wind forward until first number is found
				For i = 1 To nTmpLineLength
					If IsNumeric(Mid(tmpLine, 1, i)+"0") Then Exit For ' Add a 0 to detect cases like "-1"
				Next
				If (i = nTmpLineLength + 1) Then
					localArray(0,0) = "void"
					FillDataArrayFromString = localArray
					Exit Function
				End If
				lastBreak = i
				' Now parse rest of line. Next number starts where previous is not numeric anymore
				For i = lastBreak To nTmpLineLength
					While (i<=nTmpLineLength And IsNumeric(Mid(tmpLine, lastBreak, i-lastBreak+1)+"0")) ' Add a 0 to detect cases like "3e..."
						i = i+1
					Wend
					' When number is not numeric anymore, increase NCols, reset lastBreak
					NCols = NCols + 1
					lastBreak = i
				Next
				ReDim localArray(NCols-1, UBound(DataContentLines))
				nValidEntries = 0
				' Now do the same thing again for each line, this time also store values
				nTempLinePosition = LinePosition
				For LinePosition = nTempLinePosition To UBound(DataContentLines) - 1
					If (InStr(CommentChars,Mid(tmpLine, 1, 1))=0) Then ' first char in tmpLine is not in CommentChars
						j = 0
						lastBreak = 0
						' Wind forward until first number is found
						For i = 1 To nTmpLineLength
							If IsNumeric(Mid(tmpLine, 1, i)+"0") Then Exit For ' Add a 0 to detect cases like "-1"
						Next
						lastBreak = i
						' Now parse rest of line. Next number starts where previous is not numeric anymore
						For i = lastBreak To nTmpLineLength
							For k = i To nTmpLineLength
								If Not IsNumeric(Mid(tmpLine, lastBreak, k-lastBreak+1)+"0") Then Exit For ' Add a 0 to detect cases like "3e..."
							Next
							i = k
							' When number is not numeric anymore, increase NCols, reset lastBreak
							j = j + 1
							If (j > UBound(localArray, 1)+1) Then
								MsgBox("Auto format detection failed. Please try the manual format options.", "Cannot Determine Format")
								LinePosition = UBound(DataContentLines) - 1
								Exit For
							End If
							localArray(j-1, nValidEntries) = Mid(tmpLine, lastBreak, i-lastBreak)
							lastBreak = i
						Next
					Else
						HeaderAndComments = HeaderAndComments + tmpLine + vbNewLine
					End If
					tmpLine = Trim(DataContentLines(LinePosition + 1))
					If Len(tmpLine) = 0 Then tmpLine = "void" ' Len(string) = 0 is faster than to test for string = ""
					nTmpLineLength = Len(tmpLine) ' store in a separate variable to avoid multiple calls as this value will be required multiple times in this loop
					nValidEntries = nValidEntries + 1
				Next
				' Remove last line again when done
				ReDim Preserve localArray(NCols-1, nValidEntries-1)
			Case 1 ' char-separated
				' Store number of header lines, separator, ColOrRow, and comment chars in globals variables
				SepString = SeparatorOrLength

				NCols = UBound(Split(tmpLine,SepString))+1
				' Remove last column if empty
				NCols = IIf(Split(tmpLine, SepString)(NCols-1) = "", NCols-1, NCols)
				ReDim localArray(NCols-1, 0)
				For i = 1 To NCols
					localArray(i-1, 0) = Split(tmpLine, SepString)(i-1)
				Next
				While (LinePosition < UBound(DataContentLines))
					LinePosition = LinePosition + 1
					tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))
					If (InStr(CommentChars,Mid(tmpLine, 1, 1))=0) Then ' first char in tmpLine is not in CommentChars
						ReDim Preserve localArray(NCols-1, UBound(localArray,2)+1)
						For i = 1 To NCols
							On Error GoTo FormatError1
								localArray(i-1, UBound(localArray,2)) = Split(tmpLine, SepString)(i-1)
								GoTo NoFormatError1
							FormatError1:
								Exit For
							NoFormatError1:
								On Error GoTo 0
						Next
					Else
						HeaderAndComments = HeaderAndComments + tmpLine + vbNewLine
					End If
				Wend
			Case 2 ' fixed length entries
				FixedLength = CLng(SeparatorOrLength)
				NCols = Len(tmpLine)/FixedLength
				' Remove last column if empty
				'NCols = IIf(Split(tmpLine, Separator)(NCols-1) = "", NCols-1, NCols)
				ReDim localArray(NCols-1, 0)
				For i = 0 To NCols-1
					localArray(i, 0) = Mid(tmpLine, 1+FixedLength*i, FixedLength)
				Next
				While (LinePosition < UBound(DataContentLines))
					LinePosition = LinePosition + 1
					tmpLine = IIf(DataContentLines(LinePosition) = "", "void", DataContentLines(LinePosition))
					If (InStr(CommentChars,Mid(tmpLine, 1, 1))=0) Then ' first char in tmpLine is not in CommentChars
						ReDim Preserve localArray(NCols-1, UBound(localArray,2)+1)
						For i = 0 To NCols-1
							On Error GoTo FormatError2
								localArray(i, UBound(localArray,2)) = Mid(tmpLine, 1+FixedLength*i, FixedLength)
								GoTo NoFormatError2
							FormatError2:
								Exit For
							NoFormatError2:
								On Error GoTo 0
						Next
					Else
						HeaderAndComments = HeaderAndComments + tmpLine + vbNewLine
					End If
				Wend
		End Select
	End If

	' Restore original locale
	SetLocale(iCurrentLocale)

	' Invert matrix if necessary
	If (COLorROW = "COL") Then
		FillDataArrayFromString = localArray
	ElseIf (COLorROW = "ROW") Then
		ReDim localArray2(UBound(localArray,2), UBound(localArray,1))
		For i = 0 To UBound(localArray,1)
			For j = 0 To UBound(localArray,2)
				localArray2(j,i) = localArray(i,j)
			Next
		Next
		FillDataArrayFromString = localArray2
	End If

End Function

' ------------------------ FUNCTIONS FOR SECOND DIALOG WINDOW START HERE ---------------------------------

Function OpenImportDialog() As Integer

	Dim importArrayColsList() As String
	Dim i As Long

	ReDim importArrayColsList(UBound(DataArray,1))
	For i = 0 To UBound(importArrayColsList)
		importArrayColsList(i) = CStr(i+1)
	Next
	ReDim XDataColumn(0)
	ReDim Y1DataColumn(0)
	ReDim Y2DataColumn(0)

	Begin Dialog UserDialog 880,581,"Import data group",.OpenImportDialogFunction ' ' %GRID:10,7,1,1
		GroupBox 10,119,290,427,"Abscissa",.GroupBox1
		GroupBox 10,7,860,105,"General",.GroupBox3
		GroupBox 310,119,560,427,"Ordinate",.GroupBox2
		Text 30,63,40,14,"Title:",.Text1
		TextBox 130,56,350,21,.TitleT
		Text 30,189,60,14,"Column:",.Text2
		CheckBox 180,224,100,14,"Angle in deg",.PolarXAngleCB
		TextBox 30,245,260,266,.XDataPreview,2
		Text 320,189,50,14,"Column:",.Text3
		TextBox 320,245,260,266,.Y1DataPreview,2
		Text 600,189,50,14,"Column:",.Text4
		TextBox 600,245,260,266,.Y2DataPreview,2
		DropListBox 200,182,90,192,importArrayColsList(),.xColDLB
		DropListBox 390,182,90,192,importArrayColsList(),.y1ColDLB
		DropListBox 660,182,80,192,importArrayColsList(),.y2ColDLB
		Text 30,147,40,14,"Label:",.Text5
		TextBox 90,140,200,21,.AbscissaLabelT
		Text 320,147,40,14,"Label:",.Text6
		TextBox 370,140,210,21,.OrdinateLabelT
		OptionGroup .ReMagOG
			OptionButton 490,182,60,14,"Real",.RealOB
			OptionButton 490,196,60,14,"Mag",.MagOB
		OptionGroup .ImPhOG
			OptionButton 750,182,60,14,"Imag",.ImagOB
			OptionButton 750,196,70,14,"Phase",.PhaseOB
		Text 650,147,100,14,"Data/plot type:",.Text7
		DropListBox 750,140,90,192,DataPlotType(),.DataPlotTypeDLB
		CheckBox 600,224,110,14,"Phase in deg",.PhaseDegCB
		Text 30,35,100,14,"Preset formats:",.Text8
		DropListBox 130,28,350,192,PresetsArray(),.PresetsDLB
		CheckBox 30,91,330,14,"Import multiple data groups with a periodic shift of",.ImportMultipleCB
		TextBox 370,84,40,21,.PeriodicityT
		Text 420,91,60,14,"columns.",.Text11
		CheckBox 330,224,160,14,"Mag in dB, dB factor:",.MagdBCB
		TextBox 490,217,90,21,.dBFactorT
		Text 20,553,440,14,"Import button disabled, number of samples in columns does not match.",.NYSamplesWarningT
		CheckBox 30,224,130,14,"Sort by x values",.SortByXCB

		Text 30,525,100,14,"Rescale x data:",.Text9
		PushButton 140,518,50,21,"x (-1)",.xNegPB
		PushButton 190,518,50,21,"x 1e3",.X1e3PB
		PushButton 240,518,50,21,"x 1e-3",.X1em3PB

		Text 320,525,100,14,"Rescale y data:",.Text10
		PushButton 430,518,50,21,"x (-1)",.yNegPB
		PushButton 480,518,50,21,"x 1e3",.Y1e3PB
		PushButton 530,518,50,21,"x 1e-3",.Y1em3PB

		PushButton 680,553,90,21,"Import",.ImportPB
		PushButton 580,553,90,21,"Change File",.ChangeFilePB
		PushButton 480,553,90,21,"<< Back",.BackPB
		PushButton 780,553,90,21,"Exit",.ExitPB

	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = -1) Then

	Else

	End If

End Function

Rem See DialogFunc help topic for more information.
Dim dXScaleFactor As Double, dYScaleFactor As Double ' need these as global variables to scale x/y axes in case "multiple imports" is active

Private Function OpenImportDialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Dim i As Long, j As Long
	Dim Periodicity As Long ' Used for reading in multiple results
	Dim tmpFileName As String
	Dim importSuccess As Integer
	Dim sStartTitle As String
	Dim nStartColumnX As Long, nStartColumnY1 As Long, nStartColumnY2 As Long

	Select Case Action%
	Case 1 ' Dialog box initialization
		dXScaleFactor = 1
		dYScaleFactor = 1
		' By default, allow only "1D" as option
		ReDim DataPlotType(1)
		DataPlotType(0) = "1D"
		DataPlotType(1) = "1D Polar"
		' If DataArray has more than 2 columns, allow also "1DC"
		If (UBound(DataArray,1) > 1) Then
			ReDim Preserve DataPlotType(UBound(DataPlotType)+1)
			DataPlotType(UBound(DataPlotType)) = "1DC"
		End If
		DlgListBoxArray("DataPlotTypeDLB", DataPlotType())
		DlgValue("DataPlotTypeDLB", 0)

		DlgText("TitleT",Split(DataFileName, "\")(UBound(Split(DataFileName, "\"))))
		DlgText("AbscissaLabelT","x")
		DlgText("OrdinateLabelT","y")
		DlgValue("xColDLB",0)
		DlgValue("PolarXAngleCB",1) ' for 1D Polar settings only, assume "degree" as default
		DlgValue("y1ColDLB",1)
		DlgValue("y2ColDLB",2)
		DlgText("dBFactorT", "20")
		DlgText("PeriodicityT",Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"),DlgValue("y2ColDLB")))-Min_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"),DlgValue("y2ColDLB"))))
		DlgEnable("ImPhOG",False)
		If InStr(HeaderAndComments, "# Created by Antenna Magus") Then ' AM version >=2.5
			DlgValue("PresetsDLB",FindListIndex(PresetsArray, "Antenna Magus (v2.5 and up)"))
			DlgText("DataPlotTypeDLB", "1D")
			DlgValue("xColDLB",0)
			DlgValue("y1ColDLB",1)
			DlgValue("y2ColDLB",2)
		End If
		'Deactivate 1st column if only 1 column in array
		If (UBound(DataArray,1)=0) Then
			DlgEnable("xColDLB",False)
		Else
			XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
			UpdateDataColumn("XDataPreview", XDataColumn)
		End If
		Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
		UpdateDataColumn("Y1DataPreview", Y1DataColumn)
		'Deactivate 3rd column if only 2 columns in array
		If (UBound(DataArray,1)=1) Then
			DlgEnable("y2ColDLB",False)
		Else
			Y2DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y2ColDLB"))
			UpdateDataColumn("Y2DataPreview", Y2DataColumn)
		End If
	Case 2 ' Value changing or button pressed
		Rem ImportDialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "PresetsDLB"
				Select Case DlgText("PresetsDLB")
					Case "Antenna Magus (up to v2.5)", "Antenna Magus (v2.5 and up)"
						DlgText("DataPlotTypeDLB", "1D")
						DlgValue("xColDLB",0)
						DlgValue("y1ColDLB",1)
						DlgValue("y2ColDLB",2)
					Case Else
						' Do nothing
				End Select
				XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
				UpdateDataColumn("XDataPreview", XDataColumn)
				Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
				UpdateDataColumn("Y1DataPreview", Y1DataColumn)
				Y2DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y2ColDLB"))
				UpdateDataColumn("Y2DataPreview", Y2DataColumn)
			Case "xColDLB"
				XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
				UpdateDataColumn("XDataPreview", XDataColumn)
			Case "y1ColDLB"
				Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
				UpdateDataColumn("Y1DataPreview", Y1DataColumn)
			Case "y2ColDLB"
				Y2DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y2ColDLB"))
				UpdateDataColumn("Y2DataPreview", Y2DataColumn)
			Case "xNegPB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(XDataColumn)
					XDataColumn(i) = -XDataColumn(i)
				Next
				dXScaleFactor = -dXScaleFactor
				UpdateDataColumn("XDataPreview",XDataColumn)
			Case "X1e3PB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(XDataColumn)
					XDataColumn(i) = 1e3 * XDataColumn(i)
				Next
				dXScaleFactor = 1e3 * dXScaleFactor
				UpdateDataColumn("XDataPreview",XDataColumn)
			Case "X1em3PB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(XDataColumn)
					XDataColumn(i) = 1e-3 * XDataColumn(i)
				Next
				dXScaleFactor = 1e-3 * dXScaleFactor
				UpdateDataColumn("XDataPreview",XDataColumn)
			Case "yNegPB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(Y1DataColumn)
					Y1DataColumn(i) = -Y1DataColumn(i)
					If i <= UBound(Y2DataColumn) Then Y2DataColumn(i) = -Y2DataColumn(i)
				Next
				dYScaleFactor = -dYScaleFactor
				UpdateDataColumn("Y1DataPreview", Y1DataColumn)
				UpdateDataColumn("Y2DataPreview", Y2DataColumn)
			Case "Y1e3PB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(Y1DataColumn)
					Y1DataColumn(i) = Y1DataColumn(i)*1e3
					If i <= UBound(Y2DataColumn) Then Y2DataColumn(i) = Y2DataColumn(i)*1e3
				Next
				dYScaleFactor = 1e3 * dYScaleFactor
				UpdateDataColumn("Y1DataPreview", Y1DataColumn)
				UpdateDataColumn("Y2DataPreview", Y2DataColumn)
			Case "Y1em3PB"
				OpenImportDialogFunction = True
				For i = 0 To UBound(Y1DataColumn)
					Y1DataColumn(i) = Y1DataColumn(i)*1e-3
					If i <= UBound(Y2DataColumn) Then Y2DataColumn(i) = Y2DataColumn(i)*1e-3
				Next
				dYScaleFactor = 1e-3 * dYScaleFactor
				UpdateDataColumn("Y1DataPreview", Y1DataColumn)
				UpdateDataColumn("Y2DataPreview", Y2DataColumn)
			Case "ChangeFilePB"
				' Used to quickly load another file without going through the initial settings again
				OpenImportDialogFunction = True
				tmpFileName = GetFilePath("*.*", "All files|*.*|CSV files|*.csv|TSV files|*.tsv|TXT files|*.txt", "", "Please select data file", 0+4)
				If (Dir(tmpFileName) = "") Then ' User pressed cancel (selecting non-existing files should not be possible)
					' Do nothing, just keep existing file
					ReportWarningToWindow("Failed to load file "+tmpFileName)
				Else
					DataFileName = tmpFileName
					DataContents = ReadDataFromFile(DataFileName)
					Dim tmpDataArray() As String
					tmpDataArray = FillDataArrayFromString(DataContents, _
											ColRowFormat, _
											FileFormat, _
											SepString, _
											NHeaderLines, _
											ComChars)
					If (UBound(tmpDataArray,1)<>UBound(DataArray,1)) Then
						MsgBox("New file could not be loaded. Data format of new file is not identical to old file.","Error")
						Exit Function
					Else
						DataArray = tmpDataArray
						XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
						UpdateDataColumn("XDataPreview", XDataColumn)
						Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
						UpdateDataColumn("Y1DataPreview", Y1DataColumn)
						Y2DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y2ColDLB"))
						UpdateDataColumn("Y2DataPreview", Y2DataColumn)
					End If
				End If
			Case "ImportPB"
				DlgEnable("ImportPB", False)
				Periodicity = IIf(DlgValue("ImportMultipleCB") = 0, UBound(DataArray,1), CLng(DlgText("PeriodicityT")))
				i = 1
				sStartTitle = DlgText("TitleT")
				nStartColumnX = DlgValue("xColDLB")
				nStartColumnY1 = DlgValue("y1ColDLB")
				nStartColumnY2 = DlgValue("y2ColDLB")
				If (DlgText("DataPlotTypeDLB") = "1D") Then ' 1D type
					' 1 regular import always, then multiples if required
					importSuccess = ImportAs1D(XDataColumn,Y1DataColumn, "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					While (Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"))) + Periodicity <= UBound(DataArray,1))
						DlgValue("xColDLB", DlgValue("xColDLB")+Periodicity)
						XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
						If (dXScaleFactor <> 1) Then
							For j = 0 To UBound(XDataColumn)
								XDataColumn(j) = dXScaleFactor * XDataColumn(j)
							Next
						End If
						UpdateDataColumn("XDataPreview", XDataColumn)
						DlgValue("y1ColDLB", DlgValue("y1ColDLB")+Periodicity)
						Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
						If (dYScaleFactor <> 1) Then
							For j = 0 To UBound(Y1DataColumn)
								Y1DataColumn(j) = dYScaleFactor * Y1DataColumn(j)
							Next
						End If
						UpdateDataColumn("Y1DataPreview", Y1DataColumn)
						' FSR 04/18/2014: The following 'if' part is needed for periodic imports
						If (DlgText("PresetsDLB") = "Antenna Magus") Then
							DlgText("TitleT", Replace(Trim(Split(Split(HeaderAndComments,vbNewLine)(0),vbTab)(DlgValue("y1ColDLB"))) + ", " + Trim(Split(Split(HeaderAndComments,vbNewLine)(1),vbTab)(DlgValue("y1ColDLB"))),Chr(34),""))
							DlgText("AbscissaLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("xColDLB")),Chr(34),"")))
							DlgText("OrdinateLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("y1ColDLB")),Chr(34),"")))
						Else
							DlgText("TitleT", sStartTitle+"_"+Cstr(i))
							i = i+1
						End If
						importSuccess = ImportAs1D(XDataColumn,Y1DataColumn, "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					Wend
				ElseIf (DlgText("DataPlotTypeDLB") = "1D Polar") Then ' type 1D Polar
					' 1 regular import always, then multiples if required
					importSuccess = ImportAs1DPolar(XDataColumn,Y1DataColumn, DlgValue("PolarXAngleCB")=1, CBool(DlgValue("MagdBCB")), Evaluate(DlgText("dBFactorT")), "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					While (Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"))) + Periodicity <= UBound(DataArray,1))
						DlgValue("xColDLB", DlgValue("xColDLB")+Periodicity)
						XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
						If (dXScaleFactor <> 1) Then
							For j = 0 To UBound(XDataColumn)
								XDataColumn(j) = dXScaleFactor * XDataColumn(j)
							Next
						End If
						UpdateDataColumn("XDataPreview", XDataColumn)
						DlgValue("y1ColDLB", DlgValue("y1ColDLB")+Periodicity)
						Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
						If (dYScaleFactor <> 1) Then
							For j = 0 To UBound(Y1DataColumn)
								Y1DataColumn(j) = dYScaleFactor * Y1DataColumn(j)
							Next
						End If
						UpdateDataColumn("Y1DataPreview", Y1DataColumn)
						' FSR 04/18/2014: The following 'if' part is needed for periodic imports
						If (DlgText("PresetsDLB") = "Antenna Magus") Then
							DlgText("TitleT", Replace(Trim(Split(Split(HeaderAndComments,vbNewLine)(0),vbTab)(DlgValue("y1ColDLB"))) + ", " + Trim(Split(Split(HeaderAndComments,vbNewLine)(1),vbTab)(DlgValue("y1ColDLB"))),Chr(34),""))
							DlgText("AbscissaLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("xColDLB")),Chr(34),"")))
							DlgText("OrdinateLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("y1ColDLB")),Chr(34),"")))
						Else
							DlgText("TitleT", sStartTitle+"_"+Cstr(i))
							i = i+1
						End If
						importSuccess = ImportAs1DPolar(XDataColumn,Y1DataColumn, DlgValue("PolarXAngleCB")=1, CBool(DlgValue("MagdBCB")), Evaluate(DlgText("dBFactorT")), "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					Wend
				ElseIf (DlgText("DataPlotTypeDLB") = "1DC") Then ' 1DC type
					' 1 regular import always, then multiples if required
					importSuccess = ImportAs1DC(XDataColumn,Y1DataColumn, Y2DataColumn, DlgValue("ReMagOG")=0, CBool(DlgValue("MagdBCB")), Evaluate(DlgText("dBFactorT")), CBool(DlgValue("PhaseDegCB")), "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					While (Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"),DlgValue("y2ColDLB"))) + Periodicity <= UBound(DataArray,1))
						DlgValue("xColDLB", DlgValue("xColDLB")+Periodicity)
						XDataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("xColDLB"))
						If (dXScaleFactor <> 1) Then
							For j = 0 To UBound(XDataColumn)
								XDataColumn(j) = dXScaleFactor * XDataColumn(j)
							Next
						End If
						UpdateDataColumn("XDataPreview", XDataColumn)
						DlgValue("y1ColDLB", DlgValue("y1ColDLB")+Periodicity)
						DlgValue("y2ColDLB", DlgValue("y2ColDLB")+Periodicity)
						Y1DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y1ColDLB"))
						Y2DataColumn = FillDataColumnFromDataArray(DataArray, DlgValue("y2ColDLB"))
						If (dYScaleFactor <> 1) Then
							For j = 0 To UBound(Y1DataColumn)
								Y1DataColumn(j) = dYScaleFactor * Y1DataColumn(j)
								If j <= UBound(Y2DataColumn) Then Y2DataColumn(j) = dYScaleFactor * Y2DataColumn(j)
							Next
						End If
						UpdateDataColumn("Y1DataPreview", Y1DataColumn)
						UpdateDataColumn("Y2DataPreview", Y2DataColumn)
						' FSR 04/18/2014: The following 'if' part is needed for periodic imports
						If (DlgText("PresetsDLB") = "Antenna Magus") Then
							DlgText("TitleT", Replace(Trim(Split(Split(HeaderAndComments,vbNewLine)(0),vbTab)(DlgValue("y1ColDLB"))) + ", " + Trim(Split(Split(HeaderAndComments,vbNewLine)(1),vbTab)(DlgValue("y1ColDLB"))),Chr(34),""))
							DlgText("AbscissaLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("xColDLB")),Chr(34),"")))
							DlgText("OrdinateLabelT", Trim(Replace(Split(Split(HeaderAndComments,vbNewLine)(2),vbTab)(DlgValue("y1ColDLB")),Chr(34),"")))
						Else
							DlgText("TitleT", sStartTitle+"_"+Cstr(i))
							i = i+1
						End If
						importSuccess = ImportAs1DC(XDataColumn,Y1DataColumn, Y2DataColumn, DlgValue("ReMagOG")=0, CBool(DlgValue("MagdBCB")), Evaluate(DlgText("dBFactorT")), CBool(DlgValue("PhaseDegCB")), "1D Results\Imported Data\", DlgText("TitleT"), DlgText("TitleT"), DlgText("AbscissaLabelT"), DlgText("OrdinateLabelT"), DlgValue("SortByXCB")=1)
					Wend
				Else ' unknown selection in "DataPlotTypeDLB", should not happen
					ReportError("Unknown data/plot type.")
				End If
				' Reset dialog values
				DlgText("TitleT", sStartTitle)
				DlgValue("xColDLB", nStartColumnX)
				DlgValue("y1ColDLB", nStartColumnY1)
				DlgValue("y2ColDLB", nStartColumnY2)
				If (importSuccess = 1) Then
					'ReportInformationToWindow("Import successful.")
					If (GetApplicationName = "DS") Then
						DS.SelectTreeItem(Replace(LastPlotCreated, "Design\", ""))
					ElseIf (Left(GetApplicationName, 7) = "DS for ") Then
						SelectTreeItem(LastPlotCreated)
						DS.SelectTreeItem(Replace(LastPlotCreated, "Design\", ""))
					Else
						SelectTreeItem(LastPlotCreated)
					End If
				Else
					MsgBox("Import failed.")
				End If
				If (DlgText("DataPlotTypeDLB") = "1D Polar") Then Plot1D.PlotView("Polar")
				DlgEnable("ImportPB", True)
				OpenImportDialogFunction = True
			Case "BackPB"
				OpenImportDialogFunction = False
			Case "ExitPB"
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : ImportDialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select

	' These settings should always be applied, independent of the action. Placed at the end of the function so they are executed AFTER init
	If (DlgText("DataPlotTypeDLB") = "1D Polar") Then DlgValue("ReMagOG", 1)
	DlgEnable("ReMagOG", (DlgText("DataPlotTypeDLB") = "1DC"))
	DlgValue("ImPhOG", DlgValue("ReMagOG"))
	DlgVisible("PolarXAngleCB", DlgText("DataPlotTypeDLB") = "1D Polar")
	DlgEnable("PhaseDegCB", (DlgValue("ImPhOG") = 1) And (DlgText("DataPlotTypeDLB") = "1DC"))
	DlgEnable("MagdBCB", (DlgValue("ImPhOG") = 1) And Not (DlgText("DataPlotTypeDLB") = "1D"))
	DlgEnable("dBFactorT", (DlgValue("ImPhOG") = 1) And Not (DlgText("DataPlotTypeDLB") = "1D"))
	DlgEnable("y2ColDLB", (DlgText("DataPlotTypeDLB") = "1DC"))
	DlgEnable("y2DataPreview", (DlgText("DataPlotTypeDLB") = "1DC"))
	DlgVisible("NYSamplesWarningT", Not ((UBound(XDataColumn) = UBound(Y1DataColumn)) And ((DlgText("DataPlotTypeDLB") = "1D") Or (DlgText("DataPlotTypeDLB") = "1D Polar") Or UBound(Y1DataColumn) = UBound(Y2DataColumn)) ))
	DlgEnable("ImportPB", (UBound(XDataColumn) = UBound(Y1DataColumn)) And ((DlgText("DataPlotTypeDLB") = "1D") Or (DlgText("DataPlotTypeDLB") = "1D Polar") Or UBound(Y1DataColumn) = UBound(Y2DataColumn)) )
	If (DlgText("DataPlotTypeDLB") = "1D" Or DlgText("DataPlotTypeDLB") = "1D Polar") Then
		DlgText("PeriodicityT",Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB")))-Min_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB")))+1)
	Else
		DlgText("PeriodicityT",Max_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"),DlgValue("y2ColDLB")))-Min_LIB(Array(DlgValue("xColDLB"),DlgValue("y1ColDLB"),DlgValue("y2ColDLB")))+1)
	End If

End Function

Function FillDataColumnFromDataArray(localDataArray As Variant, columnNumber As Long) As Variant
	' Input: a string data array, assumed to be arranged in columns, and a column number
	' The function then creates a 1D array of type double and copies the values from the
	' selected column of the string data array into the double data array.
	' Returns the 1D array.
	Dim localDataColumn() As Double
	Dim tmpEntry As String
	Dim i As Long, invalidEntries As Long

	invalidEntries = 0

	ReDim localDataColumn(UBound(localDataArray,2))
	On Error GoTo InValidEntry
	For i = 0 To UBound(localDataColumn)
		localDataColumn(i-invalidEntries) = Evaluate(localDataArray(columnNumber, i))
		GoTo ValidEntry

		InValidEntry:
			invalidEntries = invalidEntries + 1
		ValidEntry:
			' Nothing else to do
		 ' If (i Mod 20000 = 0) Then ReportInformationToWindow("Column " & CStr(columnNumber) & "; entry " & CStr(i))
	Next
	On Error GoTo 0
	' Trim localDataColumn
	If (invalidEntries<=UBound(localDataColumn)) Then
		ReDim Preserve localDataColumn(UBound(localDataColumn)-invalidEntries)
	Else
		MsgBox("Selected column does not seem to contain any numeric entries.", "Error")
		ReDim localDataColumn(0)
	End If
	FillDataColumnFromDataArray = localDataColumn
End Function

Function UpdateDataColumn(ColumnName As String, DataColumn As Variant)

	Dim i As Long, ColumnString As String, sTempString As String
	Dim iCurrentLocale As Long

	ColumnString = ""
	sTempString = ""
	DlgText(ColumnName, "Updating...")
	' Switch to US locale temporaliy to format String, then switch back
	iCurrentLocale = GetLocale
	SetLocale(&H409) ' &H409 = US
	For i = 0 To UBound(DataColumn)
		sTempString = sTempString + Format(DataColumn(i),"+0.00000000000e+000; -0.00000000000e+000;+0.00000000000e+000") + vbNewLine
		If (i Mod 1000 = 0) Then
			DlgText(ColumnName, "Updating... " & CStr(i) & "/" & CStr(UBound(DataColumn)))
			ColumnString = ColumnString + sTempString
			sTempString = ""
		End If
	Next
	' flush buffer
	ColumnString = ColumnString + sTempString
	sTempString = ""
	SetLocale(iCurrentLocale) ' switch back to original locale

	DlgText(ColumnName, ColumnString)
	' Adjust axis labels and title:
	On Error GoTo HeaderError
	If (ColumnName = "XDataPreview") Then
		Select Case DlgText("PresetsDLB")
			Case "CST ASCII Export"
				DlgText("AbscissaLabelT", Trim(Split(HeaderAndComments,"]")(DlgValue("xColDLB")))+"]")
			Case "Antenna Magus (up to v2.5)"
				DlgText("AbscissaLabelT", Split(Split(HeaderAndComments,vbNewLine)(2),Chr(34))(2*DlgValue("xColDLB")+1))
			Case "Antenna Magus (v2.5 and up)"
				DlgText("AbscissaLabelT", Split(Split(HeaderAndComments,vbNewLine)(UBound(Split(HeaderAndComments,vbNewLine))-1),Chr(34))(2*DlgValue("xColDLB")+1)) ' last line in HeaderAndComments contains abscissa label
			Case Else
				' default behavior
		End Select
	ElseIf (ColumnName = "Y1DataPreview") Then
		Select Case DlgText("PresetsDLB")
			Case "CST ASCII Export"
				' Header might contain double spaces, remove them iteratively
				ColumnString = HeaderAndComments
				Do
					i = Len(ColumnString)
					ColumnString = Replace(ColumnString,"  ", " ")
				Loop Until (i = Len(ColumnString)) ' if true, no replacements took place
				ColumnString = Trim(ColumnString)
				DlgText("OrdinateLabelT", Trim(Split(HeaderAndComments,"]")(DlgValue("y1ColDLB")))+"]")
			Case "Antenna Magus (up to v2.5)"
				DlgText("TitleT", Split(Split(HeaderAndComments,vbNewLine)(0),Chr(34))(2*DlgValue("y1ColDLB")+1) + ", " + Split(Split(HeaderAndComments,vbNewLine)(1),Chr(34))(2*DlgValue("y1ColDLB")+1))
				DlgText("OrdinateLabelT", Split(Split(HeaderAndComments,vbNewLine)(2),Chr(34))(2*DlgValue("y1ColDLB")+1))
			Case "Antenna Magus (v2.5 and up)"
				If UBound(Split(HeaderAndComments,vbNewLine)) = 4 Then
					DlgText("TitleT", Split(Split(HeaderAndComments,vbNewLine)(2),Chr(34))(2*DlgValue("y1ColDLB")+1))
					DlgText("OrdinateLabelT", Split(Split(HeaderAndComments,vbNewLine)(3),Chr(34))(2*DlgValue("y1ColDLB")+1))
				ElseIf UBound(Split(HeaderAndComments,vbNewLine)) = 5 Then
					DlgText("TitleT", Split(Split(HeaderAndComments,vbNewLine)(2),Chr(34))(2*DlgValue("y1ColDLB")+1) + ", " + Split(Split(HeaderAndComments,vbNewLine)(3),Chr(34))(2*DlgValue("y1ColDLB")+1))
					DlgText("OrdinateLabelT", Split(Split(HeaderAndComments,vbNewLine)(4),Chr(34))(2*DlgValue("y1ColDLB")+1))
				End If
			Case Else
				' default behavior
		End Select
	End If
	GoTo NoHeaderError
	HeaderError:
		MsgBox("Cannot interpret header.", "Error")
		DlgValue("PresetsDLB",0)
	NoHeaderError:
		On Error GoTo 0
End Function

Function ImportAs1D(xImportColumn As Variant, _
					yImportColumn As Variant, _
					treePath As String, _
					importName As String, _
					importTitle As String, _
					xLabel As String, _
					yLabel As String, _
					bSortByX As Boolean) As Integer
	' This function takes an array as input, as well as 2 column numbers to determine which columns of the array contain the data
	' The function then imports the array as a 1D plot, saves it as importName under treePath and assigns title and labels to the plot
	' Data is assumed to be arranged in columns in importArray
	Dim importedObject As Object, importedObjectDS As Object
	Dim i As Long, j As Long
	Dim bUnsorted As Boolean, dTempX As Double, dTempY As Double

	If (UBound(xImportColumn) <> UBound(yImportColumn)) Then
		ImportAs1D = 0 ' Error
		Exit Function
	End If

	Set importedObject = Result1D("")
	For i = 0 To UBound(xImportColumn)
		importedObject.AppendXY(xImportColumn(i), yImportColumn(i))
	Next
	If bSortByX Then
		importedObject.SortByX
	End If
	importedObject.xlabel(xLabel)
	importedObject.ylabel(yLabel)
	importedObject.title(importTitle)

	' Create a duplicate object in case of "DS for "...
	If (Left(GetApplicationName,7) = "DS for ") Then
		Set importedObjectDS = DS.Result1D("")
		For i = 0 To UBound(xImportColumn)
			importedObjectDS.AppendXY(xImportColumn(i), yImportColumn(i))
		Next
		importedObjectDS.xlabel(xLabel)
		importedObjectDS.ylabel(yLabel)
		importedObjectDS.title(importTitle)
	End If

	importedObject.DeleteAt("never")
	If (GetApplicationName = "DS") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName) ' DS requires exactly one subfolder under "Results"
		importedObject.AddToTree(LastPlotCreated)
	ElseIf (Left(GetApplicationName, 7) = "DS for ") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		importedObjectDS.Save(NoForbiddenFilenameCharacters(importName+"DS.sig"))  ' DS requires .sig extension for 1D results

		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName) ' DS requires exactly one subfolder under "Results"
		importedObjectDS.AddToTree(LastPlotCreated)

		LastPlotCreated = treePath+NoForbiddenFilenameCharacters(importName) ' this is for the plot in the 3D tree, will be added below
	Else
		importedObject.Save(GetProjectPath("Result")+NoForbiddenFilenameCharacters(importName)+".sig")
		LastPlotCreated = treePath+NoForbiddenFilenameCharacters(importName)
	End If
	importedObject.AddToTree(LastPlotCreated)

	ImportAs1D = 1 ' All ok

End Function

Function ImportAs1DPolar(xImportColumn As Variant, _
							yImportColumn As Variant, _
							AngleDeg As Boolean, _
							MagdB As Boolean, _
							dBFactor As Double, _
							treePath As String, _
							importName As String, _
							importTitle As String, _
							xLabel As String, _
							yLabel As String, _
							bSortByX As Boolean) As Integer
	' This function takes an array as input, as well as 2 column numbers to determine which columns of the array contain the data
	' The function then imports the array as a 1D polar plot, saves it as importName under treePath and assigns title and labels to the plot
	' Data is assumed to be arranged in columns in importArray
	Dim importedObject As Object, importedObjectDS As Object, importedObject_Amp As Object, importedObject_Phase As Object
	Dim sAmplitudeFile As String, sPhaseFile As String
	Dim i As Long, j As Long
	Dim bUnsorted As Boolean, dTempX As Double, dTempY As Double

	If (UBound(xImportColumn) <> UBound(yImportColumn)) Then
		ImportAs1DPolar = 0 ' Error
		Exit Function
	End If

	Set importedObject = Result1DComplex("")
	Set importedObject_Amp = Result1D("")
	Set importedObject_Phase = Result1D("")

	If MagdB Then
		For i = 0 To UBound(xImportColumn)
			importedObject_Amp.AppendXY(xImportColumn(i), dBFactor ^ (yImportColumn(i) / dBFactor))
			importedObject_Phase.AppendXY(xImportColumn(i), IIf(AngleDeg, xImportColumn(i), xImportColumn(i)*180/Pi))
		Next
	Else
		For i = 0 To UBound(xImportColumn)
			importedObject_Amp.AppendXY(xImportColumn(i), yImportColumn(i))
			importedObject_Phase.AppendXY(xImportColumn(i), IIf(AngleDeg, xImportColumn(i), xImportColumn(i)*180/Pi))
		Next
	End If
	If bSortByX Then
		importedObject_Amp.SortByX()
		importedObject_Phase.SortByX()
	End If

	sAmplitudeFile = GetProjectPath("Result")+NoForbiddenFilenameCharacters(importName)+"_Amp.sig"
	sPhaseFile = GetProjectPath("Result")+NoForbiddenFilenameCharacters(importName)+"_Phase.sig"

	importedObject_Amp.Save(sAmplitudeFile)
	importedObject_Phase.Save(sPhaseFile)
	LastPlotCreated = treePath+NoForbiddenFilenameCharacters(importName)

	importedObject.LoadFromMagnitudeAndPhase(sAmplitudeFile, sPhaseFile)
	importedObject.xlabel(xLabel)
	importedObject.ylabel(yLabel)
	importedObject.title(importTitle)

	' Create a duplicate object in case of "DS for "xxx
	If (Left(GetApplicationName, 7) = "DS for ") Then
		Set importedObjectDS = DS.Result1DComplex("")
		importedObjectDS.LoadFromMagnitudeAndPhase(sAmplitudeFile, sPhaseFile)
		importedObjectDS.xlabel(xLabel)
		importedObjectDS.ylabel(yLabel)
		importedObjectDS.title(importTitle)
	End If

	If (GetApplicationName = "DS") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName)
		importedObject.AddToTree(LastPlotCreated) ' 1DC object can handle multiple subfolders in DS, but truncate to agree with 1D behavior
	ElseIf (Left(GetApplicationName, 7) = "DS for ") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		importedObject.AddToTree(treePath+NoForbiddenFilenameCharacters(importName))
		importedObjectDS.Save(NoForbiddenFilenameCharacters(importName)+"DS.sig")  ' DS requires .sig extension for 1D results
		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName)
		importedObjectDS.AddToTree(LastPlotCreated) ' 1DC object can handle multiple subfolders in DS, but truncate to agree with 1D behavior
	Else
		importedObject.Save(GetProjectPath("Result")+NoForbiddenFilenameCharacters(importName)+".sig")
		LastPlotCreated = treePath+NoForbiddenFilenameCharacters(importName)
		importedObject.AddToTree(LastPlotCreated)
	End If

	ImportAs1DPolar = 1 ' All ok

End Function

Function ImportAs1DC(xImportColumn As Variant, _
						y1ImportColumn As Variant, _
						y2ImportColumn As Variant, _
						ReIm As Boolean, _
						MagdB As Boolean, _
						dBFactor As Double, _
						PhaseDeg As Boolean, _
						treePath As String, _
						importName As String, _
						importTitle As String, _
						xLabel As String, _
						yLabel As String, _
						bSortByX As Boolean) As Integer
	' This function takes an array as input, as well as 3 column numbers to determine which columns of the array contain the data
	' The three columns are x, y1, and y2, where y1 and y2 are either Re/Im or Mag/Ph, according to ReIm
	' PhaseDeg determines if a phase was entered in degrees or rad
	' The function then imports the array as a 1DC plot, saves it as importName under treePath and assigns title and labels to the plot
	' Data is assumed to be arranged in columns in importArray
	Dim importedObject As Object, importedObjectDS As Object
	Dim i As Long, j As Long
	Dim bUnsorted As Boolean, dTempX As Double, dTempY1 As Double, dTempY2 As Double

	If ((UBound(xImportColumn) <> UBound(y1ImportColumn)) Or (UBound(xImportColumn) <> UBound(y2ImportColumn)))Then
		ImportAs1DC = 0 ' Error
		Exit Function
	End If

	Set importedObject = Result1DComplex("")
	If ReIm Then
		For i = 0 To UBound(xImportColumn)
			importedObject.AppendXY(xImportColumn(i), y1ImportColumn(i), y2ImportColumn(i))
		Next
	Else ' MagPhase
		For i = 0 To UBound(xImportColumn)
			importedObject.SetLogarithmicFactor(dBFactor)
			importedObject.AppendXY(xImportColumn(i), _
									IIf(MagdB, 10^(y1ImportColumn(i)/dBFactor), y1ImportColumn(i))*IIf(PhaseDeg, CosD(y2ImportColumn(i)), Cos(y2ImportColumn(i))), _
									IIf(MagdB, 10^(y1ImportColumn(i)/dBFactor), y1ImportColumn(i))*IIf(PhaseDeg, SinD(y2ImportColumn(i)), Sin(y2ImportColumn(i))))
		Next
	End If
	If bSortByX Then
		importedObject.SortByX
	End If

	importedObject.xlabel(xLabel)
	importedObject.ylabel(yLabel)
	importedObject.title(importTitle)

	' Create a duplicate object in case of "DS for "xxx
	If (Left(GetApplicationName, 7) = "DS for ") Then
		Set importedObjectDS = DS.Result1DComplex("")
		If ReIm Then
			For i = 0 To UBound(xImportColumn)
				importedObjectDS.AppendXY(xImportColumn(i), y1ImportColumn(i), y2ImportColumn(i))
			Next
		ElseIf PhaseDeg Then
			For i = 0 To UBound(xImportColumn)
				importedObjectDS.AppendXY(xImportColumn(i), y1ImportColumn(i)*CosD(y2ImportColumn(i)), y1ImportColumn(i)*SinD(y2ImportColumn(i)))
			Next
		Else
			For i = 0 To UBound(xImportColumn)
				importedObjectDS.SetLogarithmicFactor(dBFactor)
				importedObjectDS.AppendXY(xImportColumn(i), _
									IIf(MagdB, 10^(y1ImportColumn(i)/dBFactor), y1ImportColumn(i))*IIf(PhaseDeg, CosD(y2ImportColumn(i)), Cos(y2ImportColumn(i))), _
									IIf(MagdB, 10^(y1ImportColumn(i)/dBFactor), y1ImportColumn(i))*IIf(PhaseDeg, SinD(y2ImportColumn(i)), Sin(y2ImportColumn(i))))
			Next
		End If
		importedObjectDS.xlabel(xLabel)
		importedObjectDS.ylabel(yLabel)
		importedObjectDS.title(importTitle)
	End If

	If (GetApplicationName = "DS") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName)
		importedObject.AddToTree(LastPlotCreated) ' 1DC object can handle multiple subfolders in DS, but truncate to agree with 1D behavior
	ElseIf (Left(GetApplicationName, 7) = "DS for ") Then
		importedObject.Save(NoForbiddenFilenameCharacters(importName)+".sig")  ' DS requires .sig extension for 1D results
		importedObject.AddToTree(treePath+NoForbiddenFilenameCharacters(importName))
		importedObjectDS.Save(NoForbiddenFilenameCharacters(importName)+"DS.sig")  ' DS requires .sig extension for 1D results
		LastPlotCreated = "Design\Results\"+Split(treePath,"\")(UBound(Split(treePath,"\"))-1)+"\"+NoForbiddenFilenameCharacters(importName)
		importedObjectDS.AddToTree(LastPlotCreated) ' 1DC object can handle multiple subfolders in DS, but truncate to agree with 1D behavior
	Else
		importedObject.Save(GetProjectPath("Result")+NoForbiddenFilenameCharacters(importName)+".sig")
		LastPlotCreated = treePath+NoForbiddenFilenameCharacters(importName)
		importedObject.AddToTree(LastPlotCreated)
	End If

	ImportAs1DC = 1 ' All ok

End Function
