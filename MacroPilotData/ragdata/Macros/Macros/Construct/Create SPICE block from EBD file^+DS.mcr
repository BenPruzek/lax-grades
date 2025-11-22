'#Language "WWB-COM"

'This Macro read the .edb file and adds in the parasitic elements and import those as SPICE block.

' ================================================================================================
' Copyright 2016-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'====================
' 20-Feb-2019 ytn: fixed problem where per unit length capacitance are connected in series, they should be connected to circuit ground. Pin list display, redundant fork-endfork pair bug fix.
' 31-Oct-2016 yta: initial version based on IBIS Version 6.1.

Option Explicit
'#include "vba_globals_all.lib"

Dim iDisplayArray() As Integer, PinList() As String, PinList1() As String,  sSelectedPinList() As String, sSelectedPinListString As String, FormattedBlockString As String
Dim MaxLineNum As Integer, NumLine As Long, PathName() As String, filename_ext As String, filename As String, sSelectedPathName As String, nstr As Integer, cst_line_array() As String
Dim PinName() As String, length() As String, l() As String, c() As String, r() As String, fork() As Boolean, endfork() As Boolean
Dim SpiceString As String, BlockString As String, SpicePath As String



Sub Main

	NumLine = 0
	'This block read the formatted ebd file line by line and assign pin, len, r, l, c and fork, endfork arrarys to be used in writespice function
	Dim file3 As  String, textline3 As String
	Dim projectdir, extension As String
	projectdir = GetProjectPath("Project")
	extension = "ebd"
	file3 = GetFilePath("", extension, projectdir, "Browse for EBD file", 0)
	filename_ext = Mid(file3, InStrRev(file3, "\")+1)
	filename = Left(filename_ext, Len(filename_ext)-4)
	If file3 = "" Then
		Exit All
	End If

	BlockString = ReadEBDToString(file3)

	FormattedBlockString = FormatEBDString(BlockString)

	Dim NumDisplayPins As Integer
	NumDisplayPins = FillDisplayPinName(FormattedBlockString)

	Dim l As Integer
	ReDim PinList1(NumDisplayPins)
	For l = 0 To NumDisplayPins-1
		PinList1(l) = PinList(l)
	Next


	Begin Dialog UserDialog 320,294,"Select Pins From EBD File",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,300,273,"Pin List from: " + filename_ext,.GroupBox1
		MultiListBox 30,42,260,196,PinList1(),.MultiListBox
		OKButton 50,252,90,21
		CancelButton 170,252,100,21
		Text 20,21,280,14,"Note: Capacitance placed at beginning.",.Text1												   
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg	'displays dialog, waiting for ok


	With Block
		.Name filename_ext
		If .DoesExist Then
			.Delete
		End If
	End With

	With Block
		.Reset
		.Type("SPICEImport")
		.name(filename_ext)
		.Position(51050, 51000)
		.SetFile(SpicePath)
		.SetRelativePath("false")
		.Create
	End With

	Dim pinindex As Integer, edge As String
	Dim edgeindex As Integer
	With Block
		.name(filename_ext)
		.SetName(filename_ext)

		Dim i As Integer
		For i = 0 To .GetNumberOfPins-1
			.SetPinLayout(i, "TOP", i+1)
		Next
	End With

	'Layout of the block
	With Block
		.reset
		.name filename_ext
	End With


End Sub


Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	Select Case Action%
		Case 1 ' Dialog box initialization
			DlgListBoxArray "MultiListBox", PinList1	'displays the list of all available pins
			iDisplayArray = DlgValue("MultiListBox")

		Case 2 ' Value changing or button pressed
			If DlgItem$ = "MultiListBox" Then
				iDisplayArray = DlgValue("MultiListBox")
				ReDim Preserve PinList1(0 To UBound(iDisplayArray))
				DialogFunc = True 'do not exit the dialog
			ElseIf DlgItem$ = "OK" Then
				'Extract the pin names of the selected entries according to their indeces
				ReDim sSelectedPinList(UBound(iDisplayArray))
				Dim i As Integer
				sSelectedPinListString = ""
				sSelectedPathName = ""
				For i = LBound(iDisplayArray) To UBound(iDisplayArray)
					sSelectedPinList(i) = PinList(iDisplayArray(i))
					sSelectedPinListString = sSelectedPinListString + "_" + sSelectedPinList(i) + "_"
					PathName(i) = PathName(iDisplayArray(i))
					sSelectedPathName = sSelectedPathName + "_" + PathName(i) + "_"
				Next
				GetListFromDialog
			ElseIf DlgItem$ = "Cancel" Then
				Exit All
			End If

		Case 4 ' Focus changed
	End Select

End Function


Function ReadEBDToString(filestring As String) As String
	Dim ifile1 As Long, textline As String
	Dim BlockString As String
	BlockString = ""
	'textline = ""
	ifile1 = FreeFile

	Open filestring For Input As #ifile1
	Do Until InStr(textline, "[Begin Board Description]")
		Line Input #ifile1, textline
	Loop

	Do Until InStr(textline, "[End Board Description]")
		Line Input #ifile1, textline
		If Left(Trim(textline), 19) = "[Path Description] " Then
			BlockString = BlockString + textline + vbCrLf
			Do Until textline = "|"
				Line Input #ifile1, textline
						BlockString = BlockString + textline + vbCrLf
			Loop
		End If
	Loop

	Close #ifile1
	ReadEBDToString = BlockString	'return the blockstring of the file to be formatted later.

End Function


Function FormatEBDString(sString As String) As String
	Dim lines() As String, i As Integer, l As String, ResultString As String
	lines = Split(sString, vbCrLf)
	For i = 0 To UBound(lines) - 1
		lines(i) = Replace(lines(i), Chr(9), "")
		lines(i) = Trim(lines(i))
		If Len(lines(i)) > 0 Then
			If	Left(lines(i), 1) = "|"  Then
				lines(i) = "|"
			End If
			'Assumes file have either " R = 50" or "R=50" format, but not "R= 50"
			lines(i) = Split(lines(i), "| ")(0)	
			lines(i) = Replace(lines(i), " = ", "=")
			lines(i) = Replace(lines(i), " /", "")
			ResultString = ResultString + lines(i) + vbCrLf
		End If
	Next
	FormatEBDString = ResultString

End Function


Function FillDisplayPinName(sString As String) As Integer

	Dim lines() As String, i As Integer, j As Integer, k As Integer
	j = 0
	lines = Split(sString, vbCrLf)
	ReDim PinList(UBound(lines))
	ReDim PathName(UBound(lines))
	For i = 1 To UBound(lines)
		If InStr(lines(i), "Pin") <> 0 Then
			PathName(j) = Mid(lines(i-1), 20)
			PinList(j) = Mid(lines(i), 5)
			j = j+1
		End If
	Next
	FillDisplayPinName = j

End Function


Sub GetListFromDialog

	Dim LinesArray() As String
	LinesArray = Split(FormattedBlockString, vbCrLf)
	MaxLineNum = UBound(LinesArray)
	ReDim length(MaxLineNum)
	ReDim l(MaxLineNum)
	ReDim r(MaxLineNum)
	ReDim c(MaxLineNum)
	ReDim PinName(MaxLineNum)
	ReDim fork(MaxLineNum)
	ReDim endfork(MaxLineNum)
	'Assumes that file is properly written, i.e there's no "pin", 'node', 'fork', 'endfork', 'len' etc.
	For NumLine = 0 To UBound(LinesArray)-1
		If Left(LinesArray(NumLine), 19) = "[Path Description] " And InStr(sSelectedPathName, "_" + Mid(LinesArray(NumLine), 20) + "_") Then
			Do Until LinesArray(NumLine) = "|"
				If Left(LinesArray(NumLine), 4) =  "Pin " Then
					PinName(NumLine) = Mid(LinesArray(NumLine), 5)
				ElseIf Left(LinesArray(NumLine), 5) = "Node " Then
					PinName(NumLine) = Mid(LinesArray(NumLine), 6)
				ElseIf Left(LinesArray(NumLine), 4) = "Fork" Then
					fork(NumLine) = True
				ElseIf Left(LinesArray(NumLine), 7) = "Endfork" Then
					endfork(NumLine) = True	
				ElseIf Left(LinesArray(NumLine), 4) = "Len=" Then
					ReDim cst_line_array(4)	
					nstr = CSTSplit(LinesArray(NumLine), cst_line_array)
					ConvertUnits(cst_line_array)
					length(NumLine) = Mid(cst_line_array(0), 5)
					ReadRLCValues(cst_line_array, r, l, c, NumLine)	
				End If
				NumLine = NumLine + 1
			Loop
		End If
	Next

	GetEquivalentRLC(r, l, c, length)	
	SpicePath = WriteSpice(r, l, c, length, PinName, fork, endfork, filename)

End Sub

Sub ConvertUnits(cst_line_array() As String)

	Dim i As Integer
	Dim nstr As Integer
	nstr = UBound(cst_line_array)
	For i = 1 To nstr
		cst_line_array(i) = Replace(cst_line_array(i), "m", "e-3")
		cst_line_array(i) = Replace(cst_line_array(i), "u", "e-6")
		cst_line_array(i) = Replace(cst_line_array(i), "n", "e-9")
		cst_line_array(i) = Replace(cst_line_array(i), "p", "e-12")
	Next i

End Sub


Sub ReadRLCValues(cst_line_array() As String, r() As String, l() As String, c() As String, NumLine As Integer)

	Dim i As Integer
	Dim nstr As Integer
	nstr = UBound(cst_line_array)

	For i = 1 To nstr-1	
		If Left(UCase(cst_line_array(i)), 2) = "R=" Then
			r(NumLine) = Mid(cst_line_array(i), 3)
		ElseIf Left(UCase(cst_line_array(i)), 2) = "L=" Then
			l(NumLine) = Mid(cst_line_array(i), 3)
		ElseIf Left(UCase(cst_line_array(i)), 2) = "C=" Then
			c(NumLine) = Mid(cst_line_array(i), 3)
		End If
	Next

End Sub


Function WriteSpice(r() As String, l() As String, c() As String, length() As String, PinName() As String, fork() As Boolean, endfork() As Boolean, PathName As String) As String
	Dim i As Integer, NumLine As Integer, SpiceString As String, ifile As Integer, FilePath As String, NonZeroPins() As String, NonZeroPinCount As Integer, EndPin As String, StartPin As String
	NumLine = UBound(r)	
	ReDim NonZeroPins(NumLine)
	Dim LastEndPin As String
	Dim ForkEndPins() As String
	Dim ForkDepth As Integer
	ForkDepth = 0
	ReDim ForkEndPins(UBound(endfork))
	ifile = FreeFile
	SpiceString = ""

	FilePath = GetProjectPath("Project") + "\Result\" + PathName + ".net"
	Open FilePath For Output As #ifile		'open file to write

	For i = 0 To NumLine - 1		
		If PinName(i) <> "" Then	
			NonZeroPins(NonZeroPinCount) = PinName(i)
			NonZeroPinCount = NonZeroPinCount + 1
		End If

		If length(i) <> "" Then	
			Dim EndNode1 As String, EndNode2 As String, EndNode3 As String
			EndNode1 = GetNextNode()
			EndNode2 = GetNextNode()
			EndNode3 = GetNextNode()
			If PinName(i+1) <> "" Then
				EndPin = PinName(i+1)
			ElseIf fork(i+1) Then
				Dim endforkmatchingindex As Integer
				endforkmatchingindex = FindMatchingEndFork(fork, endfork, i+1)
				If PinName (endforkmatchingindex + 1) <> "" Then
					EndPin = PinName(endforkmatchingindex + 1)
				ElseIf PinName(i+2) <> "" Then
					EndPin = PinName(i+2)
				Else
					EndPin = EndNode3	
				End If
			Else
				EndPin = EndNode3
			End If

			If PinName(i-1) <> "" Then		
				StartPin = PinName(i-1)
			ElseIf endfork(i-1) Then	
				ForkDepth = ForkDepth - 1
				StartPin = ForkEndPins(ForkDepth)
			ElseIf fork(i-1) And endfork(i-2) Then
					StartPin = ForkEndPins(ForkDepth-2)
			Else
				StartPin = LastEndPin
			End If

			If 		r(i) <> "" And l(i) <> "" And c(i) <> "" Then 'RLC
					SpiceString = SpiceString + "C " + StartPin + " " + "0" + " " + c(i) + vbCrLf
					SpiceString = SpiceString + "L " + StartPin + " " + EndNode1 + " " + l(i) + vbCrLf
					SpiceString = SpiceString + "R " + EndNode1 + " " + EndPin + " " + r(i) + vbCrLf

			ElseIf 	r(i) = "" And l(i) <> "" And c(i) <> ""	Then'LC
					SpiceString = SpiceString + "C " + StartPin + " " + "0" + " " + c(i) + vbCrLf
					SpiceString = SpiceString + "L " + StartPin + " " + EndPin + " " + l(i) + vbCrLf

			ElseIf 	r(i) = "" And l(i) <> "" And c(i) = ""	Then'L
					SpiceString = SpiceString + "L " + StartPin + " " + EndPin + " " + l(i) + vbCrLf

			End If
					LastEndPin = EndPin
					SpiceString = SpiceString + vbCrLf
		ElseIf fork(i) Then	
			ForkEndPins(ForkDepth) = LastEndPin
			ForkDepth = ForkDepth + 1
		End If
	Next
		ReDim Preserve NonZeroPins(NonZeroPinCount)
		Dim NewPinString As String
		For i = 0 To NonZeroPinCount -1
			NewPinString = NewPinString + NonZeroPins(i) + " "
		Next
		Print #ifile, _
				"#Generate SPICE file " + vbCrLf + vbCrLf + vbCrLf + _
				".subckt " + PathName + " " + NewPinString + vbCrLf + vbCrLf + _
				SpiceString + vbCrLf + ".ends"
	Close #ifile
	WriteSpice = FilePath

End Function


'This function generates the next node to be used when fork is detected
Function GetNextNode() As String
	Static i As Integer
	i = i + 1
	GetNextNode = "z_" + cstr(i)
End Function

'This functino calculates and stores the effective value for each rlc value by multiplying the non zero length to the rlc entries.
Sub GetEquivalentRLC(r() As String, l() As String, c() As String, length() As String)
	Dim i As Integer
	For i = 0 To UBound(length)-1
		If length(i) <> "" And length(i) <> "0" Then
			If r(i) <> "" Then
				r(i) = cstr(CDbl(r(i))*CDbl(length(i)))
			End If
			If l(i) <> "" Then
				l(i) = cstr(CDbl(l(i))*CDbl(length(i)))
			End If
			If c(i) <> "" Then
				c(i) = cstr(CDbl(c(i))*CDbl(length(i)))
			End If
		End If
	Next
End Sub

Function FindMatchingEndFork(fork() As Boolean, endfork() As Boolean, forkindex As Integer)
	Dim i As Integer, ForkDepth As Integer
	For i = forkindex To UBound(fork)-1
		If fork(i) Then
			forkdepth = forkdepth + 1
		ElseIf endfork(i) Then
			forkdepth = forkdepth - 1
		End If

		If forkdepth = 0 Then
			FindMatchingEndFork = i
			Exit Function
		End If
	Next
	FindMatchingEndFork = -1
End Function
