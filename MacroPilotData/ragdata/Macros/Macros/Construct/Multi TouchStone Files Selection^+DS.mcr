' This macro Enable multiple selection in the input window of the files selection section for TOUCHSTONE Block, used within CST DS.
'
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 06-Jun-2024 amd: Dropdown List removed and Selection of Block done beforehand from the schematic
' 13-Feb-2024 amd: First Version
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit

Dim nPins As Integer
Dim sorting As String
Dim counter As Integer
Dim BlockName As String


Private Sub CheckSelectedBlock()
	On Error GoTo ErrorMsg
	Dim bType As String
	Dim nPins As Integer

	BlockName =  Mid(GetSelectedTreeItem(), InStrRev(GetSelectedTreeItem(), "\")+1)
	With Block
		.Reset
		.Name(BlockName)
		 bType = .GetTypeShortName
	End With
	If bType <> "TS"  Then
			Error
	End If
	Exit Sub

    ErrorMsg:
		MsgBox ("Please Select a TouchStone Block from the Schematic", vbCritical, "Add Multiple TOUCHSTONE Files")
		Exit All


End Sub


Function NumberExtractor(ByVal v As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Set regex = CreateObject("VBScript.RegExp")
	If sorting = "Numerically" Then

		' Regular expression to find up to 4 digits preceding ".s" and followed by any characters
    	regex.Pattern = "(\d{1,8})\.s"
    	regex.Global = True
    	Set matches = regex.Execute(v)

	    Dim numericPart As String
    	If matches.Count > 0 Then
        	' Extracts the digits (up to 4)
        	numericPart = matches(0).SubMatches(0)
        	' No need for formatting with leading zeros here since we're handling up to 4 digits
        	NumberExtractor = CLng(numericPart)
    	Else
        	' If no match is found, return an empty string or a specific value indicating no match
        	NumberExtractor = 0000 ' Default or error value if no numeric pattern is found
    	End If

    Else
    	NumberExtractor = v
	End If


End Function


Private Function FilesToList(projectdir As String,BlockName As String) As String()
	Dim file3 As String, filename_ext As String, filename As String
    Dim extension As String
    Dim fs As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim list() As String
    Dim count As Integer
	With Block
		.Reset
		.Name(BlockName)
		nPins = .GetNumberOfPins
	End With
    extension = "s" + CStr(nPins) +"p"

    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    ' Get the folder object
    Set objFolder = fs.GetFolder(projectdir)

    ' Loop through each file in the folder
    For Each objFile In objFolder.Files
        ' Check if the file has the desired extension
        If fs.GetExtensionName(objFile.Path) = extension Then
            ' Process the file
            ' Increase the count
			count = count+1
            ReDim Preserve list(1 To count)
			list(count)= objFile
        End If

    Next objFile
    FilesToList = list
    counter = count

End Function
Sub RemoveTSFile(nFiles, BlockName As String)
	Dim i As Integer
	For i=1 To nFiles-1
		With Block
			.Reset
			.Name(BlockName)
			.RemoveFile(1)
		End With
	Next i
End Sub
Sub AddListToTS(MyList()As String, BlockName As String) 'Add the TOUCHSTONE files from the sorted list to the TOUCHSTONE Block
	Dim j As Integer
	Dim FileName As String

	With Block
		.Reset()
		.Name(BlockName)
		FileName = .GetFilePath(1)
		For j = LBound(MyList) To UBound(MyList)
			If Mid(MyList(j), InStrRev(MyList(j), "\")+1) = Mid(FileName, InStrRev(FileName, "\")+1) Then
				If j > 1 Then
					.RemoveFile(1)
					.AddFile(MyList(j))
				End If

			Else
				With Block
					.Reset()
					.Name(BlockName)
					.AddFile(MyList(j))
				End With
			End If

		Next j
	End With
End Sub


Sub QuickSort(arr() As String, low As Long, high As Long)
    Dim pivot As Variant
    Dim temp As Variant
    Dim i As Long
    Dim j As Long

    If low < high Then
        pivot = NumberExtractor(arr((low + high) / 2))
        i = low
        j = high

        Do
            While NumberExtractor(arr(i)) < pivot
                i = i + 1
            Wend

            While NumberExtractor(arr(j)) > pivot
                j = j - 1
            Wend

            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop While i <= j

        QuickSort arr, low, j
        QuickSort arr, i, high
    End If
End Sub
Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed'
     If DlgItem = "OKButton" Then
		Dim fs As Object
		Dim FileDir As String
		Dim list() As String
		Set fs = CreateObject("Scripting.FileSystemObject")
		FileDir =DlgText("filename")
		If IsEmpty(FileDir) Or fs.FolderExists(FileDir) = False Then ' Check if file directory text box is empty or the given path is exsisting
			MsgBox(FileDir+" incorrect path, please enter a correct path", vbCritical, "Add Multiple TOUCHSTONE Files")
			DialogFunc = True
		Else
			FilesToList(FileDir,BlockName) 'Check if the path contains TOUCHSTONE files
			If counter < 1 Then
				MsgBox ("The entered path doesn't contain any TOUCHSTONE files.", vbCritical, "Add Multiple TOUCHSTONE Files")
				DialogFunc = True
			End If

		End If
     End If
     If DlgItem = "CancelButton" Then
	 	Exit All
	 End If
	 If DlgItem = "HelpButtonPushed" Then
		StartDESHelp "macro\common_macro_multi_TSfiles_selection"
		Exit All
	 End If
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Sub Main
	Dim list(1)
	list(0)= "Alphabetically"
	list(1) = "Numerically"
	CheckSelectedBlock()
	' show dialog
	Begin Dialog UserDialog 520,101, "Import TOUCHSTONE files to schematic", .DialogFunc
       	Text 10,15,270,15,"File Directory", .Text4
       	TextBox 100,15,405,15, .filename
       	Text 10,40,170,15,"Sorting Type"
     	DropListBox 100,40,100,15,list(),.DropListBox2
   	 	PushButton 30,65,60,20, "&OK", .OKButton
       	PushButton 100,65,60,20, "&Cancel", .CancelButton
       	PushButton 170,65,60,20, "&Help",.HelpButtonPushed

	End Dialog

	Dim dlg As UserDialog

	If Dialog(dlg)=0 Then
	Exit All
	End If

	Dim Path As String
	Dim nFiles	As Integer
	Path = dlg.filename
	sorting = list(dlg.DropListBox2)


	With Block
		.Reset
		.Name(BlockName)
		nFiles = .GetNumberOfFiles
	End With
	'Clearing Old Files from TS
	RemoveTSFile (nFiles, BlockName)

	'Storing all the s*p Files in List
	Dim MyList() As String
	MyList = FilesToList(Path,BlockName)
	'Ordering the List Alphabetically
    QuickSort (MyList, LBound(MyList), UBound(MyList))

	'add Orderd List Files To TS Block
	AddListToTS(MyList, BlockName)

	'Clean up first TS file
    With Block
    	.Reset()
    	.Name(BlockName)
    	nFiles = .GetNumberOfFiles
    	If nFiles> UBound(MyList)Then
    		RemoveTSFile(2,BlockName)
    	End If
    End With
End Sub
