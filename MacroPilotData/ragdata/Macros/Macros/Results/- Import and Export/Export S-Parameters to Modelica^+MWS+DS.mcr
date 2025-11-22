'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2020-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------------
' 23-Sep-2020 ube: enable IdEM
' 13-Aug-2020 ube,cks: first version
'------------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Dim BlockNameArray() As String, strBlockName As String
Dim sFolder As String, sFile As String

Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean

	strBlockName = BlockNameArray(DlgValue("nBlockSelected"))

	Dim bHasFrequencyBounds As Boolean
	With Block
		.name(strBlockName)
		bHasFrequencyBounds = (.GetProperty("Lower Frq Bound") = "True") And (.GetProperty("Upper Frq Bound") = "True")
	End With

	sFile   = GetProjectPath("Project")
	sFile   = Mid(sFile, 1+InStrRev(sFile,"\")) + "_" + strBlockName + ".mo"

	If (Action% = 1 ) Then
		' Allow to switch on IDEM choice
		'DlgEnable ("Group1", False)
		If UBound(BlockNameArray)=0 Then
			DlgEnable ("nBlockSelected", False)
			If bHasFrequencyBounds Then
				DlgEnable ("OptionButton3", True)
				DlgValue ("Group2", 0)
				DlgEnable ("fmax", False)
			Else
				DlgEnable ("OptionButton3", False)
				DlgEnable ("fmax", True)
				DlgValue ("Group2", 1)
			End If
		End If
		sFolder = GetProjectPath("Root")
		DlgText("Folder", sFolder)
		DlgText("Filename", sFile)
	End If

	If (Action% = 1 Or Action% = 2) Then

		If (DlgItem = "nBlockSelected") Then
			DlgText("Filename", sFile)
			If bHasFrequencyBounds Then
				DlgEnable ("OptionButton3", True)
				DlgValue ("Group2", 0)
				DlgEnable ("fmax", False)
			Else
				DlgEnable ("OptionButton3", False)
				DlgEnable ("fmax", True)
				DlgValue ("Group2", 1)
			End If
		End If

		If (DlgItem = "Group2") Then
			If DlgValue("Group2") = 0 Then
				DlgEnable ("fmax", False)
			Else
				DlgEnable ("fmax", True)
			End If
		End If

		If (DlgItem = "BrowseFolder") Then
			sFolder = DlgText("Folder")
			sFolder = GetFolder_Lib(sFolder)
			If sFolder <> "" Then
				If Right(sFolder,1) = "\" Then
					sFolder = Left(sFolder, Len(sFolder)-1)
				End If
				DlgText("Folder", sFolder)
			End If
			DialogFunc = True
		End If

		If (DlgItem = "OK") Then
			DialogFunc = False
			If 1 = 0 Then
				MsgBox "Only even numbers are accepted for number of elements.", vbInformation ,"Number of Elements"
				DialogFunc = True
			End If
		End If
	End If

End Function

Sub Main ()

	Dim nBlocks As Long, nIndex As Long, nCount As Long

	nBlocks = Block.StartBlockNameIteration()
	ReDim BlockNameArray(nBlocks)

	nCount = -1
	With Block
		.Enable3DCommands(False)
		For nIndex = 0 To nBlocks-1
			strBlockName = Block.GetNextBlockName()
			.Name strBlockName
			If (.CanUseMOR) Then
				nCount = nCount + 1
				BlockNameArray(nCount) = strBlockName
			End If
		Next nIndex
	End With
	If nCount >= 0  Then
		ReDim Preserve BlockNameArray(nCount)
	Else
		MsgBox "No Block found, for which vector fitting can be applied.", vbExclamation
		Exit All
	End If

	Begin Dialog UserDialog 490,415,"Export S-Parameters to Modelica",.DialogFunc ' %GRID:10,5,1,1
		GroupBox 20,150,450,55,"Vector Fitting Method",.GroupBox1
		GroupBox 20,10,450,130,"Block",.GroupBox2
		OKButton 30,385,90,20
		CancelButton 130,385,90,20
		DropListBox 70,53,380,20,BlockNameArray(),.nBlockSelected
		OptionGroup .Group1
			OptionButton 90,175,120,15,"Built-in",.OptionButton1
			OptionButton 240,175,130,15,"IdEM",.OptionButton2
		Text 70,30,250,15,"Select Block to be exported:",.Text2
		TextBox 220,110,230,20,.fmax
		GroupBox 20,215,450,160,"Output",.GroupBox3
		TextBox 40,260,410,20,.Folder
		PushButton 340,235,110,20,"Browse...",.BrowseFolder
		Text 40,240,90,15,"Folder:",.Text3
		Text 40,290,90,15,"Filename:",.Text4
		TextBox 40,310,410,20,.Filename
		CheckBox 40,350,410,15,"Keep existing Modelica icon, only overwrite statespace model",.KeepModelicaIcon
		OptionGroup .Group2
			OptionButton 90,85,270,15,"Block Frequency Bounds",.OptionButton3
			OptionButton 90,112,120,15,"Manual Fmax:",.OptionButton4
	End Dialog

	Dim dlg As UserDialog

	dlg.KeepModelicaIcon = 0

	If (Dialog(dlg) = 0) Then Exit All

	strBlockName = BlockNameArray(dlg.nBlockSelected)

	Dim sFmax As String
	If dlg.Group2 = 0 Then
		sFmax = ""
	Else
		sFmax = dlg.fmax
	End If

	With Block
		.name(strBlockName)
		If dlg.Group1 = 0 Then
			.SetVectorFittingMethod ("BUILT-IN")
		Else
			.SetVectorFittingMethod ("IDEM")
		End If

		' ExtractStateSpaceModel( String fmax , String task )
		.ExtractStateSpaceModel( sFmax , "" )
	End With

	sFolder = dlg.Folder
	sFile = dlg.Filename

	Dim sFolderABCD As String
	sFolderABCD = GetProjectPath("Project")+"\Export\ABCD"

	Dim sPythonScript As String, sCompleteShellCommand
	Dim sPythonExe As String
	sPythonExe = GetInstallPath + "\AMD64\python\python.exe"
	sPythonScript = GetInstallPath + "\Library\Macros\Results\- Import and Export\WriteModelicaFile.py"

	sCompleteShellCommand = Quote(sPythonExe) + " " + Quote(sPythonScript) + " " + Quote(sFile) + " " + Quote(sFolder) + " " + Quote(sFolderABCD) + " " + Cstr(dlg.KeepModelicaIcon)

	Shell(sCompleteShellCommand)

	' the batch file below is not needed, since Shell command above already did the execution
	' nevertheless batch file might be useful to leave, just in case something went wrong on a special machine

	Dim batchfile As String
	batchfile = sFolderABCD + "\batch-modelica-convert.bat"
	Open batchfile For Output As #1
		Print #1, sCompleteShellCommand
	Close #1

	' RunAndWait(Quote(batchfile))

	MsgBox(".mo file successfully generated: "+vbCrLf+vbCrLf+ sFolder +vbCrLf+sFile ,vbInformation,"Modelica file generated")

End Sub
