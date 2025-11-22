' ================================================================================================
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------
' 02-Jan-2020 ube: deactivate macro to remove it 1-2 years later (leave msgbox to direct users to the new place)
' 03-Jul-2017 jpe: adding support for multiple partitions of terra-2
' 10-Nov-2015 ube: additional output for cst-internal workflow
' 28-Aug-2015 ube: default year in the dialogue (dlg.Year) is now initialized as CSTversion year minus 1
' 03-Sep-2014 shi: test for directorybit in FoundDirectory
' 07-Jul-2014 ube: special cst-internal option (b_UBE_terra)
' 18-Dec-2013 ube,dta: removed checkbox to keep project folder (now folder always deleted, action a+d)
' 22-Feb-2011 ube: some more error handling added
' 24-Jan-2011 ube: year can now be entered as free text
' 16-Nov-2009 ube: added year 2010, default date is 1-sep-09
' 07-Aug-2008 ube: reverted change of 25-July since VBA engines are globally corrected to be backward compatible
' 25-Jul-2008 ube: new VBA engine returns different index for droplistbox, if array(0) is undefined
' 19-Jul-2007 ube: new project management tool used
'------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_wizard_delete_old_results"

Const b_UBE_terra = False   ' False / True
Const b_UBE_MACRO_ACTIVE = False   ' False / True

'#include "vba_globals_all.lib"
'------------------------------------------------------------------------------
Public projectdir As String
Public tempdir As String
Public ilevel As Long
Public DDay As Long
Public dirlast$(50)
Const NN$="Never been in this directory"
Public FilesDeleted As Long
Public datafile As String
Public hidden_dirs As String
Public cstfilename As String, keepAllResults As Boolean, keep1DResults As Boolean, keepFarfieldData As Boolean, deleteProjFolder As Boolean

'-----------------------------------------------------------------------------------------------------------------------------

Sub Main ()

	If Not b_UBE_MACRO_ACTIVE Then
		MsgBox "This macro has been deactivated, since there is new modern archiving technology available under ""File -> Preview and Organize"". " + vbCrLf  + vbCrLf + "In the Online Help additional information can be found here: "+vbCrLf+"""General Features -> Preview And Organize Projects""",vbExclamation,"Archive Multiple Projects"
		Exit All
	End If

	Dim cst_month$(12)
	cst_month$(1)        = "Jan"
	cst_month$(2)        = "Feb"
	cst_month$(3)        = "Mar"
	cst_month$(4)        = "Apr"
	cst_month$(5)        = "May"
	cst_month$(6)        = "Jun"
	cst_month$(7)        = "Jul"
	cst_month$(8)        = "Aug"
	cst_month$(9)        = "Sep"
	cst_month$(10)       = "Oct"
	cst_month$(11)       = "Nov"
	cst_month$(12)       = "Dec"

	Dim filename As String, SubdirName As String
	Dim iii As Long
	Dim rootdir As String
	Dim workdir As String

	projectdir     = GetProjectPath("Root")

	Begin Dialog UserDialog 520,266,"Archive / Delete old Results",.DialogFunc ' %GRID:5,3,1,1
		GroupBox 10,140,500,56,"",.GroupBox3
		GroupBox 10,0,500,141,"",.GroupBox1
		Text 30,21,120,14,"Root-Directory",.Text
		TextBox 30,42,460,21,.FileName
		PushButton 150,14,90,21,"Browse...",.Browse
		Text 30,77,300,14,"Handle only CST projects, last modified before",.LabelOptions
		OKButton 20,238,100,21
		CancelButton 130,238,100,21
		TextBox 330,72,30,21,.Day
		DropListBox 368,72,62,192,cst_month(),.DropListMonth
		TextBox 439,72,49,21,.Year
		OptionGroup .Group1
			OptionButton 40,98,280,14,"Handle only Root-directory",.OptionButton1
			OptionButton 40,119,350,14,"Handle Root and ALL subdirectories recursively",.OptionButton2
'		PushButton 270,245,100,21,"Help",.Help
		CheckBox 40,154,400,14,"Archive Results in CST file",.CheckArchiveResults
		CheckBox 60,175,130,14,"Keep all Results",.CheckAllResults
		CheckBox 200,175,140,14,"Keep 1D Results",.Check1DResults
		CheckBox 360,175,140,14,"Keep Farfields",.CheckFarfields
		GroupBox 10,195,500,36,"",.GroupBox2
		Text 30,210,450,15,"Note: Project Folders will be deleted after archiving.",.Text1
		'CheckBox 40,210,320,14,"Delete Project Folder after archiving",.DeleteProjFolder
	End Dialog
	Dim dlg As UserDialog

	Do
		dlg.FileName    = projectdir

		dlg.Day = "01"
		dlg.DropListMonth = 0

		' old hardwired dlg.Year = "2014"
		' new always take cstversionyear minus 1

		Dim syear As String
		syear = Cstr(-1+Cint(Mid$(GetApplicationVersion,9,4)))
		If Left(syear,2)<>"20" Then
			ReportError "Year incorrectly initialized, Exit all"
			Exit All
		End If
		dlg.Year = syear

		dlg.Group1 = 1

		dlg.CheckArchiveResults = 1
		dlg.CheckAllResults = 0
		dlg.Check1DResults = 1
		dlg.CheckFarfields = 0
		'dlg.DeleteProjFolder = 1

		If (Dialog(dlg) >= 0) Then Exit All
	Loop Until (dlg.FileName <> "") And (CLng(dlg.Day)<32)

	If (dlg.CheckArchiveResults) Then
		keepAllResults = dlg.CheckAllResults
		keep1DResults = dlg.Check1DResults
		keepFarfieldData = dlg.CheckFarfields
	Else
		keepAllResults = False
		keep1DResults = False
		keepFarfieldData = False
	End If

	deleteProjFolder = 1 ' dlg.DeleteProjFolder

	DDay = CLng(dlg.Day) + 100*(1+CLng(dlg.DropListMonth)) + 10000*CLng(dlg.Year)
	Debug.Print DDay

	Dim heute As Long
	heute = CLng(Day(Date)) + 100 * CLng(Month(Date)) + 10000 * CLng(Year(Date))

	' DEBUG	MsgBox "Die unterste Zeile muss heute sein (20000608):"+vbCrLf+vbCrLf+Str(Date)+vbCrLf+Str(heute)

	If (DDay > heute-100) Then
		Begin Dialog UserDialog 390,147, "Confirmation", .DialogFunc1 ' %GRID:10,7,1,1
			GroupBox 10,7,370,98,"",.GroupBox1
			Text 110,21,150,14,"A T T E N T I O N",.Text1
			Text 20,42,340,14,"You are going to search Projects, younger than 1 month!",.Text2
			Text 20,63,350,14,rootdir,.Text3
			Text 20,84,320,14,"files older than "+CStr(DDay)+" will be deleted",.Text5
			CancelButton 10,119,120,21
			PushButton 140,119,140,21,"I am absolutely sure",.PushButton1
			PushButton 290,119,100,21,"Help",.Help
		End Dialog
		Dim dlg2 As UserDialog
		If (Dialog(dlg2) = 0) Then Exit All
	End If

	Dim checkOnlyRoot As Long
	If dlg.Group1=0 Then
		checkOnlyRoot        = 1
	Else
		checkOnlyRoot        = 0

		Begin Dialog UserDialog 390,168 ' %GRID:10,7,1,1
			GroupBox 10,7,370,126,"",.GroupBox1
			Text 110,21,150,14,"A T T E N T I O N",.Text1
			Text 20,42,290,14,"You are going to delete RECURSIVELY !",.Text2
			Text 20,63,350,14,rootdir,.Text3
			Text 20,84,330,14,"and ALL its subdirectories are studied.",.Text4
			Text 20,112,320,14,"files older than "+CStr(DDay)+" will be deleted",.Text5
			CancelButton 20,140,80,21
			PushButton 110,140,140,21,"I am absolutely sure",.PushButton1
			PushButton 260,140,100,21,"Help",.Help
		End Dialog
		Dim dlg1 As UserDialog
		If (Dialog(dlg1) = 0) Then Exit All
	End If

	tempdir = IIf(Environ$("TEMP") <> "", Environ$("TEMP"), "C:")
	datafile = tempdir + "\CST_list_of_deleted_files.txt"
	Open datafile For Output As #1

	hidden_dirs = tempdir + "\CST_hidden_dirs.txt"
	Open hidden_dirs For Output As #3
	Close #3

	Dim scst As String

	FilesDeleted        = 0

	' =================
	Dim sa() As String, root2 As String, partWarning As String
	If b_UBE_terra Then

	Begin Dialog UserDialog 390,161,"Partition selection",.DialogFunc4 ' %GRID:10,7,1,1
		GroupBox 10,7,370,119,"",.GroupBox1
		Text 30,21,310,28,"Choose which partitions of terra-2 shall be searched/cleaned.",.Text1
		CancelButton 200,133,90,21
		OptionGroup .Partition
			OptionButton 30,56,310,21,"Clean BOTH partitions"
			OptionButton 30,77,310,21,"Clean only the D: partition (Germany+Other)"
			OptionButton 30,98,340,21,"Clean only the E: partition (all besides Germany)"
		OKButton 100,133,90,21
	End Dialog
		Dim dlg4 As UserDialog
		If (Dialog(dlg4) = 0) Then Exit All

			root2 = "\\terra-2\work\"
			
			' old FillArray sa() ,  Array("_Central","France","Germany+Other","India","Korea","NorthAmerica","SouthEastAsia","TempWorkers","UK","SouthAmerica","Italy")
			' just-for-testing    root2 = "D:\Dokumente\scripting\terra-2\test-root\"
			
			If (dlg4.Partition = 0) Then
				' this includes all folders from the network share
				FillArray sa() ,  Array("_Central","France","Germany+Other","India","Korea","NorthAmerica","SouthEastAsia","TempWorkers","UK","SouthAmerica","Italy")
				partWarning = "both partitions"
			ElseIf (dlg4.Partition = 1) Then
				' this includes only the folder actually residing on D: (the other folders are just symlinks to E:\... )
				' root2 may be set to \\terra-2\d$, if no symlinks are created on D:\ and administrative shares can be used
				FillArray sa() ,  Array("Germany+Other")
				partWarning = "the D: partition"
			ElseIf (dlg4.Partition = 2) Then
				' this includes all folders that reside on the E:\ partition
				' root2 could be set with \\terra-2\e$ if administrative shares can be used
				FillArray sa() ,  Array("_Central","France","India","Korea","NorthAmerica","SouthEastAsia","TempWorkers","UK","SouthAmerica","Italy")
				partWarning = "the E: partition"
			End If

		Begin Dialog UserDialog 280,119 ' %GRID:10,7,1,1
			GroupBox 20,7,240,49,"",.GroupBox1
			Text 40,21,220,28,"Are you sure that you want to clean "+partWarning+"?",.Text1
			CancelButton 20,91,240,21
			OKButton 20,63,240,21
		End Dialog
		Dim dlg5 As UserDialog
		If (Dialog(dlg5) = 0) Then Exit All


	Else
		root2 = dlg.FileName
		FillArray sa() ,  Array("")
	End If

	Dim s2 As String

	For Each s2 In sa
		rootdir = root2 + s2
		' MsgBox rootdir

		For iii=0 To 20
			dirlast$(iii)        = NN$
		Next

		If b_UBE_terra Then
			ReportInformation rootdir + " : " + CStr(FilesDeleted)
		End If

		ChDir rootdir

		Dim lsearch As Boolean
		lsearch        = True
		ilevel        = 0

		' ====  main loop
		Do
			' go into new Working directory

			workdir = rootdir
			For iii=0 To ilevel-1
				workdir = workdir + "\" + dirlast$(iii)
			Next
			
			Debug.Print CStr(ilevel)+"   "+CStr(lsearch)+"   "+workdir
			On Error GoTo NO_ACCESS_TO_DIRECTORY
			ChDir workdir
			On Error GoTo 0
			GoTo DIRECTORY_IS_ENTERED
			
			NO_ACCESS_TO_DIRECTORY:
				MsgBox "no access: " + workdir
				Open hidden_dirs For Append As #3
					Print #3, workdir
				Close #3
				ilevel = ilevel-1
				GoTo DIRECTORY_IS_DONE

			DIRECTORY_IS_ENTERED:
				If lsearch Then
					' NOTHING DONE HERE      ArchiveAndDelete_Files ("*.cst")
				End If
                
			DIRECTORY_IS_DONE:
				If checkOnlyRoot=1 Then
					' never check subdirectories
					ilevel        = -10

					' to be absolutely sure
					Exit Do
				Else
					If NextSubdirFound(SubdirName) Then
						scst = workdir+"\"+SubdirName+".cst"
						If Dir$(scst)<>"" Then
							' stay on same level, that directory is the project folder belonging to a cst-file
							If OlderThanDDay(scst) Then
								Print #1, "a+d " + scst
								FilesDeleted = FilesDeleted + 1
							End If

							lsearch          = False
							dirlast$(ilevel) = SubdirName
							' ilevel unchanged
						Else
							' go down one level
							lsearch          = True
							dirlast$(ilevel) = SubdirName
							ilevel           = ilevel+1
						End If
					Else
						' go up one level
						lsearch              = False
						dirlast$(ilevel)     = NN$
						ilevel               = ilevel-1
					End If
				End If
'                MsgBox CStr(ilevel)

		Loop Until ilevel < 0

	Next s2

	Close #1

	Begin Dialog UserDialog 370,98,"Summary and Execution",.DialogFunc3 ' %GRID:10,7,1,1
		GroupBox 10,7,350,56,"",.GroupBox1
		Text 30,21,310,14,"Searching Files Finished:",.Text1
		Text 20,42,330,14,"Number of CST-Files found:  " + CStr(FilesDeleted),.Text2
		PushButton 20,70,120,21,"View/Edit List",.View
		PushButton 150,70,100,21,"Execute List",.Delete
		CancelButton 260,70,80,21
	End Dialog
	Dim dlg3 As UserDialog
	If (Dialog(dlg3) = 0) Then Exit All

End Sub

Function OlderThanDDay (F$) As Boolean

	Dim dDate As Date
	Dim mfiledate As Long

	dDate		= FileDateTime(F$)
	mfiledate	= CLng(Day(dDate)) + 100 * CLng(Month(dDate)) + 10000 * CLng(Year(dDate))

	' DEBUG	MsgBox "here you can check the filedate "+vbCrLf+vbCrLf+"Filename: "+F$+vbCrLf+str1+vbCrLf+Str(mfiledate)

	OlderThanDDay = mfiledate < DDay

End Function

Function NextSubdirFound (subdir As String) As Boolean

	Dim F$

	F$ = Dir$("*",vbDirectory)
	While Not FoundDirectory(F$)
			F$ = Dir$()
	Wend

	If (F$="") Then
		NextSubdirFound = False
		subdir = "NIX gefunden"
	Else
		If (dirlast$(ilevel)=NN$) Then
			NextSubdirFound        = True
			subdir                        = F$
		Else
			While (F$<>"") And (F$ <> dirlast$(ilevel))
				F$ = Dir$()
				While Not FoundDirectory(F$)
					F$ = Dir$()
				Wend
			Wend

			If (F$="") Then
				NextSubdirFound = False
				subdir = "NIX gefunden"
			Else
				F$ = Dir$()
				While Not FoundDirectory(F$)
					F$ = Dir$()
				Wend

				If (F$="") Then
					NextSubdirFound        = False
					subdir                        = "NIX gefunden"
				Else
					NextSubdirFound        = True
					subdir                        = F$
				End If
			End If
		End If
	End If

End Function

'-----------------------------------------------------------------------------------------------------------------------------
Function FoundDirectory (F$) As Boolean

	If (F$="") Then
		FoundDirectory        = True
	Else
		On Error GoTo Skip_22
			If (F$="." Or F$=".." Or Not(CBool(GetAttr(F$) And vbDirectory))) Then
				FoundDirectory        = False
			Else
				FoundDirectory        = True
			End If
			GoTo Normal_22
		Skip_22: FoundDirectory = False
		Normal_22: On Error GoTo 0
	End If

End Function

Sub ArchiveAndDelete_Files (pattern As String)

	Dim F$
	F$ = Dir$(pattern,vbNormal)
	While (F$<>"")
		If OlderThanDDay(F$) Then
			Print #1, "a+d " + CurDir$() + "\" + F$
			FilesDeleted        = FilesDeleted + 1
		End If
		F$ = Dir$()
	Wend

End Sub

Function DialogFunc%(Item As String, action As Integer, Value As Integer)

	Dim filename As String, Index As Integer

	If (action%=1) Or (action%=2) Then

		Dim bResults As Boolean
		bResults = (DlgValue("CheckArchiveResults") = 1)

		Dim bAllRes As Boolean
		bAllRes = (DlgValue("CheckAllResults") = 1)

		DlgEnable "CheckAllResults", bResults
		DlgEnable "Check1DResults", bResults And Not bAllRes
		DlgEnable "CheckFarfields", bResults And Not bAllRes

        Select Case Item
			Case "Help"
				StartHelp HelpFileName
				DialogFunc = True
			Case "Browse"
				'filename = DlgText("FileName") + "\" + "Use this directory"
				'filename = GetFilePath(filename, "", "", "Choose Root-directory", 2)
				filename = GetFolder_Lib(DlgText("FileName"))
				If (filename <> "") Then
			        DlgText "FileName", DirName(filename)
				End If
				DialogFunc = True
			Case "OK"
				If Not IsNumeric(DlgText("Year")) Then
					MsgBox "Please check entered Year."
					DialogFunc = True
				Else
					If Clng(DlgText("Year")) > CLng(Year(Date)) Then
						MsgBox "Year is in future."
						DialogFunc = True
					End If
					If Clng(DlgText("Year")) < 2000 Then
						MsgBox "Year is too old."
						DialogFunc = True
					End If
				End If
				If Not IsNumeric(DlgText("Day")) Then
					MsgBox "Please check entered day."
					DialogFunc = True
				Else
					If Clng(DlgText("Day")) < 1 Or Clng(DlgText("Day")) > 31 Then
						MsgBox "Day must be between 1 and 31."
						DialogFunc = True
					End If
				End If
        End Select
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc3%(Item As String, action As Integer, Value As Integer)
	Select Case action
		Case 1 ' Dialog box initialization
		Case 2 ' Value changing or button pressed
			Select Case Item
				Case "Help"
					StartHelp HelpFileName
					DialogFunc3 = True
				Case "View"
					Shell("notepad.exe " + datafile, 1)
					DialogFunc3 = True
				Case "Delete"
					DeleteList
					DialogFunc3 = False
			End Select
		Case 3 ' ComboBox or TextBox Value changed
		Case 4 ' Focus changed
		Case 5 ' Idle
	End Select
End Function
'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc1%(Item As String, action As Integer, Value As Integer)
	Select Case action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case Item
		Case "Help"
			StartHelp HelpFileName
			DialogFunc1 = True
		End Select
	Case 3 ' ComboBox or TextBox Value changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------------

Sub DeleteList

	Dim SLine$,action$
	Dim archivefile As String
	Dim ilen As Long

	FilesDeleted = 0

	Dim nProblems As Long, sfile_problems As String, s1 As String

	nProblems = 0
	sfile_problems = tempdir + "\CST_access_problems_maybe_READONLY.txt"
	Open sfile_problems For Output As #3
	Close #3

	Open datafile For Input As #2
		While Not EOF(2)
			Line Input #2,SLine$
			Debug.Print SLine$
			action$ = Left(SLine$,3)
			ilen = Len(SLine$)
			cstfilename  = Right(SLine$, ilen-4)

			Debug.Print "##" + cstfilename + "##"

			On Error GoTo Problem_Delete

				Select Case action$
				Case "a+d"
					StoreInArchive(cstfilename, keepAllResults, keep1DResults, keepFarfieldData, deleteProjFolder)
					FilesDeleted = FilesDeleted + 1
				Case Else
					MsgBox "unknown action:" + vbCrLf + SLine + vbCrLf + "Exit all.", vbCritical
					Exit All
				End Select
			On Error GoTo 0
			GoTo Problem_Done

			Problem_Delete:
				Open sfile_problems For Append As #3
					Print #3, SLine$
					nProblems = nProblems + 1
				Close #3
				On Error GoTo 0
			Problem_Done:
		Wend
	Close #2

	MsgBox  "Archiving and Deleting Results Finished:" + vbCrLf + vbCrLf _
				 + "Projects (archived now)  = " + CStr(FilesDeleted) + vbCrLf + vbCrLf _
				 + "Projects with Problems   = " + CStr(nProblems)

	If (nProblems>0) Then
		Shell("notepad.exe " + sfile_problems, 1)
	End If

End Sub

Rem See DialogFunc help topic for more information.
Function DialogFunc4%(Item As String, action As Integer, Value As Integer)
	Select Case action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc4 = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc4 = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
