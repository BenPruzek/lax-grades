' *Wizard / Compare Multiple Runs
' !!! Do not change the line above !!!
'
' macro allows postprocessing of result-cache or other models, sorted in a tree structure
'
' ================================================================================================
' Copyright 2002-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 06-Jan-2020 ube: improvements of xls output file
' 23-Dec-2019 gba: adapt to new parameter file format
' 17-Nov-2015 ube: subroutine Start renamed to Start_LIB
' 10-Nov-2009 ube: if SQL database exists ("Storage.sdb"), then extract it first and then compare
' 30-Jul-2009 ube: Split replaced by CSTSplit, since otherwise compeating with standrad VBA-Split function
' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (previously only first macropath was searched)
' 18-Oct-2007 ube: converted to 2008
' 21-Oct-2005 imu: Included into Online Help
' 18-Jun-2004 ube: Particle Studio - compliant (.psf)
' 08-Jan-2004 ube: only model-files with existing log-files will be compared
' 26-May-2003 ube: FD-Solver Comparison included (cpu-time)
' 19-May-2003 ube: bugfix for recursive directories  => BaseName() replaced by Left()
' 15-May-2003 ube: MAFIA Batch added + Comparison of 0D+1D Result Postprocessing Templates
' 30-Apr-2003 ube: cst_modfile now without ".mod"-ending due to changes in vba_globals-load-routines
' 10-Apr-2003 ube: do not check log-file any longer (problems with finding all files)
' 01-Jul-2002 ube: first version
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit

Const HelpFileName = "common_preloadedmacro_wizard_compare_multiple_runs"

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "mws_evaluate-results.lib"

Public sFileExt As String
Public iversion As Integer
Public sfilename As String
Public sRootPath As String
Public sTargetPath As String, sTargetExcelFile As String
Public datafile_models As String
Public modfile(999) As String, iNmodfiles As Integer

'-----------------------------------------------------------------------------------------------------------------------------
Sub Main ()

	SpecifyModelFiles

	Dim sMacrobase As String
	Dim targetmodfile As String
	Dim sTitle As String
	
	targetmodfile = sTargetPath + "\" + Right (sTargetPath, 10) + sFileExt
	sTargetExcelFile = Replace(targetmodfile, sFileExt, ".xls")
	
	CST_MkDir (sTargetPath)

	FileCopy GetInstallPath + "\Library\Macros\Wizard\empty.cst", targetmodfile
	
	' evaluate userdefined macros

	Dim iii As Integer, jjj As Integer
	
	If (iNResultMacros > 0) Then
		For iii = 1 To iNmodfiles
			OpenFile sRootPath + "\" + modfile(iii)
			Wait 0.3
				For jjj = 1 To iNResultMacros
					sMacrobase = BaseName(sResultMacro(jjj))
					If Left(sMacrobase,10) = "0D-Result_" Then
						
						' 0D-Result: store 0d-value together with file-index (iii) in txt-file
						
						Open sTargetPath + "\" + sMacrobase + ".txt" For Append As #23
							Print #23,CStr(iii) + "   " + CStr(Evaluate0D(sResultMacro(jjj)))
						Close #23
					Else
						
						' 1D-Result: just run the macro, which creates .sig-files, automatically compared later
						
						Evaluate1D(sResultMacro(jjj))
					End If
				Next jjj
				
			Save
		Next iii
	End If
	
	' compare 1D-Results in tree
	
	OpenFile targetmodfile
	
	Dim cst_modfile As String
	Dim cst_runid As String
	
    Resulttree.EnableTreeUpdate False
	
	For iii = 1 To iNmodfiles
	
		ScreenUpdating False

			cst_modfile = sRootPath + "\" + modfile(iii)
			cst_modfile = Left(cst_modfile, Len(cst_modfile)-4)   ' cut away final 4 characters ".mod" / ".ems"

			cst_runid   = "run #" + CStr(iii)

			' now extract sql database "Storage.sdb" into ascii files

			Dim sql_folder As String
			sql_folder	= cst_modfile + "\Result\"
			UnpackDataToSigFiles (sql_folder)

			LoadProbes      (cst_modfile), (cst_runid)
			LoadVoltages    (cst_modfile), (cst_runid)
			LoadSignalFiles (cst_modfile), (cst_runid)
			LoadEnergies    (cst_modfile), (cst_runid)
			LoadBalances    (cst_modfile), (cst_runid)
			
			Load_1D (cst_modfile), (cst_runid)

			Handle_0D_Results "write", ".m0d", "MAFIA", sTargetPath, cst_modfile, CDbl(iii), cst_runid
			Handle_0D_Results "write", ".rd0", "Template", sTargetPath, cst_modfile, CDbl(iii), cst_runid

		Resulttree.UpdateTree
		ScreenUpdating True
	Next iii

	Handle_0D_Results "plot", ".m0d", "MAFIA", sTargetPath, cst_modfile, CDbl(iii), cst_runid
	Handle_0D_Results "plot", ".rd0", "Template", sTargetPath, cst_modfile, CDbl(iii), cst_runid

    Resulttree.EnableTreeUpdate True
    
	' Add graphs for userdef 0D-Results to the navigation tree 
	' (userdef 1D-Results should always produce .sig-files, which are automaticaly compared)
	
	For jjj = 1 To iNResultMacros
		sMacrobase = BaseName(sResultMacro(jjj))
		
		If Left(sMacrobase,10) = "0D-Result_" Then
		
			' Add graphs to the navigation tree
			sTitle = Replace (Right(sMacrobase,Len(sMacrobase)-10),"_","\", ,1)

			With Resulttree
				.Reset
				.Type "XYSignal"
				.Subtype "magnitude"
				.XLabel "Run ID"
				.YLabel ""
				.Title sTitle
				.File sTargetPath + "\" + sMacrobase + ".txt"
				.Name "0D Results\" + sTitle
				.Add
			End With
			
			SelectTreeItem "0D Results\" + sTitle
			
		End If
	Next jjj

	SelectTreeItem "0D Results\Port-Impedance\Port 1"
	SelectTreeItem "1D Results\Comparison\S-Parameter\|S| linear\S1(1)1(1)"

	WriteExcelFile
	Start_LIB(sTargetExcelFile)
	
	If (InStr(sTargetExcelFile,"(")<>0)Or(InStr(sTargetExcelFile,")")<>0) Then
		MsgBox "Due to round brackets () in the path, automatic start of the excel result sheet might have failed" +vbCrLf +vbCrLf _
		       +" Please open manually: "+sTargetExcelFile
	End If
	

End Sub

'-----------------------------------------------------------------------------------------------------------------------------
Sub WriteExcelFile ()

	Const Nmax_Parameter = 200		' just increase, if not sufficient
	Const sVaryingDefault = "  not used  "
	
	Dim cst_modfile As String, cst_logfile As String, cst_engfile As String
	Dim cst_line As String, cst_runid As String, lib_filename As String, cst_stmp As String
	Dim string_item(30) As String
	Dim cst_nitems As Integer
	Dim cstout_cpumin As Double
	Dim cstout_mesh As Double
	Dim cstout_energy_db As Double

	Dim cst_ParameterName(Nmax_Parameter) As String
	Dim cst_ParameterExpr(Nmax_Parameter) As String 
	Dim cst_VaryingParameterName(Nmax_Parameter) As String
	Dim cst_VaryingParameterExpr(Nmax_Parameter) As String
	
	Dim iNParameter As Integer
	Dim iNParaVary As Integer
	
	Dim sLine As String
	Dim sParName As String
	Dim sParExpr As String
	Dim sParDesc As String
	
	Dim bNewPara As Boolean
	Dim bParaIsUsed As Boolean
	Dim bNewValue As Boolean
	
	Dim iii As Integer
	Dim iip As Integer, i2 As Integer

	Dim names As Variant, expressions As Variant, values As Variant, descriptions As Variant
	Dim nparams As Integer, ipara As Integer
	
	' Loop 1: search all modfiles for new variables and count them
	
	iNParameter = 0
	For iii = 1 To iNmodfiles
		cst_modfile = sRootPath + "\" + modfile(iii)
		cst_runid = modfile(iii)

		nparams = GetProjectParameters(cst_modfile, names, expressions, values, descriptions)
		For ipara = 0 To nparams-1
			sParName = names(ipara)
			sParExpr = expressions(ipara)
			sParDesc = descriptions(ipara)
			bNewPara = True
			For iip = 1 To iNParameter
				If sParName = cst_ParameterName(iip) Then
					bNewPara = False
					Exit For
				End If
			Next iip
			If (bNewPara And sParName<>"") Then
				If (iNParameter < Nmax_Parameter) Then
					iNParameter = iNParameter + 1
					cst_ParameterName(iNParameter) = sParName
					cst_ParameterExpr(iNParameter) = sParExpr
				Else
					MsgBox "Maximum number of variables ( = "+CStr(Nmax_Parameter)+" ) is reached",vbCritical
				End If
			End If
		Next ipara
	Next iii

	' Loop 2: distinguish all parameters into 
	' 		- "fixed": used in all projects with the same value (that value is stored in cst_ParameterExpr(iip))
	'       - "varying": else  (cst_ParameterExpr(iip) set to constant sVaryingDefault)

	iNParaVary = 0
	
	For iip = 1 To iNParameter
		For iii = 1 To iNmodfiles
			cst_modfile = sRootPath + "\" + modfile(iii)

			bParaIsUsed = False
			bNewValue = False
			nparams = GetProjectParameters(cst_modfile, names, expressions, values, descriptions)
			For ipara = 0 To nparams-1
				sParName = names(ipara)
				sParExpr = expressions(ipara)
				If (sParName = cst_ParameterName(iip)) Then
					bParaIsUsed = True
					bNewValue = (sParExpr <> cst_ParameterExpr(iip))
					Exit For
				End If
			Next ipara
			
			If bNewValue Or Not bParaIsUsed Then
				iNParaVary = iNParaVary + 1
				cst_VaryingParameterName(iNParaVary) = cst_ParameterName(iip)
				cst_ParameterExpr(iip) = sVaryingDefault
				Exit For
			End If
		Next iii
	Next iip
	
	Open sTargetExcelFile For Output As #1

	Dim cst_lfirst As Boolean
	Dim cst_parnames_pretty As String
	Dim cst_parvalues_pretty As String
	cst_lfirst = True

	For iii = 1 To iNmodfiles
		cst_modfile = sRootPath + "\" + modfile(iii)
		cst_runid = modfile(iii)
		
		cst_logfile = Replace(cst_modfile,sFileExt,"\Result\Model.log")
		cst_engfile = Replace(cst_modfile,sFileExt,"\Result\1.eng")
		
		For iip = 1 To iNParaVary
			cst_VaryingParameterExpr(iip) = sVaryingDefault
		Next iip
		
		nparams = GetProjectParameters(cst_modfile, names, expressions, values, descriptions)
		For ipara = 0 To nparams-1
			For iip = 1 To iNParaVary
				sParName = names(ipara)
				sParExpr = expressions(ipara)
				If sParName = cst_VaryingParameterName(iip) Then
					cst_VaryingParameterExpr(iip) = sParExpr
					Exit For
				End If
			Next iip
		Next ipara
		
		cst_parnames_pretty = ""
		cst_parvalues_pretty = ""
		For iip = 1 To iNParaVary
			cst_parnames_pretty  = cst_parnames_pretty  + Pretty(cst_VaryingParameterName(iip))
			cst_parvalues_pretty = cst_parvalues_pretty + Pretty(cst_VaryingParameterExpr(iip))
		Next iip

		If (cst_lfirst) Then
			cst_lfirst = False
			Print #1, Pretty("run ID") + cst_parnames_pretty + _
					Pretty("energy/dB") + Pretty("cpu/min") + Pretty("meshcells")	+ Pretty("model-file")
		End If

		On Error GoTo LOGFILE_not_understood

				cstout_cpumin = 0.0
				cst_line = GrepFirstPattern(cst_logfile, "Total Solver Time        :")
				If (cst_line <> "") Then
					' I-Solver special
					i2 = InStr(cst_line,"(=")
					If i2 > 1 Then
						cst_line = Left(cst_line, i2-1)
					End If
				End If
				If (cst_line = "") Then
					cst_line = GrepFirstPattern(cst_logfile, "Run time:")
					If (cst_line <> "") Then
						' TLM-Solver special
						cst_line = Replace(cst_line,"sec"," sec")
					End If
				End If
				If (cst_line = "") Then
					cst_line = GrepFirstPattern(cst_logfile, "Total simulation time:")
				End If
				If (cst_line <> "") Then
					cst_stmp = cst_line
					cst_nitems = CSTSplit(cst_stmp, string_item)
					cstout_cpumin = RealVal(string_item(cst_nitems-2))/60
					cstout_cpumin = DRound(cstout_cpumin,2)
				Else
					cstout_cpumin = 0.0
				End If

				cst_line = GrepFirstPattern(cst_logfile, "Number of mesh cells")
				If (cst_line = "") Then
					cst_line = GrepFirstPattern(cst_logfile, "Number of cells")
				End If
				If (cst_line = "") Then
					cst_line = GrepFirstPattern(cst_logfile, "Number of meshnodes")
				End If
				If (cst_line = "") Then
					cst_line = GrepFirstPattern(cst_logfile, "Number of meshcells")
				End If
				If (cst_line = "") Then
					cstout_mesh = 0.0
				Else
					cst_stmp = cst_line
					cst_nitems = CSTSplit (cst_stmp, string_item)
					cstout_mesh = RealVal(string_item(cst_nitems-1))
				End If

				If Dir$(cst_engfile) = "" Then
					cstout_energy_db = 0.0
				Else
					cst_line = GetLastLine(cst_engfile)
					If (cst_line = "") Then
						cstout_energy_db = 0.0
					Else
						cst_stmp = cst_line
						cst_nitems = CSTSplit (cst_stmp, string_item)
						cstout_energy_db = RealVal(string_item(1))
						cstout_energy_db = DRound(cstout_energy_db,1)
					End If
				End If

				Print #1, Pretty("run #"+CStr(iii)) + cst_parvalues_pretty + _
						Pretty(cstout_energy_db) + Pretty(cstout_cpumin) + Pretty(cstout_mesh) + cst_runid
				
				GoTo INFO_is_written
						
		LOGFILE_not_understood:

				Print #1, Pretty("run #"+CStr(iii)) + cst_parvalues_pretty + Pretty("0.0")  + Pretty("0.0")  + Pretty("0.0") + cst_runid
		
		INFO_is_written:
				On Error GoTo 0
				
	Next iii

	If (iNParaVary < iNParameter) Then
		' at the end write fixed parameter values
		Print #1, " "
		Print #1, " Fixed parameters (identical in all compared files) "
		Print #1, " "
	
		For iip = 1 To iNParameter
			If cst_ParameterExpr(iip) <> sVaryingDefault Then
				' fixed parameters, used by all projects with identical values / expressions
				Print #1, Pretty(cst_ParameterName(iip)) + Pretty(cst_ParameterExpr(iip))
			End If
		Next iip
	End If

Close #1

End Sub

'-----------------------------------------------------------------------------------------------------------------------------
Sub SpecifyModelFiles ()

	sFileExt = ".cst"

	Begin Dialog UserDialog 520,245,"Postprocess multiple Simulations",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,500,140,"Specify models, to be compared",.GroupBox1
		Text 30,35,120,14,"Root-Directory",.Text
		TextBox 30,56,460,21,.FileName
		OKButton 30,217,100,21
		PushButton 150,28,90,21,"Browse...",.Browse
		CancelButton 150,217,100,21
		CheckBox 40,84,340,14,"Recursive Search through ALL subdirectories",.CheckRecursive
		Text 30,112,140,14,"File Search Pattern",.Text1
		TextBox 170,112,210,21,.Pattern
		GroupBox 10,154,500,49,"Specify empty Result Folder for comparison",.GroupBox2
		TextBox 30,175,460,21,.ResultFile
		PushButton 270,217,100,21,"Help",.Help
	End Dialog
    Dim dlg As UserDialog

    Do
		If Dir$(GetProjectPath("Result") + "Cache",vbDirectory) <> "" Then
            dlg.FileName = GetProjectPath("Result") + "Cache"
        Else
            dlg.FileName = GetProjectPath("Root")
        End If
        
        dlg.CheckRecursive = 1
        dlg.Pattern = "*" + sFileExt

        iversion  = 1
        sfilename = FindFirstFile(GetProjectPath("Root"), ShortName(GetProjectPath("Project"))+ "_compare_##", False)
        While (sfilename <> "")
            iversion  = CInt(Right$(ShortName(sfilename), 2)) + 1
            sfilename = FindNextFile
        Wend
        dlg.ResultFile = GetProjectPath("Project") + "_compare" + Format(iversion, "\_00")

        If (Dialog(dlg) >= 0) Then Exit All
    Loop Until (dlg.FileName <> "")

    sRootPath = dlg.FileName
    sTargetPath = dlg.ResultFile
    
    CST_MkDir sTargetPath

	Dim lNFiles As Long

	datafile_models        = sTargetPath + "\list_of_modelfiles.txt"

    Open datafile_models For Output As #1
    	lNFiles = 0
    
	    sfilename = FindFirstFile(sRootPath, dlg.Pattern, IIf(dlg.CheckRecursive = 1, True, False))

	    While (sfilename <> "")

			On Error GoTo No_LOG_File_No_Results_SkipThisFile

				GetAttr(sRootPath + "\" + Replace(sfilename, sFileExt, "\Result\Model.log"))	' See if log-file exists; don't use Dir() !!
				'
				' log-file exists -> results are there
				'
	            Print #1, sfilename
	            lNFiles = lNFiles + 1

			No_LOG_File_No_Results_SkipThisFile:

			On Error GoTo 0

		        sfilename = FindNextFile
	    Wend

	Close #1

	Begin Dialog UserDialog 570,126,"List of models",.DialogFunc3 ' %GRID:10,7,1,1
		Text 30,14,330,14,"Searching Files Finished:",.Text1
		Text 30,42,340,14,"Number of Result-Files, to be compared        = " + CStr(lNFiles),.Text2
		CheckBox 30,70,260,14,"Show Functions - VBA_Userdefined",.CheckVBA
		PushButton 170,98,120,21,"Continue",.Delete
		PushButton 30,98,120,21,"View/Edit List",.View
		CancelButton 310,98,100,21
		PushButton 430,98,100,21,"Help",.Help
	End Dialog
    Dim dlg3 As UserDialog
    If (Dialog(dlg3) = 0) Then Exit All

	If dlg3.CheckVBA = 1 Then
		SpecifyUserdefinedMacros
	End If

	Dim sDUMMY As String
	Open datafile_models For Input As #1
    	iNmodfiles = 0
		While Not EOF(1)
			Line Input #1, sDUMMY
			iNmodfiles = iNmodfiles + 1
			modfile(iNmodfiles) = sDUMMY
		Wend
	Close #1

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

        Dim filename As String, Extension As String, Index As Integer

        Select Case Action
                Case 1 ' Dialog box initialization
                        Extension = "zip"
                Case 2 ' Value changing or button pressed
                        Select Case Item
					 			Case "Help"
										StartHelp HelpFileName
										DialogFunc = True
                               Case "Browse"
                                        Extension = ""
                                        filename = DlgText("FileName") + "\" + "Use this directory"
                                        filename = GetFilePath(filename, "", "", "Choose Root-directory", 2)
                                        If (filename <> "") Then
                                                DlgText "FileName", DirName(filename)
												iversion  = 1
												' double dirname, because of "Use this directory"
												sfilename = FindFirstFile(DirName(DirName(filename)), ShortName(DirName(filename))+ "_compare_##", False)
												While (sfilename <> "")
												        iversion  = CInt(Right$(ShortName(sfilename), 2)) + 1
												        sfilename = FindNextFile
												Wend
                                                DlgText "ResultFile", DirName(filename) + "_compare" + Format(iversion, "\_00")
                                        End If
                                        DialogFunc = True
                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select
End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc3%(Item As String, Action As Integer, Value As Integer)

        Dim filename As String, Extension As String, Index As Integer
        Select Case Action
                Case 1 ' Dialog box initialization
                Case 2 ' Value changing or button pressed
                        Select Case Item
							Case "Help"
								StartHelp HelpFileName
								DialogFunc3 = True
                            Case "View"
                                Shell("notepad.exe " + datafile_models, 1)
                                DialogFunc3 = True
                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select
End Function
