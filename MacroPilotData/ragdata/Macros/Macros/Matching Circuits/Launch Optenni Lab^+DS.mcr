' Launch Optenni Lab

' This function will test if Optenni Lab is installed on the system, get the installation path
' of the latest version (using an auxiliary program and the registry), store the calculated impedance as a touchstone file and
' launch Optenni Lab. The touchstone file and the generated circuits are stored under the project in
' directory Results/OptenniLab

'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 09-Feb-2021 jra: Reorganized common code to OptenniLibrary.lib
' 09-Feb-2021 jra: Enabling Linux functionality
' 31-May-2019 jra: using the project name as the name of the S parameter data file
' 09-Oct-2015 jra: improved component placement on DS schematic
' 08-Dec-2011 ube: added macropath to runandwait statement, since in some cases optenni did not start properly.
' 28-Jul-2011 ube: added support for DS
' 14-Jun-2011 jra: Fixed a problem with existing empty OptenniLab directory
' 30-May-2011 jra: Added correct window titles in error messages
' 27-May-2011 jra: Fixed problem with spaces in directory or project file names
' 20-May-2011 jra: First version
'-----------------------------------------------------------------------------------------------------------------------------
Option Explicit
'#include "vba_globals_ds.lib"
'#Include "OptenniLibrary.lib"

Sub Main ()

   FillTaskNameArray
   FillPortNameArray

   Dim nport As Integer

   nport = UBound(PortNameArray)
   If nport = 0 Then
      MsgBox ("No ports found in the structure.",vbOkOnly,"No ports found")
      Exit Sub
   End If

   If UBound(TaskNameArray) = 0 Then
      MsgBox ("No S-Parameter tasks defined.",vbOkOnly,"No results found")
      Exit Sub
   End If

   Begin Dialog UserDialog 300,105,"Select S-Parameter Task" ' %GRID:10,7,1,1
      GroupBox 20,7,260,56,"Simulation Task",.GroupBox1
      DropListBox 40,28,220,192,TaskNameArray(),.Task
      OKButton 30,70,90,21
      CancelButton 130,70,90,21
   End Dialog
   Dim dlg As UserDialog
   dlg.Task       = 0
   If dlg.Task = -1 Then dlg.task = 0

   If (Dialog(dlg) = 0) Then Exit All

   ' Generate an OptenniLab directory under the Result folder of the project
   Dim OptenniDir As String
   OptenniDir = GetProjectPath("Result") +"OptenniLab"
   If Dir(OptenniDir,vbDirectory) = "" Then
      MkDir OptenniDir
   End If
   Dim sBase As String
   Dim projectFullName As String
   Dim projectName As String
   projectFullName = GetProjectPath("Project")
   Dim slashPos As Integer
   slashPos = InStrRev(projectFullName, "\")
   If slashPos = 0 Then
      MsgBox ("Could not get the project name.",vbOkOnly,"Project name not found")
      Exit Sub
   End If
   projectName = Mid(projectFullName, slashPos+1)
   
   sBase = OptenniDir+"\" + projectName

   Dim sTreePath As String
   sTreePath = "Tasks\"+TaskNameArray(dlg.task)+"\S-Parameters"

   On Error GoTo NOSPARA
   DS.TouchstoneExport(sTreePath, sBase, "50")
   ' Arguments:
   ' 1) Tree-Pfad
   ' 2) Basename of TOUCHSTONE-file (ending .sNp will be added automatically)
   ' 3) Reference impedance
   GoTo sNpSuccess
NOSPARA:
   MsgBox ("No S-Parameters found in Task "+TaskNameArray(dlg.task), _
           vbOkOnly,"No results found")
   Exit Sub

sNpSuccess:
   On Error GoTo 0

   Dim fname As String


   call InitOptenniLibrary()
   Dim optenniPath As String
   optenniPath = GetOptenniLabPath()

   If optenniPath = "" Or optenniPath = "OptenniLabNotInstalled" Then
      MsgBox ("Optenni Lab is a matching circuit generation and impedance analysis software developed by Optenni Ltd."+ _
              vbCrLf +"An Optenni Lab installation was not found."+ _
              vbCrLf +"Please contact Optenni Ltd at www.optenni.com for more information", _
              vbOkOnly ,"Optenni Lab not installed")
      Exit Sub
   End If

   Dim major As Integer
   Dim minor As Integer
   major = GetOptenniLabMajorVersion()
   minor = GetOptenniLabMinorVersion()

   ' Check Optenni Lab version. It should be at least 1.2 (1.3 for the reverse link)

   If (major < 1 Or (major = 1 And minor <2)) Then
      MsgBox ("The Optenni Lab version must be at least 1.2." + vbCrLf + _
              "The current version is " +CStr(major) + "." + CStr(minor) + _
              ".", vbOkOnly, "Incompatible version")
      Exit Sub
   End If

   fname = """"+OptenniDir +"\" + projectName + ".s" + CStr(nport) + "p"+""""
   If Not IsWindows Then
      fname = Replace$(fname, "\", "/")
   End If
   ReportInformationToWindow("Starting Optenni Lab")
   ' Launch Optenni Lab and pass the generated touchstone file name

   Dim CSTIdCommand As String
   Dim CSTid As String
   CSTid = DS.GetRegisteredDEString()
   If CSTid <> "" Then
      CSTIdCommand = " -cstinstance "+CSTid
   End If

   Dim projectPathStr As String
   If (major > 5 Or (major = 5 And minor >= 2)) Then
      projectPathStr = " -EMProjectPath """ + ProjectFullName + ".cst"""
   End If

   If (major < 2 Or (major = 2 And minor <2)) Then
      Shell(optenniPath + " -i "+fname,vbNormalFocus)
   Else
      Shell(optenniPath + " "+fname+ CSTIdCommand + projectPathStr _
            ,vbNormalFocus)
   End If


End Sub

