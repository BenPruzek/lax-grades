' Launch Optenni Lab

' This function will test if Optenni Lab is installed on the system, get the installation path
' of the latest version (using an auxiliary program and the registry), store the calculated impedance as a touchstone file and
' launch Optenni Lab. The touchstone file and the generated circuits are stored under the project in
' directory Results/OptenniLab

'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 17-Oct-2022 jra: Export of radiation patterns now ask for the angular resolution of the patterns
' 17-Oct-2022 jra: Extracted the routines which can be used from Microwave Studio from OptenniLibrary.lib to optenniLibrary_MWS.lib
' 14-Oct-2022 jra: Correct touchstone file name obtained from the Touchstone object
' 09-Feb-2021 jra: Reorganized common code to OptenniLibrary.lib
' 09-Feb-2021 jra: Enabling Linux functionality
' 22-Jan-2021 jra: fixed a problem with invalid AC efficiency results
' 03-Jun-2020 jra: added port geometry transfer and iso-latin-1 encoding
' 22-Aug-2019 jra: added XML verbatim code for file names
' 31-May-2019 jra: using the project name as the name of the S parameter data file
' 14-Apr-2018 jra: added support for sending radiation pattern data
' 28-Dec-2016 ube: ar-filter case adjusted to complex s-parameters
' 16-Nov-2016 jra: Fixed a problem with copied efficiency data
' 09-Oct-2015 jra: improved component placement on DS schematic
' 28-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 23-Apr-2012 ube: if AR-Filter results exist, those will be now used and exported to OptenniLab
' 14-Jun-2011 jra: Fixed a problem with existing empty OptenniLab directory
' 30-May-2011 jra: Added correct window titles in error messages
' 27-May-2011 jra: Fixed problem with spaces in directory or project file names
' 20-May-2011 jra: First version
'-----------------------------------------------------------------------------------------------------------------------------
'#Include "OptenniLibrary.lib"
'#Include "OptenniLibrary_MWS.lib"

Option Explicit

Sub Main () 

   If Not Resulttree.DoesTreeItemExist("1D Results\S-Parameters") Then
      MsgBox("No impedance results have been calculated for the project. Please simulate the EM model.",vbOkOnly, _
             "No results calculated")
      Exit Sub
   End If

   ' Generate an OptenniLab directory under the Result folder of the project
   Dim OptenniDir As String
   OptenniDir = GetProjectPath("Result") +"OptenniLab"
   If Dir(OptenniDir,vbDirectory) = "" Then
      MkDir OptenniDir
   End If
   Dim projectFullName As String
   Dim projectName As String
   projectFullName = GetProjectPath("Project")
   Dim slashPos As Integer
   slashPos = InStrRev(projectFullName, "\")
   If slashPos = 0 Then
      MsgBox ("Could not get the project name.",vbOkOnly, _
              "Project name not found")
      Exit Sub
   End If
   projectName = Mid(projectFullName, slashPos+1)

   Dim fname As String

   'Store the S parameters in a touchstone file
   With TOUCHSTONE
      .Reset
      .FileName ("OptenniLab\"+projectName)
      .Impedance (50)
      .FrequencyRange ("Full")
      .Renormalize (True)
      If Result1DDataExists("ar^cS1(1)1(1)") Then
         .UseARResults (True)
      Else
         .UseARResults (False)
      End If
      .Write
	  fname = .getFileName
   End With

   If GetOptenniPorts() = 0 Then
      MsgBox ("No ports found in the structure.",vbOkOnly,"No ports found")
      Exit Sub
   End If

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

   Dim efficiencyCommand As String
   efficiencyCommand = GetOptenniEfficiencyCommand()
   Dim patternCommand As String
   If (major > 4 Or (major = 4 And minor >1)) Then
      If Len(efficiencyCommand) > 0 Then
         Dim retval As Integer
         retval = MsgBox ("Do you want to transfer radiation pattern data to Optenni Lab?", _
                          vbYesNoCancel, "Send radiation pattern data")
         If retval = vbCancel Then
            Exit Sub
         End If
         If retval = vbYes Then
			Dim resolutionStr As String
			resolutionStr = InputBox("Enter angular resolution for the exported radiation pattern data (in degrees)", _
			"Enter angular resolution", "5")
			If Len(resolutionStr) = 0 Then
				Exit Sub
			End If
			
			Dim resolution As Double
			resolution = ParseOptenniResolution(resolutionStr)
			If resolution <= 0 Then
				Exit Sub
			End If
            patternCommand = GetOptenniRadiationPatternCommand(resolution)
         End If 
      End If
   End If
   If Len(patternCommand) > 0 Then
      efficiencyCommand = ""
   End If
   Dim portGeometryCommand As String
   If (major >= 5) Then
      portGeometryCommand = StoreOptenniPortGeometryFile()
   End If
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
	
   fname = """" + GetProjectPath("Result")+fname + """"
   If Not IsWindows Then
      fname = Replace$(fname, "\", "/")
   End If
   ReportInformationToWindow("Starting Optenni Lab")
   ' Launch Optenni Lab and pass the generated touchstone file name
   If (major < 2 Or (major = 2 And minor <2)) Then
      Shell(optenniPath + " -i "+fname,vbNormalFocus)
   Else
      Shell(optenniPath +" " + fname+ CSTIdCommand+ efficiencyCommand _
            +patternCommand + portGeometryCommand + projectPathStr ,vbNormalFocus)
   End If
   Exit Sub
   		
End Sub


