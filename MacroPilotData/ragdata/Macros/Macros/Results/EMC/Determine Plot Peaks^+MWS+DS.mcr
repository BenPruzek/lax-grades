' Option Explicit not possible, since tricky constructions below...

'#include "vba_globals_all.lib"
'#include "template_conversions.lib"
'#include "template_results.lib"

' Determine E-field Peaks: also Displays Markers on Peaks and Saves the Values to a File
' Works well if the Peaks are not too sharp (i.e. shaped like an impulse response)
'--------------------------------------------------------------------------------------------
' Copyright 2015-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 02-Jan-2020 ube: added online help page
' 24-Oct-2015 ube: subroutine Start renamed to Start22
' 03-Mar-2015 fsr: added drop down list to select result curve
' 20-Feb-2015 aba: added selectron for unit type in the box above/below
' 18-Feb-2015 ube: dialogue cosmetics and added into release
' 17-Feb-2015 mle: added functionality to limit peak search to an amplitude range
' 10-Feb-2015 aba,mle: changed the path to wirte text file, so it will be visible in the tree. Added name of the plot into the result file
' 10-Feb-2015 fsr: added functionality to select proper item also in design studio
' 09-Feb-2015 aba: added info box that tells the user to select a 1d result, fixed Frequency prefix output, changed peak output in file to dBuV
' 02-Feb-2015 mle: Fixed macro excitation error
' 29-Jan-2015 mle: include a frequency range and change GUI and max peak to min peak
' 28-Jan-2015 mle: first version
'--------------------------------------------------------------------------------------------

Dim result As Object
Dim aResultName() As String
Dim aResultType() As String

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

	Select Case Action
		Case 1 ' Dialog box initialization
		Case 2 ' Value changing or button pressed
			Select Case Item
				Case "Help"
					StartHelp "common_preloadedmacro_emc_determine_plot_peaks"
					DialogFunc = True
			End Select
		Case 3 ' ComboBox or TextBox Value changed
		Case 4 ' Focus changed
		Case 5 ' Idle
	End Select
End Function
Sub Main ()

	Dim sFile As String
	Dim no_peaks As Boolean              ' Flag used to indicate whether there are any peaks found
	Dim icount As Long

	Dim TreeObjectToCall As Object, sTypeFromItemNameMethodToCall As String
	Dim sSelectedTreeItem As String

	Dim MyUnitArray(3) As String

	MyUnitArray(0) = "dB"
	MyUnitArray(1) = "dBm"
	MyUnitArray(2) = "dBu"

	FillResultList_LIB(aResultName, aResultType, sName, "ALL", "ALL", "+TEMPLATE+TREE", sListOfSelectionSettings, sListOfSelectionTypeSettings)
	icount = UBound(aResultName)

	Begin Dialog UserDialog 750,210,"Peak Search",.DialogFunc ' %GRID:10,7,1,1
		DropListBox 10,35,730,175,aResultName(),.ResultListDLB
		PushButton 440,182,90,21,"Apply",.PushButton1
		CancelButton 540,182,90,21
		GroupBox 250,63,210,112,"Frequency Range",.GroupBox2
		GroupBox 10,63,240,112,"Find Peaks",.GroupBox4
		GroupBox 460,63,280,112,"Peak Properties",.GroupBox3
		TextBox 110,112,120,21,.ampupp
		TextBox 110,84,120,21,.amplow
		TextBox 320,84,120,21,.fmin
		TextBox 320,112,120,21,.fmax
		TextBox 600,84,120,21,.peakht
		TextBox 600,112,120,21,.peaknum
		Text 270,91,40,14,"Start",.Text3
		Text 270,119,40,14,"Stop",.Text5
		Text 480,91,110,14,"Peak Height (dB)",.Text4
		Text 480,119,110,14,"Max Number",.Text6
		Text 30,119,70,14,"Below",.Text1
		Text 30,91,60,14,"Above",.Text2
		Text 30,147,70,14,"Unit Type",.Text7
		DropListBox 110,140,120,121,MyUnitArray(),.UnitDropList
		Text 10,14,100,14,"Select a Curve:",.Text8
		PushButton 640,182,90,21,"Help",.Help

	End Dialog
	Dim dlg As UserDialog

	If (Left(GetApplicationName,2) = "DS") Then
		sSelectedTreeItem = Replace(DS.GetSelectedTreeItem, DSResultFolder_LIB+"\", "")
		Set TreeObjectToCall = DSResultTree
		sTypeFromItemNameMethodToCall = "GetResultTypeFromItemName"
	Else
		sSelectedTreeItem = Replace(GetSelectedTreeItem, MWSResultFolder_LIB+"\", "")
		Set TreeObjectToCall = Resulttree
		sTypeFromItemNameMethodToCall = "GetTypeFromItemName"
	End If

	dlg.ResultListDLB = FindListIndex(aResultName(), sSelectedTreeItem)
	If dlg.ResultListDLB = -1 Then dlg.ResultListDLB = 0

	dlg.peakht = "30"
	dlg.peaknum = "20"

	dlg.UnitDropList = 2

	Dim low_freq As Double
	Dim hi_freq As Double

	dlg.fmin = "[Auto:fmin]"
	dlg.fmax = "[Auto:fmax]"

	dlg.ampupp = "[Auto:max]"
	dlg.amplow = "[Auto:min]"

	If Dialog(dlg)=0 Then Exit All

	If (Left(GetApplicationName,2) = "DS") Then
		sSelectedTreeItem = DSResultFolder_LIB+"\"+aResultName(dlg.ResultListDLB)
	Else
		sSelectedTreeItem = MWSResultFolder_LIB+"\"+aResultName(dlg.ResultListDLB)
	End If

	Set result = GetLastResult_LIB(aResultName(dlg.ResultListDLB), aResultType(dlg.ResultListDLB), "1DC")
	result_mag = result.Magnitude

	' Scale magnitude to dB
	With result_mag
		For i = 0 To .GetN-1
			If .GetY(i)>0 Then
				.SetXYDouble(i,.GetX(i),20.0*Log(.GetY(i))/Log(10))
			Else
				.SetXYDouble(i,.GetX(i),-120.0)
			End If
		Next
	End With

	dpeakht = CDbl(dlg.peakht)
	dpeaknum = CDbl(dlg.peaknum)
	dfmin = CDbl(Replace(dlg.fmin, "[Auto:fmin]", CStr(result_mag.GetX(0))))
	dfmax = CDbl(Replace(dlg.fmax, "[Auto:fmax]", CStr(result_mag.GetX(result_mag.GetN-1))))
	dampupp = CDbl(Replace(dlg.ampupp, "[Auto:max]", CStr(Round((result_mag.GetY(result_mag.GetMaximumInRange(dfmin,dfmax))+120),3))))
	damplow = CDbl(Replace(dlg.amplow, "[Auto:min]", CStr(Round((result_mag.GetY(result_mag.GetMinimumInRange(dfmin,dfmax))+120),3))))

	If dlg.UnitDropList=2 Then dampupp=dampupp-120
	If dlg.UnitDropList=2 Then damplow=damplow-120

	If dlg.UnitDropList=1 Then dampupp=dampupp-60
	If dlg.UnitDropList=1 Then damplow=damplow-60

	'MsgBox (aaa)

'======================================================================================================================
' Loop to determine the number of peaks (from UBE Macro:Results / Measure Resonances and Q-values from frq-data )
'======================================================================================================================

        Const NMax = 1000.0

        Dim cst_frequency(NMax) As Single
        Dim cst_amplitude(NMax) As Single

        Dim cst_frequency_range(NMax) As Single
        Dim cst_amplitude_range(NMax) As Single

        Dim n As Long

		n = result_mag.GetFirstMaximum(dpeakht)

		NFound = 0
		NFound_range = 0
        Do
        	 If NFound > NMax-1 Then
	             ReportInformationToWindow(CStr(Time)+": Maximum Number of Peaks Reached = " + CStr(NMax))
                 Exit Do
             End If

        	 If n = -1 Then
   	             ReportInformationToWindow(CStr(Time)+": No Peaks Found: Adjust Max Peak Height")
                 no_peaks = True
                 Exit Do
             End If

	      	 If NFound_range = dpeaknum Then
   	             ReportInformationToWindow(CStr(Time)+": Maximum Specified Number of Peaks Reached: "  + CStr(dpeaknum))
                 Exit Do
             End If

             no_peaks = False

             NFound = NFound+1

             cst_frequency(NFound)=result_mag.GetX(n)
             cst_amplitude(NFound)=result_mag.GetY(n)

			If cst_frequency(NFound) >= dfmin And cst_frequency(NFound) <= dfmax And cst_amplitude(NFound) >= damplow And (cst_amplitude(NFound)-0.001) <= dampupp Then

					NFound_range = NFound_range + 1

					cst_frequency_range(NFound_range) = cst_frequency(NFound)
					cst_amplitude_range(NFound_range) = cst_amplitude(NFound)
					no_peaks = False
			End If

             n = result_mag.GetNextMaximum(dpeakht)

        Loop Until n=-1

'======================================================================================================================
' Display Markers
'======================================================================================================================
'MsgBox(basenme)

If (Left(GetApplicationName(), 2) = "DS") Then
	DS.SelectTreeItem(sSelectedTreeItem)
Else
	SelectTreeItem(sSelectedTreeItem)
End If

With Plot1D
.DeleteAllMarker
.Plot
End With

Dim index As Long
With Plot1D
      index =.GetCurveIndexOfCurveLabel(CStr(basenme))
'	 For ii = 1 To NFound
	 For ii = 1 To NFound_range
'        .AddMarker(cst_frequency(ii))
        .AddMarker(cst_frequency_range(ii))
     Next ii
    .Plot ' make changes visible
End With

'======================================================================================================================
' Print values to a text file
'======================================================================================================================

If no_peaks = False Then

Dim Fname2 As String

'Fname2=GetProjectPath("Project")+"_Peakslist.txt"
'Fname2=GetProjectPath("Project")+"\Model\3D\Peaks.txt"


Dim position As Integer
position = Len(sSelectedTreeItem) - InStrRev(sSelectedTreeItem,"\")
Fname2=GetProjectPath("Project")+"\Model\3D\Peaks_" + Right(sSelectedTreeItem,position) + ".txt"

'MsgBox(GetProjectPath("Project")+"\Model\3D\Peaks_" + Right(sSelectedTreeItem,position) + ".txt" )

Open Fname2 For Output As #2
Print #2, "# ==============================================================="
Print #2, "# File Created on:"
Print #2, "# " + Cstr(Now) ' + vbCrLf
Print #2, "# Project File:"
Print #2, "# "+GetProjectPath("Project")+".cst"
Print #2, "# Result:"
Print #2, "# " + sSelectedTreeItem
Print #2, "# ==============================================================="


'Print #2, "# Total Number of Peaks:" + Str(NFound)
'Print #2, "# ==============================================================="

'Dim enum_para As String
'enum_para = GetUnit("Frequency")
'MsgBox(enum_para)

Print #2, "# Peak Number " + vbTab + "Freq(" + Units.GetUnit("Frequency") + ")" + vbTab + "Amplitude(dBuV)" + vbTab + "Amplitude(dB)"
Print #2, "# ==============================================================="

Dim curr_freq As String
Dim curr_amp As String
Dim curr_amp_dB As String
Dim curr_num As String

'For ii = 1 To NFound
'	curr_num = CStr(ii)
'	curr_freq = CStr(cst_frequency(ii))
'	curr_amp = CStr(cst_amplitude(ii))

For ii = 1 To NFound_range
	curr_num = CStr(ii)
	curr_freq = CStr(cst_frequency_range(ii))
	curr_amp = Left$(CStr(cst_amplitude_range(ii)+120.),6)
	curr_amp_dB = Left$(CStr(cst_amplitude_range(ii)),6)

Print #2, vbTab +  curr_num + vbTab + curr_freq + vbTab + vbTab+ curr_amp + vbTab + vbTab + curr_amp_dB
Next
Print #2, "# ==============================================================="
Close #2

'MsgBox "ASCII File successfully created: " + vbCrLf + Fname2

If (Left(GetApplicationName(), 2) = "DS") Then
	DS.ReportInformationToWindow(CStr(Time)+": ASCII File successfully created: " + vbCrLf + Fname2)
Else
	ReportInformationToWindow(CStr(Time)+": ASCII File successfully created: " + vbCrLf + Fname2)
End If

If GetApplicationName <> "DS" Then
	' this should not happen in stand alone DS
	With Resulttree
		.UpdateTree
		.RefreshView
	End With
End If

'Shell("notepad.exe " + Fname2, 1)

End If


End Sub

'-----------------------------------------------------------------------------------------------------------------------------
' Useful VBA procedures and functions  (from UBE Macro:Results / Measure Resonances and Q-values from frq-data )
'-----------------------------------------------------------------------------------------------------------------------------
'
' To avoid name collision with parameters in CST MicroWaveStudio,
' all variable names in this library have the prefix "lib_"
'
Public lib_FindFileRoot As String
Public lib_FindFilePattern As String
Public lib_FindFileRecursive As Boolean
Public lib_FindFileLast As String
'-----------------------------------------------------------------------------------------------------------------------------

Sub Start22 (lib_filename As String)

        On Error GoTo Win95
        WINNT:
                Shell "cmd /c " + Quote(lib_filename)
                Exit Sub
        Win95:
                Shell "start " + Quote(lib_filename)
                Exit Sub

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function FindFirstFile (lib_rootdir As String, lib_pattern As String, lib_recursive As Boolean) As String

        If (Right$(lib_rootdir, 1) = "\") Then lib_rootdir = Left$(lib_rootdir, Len(lib_rootdir) - 1)
        If (lib_pattern = "")             Then lib_pattern = "*.*"

        lib_FindFileRoot      = lib_rootdir
        lib_FindFilePattern   = lib_pattern
        lib_FindFileRecursive = lib_recursive
        lib_FindFileLast      = Dir$(lib_rootdir + "\*.*", vbDirectory)
        FindFirstFile     = FindNextFile()

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function FindNextFile () As String

        Dim lib_subdir As String, lib_currentdir, lib_dummy As String, lib_short As String, lib_filename As String
        Dim lib_fileattribute As Integer

        lib_subdir = DirName(lib_FindFileLast)
        lib_dummy  = Dir$()

        While (Not lib_dummy Like lib_FindFilePattern)

'                If MsgBox(IIf(lib_subdir <> "", lib_subdir + "\" + lib_dummy, lib_dummy), vbQuestion + vbYesNo, "Continue") <> vbYes Then Exit All
                lib_currentdir = lib_FindFileRoot + IIf(lib_subdir <> "", "\" + lib_subdir, "")

                On Error Resume Next
                        lib_fileattribute = 0
                        lib_fileattribute = GetAttr(lib_currentdir + "\" + lib_dummy)
                On Error GoTo 0

                If (lib_dummy = "") Then
                        If (lib_subdir = "") Then Exit While
                        lib_short      = ShortName(lib_subdir)
                        lib_subdir     = DirName(lib_subdir)
                        lib_currentdir = IIf(lib_subdir <> "", lib_FindFileRoot + "\" + lib_subdir, lib_FindFileRoot)
                        lib_dummy      = Dir$(lib_currentdir + "\*.*", vbDirectory)
                        While (lib_dummy <> lib_short)
                                lib_dummy = Dir$()
                        Wend
                        lib_dummy = Dir$()
                ElseIf (lib_dummy = ".") Or (lib_dummy = "..") Then
                        lib_dummy = Dir$()
                ElseIf (lib_fileattribute = vbDirectory) And lib_FindFileRecursive Then
                        lib_subdir     = IIf(lib_subdir <> "", lib_subdir + "\" + lib_dummy, lib_dummy)
                        lib_currentdir = IIf(lib_subdir <> "", lib_FindFileRoot + "\" + lib_subdir, lib_FindFileRoot)
                        lib_dummy      = Dir$(lib_currentdir + "\*.*", vbDirectory)
                Else
                        lib_dummy = Dir$()
                End If
        Wend
        lib_filename     = IIf(lib_dummy <> "", IIf(lib_subdir <> "", lib_subdir + "\", "") + lib_dummy, "")
        lib_FindFileLast = lib_filename

        FindNextFile     = lib_filename

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function DriveName (lib_path As String) As String

        DriveName  = Left$(lib_path, 1)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function DirName (lib_path As String) As String

        Dim lib_dircount As Integer

        lib_dircount = InStrRev(lib_path, "\")
        DirName  = Left$(lib_path, IIf(lib_dircount > 1, lib_dircount - 1, 0))

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function ShortName (lib_path As String) As String

        Dim lib_dircount As Integer

        lib_dircount  = InStrRev(lib_path, "\")
        ShortName = Mid$(lib_path, lib_dircount+1, 999)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function BaseName (lib_path As String) As String

        Dim lib_dircount As Integer, lib_extcount As Integer, lib_filename As String

        lib_dircount = InStrRev(lib_path, "\")
        lib_filename = Mid$(lib_path, lib_dircount+1)
        lib_extcount = InStrRev(lib_filename, ".")
        BaseName = Left$(lib_filename, IIf(lib_extcount > 0, lib_extcount-1, 999))

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function ExtName (lib_filename As String) As String

        Dim lib_extcount As Integer

        lib_extcount = InStrRev(lib_filename, ".")
        ExtName = IIf(lib_extcount > 0, Mid$(lib_filename, lib_extcount+1), "")

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function FullPath(lib_filename As String, lib_directory As String) As String

        FullPath = IIf(Mid$(lib_filename,2,1) = ":", lib_filename, lib_directory + "\" + lib_filename)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function ShortPath(lib_filename As String, lib_directory As String) As String

        If (LCase$(DirName(lib_filename)) = LCase$(lib_directory)) Then
                ShortPath = ShortName(lib_filename)
        Else
                ShortPath = lib_filename
        End If

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function Quote (lib_Text As String) As String

        Quote = Chr$(34) + lib_Text + Chr$(34)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function GetString(lib_application As String, lib_section As String, lib_key As String, lib_default As String) As String

        Dim lib_dummy As String

        lib_dummy = GetSetting(lib_application, lib_section, lib_key)
        GetString = IIf(lib_dummy <> "", lib_dummy, lib_default)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Function GetInteger(lib_application As String, lib_section As String, lib_key As String, lib_default As Integer) As Integer

        Dim lib_dummy As String

        lib_dummy  = GetSetting(lib_application, lib_section, lib_key)
        GetInteger = IIf(lib_dummy <> "", Val(lib_dummy), lib_default)

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Sub SaveString(lib_application As String, lib_section As String, lib_key As String, lib_value As String)

        SaveSetting lib_application, lib_section, lib_key, lib_value

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Sub SaveInteger(lib_application As String, lib_section As String, lib_key As String, lib_value As Integer)

        SaveSetting lib_application, lib_section, lib_key, CStr(lib_value)

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function RealVal(lib_Text As Variant) As Double

        If (CDbl("0.5") > 1) Then
                RealVal = CDbl(Replace(lib_Text, ".", ","))
        Else
		On Error Resume Next
                	RealVal = CDbl(lib_Text)
		On Error GoTo 0
        End If

End Function

'-----------------------------------------------------------------------------------------------------------------------------

Sub CheckApplication (lib_appname As String, lib_filename As String, lib_downloadsource As String)

        Dim lib_newline As String, lib_appdir As String, lib_hotlink As String
        lib_newline = Chr$(10) + Chr$(13)
        lib_appdir  = DirName(lib_filename)
        lib_hotlink = lib_appdir + lib_appname + ".url"

        If (Dir$(lib_filename, vbNormal) = "") Then
                If (MsgBox( _
                                lib_appname + " must be installed in" + lib_newline + lib_newline + _
                                lib_appdir                            + lib_newline + lib_newline + _
                                "Download " + lib_appname + " from"   + lib_newline + lib_newline + _
                                lib_downloadsource, _
                                vbOkCancel + vbQuestion, _
                                lib_appname + " is missing" _
                        ) = vbOK) Then
                        If (Dir$(lib_appdir, vbDirectory) = "") Then
                                MkDir lib_appdir
                        End If
                        Open lib_hotlink For Output As #1
                                Print #1, "[DEFAULT]"
                                Print #1, "BASEURL=" + lib_downloadsource
                                Print #1, "[InternetShortcut]"
                                Print #1, "URL=" + lib_downloadsource
                        Close #1
                        Shell "explorer " + lib_downloadsource
                End If
                Exit All
        End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Sub ShowHelp (lib_name As String)

        Dim lib_filename As String, lib_extension() As Variant, lib_index As Integer

        lib_extension = Array("htm", "html", "txt", "rtf", "doc", "hlp")

        lib_index = 0
        Do
                    lib_filename = GetMacroPath + "\" + Dir$(GetMacroPath + "\*" + lib_name + "." + lib_extension(lib_index))
                    lib_index = lib_index + 1
        Loop Until (lib_filename <> "" Or lib_extension(lib_index) = "")

        Start22 lib_filename

End Sub

'-----------------------------------------------------------------------------------------------------------------------------
