' *File / Open directory Global Macro Path
' !!! Do not change the line above !!!
' ================================================================================================
' Copyright 2003-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 13-Jul-2016 ube: add selection of multiple library pathes
' 16-Feb-2016 tsi: Allow file browsing for Linux.
' 05-Jan-2010 ube: include quotes around pathname, otherwise komma in path not working
' 05-Jan-2003 ube: first version
'------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

Sub Main

	Dim i As Long, sLibPathes() As String
	For i = 0 To GetNumberOfLibraryPathes()-1
		ReDim Preserve sLibPathes(i)
		sLibPathes(i) = GetLibraryPath(i)
	Next

	Begin Dialog UserDialog 700,133,"Choose Library Path to open" ' %GRID:10,7,1,1
		ListBox 10,14,680,84,sLibPathes(),.ListBox1
		OKButton 10,105,90,21
		CancelButton 110,105,90,21
	End Dialog
	Dim dlg As UserDialog

	If (Dialog(dlg) <> 0) Then
		Dim lib_ProjectDir As String
	    Dim fileBrowser As String
	    Dim args As String
	    fileBrowser = GetDefaultFileBrowser + " "
		lib_ProjectDir = sLibPathes(dlg.ListBox1)
	    args = MakeNativePath(lib_ProjectDir)
		Shell fileBrowser & Chr$(34) + args + Chr$(34), 1
	End If

End Sub
