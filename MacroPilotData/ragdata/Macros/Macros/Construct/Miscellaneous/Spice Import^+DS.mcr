'#Language "WWB-COM"

' This macro imports a SPICE netlist and generates a single CST DESIGN STUDIO projet block from it, representing a toplevel subcircuit, from it on the schematic.
'-------------------------------------------
' Version history:
' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 19-Aug-2022 cri: Created first version.
' 24-Aug-2022 cri: Replaced edit text field for subcircuit name by non-editable dropdown box extracted after browing circuit file.
' 26-Aug-2022 cri: Removed unnecessary check if project and netlist file are on the same drive. 
' ---------------------------------------------------------------------------------------------------------------------------------
Option Explicit

Dim sPythonScript As String, sCompleteShellCommand
Dim sPythonExe As String
Const macroPath = "\Library\Macros\Construct\Miscellaneous\"
Const CSTStudioAppdataPath = "\DASSAULT_SYSTEMES\CSTStudioSuite\"

Public Function BrowseNetlistFile() As String
	BrowseNetlistFile = GetFilePath("*.sp; *.cir; *.net; *.txt", "SPICE files|*.*", , "Please select SPICE file to be imported", 0)
End Function

Public Function OnOK(ByVal netlistFileName As String, ByVal subcktName As String, ByVal dialect As String)
	DS.ReportInformationToWindow("Starting SPICE to schematic import.")
	Dim projectPath As String
	projectPath = GetProjectPath("Root")
	DS.ImportSPICEFromFile(netlistFileName, subcktName, dialect)
	DS.ReportInformationToWindow("Netlist " + netlistFileName + " has been successfully imported.")
End Function

Function FileDlgFunction(identifier$, action, suppvalue)

Dim myarray$(3)

Dim msgtext As Variant

Dim x As Integer

For x= 0 To 2

   myarray$(x)=Chr$(x+65)

Next x

   Select Case action

      Case 1

      Case 2  'user changed control or clicked a button

         If DlgControlId(identifier$)=3 Then

             If DlgListBoxArray(2)=0 Then

                     DlgListBoxArray 2, myarray$()

             End If

         End If

      End Select

End Function

Public Function IsArrayEmpty(arr As Variant) As Long
  On Error GoTo handler

  Dim lngUpper As Long
  lngUpper = UBound(arr)
  IsArrayEmpty = False
  Exit Function

handler:
	IsArrayEmpty = True
End Function


Private Function DialogFunc(DlgItem$, action%, suppvalue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses Any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
	Dim fileCreated As Boolean
	Select Case action%
	Case 1 ' Dialog box initialization
		DlgEnable ("filename", False)
    Case 2 ' Value changing or button pressed'
    	DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
		Case "browse"
			Dim netlistFile As String
			netlistFile = BrowseNetlistFile()
 			DlgText("fileName", netlistFile)
 			Dim toplevelSubckts As Variant
 			toplevelSubckts = DS.GetToplevelSubcircuits(netlistFile,DlgText("dialect"))

 			If IsArrayEmpty(toplevelSubckts) Then
 				ReDim Preserve subckts$(0)
 				DlgListBoxArray "subckt",subckts$()
 			Else
				Dim N As Integer
				N = UBound(toplevelSubckts)
				ReDim Preserve subckts$(N)
				Dim nIndex As Long
				For nIndex = 0 To N
					subckts(nIndex)=toplevelSubckts(nIndex)
				Next nIndex
				DlgListBoxArray "subckt",subckts$()
 				DlgValue("subckt", 0)
 			End If

		Case "OKButtonPushed"
			OnOK(DlgText("filename"), DlgText("subckt"), DlgText("dialect"))
			Exit All
		Case "HelpButtonPushed"
			StartDESHelp "macro\common_macro_spice2schematicimport"
		Case "CancelButtonPushed"
			Exit All
	End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    	DialogFunc = True ' Prevent button press from closing the dialog box

    Case 6 ' Function key
    End Select
End Function

Sub Main()
    Dim lists$(3)
    lists$(0) = "SPICE3f5"
    lists$(1) = "PSPICE"
    lists$(2) = "HSPICE"
    lists$(3) = "Combined"
    Dim subckts$(0)
    Begin Dialog UserDialog 520,160, "Import SPICE netlist to schematic", .DialogFunc
        PushButton 10,15,80,20,"&Browse...", .browse
        Text 10,40,270,15,"File name", .Text3
        TextBox 100,40,405,15, .filename
		Text 10,65,270,15,"Subcircuit", .Text4
        DropListBox 100,65,405,15,subckts$(),.subckt
		Text 10,90,170,15,"Spice format"
      	DropListBox 100,90,100,15,lists$(),.dialect
        PushButton 30,124,60,20, "&OK", .OKButtonPushed
        PushButton 100,124,60,20, "&Cancel", .CancelButtonPushed
        PushButton 170,124,60,20, "&Help",.HelpButtonPushed
    End Dialog
    Dim dlg As UserDialog
    dlg.filename = ""
    dlg.subckt = 0
    dlg.dialect = 3
    Dialog dlg ' show dialog
End Sub
