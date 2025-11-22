'#Language "WWB-COM"

' ==============================================================================================================================================================
' This macro imports multiple nearfield or farfield sources.'
'
' Copyright 2021-2023 Dassault Systemes Deutschland GmbH
' ==============================================================================================================================================================
' History of Changes
' ------------------
' 25-Sep-2022 ube: included in distribution
' 17-Dec-2021 dta: initial version
' ==============================================================================================================================================================

Option Explicit

'#include "vba_globals_all.lib"

Sub Main

	Begin Dialog UserDialog 330,175,"Import multiple sources",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 20,7,290,56,"Source type",.GroupBox2
		PushButton 60,28,90,21,"Nearfield",.nearfieldPB
		PushButton 180,28,90,21,"Farfield",.farfieldPB
		CancelButton 20,147,90,21
		GroupBox 20,63,290,77,"",.GroupBox1
		Text 40,77,260,56,"INFO:" + vbCrLf + vbCrLf + "All the sources of the same type present in the selected folder will be imported.",.Text2
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "nearfieldPB"
				If (ImportNearfieldSource = -1) Then ' error
					DialogFunc = True
				End If
			Case "farfieldPB"
				If (ImportFarfieldSource = -1) Then ' error
					DialogFunc = True
				End If
			Case "Cancel"
				DialogFunc = False
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function ImportNearfieldSource() As Integer

	Dim sourceFullName As String, sourceName As String, sourceFolder As String
	Dim index As Long
	Dim sCommand As String

	sourceFullName = GetFilePath("*.fsm;","Nearfield source Files|*.fsm|All Files|*.*",GetProjectPath("Root"),"Select Field source monitor file to load",0)
	If sourceFullName = "" Then Exit All


	sourceName=Split(sourceFullName,"\")(UBound(Split(sourceFullName,"\")))
	sourceName = Split(sourceName,".fsm")(0)

	'===============Get root path============
    index=InStrRev (sourceFullName, "\")			'Returns the index of the last "\" in the datafilename path
    sourceFolder= Left$(sourceFullName,index-1)		'Trim the root path starting from the given index-1
	'========================================

	'MsgBox(sourceName)

	sourceName=FindFirstFile(sourceFolder,"*.fsm*", False)
		sourceFullName=sourceFolder+"\"+sourceName
		sourceName = Split(sourceName,".fsm")(0)

		While (sourceName <> "")


	'		With FieldSource
    ' .Reset
    ' .Name "fs1"
    ' .FileName "E:\004_New_Release_Testing\2022\zzz_Macro_Template\006_Import multiple sources\nearfield sources\field-source (f=1)_compressed_1.fsm"
    ' .Id "0"
    ' .ImportToActiveCoordinateSystem "True"
    ' .Read
	'	End With

			'@ define farfield source
				sCommand = ""
				sCommand = sCommand + "With FieldSource " + vbLf
				sCommand = sCommand + ".Reset " + vbLf
				sCommand = sCommand + ".Name """+sourceName+"""" + vbLf
				FieldSource.FileName sourceFullName
				sCommand = sCommand + ".Filename """+sourceFullName+"""" + vbLf
				sCommand = sCommand + ".Id """+FieldSource.GetNextId()+"""" + vbLf
				sCommand = sCommand + ".ImportToActiveCoordinateSystem ""True""" + vbLf
				sCommand = sCommand + ".Read " + vbLf
				sCommand = sCommand + "End With"
				AddToHistory "define field source: "+sourceName+"", sCommand


			sourceName = FindNextFile()
			If sourceName<>"" Then
				sourceFullName=sourceFolder+"\"+sourceName
				sourceName = Split(sourceName,".fsm")(0)
			End If

		Wend

'=============================================
		UpdateTree

	ImportNearfieldSource=0

End Function

Function ImportFarfieldSource() As Integer

	Dim sourceFullName As String, sourceName As String, sourceFolder As String
	Dim index As Long
	Dim Position(2) As Double
	Dim ThetaZAxis(2) As Double
	Dim PhiXAxis(2) As Double
	Dim FarfieldAlignment (2) As String
	Dim CalcMPCoeff As String, EnableMPCoeff As Boolean


	Dim SetFarfieldAlignment As String

	Dim sCommand As String

	ImportFarfieldSource = -1 ' error

	Position(0) = 0
	Position(1) = 0
	Position(2) = 0
	ThetaZAxis(0) = 0
	ThetaZAxis(1) = 0
	ThetaZAxis(2) = 1
	PhiXAxis(0) = 1
	PhiXAxis(1) = 0
	PhiXAxis(2) = 0

	FarfieldAlignment(0)="User defined"
	FarfieldAlignment(1)="Current WCS"
	FarfieldAlignment(2)="Source file"

	sourceFullName = GetFilePath("*.ffs","Farfield Files|*.ffs|All Files|*.*",GetProjectPath("Root"),"Select farfield to load",0)
	If (sourceFullName = "") Then
		ImportFarfieldSource = -1
		Exit Function
	Else
		sourceName = Split(sourceFullName,"\")(UBound(Split(sourceFullName,"\")))
		sourceName = Split(sourceName,".ffs")(0)

	'===============Get root path============
    index=InStrRev (sourceFullName, "\")			'Returns the index of the last "\" in the datafilename path
    sourceFolder= Left$(sourceFullName,index-1)		'Trim the root path starting from the given index-1
	'========================================

	Begin Dialog UserDialog 510,287,"Farfield sources import settings",.DialogFuncFarfieldImport ' %GRID:10,7,1,1

		OKButton 410,14,90,21
		CancelButton 410,42,90,21
		GroupBox 20,63,360,49,"Position",.GroupBox2
		GroupBox 20,119,360,49,"Start for theta (z'-axis)",.GroupBox3
		GroupBox 20,175,360,49,"Start for phi (x'-axis)",.GroupBox4
		GroupBox 20,231,360,49,"Multipole coefficients",.GroupBox5
		TextBox 60,84,60,21,.xPosT
		TextBox 180,84,60,21,.yPosT
		TextBox 300,84,60,21,.zPosT
		TextBox 60,140,60,21,.XTheta
		GroupBox 20,7,360,49,"Alignment",.GroupBox1
		TextBox 60,196,60,21,.XPhi
		TextBox 180,140,60,21,.YTheta
		TextBox 180,196,60,21,.YPhi
		TextBox 300,140,60,21,.ZTheta
		TextBox 300,196,60,21,.ZPhi
		CheckBox 60,252,20,14,"CheckBox1",.MPCoeff
		Text 90,252,220,14,"Calculate multipole coefficients",.Text1
		Text 40,91,10,14,"X",.Text2
		Text 160,91,10,14,"Y",.Text3
		Text 280,91,10,14,"Z",.Text4
		Text 40,145,10,14,"X",.Text5
		Text 160,145,10,14,"Y",.Text6
		Text 280,145,10,14,"Z",.Text7
		Text 40,199,10,14,"X",.Text8
		Text 160,199,10,14,"Y",.Text9
		Text 280,199,10,14,"Z",.Text10
		DropListBox 60,28,250,21,FarfieldAlignment(),.Alignment


	End Dialog

		Dim dlg As UserDialog


		dlg.xPosT = cstr(Position(0))
		dlg.yPosT = cstr(Position(1))
		dlg.zPosT = cstr(Position(2))
		dlg.XTheta = cstr(ThetaZAxis(0))
		dlg.YTheta = cstr(ThetaZAxis(1))
		dlg.ZTheta = cstr(ThetaZAxis(2))
		dlg.XPhi = cstr(PhiXAxis(0))
		dlg.YPhi = cstr(PhiXAxis(1))
		dlg.ZPhi = cstr(PhiXAxis(2))

		If (Dialog(dlg) = 0) Then ' User pressed Cancel
			ImportFarfieldSource = -1
			Exit Function
		End If


		Position(0) = Evaluate(dlg.xPosT)
		Position(1) = Evaluate(dlg.yPosT)
		Position(2) = Evaluate(dlg.zPosT)
		ThetaZAxis(0) = Evaluate(dlg.XTheta)
		ThetaZAxis(1) = Evaluate(dlg.YTheta)
		ThetaZAxis(2) = Evaluate(dlg.ZTheta)
		PhiXAxis(0) = Evaluate(dlg.XPhi)
		PhiXAxis(1) = Evaluate(dlg.YPhi)
		PhiXAxis(2) = Evaluate(dlg.ZPhi)
		EnableMPCoeff=Evaluate(dlg.MPCoeff)


		MakeSureParameterExists "ffs_pos_x", 	Position(0)
		SetParameterDescription  ( "ffs_pos_x",  "Farfield source position in X"  )
		MakeSureParameterExists "ffs_pos_y", 	Position(1)
		SetParameterDescription  ( "ffs_pos_y",  "Farfield source position in Y"  )
		MakeSureParameterExists "ffs_pos_z", 	Position(2)
		SetParameterDescription  ( "ffs_pos_z",  "Farfield source position in Z"  )

		If dlg.MPCoeff Then
			CalcMPCoeff="true"
		Else
			CalcMPCoeff="false"
		End If


		If dlg.Alignment=0 Then
			SetFarfieldAlignment="user"
		ElseIf dlg.Alignment=1 Then
			SetFarfieldAlignment="currentwcs"
		Else
			SetFarfieldAlignment="sourcefile"
		End If

'=============================================

		sourceName=FindFirstFile(sourceFolder,"*.ffs*", False)
		sourceFullName=sourceFolder+"\"+sourceName
		sourceName = Split(sourceName,".ffs")(0)

		While (sourceName <> "")

			'@ define farfield source
				sCommand = ""
				sCommand = sCommand + "With FarfieldSource " + vbLf
				sCommand = sCommand + ".Reset " + vbLf
				sCommand = sCommand + ".Name """+sourceName+"""" + vbLf
				sCommand = sCommand + ".Id """+FARFIELDSOURCE.GetNextId+"""" + vbLf
				sCommand = sCommand + ".UseCopyOnly ""True""" + vbLf

				sCommand = sCommand + ".SetPosition ""ffs_pos_x"",""ffs_pos_y"", ""ffs_pos_z""" + vbLf
				sCommand = sCommand + ".SetTheta0XYZ """+CStr(ThetaZAxis(0))+""","""+CStr(ThetaZAxis(1))+""", """+CStr(ThetaZAxis(2))+"""" + vbLf
				sCommand = sCommand + ".SetPhi0XYZ """+CStr(PhiXAxis(0))+""","""+CStr(PhiXAxis(1))+""", """+CStr(PhiXAxis(2))+"""" + vbLf

				sCommand = sCommand + ".Import """+sourceFullName+"""" + vbLf
				sCommand = sCommand + ".UseMultipoleFFS """+CalcMPCoeff+"""" + vbLf
				sCommand = sCommand + ".SetAlignmentType """+SetFarfieldAlignment+"""" + vbLf
				sCommand = sCommand + ".SetMultipoleDegree ""1""" + vbLf
				sCommand = sCommand + ".SetMultipoleCalcMode ""automatic""" + vbLf
				sCommand = sCommand + ".Store " + vbLf
				sCommand = sCommand + "End With"
				AddToHistory "define farfield source: "+sourceName+"", sCommand


			sourceName = FindNextFile()
			If sourceName<>"" Then
				sourceFullName=sourceFolder+"\"+sourceName
				sourceName = Split(sourceName,".ffs")(0)
			End If

		Wend

'=============================================
		UpdateTree

		ImportFarfieldSource = 0

	End If

End Function

Private Function DialogFuncFarfieldImport(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
				DlgValue  "MPCoeff",1
				DlgEnable "MPCoeff",True
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
				Case "Alignment"
		    		If (SuppValue = 1 Or SuppValue=2) Then
						DlgEnable "xPosT",False
						DlgEnable "yPosT",False
						DlgEnable "zPosT",False
						DlgEnable "XTheta",False
						DlgEnable "YTheta",False
						DlgEnable "ZTheta",False
						DlgEnable "XPhi",False
						DlgEnable "YPhi",False
						DlgEnable "ZPhi",False
						DialogFuncFarfieldImport = True
		    		Else
						DlgEnable "xPosT",True
						DlgEnable "yPosT",True
						DlgEnable "zPosT",True
						DlgEnable "XTheta",True
						DlgEnable "YTheta",True
						DlgEnable "ZTheta",True
						DlgEnable "XPhi",True
						DlgEnable "YPhi",True
						DlgEnable "ZPhi",True
						DialogFuncFarfieldImport = True
					End If
				Case "Cancel"
				DialogFuncFarfieldImport = False
				Exit All
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
