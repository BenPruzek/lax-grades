
'The macro allows to convert the farfield data from FEKO (.ffe format) to CST (.ffs format)
' useful to extract data on Faces.
'
' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 22-May-2015 dta: made more robust data extraction
' 06-Aug-2014 dta: support of multifrequency file
' 05-Aug-2014 dta: made more robust the data extraction
' 22-Jul-2014 dta: initial version
''-----------------------------------------------------------------------------------------------------------------------------
Option Explicit

'#include "vba_globals_all.lib"


Sub Main

	Dim infile As String
	Dim outfile As String
	Dim inline As String
	Dim Frequencies() As String, Theta_Samples As String, Phi_Samples As String
	Dim Freq_samples As Integer
	Dim Origin_X As String, Origin_Y As String, Origin_Z As String
	Dim Theta As String, Phi As String, ETheta_Re As String, ETheta_Im As String, EPhi_Re As String, EPhi_Im As String
	Dim Write_FFS_Header As Boolean, Is_Next_Freq As Boolean
	Dim Temp_String As String
	Dim linein_error As Boolean, fval_ok As Boolean
	Dim count As Long

	Dim sep As String, sep_del As String    'separator used to properly get data from each line

	
	Begin Dialog UserDialog 650,112,"Convert Farfield data from FEKO to CST",.DialogFunc ' %GRID:10,7,1,1
		Text 20,35,80,14,"Input File",.Text1
		Text 20,63,80,14,"Output File",.Text2
		TextBox 100,28,440,21,.Infile
		TextBox 100,56,440,21,.Outfile
		PushButton 550,28,90,21,"Browse",.Browseinputfile
		OKButton 110,84,90,21
		CancelButton 210,84,90,21
		Text 10,7,630,14,"The scritp will convert a [multifrequency] farfield (.ffe) exported from FEKO (v.6.1 or later) to CST (.ffs)",.Text3
	End Dialog
	Dim dlg As UserDialog
	If (Dialog(dlg) = 0) Then Exit All


	'infile = "farfield_data.ffe"
	infile = dlg.Infile
	'outfile = "farfield_dta.ffs"
	outfile = dlg.Outfile
	linein_error = False

	'Inizialization of variables

	Write_FFS_Header=False  		'the boolean value is used to write the header only once in the file
	Is_Next_Freq=False		         'next frequency sample will be written

	Freq_samples=0

	Origin_X="0.000000e+000"
	Origin_Y="0.000000e+000"
	Origin_Z="0.000000e+000"

	sep = " "       'substitute occurances with a empty space
	sep_del=""		'substiture occurances deleting them

	'=====================================================================================
	'Retrieve number of frequency samples
	'======================================================================================

	'---Loop for reading lines
	Open infile For  Input As #1
	While Not EOF(1)

		'--- data from file
		Line Input #1,inline


		If (Mid$(inline,2,9))="Frequency" Then					'record  frequency
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)

			Temp_String = (Split(inline,sep)(1))

			Freq_samples = Freq_samples+1
			ReDim Preserve Frequencies(Freq_samples)
			Frequencies(Freq_samples-1)=Temp_String

		End If
	Wend

	Close #1

	'=====================================================================================
	'======================================================================================

	'---Loop for reading lines
	Open infile For  Input As #1
	Open outfile For Output As #2
	While Not EOF(1)
	
		'--- data from file
		Line Input #1,inline

	If Mid$(inline,1,1)<>"#" And inline<>"" And Mid$(inline,1,2)<>"**" Then    'jump to recording data section

		GoTo RECORDDATALINE

	ElseIf (Mid$(inline,1,2)="##" Or Mid$(inline,1,2)="**")Or inline="" Then  'jump to next line since this is a comment or an empty line
		GoTo WLOOPEND
	ElseIf Mid$(inline,1,1)="#" Then 		'read data to be transferred to  the header of the .ffs file
		If (Mid$(inline,2,9))="Frequency" Then
			Is_Next_Freq=True

		GoTo WLOOPEND

		ElseIf (Mid$(inline,2,17))="Coordinate System" Then
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)
			Temp_String = (Split(inline,sep)(2))
			If Temp_String<>"Spherical" Then
				MsgBox "Only Spherical coordinate system is currently supported",vbOkOnly+vbCritical,"Execution stopped"
			Exit Sub
			End If
		GoTo WLOOPEND
		ElseIf (Mid$(inline,2,6))="Origin" Then
			inline = Replace(inline,"(",sep,,)
			inline = Replace(inline,")",sep,,)
			inline = Replace(inline,",",sep,,)
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)

			Origin_X = (Split(inline,sep)(1))
			Origin_Y = (Split(inline,sep)(2))
			Origin_Z = (Split(inline,sep)(3))

			GoTo WLOOPEND
		ElseIf (Mid$(inline,2,12))="No. of Theta" Then		'record Theta samples
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)
			Temp_String = (Split(inline,sep)(4))
			Theta_Samples=Temp_String
			GoTo WLOOPEND
		ElseIf (Mid$(inline,2,10))="No. of Phi" Then		'record Phi samples
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)
			Temp_String = (Split(inline,sep)(4))
			Phi_Samples=Temp_String
			GoTo WLOOPEND
		ElseIf (Mid$(inline,2,14))="Far Field Type" Then
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)
			Temp_String = (Split(inline,sep)(3))
			If (Temp_String<>"Gain" And Temp_String<>"Directivity")  Then
				MsgBox "Only Gain or Directivity are supported",vbOkOnly+vbCritical,"Execution stopped"
			Exit Sub
			End If
			GoTo WLOOPEND
		Else    			'skip lines starting with # and not containing useful info
			GoTo WLOOPEND
		End If

		RECORDDATALINE:
		If Write_FFS_Header=False Then   'write FFS header

		Print #2, "// CST Farfield Source File"
		Print #2, ""
		Print #2, "// Version:"
		Print #2, "3.0"
		Print #2, ""
		Print #2, "// Data Type"
		Print #2, "Farfield"
		Print #2, ""
		Print #2, "// #Frequencies"
		Print #2, Cstr(Freq_samples)
		Print #2, ""
		Print #2, "// Position"
		Print #2, Origin_X+" "+Origin_Y+" "+Origin_Z
		Print #2, ""
		Print #2, "// zAxis"
		Print #2, "0.000000e+000 0.000000e+000 1.000000e+000"
		Print #2, ""
		Print #2, "// xAxis"
		Print #2, "1.000000e+000 0.000000e+000 0.000000e+000"
		Print #2, ""
		Print #2, "// Radiated/Accepted/Stimulated Power, Frequency"

	For count=0 To Freq_samples-1
		Print #2, "0.000000e+000"
		Print #2, "0.000000e+000
		Print #2, "0.000000e+000"
		Print #2, Frequencies(count)
		Print #2, ""
	Next count
		Write_FFS_Header=True

		End If

		If Is_Next_Freq Then        'write header
		Print #2, ""
		Print #2, "// >> Total #phi samples, total #theta samples"
		Print #2, Phi_Samples+" "+Theta_Samples
		Print #2, ""
		Print #2, "// >> Phi, Theta, Re(E_Theta), Im(E_Theta), Re(E_Phi), Im(E_Phi):"

		Is_Next_Freq=False

		End If

	inline = Replace(inline,sep+sep+sep,sep,,)
	inline = Replace(inline,vbTab,sep,,)
	inline = Replace(inline,sep+sep,sep,,)

	'==========verify if there are still empty spaces to be removed
	While Mid$(inline,1,1)=" "
		inline = Replace(inline,sep,sep_del,,1)
	Wend

	Theta = (Split(inline,sep)(0))
	Phi = (Split(inline,sep)(1))
	ETheta_Re = (Split(inline,sep)(2))
	ETheta_Im = (Split(inline,sep)(3))
	EPhi_Re = (Split(inline,sep)(4))
	EPhi_Im = (Split(inline,sep)(5))

	If ETheta_Re="NaN" Or ETheta_Im="NaN" Or ETheta_Re="NaN" Or ETheta_Im="NaN" Then
	MsgBox "The line related to the direction Phi="+CDbl(Phi)+", Theta="+CDbl(Theta)+" doesn't contain valid numeric data (NaN). Check your file.",vbOkOnly+vbCritical,"Execution stopped"
	Exit Sub
	End If

	Print #2, Phi+" "+Theta+" "+ETheta_Re+" "+ETheta_Im+" "+EPhi_Re+" "+EPhi_Im

	End If

	WLOOPEND:
	Wend

	Close #1
	Close #2
	
End Sub
'-------------------------------------------------------------------------------------
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean

	Dim Extension As String, projectdir As String, filename As String
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        Select Case DlgItem
        	Case "Browseinputfile"
        		Extension = "ffe"
                projectdir = GetProjectPath("Root")
                filename = GetFilePath(,Extension, projectdir, "Browse for farfield file", 0)
                If (filename <> "") Then
                    DlgText "Infile", filename
                    DlgText "Outfile", Dirname(filename)+"\"+Basename(filename)+"_CST.ffs"
                End If
        	dialogfunc = True
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
