'The macro allows to convert the farfield data from HFSS (.ffd format) to CST (.ffs format)

' ================================================================================================
' Copyright 2014-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 29-Nov-2024 dta: check angular range  (theta=0-180, phi=0-360), otherwise error and exit for now
' 28-Oct-2024 dta: initial version
'--------------------------------------------------------------------------------------------------------------------------
Dim Global_Infile As String, Global_Outfile As String

Option Explicit

'#include "vba_globals_all.lib"

Sub Main

	Dim linein_error As Boolean, fval_ok As Boolean
	
	Begin Dialog UserDialog 740,112,"Convert Farfield data from HFSS to CST",.DialogFunc2 ' %GRID:10,7,1,1
		Text 20,35,80,14,"Input File",.Text1
		Text 20,63,80,14,"Output File",.Text2
		TextBox 100,28,530,21,.Infile
		TextBox 100,56,530,21,.Outfile
		PushButton 640,28,90,21,"Browse",.Browseinputfile
		OKButton 110,84,90,21
		CancelButton 210,84,90,21
		Text 10,7,660,14,"The scritp will convert a [multifrequency] farfield (.ffd) exported from HFSS to CST farfield source(.ffs).",.Text3
	End Dialog
	Dim dlg2 As UserDialog
	If (Dialog(dlg2) = 0) Then Exit All


	'infile = "farfield_data.ffe"
	Global_Infile= dlg2.Infile
	'outfile = "farfield_dta.ffs"
	Global_Outfile = dlg2.Outfile
	linein_error = False



	Begin Dialog UserDialog 510,98,"Convert Farfield data from HFSS to CST",.DialogFunc ' %GRID:10,7,1,1
		Text 20,14,260,14,"Converting file... "
		Text 20,42,490,14,"",.OutputT
		OKButton 20,70,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

Function ExportFFDtoFFS(DataFileName As String) As Integer

	Dim inline As String
	Dim Frequencies() As String
	Dim Theta_Start As String, Theta_Stop As String, Theta_Samples As String, Phi_Start As String, Phi_Stop As String, Phi_Samples As String
	Dim Theta_Step As Double, Phi_Step As Double
	Dim Phi_i As Long, Theta_i As Long, freq_i As Long
	Dim Freq_samples As Integer
	Dim Origin_X As String, Origin_Y As String, Origin_Z As String
	Dim Theta As String, Phi As String, ETheta_Re As String, ETheta_Im As String, EPhi_Re As String, EPhi_Im As String
	Dim Write_FFS_Header As Boolean, Is_Next_Freq As Boolean
	Dim Temp_String As String
	Dim linein_error As Boolean, fval_ok As Boolean
	Dim count As Long, line_number As Long
	Dim sep As String, sep_del As String    'separator used to properly get data from each line

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
	'Retrieve Theta and Phi samples
	'======================================================================================

	'---Loop for reading lines
	Open Global_Infile For Input As #1

		'--- data from file
		Line Input #1,inline

			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)

		'==========verify if there are still empty spaces to be removed
		While Mid$(inline,1,1)=" "
			inline = Replace(inline,sep,sep_del,,1)
		Wend

		Theta_Start = (Split(inline,sep)(0))
		Theta_Stop = (Split(inline,sep)(1))
		Theta_Samples = (Split(inline,sep)(2))
		Theta_Step=(CdBl(Theta_Stop)-CdBl(Theta_Start))/(CdBl(Theta_Samples)-1)



		Line Input #1,inline

			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)

		'==========verify if there are still empty spaces to be removed
		While Mid$(inline,1,1)=" "
			inline = Replace(inline,sep,sep_del,,1)
		Wend

		Phi_Start = (Split(inline,sep)(0))
		Phi_Stop = (Split(inline,sep)(1))
		Phi_Samples = (Split(inline,sep)(2))
		Phi_Step=(CdBl(Phi_Stop)-CdBl(Phi_Start))/(CdBl(Phi_Samples)-1)


		'=====================================================================================
		'perform a check on the angular range validity
		'=====================================================================================

		If CDbl(Theta_Stop)-CDbl(Theta_Start)>180 Or CdBl(Theta_Start)<>0 Then
			ReportError("Please check the Theta frequency range. Allowed range is from 0  to 180 ")
			'Exit All
		End If

		If CDbl(Phi_Stop)-CDbl(Phi_Start)>360 Or CdBl(Phi_Start)<>0 Then
			ReportError("Please check the Phi frequency range. Allowed range is from 0  to 360 ")
			'Exit All
		End If

	Close #1

	'=====================================================================================
	'Retrieve number & frequency sample value
	'======================================================================================

	'---Loop for reading lines
	Open Global_Infile For  Input As #1
	While Not EOF(1)

		'--- data from file
		Line Input #1,inline


		If (Mid$(inline,1,9))="Frequency" Or (Mid$(inline,1,9))="frequency" Then					'record  frequency
			inline = Replace(inline,sep+sep+sep,sep,,)
			inline = Replace(inline,vbTab,sep,,)
			inline = Replace(inline,sep+sep,sep,,)

			Temp_String = (Split(inline,sep)(1))

			Freq_samples = Freq_samples+1
			ReDim Preserve Frequencies(Freq_samples-1)
			Frequencies(Freq_samples-1)=Temp_String

		End If
	Wend

	Close #1

	'=====================================================================================
	'---Loop for reading field values
	'=====================================================================================

	Open Global_Outfile For Output As #2

	Print #2, "// CST Farfield Source File"					'write CST ffs header part
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

	For freq_i=1 To Freq_samples				'Loop over frequency

					Print #2, ""
					Print #2, "// >> Total #phi samples, total #theta samples"
					Print #2, Phi_Samples+" "+Theta_Samples
					Print #2, ""
					Print #2, "// >> Phi, Theta, Re(E_Theta), Im(E_Theta), Re(E_Phi), Im(E_Phi):"


		For Phi_i=1 To CLng(Phi_Samples) 		'Loop over Phi samples

			Open Global_Infile For Input As #1

			line_number=0

			Theta_i=1


			While Not EOF(1)

				'--- data from file
				Line Input #1,inline

				line_number=line_number+1

				If line_number>(3+freq_i+(freq_i-1)*(CLng(Phi_Samples)*CLng(Theta_Samples))) And line_number<(3+freq_i+freq_i*(CLng(Phi_Samples)*CLng(Theta_Samples))+1) And line_number=(3+freq_i+(freq_i-1)*(CLng(Phi_Samples)*CLng(Theta_Samples))+(Phi_i+(Theta_i-1)*CLng(Phi_Samples)))    Then    'jump to recording data section


					inline = Replace(inline,sep+sep+sep,sep,,)
					inline = Replace(inline,vbTab,sep,,)
					inline = Replace(inline,sep+sep,sep,,)

					'==========verify if there are still empty spaces to be removed
					While Mid$(inline,1,1)=" "
						inline = Replace(inline,sep,sep_del,,1)
					Wend

					'Theta = (Split(inline,sep)(0))
					'Phi = (Split(inline,sep)(1))
					ETheta_Re = (Split(inline,sep)(0))
					ETheta_Im = (Split(inline,sep)(1))
					EPhi_Re = (Split(inline,sep)(2))
					EPhi_Im = (Split(inline,sep)(3))


					Print #2,CStr((Phi_i-1)*Phi_Step)+" "+CStr((Theta_i-1)*Theta_Step)+" "+ETheta_Re+" "+ETheta_Im+" "+EPhi_Re+" "+EPhi_Im

					If Theta_i=CLng(Theta_Samples) Then

						GoTo WLOOPEND
					End If

					Theta_i=Theta_i+1

				End If


			Wend

			WLOOPEND:

			Close #1

		Next Phi_i
	Next freq_i

	Close #2


DlgText("OutputT", "Done!")
DlgEnable("OK", True)

ReportInformationToWindow ("Farfield source successfully created in the directory: "+Global_Outfile)

End Function
'-------------------------------------------------------------------------------------
Function dialogfunc2(DlgItem$, Action%, SuppValue%) As Boolean

	Dim Extension As String, projectdir As String, filename As String
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        Select Case DlgItem
        	Case "Browseinputfile"
        		Extension = "ffd"
                projectdir = GetProjectPath("Root")
                filename = GetFilePath(,Extension, projectdir, "Browse for farfield file", 0)
                If (filename <> "") Then
                    DlgText "Infile", filename
                    DlgText "Outfile", Dirname(filename)+"\"+Basename(filename)+"_CST.ffs"
                End If
        	dialogfunc2 = True
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Private Function dialogfunc(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable("OK", False)
	Case 2 ' Value changing or button pressed
		Rem DialogFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunc = True ' Continue getting idle actions
		ExportFFDtoFFS(Global_Infile)
	Case 6 ' Function key
	End Select
End Function
