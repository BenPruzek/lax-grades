'#Language "WWB-COM"
'#include "vba_globals_all.lib"

' ================================================================================================
' This macro helps users to
' 1/ automatically import field sources into a virtual reverb chamber simulation
' 2/ calculate and display shielding effectiveness for the selected probe(s) in VRC DUT model
' ================================================================================================
' Copyright 2024-2024 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 17-Dec-2024 wxu: adjusted names of the output, spectrum file unit conversion, added detection of v2026 reference data
' 15-Dec-2024 wxu: removed import field source, integrated standard dev. calculation, renamed output per customer feedback
' 19-Nov-2024 wxu: added a small check to ensure all the results are present in 1D Result Folder
' 18-Oct-2024 wxu,djs: initial version
' ================================================================================================
Option Explicit
Public available_probes() As String, selected_probes() As String
Public sAnalyticalResultFilePath As String
Public freq() As Double, e_avg_x() As Double, e_avg_y() As Double, e_avg_z() As Double, e_avg_abs() As Double
Public nFs As Long, nSelectedProbes As Long
Public e_incoming As Double, bScale As Boolean, bPeak As Boolean

Public Type ProbeData
	probe_name As String
	file_name As String
	o1DRes() As Object 'store x,y,z,abs probe 1D data for each probe
	o1DEnv(3) As Object 'envelope of the results in x,y,z and abs
	o1DSE(3) As Object
End Type

Dim Probes() As ProbeData
Public Const sSE_Result_Root = "1D Results\Radiated Immunity\Reverberation Chamber\"

Public sRefResNames As Variant,CST_Version As Integer
Public Const default_avg_field_file = "plane_waves_spectrum_avg.dat"
Public Const detault_avg_field_label = "Average E Field without EUT"

Public Const sDir = Array("X","Y","Z","Abs")
Public Const sDelimiters = Array(Chr(9)," ", ",",";",vbCrLf) 'supported delimiters
Public debug_flag As Boolean 'optional output to indicate calculation progress

Sub Main

	debug_flag = True

	CST_Version = CInt(Mid(GetApplicationVersion,9,4)) 'obtain CST version number as an integer

	If CST_Version < 2026 Then
		Dim ref_path As String
		ref_path = sSE_Result_Root + detault_avg_field_label

		sRefResNames = Array(ref_path + "\X",ref_path + "\Y",ref_path +"\Z",ref_path + "\Abs")
	Else
		sRefResNames = Array("1D Results\Reverberation Chamber\Average\Ex","1D Results\Reverberation Chamber\Average\Ey", _
										"1D Results\Reverberation Chamber\Average\Ez","1D Results\Reverberation Chamber\Average\Etot")
		If Not Ref_Results_Exist()  Then
			MsgBox("Please execute VBA Macro ->Solver ->Sources-> Reverberation Chamber... first to make reference data available",vbCritical)
			Exit All
		End If
	End If

	Dim probe_names_var As Variant,i As Long
	probe_names_var = get_all_items_under("Probes")

	If Not IsEmpty(probe_names_var) Then
		ReDim available_probes(UBound(probe_names_var))
		For i = 0 To UBound(probe_names_var)
			available_probes(i) = cstr(probe_names_var(i))
			available_probes(i) = Split(available_probes(i),"\")(1)
		Next
	Else
		ReDim available_probes(0)
		available_probes(0) ="No probe found"
	End If

	Begin Dialog UserDialogSE 510,350,"Reverb Chamber Shielding Effectiveness Post-processing",.DiagFuncSE ' %GRID:10,7,1,1
		Text 40,14,120,21,"Select Probes for",.Text1
		OptionGroup .optionSE
			OptionButton 210,14,210,21,"Shielding Effectiveness",.oSE
			OptionButton 210,35,200,21,"Standard Deviation",.oSTDDEV
		MultiListBox 40,35,140,119,available_probes(),.avail_p_list
		Text 210,77,150,14,"Selected probes",.Text3
		TextBox 210,98,260,49,.tbSelectedProbes,2
		Text 40,168,270,14,"Reference (Average field w/o EUT)",.Text2
		OptionGroup .optionRef
			OptionButton 50,189,160,14,"External file",.oFile
			OptionButton 230,189,240,14,"Existing 1D data",.oExisting
		TextBox 40,217,350,21,.analytical_result_file
		PushButton 400,217,90,21,"Browse",.pbBrowse
		CheckBox 20,252,330,14,"Source Injection (incoming E field) Level",.cbScale
		TextBox 50,273,120,21,.tbIncoming
		Text 180,278,50,14,"V/m",.Text4
		CheckBox 20,308,190,14,"Output probe peaks",.cbOutputPeak
		OKButton 290,322,90,21
		CancelButton 400,322,90,21
	End Dialog

	Dim dlg_se As UserDialogSE
	If Not Dialog(dlg_se) Then
		Exit Sub
	End If

End Sub

Rem DialogFuncSE handles the shielding effectiveness sub-dialog
Private Function DiagFuncSE(DlgItem$, Action%, SuppValue?) As Boolean
	Dim sAnalyticalResultFileModel3D As String
	Dim nSelected As Integer, i As Integer, iSelected() As Integer

	Select Case Action%
	Case 1 ' Dialog box initialization
		If (Ref_Results_Exist()) Then	'if reference data does not exist, disable the option
			DlgValue("optionRef",0)
			DlgEnable("oExisting",False)
		End If
		'leave rest of the options disabled until user picks some probes
		DlgEnable("optionRef", False)
		DlgEnable("analytical_result_file", False)
		DlgText("analytical_result_file","")
		DlgEnable("pbBrowse", False)
		DlgEnable("cbScale", False)
		DlgEnable("tbIncoming", False)
		DlgEnable("cbOutputPeak",False)
		DlgEnable("OK", False)
		e_incoming = 3.0
		bScale = False
		bPeak = False

	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "avail_p_list"
				iSelected = DlgValue("avail_p_list")
				nSelected = UBound(iSelected)
				If nSelected = -1 Or (DlgValue("optionSE") = 1 And nSelected < 7) Then
					'Case -1 'user de-selected all probes
						DlgText("tbSelectedProbes", "Please select more probes, stand dev calc requires at least 8")
						DlgEnable("optionRef", False)
						DlgEnable("analytical_result_file", False)
						DlgText("analytical_result_file","")
						DlgEnable("pbBrowse", False)
						DlgEnable("cbScale", False)
						DlgEnable("cbOutputPeak", False)
						DlgEnable("OK", False)

				Else
					'Case Else 'user selected probes
						ReDim selected_probes (nSelected)
						For i = 0 To nSelected
							selected_probes(i) = available_probes(iSelected(i))
						Next
						Dim display_text As String
						display_text = ""
						For i = 0 To UBound(selected_probes)
							display_text += selected_probes(i)
							display_text += "; "
						Next
						DlgText("tbSelectedProbes", display_text)

						set_ref_file_field()

						If DlgValue("optionSE") = 0 Then
							DlgEnable("cbScale",True)
							DlgValue("cbScale",1)
							DlgEnable("tbIncoming", True)
							DlgText("tbIncoming",cstr(e_incoming))
						End If

				End If
				DiagFuncSE = True ' Prevent button press from closing the dialog box
			Case "optionSE"
				If DlgValue("optionSE") = 0 Then
					DlgEnable("cbScale", True)
					DlgEnable("tbIncoming", True)
				Else
					DlgEnable("cbScale", False)
					DlgEnable("tbIncoming", False)

					selected_probes = available_probes

					display_text = ""
					For i = 0 To UBound(selected_probes)
						display_text += selected_probes(i)
						display_text += "; "
					Next
					DlgText("tbSelectedProbes", display_text)

					'Select all probes by default, allowing user to overwrite
					nSelected = UBound(available_probes)
					ReDim iSelected(nSelected)
					For i = 0 To nSelected
						iSelected(i) = i
					Next
					DlgValue("avail_p_list",iSelected)
					If nSelected >= 8 Then 'need at least 8 points for standard deviation
						set_ref_file_field()
					End If
				End If
				DiagFuncSE = True ' Prevent button press from closing the dialog box
			Case "optionRef"
				If DlgValue("optionRef") = 0 Then 'user chooses external file, this means 1D results available, otherwise will default to this option

						sAnalyticalResultFileModel3D = GetProjectPath("Model3D") + default_avg_field_file
						If Dir(sAnalyticalResultFileModel3D) <> "" Then
							sAnalyticalResultFilePath = sAnalyticalResultFileModel3D
							DlgText("analytical_result_file",sAnalyticalResultFilePath)
							DlgEnable("analytical_result_file",True)
							DlgEnable("pbBrowse",True)
							DlgEnable("cbOutputPeak",True)
							DlgEnable("OK", True)
						Else
							sAnalyticalResultFilePath = ""
							DlgText("analytical_result_file",sAnalyticalResultFilePath)
							DlgEnable("analytical_result_file",True)
							DlgEnable("pbBrowse",True)
						End If

				Else 'Existing 1D
						DlgEnable("pbBrowse",False)
						If  Ref_Results_Exist() Then
							DlgValue("optionRef",1)

							'disable the file option for 2026 and beyond
							If CST_Version >=2026 Then DlgEnable("oFile", False)

							DlgText("analytical_result_file", Left(sRefResNames(0), InStrRev(sRefResNames(0),"\")))
							DlgEnable("analytical_result_file", False)
							DlgEnable("OK", True)
						Else
							DlgText("analytical_result_file", "")
							DlgEnable("OK", False)
						End If
				End If
				If DlgValue("optionSE") = 0 Then DlgEnable("cbScale", True)

			Case "pbBrowse"
					sAnalyticalResultFilePath = GetFilePath(default_avg_field_file, _
												"*.txt,*.csv", sAnalyticalResultFilePath, "Please select analytical result file", 0 )
					DlgText("analytical_result_file",sAnalyticalResultFilePath)
					If sAnalyticalResultFilePath <> "" Then
						DlgEnable("cbOutputPeak", True)
						DlgEnable("OK",True)
					Else
						DlgEnable("cbOutputPeak", False)
						DlgEnable("OK",False)
					End If

				DiagFuncSE = True ' Prevent button press from closing the dialog box
			Case "cbScale"
				If DlgValue("cbScale",True) Then
					DlgEnable("tbIncoming", True)
					DlgText("tbIncoming", cstr(e_incoming))
					bScale = True
				Else
					DlgEnable("tbIncoming", False)
					DlgText("tbIncoming","")
					bScale = False
				End If
			Case "OK"
				Select Case DlgValue("optionRef")
				Case 0 'File
					sAnalyticalResultFilePath = DlgText("analytical_result_file")

					Dim sAnalyticlFileModel3D As String, file_name_index As Integer, avg_field_tree_label As String

					avg_field_tree_label = detault_avg_field_label ' + "(" + default_avg_field_file + ")"

					file_name_index = InStrRev(sAnalyticalResultFilePath,"\") + 1
					sAnalyticlFileModel3D = GetProjectPath("Model3D") + Mid(sAnalyticalResultFilePath,file_name_index)

					If Dir(sAnalyticlFileModel3D) = "" Then 'no previous file exist
						FileCopy sAnalyticalResultFilePath, sAnalyticlFileModel3D
					ElseIf sAnalyticalResultFilePath <> sAnalyticlFileModel3D Then 'user picked a new file
						If ResultTree.DoesTreeItemExist( avg_field_tree_label ) Then 'previosu tree item exist, remove it
							With ResultTree
								.Name avg_field_tree_label
								.Delete
							End With
						End If
						Kill sAnalyticlFileModel3D 'delete the previous file in model3d
						FileCopy sAnalyticalResultFilePath, sAnalyticlFileModel3D 'copy the new file into model3d
					End If

					With ResultTree 'add to tree for later review
						.File sAnalyticlFileModel3D
						.Name avg_field_tree_label
						.Type "Notefile"
						.DeleteAt "never" ' Survive rebuilds and delete results
						.Add
					End With

					Load_Analytical_Solution_to_1DResult(sAnalyticalResultFilePath)

				Case 1 '1D Result in NT already exist
					'not loading e_avg arrays, using Result1D object directly now
				    'Load_Existing_Reference_Data()
				End Select

				bScale = IIf(DlgValue("cbScale"), True, False)

				If bScale Then  e_incoming = cDbl(DlgText("tbIncoming"))

				bPeak = IIf(DlgValue("cbOutputPeak"), True,False)

				Dim Calc_option As String

				If DlgValue("optionSE") = 0 Then
					Calc_option = "Shielding Effectiveness"
				Else
					Calc_option = "Standard Deviation"
				End If

				If Calcuate_SE_OR_STDDEV_at_Selected_Probes(Calc_option) Then
					MsgBox(Calc_option + " results are available under " + sSE_Result_Root + " folder", vbInformation)
				Else
					MsgBox(Calc_option + " calculation unsuccessful", vbCritical)
				End If

		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DiagFuncVRC = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function Calcuate_SE_OR_STDDEV_at_Selected_Probes(Calc_option As String) As Boolean
	'Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	Dim ii As Long, jj As Long, kk As Long,nPoints As Long, nPoints_ref As Long, n As Long
	Dim filename As String, dfreq As Double

	If debug_flag Then reportinformation("Getting results from 1D Results\Probes...")
	If Not Get_Probe_1D_Results() Then
		Reportinformation("Can't load probe results, please inspect 1D Results\Probes folder")
		Exit Function
	End If
	If debug_flag Then reportinformation("Done.")

	If debug_flag Then reportinformation("Calculating envelope of results...")
	If Not Get_Probe_Envelope() Then
		Reportinformation("Envelope calculation error")
		Exit Function
	End If
	If debug_flag Then reportinformation("Done.")

	If Calc_option = "Shielding Effectiveness" Then
		If debug_flag Then reportinformation("Calculating shielding effectiveness...")

		Dim e_max() As Double, dSE As Double, dFactor As Double
		Dim oSE(3) As Object, oScale(3) As Object, oFactor(3) As Object, oRef As Object

		For ii = 0 To nSelectedProbes -1 'Probe loop ii--> probe index
			For jj = 0 To 3 'Component/Direction/Polorization loop, jj-> direction index
				filename = ResultTree.GetFileFromTreeItem(sRefResNames(jj))
				Set oRef = Result1D(filename)
				Set oFactor(jj)  = Result1D("")

				nPoints_ref = oRef.GetN

				ReDim e_max(nPoints_ref -1)

				For n = 0 To nPoints_ref -1	'Frequency loop
					e_max (n) = Probes(ii).o1DEnv(jj).GetY(n)
					If  e_max(n) > 0 Then	'here asuming all e field mag data are >0, if it is 0, it is out of the range of frequency
						dfreq = oRef.GetX(n)
						'Incoming -> reference factor
						dFactor = e_incoming/oRef.GetY(n)
						oFactor(jj).AppendXY(dfreq,dFactor)
						'Shielding Effectiveness
						dSE = 20*CST_Log10(oRef.GetY(n)/e_max(n))
						Probes(ii).o1DSE(jj).AppendXY (dfreq,dSE)
					Else
						Exit For
					End If
				Next
			Next
		Next

		For ii = nSelectedProbes -1 To 0 STEP -1
			For jj = 2 To 0 STEP -1 'Plot X,Y,Z component only, exclude abs
				With Probes(ii).o1DSE(jj)
					.Type "db"
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency") )
					.SetYLabelAndUnit("S.E. " + sDir(jj) + " Component", "dB")
					.SetLogarithmicFactor(10.0)
					.Title "Shielding Effectiveness"
					.save( "e_se_" + cstr(ii) + "_" + sDir(jj) + ".sig")
					.AddtoTree(sSE_Result_Root + "Shielding Effectiveness\" + Probes(ii).probe_name + "(" + sDir(jj) + ")" )
				End With
			Next
		Next

				'scale the envelope to the new incoming e field value: (default to 3 v/m, but user can change)
		If bScale Then
			For ii = nSelectedProbes -1 To 0 STEP -1
				For jj = 3 To 0 STEP -1
						Set oScale(jj) = Probes(ii).o1DEnv(jj).copy
						oScale(jj).ComponentMult(oFactor(jj))

						With oScale(jj)
							.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
							.SetYLabelAndUnit("Probe Field at Source Injection Level " + sDir(jj) + " Component", "V/m")
							.SetLogarithmicFactor(20.0)
							.Title "Probe Field at Source Injection Level " + cstr(e_incoming)
							.save("e_env_scale_" + cstr(ii) + "_" + sDir(jj) + ".sig")
							.AddtoTree(sSE_Result_Root + "Probe Field at Source Injection Level\" + Probes(ii).probe_name + "(" + sDir(jj) + ")" )
						End With
				Next
			Next
		End If

	ElseIf Calc_option = "Standard Deviation" Then
		If debug_flag Then reportinformation("Calculating standard deviation...")

		Dim sum(2) As Double, avg(2) As Double, std_dev(2) As Double, avgT As Double, std_devT As Double
		Dim Env_Avg(3) As Object, SD(3) As Object

		For jj = 0 To 3 'Component/Direction/Polorization loop, jj-> direction index, here only 0-2 for X,Y,Z, the Total is calculated differently
			Set Env_Avg(jj) = Result1D("")
			Set SD(jj) = Result1D("")
		Next jj

		filename = ResultTree.GetFileFromTreeItem(sRefResNames(0))

		Set oRef = Result1D(filename)
		nPoints_ref = oRef.GetN

		For n = 0 To nPoints_ref -1	'Frequency loop

			avgT = 0.0
			std_devT = 0.0

			For jj = 0 To 2	'X,Y,Z Direction
				sum(jj) = 0
				avg(jj) = 0
				std_dev(jj) = 0

				For ii = 0 To nSelectedProbes -1 'Probe loop ii--> probe index
					dfreq = Probes(ii).o1DEnv(jj).GetX(n)
					sum(jj) += Probes(ii).o1DEnv(jj).GetY(n)
				Next ii

				avg(jj) = sum(jj) / nSelectedProbes
				avgT += avg(jj)

				For ii = 0 To nSelectedProbes -1 'Probe loop ii--> probe index
					std_dev(jj) += (Probes(ii).o1DEnv(jj).GetY(n) - avg(jj))^2
				Next ii

				std_dev(jj) = Sqr (std_dev(jj) / (nSelectedProbes-1) )
				std_dev(jj) = 20*Log( (std_dev(jj) + avg(jj)) / avg(jj)) / Log(10)

				Env_Avg(jj).AppendXY(dfreq, avg(jj))
				SD(jj).AppendXY(dfreq, std_dev(jj))
			Next jj

			avgT = avgT /3
			For ii = 0 To nSelectedProbes -1 'Probe loop ii--> probe index
				For jj = 0 To 2
					std_devT += (Probes(ii).o1DEnv(jj).GetY(n) - avgT)^2
				Next jj
			Next ii

			std_devT = Sqr(std_devT/(nSelectedProbes*3-1))
			std_devT = 20*Log( (std_devT + avgT) / avgT) / Log(10)
			SD(3).AppendXY(dfreq, std_devT)
	 Next n

		'write total first, so it appears last
    	With SD(3)
			.Type "db"
			.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
			.SetYLabelAndUnit "Standard Deviation", "dB"
			.Title "Standard Deviation of Probes"
			.Save getprojectbasename + "temp_sd_e_field_complex_object_" + "Total" + ".sig"
			.AddToTree sSE_Result_Root + "Standard Deviation of Probes\Total"
		End With

		For jj = 2 To 0 STEP -1
        	With SD(jj)
				.Type "db"
				.SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
				.SetYLabelAndUnit "Standard Deviation", "dB"
				.Title "Standard Deviation of Probes"
				.Save getprojectbasename + "temp_sd_e_field_complex_object_" + sDir(jj) + ".sig"
				.AddToTree sSE_Result_Root + "Standard Deviation of Probes\" + sDir(jj)
			End With

			If bPeak Then
	            With Env_Avg(jj)
	              .SetXLabelAndUnit "Frequency" , Units.GetUnit("Frequency")
	              .SetYLabelAndUnit "Average Peak E Field" , "V/m"
				  .SetLogarithmicFactor(20.0)
	              .Title "Average Probe Peaks"
	              .Save getprojectbasename + "temp_avg_e_field_complex_object_" + sDir(jj) + ".sig"
	              .AddToTree sSE_Result_Root + "Probe Peaks\Average " + sDir(jj)
		       	End With
			End If
		Next

	End If

	If bPeak Then
		For ii = nSelectedProbes-1 To 0 STEP -1
			For jj = 3 To 0 STEP -1
				With Probes(ii).o1DEnv(jj)
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
					.SetYLabelAndUnit("E field peak " + sDir(jj) + " Component" ,"V/m")
					.SetLogarithmicFactor(20.0)
					.Title "Max probe values in all MC"
					.save("e_env_" + cstr(ii) + "_" + sDir(jj) + ".sig")
					.AddtoTree(sSE_Result_Root + "Probe Peaks\" + Probes(ii).probe_name + "(" + sDir(jj) + ")")
				End With
			Next
		Next
	End If

	If debug_flag Then reportinformation("Done.")

	Calcuate_SE_OR_STDDEV_at_Selected_Probes = True
End Function


Function Get_Probe_1D_Results() As Boolean
	Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	Dim excitation_names As Variant
	Dim ii As Long, jj As Long, kk As Long

	excitation_names = get_all_items_under("Field Sources")
	If IsEmpty(excitation_names) Then
		reportinformation("No field source excitation found, please import vrc field sources first")
		Exit Function
	Else
		nFs = UBound(excitation_names) + 1 'nFS: number of Field sources, global variable,obtained here
		For ii = 0 To nFs -1
			excitation_names(ii) = "[" + Split(excitation_names(ii),"\")(1) + "]"
		Next
	End If

	nSelectedProbes = UBound(selected_probes) + 1 'nSelectedProbes is a global variable and obtained here

	If nSelectedProbes <=0 Then
		reportinformation ("No probe selected, existing...")
		Get_Probe_1D_Results = False
		Exit Function
	End If

	ReDim Probes(nSelectedProbes -1)


	For ii = 0 To nSelectedProbes -1
		Probes(ii).probe_name = selected_probes(ii)
		ReDim Probes(ii).o1DRes(3,nFs)

		For jj = 0 To 3
			Set Probes(ii).o1DSE(jj) = Result1D("")
			For kk = 0 To nFs - 1
				Set Probes(ii).o1DRes(jj,kk) = Result1D("")
			Next kk
		Next jj
	Next ii

	nResults = ResultTree.GetTreeResults("1D Results\Probes\E-Field","0D/1D","",paths,types,files,info)

	If nResults = 0 Then
		Reportinformation("No resulst from selected probe found")
		Get_Probe_1D_Results = 0
		Exit Function
	End If

	'organizing the results and store the filenames to probedata structure
	'Public Type ProbeData
	'probe_name As String
	'file_name As String
	'o1DRes() As Object 'store x,y,z,abs probe 1D data for each probe
	'o1DEnv(3) As Object 'envelope of the results in x,y,z and abs
	'o1DSE(3) As Object
	'End Type
	'It is expected that each probe contains 4 components (X,Y,Z,Abs) for each [excitation]

	For ii = 0 To nResults-1
		For jj = 0 To nSelectedProbes - 1
			If  InStr(paths(ii),Probes(jj).probe_name) Then
				If InStr(paths(ii),"(X)") Then
					For kk = 0 To nFs -1
						If InStr(paths(ii),excitation_names(kk)) Then Set Probes(jj).o1DRes(0,kk) = Result1DComplex(files(ii)).Magnitude()
					Next
				ElseIf InStr(paths(ii), "(Y)") Then
					For kk = 0 To nFs -1
						If InStr(paths(ii),excitation_names(kk)) Then Set Probes(jj).o1DRes(1,kk) = Result1DComplex(files(ii)).Magnitude()
					Next
				ElseIf InStr(paths(ii), "(Z)") Then
					For kk = 0 To nFs -1
						If InStr(paths(ii),excitation_names(kk)) Then Set Probes(jj).o1DRes(2,kk) = Result1DComplex(files(ii)).Magnitude()
					Next
				ElseIf InStr(paths(ii), "(Abs)") Then
					For kk = 0 To nFs -1
						If InStr(paths(ii),excitation_names(kk)) Then Set Probes(jj).o1DRes(3,kk) = Result1DComplex(files(ii)).Magnitude()
					Next
				End If
			End If
		Next
	Next
	Get_Probe_1D_Results = True
End Function
Function Get_Probe_Envelope() As Boolean
	Dim filename As String, n As Long,e_max() As Double, dfreq As Double,dSE As Double, dFactor As Double
	Dim min_x As Double, max_x As Double, x_ref As Double
	Dim nPoints As Long, nPoints_ref As Long
	Dim oRef As Object
	Dim ii As Long, jj As Long, kk As Long

	For ii = 0 To nSelectedProbes -1 'Probe loop ii--> probe index	'nSelectedProbes is obtained from Get_Probe_1D_Results()
		For jj = 0 To 3 'Component/Direction/Polorization loop, jj-> direction index
			filename = ResultTree.GetFileFromTreeItem(sRefResNames(jj))
			Set oRef = Result1D(filename)
			Set Probes(ii).o1DEnv(jj) = Result1D("")

			nPoints_ref = oRef.GetN

			For kk = 0 To nFs -1	'Field Source/Excitation loop, 'nFS is obtained from Get_Probe_1D_Results()
				'dealing with end points
				nPoints = Probes(ii).o1DRes(jj,kk).GetN

				If nPoints = 0 Then 'added a check 11/19/2024 to warn user if the data are not complete
					Reportinformation("Missing fs=" + cstr(kk + 1) + " direction =" + sDir(jj) + " Probe =" + Probes(ii).probe_name)
					MsgBox("Simulation results are not complete, expecting frequency results from all field sources and x,y,z,abs components!" + vbLf + _
							"Please check 1D Results\Probes\E-Field folder",vbCritical)
					Get_Probe_Envelope = False
					Exit Function
				End If

				' first and last points of frequency
				min_x = Probes(ii).o1DRes(jj,kk).GetX(0)
				max_x = Probes(ii).o1DRes(jj,kk).GetX(nPoints - 1)

				' if the first frequency is greater than ref, the MakeCompatibleTo method will set the Y value at index 0 to be 0
				' here we test if the first and last frequencies are within 0.001 of a reference point, if yes, set X(0), to be ref_X(point), and X(n_point) to be X_ref(point)

				Dim i_ref As Long
				'looping through all reference data points, and see if any point are close enough to the first and last X of the probe frequency
				For i_ref = 0 To nPoints_ref - 1

					x_ref = oRef.GetX(i_ref)
					'if x_ref is teeny bit more than x_ref
					If min_x > x_ref And min_x - x_ref < 1E-3 Then	'snap to the ref point if it is close enough, preventing the x_ref point goes to 0
						Probes(ii).o1DRes(jj,kk).setX(0,x_ref)
					End If
						'if x_ref is teeny bit less than x_ref
					If max_x < x_ref And x_ref - max_x < 1E-3 Then  'snap to the ref point if it is close enough, preventing the x_ref point goes to 0
						Probes(ii).o1DRes(jj,kk).setX(nPoints - 1, x_ref)
					End If
				Next

				'after MakeCompatibleTo method is called, the number of points becomes the same between data and reference

				Probes(ii).o1DRes(jj,kk).MakeCompatibleTo(oRef) 'map the result to reference freq. returns 0 if the freq of reference > data freq
			Next

			ReDim e_max(nPoints_ref -1)

			For n = 0 To nPoints_ref -1	'Frequency loop
				For kk = 0 To nFs -1
					If Probes(ii).o1DRes(jj,kk).GetY(n) > e_max(n) Then e_max(n) = Probes(ii).o1DRes(jj,kk).GetY(n)
				Next
				If e_max(n) > 0 Then	'here asuming all e field mag data are >0, if it is 0, it is out of the range of frequency
					dfreq = oRef.GetX(n)
					'Envelope
					Probes(ii).o1DEnv(jj).AppendXY (dfreq,e_max(n))
				Else
					Exit For
				End If
			Next
		Next
	Next
	Get_Probe_Envelope = True
End Function


Function Load_Analytical_Solution_to_1DResult(sAnalyticalResultFilePath)

				Read_Analytical_Result_File(sAnalyticalResultFilePath)

				Dim oAnalytical1D_X As Object, oAnalytical1D_Y As Object, oAnalytical1D_Z As Object
				Dim oAnalytical1D_Abs As Object, ii As Long
				Set oAnalytical1D_X = Result1D("")
				Set oAnalytical1D_Y = Result1D("")
				Set oAnalytical1D_Z = Result1D("")
				Set oAnalytical1D_Abs = Result1D("")

				For ii = 0 To UBound(freq)
					oAnalytical1D_X.AppendXY(freq(ii),e_avg_x(ii))
					oAnalytical1D_Y.AppendXY(freq(ii),e_avg_y(ii))
					oAnalytical1D_Z.AppendXY(freq(ii),e_avg_z(ii))
					oAnalytical1D_Abs.AppendXY(freq(ii),e_avg_abs(ii))
				Next

				With oAnalytical1D_Abs
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
					.SetYLabelAndUnit("Average E field Abs", "V/m")
					.SetLogarithmicFactor(20.0)
					.title("Average E Field w/o EUT")
					.save("e_avg_abs.sig")
					.AddtoTree(sRefResNames(3))
					.DeleteAt("never")
				End With

				With oAnalytical1D_Z
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
					.SetYLabelAndUnit("Average E field Z Component", "V/m")
					.SetLogarithmicFactor(20.0)
					.title("Average E Field w/o EUT")
					.save("e_avg_z.sig")
					.AddtoTree(sRefResNames(2))
					.DeleteAt("never")
				End With

				With oAnalytical1D_Y
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
					.SetYLabelAndUnit("Average E field Y Component","V/m")
					.SetLogarithmicFactor(20.0)
					.title("Average E Field w/o EUT")
					.save("e_avg_y.sig")
					.AddtoTree(sRefResNames(1))
					.DeleteAt("never")
				End With

				With oAnalytical1D_X
					.SetXLabelAndUnit("Frequency" , Units.GetUnit("Frequency"))
					.SetYLabelAndUnit("Average E field X Component", "V/m")
					.SetLogarithmicFactor(20.0)
					.title("Average E Field w/o EUT")
					.save("e_avg_x.sig")
					.AddtoTree(sRefResNames(0))
					.DeleteAt("never")
				End With

End Function


Function Read_Analytical_Result_File(sFileName As String)
	Dim sFileContents As String, sFileContentsArray() As String

	sFileContents = TextFileToString_LIB(sFileName)
	sFileContentsArray() = Split(sFileContents, vbNewLine)

	If UBound(sFileContentsArray) = 0 Then ' Could be a Linux file, try LF as separator
		sFileContentsArray() = Split(sFileContents, Chr(10))
	End If

	Dim ii As Long, numPoints As Long, current_line() As String

	For ii = UBound(sFileContentsArray) To 0 STEP -1
		If sFileContentsArray(ii) <> "" Then Exit For
	Next
	numPoints = ii

	ReDim freq(numPoints), e_avg_x(numPoints),  e_avg_y(numPoints), e_avg_z(numPoints), e_avg_abs(numPoints)
	For ii = 0 To numPoints
		current_line = mySplit(sFileContentsArray(ii))
		If UBound(current_line) = 4 Then
			' the spectrum average file will be using GHz as unit, need to be converted to the current project unit
			freq(ii) = cDbl(current_line(0)) * 1e9 /Units.GetFrequencyUnitToSI
			e_avg_x(ii) = cDbl(current_line(1))
			e_avg_y(ii) = cDbl(current_line(2))
			e_avg_z(ii) = cDbl(current_line(3))
			e_avg_abs(ii) = cDbl(current_line(4)) 'this should be the magnitude column
		Else
			MsgBox("Unexpected analytical file format, please contact CST", vbCritical)
			Exit All
		End If
	Next
End Function

Function Ref_Results_Exist() As Boolean
	Dim res_name As String
	For Each res_name In sRefResNames
		If Not ResultTree.DoesTreeItemExist(res_name) Then
			Ref_Results_Exist = False
			Exit Function
		End If
	Next
	Ref_Results_Exist = True
End Function

Function get_all_items_under(sRoot) As Variant
	Dim existing_item As String, item_array() As Variant, i As Long
	i = -1
	existing_item = ResultTree.GetFirstChildName(sRoot)
	While existing_item <> ""
		i += 1
		ReDim Preserve item_array(i)
		item_array(i) = existing_item
		existing_item = ResultTree.GetNextItemName(existing_item)
	Wend
get_all_items_under = item_array
End Function

Function find_all_files_with_extension(path As String, ext As String) As Variant
	Dim fname As String, farray As Variant, i As Long
	i = -1
	fname = Dir(path + ext)
	While fname <> ""
		i += 1
		ReDim Preserve farray(i)
		farray(i) = path + fname
		fname = Dir()
	Wend

find_all_files_with_extension = farray

End Function
Function mySplit(sCodeline As String) As String()
	Dim ii As Integer, lastNonEmpty As Integer
	Dim sLineArray() As String, Sep As String

	'Replace all supported delimiters with tab
	For Each Sep In sDelimiters
		sCodeline = Replace(sCodeline,Sep,Chr(9))
	Next

	sLineArray = Split(sCodeline,Chr(9))

	lastNonEmpty = -1
	For ii = 0 To UBound(sLineArray)
		If sLineArray(ii) <> "" Then
			lastNonEmpty  = lastNonEmpty + 1
			sLineArray(lastNonEmpty) = sLineArray(ii)
		End If
	Next

	If lastNonEmpty > -1 Then ReDim Preserve sLineArray(lastNonEmpty)

	Return sLineArray
End Function

Function set_ref_file_field() As Boolean
	Dim sAnalyticalResultFileModel3D As String, display_name As String

	'set the default reference options
	If  Ref_Results_Exist() Then
		DlgEnable("optionRef", True)
		DlgValue("optionRef", 1)
		DlgText("analytical_result_file", Left(sRefResNames(0), InStrRev(sRefResNames(0),"\")))
		DlgEnable("oExisting",True)
		DlgEnable("oFile", True)
		DlgEnable("cbOutputPeak",True)
		DlgEnable("OK", True)

	Else
		DlgValue("optionRef", 0)
		DlgEnable("oFile",True)
		sAnalyticalResultFileModel3D = GetProjectPath("Model3D") + default_avg_field_file
		If Dir(sAnalyticalResultFileModel3D) <> "" Then
			sAnalyticalResultFilePath = sAnalyticalResultFileModel3D
			DlgText("analytical_result_file",sAnalyticalResultFilePath)
			DlgEnable("analytical_result_file",True)
			DlgEnable("pbBrowse",True)
			DlgEnable("cbOutputPeak",True)
			DlgEnable("OK", True)
		Else
			sAnalyticalResultFilePath = ""
			DlgText("analytical_result_file",sAnalyticalResultFilePath)
			DlgEnable("analytical_result_file",True)
			DlgEnable("pbBrowse",True)
		End If
	End If
End Function
