' ================================================================================================
' Copyright 2022-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 08-Jan-2024 vpn: adapt to introduce settings for 1D monitors
' 23-May-2022 vpn: First version
' ================================================================================================

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

''''''''''''''''''''''''''' define used types
' define type for antenna-element info
Type AntElem
	ThetaX					As Double
	ThetaY					As Double
	ThetaZ					As Double
	PhiX					As Double
	PhiY					As Double
	PhiZ					As Double
	OrigX					As Double
	OrigY					As Double
	OrigZ					As Double
	Amplitude				As Double
	Phase					As Double
	Noise					As Double
	IsNoiseInDBm			As Boolean
	FileName				As String
End Type

' define the type of curve
Type CurveSettings
	CurveName				As String
	Velocity				As Double
	SamplingStep			As Double
	OnlyComputeCornerPoint	As Boolean
	NumElem					As Integer
	Active					As Boolean
	Elem()					As AntElem
End Type

''''''''''''''''''''''''''' define global variables
Dim id, description As String
Dim elem_settings_intermediate() As AntElem
Dim elem_settings() As AntElem
Dim curve_settings() As CurveSettings
Dim max_curve_name_length As Integer

''''''''''''''''''''''''''' define functions
' pad string
Function pad_string(str_source As String, pad_len As Long, Optional pad_char As String = " ") As String
	If (pad_len > Len(str_source)) Then
		pad_len = pad_len - Len(str_source)
	Else
		pad_len = 1
	End If
	pad_string = String(pad_len, pad_char)
	pad_string = str_source + pad_string
End Function

' assign antenna element default value
Function assign_antelem_default_value(ByRef Elem() As AntElem)
	ReDim Elem(0)
	Elem(0).ThetaX			= 1.0
	Elem(0).ThetaY			= 0.0
	Elem(0).ThetaZ			= 0.0
	Elem(0).PhiX			= 0.0
	Elem(0).PhiY			= 0.0
	Elem(0).PhiZ			= 1.0
	Elem(0).OrigX			= 0.0
	Elem(0).OrigY			= 0.0
	Elem(0).OrigZ			= 0.0
	Elem(0).Amplitude		= 1.0
	Elem(0).Phase			= 0.0
	Elem(0).Noise			= 0.0
	Elem(0).IsNoiseInDBm	= True
	Elem(0).FileName		= ""
End Function

' function to convert to string
Function toStr(value As String)
	If IsNumeric(value) Then
		Dim x As Double
		x = CDbl(value)
		toStr = Format(x, "0.###")
	Else
		toStr = value
	End If
End Function

' write curve setting
Function write_curve_settings_to_history()
	Dim caption1 As String
	Dim command1 As String
	Dim sIsActive As String
	Dim sIsCompNativ As String
	caption1 = "asymptotic solver: add 1D monitors"
	command1 = "With AsymptoticSolver" + vbCrLf
	command1 = command1 + "     .Reset1DMonList """ + CStr(UBound(curve_settings) + 1) + """" + vbCrLf
	' add to command
	For I=0 To UBound(curve_settings)
		If (curve_settings(I).Active) Then
			sIsActive = "True"
		Else
			sIsActive = "False"
		End If
		If (curve_settings(I).OnlyComputeCornerPoint) Then
			sIsCompNativ = "True"
		Else
			sIsCompNativ = "False"
		End If
		command1 = command1 + "     .Add1DMon """ + curve_settings(I).CurveName + _
		""", """ + sIsActive + _
		""", """ + sIsCompNativ + _
		""", """ + CStr(UBound(curve_settings(I).Elem) + 1) + _
		""", """ + CStr(curve_settings(I).SamplingStep) + _
		""", """ + CStr(curve_settings(I).Velocity) + """ " + vbCrLf
		' elements
		For J=0 To UBound(curve_settings(I).Elem)
			Dim sIsNoiseInDBm As String
			Dim sFFSName As String
			If (curve_settings(I).Elem(J).IsNoiseInDBm) Then
				sIsNoiseInDBm = "True"	' True
			Else
				sIsNoiseInDBm = "False"	' False
			End If
			command1 = command1 + "     .Add1DMonElem """ + curve_settings(I).CurveName + _
			""", """ + CStr(curve_settings(I).Elem(J).OrigX) + _
			""", """ + CStr(curve_settings(I).Elem(J).OrigY) + _
			""", """ + CStr(curve_settings(I).Elem(J).OrigZ) + _
			""", """ + CStr(curve_settings(I).Elem(J).ThetaX) + _
			""", """ + CStr(curve_settings(I).Elem(J).ThetaY) + _
			""", """ + CStr(curve_settings(I).Elem(J).ThetaZ) + _
			""", """ + CStr(curve_settings(I).Elem(J).PhiX) + _
			""", """ + CStr(curve_settings(I).Elem(J).PhiY) + _
			""", """ + CStr(curve_settings(I).Elem(J).PhiZ) + _
			""", """ + CStr(curve_settings(I).Elem(J).Amplitude) + _
			""", """ + CStr(curve_settings(I).Elem(J).Phase) + _
			""", """ + CStr(curve_settings(I).Elem(J).Noise) + _
			""", """ + sIsNoiseInDBm + _
			""", """ + curve_settings(I).Elem(J).FileName + _
			""", """ + "True" + _
			""", """ + "" + """ " + vbCrLf
		Next
	Next
	command1 = command1 + "End With"
	AddToHistory caption1, command1
End Function

' Obtain the curve information
' 1. Read from the text file
' 2. If the text file is not available, add a default value for each curve
Function obtain_curve_info(ByRef is_any_curve As Boolean, ByRef max_curve_name_length As Integer)
	Dim ssc As String
	Dim curveList(1001) As String
	Dim nIndex As Integer
	is_any_curve = False
	max_curve_name_length = 0

	'Get list of curves (here also include closed and unconnected segments)
	ssc = ResultTree.GetFirstChildName ("Curves")  ' 6 characters long
	Dim icount As Integer
	icount = 0
	While ssc <> ""
		curveList(icount)=Right(ssc, Len(ssc)-7)
		icount = icount + 1
		ssc=ResultTree.GetNextItemName (ssc)
	Wend

	' Fill the info for curve_settings
	If (icount > 0) Then
		is_any_curve = True
		Dim name_length As Integer
		ReDim curve_settings(icount-1)
		For I=0 To icount-1
			name_length = Len(curveList(I))
			max_curve_name_length = Max(max_curve_name_length, name_length)
			curve_settings(I).CurveName 				= curveList(I)
			nIndex = AsymptoticSolver.Get1DMonIndexFromName(curve_settings(I).CurveName)
			If (nIndex) < 0 Then
				' assign default values
				curve_settings(I).Velocity 					= 0
				curve_settings(I).SamplingStep 				= 0.5
				curve_settings(I).OnlyComputeCornerPoint	= False
				curve_settings(I).NumElem					= 1
				curve_settings(I).Active					= True
				assign_antelem_default_value(curve_settings(I).Elem)
			Else
				' get data from asymptotic solver
				Dim monitor_info_array() As String
				monitor_info_array = AsymptoticSolver.GetFromIndex1DMonInfo(nIndex)
				If (CInt(monitor_info_array(1)) = 1) Then
					curve_settings(I).Active = True
				Else
					curve_settings(I).Active = False
				End If
				If (CInt(monitor_info_array(3)) = 1) Then
					curve_settings(I).OnlyComputeCornerPoint = True
				Else
					curve_settings(I).OnlyComputeCornerPoint = False
				End If
				curve_settings(I).SamplingStep 				= CDbl(monitor_info_array(4))
				curve_settings(I).Velocity 					= CDbl(monitor_info_array(5))
				curve_settings(I).NumElem					= CInt(monitor_info_array(6))
				ReDim curve_settings(I).Elem(curve_settings(I).NumElem - 1)
				For J=0 To UBound(curve_settings(I).Elem)
					Dim monitor_elem_info_array() As String
					monitor_elem_info_array = AsymptoticSolver.GetFromIndex1DMonElemInfo(nIndex, J)
					curve_settings(I).Elem(J).OrigX 	= CDbl(monitor_elem_info_array(0))
					curve_settings(I).Elem(J).OrigY 	= CDbl(monitor_elem_info_array(1))
					curve_settings(I).Elem(J).OrigZ 	= CDbl(monitor_elem_info_array(2))
					curve_settings(I).Elem(J).ThetaX 	= CDbl(monitor_elem_info_array(3))
					curve_settings(I).Elem(J).ThetaY 	= CDbl(monitor_elem_info_array(4))
					curve_settings(I).Elem(J).ThetaZ 	= CDbl(monitor_elem_info_array(5))
					curve_settings(I).Elem(J).PhiX 		= CDbl(monitor_elem_info_array(6))
					curve_settings(I).Elem(J).PhiY 		= CDbl(monitor_elem_info_array(7))
					curve_settings(I).Elem(J).PhiZ 		= CDbl(monitor_elem_info_array(8))
					curve_settings(I).Elem(J).Amplitude = CDbl(monitor_elem_info_array(9))
					curve_settings(I).Elem(J).Phase 	= CDbl(monitor_elem_info_array(10))
					curve_settings(I).Elem(J).Noise 	= CDbl(monitor_elem_info_array(11))
					If (CInt(monitor_elem_info_array(12)) = 1) Then
						curve_settings(I).Elem(J).IsNoiseInDBm = True
					Else
						curve_settings(I).Elem(J).IsNoiseInDBm = False
					End If
					curve_settings(I).Elem(J).FileName 	= monitor_elem_info_array(13)
				Next
			End If
		Next
	End If
End Function

Function check_value()
	Dim string_ThetaX 		As String
	Dim string_ThetaY 		As String
	Dim string_ThetaZ 		As String
	Dim string_PhiX 		As String
	Dim string_PhiY 		As String
	Dim string_PhiZ 		As String
	Dim string_OrigX 		As String
	Dim string_OrigY 		As String
	Dim string_OrigZ 		As String
	Dim string_Amplitude 	As String
	Dim string_Phase		As String
	Dim string_Noise		As String
	Dim FileName			As String
	Dim ThetaX 				As Double
	Dim ThetaY 				As Double
	Dim ThetaZ 				As Double
	Dim PhiX 				As Double
	Dim PhiY 				As Double
	Dim PhiZ 				As Double
	Dim OrigX 				As Double
	Dim OrigY 				As Double
	Dim OrigZ 				As Double
	Dim Phase				As Double
	Dim Noise				As Double
	Dim NoiseInDBm			As Integer
	string_ThetaX 	= DlgText$("Editz1")
	string_ThetaY 	= DlgText$("Editz2")
	string_ThetaZ 	= DlgText$("Editz3")
	string_PhiX 	= DlgText$("Editx1")
	string_PhiY		= DlgText$("Editx2")
	string_PhiZ		= DlgText$("Editx3")
	string_OrigX	= DlgText$("EditOrg_x")
	string_OrigY	= DlgText$("EditOrg_y")
	string_OrigZ	= DlgText$("EditOrg_z")
	string_Amplitude= DlgText$("EditAmpl")
	string_Phase	= DlgText$("EditPhase")
	string_Noise	= DlgText$("EditNoise")
	FileName		= DlgText$("EditFilename")
	NoiseInDBm		= DlgValue("EditIsNoiseInDBm")
	OrigX = CDbl(string_OrigX)
	OrigY = CDbl(string_OrigY)
	OrigZ = CDbl(string_OrigZ)
	ThetaX = CDbl(string_ThetaX)
	ThetaY = CDbl(string_ThetaY)
	ThetaZ = CDbl(string_ThetaZ)
	PhiX = CDbl(string_PhiX)
	PhiY = CDbl(string_PhiY)
	PhiZ = CDbl(string_PhiZ)
	If (Not IsNumeric(string_OrigX)) Then
		MsgBox "Incorrect value for origin-x", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_OrigY)) Then
		MsgBox "Incorrect value for origin-Y", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_OrigZ)) Then
		MsgBox "Incorrect value for origin-Z", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_ThetaX)) Then
		MsgBox "Incorrect value for theta-x", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_ThetaY)) Then
		MsgBox "Incorrect value for theta-y", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_ThetaZ)) Then
		MsgBox "Incorrect value for theta-z", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_PhiX)) Then
		MsgBox "Incorrect value for phi-x", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_PhiY)) Then
		MsgBox "Incorrect value for phi-y", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_PhiZ)) Then
		MsgBox "Incorrect value for phi-z", vbOkOnly
		check_value = False
	ElseIf (ThetaX*PhiX + ThetaY*PhiY + ThetaZ*PhiZ > 0.000001) Then
		MsgBox "Incorrect value of theta and phi orientation: theta and phi are not orthogonal ", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_Amplitude)) Then
		MsgBox "Incorrect value for amplitude", vbOkOnly
		check_value = False
	ElseIf(CDbl(string_Amplitude) < 0.0) Then
		MsgBox "Incorrect value for amplitude", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_Phase)) Then
		MsgBox "Incorrect value for phase", vbOkOnly
		check_value = False
	ElseIf (Not IsNumeric(string_Noise)) Then
		MsgBox "Incorrect value for noise", vbOkOnly
		check_value = False
	Else
		Noise = CDbl(string_Noise)
		If (Noise < 0.0 And NoiseInDBm = 0) Then
			MsgBox "Incorrect value for Noise", vbOkOnly
			check_value = False
		Else
			check_value = True
		End If
	End If
End Function

' function for antenna elements dialog
Private Function antenna_element_dialog_function(DlgItem$, Action%, SuppValue&) As Boolean
	Dim list_array() As String
	Dim elem_settings_tmp() As AntElem
	Dim index_ As Integer
	Dim iselected As Integer

	If (DlgItem = "elem_button_ok") Then
		' OK button pressed: update value from elem_settings_intermediate to elem_settings
		ReDim elem_settings(UBound(elem_settings_intermediate))
		For I=0 To UBound(elem_settings)
			elem_settings(I).ThetaX			= elem_settings_intermediate(I).ThetaX
			elem_settings(I).ThetaY			= elem_settings_intermediate(I).ThetaY
			elem_settings(I).ThetaZ			= elem_settings_intermediate(I).ThetaZ
			elem_settings(I).PhiX			= elem_settings_intermediate(I).PhiX
			elem_settings(I).PhiY			= elem_settings_intermediate(I).PhiY
			elem_settings(I).PhiZ			= elem_settings_intermediate(I).PhiZ
			elem_settings(I).OrigX			= elem_settings_intermediate(I).OrigX
			elem_settings(I).OrigY			= elem_settings_intermediate(I).OrigY
			elem_settings(I).OrigZ			= elem_settings_intermediate(I).OrigZ
			elem_settings(I).Amplitude		= elem_settings_intermediate(I).Amplitude
			elem_settings(I).Phase			= elem_settings_intermediate(I).Phase
			elem_settings(I).Noise 			= elem_settings_intermediate(I).Noise
			elem_settings(I).IsNoiseInDBm	= elem_settings_intermediate(I).IsNoiseInDBm
			elem_settings(I).FileName		= elem_settings_intermediate(I).FileName
		Next
	End If
	Select Case Action%
	Case 1
	Case 2
		If (DlgItem = "BrowseFile") Then
			' keep dialog open
			antenna_element_dialog_function = True
			' select file
			Dim extension As String
			extension = "ffs"
			Dim sfilename As String
			sfilename = GetFilePath("", extension, GetProjectPath("Result"), , 0)
			Dim ilen As Integer
			ilen = Len(GetProjectPath("Project"))
			If LCase(Left(sfilename,ilen)) = LCase(GetProjectPath("Project")) Then
				sfilename = Mid(sfilename, ilen+2)
			End If
			DlgText "EditFilename", sfilename
		ElseIf (DlgItem = "button_add") Then
			' keep dialog open
			antenna_element_dialog_function = True
			' add : check value + update line (update to elem_settings_intermediate)
			If (check_value()) Then
				' create temporary variable to store the new line
				ReDim elem_settings_tmp(UBound(elem_settings_intermediate)+1)
				For I=0 To UBound(elem_settings_intermediate)
					elem_settings_tmp(I).ThetaX 		= elem_settings_intermediate(I).ThetaX
					elem_settings_tmp(I).ThetaY 		= elem_settings_intermediate(I).ThetaY
					elem_settings_tmp(I).ThetaZ 		= elem_settings_intermediate(I).ThetaZ
					elem_settings_tmp(I).PhiX 			= elem_settings_intermediate(I).PhiX
					elem_settings_tmp(I).PhiY 			= elem_settings_intermediate(I).PhiY
					elem_settings_tmp(I).PhiZ 			= elem_settings_intermediate(I).PhiZ
					elem_settings_tmp(I).OrigX 			= elem_settings_intermediate(I).OrigX
					elem_settings_tmp(I).OrigY 			= elem_settings_intermediate(I).OrigY
					elem_settings_tmp(I).OrigZ 			= elem_settings_intermediate(I).OrigZ
					elem_settings_tmp(I).Amplitude 		= elem_settings_intermediate(I).Amplitude
					elem_settings_tmp(I).Phase 			= elem_settings_intermediate(I).Phase
					elem_settings_tmp(I).Noise 			= elem_settings_intermediate(I).Noise
					elem_settings_tmp(I).IsNoiseInDBm	= elem_settings_intermediate(I).IsNoiseInDBm
					elem_settings_tmp(I).FileName 		= elem_settings_intermediate(I).FileName
				Next
				Dim NoiseInDBm			As Integer
				index_ = UBound(elem_settings_intermediate)+1
				elem_settings_tmp(index_).ThetaX 		= CDbl(DlgText$("Editz1"))
				elem_settings_tmp(index_).ThetaY 		= CDbl(DlgText$("Editz2"))
				elem_settings_tmp(index_).ThetaZ 		= CDbl(DlgText$("Editz3"))
				elem_settings_tmp(index_).PhiX 			= CDbl(DlgText$("Editx1"))
				elem_settings_tmp(index_).PhiY 			= CDbl(DlgText$("Editx2"))
				elem_settings_tmp(index_).PhiZ 			= CDbl(DlgText$("Editx3"))
				elem_settings_tmp(index_).OrigX 		= CDbl(DlgText$("EditOrg_x"))
				elem_settings_tmp(index_).OrigY 		= CDbl(DlgText$("EditOrg_y"))
				elem_settings_tmp(index_).OrigZ 		= CDbl(DlgText$("EditOrg_z"))
				elem_settings_tmp(index_).Amplitude 	= CDbl(DlgText$("EditAmpl"))
				elem_settings_tmp(index_).Phase 		= CDbl(DlgText$("EditPhase"))
				elem_settings_tmp(index_).Noise 		= CDbl(DlgText$("EditNoise"))
				NoiseInDBm		= DlgValue("EditIsNoiseInDBm")
				If (NoiseInDBm = 0) Then
					elem_settings_tmp(index_).IsNoiseInDBm = False
				Else
					elem_settings_tmp(index_).IsNoiseInDBm = True
				End If
				elem_settings_tmp(index_).FileName 	= DlgText("EditFilename")
				' pass value back to elem_settings_intermediate
				ReDim elem_settings_intermediate(UBound(elem_settings_tmp))
				For I=0 To UBound(elem_settings_tmp)
					elem_settings_intermediate(I).ThetaX = elem_settings_tmp(I).ThetaX
					elem_settings_intermediate(I).ThetaY = elem_settings_tmp(I).ThetaY
					elem_settings_intermediate(I).ThetaZ = elem_settings_tmp(I).ThetaZ
					elem_settings_intermediate(I).PhiX = elem_settings_tmp(I).PhiX
					elem_settings_intermediate(I).PhiY = elem_settings_tmp(I).PhiY
					elem_settings_intermediate(I).PhiZ = elem_settings_tmp(I).PhiZ
					elem_settings_intermediate(I).OrigX = elem_settings_tmp(I).OrigX
					elem_settings_intermediate(I).OrigY = elem_settings_tmp(I).OrigY
					elem_settings_intermediate(I).OrigZ = elem_settings_tmp(I).OrigZ
					elem_settings_intermediate(I).Amplitude = elem_settings_tmp(I).Amplitude
					elem_settings_intermediate(I).Phase = elem_settings_tmp(I).Phase
					elem_settings_intermediate(I).Noise = elem_settings_tmp(I).Noise
					elem_settings_intermediate(I).IsNoiseInDBm = elem_settings_tmp(I).IsNoiseInDBm
					elem_settings_intermediate(I).FileName = elem_settings_tmp(I).FileName
				Next
				' update the list
				ReDim list_array(UBound(elem_settings_intermediate))
				For I=0 To UBound(list_array)
					list_array(I) = get_antelem_line(elem_settings_intermediate(I))
				Next
				DlgListBoxArray "elem_list_box", list_array
			End If
		ElseIf (DlgItem = "button_remove") Then
			' keep dialog open
			antenna_element_dialog_function = True
			' remove: if only one element left, not allowed to delete
			If (UBound(elem_settings_intermediate) = 0) Then
				MsgBox "At least one element should be available", vbOkOnly
			Else
				ReDim elem_settings_tmp(UBound(elem_settings_intermediate)-1)
				index_ = 0
				iselected = DlgValue("elem_list_box")
				For I=0 To UBound(elem_settings_intermediate)
					If (iselected <> I) Then
						elem_settings_tmp(index_).ThetaX 		= elem_settings_intermediate(I).ThetaX
						elem_settings_tmp(index_).ThetaY 		= elem_settings_intermediate(I).ThetaY
						elem_settings_tmp(index_).ThetaZ 		= elem_settings_intermediate(I).ThetaZ
						elem_settings_tmp(index_).PhiX 			= elem_settings_intermediate(I).PhiX
						elem_settings_tmp(index_).PhiY 			= elem_settings_intermediate(I).PhiY
						elem_settings_tmp(index_).PhiZ 			= elem_settings_intermediate(I).PhiZ
						elem_settings_tmp(index_).OrigX 		= elem_settings_intermediate(I).OrigX
						elem_settings_tmp(index_).OrigY 		= elem_settings_intermediate(I).OrigY
						elem_settings_tmp(index_).OrigZ 		= elem_settings_intermediate(I).OrigZ
						elem_settings_tmp(index_).Amplitude 	= elem_settings_intermediate(I).Amplitude
						elem_settings_tmp(index_).Phase 		= elem_settings_intermediate(I).Phase
						elem_settings_tmp(index_).Noise 		= elem_settings_intermediate(I).Noise
						elem_settings_tmp(index_).IsNoiseInDBm	= elem_settings_intermediate(I).IsNoiseInDBm
						elem_settings_tmp(index_).FileName 		= elem_settings_intermediate(I).FileName
						index_ = index_ + 1
					End If
				Next
				' pass value back to elem_settings_intermediate
				ReDim elem_settings_intermediate(UBound(elem_settings_tmp))
				For I=0 To UBound(elem_settings_tmp)
					elem_settings_intermediate(I).ThetaX = elem_settings_tmp(I).ThetaX
					elem_settings_intermediate(I).ThetaY = elem_settings_tmp(I).ThetaY
					elem_settings_intermediate(I).ThetaZ = elem_settings_tmp(I).ThetaZ
					elem_settings_intermediate(I).PhiX = elem_settings_tmp(I).PhiX
					elem_settings_intermediate(I).PhiY = elem_settings_tmp(I).PhiY
					elem_settings_intermediate(I).PhiZ = elem_settings_tmp(I).PhiZ
					elem_settings_intermediate(I).OrigX = elem_settings_tmp(I).OrigX
					elem_settings_intermediate(I).OrigY = elem_settings_tmp(I).OrigY
					elem_settings_intermediate(I).OrigZ = elem_settings_tmp(I).OrigZ
					elem_settings_intermediate(I).Amplitude = elem_settings_tmp(I).Amplitude
					elem_settings_intermediate(I).Phase = elem_settings_tmp(I).Phase
					elem_settings_intermediate(I).Noise = elem_settings_tmp(I).Noise
					elem_settings_intermediate(I).IsNoiseInDBm = elem_settings_tmp(I).IsNoiseInDBm
					elem_settings_intermediate(I).FileName = elem_settings_tmp(I).FileName
				Next
				' update the list
				ReDim list_array(UBound(elem_settings_intermediate))
				For I=0 To UBound(list_array)
					list_array(I) = get_antelem_line(elem_settings_intermediate(I))
				Next
				DlgListBoxArray "elem_list_box", list_array
			End If
		ElseIf (DlgItem = "button_update") Then
			' keep dialog open
			antenna_element_dialog_function = True
			' update: check value + update line (update to elem_settings_intermediate)
			If (check_value()) Then
				' update elem_settings_intermediate
				iselected = DlgValue("elem_list_box")
				elem_settings_intermediate(iselected).ThetaX 		= CDbl(DlgText$("Editz1"))
				elem_settings_intermediate(iselected).ThetaY 		= CDbl(DlgText$("Editz2"))
				elem_settings_intermediate(iselected).ThetaZ 		= CDbl(DlgText$("Editz3"))
				elem_settings_intermediate(iselected).PhiX 			= CDbl(DlgText$("Editx1"))
				elem_settings_intermediate(iselected).PhiY 			= CDbl(DlgText$("Editx2"))
				elem_settings_intermediate(iselected).PhiZ 			= CDbl(DlgText$("Editx3"))
				elem_settings_intermediate(iselected).OrigX 		= CDbl(DlgText$("EditOrg_x"))
				elem_settings_intermediate(iselected).OrigY 		= CDbl(DlgText$("EditOrg_y"))
				elem_settings_intermediate(iselected).OrigZ 		= CDbl(DlgText$("EditOrg_z"))
				elem_settings_intermediate(iselected).Amplitude 	= CDbl(DlgText$("EditAmpl"))
				elem_settings_intermediate(iselected).Phase 		= CDbl(DlgText$("EditPhase"))
				elem_settings_intermediate(iselected).Noise 		= CDbl(DlgText$("EditNoise"))
				NoiseInDBm		= DlgValue("EditIsNoiseInDBm")
				If (NoiseInDBm = 0) Then
					elem_settings_intermediate(iselected).IsNoiseInDBm = False
				Else
					elem_settings_intermediate(iselected).IsNoiseInDBm = True
				End If
				elem_settings_intermediate(iselected).FileName 	= DlgText("EditFilename")
				' update the list
				ReDim list_array(UBound(elem_settings_intermediate))
				For I=0 To UBound(list_array)
					list_array(I) = get_antelem_line(elem_settings_intermediate(I))
				Next
				DlgListBoxArray "elem_list_box", list_array
			End If
		End If
	End Select

End Function

' get line for antenna element
Function get_antelem_line(Elem As AntElem)
	get_antelem_line = "(" + toStr(CStr(Elem.ThetaX)) + ";" + toStr(CStr(Elem.ThetaY)) + ";" + toStr(CStr(Elem.ThetaZ)) + ")" + "     " + "(" + toStr(CStr(Elem.PhiX)) + ";" + toStr(CStr(Elem.PhiY)) + ";" + toStr(CStr(Elem.PhiZ)) + ")" + "     " + "(" + toStr(CStr(Elem.OrigX)) + ";" + toStr(CStr(Elem.OrigY)) + ";" + toStr(CStr(Elem.OrigZ)) + ")"
	get_antelem_line = get_antelem_line + "     " + "(" + toStr(CStr(Elem.Amplitude)) + ";" + toStr(CStr(Elem.Phase)) + ")"
	get_antelem_line = get_antelem_line + "     " + toStr(CStr(Elem.Noise))
	If (Elem.IsNoiseInDBm) Then
		get_antelem_line = get_antelem_line + "     " + "True"
	Else
		get_antelem_line = get_antelem_line + "     " + "False"
	End If
	get_antelem_line = get_antelem_line + "     " + Elem.FileName
End Function

' function when user wants to edit any curve item
Sub edit_antenna_elements()
	Dim edit_antenna_elements_dialog_caption As String
	Dim list_box_string1 As String
	Dim list_box_string2 As String
	Dim list_array() As String
	edit_antenna_elements_dialog_caption = "Edit antenna elements"
	list_box_string1 = " Start-for-theta         Start-for-phi           Origin           Ampl-Phase           Noise    Is-dBm   FileName"
	list_box_string2 = "     (x,y,z)                      (x,y,z)                (x,y,z)           (ampl,phase)"
	' copy value from elem_settings to elem_settings_intermediate, from then work on elem_settings_intermediate
	ReDim elem_settings_intermediate(UBound(elem_settings))
	For I=0 To UBound(elem_settings)
		elem_settings_intermediate(I).ThetaX		= elem_settings(I).ThetaX
		elem_settings_intermediate(I).ThetaY		= elem_settings(I).ThetaY
		elem_settings_intermediate(I).ThetaZ		= elem_settings(I).ThetaZ
		elem_settings_intermediate(I).PhiX			= elem_settings(I).PhiX
		elem_settings_intermediate(I).PhiY			= elem_settings(I).PhiY
		elem_settings_intermediate(I).PhiZ			= elem_settings(I).PhiZ
		elem_settings_intermediate(I).OrigX			= elem_settings(I).OrigX
		elem_settings_intermediate(I).OrigY			= elem_settings(I).OrigY
		elem_settings_intermediate(I).OrigZ			= elem_settings(I).OrigZ
		elem_settings_intermediate(I).Amplitude		= elem_settings(I).Amplitude
		elem_settings_intermediate(I).Phase			= elem_settings(I).Phase
		elem_settings_intermediate(I).Noise 		= elem_settings(I).Noise
		elem_settings_intermediate(I).IsNoiseInDBm	= elem_settings(I).IsNoiseInDBm
		elem_settings_intermediate(I).FileName		= elem_settings(I).FileName
	Next

	' update list_array
	ReDim list_array(UBound(elem_settings_intermediate))
	For I=0 To UBound(list_array)
		list_array(I) = get_antelem_line(elem_settings_intermediate(I))
	Next
	' define dialog
	Begin Dialog EditAntennaElementDialog 650, 540, edit_antenna_elements_dialog_caption, .antenna_element_dialog_function
		Text 300,15,150,14,"x",.Text1
		Text 400,15,150,14,"y",.Text2
		Text 500,15,150,14,"z",.Text3
		Text 70,32,150,14,"Start for theta (z'-axis):",.Text4
		TextBox 270,32,90,21,.Editz1
		TextBox 370,32,90,21,.Editz2
		TextBox 470,32,90,21,.Editz3
		Text 70,60,150,14,"Start for phi (x'-axis):",.Text5
		TextBox 270,60,90,21,.Editx1
		TextBox 370,60,90,21,.Editx2
		TextBox 470,60,90,21,.Editx3
		Text 70,88,250,14,"Origin (rel. to point center):",.Text6
		TextBox 270,88,90,21,.EditOrg_x
		TextBox 370,88,90,21,.EditOrg_y
		TextBox 470,88,90,21,.EditOrg_z
		Text 160,119,90,14,"Amplitude:",.Text7
		Text 160,147,90,14,"Phase:",.Text8
		Text 160,175,90,14,"Noise:",.Text9
		Text 160,203,90,14,"Filename:",.Text10
		TextBox 270,116,170,21,.EditAmpl
		TextBox 270,144,170,21,.EditPhase
		TextBox 270,175,170,21,.EditNoise
		OptionGroup .EditIsNoiseInDBm
			OptionButton 450,172,50,21,"mW",.OptionButton1
			OptionButton 510,172,50,21,"dBm",.OptionButton2
		TextBox 270,201,170,21,.EditFilename
		PushButton 460,201,100,21,"Browse file...",.BrowseFile
		PushButton 160,230,100,25,"Add",.button_add
		PushButton 260,230,100,25,"Remove",.button_remove
		PushButton 360,230,100,25,"Update",.button_update
		Text 10, 269, 650, 21,list_box_string1,.Text11
		Text 10, 285, 650, 21,list_box_string2,.Text12
		ListBox 10,305,620,200,list_array(),.elem_list_box
		OKButton 160,510,100,25,.elem_button_ok
		CancelButton 260,510,100,25
	End Dialog

	' assign value to dialog + default values
	Dim edit_antenna_element_dialog As EditAntennaElementDialog
	edit_antenna_element_dialog.Editz1 		= "1.0"
	edit_antenna_element_dialog.Editz2 		= "0.0"
	edit_antenna_element_dialog.Editz3 		= "0.0"
	edit_antenna_element_dialog.Editx1 		= "0.0"
	edit_antenna_element_dialog.Editx2 		= "0.0"
	edit_antenna_element_dialog.Editx3 		= "1.0"
	edit_antenna_element_dialog.EditOrg_x 	= "0.0"
	edit_antenna_element_dialog.EditOrg_y 	= "0.0"
	edit_antenna_element_dialog.EditOrg_z 	= "0.0"
	edit_antenna_element_dialog.EditAmpl 	= "1.0"
	edit_antenna_element_dialog.EditPhase 	= "0.0"
	edit_antenna_element_dialog.EditNoise 	= "0.0"
	edit_antenna_element_dialog.EditIsNoiseInDBm = 1

	' show the dialog
	element_dlg_button = Dialog(edit_antenna_element_dialog)

	If (element_dlg_button = -1) Then
		' OK button pressed: update value from elem_settings_intermediate to elem_settings
		ReDim elem_settings(UBound(elem_settings_intermediate))
		For I=0 To UBound(elem_settings)
			elem_settings(I).ThetaX			= elem_settings_intermediate(I).ThetaX
			elem_settings(I).ThetaY			= elem_settings_intermediate(I).ThetaY
			elem_settings(I).ThetaZ			= elem_settings_intermediate(I).ThetaZ
			elem_settings(I).PhiX			= elem_settings_intermediate(I).PhiX
			elem_settings(I).PhiY			= elem_settings_intermediate(I).PhiY
			elem_settings(I).PhiZ			= elem_settings_intermediate(I).PhiZ
			elem_settings(I).OrigX			= elem_settings_intermediate(I).OrigX
			elem_settings(I).OrigY			= elem_settings_intermediate(I).OrigY
			elem_settings(I).OrigZ			= elem_settings_intermediate(I).OrigZ
			elem_settings(I).Amplitude		= elem_settings_intermediate(I).Amplitude
			elem_settings(I).Phase			= elem_settings_intermediate(I).Phase
			elem_settings(I).Noise 			= elem_settings_intermediate(I).Noise
			elem_settings(I).IsNoiseInDBm	= elem_settings_intermediate(I).IsNoiseInDBm
			elem_settings(I).FileName		= elem_settings_intermediate(I).FileName
		Next
	ElseIf (element_dlg_button = 0) Then
		' Cancel button pressed: do nothing
	End If

End Sub

' function for curve dialog
Private Function curve_dialog_function(DlgItem$, Action%, SuppValue&) As Boolean
	' the below codes can be commented later if needed
	DlgEnable("number_elem",False)
	DlgVisible("number_elem",False)
	DlgEnable("Text5",False)
	DlgVisible("Text5",False)
	DlgEnable("edit_elements",False)
	DlgVisible("edit_elements",False)
	
	' disable number_elem
	DlgEnable("number_elem",False)
	DlgEnable("iselected_string",False)
	DlgVisible("iselected_string",False)
	' disable sampling_rate when needed
	If (DlgValue("active") = 0) Then
		DlgVisible("velocity", False)		' uncomment here later if needed
		DlgVisible("Text1", False)			' uncomment here later if needed
		DlgVisible("Text2", False)			' uncomment here later if needed
		DlgEnable("velocity", False)
		DlgEnable("sampling_rate", False)
		DlgEnable("is_compute_corner", False)
		DlgEnable("edit_elements", False)
	Else
		DlgVisible("velocity", False)		' uncomment here later if needed
		DlgVisible("Text1", False)			' uncomment here later if needed
		DlgVisible("Text2", False)			' uncomment here later if needed
		DlgEnable("velocity", True)
		DlgEnable("sampling_rate", True)
		DlgEnable("is_compute_corner", True)
		DlgEnable("edit_elements", True)
	End If
	If (DlgValue("is_compute_corner") = 0) Then
		DlgEnable("sampling_rate",True)
	Else
		DlgEnable("sampling_rate",False)
	End If
	' check value when user presses OK button
	If (DlgItem = "curve_button_ok") Then
		Dim velocity As Integer
		Dim sampling_rate As Integer
		Dim string_velocity As String
		Dim string_sampling_rate As String
		string_velocity = DlgText$("velocity")
		string_sampling_rate = DlgText$("sampling_rate")
		If (Not IsNumeric(string_velocity)) Then
			MsgBox "Incorrect value for velocity", vbOkOnly
			curve_dialog_function = True	' not close the dialog
		ElseIf (Not IsNumeric(string_sampling_rate)) Then
			MsgBox "Incorrect value for sampling rate", vbOkOnly
			curve_dialog_function = True	' not close the dialog
		Else
			velocity = CDbl(string_velocity)
			sampling_rate = CDbl(DlgText$("sampling_rate"))
			If (velocity < 0) Then
				MsgBox "Incorrect value for velocity", vbOkOnly
				curve_dialog_function = True	' not close the dialog
			ElseIf (sampling_rate < 0) Then
				MsgBox "Incorrect value for sampling rate", vbOkOnly
				curve_dialog_function = True	' not close the dialog
			End If
		End If
	End If
	' open new dialog when user presses edit button
	If (DlgItem = "edit_elements") Then
		curve_dialog_function = True	' not close the dialog
		edit_antenna_elements()
		DlgText "number_elem", CStr(UBound(elem_settings)+1)
	End If
End Function

' function when user wants to edit any curve item
Sub edit_curve_item(iselected As Integer)
	Dim edit_curve_dialog_caption As String
	Dim iselected_string As String
	iselected_string = CStr(iselected)
	edit_curve_dialog_caption = "Edit 1D monitor - "
	edit_curve_dialog_caption = edit_curve_dialog_caption + curve_settings(iselected).CurveName
	' define dialog
	Begin Dialog EditCurveDialog 400,160, edit_curve_dialog_caption, .curve_dialog_function
		CheckBox 20,20,70,15,"Active",.active
		Text 20,45,150,14,"Velocity",.Text1
		TextBox 170,43,90,21,.velocity
		Text 270,45,150,14,"m/s",.Text2
		Text 20,70,150,14,"Sampling-rate",.Text3
		TextBox 170,68,90,21,.sampling_rate
		Text 270,70,150,14,"steps/wavelength",.Text4
		Text 270,70,150,14,iselected_string,.iselected_string
		CheckBox 100,20,200,14,"Only compute native points",.is_compute_corner
		Text 20,95,150,14,"Number of elements",.Text5
		TextBox 170,92,90,21,.number_elem
		OKButton 20,120,100,25,.curve_button_ok
		CancelButton 150,120,100,25
		PushButton 280,120,100,25,"Edit Elements",.edit_elements
	End Dialog

	Dim edit_curve_dialog As EditCurveDialog
	' assign values for dialog
	edit_curve_dialog.velocity = CStr(curve_settings(iselected).Velocity)
	edit_curve_dialog.sampling_rate = CStr(curve_settings(iselected).SamplingStep)
	edit_curve_dialog.is_compute_corner = curve_settings(iselected).OnlyComputeCornerPoint
	edit_curve_dialog.number_elem = CStr(curve_settings(iselected).NumElem)
	edit_curve_dialog.active = curve_settings(iselected).Active
	' temporary copy the data from curve_settings(iselected).Elem to elem_settings
	ReDim elem_settings(curve_settings(iselected).NumElem-1)
	For I=0 To UBound(elem_settings)
		elem_settings(I).OrigX 			= curve_settings(iselected).Elem(I).OrigX
		elem_settings(I).OrigY 			= curve_settings(iselected).Elem(I).OrigY
		elem_settings(I).OrigZ 			= curve_settings(iselected).Elem(I).OrigZ
		elem_settings(I).ThetaX 		= curve_settings(iselected).Elem(I).ThetaX
		elem_settings(I).ThetaY 		= curve_settings(iselected).Elem(I).ThetaY
		elem_settings(I).ThetaZ 		= curve_settings(iselected).Elem(I).ThetaZ
		elem_settings(I).PhiX 			= curve_settings(iselected).Elem(I).PhiX
		elem_settings(I).PhiY 			= curve_settings(iselected).Elem(I).PhiY
		elem_settings(I).PhiZ 			= curve_settings(iselected).Elem(I).PhiZ
		elem_settings(I).Amplitude 		= curve_settings(iselected).Elem(I).Amplitude
		elem_settings(I).Phase 			= curve_settings(iselected).Elem(I).Phase
		elem_settings(I).Noise 			= curve_settings(iselected).Elem(I).Noise
		elem_settings(I).IsNoiseInDBm 	= curve_settings(iselected).Elem(I).IsNoiseInDBm
		elem_settings(I).FileName 		= curve_settings(iselected).Elem(I).FileName
	Next

	' show the dialog
	curve_dlg_button = Dialog (edit_curve_dialog)

	' check either OK or Cancel button is pressed
	If (curve_dlg_button = -1) Then
		' if OK button is pressed:
		curve_settings(iselected).Velocity = CDbl(edit_curve_dialog.velocity)
		curve_settings(iselected).SamplingStep = CDbl(edit_curve_dialog.sampling_rate)
		curve_settings(iselected).OnlyComputeCornerPoint = edit_curve_dialog.is_compute_corner
		curve_settings(iselected).Active = edit_curve_dialog.active
		' update information of antenna element
		curve_settings(iselected).NumElem = CInt(edit_curve_dialog.number_elem)
		ReDim curve_settings(iselected).Elem(UBound(elem_settings))
		For I=0 To UBound(elem_settings)
			curve_settings(iselected).Elem(I).OrigX 		= elem_settings(I).OrigX
			curve_settings(iselected).Elem(I).OrigY			= elem_settings(I).OrigY
			curve_settings(iselected).Elem(I).OrigZ 		= elem_settings(I).OrigZ
			curve_settings(iselected).Elem(I).ThetaX 		= elem_settings(I).ThetaX
			curve_settings(iselected).Elem(I).ThetaY 		= elem_settings(I).ThetaY
			curve_settings(iselected).Elem(I).ThetaZ 		= elem_settings(I).ThetaZ
			curve_settings(iselected).Elem(I).PhiX 			= elem_settings(I).PhiX
			curve_settings(iselected).Elem(I).PhiY 			= elem_settings(I).PhiY
			curve_settings(iselected).Elem(I).PhiZ 			= elem_settings(I).PhiZ
			curve_settings(iselected).Elem(I).Amplitude 	= elem_settings(I).Amplitude
			curve_settings(iselected).Elem(I).Phase 		= elem_settings(I).Phase
			curve_settings(iselected).Elem(I).Noise 		= elem_settings(I).Noise
			curve_settings(iselected).Elem(I).IsNoiseInDBm 	= elem_settings(I).IsNoiseInDBm
			curve_settings(iselected).Elem(I).FileName 	= elem_settings(I).FileName
		Next
	ElseIf ((curve_dlg_button = 0)) Then
		' if Cancel button is pressed: do nothing
	End If
End Sub

' create list heading
Function create_list_heading(max_curve_name_length As Integer, curve_name As String, velocity As String, sampling_rate As String, only_computed_corner_point As String, num_elem As String, is_active As String)
	create_list_heading = pad_string(curve_name, max_curve_name_length)
	create_list_heading = create_list_heading + pad_string(is_active, 10)
	create_list_heading = create_list_heading + pad_string(only_computed_corner_point, 40)
	create_list_heading = create_list_heading + pad_string(sampling_rate, 25)
	'create_list_heading = create_list_heading + pad_string(num_elem, 20)				' uncomment here and comment the below line later if needed
	'create_list_heading = create_list_heading + pad_string(velocity, 20)				' uncomment here and comment the below line later if needed
End Function

' convert item to line
Function get_curve_line(Item As CurveSettings, max_curve_name_length As Integer)
	Dim only_computed_corner_point As String
	Dim is_active As String
	Dim sampling_step As String
	If (Item.OnlyComputeCornerPoint) Then
		only_computed_corner_point = "True"
		sampling_step = ""
	Else
		only_computed_corner_point = "False"
		sampling_step = CStr(Item.SamplingStep)
	End If
	If (Item.Active) Then
		is_active = "True"
	Else
		is_active = "False"
	End If
	get_curve_line = create_list_heading(max_curve_name_length, Item.CurveName, CStr(Item.Velocity), sampling_step, only_computed_corner_point, CStr(Item.NumElem), is_active)
End Function

' function for main dialog
Private Function main_dialog_function(DlgItem$, Action%, SuppValue&) As Boolean
	If (DlgItem = "checkBox") Then
		If (DlgValue("checkBox") = 1) Then
			DlgEnable("main_list_box", True)
			DlgEnable("main_edit", True)
		Else
			DlgEnable("main_list_box", False)
			DlgEnable("main_edit", False)
		End If
	ElseIf (DlgItem = "main_edit") Then
		main_dialog_function = True
		Dim iselected As Integer
		iselected = DlgValue("main_list_box")
		edit_curve_item(iselected)
		' update the value of list in main dialog
		Dim list_array() As String
		ReDim list_array(UBound(curve_settings))
		For I=0 To UBound(curve_settings)
			list_array(I) = get_curve_line(curve_settings(I), max_curve_name_length)
		Next
		DlgListBoxArray "main_list_box", list_array
		DlgValue "main_list_box", iselected
	End If
End Function

Sub Main ()
	id = "EnableChannelPropagationInPPMode"
	description = "If True, channel parameters can be computed in post-processing using ""Channel Computation"" template."
	caption = "Compute channel parameters in post-processing"
	Dim main_dialog_caption As String
	main_dialog_caption = "1D monitor definition for channel computation"
	Dim list_heading As String
	Dim list_array() As String
	Dim is_any_curve As Boolean
	Dim default_max_curve_name_length As Integer

	default_max_curve_name_length = 20

	' Obtain a list of curve with relevant information
	obtain_curve_info(is_any_curve, max_curve_name_length)
	max_curve_name_length = Max(default_max_curve_name_length, max_curve_name_length)

	' add info to list_array
	If (is_any_curve) Then
		ReDim list_array(UBound(curve_settings))
		For I=0 To UBound(curve_settings)
			list_array(I) = get_curve_line(curve_settings(I), max_curve_name_length)
		Next
	End If

	' Define the main dialog
	'list_heading = "Curve-name   Active    Only-compute-native-points    Sampling-rate    Num-elements    Velocity (m/s)"			' uncomment here and comment the below line later if needed
	list_heading = "Curve-name   Active    Only-compute-native-points    Sampling-rate"
	Begin Dialog MainDialog 700, 315, main_dialog_caption, .main_dialog_function
		CheckBox 20,11,400,15,caption,.checkBox
		Text 20,35,650,25,description,.Text1
		Text 20,65,650,14, list_heading,.main_text_heading
		ListBox 20,85,650,190,list_array(),.main_list_box
		PushButton 20,280,100,25,"Edit Entry",.main_edit
		OKButton 150,280,100,25
		CancelButton 280,280,100,25
	End Dialog

	' show the dialog
	Dim main_dlg As MainDialog
	main_dlg.checkBox = AsymptoticSolver.Get(id)
	main_dlg_button = Dialog (main_dlg)

	' check either OK or Cancel button is pressed
	If (main_dlg_button = -1) Then
		' OK button pressed: store to file
		MsgBox("Changes are only applied when the simulation and the template are re-run.")
		If (is_any_curve) Then
			write_curve_settings_to_history()
		End If
		AddToHistory "asymptotic solver: set " + caption, "AsymptoticSolver.Set """ + id + """, " + CStr(main_dlg.checkBox)
	ElseIf (main_dlg_button = 0) Then
		' Cancel button pressed: do nothing
	End If
End Sub
