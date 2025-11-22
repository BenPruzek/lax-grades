' Extrudes multiple picked faces
' User can enter shapenames and component
' Material will be cloned from the original shape
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 22-Nov-2006 msc: First version

Sub Main
	Begin Dialog UserDialog 330,174,"Extrude Multiple Current Ports" ' %GRID:2,2,1,1
		GroupBox 10,7,310,49,"Properties",.GroupBox1
		TextBox 148,24,70,20,.sExtr
		Text 20,28,120,14,"Extrusion Height",.Text1
		OKButton 190,147,90,21
		CancelButton 60,147,90,21
		GroupBox 10,70,310,70,"Current Port Properties",.GroupBox2
		OptionGroup .Group1
			OptionButton 40,91,100,14,"CurrentPort",.bCP
			OptionButton 40,112,110,14,"Voltage Port",.bVP
		TextBox 180,114,80,20,.TextBox1
		Text 184,98,70,14,"Value",.Text2
		Text 268,118,40,14,"A or V",.Text3
	End Dialog
	Dim dlg As UserDialog

	dlg.sExtr  = "0"
	bCP = True
	bVP = False
	If (Dialog(dlg) = 0) Then Exit All

	Dim dExtrHeight As Double
	dExtrHeight = CDbl(dlg.sExtr)

	Dim n_faces As Integer
	n_faces = Pick.GetNumberOfPickedFaces

	If (n_faces = 0) Then 
		MsgBox "No Faces picked."
		Exit All
	End If

	Dim i_fids()
	ReDim i_fids(n_faces)

	Dim s_names()
	ReDim s_names(n_faces)

	Dim s_sname As String
	Dim i_fid As Long

	For i=1 To n_faces
		s_name  = Pick.GetPickedFaceFromIndex(i,i_fid)
		s_names(i) = s_name
		i_fids(i) = i_fid
	Next i

	' Extrude the pins first
    Dim sCommand As String
    Dim sTmp1 As String
    Dim sTmp2 As String
    Dim sTmp3 As String

	sCommand = ""
    sTmp1 = ""
    sTmp2 = ""
	sTmp3 = ""
	sCommand = sCommand + "Pick.ClearAllPicks" + vbLf

	Pick.ClearAllPicks

	For i=1 To n_faces

		sTmp1 = s_names(i)
		sTmp2 = i_fids(i)
		sTmp3 = Solid.GetMaterialNameForShape(s_names(i))

		' Face Repick
		sCommand = sCommand	+ "Pick.PickFaceFromId(""
		sCommand = sCommand + sTmp1 +"""
		sCommand = sCommand + "," + """
		sCommand = sCommand + sTmp2 + """
		sCommand = sCommand + ")" + vbLf

		' Offset face
	 	sCommand = sCommand	+ "Solid.OffsetSelectedFaces(""
	 	sCommand = sCommand	+ dlg.sExtr + """
		sCommand = sCommand + ")" + vbLf

		' Assign Current Port
		sCommand = sCommand + "With CurrentPort" + vbLf
		'With CurrentPort
		sCommand = sCommand + ".Reset" + vbLf
		'.Reset
		sCommand = sCommand + ".Name ""
		sCommand = sCommand + "currentport" + CStr(i) + """
		sCommand = sCommand + vbLf
		'.Name "currentport1"
		sCommand = sCommand + ".Value ""
		sCommand = sCommand + dlg.TextBox1
		sCommand = sCommand + """
		sCommand = sCommand + vbLf
		'.Value "1"
 		If dlg.Group1=False Then 'Current Port
			sCommand = sCommand + ".ValueType " + """
			sCommand = sCommand + "Current" + """
			sCommand = sCommand + vbLf
		Else ' Potential
			sCommand = sCommand + ".ValueType " + """
			sCommand = sCommand + "Potential" + """
			sCommand = sCommand + vbLf
		End If

		' Face for current port
		sCommand = sCommand	+ ".Face ""
		sCommand = sCommand + sTmp1 +"""
		sCommand = sCommand + "," + """
		sCommand = sCommand + sTmp2 + """
		sCommand = sCommand + vbLf

		'.Face "component1:solid1_9", "21"
		sCommand = sCommand + ".Create" + vbLf
		' .Create
		sCommand = sCommand + "End With" + vbLf
		'End With

	Next i

	AddToHistory "Extrude Multiple Current Ports", sCommand

	Pick.ClearAllPicks

End Sub
