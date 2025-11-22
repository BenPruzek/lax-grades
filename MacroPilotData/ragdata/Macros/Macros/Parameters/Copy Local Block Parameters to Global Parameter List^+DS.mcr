
' ------------------------------------------------------------------------------------------------
' Copyright 2009-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' 13-Jul-2022 ube: corrected blockarray index
' 22-Jul-2021 ube: added option to always add the block name as prefix
' 19-Jul-2019 fcu: updated .GetTypeName() to work correctly after fix of CST-52569
' 05-Sep-2017 fsr: Macro now works with block names that contain space, minus, or plus
' 22-Nov-2013 fsr: user may now select 'All blocks'; only include CST blocks; operations carried out in background mode without opening subprojects
' 17-Jun-2011 fsr: user may choose if the MWS block that opens for the transfer should remain open or not; check if at least 1 block defined
' 27-Apr-2011 fsr: user may choose to "repeat keep/rename for all remaining parameters"
' 02-Mar-2011 fsr: if DS parameter already exists, user is informed for naming conflict
' 19-Mar-2009 ube: first version
'-------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub Main

	Dim sName As String
	Dim sType As String
	Dim sValue As String
	Dim parameterConflictResolution As Integer ' 0: Keep name, 1: Rename, 2: Keep and repeat for all, 3: Rename and repeat for all
	Dim bLocalParametersWin As Boolean
	Dim bAlwaysAddBlocknamePrefix As Boolean

	Dim nblocks As Integer, i1 As Integer, b

	parameterConflictResolution = -1

	With Block
		nblocks = .StartBlockNameIteration  ' Resets the block iterator and returns the number of blocks.

		If nblocks = 0 Then
			MsgBox("No parametric blocks defined!","Error")
			Exit Sub
		End If

		Dim aNames() As String, sTempName As String
		ReDim aNames(nblocks)

		Dim i2 As Integer
		i2 = 0

		aNames(0) = "All blocks"
		For i1=1 To nblocks
			With Block
				.Enable3DCommands(False)
				sTempName = .GetNextBlockName
				.Name sTempName
				If (InStr(.GetTypeName, "STUDIO")>0) Then
					i2 = i2+1
					aNames(i2) = sTempName
				End If
			End With
		Next i1
	End With

	Begin Dialog UserDialog 510,140,"Create global parameters from local block parameters" ' %GRID:10,7,1,1
		GroupBox 20,7,470,98,"Select Block",.GroupBox1
		DropListBox 40,28,430,21,aNames(),.aNames
		OKButton 300,112,90,21
		CancelButton 400,112,90,21
		CheckBox 40,56,430,14,"Copy local values to global values if parameter already exists",.LocalWinsCB
		CheckBox 40,77,410,14,"Always add block name as prefix to parameter name",.AlwaysAddBlocknamePrefix
'		CheckBox 40,56,280,14,"Close subproject/block again after import",.CloseAfterImportCB
	End Dialog
	Dim dlg As UserDialog
'	dlg.CloseAfterImportCB = 1

	If (Dialog(dlg) = 0) Then Exit All

	Dim nStart As Long, nEnd As Long, n As Long
	Dim sBlockName As String, sBlockNameNoForbiddenChars As String

	bLocalParametersWin = CBool(dlg.LocalWinsCB)
	bAlwaysAddBlocknamePrefix = CBool(dlg.AlwaysAddBlocknamePrefix)
	If aNames(dlg.aNames) = "All blocks" Then
		nStart = 1
		nEnd = nblocks
	Else ' loop will only have 1 item
		nStart = dlg.aNames
		nEnd = dlg.aNames
	End If

	For n = nStart To nEnd
		sBlockName = aNames(n)
		sBlockNameNoForbiddenChars = Replace(Replace(Replace(sBlockName, " ", "_"), "-", "_"), "+", "_")
		With Block
			If (sBlockName = "") Then Exit For ' done
			'DS.ReportInformationToWindow("Copying local parameters from block: '"+aNames(n)+"'.")
			.Name sBlockName
			.Enable3DCommands(False)
			.StartPropertyIteration
			.GetNextProperty(sName, sType, sValue)
			While StrComp(sName, "", 1)
				If (IsNumeric(sValue)) Then ' do not transfer dependent parameters
					If bAlwaysAddBlocknamePrefix Then
						DS.StoreParameter(sBlockNameNoForbiddenChars+"_"+sName, sValue)
						.SetDoubleProperty(sName, sBlockNameNoForbiddenChars+"_"+sName)
					Else
						If (ParameterExists(sName)) Then
							If (parameterConflictResolution=-1) Or (parameterConflictResolution=0) Or (parameterConflictResolution=1) Then ' not defined yet or user did not check "repeat for all"
								parameterConflictResolution = ParameterExistsWarningGUI(sName, sBlockNameNoForbiddenChars)
							End If
							If (parameterConflictResolution=0) Or (parameterConflictResolution=2) Then ' keep the name
								If bLocalParametersWin Then DS.StoreParameter(sName, sValue)
								.SetDoubleProperty(sName, sName)
							ElseIf (parameterConflictResolution=1) Or (parameterConflictResolution=3) Then ' rename
								DS.StoreParameter(sBlockNameNoForbiddenChars+"_"+sName, sValue)
								.SetDoubleProperty(sName, sBlockNameNoForbiddenChars+"_"+sName)
							End If
						Else
							DS.StoreParameter(sName, sValue)
							.SetDoubleProperty(sName, sName)
						End If
					End If
				End If
				.GetNextProperty(sName, sType, sValue)
			Wend
		End With
	Next n

	If InStr(GetApplicationName(), "DS for")>0 Then
		MsgBox("Done! A parametric structure update might be necessary, please check the 3D tab.", "Copy local parameters to global parameters")
	Else
		MsgBox("Done!", "Copy local parameters to global parameters")
	End If
	' Close MWS if checkbox is active. DS will remain open.
'	If (dlg.CloseAfterImportCB = 1) Then
'		Dim studio As Object
'		' get currently open CST STUDIO SUITE
'		Set studio = CreateObject("CSTStudio.Application")
'		Dim mwsproj As Object
'		' get active MWS (or EMS, PS, ...)
'		Set mwsproj = studio.Active3D
'		If Not(mwsproj Is Nothing) Then mwsproj.Quit
'	End If

End Sub

Function ParameterExists(sName As String) As Boolean
	Dim i As Long
	ParameterExists = False
	For i = 0 To DS.GetNumberOfParameters-1
		If sName = DS.GetParameterName(i) Then
			ParameterExists = True
			Exit For
		End If
	Next
End Function

Function ParameterExistsWarningGUI(paraName As String, blockName As String) As Integer
	Begin Dialog UserDialog 550+15*(Len(paraName)+Len(blockName)+1),84,"Parameter name conflict" ' %GRID:10,7,1,1
		Text 20,14,520+15*(Len(paraName)+Len(blockName)+1),14,"The parameter '"+paraName+"' already exists. Would you like to keep the name or rename it to '"+blockName+"_"+paraName+"'?",.outputT
		PushButton 340+15*(Len(paraName)+Len(blockName)+1),49,90,21,"Keep Name",.KeepPB
		PushButton 440+15*(Len(paraName)+Len(blockName)+1),49,90,21,"Rename",.RenamePB
		CheckBox 80+15*(Len(paraName)+Len(blockName)+1),56,250,14,"Repeat for all remaining parameters:",.RepeatCB
	End Dialog
	Dim dlg As UserDialog
	' 0 = Keep name, 1 = Rename; 2 = Keep name for all; 3 = Rename for all
	ParameterExistsWarningGUI = Dialog(dlg)-1 + 2*dlg.RepeatCB
End Function
