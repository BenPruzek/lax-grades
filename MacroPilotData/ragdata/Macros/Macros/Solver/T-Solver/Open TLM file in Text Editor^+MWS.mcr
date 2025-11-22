'#Language "WWB-COM"

' ================================================================================================
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 08-Mar-2012 ube: First version
' ================================================================================================
Option Explicit

' Opens the model.tlm file of the current project

Sub Main

	Dim sTLMFile As String

	sTLMFile = GetProjectPath("Result")+"MS\Model.tlm"

	If (Dir(sTLMFile)="") Then
		MsgBox("Cannot find TLM file. Please start the TLM solver first.","Error")
		Exit All
	Else
		Shell "notepad " & sTLMFile, 3
	End If

End Sub
