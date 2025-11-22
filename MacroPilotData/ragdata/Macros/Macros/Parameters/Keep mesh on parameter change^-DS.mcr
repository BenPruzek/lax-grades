Option Explicit
'#include "vba_globals_all.lib"

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ---------------------------------------------------------------------------------------------------------------------------------------------
' 23-Feb-2016 ube: Strg replaced by Ctrl
' 13-May-2014 mfe: first version - created by modifying macro 'Record current Parameter Values^+DS.rtp'
' ---------------------------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

	Select Case Action
		Case 1 ' Dialog box initialization
		Case 2 ' Value changing or button pressed
		Case 3 ' ComboBox or TextBox Value changed
		Case 4 ' Focus changed
		Case 5 ' Idle
	End Select
End Function

Function Define(sName As String, bCreate As Boolean) As Boolean

	Dim ParameterArray(1000) As String, IndexSelection() As Integer
	Dim ii As Integer, jj As Integer, jnow As Integer, sArray() As String
	Const Separator = "#"
	Dim sAllSelectedParaNames As String
	Dim sPara As String

	' get all parameter names, and put the already marked ones into 'sAllSelectedParaNames'
	For ii = 0 To GetNumberOfParameters-1
		sPara = GetParameterName(ii)
		ParameterArray(ii) = sPara
		If (IsKeepMeshParameter (sPara)) Then
			sAllSelectedParaNames += sPara + Separator
		End If
	Next ii

	sArray = Split(sAllSelectedParaNames,Separator)
	jnow = -1
	ReDim IndexSelection(1000)

	' the following lines setup index-array according to previous sAllSelectedParaNames,
	' so that previous selected parameters are again selected when reopening the dialogue

	For jj = 0 To UBound(sArray)
		For ii = 0 To GetNumberOfParameters-1
			If sArray(jj) = ParameterArray(ii) Then
				jnow = jnow + 1
				IndexSelection(jnow)=ii
			End If
		Next ii
	Next jj

	If jnow>=0 Then ReDim Preserve IndexSelection(jnow)

	Begin Dialog UserDialog 450,217,"Keep mesh on parameter change", .DialogFunc ' %GRID:10,7,1,1
		MultiListBox 20,45,400,140,ParameterArray(),.ParameterArray
		Text 20,10,420,14,"Please multi-select all parameters, which do not affect the mesh:",.Text1
		Text 20,25,420,14,"(Press 'Ctrl'+'left mouse' to deselect)",.Text2
		OKButton 20,189,90,21
		CancelButton 120,189,90,21
	End Dialog
	Dim dlg As UserDialog
	If jnow>=0 Then dlg.ParameterArray = IndexSelection

	If (Not Dialog(dlg)) Then
		Define = False
	Else
		Define = True

		' reset all parameters
		For ii = 0 To GetNumberOfParameters-1
			KeepMeshOnParameterChange GetParameterName(ii), False
		Next ii

		' now set the selected parameters to true
		If UBound(dlg.ParameterArray)>=0 Then
			KeepMeshOnParameterChange ParameterArray(dlg.ParameterArray(0)), True
		End If

		For jj=1 To UBound(dlg.ParameterArray)
			KeepMeshOnParameterChange ParameterArray(dlg.ParameterArray(jj)), True
		Next
	End If
End Function


Sub Main ()

	If (Define("test", True)) Then

	End If

End Sub
