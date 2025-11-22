' ================================================================================================
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 29-Nov-2013 ube: First version
' ================================================================================================

Sub Main ()

	Dim sBoolean As String
	If (MsgBox "The TET FD Solver can use the Green and Sandy Model to simulate partially magnetized ferrites. "+vbCrLf+vbCrLf+"Please choose settings for this model:"+vbCrLf+vbCrLf+"   Yes : Enable it."+vbCrLf+"   No  : Disable it.",vbInformation+vbYesNo,"Partially Magnetized Ferrites")=vbYes Then
		sBoolean = "True"
	Else
		sBoolean = "False"
	End If

	AddToHistory "(*) Green-Sandy-Model", "FDSolver.SetUseGreenSandyFerriteModel " + sBoolean 

End Sub
