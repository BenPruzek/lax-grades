
' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 26-Sep-2011 ube: First version
' ================================================================================================

Sub Main ()

	If (MsgBox "This macro defines specific post processing options for slow wave structures." +vbCrLf+vbCrLf + "Assumptions:" +vbCrLf + "  - Parameter ""phase"" is used for periodic boundary phase shift"+vbCrLf + "  - Results will be available after a Parameter Sweep of ""phase"""+vbCrLf+vbCrLf+"Do you want to enable those postprocessing steps?",vbInformation+vbYesNo," Slow Wave userdefined Watch")=vbYes Then

		AddToHistory "add watch: userdefined", "ParameterSweep.AddUserdefinedWatch"

		FileCopy GetInstallPath + "\Library\Macros\Solver\E-Solver\Model.pfc",getprojectpath("Model3D")+"Model.pfc"

		MsgBox "successfully defined"

	End If

End Sub
