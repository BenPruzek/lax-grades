
Option Explicit
' Little Wrapper to call the Via Wizard from CST Macros
' 091106 msc: initial
' ================================================================================================
' Copyright 2009-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 10-Nov-2009 ube: First version
' ================================================================================================

Sub Main
	Dim sCallWizard As String

	sCallWizard = GetInstallPath + "\Library\Macros\Wizard\"
	sCallWizard = sCallWizard + "via.bat"

	Shell(sCallWizard, vbHide)
End Sub
