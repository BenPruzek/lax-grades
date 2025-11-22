Option Explicit
'#include "vba_globals_all.lib"

' ================================================================================================
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 22-May-2006 ube: First version
' ================================================================================================

Sub Main ()

	MsgBox  "Since v2021 copy/paste of result templates between cst-projects is possible directly from the Result Template dialog, therefore this macro is no longer required."+vbCrLf+vbCrLf+ _
			"Following workflow has to be applied:"+vbCrLf+vbCrLf+ _
			" 1) open first project, from where result templates should be copied, multi select templates, Right Mouse - Copy"+vbCrLf+vbCrLf+ _
			" 2) open second project, where result templates should be included, open template dialogue, Right Mouse - Paste"+vbCrLf+vbCrLf+ _
			"" _
			,vbInformation, "Import Result Templates"

End Sub
