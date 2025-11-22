' ================================================================================================
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'-----------------------------------------------------------------------------------------
' 31-Jan-2014 dta: added screen update at the end of the execution
' 15-May-2013 tgl,ube: first version
'-----------------------------------------------------------------------------------------

Option Explicit

Sub Main ()

	Dim sCommand As String

	sCommand = ""
	' create the group to add the solids to; additionally set up the meshing special
	sCommand = sCommand + "Group.Add ""materialindependent-meshing"", ""mesh""" + vbLf
	sCommand = sCommand + "With MeshSettings" + vbLf
	sCommand = sCommand + "With .ItemMeshSettings (""group$materialindependent-meshing"")" + vbLf
	sCommand = sCommand + ".SetMeshType ""Tet""" + vbLf
	sCommand = sCommand + ".Set ""MaterialIndependent"", 1" + vbLf
	sCommand = sCommand + "End With" + vbLf + "End With" + vbLf
		
	AddToHistory "Macro: Create Material-independent Mesh Group", sCommand

	ScreenUpdating True

End Sub
