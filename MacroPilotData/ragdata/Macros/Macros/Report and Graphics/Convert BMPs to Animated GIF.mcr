'#Language "WWB-COM"

' Copyright 2010-2023 Dassault Systemes Deutschland GmbH
'-------------------------------------------------------------------------------------------------
' History of Changes
' ------------------------------------------------------------------------------------------------
' 11-Feb-2020 mwl: use video_creation.lib instead of doing everything manually
' 18-Dec-2013 fsr:  Exit properly if user cancels file path selection
' 15-Oct-2010 fsr:  first version - driven by PS-users
'-------------------------------------------------------------------------------------------------

Option Explicit

'#include "video_creation.lib"
'#include "vba_globals_all.lib"

Sub Main

	Dim imageDirectory As String
	imageDirectory = GetFilePath("*.bmp", "Bitmap files|*.bmp", GetProjectPath("Result"), "Select first bitmap file", 0)
	If imageDirectory = "" Then ' user pressed 'Cancel'
		Exit All
	End If

	GenerateMovieFromImagesInDirectory_Interactive(DirName(imageDirectory), "*.bmp", imageDirectory, "common_preloadedmacro_Graphics_GenerateVideoFromImages")

End Sub
