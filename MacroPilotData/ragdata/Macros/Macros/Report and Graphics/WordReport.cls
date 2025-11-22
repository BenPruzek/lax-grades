VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_PredeclaredId = False
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Name = "WordReport"

'=============================================================================
'
' WordReport: Example class module for generating simulation reports
'             with Microsoft Word 2000.
' Copyright 1999-2021 Dassault Systemes Deutschland GmbH
'
' To use this class you need to include the following line on top
' of your code module:
'                          '#Uses "WordReport.cls"
'
' You may then obtain an object of this class by:
'                          Dim report As New WordReport
'
' This object now provides you with a couple of useful methods
' (defined below):
'
' NewFile <filename>			Create a new Word document with the
'                               specified file name. All subsequent
'                               output will be placed into this file.
'
' CloseFile                     Close the currently active Word document.
'                               All changes to the document will be saved.
'
' WriteText <string>            Write the specified text into the active
'                               Word document.
'
' NewLine                       Add a new line to the active Word document.
'
' NewPage                       Add a page break to the active Word document.
'
' WriteParagraph <string>       Write the specified text including a new 
'                               line to the active Word document.
'
' WriteTitle <string>           Write the specified text centered, bold face
'                               using a large font to the active Word document.
'
' WriteHeading <string>         Write the specified text bold face to the active 
'                               Word document.
'
' WriteCurrentPlot              Place a screen-shot of the current plot 
'                               window into the word document as a picture.
'
' WriteStructurePlot            Place a screen-shot of the current structure 
'                               view into the word document as a picture.
'
' WriteResultPlot <path>        Place a picture of the specified result
'                               item into the active word document. The
'                               path specifies the location of the result
'                               in the navigation tree. 
'                               (example: "1D Results\Port signals")
'
' WriteParameterTable           Place a table containing all defined
'                               parameters and their current values into
'                               the active word document.
'
' GetWordObject                 Return a handle to the Word application
'                               object which contains the currently
'                               active word document.
'
'=============================================================================

'Option ExplicitOn
Dim word As Object

'------------------------------------------------------------------------
' Public class interface

Public Sub NewFile (FName As String)
    
    word.Documents.Add
    word.ActiveDocument.SaveAs FileName:=FName
    
End Sub

Public Sub CloseFile ()

	word.ActiveDocument.Close SaveChanges:=True

End Sub

Public Sub WriteText (Contents As String)

	word.Selection.TypeText Text:=Contents

End Sub

Public Sub NewLine ()

	word.Selection.TypeParagraph

End Sub


Public Sub NewPage ()

	word.Selection.InsertBreak Type:=1
	
End Sub

Public Sub WriteParagraph (Contents As String)

	WriteText(Contents)
	NewLine
	
End Sub

Public Sub WriteTitle (Contents As String)

    word.Selection.Font.Bold = True
    word.Selection.ParagraphFormat.Alignment = 1
    word.Selection.Font.Size = 14

	WriteText(Contents)
	NewLine

    word.Selection.Font.Bold = False
    word.Selection.ParagraphFormat.Alignment = 0
    word.Selection.Font.Size = 10
    	
End Sub

Public Sub WriteHeading (Contents As String)

	word.Selection.Font.Bold   = True
	WriteText(Contents)
	word.Selection.Font.Bold   = False
	NewLine
	NewLine
	
End Sub


Public Sub WriteCurrentPlot ()

   ExportImageToClipboard 0,0	
   
   word.Selection.TypeParagraph
   word.Selection.Paste
   word.Selection.TypeParagraph
   
End Sub

Public Sub WriteStructurePlot ()

   SelectTreeItem "Components"
   SelectModelView
   WriteCurrentPlot
   
End Sub

Public Sub WriteResultPlot (Path As String)

   SelectTreeItem Path
   WriteCurrentPlot
   
End Sub

Public Sub WriteParameterTable ()

	Dim ipar, npars As Long
	
	' Determine the number of parameters
	
	npars = GetNumberOfParameters()
	
	If (npars <> 0) Then
	
		' First of all create a new table in the word document
		
	    word.ActiveDocument.Tables.Add Range:=word.Selection.Range, NumRows:=npars+1, NumColumns:= 2
	    word.Selection.Tables(1).AutoFormat ApplyBorders :=True, _
	    	ApplyShading:=True, ApplyFont:=True, ApplyColor:=False, _
	        ApplyHeadingRows:=True, ApplyLastRow:=False, ApplyFirstColumn:=True, _
	        ApplyLastColumn:=False, AutoFit:=True
	    
	    ' Now insert the column headings
	        
	    word.Selection.Font.Italic = True
	    word.Selection.Font.Bold   = True
	    word.Selection.TypeText Text:="Parameter"
	    word.Selection.MoveRight
	    word.Selection.Font.Italic = True
	    word.Selection.Font.Bold   = True
	    word.Selection.TypeText Text:="Value"
	    word.Selection.MoveRight
	    word.Selection.MoveRight
	    
	    ' Finally insert the parameters (name and value)
	    
	    For ipar=1 To npars
	    
		    word.Selection.TypeText Text:=GetParameterName(ipar-1)
		    word.Selection.MoveRight
		    word.Selection.TypeText Text:= CStr(GetParameterNValue(ipar-1))
		    word.Selection.MoveRight
		    word.Selection.MoveRight
	    
	    Next ipar
	    
    End If

End Sub

Public Function GetWordObject () As Object

	GetWordObject = word

End Function

'------------------------------------------------------------------------
' Constructor and Destructor

Private Sub Class_Initialize()
    
    ' Make connection to Word

	Set word = CreateObject("Word.Application")    
End Sub

Private Sub Class_Terminate()
    
    ' Terminate connection to word

	word.Quit    
    Set word = Nothing
End Sub
