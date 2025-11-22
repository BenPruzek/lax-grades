' *Filter Analysis / Smith Chart - Phase Rotation
' !!!
'
'--------------------------------------------------------------------------------------------
' Use this macro to grafically "rotate" the Smith-Chart-Data
' NOTE, that this is just a phase-rotation and doesn't work as a deembedding function
'
' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 20-Mar-2014  fhi,mbr:  Initial Version , redesign for V2104
'
'#Language "WWB-COM"
'

Option Explicit

'#include "mws_ports.lib"
'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "complex.lib"

Const HelpFileName = "common_preloadedmacro_filter_analysis_smith_chart_-_phase_rotation"

Public p_portnr As String
Public p_modenr As String
Public p_tics_on As Integer
Public p_phase_offset As String
Public p_nr_of_markers As String
Public p_only_even_90 As Integer
Public p_ar_on As Integer

Sub ReadHeader( ByVal sFilename As String, sData As String )

	Dim nFileNum As Integer
	Dim sNextLine As String
	Dim nLineCount As Long
	nLineCount = 0

	nFileNum = FreeFile
	Open sFilename For Input As nFileNum

	Do While ( Not EOF( nFileNum ) And ( nLineCount < 30 ) )

		nLineCount = nLineCount + 1
    	Line Input #nFileNum, sNextLine

		Dim sArray() As String
		sArray = Split( sNextLine, vbTab )
		If UBound( sArray ) = 4 Then
			'now comes the data. implicit assumtion: we have five data columns
			Close nFileNum
			Exit Sub
		Else
			'there is more header information
			sData = sData + sNextLine + vbCrLf
		End If
	Loop


	ReportError( "ASCII file format is not supported. Maybe reference impedance data is missing?" )

	Close nFileNum

End Sub

Sub ReplaceHeaderValue( sHeader As String, ByVal sHeaderKey As String, ByVal sNewValue As String )

	' replace an existing value in the header with the new one.
	' example: "NPoints   =   23 " to "NPoints   =   42"

	Dim keypos As Long
	keypos = InStr( sHeader, sHeaderKey )

	If keypos = 0 Then
		ReportError("Header Key not found: " + sHeaderKey )
	End If

	Dim equalpos As Long
	equalpos = InStr( Mid( sHeader, keypos ) , "= " ) + keypos

	Dim sRight As String
	sRight = Mid( sHeader, equalpos )

	Dim sLeft As String
	sLeft = Left( sHeader, equalpos )

	Dim crlfpos As Long
	crlfpos = InStr( sRight, vbCrLf )

	sRight = sNewValue +  Mid( sRight, crlfpos )

	sHeader = Left( sHeader, equalpos ) + sRight

End Sub


Sub WriteData( iOutputFile As Integer, o As Object, o_imp As Object, bUseCommaAsSeparator As Boolean )
	If o.GetN = 0 Then
		ReportError( "Empty data object." )
	ElseIf o_imp.GetN = 0 Then
		ReportError( "Empty reference impedance data object." )
	End If

	o_imp.MakeCompatibleTo( o )

	Dim N As Long
	Dim x As Double
	Dim yre As Double
	Dim yim As Double

	Dim x_imp As Double
	Dim yre_imp As Double
	Dim yim_imp As Double

	For N = 0 To o.GetN - 1
		o.GetDataFromIndex( N, x, yre, yim )
		o_imp.GetDataFromIndex( N, x_imp, yre_imp, yim_imp )
		Dim sLine As String
		sLine = CStr( x ) + vbTab + CStr( yre ) + vbTab + Cstr( yim ) + vbTab + Cstr( yre_imp ) + vbTab + Cstr( yim_imp )
		If bUseCommaAsSeparator Then
			sLine = Replace( sLine, ".", "," )
		Else 'fhi
		'	sLine = Replace( sLine, ",", "." )
		End If
		Print #iOutputFile, sLine
	Next

End Sub

Sub WriteHeader( iOutputFile As Integer, sHeader As String )
	Print #iOutputFile, sHeader; 'have no line break at the end!
End Sub

Sub CreateFileInClipboardFormat( sFilename As String, sTargetName As String, sHeader As String, bUseCommaAsSeparator As Boolean, o As Object, o_imp As Object )

	'adapt header information to new data
	ReplaceHeaderValue( sHeader, "Npoints", CStr( o.GetN ) )
	ReplaceHeaderValue( sHeader, "Curvelabel", sTargetName )

	'write header and data columns
	Dim iOutputFile As Integer
	iOutputFile = FreeFile()
	Open sFilename For Output As #iOutputFile

	WriteHeader( iOutputFile, sHeader )
	WriteData( iOutputFile, o, o_imp, bUseCommaAsSeparator )

	Close #iOutputFile

End Sub

Function xGetDecimalSeparator() As String
	Dim cst_separator    As String
	cst_separator    = Mid$(CStr(0.5), 2, 1)
	xGetDecimalSeparator= cst_separator
End Function

Function DecimalSeparatorIsComma() As Boolean

	'determine decimal symbol: comma or dot?
	DecimalSeparatorIsComma = False
	Dim sSeparator As String
	sSeparator = GetDecimalSeparator()
	If sSeparator = "," Then
		DecimalSeparatorIsComma = True
	ElseIf sSeparator <> "." Then
		ReportError( "Decimal separator symbol must be dot or comma." )
	End If

End Function

Sub CreateTreeFolder( sFolder As String )
	Resulttree.Reset
	Resulttree.Name( sFolder )
	Resulttree.Type( "folder" )
	Resulttree.File( "temp_filename" )
	Resulttree.DeleteAt( "never" )
	Resulttree.Add
End Sub

Sub CreateRotatedSmithChartFromTreeItem( sTreeItem As String, theta As Double, sTargetFolder As String, sTargetName As String )

	'get data object from tree item
    Dim resultID As String
    resultID = GetLastResultID()  
	Dim o As Object
	Set o = Resulttree.GetResultFromTreeItem( sTreeItem, resultID )
    
    If o Is Nothing Then
       ReportError("No data available for tree item: " + sTreeItem )
       Exit sub
    End If
    

	'get the ref impedance object from tree item
	Dim o_imp As Object
	Set o_imp = Resulttree.GetImpedanceResultFromTreeItem( sTreeItem, resultID )
    
    If o Is Nothing Then
       ReportError("No impedance data available for tree item: " + sTreeItem )
       Exit sub
    End If
    


	'rotate. ref-impedance could be modified here as well.
	o.ScalarMultReIm( Cos( theta ), Sin( theta ) )

	CreateTreeFolder( sTargetFolder )

	'create a valid ascii file in clipboard format
	Dim sTempFileName As String
	sTempFileName =  GetProjectPath( "Temp" ) + "ascii_data.txt"
	SelectTreeItem( sTreeItem )
	StoreCurvesInASCIIFile( sTempFileName )

	'read clipboard format header and determine decimal separator
	Dim sHeader As String
	ReadHeader( sTempFileName, sHeader )

	'store the modified data in clipboard format
	CreateFileInClipboardFormat( sTempFileName, sTargetName, sHeader, DecimalSeparatorIsComma(), o, o_imp )

	'paste it to get a smith chart
	PasteCurvesFromASCIIFile( sTargetFolder, sTempFileName )

	If p_tics_on = 1 Then
 		add_the_markers (o, sTargetFolder, sTargetName)
 	End If
	'plot as smith chart
	SelectTreeItem( sTargetFolder+"\"+sTargetName )
	Plot1D.PlotView( "Smith" )

End Sub

Sub add_the_markers (o As Object , sTargetFolder As String, sTargetName As String)

	Dim number_of_markers As Integer, i As Integer, phi1 As Double, phi2 As Double
 	Dim addmarker() As Double, marker_loop As Integer
 	number_of_markers = 20
 	ReDim addmarker (number_of_markers)
 	Dim marker_index As Integer, f1 As Double, f2 As Double, yre As Double, yim As Double

	SelectTreeItem( sTargetFolder+"\"+sTargetName )

 	For i = 0 To o.getN-2
 		o.GetDataFromIndex( i, f1, yre, yim )
 		phi1= atn2(yim,yre)*180/pi
 		o.GetDataFromIndex( i+1, f2, yre, yim )
 		phi2= atn2(yim,yre)*180/pi
 		If p_only_even_90 = 0 Then		'don't skip markers at +/-j
         If phi1 > -90 And phi2 < -90 Then		' -j
        	'MsgBox  cstr(f1)+": "+cstr(phi1)+" "+cstr(f2)+ ": "+cstr( phi2)
			addmarker(marker_index) = interpol_deg(f1,f2,phi1,phi2,-90)
			marker_index= marker_index+1
			'MsgBox cstr(	addmarker(marker_index-1))
         End If
         If phi1 > 90 And phi2 < 90 Then		' +j
        	'MsgBox  cstr(f1)+": "+cstr(phi1)+" "+cstr(f2)+ ": "+cstr( phi2)
			addmarker(marker_index) = interpol_deg(f1,f2,phi1,phi2,90)
			marker_index= marker_index+1
			'MsgBox cstr(	addmarker(marker_index-1))
         End If
        Else
         If phi1 > 0 And phi1 < 45 And phi2 < 0 And phi2 > -45 Then		' +1
        	'MsgBox  cstr(f1)+": "+cstr(phi1)+" "+cstr(f2)+ ": "+cstr( phi2)
			addmarker(marker_index) = interpol_deg(f1,f2,phi1,phi2,0)
			marker_index= marker_index+1
			'MsgBox cstr(	addmarker(marker_index-1))
         End If
         If phi1 < -135 And phi1 >= -180 And phi2 <=180  And phi2 > 135 Then		' -1
        	'MsgBox  cstr(f1)+": "+cstr(phi1)+" "+cstr(f2)+ ": "+cstr( phi2)
			addmarker(marker_index) = interpol_deg(f1,f2,phi1+360,phi2,180)
			marker_index= marker_index+1
			'MsgBox cstr(	addmarker(marker_index-1))
         End If
        End If
 	Next i

 	If p_only_even_90 = 0 Then		'bandwidth for markers 1 and 2 at +/-j
 		ReportInformationToWindow( "Bandwidth for 1st Resonator: "+ cstr(addmarker(1)-addmarker(0))+ " "+Units.GetUnit("Frequency") )
 	Else
		ReportInformationToWindow( "Bandwidth for 2nd Resonator: "+ cstr(addmarker(2)-addmarker(0))+ " "+Units.GetUnit("Frequency") )
 	End If

 	If p_tics_on = 1 Then
 		Plot1D.DeleteAllMarker
 		For marker_loop = 0 To marker_index-1  ' plot markers
   			Plot1D.addmarker addmarker(marker_loop)
 		Next marker_loop
 	End If

End Sub

Function interpol_deg (x1 As Double, x2 As Double, p1 As Double, p2 As Double, p As Double) As Double
   interpol_deg = x1+((x2-x1)/(p2-p1))*(p - p1)
End Function



Sub Main

	ActivateScriptSettings True
	ClearScriptSettings

	Begin Dialog UserDialog 280,259,"Smith Chart: Phase Rotation",.DialogFunction ' %GRID:10,7,1,1
		PushButton 10,217,130,21,"Perform Rotation",.phasechanged
		PushButton 150,217,130,21,"Delete all Markers",.delete_marker
		CancelButton 10,238,130,21
		PushButton 150,238,130,21,"Help",.Help
		GroupBox 10,56,270,49,"Phase",.Phase
		GroupBox 10,105,270,63,"Marker",.Marker
		GroupBox 10,7,270,42,"Port",.Port
		Text 20,28,80,14,"Port-Nr:",.Text5
		Text 150,28,60,14,"Mode-Nr:",.Text6
		DropListBox 80,21,60,121,PortNumberArray(),.portnr
		DropListBox 220,21,50,121,ModeNumberArray(),.modenr
		Text 20,77,140,21,"Phase Offset (in deg)",.Text2
		TextBox 170,70,100,21,.phase_offset
		CheckBox 20,119,160,21,"Markers ON",.tics_on
		CheckBox 20,140,200,21,"Skip Markers at +/- 90 deg",.only_even_90
		GroupBox 10,175,270,42,"S-Parameter",.GroupBox1
		CheckBox 20,196,140,14,"Use AR-Results",.ar_on
	End Dialog
	Dim dlg As UserDialog

	Dim getnofp As Integer
	Dim cst_index_sm As Integer
	Dim param_exist_flag As Boolean

	' set defaults
	dlg.phase_offset = "0"
	dlg.tics_on = 1
	'dlg.nr_of_markers = "10"
	dlg.ar_on =1
	dlg.only_even_90 = 1

	getnofp = getNumberofparameters
	'run thru parameters to find out settings
	For cst_index_sm = 0 To getnofp
	 If getparametername (cst_index_sm) = "smith_phase_offset" Then
	  dlg.phase_offset = restoredoubleparameter ("smith_phase_offset")
	  param_exist_flag=True
	 Else
	  param_exist_flag=False
	 End If
	Next cst_index_sm

	'Dialog Window
      If param_exist_flag = True Then
	   dlg.phase_offset = restoredoubleparameter ("smith_phase_offset")
	  End If

	'Dialog Window
    Do
      If param_exist_flag = True Then
	   dlg.phase_offset = restoredoubleparameter ("smith_phase_offset")
	  End If
	  If Dialog(dlg)=0   Then Exit All
    Loop Until (dlg.phase_offset <> "")

End Sub

Sub plot_it_now

	Dim no_modes_flag As Boolean, ar_on_string As String, sTreeItem As String,   theta As Double,  sTargetFolder As String,  sTargetName As String

	'compose the tree-name
	If p_ar_on = 1 And (selecttreeitem( "1D Results\S-Parameters (AR)"))  Then
		ar_on_string= " (AR)"
	Else
		ar_on_string=""
	End If

	sTreeItem = "1D Results\S-Parameters"+ar_on_string+"\S"+ GetScriptSetting("chirp_port","1")+"("+GetScriptSetting("chirp_mode","1") +"),"+  _
											 GetScriptSetting("chirp_port","1")+"("+GetScriptSetting("chirp_mode","1") +")"
	If Not selecttreeitem (sTreeItem) Then	'if mode indices not available
		sTreeItem = "1D Results\S-Parameters"+ar_on_string+"\S"+ GetScriptSetting("chirp_port","1")+","+GetScriptSetting("chirp_port","1")
		no_modes_flag= True
	Else
		no_modes_flag = False
	End If

	theta = cdbl(p_phase_offset)*pi/180
	ReportInformationToWindow( "Rotating '" + sTreeItem + "' by "  + CStr( p_phase_offset )  + " degrees." )

	sTargetFolder = "1D Results\Smith_Chart_Rotated"+ar_on_string
	sTargetName ="S"+ GetScriptSetting("chirp_port","1")+"("+GetScriptSetting("chirp_mode","1") +"),"+  _
											 GetScriptSetting("chirp_port","1")+"("+GetScriptSetting("chirp_mode","1") +")_rotated"

	If no_modes_flag Then	'mode indices not available
		sTargetName ="S"+ GetScriptSetting("chirp_port","1")+ ","+ GetScriptSetting("chirp_port","1") +"_rotated"
	End If

	StoreDoubleParameter "smith_phase_offset", GetDouble_new(p_phase_offset)

	With Resulttree
		.Name sTargetFolder+"\"+sTargetName
		.delete
	End With

    StoreScriptSetting("complete_folder_info",sTargetFolder+"\"+sTargetName)
    CreateRotatedSmithChartFromTreeItem( sTreeItem, theta, sTargetFolder, sTargetName )

End Sub

Sub  delete_the_markers()
	Dim sTargetName As String
	sTargetName = GetScriptSetting("complete_folder_info","1D Results\Smith_Chart_Rotated\S1,1_rotated")
	selecttreeitem sTargetName
    Plot1D.DeleteAllMarker
	Plot1D.plotview "smith"
End Sub



'Function DialogFunc%(DlgItem As String, Action As Integer, SuppValue As Integer)
	Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean


    Dim m As String
    Dim file As String
    Dim basepath As String

    Debug.Print "Action=";Action
    Debug.Print DlgItem
    Debug.Print "SuppValue=";SuppValue

    Select Case Action%
    Case 1 ' Dialog box initialization
    	FillPortNumberArray()
		FillModeNumberArray (GetScriptSetting("chirp_port", PortNumberArray(0)))
		DlgListBoxArray("portnr", PortNumberArray)
		DlgListBoxArray("modenr", ModeNumberArray)

		DlgValue("portnr", FindListIndex(PortNumberArray(), GetScriptSetting("chirp_port","1")))
		If DlgValue("portnr")<0 Or DlgValue("portnr")>UBound(PortNumberArray) Then DlgValue("portnr", 0)
		DlgValue("modenr", FindListIndex(ModeNumberArray(), (GetScriptSetting("chirp_mode","1")))) '1
		If DlgValue("modenr")<0 Or DlgValue("modenr")>UBound(ModeNumberArray) Then DlgValue("modenr", 0)

    Case 2 ' Value changing or button pressed
      Select Case DlgItem
		Case "Help"
			StartHelp HelpFileName
			DialogFunction = True

        Case "delete_marker" 	'example for other entries, e.g. help is quite useful here

			 delete_the_markers
            DialogFunction = True 				'do not exit the dialog

        Case "portnr"
            FillModeNumberArray(PortNumberArray(DlgValue("portnr")))
			DlgListBoxArray("modenr", ModeNumberArray)
			DlgValue("modenr", 0)

        Case "phasechanged"
            'MsgBox "addplot"

            StoreScriptSetting("chirp_port",PortNumberArray(DlgValue("portnr")))
			StoreScriptSetting("chirp_mode",ModeNumberArray(DlgValue("modenr")))

            p_portnr = DlgText("portnr")
            p_modenr = DlgValue ("modenr")
            p_only_even_90 = DlgValue ("only_even_90")
            p_tics_on = DlgValue ("tics_on")
            p_phase_offset = DlgText ("phase_offset")
            p_ar_on = DlgValue ("ar_on")

            plot_it_now

            DialogFunction=True


            Case Else
            'Beep
        	'
      End Select
    Case 3 ' Combo or text value changed
      'MsgBox "combo-box"
      DialogFunction = True 				'do not exit the dialog
	  'm = DlgText ("ComboBox1")
	  'MsgBox m
    Case 4 ' Focus changed
       Debug.Print "DlgFocus=""";DlgFocus();""""
    End Select
End Function


Function GetDouble_new(value As String) As Double
	Dim cst_separator    As String
	cst_separator    = Mid$(CStr(0.5), 2, 1)
	GetDouble_new        = CVar(CDbl(Replace(value, ".", cst_separator)))
End Function
