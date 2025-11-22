
Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        If (DlgItem = "PushButton1") Then
        
          DlgText "Xmin", RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XLow")
          DlgText "Xmax", RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XHigh")
          DlgText "Ymin", RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YLow")
          DlgText "Ymax", RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YHigh")
  
          dialogfunc = True
        End If
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function

Sub Main ()
 
 	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmin", ""
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmax", ""
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymin", ""
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymax", ""

    Begin Dialog UserDialog 400,154,"Define Substrate Dimensions", .dialogfunc ' %GRID:10,6,1,1
		Text 20,14,90,14,"Xmin",.Text1
		Text 210,14,90,14,"Xmax",.Text2
		Text 20,63,90,14,"Ymin",.Text3
		Text 210,63,90,14,"Ymax",.Text4
		TextBox 20,30,160,18,.Xmin
		TextBox 20,78,160,18,.Ymin
		TextBox 210,30,160,18,.Xmax
		TextBox 210,78,160,18,.Ymax
		OKButton 20,126,90,21
		CancelButton 120,126,90,21
		PushButton 220,126,90,21,"Reset",.PushButton1
	End Dialog
	Dim dlg As UserDialog
	
	dlg.Xmin = Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XLow"), ",", ".")
	dlg.Xmax = Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XHigh"), ",", ".")
	dlg.Ymin = Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YLow"), ",", ".")
	dlg.Ymax = Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YHigh"), ",", ".")   
  
	On Error GoTo Finish
	Dialog dlg
	
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmin", dlg.Xmin
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmax", dlg.Xmax
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymin", dlg.Ymin
	StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymax", dlg.Ymax
	
Finish:

End Sub
