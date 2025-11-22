' Define Substrate Dimensions
' !!! Do not change the line above !!!

' 24-Jul-2009 ube: GetMacroPath replaced by GetInstallPath + "\Library\Macros" (pervisouly only first macropath was search)

Sub Main ()

    Dim sXmin As String, sXmax As String, sYmin As String, sYmax As String

    BeginHide
    
    RunScript GetInstallPath + "\Library\Macros\File\ADS Converter 3.bas"
    
	sXmin = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Xmin")
	sXmax = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Xmax")
	sYmin = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Ymin")
	sYmax = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Ymax")

    Assign "sXmin"
    Assign "sXmax"
    Assign "sYmin"
    Assign "sYmax"
    EndHide
    
    If (sXmin <> "" And sXmax <> "" And sYmin <> "" And sYmax <> "") Then
    
		StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmin", sXmin
		StoreGlobalDataValue "Macros\ADS Import\New Bounds\Xmax", sXmax
		StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymin", sYmin
		StoreGlobalDataValue "Macros\ADS Import\New Bounds\Ymax", sYmax
		
	    RunScript GetInstallPath + "\Library\Macros\File\ADS Converter 2.bas"
	End If

End Sub
