
Sub BuildSubstrateBricks()
  Dim dXL As Double, dXH As Double, dYL As Double, dYH As Double
  
  nLayers = CInt(RestoreGlobalDataValue("Macros\ADS Import\Substrates\Count"))
  RestoreBBoxBounds dXL, dXH, dYL, dYH
  
  Dim i As Integer
  Dim dFrom As Double
  Dim dTo As Double
  
  dFrom = 0.0
  dTo = 0.0
  
  For i=nLayers-1 To 0 STEP -1
    Dim dThick As Double
    dThick = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Substrates\Thickness" + CStr(i)))
    
    If (dThick <> 0.0) Then
    
	    dFrom = dTo
	    dTo   = dFrom + dThick
	    
	    With Brick
		     .Reset 
		     .Name "solid1" 
		     .Layer "Substrate"  + CStr(i)
		     .Xrange CStr(dXL), CStr(dXH) 
		     .Yrange CStr(dYL), CStr(dYH) 
		     .Zrange CStr(dFrom), CStr(dTo)
		     .Create
		End With
		
	End If
  
  Next i
  
End Sub

Sub RestoreStructureBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  dXL = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XLow"))
  dXH = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XHigh"))
  dYL = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YLow"))
  dYH = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YHigh"))
End Sub

Sub RestoreBBoxBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  dXL = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XLow"))
  dXH = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XHigh"))
  dYL = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YLow"))
  dYH = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YHigh"))
End Sub

Sub StoreBBoxBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XLow", CStr(dXL)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XHigh", CStr(dXH)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YLow", CStr(dYL)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YHigh", CStr(dYH)
End Sub


Function GetShapeNameFromPoint(dX As Double, dY As Double, dZ As Double) As String

  Dim nShapes As Integer, i As Integer
  
  nShapes = Solid.GetNumberOfShapes()
  Dim sName As String
  sName = ""
  
  For i=0 To nShapes-1
    Dim sShapeName As String
    sShapeName = Solid.GetNameOfShapeFromIndex(i)
    
    If (Solid.IsPointInsideShape(dX, dY, dZ, sShapeName)) Then
      sName = sShapeName
    End If
  
  Next i 

  GetShapeNameFromPoint = sName
	
End Function

Sub Main ()

    Dim sXmin As String, sXmax As String, sYmin As String, sYmax As String

	sXmin = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Xmin")
	sXmax = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Xmax")
	sYmin = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Ymin")
	sYmax = RestoreGlobalDataValue("Macros\ADS Import\New Bounds\Ymax")
	
	Dim dDXMax As Double, dDXMin As Double, dDYMax As Double, dDYMin As Double
	dDXMax = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XHigh") + "- (" + sXmax + ")", ",", "."))
	dDXMin = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XLow") + "- (" + sXmin + ")", ",", "."))
	dDYMax = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YHigh") + "- (" + sYmax + ")", ",", "."))
	dDYMin = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YLow") + "- (" + sYmin + ")", ",", "."))
		
	Dim nPort As Integer
	Dim dX As Double, dY As Double, dZ As Double
	
	ResultTree.EnableTreeUpdate False
	ScreenUpdating False
	SetLock True

 	If (RestoreGlobalDataValue("Macros\ADS Import\Ports\XMin") <> "0" And dDXMin <> 0.0) Then
   	  nPort = CInt(RestoreGlobalDataValue("Macros\ADS Import\Ports\PNXMin"))
	  dX = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XLow"), ",", "."))
	  dY = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Y"))
	  dZ = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Z"))
	  
	  Pick.ClearAllPicks 
	  Pick.PickFaceFromPoint GetShapeNameFromPoint(dX, dY, dZ), dX, dY, dZ 
	  Solid.OffsetSelectedFaces CStr(dDXMin)
	End If
	
 	If (RestoreGlobalDataValue("Macros\ADS Import\Ports\XMax") <> "0" And dDXMax <> 0.0) Then
   	  nPort = CInt(RestoreGlobalDataValue("Macros\ADS Import\Ports\PNXMax"))
	  dX = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XHigh"), ",", "."))
	  dY = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Y"))
	  dZ = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Z"))
	  
	  Pick.ClearAllPicks 
	  Pick.PickFaceFromPoint GetShapeNameFromPoint(dX, dY, dZ), dX, dY, dZ 
	  Solid.OffsetSelectedFaces CStr(-dDXMax)
	End If
	
 	If (RestoreGlobalDataValue("Macros\ADS Import\Ports\YMin") <> "0" And dDYMin <> 0.0) Then
   	  nPort = CInt(RestoreGlobalDataValue("Macros\ADS Import\Ports\PNYMin"))
	  dY = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YLow"), ",", "."))
	  dX = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "X"))
	  dZ = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Z"))
	  
	  Pick.ClearAllPicks 
	  Pick.PickFaceFromPoint GetShapeNameFromPoint(dX, dY, dZ), dX, dY, dZ 
	  Solid.OffsetSelectedFaces CStr(dDYMin)
	End If
	
 	If (RestoreGlobalDataValue("Macros\ADS Import\Ports\YMax") <> "0" And dDYMax <> 0.0) Then
   	  nPort = CInt(RestoreGlobalDataValue("Macros\ADS Import\Ports\PNYMax"))
	  dY = Evaluate(Replace(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YHigh"), ",", "."))
	  dX = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "X"))
	  dZ = Evaluate(RestoreGlobalDataValue("Macros\ADS Import\Ports\P" + CStr(nPort) + "Z"))
	  
	  Pick.ClearAllPicks 
	  Pick.PickFaceFromPoint GetShapeNameFromPoint(dX, dY, dZ), dX, dY, dZ 
	  Solid.OffsetSelectedFaces CStr(-dDYMax)
	End If
	
	StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XLow",  sXmin
	StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XHigh", sXmax
	StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YLow",  sYmin
	StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YHigh", sYmax

	BuildSubstrateBricks

	ResultTree.EnableTreeUpdate True
	ScreenUpdating True
	SetLock False
	  
Finish:  
End Sub
