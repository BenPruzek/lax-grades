
Option Explicit

' This function reads a sequence of characters from the given file stream
' until a specified delimit character is reached.

Dim bReplaceDotByColon As Boolean
Dim bShowOhmicSheetWarning As Boolean

Function ReadBlock(nFileStream As Integer, cLimit1 As String, cLimit2 As String) As String

	Dim bRead As Boolean, c As String, sContents As String

	bRead = True

	On Error GoTo Finish

    Do
       c = Input$(nFileStream, 1)
	   If (c = cLimit1 Or c = cLimit2) Then
	     bRead = False
	   Else
	     If (c <> Chr(10) And c <> Chr(13)) Then
	       sContents = sContents + c
	     Else
	       sContents = sContents + " "
	     End If
	   End If

	Loop Until bRead = False

	Finish:

	If (bRead) Then
	  sContents = ""
	End If

	ReadBlock = LCase(sContents)

End Function

Function GetItem(sLine As String)

  sLine = Trim(sLine)

  Dim nIndex As Integer, sItem As String

  sLine = Replace(sLine, Chr(9), " ")

  nIndex = InStr(sLine, " ")

  If (nIndex = 0) Then
    sItem = sLine
    Dim nSemicolon As Integer
	nSemicolon = InStr(sItem, ";")
	If (nSemicolon <> 0) Then
		sItem = Left(sItem, nSemicolon-1)
	End If
    sLine = ""
  Else
    sItem = Left(sLine, nIndex)
    sLine = Mid(sLine, nIndex+1)
  End If

  sItem = Trim(sItem)

  GetItem = sItem

End Function

Function GetDouble(sItem As String) As Double

  If (bReplaceDotByColon) Then
    sItem = Replace(sItem, ".", ",")
  End If

  sItem = Trim(sItem)

  If (sItem = "infinity") Then
    GetDouble = 1e30
  Else
    GetDouble = CVar(sItem)
  End If

End Function

Function GetInt(sItem As String) As Long

    GetInt = CVar(sItem)

End Function

Function GetValueForCommand(sLine As String, sCommand As String) As String

	Dim nIndex As Integer, sTmpLine As String, sValue As String
	Dim nIndexLeftBracket As Integer, nIndexRightBracket As Integer
	sValue = ""
	sTmpLine = sLine

	If (sTmpLine <> "") Then
		nIndex = InStr(sTmpLine, sCommand)
		If (nIndex <> 0) Then
			sTmpLine = Trim(Mid(sTmpLine, nIndex))

 			nIndex = InStr(sTmpLine, "=")
 			If (nIndex <> 0) Then
				sTmpLine = Trim(Mid(sTmpLine, nIndex+1))
				nIndex = InStr(sTmpLine, " ")
				nIndexLeftBracket = InStr(sTmpLine, "{")
				nIndexRightBracket = InStr(sTmpLine, "}")

 				If (nIndex <> 0) Then
 					If ((nIndexLeftBracket < nIndex) And (nIndex < nIndexRightBracket) ) Then
 						sValue = Trim(Left(sTmpLine, nIndexRightBracket))
 					Else
						sValue = Trim(Left(sTmpLine, nIndex-1))
					End If
				Else
					sValue = sTmpLine
 				End If
 			End If
		End If
	End If

 	GetValueForCommand = sValue

End Function

Function GetMaskIndexById(nMask As Integer, nMaskId() As Integer, nMaskIdToFind As Integer) As Integer

	Dim nIndex As Integer, nCounter As Integer
	nIndex = -1
	nCounter = 0

	While (nCounter < nMask)
		If( nMaskId(nCounter) = nMaskIdToFind ) Then
			nIndex = nCounter
			nCounter = nMask
		End If
		nCounter = nCounter+1
	Wend

	GetMaskIndexById = nIndex

End Function

Function GetMaskIdByName(nMask As Integer, nMaskId() As Integer, nMaskName() As String, nMaskNameToFind As String) As Integer

	Dim nIndex As Integer, nCounter As Integer
	nIndex = -1
	nCounter = 0

	While (nCounter < nMask)
		If( nMaskName(nCounter) = nMaskNameToFind ) Then
			nIndex = nCounter
			nCounter = nMask
		End If
		nCounter = nCounter+1
	Wend

	GetMaskIdByName = nMaskId(nIndex)

End Function

Function GetMaterialIndexByName(nMaterial As Integer, sMaterialName() As String, sMaterialNameToFind As String) As Integer

	Dim nIndex As Integer, nCounter As Integer
	nIndex = -1
	nCounter = 0

	While (nCounter < nMaterial)
		If( sMaterialName(nCounter) = sMaterialNameToFind ) Then
			nIndex = nCounter
			nCounter = nMaterial
		End If
		nCounter = nCounter+1
	Wend

	GetMaterialIndexByName = nIndex

End Function

Function GetOperationIndexByName(nOperation As Integer, sOperationName() As String, sOperationNameToFind As String) As Integer

	Dim nIndex As Integer, nCounter As Integer
	nIndex = -1
	nCounter = 0

	While (nCounter < nOperation)
		If( sOperationName(nCounter) = sOperationNameToFind ) Then
			nIndex = nCounter
			nCounter = nOperation
		End If
		nCounter = nCounter+1
	Wend

	GetOperationIndexByName = nIndex

End Function


Function GetStackType(sStackLine As String) As String

	Dim nIndex As Integer, sType As String
	sType = ""

	nIndex = InStr(sStackLine, " ")
	If (nIndex <> 0) Then
	  sType = Trim(Left(sStackLine, nIndex-1))
	  sStackLine = Trim(Mid(sStackLine, nIndex+1))
	End If

	GetStackType = sType

End Function

Function ScanStackLayer(sLayerLine As String, sLayerName As String, dHeight As Double, sMaterialName As String, nMaskId As Integer, sMaskName As String, sMaskValues As String)

	Dim sMaskValue As String
	Dim sMaskValuesToProcess As String

	sLayerName = GetValueForCommand(sLayerLine, "name")
	dHeight = GetDouble(GetValueForCommand(sLayerLine, "height"))
	sMaterialName = GetValueForCommand(sLayerLine, "material")

	If (sMaskValues = "") Then
    	sMaskValues = GetValueForCommand(sLayerLine, "mask")
    End If

	If (sMaskValues <> "") Then
		' List of masks
		Dim nIndex As Integer
		nIndex = InStr(sMaskValues, " ")
		If (nIndex <> 0) Then
			sMaskValue = Trim(Left(sMaskValues, nIndex-1)) + "}"
			sMaskValuesToProcess = "{" + Trim(Mid(sMaskValues, nIndex+1))
		Else
			sMaskValue = sMaskValues
		End If

		' Entferne geschweifte Klammern
		sMaskValue = Trim(Mid(sMaskValue, 2))
		sMaskValue = Trim(Left(sMaskValue, Len(sMaskValue)-1))

		If IsNumeric(sMaskValue) Then
			nMaskId = GetInt(sMaskValue)
		Else
			sMaskName = sMaskValue
		End If

	Else
		nMaskId = -1
	End If

	ScanStackLayer = sMaskValuesToProcess

End Function

Function ScanStackInterface(sInterfaceLine As String, sInterfaceName As String, nMaskId As Integer, sMaskName As String, sMaskValues As String)

	Dim sMaskValue As String
	Dim sMaskValuesToProcess As String
	
	sInterfaceName = GetValueForCommand(sInterfaceLine, "name")

	If (sMaskValues = "") Then
    		sMaskValues = GetValueForCommand(sInterfaceLine, "mask")
    	End If
    	
	If (sMaskValues <> "") Then
		' List of masks
		Dim nIndex As Integer	
		nIndex = InStr(sMaskValues, " ")
		If (nIndex <> 0) Then
			sMaskValue = Trim(Left(sMaskValues, nIndex-1)) + "}"
			sMaskValuesToProcess = "{" + Trim(Mid(sMaskValues, nIndex+1))
		Else
			sMaskValue = sMaskValues
		End If
		
		' Entferne geschweifte Klammern
		sMaskValue = Trim(Mid(sMaskValue, 2))
		sMaskValue = Trim(Left(sMaskValue, Len(sMaskValue)-1))

		If IsNumeric(sMaskValue) Then
			nMaskId = GetInt(sMaskValue)
		Else
			sMaskName = sMaskValue
		End If
	Else
		nMaskId = -1
	End If
	
	ScanStackInterface = sMaskValuesToProcess

End Function

Sub ScanMaterial(sMaterial As String, nMaterial As Integer, sMaterialName() As String, dMaterialEpsilon() As Double, dMaterialMue() As Double, dMaterialConductivity() As Double, dMaterialLossTangent() As Double)

	Dim nIndex As Integer, sName As String, dE As Double, dM As Double, dCond As Double, dLossTan As Double
	sName = ""
	dE = 1.0
	dM = 1.0
	dCond = 0.0
	dLossTan = 0.0

	On Error GoTo Finish

	nIndex = InStr(sMaterial, " ")
	If (nIndex = 0) Then
		GoTo Finish
	End If

	sMaterial = Trim(Mid(sMaterial, nIndex+1))
    nIndex = InStr(sMaterial, " ")
    If (nIndex = 0) Then
		GoTo Finish
	End If

    sName = Trim(Left(sMaterial, nIndex-1))

    Dim sTmp As String
    sTmp = GetValueForCommand(sMaterial, "permittivity")
    If (sTmp <> "") Then
    	dE = GetDouble(sTmp)
    End If

    sTmp = GetValueForCommand(sMaterial, "permeability")
    If (sTmp <> "") Then
    	dM = GetDouble(sTmp)
    End If

    sTmp = GetValueForCommand(sMaterial, "conductivity")
    If (sTmp <> "") Then
    	dCond = GetDouble(sTmp)
    End If

    sTmp = GetValueForCommand(sMaterial, "losstangent")
    If (sTmp <> "") Then
    	dLossTan = GetDouble(sTmp)
    End If

    Finish:

    If (sName <> "") Then
    	ReDim Preserve sMaterialName(nMaterial)
      	ReDim Preserve dMaterialEpsilon(nMaterial)
        ReDim Preserve dMaterialMue(nMaterial)
        ReDim Preserve dMaterialConductivity(nMaterial)
        ReDim Preserve dMaterialLossTangent(nMaterial)

	  	sMaterialName(nMaterial) = sName
	  	dMaterialEpsilon(nMaterial) = dE
	  	dMaterialMue(nMaterial) = dM
	  	dMaterialConductivity(nMaterial) = dCond
	  	dMaterialLossTangent(nMaterial) = dLossTan

	  	nMaterial = nMaterial + 1
    End If

End Sub

Sub ScanMask(sMaskLine As String, nMask As Integer, nMaskId() As Integer, sMaskType() As String, sMaskName() As String, sMaskMaterial() As String, sMaskOperation() As String)

	Dim nIndex As Integer, nId As Integer, sName As String, sType As String, sMat As String, sOperator As String
	nId = -1
	sName = ""
	sType = "positive"
	sMat = ""
	sOperator = ""

	On Error GoTo Finish

	nIndex = InStr(sMaskLine, " ")
	If (nIndex = 0) Then
		GoTo Finish
	End If

	sMaskLine = Trim(Mid(sMaskLine, nIndex+1))
    nIndex = InStr(sMaskLine, " ")
    If (nIndex = 0) Then
		GoTo Finish
	End If

    nId = GetInt(Trim(Left(sMaskLine, nIndex-1)))

    nIndex = InStr(sMaskLine, "negative")
    If (nIndex <> 0) Then
		sType = "negative"
	End If

    sName = GetValueForCommand(sMaskLine, "name")
    sMat = GetValueForCommand(sMaskLine, "material")
    sOperator = GetValueForCommand(sMaskLine, "operation")

    Finish:

    If (nId <> -1) Then
    	ReDim Preserve nMaskId(nMask)
    	ReDim Preserve sMaskName(nMask)
    	ReDim Preserve sMaskType(nMask)
      	ReDim Preserve sMaskMaterial(nMask)
        ReDim Preserve sMaskOperation(nMask)

	  	nMaskId(nMask) = nId
	  	sMaskName(nMask) = sName
	  	sMaskType(nMask) = sType
	  	sMaskMaterial(nMask) = sMat
	  	sMaskOperation(nMask) = sOperator

	  	nMask = nMask + 1
    End If

End Sub

Sub ScanOperation(sOperLine As String, nOperation As Integer, sOperationName() As String, sOperationType() As String, dOperationThickness() As Double)

	Dim nIndex As Integer, sName As String, sType As String, dThickness As Double
	sName = ""
	sType = ""
	dThickness = 0.0

	On Error GoTo Finish

	nIndex = InStr(sOperLine, " ")
	If (nIndex = 0) Then
		GoTo Finish
	End If

	sOperLine = Trim(Mid(sOperLine, nIndex+1))
    nIndex = InStr(sOperLine, " ")
    If (nIndex = 0) Then
		GoTo Finish
	End If

    sName = Trim(Left(sOperLine, nIndex-1))

    sOperLine = Trim(Mid(sOperLine, nIndex+1))
    nIndex = InStr(sOperLine, " ")
    If (nIndex = 0) Then
		sType = sOperLine
	Else
		sType = Trim(Left(sOperLine, nIndex-1))
	End If

	nIndex = InStr(sType, "intrude")
	If (nIndex <> 0) Then
		sType = "intrude"
    	dThickness = GetDouble(GetValueForCommand(sOperLine, "intrude"))
    	'If (dThickness <= 1e-6) Then
    		'dThickness = 0.0
    	'End If
    End If
    nIndex = InStr(sType, "expand")
	If (nIndex <> 0) Then
		sType = "expand"
    	dThickness = GetDouble(GetValueForCommand(sOperLine, "expand"))
    End If

    If IsCSTVersionOfCurrentBlock_GreaterEqualThan(2011, 6) And (sType = "intrude" Or sType = "expand") Then
    	nIndex = InStr(sOperLine, "down")
		If (nIndex <> 0) Then
			dThickness = - dThickness
		End If
    End If

    Finish:

    If (sName <> "") Then
    	ReDim Preserve sOperationName(nOperation)
    	ReDim Preserve sOperationType(nOperation)
    	ReDim Preserve dOperationThickness(nOperation)

	  	sOperationName(nOperation) = sName
	  	sOperationType(nOperation) = sType
	  	dOperationThickness(nOperation) = dThickness

	  	nOperation = nOperation + 1
    End If
End Sub


Sub ScanLayerBlock(sBlock As String, dScale As Double, dThickness() As Double, dEpsilon() As Double, dMue() As Double, dLossTang() As Double, dLossCond() As Double)

  Dim nLayer As Integer, dT As Double, dE As Double, dM As Double, dL As Double

  nLayer = GetInt(GetItem(sBlock))
  GetItem(sBlock)
  dT = GetDouble(GetItem(sBlock))

  If (dT > 1e10) Then dT = 0.0

  Dim sLossType As String

  GetItem(sBlock)
  sLossType = Trim(GetItem(sBlock))
  dE = GetDouble(GetItem(sBlock))
  dL = GetDouble(GetItem(sBlock))

  GetItem(sBlock)
  GetItem(sBlock)
  dM = GetDouble(GetItem(sBlock))

  If (nLayer = 0) Then
    ReDim Preserve dThickness(0)
    ReDim Preserve dEpsilon(0)
    ReDim Preserve dMue(0)
    ReDim Preserve dLossTang(0)
    ReDim Preserve dLossCond(0)
  ElseIf (nLayer > UBound(dThickness)-1) Then
    ReDim Preserve dThickness(nLayer)
    ReDim Preserve dEpsilon(nLayer)
    ReDim Preserve dMue(nLayer)
    ReDim Preserve dLossTang(nLayer)
    ReDim Preserve dLossCond(nLayer)
  End If

  dThickness(nLayer) = dT * dScale
  dEpsilon(nLayer) = dE
  dMue(nLayer) = dM

  If sLossType = "losstangent" Then
	dLossTang(nLayer) = dL
	dLossCond(nLayer) = 0.0
  ElseIf sLossType = "conductivity" Then
	dLossTang(nLayer) = 0.0
	dLossCond(nLayer) = dL
  End If

End Sub

Function GetScaleFromDropboxNumber(sUnit As String)

	sUnit = Trim(sUnit)
    Dim dScale As Double
    dScale = 1.0

    Select Case sUnit
    Case "0" '"nm"
      dScale = 1e-9
    Case "1" '"um"
      dScale = 1e-6
    Case "2" '"mm"
      dScale = 1e-3
    Case "3" '"cm"
      dScale = 1e-2
    Case "4" '"metre"
      dScale = 1.0
    'Case "km"
    '  dScale = 1e3
    'Case "nin"
    '  dScale = 1e-9 * 2.54e-2
    'Case "uinch"
    '  dScale = 1e-6 * 2.54e-2
    Case "5" '"mil"
      dScale = 1e-3 * 2.54e-2
    Case "6" '"inch"
      dScale = 2.54e-2
    'Case "foot"
    '  dScale = 0.0
    'Case "yard"
    '  dScale = 0.0
    'Case "mile"
    '  dScale = 0.0
    End Select

  GetScaleFromDropboxNumber = dScale * 1e6

End Function

Function GetScaleFromName(sUnit As String)

    sUnit = Trim(sUnit)
    Dim dScale As Double
    dScale = 1.0

    Select Case sUnit
    Case "nm"
      dScale = 1e-9
    Case "um"
      dScale = 1e-6
    Case "mm"
      dScale = 1e-3
    Case "cm"
      dScale = 1e-2
    Case "metre"
      dScale = 1.0
    Case "km"
      dScale = 1e3
    Case "nin"
      dScale = 1e-9 * 2.54e-2
    Case "uinch"
      dScale = 1e-6 * 2.54e-2
    Case "mil"
      dScale = 1e-3 * 2.54e-2
    Case "inch"
      dScale = 2.54e-2
    Case "foot"
      dScale = 0.0
    Case "yard"
      dScale = 0.0
    Case "mile"
      dScale = 0.0
    End Select

  GetScaleFromName = dScale * 1e6

End Function

' This subroutine reads the substrate file and fills the arrays with thickness, permittivity, permeability

Sub ReadLayers(sFileName As String, dThickness() As Double, dEpsilon() As Double, dMue() As Double, dLossTang() As Double, dLossCond() As Double, sTopPlane As String, sBottomPlane As String)

  Dim sLine As String

  Open sFileName For Input As #1

  On Error GoTo Finish
  Line Input #1, sLine

  Dim dScale As Double
  dScale = 1.0

  While True

    sLine = LCase(Trim(sLine))

    If (sLine <> "") Then
      Dim sCommand As String, nIndex As Integer

      nIndex = InStr(sLine, " ")
      If (nIndex = 0) Then
        sCommand = Trim(sLine)
      Else
        sCommand = Trim(Left(sLine, nIndex))
      End If

      If (sCommand = "units") Then
        Dim sUnit As String
        sUnit = Trim(Mid(sLine, nIndex))

		dScale = GetScaleFromName(sUnit)

      ElseIf (sCommand = "layers") Then

        Do
	        Dim sBlock As String
	        sBlock = Trim(ReadBlock(1, ",", ";"))

	        If (sBlock <> "") Then
	          Dim sTmpBlock
	          sTmpBlock = sBlock

	          ScanLayerBlock sTmpBlock, dScale, dThickness(), dEpsilon(), dMue(), dLossTang(), dLossCond()

	        End If

        Loop Until sBlock = ""

      ElseIf (sCommand = "topplane") Then

        sTopPlane = Trim(Mid(sLine, nIndex))
        If (sTopPlane <> "open") Then
	        sTopPlane = Trim(Left(sTopPlane, InStr(sTopPlane, " ")))
	    End If

      ElseIf (sCommand = "bottomplane") Then

        sBottomPlane = Trim(Mid(sLine, nIndex))
        If (sBottomPlane <> "open") Then
        	sBottomPlane = Trim(Left(sBottomPlane, InStr(sBottomPlane, " ")))
        End If

      End If

    End If

    Line Input #1, sLine
  Wend

  Finish:
  Close #1

End Sub

Sub ReadLayersAndPlanes(sFileName As String, sLayerNames() As String, dThickness() As Double, dEpsilon() As Double, dMue() As Double, dLossTang() As Double, dLossCond() As Double, sTopPlane As String, sBottomPlane As String, nPlanes As Integer, nPlaneId() As Integer, nLayerFrom() As Integer, sPlaneTypes() As String, dPlaneCond() As Double, dPlaneImpedance() As Double, dPlaneThickness() As Double, nVias As Integer, nViaId() As Integer, nViaStart() As Integer, nViaEnd() As Integer, bBoundaryOutlineDefined As Boolean)

  Dim sLine As String

  Open sFileName For Input As #1

  On Error GoTo Finish
  Line Input #1, sLine

  Dim dScale As Double
  dScale = 1.0

  ' Material Info
  Dim nMaterial As Integer, sMaterialName() As String, dMaterialEpsilon() As Double, dMaterialMue() As Double, dMaterialConductivity() As Double, dMaterialLossTangent() As Double
  ' Mask Info
  Dim nMask As Integer, nMaskId() As Integer, sMaskType() As String, sMaskName() As String, sMaskMaterial() As String, sMaskOperation() As String
  ' Operation Info
  Dim nOperation As Integer, sOperationName() As String, sOperationType() As String, dOperationThickness() As Double
  ' Layer Info
  Dim nLayer As Integer

  nLayer = 0
  nMaterial = 2
  ReDim Preserve sMaterialName(nMaterial)
  ReDim Preserve dMaterialEpsilon(nMaterial)
  ReDim Preserve dMaterialMue(nMaterial)
  ReDim Preserve dMaterialConductivity(nMaterial)
  ReDim Preserve dMaterialLossTangent(nMaterial)

  sMaterialName(0) = "air"
  dMaterialEpsilon(0) = 1.0
  dMaterialMue(0) = 1.0
  dMaterialConductivity(0) = 0.0
  dMaterialLossTangent(0) = 0.0

  sMaterialName(1) = "perfect_conductor"
  dMaterialEpsilon(1) = 1.0
  dMaterialMue(1) = 1.0
  dMaterialConductivity(1) = 0.0
  dMaterialLossTangent(1) = 0.0

  While True

    sLine = LCase(Trim(sLine))

    If (sLine <> "") Then
      Dim sCommand As String, nIndex As Integer

      nIndex = InStr(sLine, " ")
      If (nIndex = 0) Then
        sCommand = Trim(sLine)
      Else
        sCommand = Trim(Left(sLine, nIndex))
      End If

      If (sCommand = "units") Then

        Dim sUnit As String

        Do
	        Dim sUnitsLine As String
	        Line Input #1, sUnitsLine
	        sUnitsLine = LCase(Trim(sUnitsLine))

	        If (sUnitsLine = "end_units") Then
			   sUnitsLine = ""
			ElseIf (sUnit = "") Then
			   sUnit = GetValueForCommand(sUnitsLine, "distance")
	        End If
        Loop Until sUnitsLine = ""

		dScale = GetScaleFromName(sUnit)

      ElseIf (sCommand = "begin_material") Then

        Do
	        Dim sMaterial As String
	        Line Input #1, sMaterial
	        sMaterial = LCase(Trim(sMaterial))

	        If (sMaterial = "end_material") Then
			   sMaterial = ""
			Else
			  If (sMaterial <> "") Then
	          	Dim sTmpMaterial As String
	          	sTmpMaterial = sMaterial
	          	ScanMaterial sTmpMaterial, nMaterial, sMaterialName(), dMaterialEpsilon(), dMaterialMue(), dMaterialConductivity(), dMaterialLossTangent()
	          End If
	        End If

        Loop Until sMaterial = ""

      ElseIf (sCommand = "begin_mask") Then

      	nMask = 0
        Do
	        Dim sMask As String
	        Line Input #1, sMask
	        sMask = LCase(Trim(sMask))

	        If (sMask = "end_mask") Then
			   sMask = ""
			Else
			  If (sMask <> "") Then
	          	Dim sTmpMask As String
	          	sTmpMask = sMask
	          	ScanMask sTmpMask, nMask, nMaskId(), sMaskType(), sMaskName(), sMaskMaterial(), sMaskOperation()
	          End If
	        End If

        Loop Until sMask = ""

      ElseIf (sCommand = "begin_operation") Then

      	nOperation = 0
        Do
	        Dim sOperation As String
	        Line Input #1, sOperation
	        sOperation = LCase(Trim(sOperation))

	        If (sOperation = "end_operation") Then
			   sOperation = ""
			Else
			  If (sOperation <> "") Then
	          	Dim sTmpOper As String
	          	sTmpOper = sOperation
	          	ScanOperation sTmpOper, nOperation, sOperationName(), sOperationType(), dOperationThickness()
	          End If
	        End If

        Loop Until sOperation = ""

      ElseIf (sCommand = "begin_stack") Then

	    Do
	        Dim sStackLine As String
	        Dim sMaskValues As String
	        
	        If (sMaskValues = "") Then
	        	Line Input #1, sStackLine
	        End If

	        sStackLine = LCase(Trim(sStackLine))

	        Dim nCurrentMaskId As Integer
	        nCurrentMaskId = -1
	        Dim sCurrentMaskName As String
	        sCurrentMaskName = ""

	        If (sStackLine = "end_stack") Then
			   sStackLine = ""
			Else
			  If (sStackLine <> "") Then

			  	Dim sStackType As String
			  	Dim sTmpStackline As String
				sTmpStackline = sStackLine
			  	sStackType = GetStackType(sStackLine)

			  	If (sStackType = "interface") Then

			  		Dim sInterfaceName As String
			  		Dim sTmpInterface As String
	          		sTmpInterface = sStackLine
			  		sMaskValues = ScanStackInterface(sTmpInterface, sInterfaceName, nCurrentMaskId, sCurrentMaskName, sMaskValues)

			  		If (sMaskValues <> "") Then
						sStackLine = sTmpStackline
			  		End If

			  	Else

			  		Dim sLayerName As String, dHeight As Double, sMatName As String

			  		If (sStackType = "top") Then

			  			Dim nTopIndex As Integer
						nTopIndex = InStr(sStackLine, " ")
						If (nTopIndex <> 0) Then
							sTopPlane = Trim(Left(sStackLine, nTopIndex))
						End If

						sLayerName = "topplane"
						dHeight = 0.0
						sMatName = GetValueForCommand(sStackLine, "material")

			  		ElseIf (sStackType = "layer") Then

			  			Dim sTmpLayer As String
	          			sTmpLayer = sStackLine
	          			sMaskValues = ScanStackLayer(sTmpLayer, sLayerName, dHeight, sMatName, nCurrentMaskId, sCurrentMaskName, sMaskValues)

	          			If (sMaskValues <> "") Then
	          				sStackLine = sTmpStackline
	          			End If

			  		ElseIf (sStackType = "bottom") Then

			  			Dim nBottomIndex As Integer
						nBottomIndex = InStr(sStackLine, " ")
						If (nBottomIndex <> 0) Then
							sBottomPlane = Trim(Left(sStackLine, nBottomIndex))
						End If

						sLayerName = "bottomplane"
						dHeight = 0.0
						sMatName = GetValueForCommand(sStackLine, "material")

	          		End If

	          		' For ADS 2014 format, it happens that there are lines in the BEGIN STACK part which are no layer definitions.
	          		' These lines have no material defined and are skipped.

					If (sMatName <> "") Then

						' Do not add a new layer when its name has not changed!
						Dim bAddNewLayer As Boolean
						bAddNewLayer = True
						If (nLayer <> 0) Then
							If (sLayerNames(nLayer-1) = sLayerName) Then
								bAddNewLayer = False
							End If
						End If

	          			If (bAddNewLayer) Then

	          		Dim nMatInfoIndex As Integer
			  		nMatInfoIndex = GetMaterialIndexByName nMaterial, sMaterialName(), sMatName

			  		ReDim Preserve sLayerNames(nLayer)
				    ReDim Preserve dThickness(nLayer)
				    ReDim Preserve dEpsilon(nLayer)
				    ReDim Preserve dMue(nLayer)
				    ReDim Preserve dLossTang(nLayer)
				    ReDim Preserve dLossCond(nLayer)

				    sLayerNames(nLayer) = sLayerName
	          		dThickness(nLayer) = dHeight * dScale
	          		dEpsilon(nLayer) = dMaterialEpsilon(nMatInfoIndex)
	          		dMue(nLayer) = dMaterialMue(nMatInfoIndex)
	          		dLossTang(nLayer) = dMaterialLossTangent(nMatInfoIndex)
	          		dLossCond(nLayer) = dMaterialConductivity(nMatInfoIndex)

	          		nLayer = nLayer + 1

	          	End If

	          	End If

	          	End If

	          	If (sCurrentMaskName <> "") Then
					nCurrentMaskId = GetMaskIdByName nMask, nMaskId(), sMaskName(), sCurrentMaskName
	          	End If

	          	If (nCurrentMaskId <> -1 ) Then

		  			Dim nMaskInfoIndex As Integer, nMaterialInfoIndex As Integer, nOperationInfoIndex As Integer
			  		nMaskInfoIndex = GetMaskIndexById nMask, nMaskId(), nCurrentMaskId
			  		nMaterialInfoIndex = GetMaterialIndexByName nMaterial, sMaterialName(), sMaskMaterial(nMaskInfoIndex)
			  		nOperationInfoIndex = GetOperationIndexByName nOperation, sOperationName(), sMaskOperation(nMaskInfoIndex)

			  		If (sOperationType(nOperationInfoIndex) = "sheet" ) Then

			  			ReDim Preserve nPlaneId(nPlanes)
      					ReDim Preserve nLayerFrom(nPlanes)
        				ReDim Preserve sPlaneTypes(nPlanes)
        				ReDim Preserve dPlaneCond(nPlanes)
      					ReDim Preserve dPlaneImpedance(nPlanes)
        				ReDim Preserve dPlaneThickness(nPlanes)

				        nPlaneId(nPlanes) = nCurrentMaskId
				        sPlaneTypes(nPlanes) = "wall"
				        nLayerFrom(nPlanes) = 1

				        dPlaneCond(nPlanes) = dMaterialConductivity(nMaterialInfoIndex)
			  			dPlaneImpedance(nPlanes) = 0.0
			  			dPlaneThickness(nPlanes) = dOperationThickness(nOperationInfoIndex) * 1e6

				        nPlanes = nPlanes + 1

				    ElseIf (sOperationType(nOperationInfoIndex) = "boundaryoutline") Then

				    	bBoundaryOutlineDefined = True

				    	ReDim Preserve nPlaneId(nPlanes)
      					ReDim Preserve nLayerFrom(nPlanes)
        				ReDim Preserve sPlaneTypes(nPlanes)
        				ReDim Preserve dPlaneCond(nPlanes)
      					ReDim Preserve dPlaneImpedance(nPlanes)
        				ReDim Preserve dPlaneThickness(nPlanes)

			  			nPlaneId(nPlanes) = nCurrentMaskId
			  			nLayerFrom(nPlanes) = nLayer

						sPlaneTypes(nPlanes) = "boundary"

			  			'dPlaneCond(nPlanes) = dMaterialConductivity(nMaterialInfoIndex)
			  			'dPlaneImpedance(nPlanes) = 0.0
			  			'dPlaneThickness(nPlanes) = dHeight * dScale

			  			nPlanes = nPlanes + 1

			  		ElseIf (sOperationType(nOperationInfoIndex) = "intrude" Or sOperationType(nOperationInfoIndex) = "expand") Then

			  			ReDim Preserve nPlaneId(nPlanes)
      					ReDim Preserve nLayerFrom(nPlanes)
        				ReDim Preserve sPlaneTypes(nPlanes)
        				ReDim Preserve dPlaneCond(nPlanes)
      					ReDim Preserve dPlaneImpedance(nPlanes)
        				ReDim Preserve dPlaneThickness(nPlanes)

			  			nPlaneId(nPlanes) = nCurrentMaskId
			  			nLayerFrom(nPlanes) = nLayer

			  			If (sMaskType(nMaskInfoIndex) = "negative") Then
							sPlaneTypes(nPlanes) = "slot"
			  			Else
			  				sPlaneTypes(nPlanes) = "strip"
			  			End If


			  			dPlaneCond(nPlanes) = dMaterialConductivity(nMaterialInfoIndex)
			  			dPlaneImpedance(nPlanes) = 0.0
			  			If (dPlaneCond(nPlanes) <> 0.0 Or (IsCSTVersionOfCurrentBlock_GreaterEqualThan(2011, 6) And sMaterialName(nMaterialInfoIndex) = "perfect_conductor")) Then
			  				dPlaneThickness(nPlanes) = dOperationThickness(nOperationInfoIndex) * 1e6
			  			Else
							dPlaneThickness(nPlanes) = 0.0
			  			End If

			  			nPlanes = nPlanes + 1

			  		ElseIf (sOperationType(nOperationInfoIndex) = "drill" ) Then

			  			ReDim Preserve nViaId(nVias)
      					ReDim Preserve nViaStart(nVias)
        				ReDim Preserve nViaEnd(nVias)

			  			nViaId(nVias) = nCurrentMaskId
			  			nViaStart(nVias) = nLayer-1
				  		nViaEnd(nVias) = nLayer-1

				  		nVias = nVias + 1

				  	End If

			  	End If

	          End If

	        End If

        Loop Until sStackLine = ""

      End If

    End If

    Line Input #1, sLine
  Wend

  Finish:
  Close #1

End Sub

Sub UpdateMetalThickness(bUseSheets As Boolean, dMinimumMetalThickness As Double, nPlanes As Integer, dPlaneThickness() As Double)

  Dim i As Integer

  For i=0 To nPlanes-1
    If (bUseSheets) Then
      dPlaneThickness(i) = 0.0
    Else
	    If (dPlaneThickness(i) = 0.0) Then
	      dPlaneThickness(i) = dMinimumMetalThickness
	    End If
	End If
  Next i

End Sub

Function max(a As Double, b As Double)
	If a > b Then
	  max = a
	Else
	  max = b
	End If
End Function

Sub UpdateSubstrateThickness(nPlanes As Integer, nVias As Integer, dPlaneThickness() As Double, dSubstrateThickness() As Double, nLayerFrom() As Integer, sPlaneType() As String, nViaStart() As Integer, nViaEnd() As Integer, dPlaneCoords() As Double, dViaStart() As Double, dViaEnd() As Double)

  Dim nTotal As Integer
  nTotal = UBound(dSubstrateThickness)

  Dim dAddHeight() As Double
  ReDim dAddHeight(nTotal)

  Dim i As Integer

  For i=0 To nTotal
    dAddHeight(i) = 0.0
  Next i

  For i=0 To nPlanes-1
    If dPlaneThickness(i) > 0.0 Then
		dAddHeight(nLayerFrom(i)-1) = max(dPlaneThickness(i), dAddHeight(nLayerFrom(i)-1))
    Else
		dAddHeight(nLayerFrom(i)) = max(dPlaneThickness(i), -dPlaneThickness(i))
    End If
  Next i

  For i=0 To nTotal
    If (dSubstrateThickness(i) > 0.0) Then
      dSubstrateThickness(i) = dSubstrateThickness(i) + dAddHeight(i)
    End If
  Next i

  ReDim dPlaneCoords(nPlanes)

  For i=0 To nPlanes-1
      dPlaneCoords(i) = GetZCoordinateOfLayerTop(nLayerFrom(i), dSubstrateThickness)
  Next i

  ReDim dViaStart(nVias)
  ReDim dViaEnd(nVias)

  For i=0 To nVias-1
      dViaStart(i) = GetZCoordinateOfLayerTop(nViaStart(i), dSubstrateThickness)
	  dViaEnd(i)   = GetZCoordinateOfLayerBottom(nViaEnd(i), dSubstrateThickness)
  Next i
End Sub


Function GetSubstrateHeight(dThickness() As Double)

  Dim nTotal As Integer
  nTotal = UBound(dThickness)

  Dim dHeight As Double
  dHeight = 0.0

  Dim i As Integer

  For i=0 To nTotal
    dHeight = dHeight + dThickness(i)
  Next i

  GetSubstrateHeight = dHeight

End Function

Function GetZCoordinateOfLayerTop(nLayer As Integer, dThickness() As Double)
  Dim nTotal As Integer
  nTotal = UBound(dThickness)

  Dim dCoord As Double
  dCoord = 0.0

  Dim i As Integer

  For i=nTotal To nLayer STEP -1
    dCoord = dCoord + dThickness(i)
  Next i

  GetZCoordinateOfLayerTop = dCoord

End Function

Function GetZCoordinateOfLayerBottom(nLayer As Integer, dThickness() As Double)
  Dim nTotal As Integer
  nTotal = UBound(dThickness)

  Dim dCoord As Double
  dCoord = 0.0

  Dim i As Integer

  For i=nTotal To nLayer+1 STEP -1
    dCoord = dCoord + dThickness(i)
  Next i

  GetZCoordinateOfLayerBottom = dCoord

End Function

Sub ReadPlanes(sFileName As String, nPlanes As Integer, nPlaneId() As Integer, nLayerFrom() As Integer, sPlaneTypes() As String, dPlaneCond() As Double, dPlaneImpedance() As Double, dPlaneThickness() As Double, nVias As Integer, nViaId() As Integer, nViaStart() As Integer, nViaEnd() As Integer, dThickness() As Double)

  Dim sBlock As String

  Open sFileName For Input As #1
  On Error GoTo finish

  Do
    sBlock = Trim(ReadBlock(1, ",", ";"))

    If (sBlock <> "") Then
      Dim nPlane As Integer, sType As String, nLayerFromLocal As Integer, nLayerToLocal As Integer, sItem As String, nItem As Integer
      Dim sFirst As String, sSecond As String

      nPlane = GetInt(GetItem(sBlock))
      sType  = Trim(GetItem(sBlock))

	  sItem = GetItem(sBlock)
	  nItem = InStr(sItem, "-")

	  If (nItem = 0) Then
	    nLayerFromLocal = GetInt(sItem)
	    nLayerToLocal   = GetInt(sItem)
	  Else
        sFirst  = Left(sItem, nItem-1)
        sSecond = Mid(sItem, nItem+1)

	    nLayerFromLocal = GetInt(sFirst)
	    nLayerToLocal   = GetInt(sSecond)
	  End If

      Select Case sType
      Case "wall"
      	ReDim Preserve nPlaneId(nPlanes)
      	ReDim Preserve dPlaneCoords(nPlanes)
        ReDim Preserve sPlaneTypes(nPlanes)
        ReDim Preserve nLayerFrom(nPlanes)

        nPlaneId(nPlanes) = nPlane
        sPlaneTypes(nPlanes) = "wall"
        nLayerFrom(nPlanes) = 1

        nPlanes = nPlanes + 1

        ' ensure that loop does not finish after this
        sBlock = "precedence 0"

      Case "strip"
        ReDim Preserve nPlaneId(nPlanes)
        ReDim Preserve dPlaneCoords(nPlanes)
        ReDim Preserve sPlaneTypes(nPlanes)
        ReDim Preserve dPlaneCond(nPlanes)
        ReDim Preserve dPlaneImpedance(nPlanes)
        ReDim Preserve nLayerFrom(nPlanes)
        ReDim Preserve dPlaneThickness(nPlanes)

        Dim sCondType As String, sUnitType As String
        sCondType = Trim(GetItem(sBlock))

        Dim dUnitScale As Double

        nPlaneId(nPlanes) = nPlane
        sPlaneTypes(nPlanes) = "strip"
        nLayerFrom(nPlanes) = nLayerFromLocal

		If sCondType = "condthickness" Then
			dPlaneCond(nPlanes)      = GetDouble(GetItem(sBlock))
			dPlaneThickness(nPlanes) = GetDouble(GetItem(sBlock)) * 1e6
			dPlaneImpedance(nPlanes) = 0.0
		ElseIf sCondType = "conductivity" Then
			dPlaneCond(nPlanes)      = GetDouble(GetItem(sBlock))
			dPlaneImpedance(nPlanes) = 0.0
			GetItem(sBlock)
			If (Trim(GetItem(sBlock)) = "thickness") Then
				dPlaneThickness(nPlanes) = GetDouble(GetItem(sBlock)) * 1e6
				sUnitType = LCase(Trim(GetItem(sBlock)))

				dUnitScale = 1.0

				If (sUnitType = "meter") Then
				   dUnitScale = 1.0
				ElseIf (sUnitType = "mm") Then
				   dUnitScale = 0.001
				ElseIf (sUnitType = "um") Then
				   dUnitScale = 1.0e-6
				ElseIf (sUnitType = "mil") Then
				   dUnitScale = 2.54e-5
				Else
				   dUnitScale = 1.0
				End If

				dPlaneThickness(nPlanes) = dPlaneThickness(nPlanes) * dUnitScale
			End If
		ElseIf sCondType = "impedance" Then
			dPlaneCond(nPlanes)      = 0.0
			dPlaneImpedance(nPlanes) = GetDouble(GetItem(sBlock))

			GetItem(sBlock)
			If (Trim(GetItem(sBlock)) = "thickness") Then
				dPlaneThickness(nPlanes) = GetDouble(GetItem(sBlock)) * 1e6
				sUnitType = LCase(Trim(GetItem(sBlock)))

				dUnitScale = 1.0

				If (sUnitType = "meter") Then
				   dUnitScale = 1.0
				ElseIf (sUnitType = "mm") Then
				   dUnitScale = 0.001
				ElseIf (sUnitType = "um") Then
				   dUnitScale = 1.0e-6
				ElseIf (sUnitType = "mil") Then
				   dUnitScale = 2.54e-5
				Else
				   dUnitScale = 1.0
				End If

				dPlaneThickness(nPlanes) = dPlaneThickness(nPlanes) * dUnitScale
			End If

		Else
			dPlaneCond(nPlanes)      = 0.0
			dPlaneImpedance(nPlanes) = 0.0
			dPlaneThickness(nPlanes) = 0.0
		End If

        nPlanes = nPlanes + 1

      Case "slot"
        ReDim Preserve nPlaneId(nPlanes)
        ReDim Preserve sPlaneTypes(nPlanes)
        ReDim Preserve dPlaneCond(nPlanes)
        ReDim Preserve dPlaneThickness(nPlanes)
        ReDim Preserve nLayerFrom(nPlanes)
        ReDim Preserve dPlaneImpedance(nPlanes)

        nPlaneId(nPlanes) = nPlane
        sPlaneTypes(nPlanes) = "slot"
        nLayerFrom(nPlanes) = nLayerFromLocal

		If (Trim(GetItem(sBlock)) = "condthickness") Then
			dPlaneCond(nPlanes)      = GetDouble(GetItem(sBlock))
			dPlaneImpedance(nPlanes) = 0.0
			dPlaneThickness(nPlanes) = GetDouble(GetItem(sBlock)) * 1e6
		ElseIf sCondType = "conductivity" Then
			dPlaneCond(nPlanes)      = GetDouble(GetItem(sBlock))
			dPlaneImpedance(nPlanes) = 0.0
			GetItem(sBlock)
			If (Trim(GetItem(sBlock)) = "thickness") Then
				dPlaneThickness(nPlanes) = GetDouble(GetItem(sBlock)) * 1e6
				sUnitType = LCase(Trim(GetItem(sBlock)))

				dUnitScale = 1.0

				If (sUnitType = "meter") Then
				   dUnitScale = 1.0
				ElseIf (sUnitType = "mm") Then
				   dUnitScale = 0.001
				ElseIf (sUnitType = "um") Then
				   dUnitScale = 1.0e-6
				ElseIf (sUnitType = "mil") Then
				   dUnitScale = 2.54e-5
				Else
				   dUnitScale = 1.0
				End If

				dPlaneThickness(nPlanes) = dPlaneThickness(nPlanes) * dUnitScale
			End If
		Else
			dPlaneCond(nPlanes)      = 0.0
			dPlaneImpedance(nPlanes) = 0.0
			dPlaneThickness(nPlanes) = 0.0
		End If

        nPlanes = nPlanes + 1

      Case "via"

	      ReDim Preserve nViaId(nVias)
	      ReDim Preserve nViaStart(nVias)
	      ReDim Preserve nViaEnd(nVias)

	      nViaId(nVias) = nPlane
	      nViaStart(nVias) = nLayerFromLocal
	      nViaEnd(nVias)   = nLayerToLocal

	      nVias = nVias + 1

      End Select

    End If

  Loop Until sBlock = ""

Finish:
  Close #1


End Sub

Sub GetViaRange(nVia As Integer, dFrom As Double, dTo As Double, nVias As Integer, nViaId() As Integer, dViaStart() As Double, dViaEnd() As Double)

  Dim i As Integer
  For i=0 To nVias-1
    If (nViaId(i) = nVia) Then
      dFrom = dViaStart(i)
      dTo = dViaEnd(i)
    End If
  Next i

End Sub

Function GetPlaneCoord(nPlane As Integer, nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double)

  Dim dCoord As Double
  dCoord = -1.0

  Dim i As Integer
  For i=0 To nPlanes-1
    If (nPlaneId(i) = nPlane) Then
      dCoord = dPlaneCoords(i)
    End If
  Next i

  GetPlaneCoord = dCoord

End Function

Function GetPlaneThickness(nPlane As Integer, nPlanes As Integer, nPlaneId() As Integer, dPlaneThickness() As Double)

  Dim dThickness As Double
  dThickness = -1.0

  Dim i As Integer
  For i=0 To nPlanes-1
    If (nPlaneId(i) = nPlane) Then
      dThickness = dPlaneThickness(i)
    End If
  Next i

  GetPlaneThickness = dThickness

End Function

Function GetMinimumPlaneThickness(nPlanes As Integer, dPlaneThickness() As Double)

  Dim dThickness As Double
  dThickness = 1e30

  Dim i As Integer

  For i=0 To nPlanes-1
    If (Abs(dPlaneThickness(i)) < dThickness) Then
		dThickness = Abs(dPlaneThickness(i))
    End If
  Next i

  GetMinimumPlaneThickness = dThickness

End Function


Function GetPlaneCond(nPlane As Integer, nPlanes As Integer, nPlaneId() As Integer, dPlaneCond() As Double)

  Dim dCond As Double
  dCond = -1.0

  Dim i As Integer
  For i=0 To nPlanes-1
    If (nPlaneId(i) = nPlane) Then
      dCond = dPlaneCond(i)
    End If
  Next i

  GetPlaneCond = dCond

End Function

Function GetPlaneImpedance(nPlane As Integer, nPlanes As Integer, nPlaneId() As Integer, dPlaneImpedance() As Double)

  Dim dImpedance As Double
  dImpedance = -1.0

  Dim i As Integer
  For i=0 To nPlanes-1
    If (nPlaneId(i) = nPlane) Then
      dImpedance = dPlaneImpedance(i)
    End If
  Next i

  GetPlaneImpedance = dImpedance

End Function

Function GetPlaneType(nPlane As Integer, nPlanes As Integer, nPlaneId() As Integer, sPlaneType() As String) As String

  Dim sType As String
  sType = "strip"

  Dim i As Integer
  For i=0 To nPlanes-1
    If (nPlaneId(i) = nPlane) Then
      sType = sPlaneType(i)
    End If
  Next i

  GetPlaneType = sType

End Function

Function AreLinesColinear(dX1A As Double, dY1A As Double, dX2A As Double, dY2A As Double, dX1B As Double, dY1B As Double, dX2B As Double, dY2B As Double) As Boolean

	Dim dDelta As Double
	dDelta = 1e-5

	Dim dVxA As Double, dVyA As Double, dVAAbs As Double

	dVxA   = dX2A - dX1A
	dVyA   = dY2A - dY1A
	dVAAbs = Sqr(dVxA * dVxA +dVyA * dVyA)
	dVxA   = dVxA / dVAAbs
	dVyA   = dVyA / dVAAbs

	Dim dVxB As Double, dVyB As Double, dVBAbs As Double

	dVxB   = dX2B - dX1B
	dVyB   = dY2B - dY1B
	dVBAbs = Sqr(dVxB * dVxB +dVyB * dVyB)
	dVxB   = dVxB / dVBAbs
	dVyB   = dVyB / dVBAbs

	If (Abs(dVxA * dVxB + dVyA * dVyB) < 0.999) Then
		AreLinesColinear = False
		Exit Function
	End If

	Dim dTx As Double, dTy As Double, dS1 As Double, dS2 As Double

	dTx = dX1B - dX1A
	dTy = dY1B - dY1A

	dS1 = dTx * dVxA + dTy * dVyA

	dTx = dTx - dS1 * dVxA
	dTy = dTy - dS1 * dVyA

	If (Abs(dTx) > dDelta Or Abs(dTy) > dDelta) Then
		AreLinesColinear = False
		Exit Function
	End If

	dTx = dX2B - dX1A
	dTy = dY2B - dY1A

	dS2 = dTx * dVxA + dTy * dVyA

	If ((dS1 < 0.0 And dS2 < 0.0) Or dS1 > dVAAbs And dS2 > dVAAbs) Then
		AreLinesColinear = False
	Else
		AreLinesColinear = True
	End If

End Function

Sub ProcessPolygon(nPlane As Integer, sPlaneType As String, dCond As Double, dImpedance As Double, sPoints As String, dFrom As Double, dTo As Double, nCount As Long, dScale As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double, nPolygonCount As Long, bPolygonFailed As Boolean, sUseSimplification As String, iSimplifyMinPointsArc As Integer, iSimplifyMinPointsCircle As Integer, dSimplifyAngle As Double, dSimplifyAdjacentTol As Double, dSimplifyRadiusTol As Double, dSimplifyAngleTang As Double, dSimplifyEdgeLength As Double)

Dim bFirst As Boolean
bFirst = True

Dim sMaterialName As String
sMaterialName = ""

If (sPlaneType <> "wall" And sPlaneType <> "boundary") Then

	Component.New "Plane" + CVar(nPlane)

	If (dCond = 0.0 And dImpedance = 0.0) Then

		sMaterialName = "Plane" + CVar(nPlane)

		With Material
		     .Reset
		     .Name sMaterialName
		     .FrqType "hf"
		     .Type "Pec"
		     .SetMaterialUnit "GHz", "mm"
		     .Rho "0.0"
		     .Colour "0.952941", "0.972549", "0.219608"
		     .Wireframe "False"
		     .Transparency "0"
		     .Create
		End With

	ElseIf (dCond <> 0.0 And dImpedance = 0.0) Then

		sMaterialName = "Plane" + CVar(nPlane)

		With Material
		     .Reset
		     .Name sMaterialName
		     .FrqType "hf"
		     .Type "Lossy metal"
		     .Kappa CVar(dCond)
		     .SetMaterialUnit "GHz", "mm"
		     .Rho "0.0"
		     .Colour "0.952941", "0.972549", "0.219608"
		     .Wireframe "False"
		     .Transparency "0"
		     .Create
		End With

	ElseIf (dCond = 0.0 And dImpedance <> 0.0) Then

		If (dFrom = dTo) Then

			sMaterialName = "Plane" + CVar(nPlane)

			With Material
			     .Reset
			     .Name sMaterialName
			     .FrqType "hf"
			     .Type "Lossy metal"
			     .OhmicSheet CVar(dImpedance)
			     .SetMaterialUnit "GHz", "mm"
			     .Rho "0.0"
			     .Colour "0.34902", "0.968628", "0.67451"
			     .Wireframe "False"
			     .Transparency "0"
			     .Create
			End With

		Else

			sMaterialName = "Plane" + CVar(nPlane) + "_Normal"

			bShowOhmicSheetWarning = True

			With Material
			     .Reset
			     .Name "Plane" + CVar(nPlane) + "_Normal"
			     .FrqType "hf"
			     .Type "Normal"
			     .Epsilon "1"
                    .Mu("1")
			     .SetMaterialUnit "GHz", "mm"
			     .Kappa CVar(1.0 / dImpedance / Abs(dTo - dFrom) * 1e6)
			     .TanD "0.0"
			     .TanDFreq "0.0"
			     .TanDGiven "False"
			     .TanDModel "ConstTanD"
			     .KappaM "0.0"
			     .TanDM "0.0"
			     .TanDMFreq "0.0"
			     .TanDMGiven "False"
			     .TanDMModel "ConstTanD"
			     .DispModelEps "None"
                    .DispModelMu("None")
			     .DispersiveFittingSchemeEps "General 1st"
			     .DispersiveFittingSchemeMue "General 1st"
			     .UseGeneralDispersionEps "False"
			     .UseGeneralDispersionMue "False"
			     .Rho "0.0"
			     .ThermalConductivity "0"
			     .SetActiveMaterial "hf"
			     .Colour "0.34902", "0.968628", "0.67451"
			     .Wireframe "False"
			     .Transparency "0"
			     .Create
			End With

			With Material
			     .Reset
			     .Name "Plane" + CVar(nPlane) + "_OhmicSheet"
			     .FrqType "hf"
			     .Type "Lossy metal"
			     .OhmicSheet CVar(dImpedance)
			     .SetMaterialUnit "GHz", "mm"
			     .Rho "0.0"
			     .Colour "0.34902", "0.968628", "0.67451"
			     .Wireframe "False"
			     .Transparency "0"
			     .Create
			End With
		End If

	Else
		ReportWarningToWindow "Invalid or missing material information for Plane" + CVar(nPlane)
	End If
End If

Dim dXP() As Double, dYP() As Double
Dim nPoints As Integer
nPoints = 0

Do
	Dim sCoord As String
	sCoord = Trim(GetItem(sPoints))

	Dim nIndex As Integer
	nIndex = InStr(sCoord, ",")

	Dim dX As Double, dY As Double
	dX = dScale * GetDouble(Left(sCoord, nIndex-1))
	dY = dScale * GetDouble(Mid(sCoord, nIndex+1))

	ReDim Preserve dXP(nPoints), dYP(nPoints)
	dXP(nPoints) = dX
	dYP(nPoints) = dY

	nPoints = nPoints + 1

	dXL = IIf(dXL > dX, dX, dXL)
	dXH = IIf(dXH < dX, dX, dXH)
	dYL = IIf(dYL > dY, dY, dYL)
	dYH = IIf(dYH < dY, dY, dYH)

Loop Until sPoints = ""

If (dXP(0) <> dXP(nPoints-1) Or dYP(0) <> dYP(nPoints-1)) Then
	ReDim Preserve dXP(nPoints), dYP(nPoints)
	dXP(nPoints) = dXP(0)
	dYP(nPoints) = dYP(0)
	nPoints = nPoints + 1
End If

nPoints = nPoints-1

Dim dXC(3) As Double, dYC(3) As Double
Dim bCorrection As Boolean
bCorrection = False

Dim iSeg As Integer
For iSeg = 0 To nPoints-1
  Dim dX1 As Double, dY1 As Double
  Dim dX2 As Double, dY2 As Double

  dX1 = dXP(iSeg)
  dY1 = dYP(iSeg)
  dX2 = dXP((iSeg+1) Mod nPoints)
  dY2 = dYP((iSeg+1) Mod nPoints)

  Dim iSeg2 As Integer
  For iSeg2 = 0 To iSeg-1
    Dim dX12 As Double, dY12 As Double
  	Dim dX22 As Double, dY22 As Double

	  dX12 = dXP(iSeg2)
	  dY12 = dYP(iSeg2)
	  dX22 = dXP((iSeg2+1) Mod nPoints)
	  dY22 = dYP((iSeg2+1) Mod nPoints)

      If (AreLinesColinear(dX1, dY1, dX2, dY2, dX12, dY12, dX22, dY22)) Then

        Dim dX1N As Double, dY1N As Double, dX2N As Double, dY2N As Double
   		Dim dX1B As Double, dY1B As Double, dX2A As Double, dY2A As Double

        dX1B = dXP((iSeg-1) Mod nPoints)
	    dY1B = dYP((iSeg-1) Mod nPoints)
	    dX2A = dXP((iSeg+2) Mod nPoints)
	    dY2A = dYP((iSeg+2) Mod nPoints)

	    dX1N = 0.99 * (dX1 - dX1B) + dX1B
	    dY1N = 0.99 * (dY1 - dY1B) + dY1B
	    dX2N = 0.01 * (dX2A - dX2) + dX2
	    dY2N = 0.01 * (dY2A - dY2) + dY2

        dXP(iSeg) = dX1N
        dYP(iSeg) = dY1N
        dXP((iSeg+1) Mod nPoints) = dX2N
        dYP((iSeg+1) Mod nPoints) = dY2N

        bCorrection = True

        dXC(0) = dX1
        dXC(1) = dX1N
        dXC(2) = dX2N
        dXC(3) = dX2

        dYC(0) = dY1
        dYC(1) = dY1N
        dYC(2) = dY2N
        dYC(3) = dY2

      End If

  Next iSeg2
Next iSeg

nPoints = nPoints+1
dXP(nPoints-1) = dXP(0)
dYP(nPoints-1) = dYP(0)

If (sPlaneType <> "wall" And sPlaneType <> "boundary") Then

	On Error GoTo Failed

	With Extrude
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Component "Plane" + CVar(nPlane)
	     .Material sMaterialName
	     .Mode "Pointlist"
	     .Height CVar(dTo - dFrom)
	     .Twist "0.0"
	     .Taper "0.0"
	     .Origin "0.0", "0.0", CVar(dFrom)
	     .Uvector "1.0", "0.0", "0.0"
	     .Vvector "0.0", "1.0", "0.0"
	     .Point Replace(CVar(dXP(0)), ",", "."), Replace(CVar(dYP(0)), ",", ".")

	     Dim iPoint As Integer
	     For iPoint = 1 To nPoints-1
	       .LineTo Replace(CVar(dXP(iPoint)), ",", "."), Replace(CVar(dYP(iPoint)), ",", ".")
	     Next iPoint

	     .SetSimplifyActive sUseSimplification
	     .SetSimplifyMinPointsArc iSimplifyMinPointsArc
	     .SetSimplifyMinPointsCircle iSimplifyMinPointsCircle
	     .SetSimplifyAngle dSimplifyAngle
	     .SetSimplifyAdjacentTol dSimplifyAdjacentTol
	     .SetSimplifyRadiusTol dSimplifyRadiusTol
	     .SetSimplifyAngleTang dSimplifyAngleTang
	     .SetSimplifyEdgeLength dSimplifyEdgeLength

		If nPoints > 3 Then
	    	.Create
	    End If
	End With

	GoTo ExtrusionComplete

	Failed:

	bPolygonFailed = True

	Curve.NewCurve "failed polygons"

	With Polygon3D
	     .Reset
	     .Name "polygon" + CVar(nPolygonCount)
	     .Curve "failed polygons"
	     .Point Replace(CVar(dXP(0)), ",", "."), Replace(CVar(dYP(0)), ",", "."), dFrom
	     For iPoint = 1 To nPoints-1
	       .Point Replace(CVar(dXP(iPoint)), ",", "."), Replace(CVar(dYP(iPoint)), ",", "."), dFrom
	     Next iPoint
	     .Create
	End With

	ExtrusionComplete:

ElseIf (sPlaneType = "boundary") Then

	' for boundary outlines we consider the geometry to be the dimesions for the bounding box but we do not create any geometry.

Else
	Curve.NewCurve "wall polygons"

	With Polygon3D
	     .Reset
	     .Name "polygon" + CVar(nPolygonCount)
	     .Curve "wall polygons"
	     .Point Replace(CVar(dXP(0)), ",", "."), Replace(CVar(dYP(0)), ",", "."), dFrom
	     For iPoint = 1 To nPoints-1
	       .Point Replace(CVar(dXP(iPoint)), ",", "."), Replace(CVar(dYP(iPoint)), ",", "."), dFrom
	     Next iPoint
	     .Create
	End With

End If

nCount = nCount+1

If (bCorrection) Then

	With Extrude
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer sMaterialName
	     .Mode "Pointlist"
	     .Height CVar(dTo - dFrom)
	     .Twist "0.0"
	     .Taper "0.0"
	     .Origin "0.0", "0.0", CVar(dFrom)
	     .Uvector "1.0", "0.0", "0.0"
	     .Vvector "0.0", "1.0", "0.0"
	     .Point Replace(CVar(dXC(0)), ",", "."), Replace(CVar(dYC(0)), ",", ".")
	     .LineTo Replace(CVar(dXC(1)), ",", "."), Replace(CVar(dYC(1)), ",", ".")
	     .LineTo Replace(CVar(dXC(2)), ",", "."), Replace(CVar(dYC(2)), ",", ".")
	     .LineTo Replace(CVar(dXC(3)), ",", "."), Replace(CVar(dYC(3)), ",", ".")
	     .SetSimplifyActive sUseSimplification
	     .SetSimplifyMinPointsArc iSimplifyMinPointsArc
	     .SetSimplifyMinPointsCircle iSimplifyMinPointsCircle
	     .SetSimplifyAngle dSimplifyAngle
	     .SetSimplifyAdjacentTol dSimplifyAdjacentTol
	     .SetSimplifyRadiusTol dSimplifyRadiusTol
	     .SetSimplifyAngleTang dSimplifyAngleTang
	     .SetSimplifyEdgeLength dSimplifyEdgeLength
	     .Create
	End With

	nCount = nCount+1

End If

End Sub

Sub ProcessViaPolygon(nPlane As Integer, dCond As Double, sPoints As String, dFrom As Double, dTo As Double, nCount As Long, dScale As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double, nPolygonCount As Long, bPolygonFailed As Boolean, sUseSimplification As String, iSimplifyMinPointsArc As Integer, iSimplifyMinPointsCircle As Integer, dSimplifyAngle As Double, dSimplifyAdjacentTol As Double, dSimplifyRadiusTol As Double, dSimplifyAngleTang As Double, dSimplifyEdgeLength As Double)

Dim bFirst As Boolean
bFirst = True

Dim nSemicolon As Integer
nSemicolon = InStr(sPoints, ";")
sPoints = Left(sPoints, nSemicolon-1)

Component.New "Via" + CVar(nPlane) 

If (dCond = 0.0) Then

	With Material
	     .Reset
	     .Name "Via" + CVar(nPlane)
	     .FrqType "hf"
	     .Type "Pec"
	     .Rho "0.0"
	     .Colour "0.952941", "0.972549", "0.219608"
	     .Wireframe "False"
	     .Transparency "0"
	     .Create
	End With

Else

	With Material
	     .Reset
	     .Name "Via" + CVar(nPlane)
	     .FrqType "hf"
	     .Type "Lossy metal"
	     .Kappa CVar(dCond)
	     .Rho "0.0"
	     .Colour "0.952941", "0.972549", "0.219608"
	     .Wireframe "False"
	     .Transparency "0"
	     .Create
	End With

End If

Dim dXP() As Double, dYP() As Double
Dim nPoints As Integer
nPoints = 0

Do
	Dim sCoord As String
	sCoord = Trim(GetItem(sPoints))

	Dim nIndex As Integer
	nIndex = InStr(sCoord, ",")

	Dim dX As Double, dY As Double
	dX = dScale * GetDouble(Left(sCoord, nIndex-1))
	dY = dScale * GetDouble(Mid(sCoord, nIndex+1))

	ReDim Preserve dXP(nPoints), dYP(nPoints)
	dXP(nPoints) = dX
	dYP(nPoints) = dY

	nPoints = nPoints + 1

	dXL = IIf(dXL > dX, dX, dXL)
	dXH = IIf(dXH < dX, dX, dXH)
	dYL = IIf(dYL > dY, dY, dYL)
	dYH = IIf(dYH < dY, dY, dYH)

Loop Until sPoints = ""

If (dXP(0) <> dXP(nPoints-1) Or dYP(0) <> dYP(nPoints-1)) Then
	ReDim Preserve dXP(nPoints), dYP(nPoints)
	dXP(nPoints) = dXP(0)
	dYP(nPoints) = dYP(0)
	nPoints = nPoints + 1
End If

nPoints = nPoints-1

Dim dXC(3) As Double, dYC(3) As Double
Dim bCorrection As Boolean
bCorrection = False

Dim iSeg As Integer
For iSeg = 0 To nPoints-1
  Dim dX1 As Double, dY1 As Double
  Dim dX2 As Double, dY2 As Double

  dX1 = dXP(iSeg)
  dY1 = dYP(iSeg)
  dX2 = dXP((iSeg+1) Mod nPoints)
  dY2 = dYP((iSeg+1) Mod nPoints)

  Dim iSeg2 As Integer
  For iSeg2 = 0 To iSeg-1
    Dim dX12 As Double, dY12 As Double
  	Dim dX22 As Double, dY22 As Double

	  dX12 = dXP(iSeg2)
	  dY12 = dYP(iSeg2)
	  dX22 = dXP((iSeg2+1) Mod nPoints)
	  dY22 = dYP((iSeg2+1) Mod nPoints)

      If (AreLinesColinear(dX1, dY1, dX2, dY2, dX12, dY12, dX22, dY22)) Then

        Dim dX1N As Double, dY1N As Double, dX2N As Double, dY2N As Double
   		Dim dX1B As Double, dY1B As Double, dX2A As Double, dY2A As Double

        dX1B = dXP((iSeg-1) Mod nPoints)
	    dY1B = dYP((iSeg-1) Mod nPoints)
	    dX2A = dXP((iSeg+2) Mod nPoints)
	    dY2A = dYP((iSeg+2) Mod nPoints)

	    dX1N = 0.99 * (dX1 - dX1B) + dX1B
	    dY1N = 0.99 * (dY1 - dY1B) + dY1B
	    dX2N = 0.01 * (dX2A - dX2) + dX2
	    dY2N = 0.01 * (dY2A - dY2) + dY2

        dXP(iSeg) = dX1N
        dYP(iSeg) = dY1N
        dXP((iSeg+1) Mod nPoints) = dX2N
        dYP((iSeg+1) Mod nPoints) = dY2N

        bCorrection = True

        dXC(0) = dX1
        dXC(1) = dX1N
        dXC(2) = dX2N
        dXC(3) = dX2

        dYC(0) = dY1
        dYC(1) = dY1N
        dYC(2) = dY2N
        dYC(3) = dY2

      End If

  Next iSeg2
Next iSeg

nPoints = nPoints+1
dXP(nPoints-1) = dXP(0)
dYP(nPoints-1) = dYP(0)

On Error GoTo Failed

With Extrude
     .Reset
     .Name "solid" + CVar(nCount)
     .Layer "Via" + CVar(nPlane)
     .Mode "Pointlist"
     .Height CVar(dTo - dFrom)
     .Twist "0.0"
     .Taper "0.0"
     .Origin "0.0", "0.0", CVar(dFrom)
     .Uvector "1.0", "0.0", "0.0"
     .Vvector "0.0", "1.0", "0.0"
     .Point Replace(CVar(dXP(0)), ",", "."), Replace(CVar(dYP(0)), ",", ".")

     Dim iPoint As Integer
     For iPoint = 1 To nPoints-1
       .LineTo Replace(CVar(dXP(iPoint)), ",", "."), Replace(CVar(dYP(iPoint)), ",", ".")
     Next iPoint
     
     .SetSimplifyActive sUseSimplification
     .SetSimplifyMinPointsArc iSimplifyMinPointsArc
     .SetSimplifyMinPointsCircle iSimplifyMinPointsCircle
     .SetSimplifyAngle dSimplifyAngle
     .SetSimplifyAdjacentTol dSimplifyAdjacentTol
     .SetSimplifyRadiusTol dSimplifyRadiusTol
     .SetSimplifyAngleTang dSimplifyAngleTang
     .SetSimplifyEdgeLength dSimplifyEdgeLength

	If nPoints > 3 Then
	    .Create
	End If
End With

GoTo ExtrusionComplete

Failed:

bPolygonFailed = True

Curve.NewCurve "failed polygons"

With Polygon3D
     .Reset
     .Name "polygon" + CVar(nPolygonCount)
     .Curve "failed polygons"
     .Point Replace(CVar(dXP(0)), ",", "."), Replace(CVar(dYP(0)), ",", "."), dFrom
     For iPoint = 1 To nPoints-1
       .Point Replace(CVar(dXP(iPoint)), ",", "."), Replace(CVar(dYP(iPoint)), ",", "."), dFrom
     Next iPoint
     .Create
End With

ExtrusionComplete:

nCount = nCount+1

If (bCorrection) Then

	With Extrude
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer "Via" + CVar(nPlane)
	     .Mode "Pointlist"
	     .Height CVar(dTo - dFrom)
	     .Twist "0.0"
	     .Taper "0.0"
	     .Origin "0.0", "0.0", CVar(dFrom)
	     .Uvector "1.0", "0.0", "0.0"
	     .Vvector "0.0", "1.0", "0.0"
	     .Point Replace(CVar(dXC(0)), ",", "."), Replace(CVar(dYC(0)), ",", ".")
	     .LineTo Replace(CVar(dXC(1)), ",", "."), Replace(CVar(dYC(1)), ",", ".")
	     .LineTo Replace(CVar(dXC(2)), ",", "."), Replace(CVar(dYC(2)), ",", ".")
	     .LineTo Replace(CVar(dXC(3)), ",", "."), Replace(CVar(dYC(3)), ",", ".")
	     .SetSimplifyActive sUseSimplification
	     .SetSimplifyMinPointsArc iSimplifyMinPointsArc
	     .SetSimplifyMinPointsCircle iSimplifyMinPointsCircle
	     .SetSimplifyAngle dSimplifyAngle
	     .SetSimplifyAdjacentTol dSimplifyAdjacentTol
	     .SetSimplifyRadiusTol dSimplifyRadiusTol
	     .SetSimplifyAngleTang dSimplifyAngleTang
	     .SetSimplifyEdgeLength dSimplifyEdgeLength
	     .Create
	End With

	nCount = nCount+1

End If

End Sub

Sub ProcessViaLine(nPlane As Integer, sPoints As String, dFrom As Double, dTo As Double, nCount As Long, dScale As Double, dThick As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double)

Dim nSemicolon As Integer
nSemicolon = InStr(sPoints, ";")
sPoints = Left(sPoints, nSemicolon-1)

Component.New "Via" + CVar(nPlane) 

With Material
     .Reset
     .Name "Via" + CVar(nPlane)
     .FrqType "hf"
     .Type "Pec"
     .Rho "0.0"
     .Colour "0.952941", "0.972549", "0.219608"
     .Wireframe "False"
     .Transparency "0"
     .Create
End With

Dim sCoord As String
Dim nIndex As Integer

sCoord = Trim(GetItem(sPoints))
nIndex = InStr(sCoord, ",")

Dim c1x As Double, c1y As Double
c1x = dScale * GetDouble(Left(sCoord, nIndex-1))
c1y = dScale * GetDouble(Mid(sCoord, nIndex+1))

sCoord = Trim(GetItem(sPoints))
nIndex = InStr(sCoord, ",")

Dim c2x As Double, c2y As Double
c2x = dScale * GetDouble(Left(sCoord, nIndex-1))
c2y = dScale * GetDouble(Mid(sCoord, nIndex+1))

Dim dDeltaX As Double, dDeltaY As Double

dDeltaX = c1x - c2x
dDeltaY = c1y - c2y

Dim dX As Double, dY As Double

If (Abs(dDeltaX) < Abs(dDeltaY)) Then

	With Brick
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer "Via"  + CVar(nPlane)
	     .Xrange c1x - dThick, c2x + dThick
	     .Yrange c1y, c2y
	     .Zrange dTo, dFrom
	     .Create
	End With

	Solid.SetAutomeshParameters "Via"  + CVar(nPlane) + ":solid" + CVar(nCount), "-1", "True"

	dX = c1x - dThick
	dY = c1y

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

	dX = c1x + dThick
	dY = c2y

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

Else

	With Brick
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer "Via"  + CVar(nPlane)
	     .Xrange c1x, c2x
	     .Yrange c1y - dThick, c2y + dThick
	     .Zrange dFrom, dTo
	     .Create
	End With

	Solid.SetAutomeshParameters "Via"  + CVar(nPlane) + ":solid" + CVar(nCount), "-1", "True"

	dX = c1x
	dY = c1y - dThick

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

	dX = c2x
	dY = c1y + dThick

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

End If

nCount = nCount+1

End Sub

Sub ProcessLine(nPlane As Integer, dCond As Double, sPoints As String, dFrom As Double, dTo As Double, nCount As Long, dScale As Double, dThick As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double)

Dim nSemicolon As Integer
nSemicolon = InStr(sPoints, ";")
sPoints = Left(sPoints, nSemicolon-1)

Component.New "Plane" + CVar(nPlane) 

If (dCond = 0.0) Then

	With Material
	     .Reset
	     .Name "Plane" + CVar(nPlane)
	     .FrqType "hf"
	     .Type "Pec"
	     .Rho "0.0"
	     .Colour "0.952941", "0.972549", "0.219608"
	     .Wireframe "False"
	     .Transparency "0"
	     .Create
	End With

Else

	With Material
	     .Reset
	     .Name "Plane" + CVar(nPlane)
	     .FrqType "hf"
	     .Type "Lossy metal"
	     .Kappa CVar(dCond)
	     .Rho "0.0"
	     .Colour "0.952941", "0.972549", "0.219608"
	     .Wireframe "False"
	     .Transparency "0"
	     .Create
	End With

End If

Dim sCoord As String
Dim nIndex As Integer

sCoord = Trim(GetItem(sPoints))
nIndex = InStr(sCoord, ",")

Dim c1x As Double, c1y As Double
c1x = dScale * GetDouble(Left(sCoord, nIndex-1))
c1y = dScale * GetDouble(Mid(sCoord, nIndex+1))

sCoord = Trim(GetItem(sPoints))
nIndex = InStr(sCoord, ",")

Dim c2x As Double, c2y As Double
c2x = dScale * GetDouble(Left(sCoord, nIndex-1))
c2y = dScale * GetDouble(Mid(sCoord, nIndex+1))

Dim dDeltaX As Double, dDeltaY As Double

dDeltaX = c1x - c2x
dDeltaY = c1y - c2y

Dim dX As Double, dY As Double

If (Abs(dDeltaX) < Abs(dDeltaY)) Then

	With Brick
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer "Plane"  + CVar(nPlane)
	     .Xrange c1x - dThick, c2x + dThick
	     .Yrange c1y, c2y
	     .Zrange dTo, dFrom
	     .Create
	End With

	Solid.SetAutomeshParameters "Plane"  + CVar(nPlane) + ":solid" + CVar(nCount), "-1", "True"

	dX = c1x - dThick
	dY = c1y

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

	dX = c1x + dThick
	dY = c2y

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

Else

	With Brick
	     .Reset
	     .Name "solid" + CVar(nCount)
	     .Layer "Plane"  + CVar(nPlane)
	     .Xrange c1x, c2x
	     .Yrange c1y - dThick, c2y + dThick
	     .Zrange dFrom, dTo
	     .Create
	End With

	Solid.SetAutomeshParameters "Plane"  + CVar(nPlane) + ":solid" + CVar(nCount), "-1", "True"

	dX = c1x
	dY = c1y - dThick

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

	dX = c2x
	dY = c1y + dThick

    dXL = IIf(dXL > dX, dX, dXL)
    dXH = IIf(dXH < dX, dX, dXH)
    dYL = IIf(dYL > dY, dY, dYL)
    dYH = IIf(dYH < dY, dY, dYH)

End If

nCount = nCount+1

End Sub

Function IsViaLayer(nPlane As Integer, nVias As Integer, nViaId() As Integer)

  Dim bFound As Boolean, i As Integer
  bFound = False

  For i=0 To nVias-1
    If (nViaId(i) = nPlane) Then
      bFound = True
    End If
  Next i

  IsViaLayer = bFound

End Function



Sub ReadGeometry(sFileName As String, sPlaneType() As String, dPlaneThickness() As Double, nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, dPlaneCond() As Double, dPlaneImpedance() As Double, nVias As Integer, nViaId() As Integer, dViaStart() As Double, dViaEnd() As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double, nEntitiesRead As Long, bPolygonFailed As Boolean, sUseSimplification As String, iSimplifyMinPointsArc As Integer, iSimplifyMinPointsCircle As Integer, dSimplifyAngle As Double, dSimplifyAdjacentTol As Double, dSimplifyRadiusTol As Double, dSimplifyAngleTang As Double, dSimplifyEdgeLength As Double, dMinimumMetalThickness As Double)

  Dim OpenError As Boolean
  OpenError = OpenTextFile(sFileName)
  If OpenError Then GoTo failed
  On Error GoTo failed
  On Error Resume Next

  Dim sLine As String
  Dim nPolyCount As Long, nViaCount As Long, nPolygonCount As Long
  nPolyCount = 0
  nViaCount = 0
  nPolygonCount = 0

  Dim dScale As Double
  dScale = 1.0

  Dim dMetalT As Double
  dMetalT = 0.0

  Dim numLinesToPreserve As Integer
  numLinesToPreserve = 1000
  Dim TextLines() As String
  Dim n As Integer, m As Integer

  ReDim Preserve TextLines(numLinesToPreserve)

  n = 0
  TextLines(n) = ReadLine
  While TextLines(n) <> ""
  	If( n > numLinesToPreserve ) Then
  		numLinesToPreserve = numLinesToPreserve + 1000
		ReDim Preserve TextLines(numLinesToPreserve)
  	End If
	n = n+1
	TextLines(n) = ReadLine
  Wend

  CloseTextFile

  m = n-1

  For n = 0 To m
  	sLine = ""
	sLine = TextLines(n)
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

      Dim sBlock As String
      sBlock = sLine

      Dim sCommand As String, sWhat As String
      sCommand = Trim(GetItem(sBlock))

	  Select Case sCommand
	  Case "units"
	    sWhat = Trim(GetItem(sBlock))
		Dim nColonIndex As Integer

		nColonIndex = InStr(sWhat, ",")
		sWhat = Trim(Left(sWhat, nColonIndex-1))

   		dScale = GetScaleFromName(sWhat)

	  Case "add"

	    While (InStr(sBlock, ";") = 0)
	      Dim sLine2 As String
	      Line Input #1, sLine2
	      sLine2 = Trim(LCase(sLine2))
	      sBlock = sBlock + " " + sLine2
	    Wend

	    sWhat = Trim(GetItem(sBlock))

	    Dim nPlane As Integer
	    nPlane = GetInt(Mid(sWhat, 2))
        Dim dFrom As Double, dTo As Double, dCond As Double, dImpedance As Double
        Dim sType As String

	    If (Left(sWhat, 1) = "p") Then

	      If (IsViaLayer(nPlane, nVias, nViaId())) Then

			  sBlock = Trim(sBlock)

			  If (Mid(sBlock, 1, 1) = ":") Then
			      GetItem(sBlock)
			  End If

	  	      GetViaRange nPlane, dFrom, dTo, nVias, nViaId(), dViaStart(), dViaEnd()

			  ProcessViaPolygon nPlane, dCond, sBlock, dFrom, dTo, nViaCount, dScale, dXL, dXH, dYL, dYH, nPolygonCount, bPolygonFailed, sUseSimplification, iSimplifyMinPointsArc, iSimplifyMinPointsCircle, dSimplifyAngle, dSimplifyAdjacentTol, dSimplifyRadiusTol, dSimplifyAngleTang, dSimplifyEdgeLength
			  nEntitiesRead = nEntitiesRead + 1

	      Else

			  sBlock = Trim(sBlock)

			  If (Mid(sBlock, 1, 1) = ":") Then
			      GetItem(sBlock)
			  End If

			  sType = GetPlaneType(nPlane, nPlanes, nPlaneId(), sPlaneType())

			  dFrom = GetPlaneCoord(nPlane, nPlanes, nPlaneId(), dPlaneCoords())
			  dTo   = dFrom + GetPlaneThickness(nPlane, nPlanes, nPlaneId, dPlaneThickness)

			  dCond     = GetPlaneCond(nPlane, nPlanes, nPlaneId(), dPlaneCond())
			  dImpedance = GetPlaneImpedance(nPlane, nPlanes, nPlaneId(), dPlaneImpedance())

	          ProcessPolygon nPlane, sType, dCond, dImpedance, sBlock, dFrom, dTo, nPolyCount, dScale, dXL, dXH, dYL, dYH, nPolygonCount, bPolygonFailed, sUseSimplification, iSimplifyMinPointsArc, iSimplifyMinPointsCircle, dSimplifyAngle, dSimplifyAdjacentTol, dSimplifyRadiusTol, dSimplifyAngleTang, dSimplifyEdgeLength
	   		  nEntitiesRead = nEntitiesRead + 1

	      End If

	    ElseIf (Left(sWhat, 1) = "l") Then

	      If (IsViaLayer(nPlane, nVias, nViaId())) Then

		      sBlock = Trim(sBlock)

		      If (Mid(sBlock, 1, 1) = ":") Then
				  GetItem(sBlock)
			  End If

	  	      GetViaRange nPlane, dFrom, dTo, nVias, nViaId(), dViaStart(), dViaEnd()

	  	      dMetalT = GetPlaneThickness(nPlane, nPlanes, nPlaneId, dPlaneThickness)

	  	      Dim dViaThickness As Double
	  	      dViaThickness = dMinimumMetalThickness

	  	      If (dViaThickness = 0.0) Then
				dViaThickness = GetMinimumPlaneThickness(nPlanes, dPlaneThickness)
	  	      End If

			  ProcessViaLine nPlane, sBlock, dFrom, dTo, nViaCount, dScale, 0.5 * dViaThickness, dXL, dXH, dYL, dYH
	   		  nEntitiesRead = nEntitiesRead + 1

	   	  Else

			  sBlock = Trim(sBlock)

			  If (Mid(sBlock, 1, 1) = ":") Then
			      GetItem(sBlock)
			  End If

			  dFrom = GetPlaneCoord(nPlane, nPlanes, nPlaneId(), dPlaneCoords())
			  dTo   = dFrom + GetPlaneThickness(nPlane, nPlanes, nPlaneId, dPlaneThickness)

			  dCond      = GetPlaneCond(nPlane, nPlanes, nPlaneId(), dPlaneCond())
			  dImpedance = GetPlaneImpedance(nPlane, nPlanes, nPlaneId(), dPlaneImpedance())

	  	      dMetalT = GetPlaneThickness(nPlane, nPlanes, nPlaneId, dPlaneThickness)

	          'ProcessLine nPlane, dCond, sBlock, dFrom, dTo, nPolyCount, dScale, dMetalT, dXL, dXH, dYL, dYH
	   		  'nEntitiesRead = nEntitiesRead + 1

		  End If

	    End If

      End Select

    End If

  Next n

  GoTo Finish

Failed:
	ReportWarningToWindow "Import process failed."
Finish:

End Sub

Sub StoreStructureBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  StoreGlobalDataValue "Macros\ADS Import\Structure Bounds\XLow", CVar(dXL)
  StoreGlobalDataValue "Macros\ADS Import\Structure Bounds\XHigh", CVar(dXH)
  StoreGlobalDataValue "Macros\ADS Import\Structure Bounds\YLow", CVar(dYL)
  StoreGlobalDataValue "Macros\ADS Import\Structure Bounds\YHigh", CVar(dYH)
End Sub

Sub StoreBBoxBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XLow", CVar(dXL)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\XHigh", CVar(dXH)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YLow", CVar(dYL)
  StoreGlobalDataValue "Macros\ADS Import\BBox Bounds\YHigh", CVar(dYH)
End Sub

Sub RestoreStructureBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  dXL = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XLow"))
  dXH = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\XHigh"))
  dYL = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YLow"))
  dYH = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\Structure Bounds\YHigh"))
End Sub

Sub RestoreBBoxBounds(dXL As Double, dXH As Double, dYL As Double, dYH As Double)
  dXL = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XLow"))
  dXH = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\XHigh"))
  dYL = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YLow"))
  dYH = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\BBox Bounds\YHigh"))
End Sub

Sub BuildSubstrateLayers(dThickness() As Double, dEpsilon() As Double, dMue() As Double, dLossTang() As Double, dLossCond() As Double)

  Dim nTotal As Integer
  nTotal = UBound(dThickness)

  Dim i As Integer

  For i=0 To nTotal
  
  Component.New "Substrate" + CVar(i)
  
	With Material
	     .Reset
	     .Name "Substrate" + CVar(i)
	     .FrqType "hf"
	     .Type "Normal"
	     .Epsilon CVar(dEpsilon(i))
            .Mue(CVar(dMue(i)))
	     .Kappa CVar(dLossCond(i))
	     .TanD CVar(dLossTang(i))
	     .TanDFreq "0.0"

	     If dLossTang(i) <> 0.0 Then
	     	.TanDGiven "True"
	     Else
	     	.TanDGiven "False"
	     End If

	     .TanDModel "ConstTanD"
	     .KappaM "0.0"
	     .TanDM "0.0"
	     .TanDMFreq "0.0"
	     .TanDMGiven "False"
	     .DispModelEps "None"
            .DispModelMue("None")
	     .Rho "0.0"
	     Select Case i Mod 3
	     Case 0
	     .Colour "0.654902", "0.741176", "0.996078"
	     Case 1
         .Colour "0.980392", "0.486275", "0.968628"
         Case 2
         .Colour "0.988235", "0.807843", "0.627451"
         End Select
	     .Wireframe "False"
	     .Transparentoutline "True"
	     .Transparency "95"
	     .Create
	End With

  Next i
End Sub

Function GetSolidOnLayer(sLayer As String) As String

    Dim nShapes As Integer, i As Integer
    Dim sName As String

    nShapes = Solid.GetNumberOfShapes()

    GetSolidOnLayer = ""

    sLayer = sLayer + ":"

	For i=0 To nShapes-1

		sName = Solid.GetNameOfShapeFromIndex(i)

		If (Left(sName, Len(sLayer)) = sLayer) Then

			GetSolidOnLayer = sName

			Exit For
		End If
	Next i

End Function

Sub JoinPlanesAndVias(nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, dThickness() As Double, nVias As Integer, nViaId() As Integer, dViaStart() As Double, dViaEnd() As Double)

  Dim i As Integer, j As Integer

  For i=0 To nVias-1
     For j=0 To nPlanes-1

	  Dim sViaName As String, sPlaneName As String
      sViaName = GetSolidOnLayer("Via" + CVar(nViaId(i)))
      sPlaneName = GetSolidOnLayer("Plane" + CVar(nPlaneId(j)))

      If (sViaName <> "" And sPlaneName <> "") Then

      	Solid.Insert sViaName, sPlaneName

      End If

     Next j
  Next i

End Sub


Sub MergeConductingLayers(nPlanes As Integer, nPlaneId() As Integer, nVias As Integer, nViaId() As Integer)
  Dim i As Integer
  Dim sName As String

  For i=0 To nPlanes-1
    sName = "Plane" + CVar(nPlaneId(i))
    On Error GoTo Done1
    Solid.MergeMaterialsOfComponent sName
    Done1:
  Next i

  For i=0 To nVias-1
    sName = "Via" + CVar(nViaId(i))
    On Error GoTo Done2
    Solid.MergeMaterialsOfComponent sName
    Done2:
  Next i

End Sub

Sub BuildConductingSlotLayers(nPlanes As Integer, nPlaneId() As Integer, dPlaneThickness() As Double, dPlaneCoords() As Double, sPlaneTypes() As String, dXL As Double, dXH As Double, dYL As Double, dYH As Double)

  Dim i As Integer
  Dim sName As String

  For i=0 To nPlanes-1
    If sPlaneTypes(i) = "slot" Then

	    sName = "Plane" + CVar(nPlaneId(i))
	    On Error GoTo Done1

	    ' determine the name of the metal layer
	    Dim nShapes As Long, index As Long
	    nShapes = Solid.GetNumberOfShapes

		Dim sShapeName As String

		For index = 0 To nShapes
			sShapeName = Solid.GetNameOfShapeFromIndex(index)

			If (Left(sShapeName, InStr(sShapeName, ":")-1) = sName) Then
				Exit For
			End If
		Next index

	    ' build a sheet for the whole layer

		With Brick
		     .Reset
		     .Name "slot"
		     .Layer sName
		     .Xrange CVar(dXL), CVar(dXH)
		     .Yrange CVar(dYL), CVar(dYH)
		     .Zrange CVar(dPlaneCoords(i)), CVar(dPlaneCoords(i)+dPlaneThickness(i))
		     .Create
		End With

	    ' subtract the slots

	    Solid.Subtract sName + ":slot", sShapeName

	    Done1:
	End If
  Next i


End Sub

Sub StoreSubstrateData(dThickness() As Double)
  Dim nLayers As Integer
  nLayers = UBound(dThickness)+1

  StoreGlobalDataValue "Macros\ADS Import\Substrates\Count", CVar(nLayers)

  Dim i As Integer

  For i=0 To nLayers-1
	StoreGlobalDataValue "Macros\ADS Import\Substrates\Thickness" + CVar(i), CVar(dThickness(i))
  Next i

End Sub

Sub BuildSubstrateBricks()
  Dim dXL As Double, dXH As Double, dYL As Double, dYH As Double, nLayers As Integer

  nLayers = GetInt(RestoreGlobalDataValue("Macros\ADS Import\Substrates\Count"))
  RestoreBBoxBounds dXL, dXH, dYL, dYH

  Dim i As Integer
  Dim dFrom As Double
  Dim dTo As Double

  dFrom = 0.0
  dTo = 0.0

  For i=nLayers-1 To 0 STEP -1
    Dim dThick As Double
    dThick = GetDouble(RestoreGlobalDataValue("Macros\ADS Import\Substrates\Thickness" + CVar(i)))

    If (dThick <> 0.0) Then

	    dFrom = dTo
	    dTo   = dFrom + dThick

	    With Brick
		     .Reset
		     .Name "solid1"
		     .Layer "Substrate"  + CVar(i)
		     .Xrange CVar(dXL), CVar(dXH)
		     .Yrange CVar(dYL), CVar(dYH)
		     .Zrange CVar(dFrom), CVar(dTo)
		     .Create
		End With

		Dim sSolidName As String
		sSolidName = "Substrate"  + CVar(i) + ":solid1"

		With Solid
			.SetMeshProperties sSolidName, "PBA", "True"
			.SetAutomeshParameters sSolidName, "0", "True"
			.SetAutomeshStepwidth sSolidName, "0", "0", CVar((dTo - dFrom) / 4.0)
			.SetAutomeshExtendwidth sSolidName, "0", "0", CVar(dTo - dFrom)
		End With

	End If

  Next i

End Sub

Function GetXMLCommand(sLine As String) As String

	Dim sCommand As String
	sCommand = sLine

	Dim nIndex As Integer, nIndexEnd As Integer
	nIndex = InStr(sCommand, "<")
	sCommand = Mid(sCommand, nIndex+1)
	nIndex = InStr(sCommand, ">")
	sCommand = Left(sCommand, nIndex-1)

	nIndex = InStr(sCommand, "=")
	If nIndex <> 0 Then
		sCommand = Left(sCommand, nIndex-1)
	End If

	GetXMLCommand = sCommand

End Function

Function GetXMLValue(sLine As String) As String

	Dim sValue As String
	sValue = sLine

	Dim nIndex As Integer

	nIndex =InStr(sValue, "=")
	If nIndex <> 0 Then
		sValue = Mid(sValue, nIndex+2)
		nIndex = InStr(sValue, ">")
		sValue = Left(sValue, nIndex-2)
	Else
		nIndex = InStr(sValue, ">")
		sValue = Mid(sValue, nIndex+1)
		nIndex = InStr(sValue, "<")
		sValue = Left(sValue, nIndex-1)
	End If

	GetXMLValue = Trim(sValue)

End Function


Function IsXMLCommandEnd(sLine As String, sForCommand As String) As Boolean

	Dim sCommand As String
	sCommand = sLine

	Dim nIndex As Integer
	nIndex = InStr(sCommand, "<")
	sCommand = Mid(sCommand, nIndex+2)
	nIndex = InStr(sCommand, ">")
	sCommand = Left(sCommand, nIndex-1)

	Dim bEnd As Boolean
	bEnd = False
	If( sForCommand = sCommand ) Then
		bEnd = True
	End If

	IsXMLCommandEnd = bEnd

End Function

Function ReadEMStateFile(sFileName As String, dSubstrateLateralExtension As Double, dSubstrateVerticalExtension As Double, sSubstrateBoundaryCond As String) As Boolean

  ReadEMStateFile = False

  Dim OpenError As Boolean
  OpenError = OpenTextFile(sFileName)
  If OpenError Then GoTo finish
  On Error GoTo finish
  On Error Resume Next

  Dim numLinesToPreserve As Integer
  numLinesToPreserve = 1000
  Dim TextLines() As String
  Dim n As Integer, m As Integer

  ReDim Preserve TextLines(numLinesToPreserve)

  n = 0
  TextLines(n) = ReadLine
  While TextLines(n) <> ""
  	If( n > numLinesToPreserve ) Then
  		numLinesToPreserve = numLinesToPreserve + 1000
		ReDim Preserve TextLines(numLinesToPreserve)
  	End If
	n = n+1
	TextLines(n) = ReadLine
  Wend

  CloseTextFile

  Dim sLine As String

  m = n-1

  For n = 0 To m
  	sLine = ""
	sLine = TextLines(n)
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

    	Dim sCommand As String
	    sCommand = GetXMLCommand(sLine)

	    Dim sValue As String
	    sValue = GetXMLValue(sLine)

	    If (sCommand = "substratelateralextensionvalue") Then

			dSubstrateLateralExtension = GetDouble(sValue)

	    ElseIf (sCommand = "substratelateralextensionunit") Then

	    	dSubstrateLateralExtension = dSubstrateLateralExtension * GetScaleFromDropboxNumber(sValue)

	    ElseIf (sCommand = "substrateverticalextensionvalue") Then

	    	dSubstrateVerticalExtension = GetDouble(sValue)

	    ElseIf (sCommand = "substrateverticalextensionunit") Then

	    	dSubstrateVerticalExtension = dSubstrateVerticalExtension * GetScaleFromDropboxNumber(sValue)

	    ElseIf (sCommand = "substratewallboundary") Then

	    	' Dialog: "EM Setup for simulation" FEM -> Options tab -> Global tab -> Substrate wall boundary dropbox:
	    	' 0: Open, 1: Perfect conductor, 2: Perfect MagWall
	    	If (sValue = "0") Then
				sSubstrateBoundaryCond = "open"
	    	ElseIf (sValue = "1") Then
				sSubstrateBoundaryCond = "electric"
	    	ElseIf (sValue = "2") Then
				sSubstrateBoundaryCond = "magnetic"
	    	End If

	    	ReadEMStateFile = True

	    	GoTo Finish

	   End If

   End If

Next n

Finish:

End Function

Function GetPinIndexById(nPin As Integer, nPinId() As Integer, nPinIdToFind As Integer) As Integer

	Dim nIndex As Integer, nCounter As Integer
	nIndex = -1
	nCounter = 0

	While (nCounter < nPin)
		If( nPinId(nCounter) = nPinIdToFind ) Then
			nIndex = nCounter
			nCounter = nPin
		End If
		nCounter = nCounter+1
	Wend

	GetPinIndexById = nIndex

End Function

Function GetShapeNameFromPointCoordinates(dX As Double, dY As Double, dZ As Double) As String

	Dim sShapeName As String, sMaterialName As String, sMaterialType As String
	Dim bIsOnBoundaryOfShape As Boolean

	GetShapeNameFromPointCoordinates = ""

	Dim count%, numTotalObj%
	numTotalObj = Solid.GetNumberOfShapes
	For count = 0 To (numTotalObj-1)

		sShapeName = Solid.GetNameOfShapeFromIndex(count)
		sMaterialName = Solid.GetMaterialNameForShape(sShapeName)
		sMaterialType = Material.GetTypeOfMaterial(sMaterialName)

		If sMaterialType <> "Normal" Then

			bIsOnBoundaryOfShape = Solid.IsPointOnAnyEdgeOfShape(dX, dY, dZ, sShapeName)

			If bIsOnBoundaryOfShape = True Then

				GetShapeNameFromPointCoordinates = sShapeName
				GoTo Found

			End If

		End If
	Next

	Found:

End Function

Function CreateMultipinWaveguidePort(nPortNumber As Integer, sOrientation As String, nPortIndex As Integer, dXRange As Double, dYRange As Double, nPins As Integer, sPinIndexToOrientation() As String, nPinIndexToPort() As Integer, dUPinValues() As Double, dVPinValues() As Double)

	CreateMultipinWaveguidePort = False

	On Error GoTo Failed

    With Port
     .Reset
     .PortNumber CVar(nPortNumber)
     .Label ""
     .NumberOfModes "1"
     .AdjustPolarization "False"
     .PolarizationAngle "0.0"
     .ReferencePlaneDistance "0.0"
     .TextSize "50"
     .TextMaxLimit "1"
     .Coordinates "Full"
     .Orientation sOrientation
     .PortOnBound "False"
     .ClipPickedPortToBound "False"
     .Xrange CVar(dXRange), CVar(dXRange)
     .Yrange CVar(dYRange), CVar(dYRange)
     .Zrange "0", "0"
     .XrangeAdd "0.0", "0.0"
     .YrangeAdd "0.0", "0.0"
     .ZrangeAdd "0.0", "0.0"
     .SingleEnded "False"
     .ConsiderForStructureBoundary "False"

     Dim i As Integer, nPinCnt As Integer
     nPinCnt = 1
  	 For i=0 To nPins-1
		If sPinIndexToOrientation(i) = sOrientation And nPinIndexToPort(i) = nPortIndex Then

			.AddPotentialNumerically CVar(nPinCnt), "positive",  CVar(dUPinValues(i)),  CVar(dVPinValues(i))
			nPinCnt = nPinCnt + 1

		End If
  	 Next i

     .SingleEnded "0"
     .Create
   End With

   CreateMultipinWaveguidePort = True
   GoTo CreationComplete

   Failed:
   CreateMultipinWaveguidePort = False

   CreationComplete:

End Function

Function CreateDiscreteFacePort(nPort As Integer, dX_Plus As Double, dY_Plus As Double, dZ_Plus As Double, dX_Minus As Double, dY_Minus As Double, dZ_Minus As Double, bMoveEdge As Boolean) As Boolean

	CreateDiscreteFacePort = False

	Dim sShapeName_Minus As String, sShapeName_Plus As String
	
	Dim bEdgesPicked As Boolean
	bEdgesPicked = True
	
	If IsBuildingModel() Then
	
    bEdgesPicked = FALSE
    On Error GoTo PickDone
	
    sShapeName_Plus  = GetShapeNameFromPointCoordinates(dX_Plus, dY_Plus, dZ_Plus)
    sShapeName_Minus = GetShapeNameFromPointCoordinates(dX_Minus, dY_Minus, dZ_Minus)

    If sShapeName_Plus <> "" And sShapeName_Minus <> "" Then

      Pick.PickEdgeFromPoint sShapeName_Plus, dX_Plus, dY_Plus, dZ_Plus
      Pick.PickEdgeFromPoint sShapeName_Minus, dX_Minus, dY_Minus, dZ_Minus

      bEdgesPicked = True

    ElseIf bMoveEdge And sShapeName_Plus <> ""Then

      Pick.PickEdgeFromPoint sShapeName_Plus, dX_Plus, dY_Plus, dZ_Plus
      Pick.MoveEdge "-1", "0.0", "0.0", CVar(-dZ_Plus), "True"

      bEdgesPicked = True

    ElseIf bMoveEdge And sShapeName_Minus <> "" Then

      ' ---- This case is not yet tested!
      Pick.PickEdgeFromPoint sShapeName_Minus, dX_Minus, dY_Minus, dZ_Minus
      Pick.MoveEdge "-1", "0.0", "0.0", CVar(-dZ_Minus), "True"

      bEdgesPicked = True

    End If
    
    PickDone:
    
  End If

 	If bEdgesPicked = True Then

		On Error GoTo Failed

		With DiscreteFacePort
		     .Reset
		     .PortNumber CVar(nPort)
		     .Type "SParameter"
		     .Label ""
		     .Impedance "50.0"
		     .VoltagePortImpedance "0.0"
		     .VoltageAmplitude "1.0"
		     .SetP1 "True", CVar(dX_Plus), CVar(dY_Plus), CVar(dZ_Plus)
		     .SetP2 "True", CVar(dX_Minus), CVar(dY_Minus), CVar(dZ_Minus)
		     .LocalCoordinates "False"
		     .InvertDirection "False"
		     .CenterEdge "True"
		     .Monitor "True"
		     .UseProjection "True"
		     .ReverseProjection "False"
		     .Create
		End With

		CreateDiscreteFacePort = True
		GoTo CreationComplete

		Failed:
		CreateDiscreteFacePort = False

	End If

	CreationComplete:

End Function

Function GetPortLocationIfOnBoundary(dX As Double, dY As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double) As String

	GetPortLocationIfOnBoundary = ""

  	Dim dDeltaX As Double, dDeltaY As Double
    dDeltaX = (dXH - dXL) / 1000.0
    dDeltaY = (dYH - dYL) / 1000.0

    Dim bXmin As Boolean, bYmin As Boolean, bXmax As Boolean, bYmax As Boolean

    bXmin = False
    bXmax = False
    bYmin = False
    bYmax = False

    If (Abs(dX - dXL) < dDeltaX) Then bXmin = True
    If (Abs(dX - dXH) < dDeltaX) Then bXmax = True
    If (Abs(dY - dYL) < dDeltaY) Then bYmin = True
    If (Abs(dY - dYH) < dDeltaY) Then bYmax = True

    Dim nCount As Integer
    nCount = 0
    If (bXmin) Then nCount = nCount+1
    If (bXmax) Then nCount = nCount+1
    If (bYmin) Then nCount = nCount+1
    If (bYmax) Then nCount = nCount+1

    If nCount = 1 Then
        If (bXmin) Then GetPortLocationIfOnBoundary = "XMin"
        If (bXmax) Then GetPortLocationIfOnBoundary = "XMax"
        If (bYmin) Then GetPortLocationIfOnBoundary = "YMin"
        If (bYmax) Then GetPortLocationIfOnBoundary = "YMax"
    End If

End Function

Function ReadPins(sPinFileName As String, nPinIds() As Integer, nPlanesForPins() As Integer, dXPinValues() As Double, dYPinValues() As Double, nPins As Integer)

  ReadPins = False

  ' <!-- note: all pin coordinates are in meter --> multiply here by 1e3
  Dim dScale As Double
  dScale = 1e6

  Open sPinFileName For Input As #1
  On Error GoTo finish

  Dim sLine As String
  Dim sCommand As String
  Dim sValue As String

  Do
    Line Input #1, sLine
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

	    sCommand = GetXMLCommand(sLine)

	    If (sCommand = "pin") Then

	    	Dim nPlane As Integer
	    	Dim nPin As Integer
	    	Dim dX As Double, dY As Double, dZ As Double

	    	Line Input #1, sLine
	    	sLine = Trim(LCase(sLine))
	      	While (IsXMLCommandEnd(sLine, "pin") = False)

				sCommand = GetXMLCommand(sLine)
				If (sCommand = "name") Then
					sValue = GetXMLValue(sLine)
					nPin = GetInt(Mid(sValue, 2))
				ElseIf (sCommand = "layer") Then
	      			sValue = GetXMLValue(sLine)
					nPlane = GetInt(sValue)
				ElseIf (sCommand = "x") Then
					sValue = GetXMLValue(sLine)
					dX = GetDouble(sValue) * dScale
				ElseIf (sCommand = "y") Then
					sValue = GetXMLValue(sLine)
					dY = GetDouble(sValue) * dScale
				End If

	      		Line Input #1, sLine
	    		sLine = Trim(LCase(sLine))
	      	Wend

	        ReDim Preserve nPinIds(nPins)
			ReDim Preserve dXPinValues(nPins)
			ReDim Preserve dYPinValues(nPins)
			ReDim Preserve nPlanesForPins(nPins)

	        nPinIds(nPins) = nPin
	        dXPinValues(nPins) = dX
	        dYPinValues(nPins) = dY
	        nPlanesForPins(nPins) = nPlane

	        nPins = nPins + 1

    	End If
    End If

  Loop Until sLine = ""

  ReadPins = True

  Finish:
  Close #1

End Function

Sub AddPort( nPorts As Integer, dPortValues() As Double, nPortNum() As Integer, dValue As Double, nPortNumber As Integer, sPinIndexToOrientation() As String, nPinIndexToPort() As Integer, pinIndex As Integer, sLocation As String)

	Dim nPortIndex As Integer
	Dim i As Integer

	For i=0 To nPorts-1
		If dPortValues(i) = dValue Then
			nPortIndex = i
			GoTo SetPinValues1
		End If
	Next i

	ReDim Preserve dPortValues(nPorts)
	ReDim Preserve nPortNum(nPorts)
	dPortValues(nPorts) = dValue
	nPortNum(nPorts) = nPortNumber

	nPortIndex = nPorts
	nPorts = nPorts + 1

	SetPinValues1:

	sPinIndexToOrientation(pinIndex) = LCase(sLocation)
	nPinIndexToPort(pinIndex) = nPortIndex

End Sub

Sub ReadPinsAndPorts(sPinFileName As String, sPortFileName As String, nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, dPlaneThickness() As Double, dXL As Double, dXH As Double, dYL As Double, dYH As Double, bCreateWaveguidePorts As Boolean, bUseWaveguidePorts As Boolean)

  StoreGlobalDataValue "Macros\ADS Import\Ports\XMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\XMax", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMax", "0"

  Dim nPinIds() As Integer, nPlanesForPins() As Integer
  Dim dXPinValues() As Double, dYPinValues() As Double, dZPinValues() As Double
  Dim nPins As Integer
  nPins = 0

  ReadPins(sPinFileName, nPinIds(), nPlanesForPins(), dXPinValues(), dYPinValues(), nPins)

  Open sPortFileName For Input As #2
  On Error GoTo finish2

  ' --------------------- Multipin waveguide ports ----------------------------------------------
  ' For creating the multipin waveguide ports we have to collect the following data:
  ' Ports can be in xmin, xmax, ymin, and ymax orientation.
  ' Note we want to create inner mulitpin ports therefore we have for example in xmin orientation to store the xmin double values.
  ' Several pins may be related to the same xmin double value.
  '
  ' Collect in a list the double values along xmin: dXMin_Ports (same for xmax, ymin, ymax)
  Dim nXMin_Ports As Integer, nXMax_Ports As Integer, nYMin_Ports As Integer, nYMax_Ports As Integer
  Dim dXMin_Ports() As Double, dXMax_Ports() As Double, dYMin_Ports() As Double, dYMax_Ports() As Double
  ' For each pin we need to store to which port it is related and to which port orientation ("xmin", "xmax", "ymin", or "ymax")
  ' For example nPinIndexToPort(0) = 1 and sPinIndexToOrientation(0) = "xmin" -> dXMin_Ports(1)
  Dim nPinIndexToPort() As Integer
  Dim sPinIndexToOrientation() As String
  ' Collect in a list the port numbers (index of nXMin_PortNum is related to the index of dXMin_Layers)
  Dim nXMin_PortNum() As Integer, nXMax_PortNum() As Integer, nYMin_PortNum() As Integer, nYMax_PortNum() As Integer

  nXMin_Ports = 0
  nXMax_Ports = 0
  nYMin_Ports = 0
  nYMax_Ports = 0
  ReDim Preserve sPinIndexToOrientation(nPins)
  ReDim Preserve nPinIndexToPort(nPins)
  ReDim Preserve dZPinValues(nPins)

  Dim sLine As String
  Dim sCommand As String
  Dim sValue As String

  Do
    Line Input #2, sLine
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

	    sCommand = GetXMLCommand(sLine)

	    If (sCommand = "port id") Then

	    	' ------------- Read port number, minus pin and plus pin -------------

	    	Dim nPortNumber As Integer
	    	Dim nPlusPin As Integer
	    	Dim nMinusPin As Integer

	    	nPlusPin = -1
	    	nMinusPin = -1

	    	sValue = GetXMLValue(sLine)
			nPortNumber = GetInt(sValue)

	    	Line Input #2, sLine
	    	sLine = Trim(LCase(sLine))
	      	While (IsXMLCommandEnd(sLine, "port") = False)

	      		sCommand = GetXMLCommand(sLine)
				If (sCommand = "plus_pin") Then
	      			sValue = GetXMLValue(sLine)
					nPlusPin = GetInt(Mid(sValue, 2))
				ElseIf (sCommand = "minus_pin") Then
	      			sValue = GetXMLValue(sLine)
					nMinusPin = GetInt(Mid(sValue, 2))
				End If

	      		Line Input #2, sLine
	    		sLine = Trim(LCase(sLine))

	      	Wend

	      	' No pins have been defined -> no port can be created.
	      	If nPlusPin = -1 And nMinusPin = -1 Then
	      		GoTo ReadNextLine
	      	End If

	      	' ------------- Determine X, Y, Z Coordinates for minus and plus pin -------------

	      	Dim bMovePickedEdge As Boolean
	      	Dim dX_Plus As Double, dY_Plus As Double, dZ_Plus As Double
	      	Dim dX_Minus As Double, dY_Minus As Double, dZ_Minus As Double
	      	Dim pinIndex_Plus As Integer, pinIndex_Minus As Integer
			Dim nPinPlane As Integer

			' Move picked edge for discrete face port only when one pin is defined in .prt file for current port
			bMovePickedEdge = False

	      	If nPlusPin <> -1 Then
				pinIndex_Plus = GetPinIndexById(nPins, nPinIds(), nPlusPin)

	      		nPinPlane = nPlanesForPins(pinIndex_Plus)
	      		dZ_Plus = GetPlaneCoord(nPinPlane, nPlanes, nPlaneId(), dPlaneCoords())
	      	End If

	      	If nMinusPin <> -1 Then
				pinIndex_Minus = GetPinIndexById(nPins, nPinIds(), nMinusPin)

	      		nPinPlane = nPlanesForPins(pinIndex_Minus)
	      		dZ_Minus = GetPlaneCoord(nPinPlane, nPlanes, nPlaneId(), dPlaneCoords())
	      	End If

	      	If nPlusPin = -1 Then
	      		' ---- This case is not yet tested!
				pinIndex_Plus = pinIndex_Minus
	      		dZ_Plus  = 0.0
	      		bMovePickedEdge = True
	      	End If

	      	If nMinusPin = -1 Then
	      		pinIndex_Minus = pinIndex_Plus
	      		dZ_Minus = 0.0
	      		bMovePickedEdge = True
	      	End If

	      	dX_Plus = dXPinValues(pinIndex_Plus)
	      	dY_Plus = dYPinValues(pinIndex_Plus)

	      	dX_Minus = dXPinValues(pinIndex_Minus)
	      	dY_Minus = dYPinValues(pinIndex_Minus)

	      	' ------------- Now create the ports -------------
	      	' Note: Ports are created from plus pin to minus pin.

	      	Dim sLocation As String

	      	' bUseWaveguidePorts: user wants to create waveguide ports on boundary.
	      	' bCreateWaveguidePorts: there is an additional space definied to be added to the bounding box.
	      	If bUseWaveguidePorts = True Then
	      		sLocation = GetPortLocationIfOnBoundary(dX_Plus, dY_Plus, dXL, dXH, dYL, dYH)
			End If

			' The port is located at a boundary, and the user wants to create waveguide ports.
			If sLocation <> "" Then

			   ' Store all informations to create later the multipin waveguide ports.
			   dZPinValues(pinIndex_Plus) = dZ_Plus

		       If sLocation = "XMin" Then
		       		AddPort( nXMin_Ports, dXMin_Ports(), nXMin_PortNum(), dX_Plus, nPortNumber, sPinIndexToOrientation(), nPinIndexToPort(), pinIndex_Plus, sLocation )
  			   ElseIf sLocation = "XMax" Then
  			   		AddPort( nXMax_Ports, dXMax_Ports(), nXMax_PortNum(), dX_Plus, nPortNumber, sPinIndexToOrientation(), nPinIndexToPort(), pinIndex_Plus, sLocation )
  			   ElseIf sLocation = "YMin" Then
  			   		AddPort( nYMin_Ports, dYMin_Ports(), nYMin_PortNum(), dY_Plus, nPortNumber, sPinIndexToOrientation(), nPinIndexToPort(), pinIndex_Plus, sLocation )
  			   Else
  			   		AddPort( nYMax_Ports, dYMax_Ports(), nYMax_PortNum(), dY_Plus, nPortNumber, sPinIndexToOrientation(), nPinIndexToPort(), pinIndex_Plus, sLocation )
		       End If

			' The port is located inside the structure
	      	ElseIf CreateDiscreteFacePort(nPortNumber, dX_Plus, dY_Plus, dZ_Plus, dX_Minus, dY_Minus, dZ_Minus, bMovePickedEdge) = False Then

				With DiscretePort
				     .Reset
				     .PortNumber CVar(nPortNumber)
				     .Type "SParameter"
				     .Impedance "50.0"
				     .Voltage "1.0"
				     .Current "1.0"
				     .Point1 CVar(dX_Plus), CVar(dY_Plus), CVar(dZ_Plus)
				     .Point2 CVar(dX_Minus), CVar(dY_Minus), CVar(dZ_Minus)
				     .UsePickedPoints "False"
				     .LocalCoordinates "False"
				     .Monitor "False"
				     .Create
				End With

	      	End If

	    End If

	End If

	ReadNextLine:

  Loop Until sLine = ""


Finish2:
  Close #2

  ' Now create multipin waveguide ports in x and y directions.
  Dim n As Integer
  For n=0 To nXMin_Ports-1
  	CreateMultipinWaveguidePort(nXMin_PortNum(n), "xmin", n, dXMin_Ports(n), 0.0, nPins, sPinIndexToOrientation, nPinIndexToPort, dYPinValues, dZPinValues)
  Next n
  For n=0 To nXMax_Ports-1
  	CreateMultipinWaveguidePort(nXMax_PortNum(n), "xmax", n, dXMax_Ports(n), 0.0, nPins, sPinIndexToOrientation, nPinIndexToPort, dYPinValues, dZPinValues)
  Next n
  For n=0 To nYMin_Ports-1
  	CreateMultipinWaveguidePort(nYMin_PortNum(n), "ymin", n, 0.0, dYMin_Ports(n), nPins, sPinIndexToOrientation, nPinIndexToPort, dZPinValues, dXPinValues)
  Next n
  For n=0 To nYMax_Ports-1
  	CreateMultipinWaveguidePort(nYMax_PortNum(n), "ymax", n, 0.0, dYMax_Ports(n), nPins, sPinIndexToOrientation, nPinIndexToPort, dZPinValues, dXPinValues)
  Next n

End Sub


Sub ReadNewPortFiles(sPinFileName As String, dXL As Double, dXH As Double, dYL As Double, dYH As Double, nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, dPlaneThickness() As Double, bCreateWaveguidePorts As Boolean)

  StoreGlobalDataValue "Macros\ADS Import\Ports\XMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\XMax", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMax", "0"

' <!-- note: all port coordinates are in meter --> multiply here by 1e3
  Dim dScale As Double
  dScale = 1e6

  Open sPinFileName For Input As #1
  On Error GoTo finish

  Dim sLine As String

  Do
    Line Input #1, sLine
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

    	Dim sCommand As String
	    sCommand = GetXMLCommand(sLine)

	    If (sCommand = "pin") Then

	    	Dim sValue As String
	    	Dim nPlane As Integer
	    	Dim nPort As Integer
	    	Dim dX As Double, dY As Double, dZ As Double

	    	Line Input #1, sLine
	    	sLine = Trim(LCase(sLine))
	      	While (IsXMLCommandEnd(sLine, "pin") = False)

				sCommand = GetXMLCommand(sLine)
				If (sCommand = "name") Then
					sValue = GetXMLValue(sLine)
					nPort = GetInt(Mid(sValue, 2))
				ElseIf (sCommand = "layer") Then
	      			sValue = GetXMLValue(sLine)
					nPlane = GetInt(sValue)
				ElseIf (sCommand = "x") Then
					sValue = GetXMLValue(sLine)
					dX = GetDouble(sValue) * dScale
				ElseIf (sCommand = "y") Then
					sValue = GetXMLValue(sLine)
					dY = GetDouble(sValue) * dScale
				End If

	      		Line Input #1, sLine
	    		sLine = Trim(LCase(sLine))
	      	Wend

	        dZ = GetPlaneCoord(nPlane, nPlanes, nPlaneId(), dPlaneCoords()) + 0.5 * GetPlaneThickness(nPlane, nPlanes, nPlaneId(), dPlaneThickness())

	        Dim dDeltaX As Double, dDeltaY As Double
	        dDeltaX = (dXH - dXL) / 1000.0
	        dDeltaY = (dYH - dYL) / 1000.0

	        Dim bXmin As Boolean, bYmin As Boolean, bXmax As Boolean, bYmax As Boolean

	        bXmin = False
	        bXmax = False
	        bYmin = False
	        bYmax = False

	        If (Abs(dX - dXL) < dDeltaX) Then bXmin = True
	        If (Abs(dX - dXH) < dDeltaX) Then bXmax = True
	        If (Abs(dY - dYL) < dDeltaY) Then bYmin = True
	        If (Abs(dY - dYH) < dDeltaY) Then bYmax = True

	        Dim nCount As Integer
	        nCount = 0
	        If (bXmin) Then nCount = nCount+1
	        If (bXmax) Then nCount = nCount+1
	        If (bYmin) Then nCount = nCount+1
	        If (bYmax) Then nCount = nCount+1

	        If (nCount = 1 And bCreateWaveguidePorts = True) Then
	          ' The port is located at a boundary

	            Dim sLocation As String
		        If (bXmin) Then sLocation = "XMin"
		        If (bXmax) Then sLocation = "XMax"
		        If (bYmin) Then sLocation = "YMin"
		        If (bYmax) Then sLocation = "YMax"

		        If (RestoreGlobalDataValue("Macros\ADS Import\Ports\" + sLocation) = "0") Then

			        With Port
				     .Reset
				     .PortNumber CVar(nPort)
				     .NumberOfModes "1"
				     .AdjustPolarization False
				     .PolarizationAngle "0.0"
				     .ReferencePlaneDistance "0"
				     .TextSize "50"
				     .Coordinates "Full"
				     .Location LCase(Trim(sLocation))
				     .Create
					End With

					StoreGlobalDataValue "Macros\ADS Import\Ports\" + sLocation, "1"
					StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "X", CVar(dX)
					StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "Y", CVar(dY)
					StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "Z", CVar(dZ)
					StoreGlobalDataValue "Macros\ADS Import\Ports\PN" + sLocation, CVar(nPort)
				End If

			Else
			  ' The port is located inside the structure

				With DiscretePort
				     .Reset
				     .PortNumber CVar(nPort)
				     .Type "SParameter"
				     .Impedance "50.0"
				     .Voltage "1.0"
				     .Current "1.0"
				     .Point1 CVar(dX), CVar(dY), CVar(dZ)
				     .Point2 CVar(dX), CVar(dY), "0.0"
				     .UsePickedPoints "False"
				     .LocalCoordinates "False"
				     .Monitor "False"
				     .Create
				End With

	        End If

    	End If

      End If

  Loop Until sLine = ""

Finish:
  Close #1

End Sub

Sub ReadPortFile(sFileName As String, dXL As Double, dXH As Double, dYL As Double, dYH As Double, nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, dPlaneThickness() As Double)

  StoreGlobalDataValue "Macros\ADS Import\Ports\XMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\XMax", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMin", "0"
  StoreGlobalDataValue "Macros\ADS Import\Ports\YMax", "0"

  Open sFileName For Input As #1
  On Error GoTo finish

  Dim sLine As String

  Dim dScale As Double
  dScale = 1.0

  Do
    Line Input #1, sLine
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

      Dim sBlock As String
      sBlock = sLine

      Dim sCommand As String, sWhat As String
      sCommand = Trim(GetItem(sBlock))

	  Select Case sCommand
	  Case "units"
	    sWhat = Trim(GetItem(sBlock))
		Dim nColonIndex As Integer

		nColonIndex = InStr(sWhat, ",")
		sWhat = Trim(Left(sWhat, nColonIndex-1))

   		dScale = GetScaleFromName(sWhat)

	  Case "add"
	    sWhat = Trim(GetItem(sBlock))

	    Dim nPlane As Integer
	    nPlane = GetInt(Mid(sWhat, 2))

	    Dim nIndex As Integer
	    nIndex = InStr(sBlock, "'")

	    sBlock = Mid(sBlock, nIndex+1)

	    nIndex = InStr(sBlock, ",")
	    Dim nPort As Integer

	    nPort = GetInt(Left(sBlock, nIndex-1))

        nIndex = InStr(sBlock, "'")
        sBlock = Mid(sBlock, nIndex+1)

        Dim sX As String, sY As String

        nIndex = InStr(sBlock, ",")
        sX = Left(sBlock, nIndex-1)
        sY = Mid(sBlock, nIndex+1)

        nIndex = InStr(sY, ";")
        sY = Left(sY, nIndex-1)

        Dim dX As Double, dY As Double, dZ As Double
        dX = GetDouble(sX) * dScale
        dY = GetDouble(sY) * dScale
        dZ = GetPlaneCoord(nPlane, nPlanes, nPlaneId(), dPlaneCoords()) + 0.5 * GetPlaneThickness(nPlane, nPlanes, nPlaneId(), dPlaneThickness())

        Dim dDeltaX As Double, dDeltaY As Double
        dDeltaX = (dXH - dXL) / 1000.0
        dDeltaY = (dYH - dYL) / 1000.0

        Dim bXmin As Boolean, bYmin As Boolean, bXmax As Boolean, bYmax As Boolean

        bXmin = False
        bXmax = False
        bYmin = False
        bYmax = False

        If (Abs(dX - dXL) < dDeltaX) Then bXmin = True
        If (Abs(dX - dXH) < dDeltaX) Then bXmax = True
        If (Abs(dY - dYL) < dDeltaY) Then bYmin = True
        If (Abs(dY - dYH) < dDeltaY) Then bYmax = True

        Dim nCount As Integer
        nCount = 0
        If (bXmin) Then nCount = nCount+1
        If (bXmax) Then nCount = nCount+1
        If (bYmin) Then nCount = nCount+1
        If (bYmax) Then nCount = nCount+1

        If (nCount = 1) Then
          ' The port is located at a boundary

            Dim sLocation As String
	        If (bXmin) Then sLocation = "XMin"
	        If (bXmax) Then sLocation = "XMax"
	        If (bYmin) Then sLocation = "YMin"
	        If (bYmax) Then sLocation = "YMax"

	        If (RestoreGlobalDataValue("Macros\ADS Import\Ports\" + sLocation) = "0") Then

		        With Port
			     .Reset
			     .PortNumber CVar(nPort)
			     .NumberOfModes "1"
			     .AdjustPolarization False
			     .PolarizationAngle "0.0"
			     .ReferencePlaneDistance "0"
			     .TextSize "50"
			     .Coordinates "Full"
			     .Location LCase(Trim(sLocation))
			     .Create
				End With

				StoreGlobalDataValue "Macros\ADS Import\Ports\" + sLocation, "1"
				StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "X", CVar(dX)
				StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "Y", CVar(dY)
				StoreGlobalDataValue "Macros\ADS Import\Ports\P" + CVar(nPort) + "Z", CVar(dZ)
				StoreGlobalDataValue "Macros\ADS Import\Ports\PN" + sLocation, CVar(nPort)
			End If

		Else
		  ' The port is located inside the structure

			With DiscretePort
			     .Reset
			     .PortNumber CVar(nPort)
			     .Type "SParameter"
			     .Impedance "50.0"
			     .Voltage "1.0"
			     .Current "1.0"
			     .Point1 CVar(dX), CVar(dY), CVar(dZ)
			     .Point2 CVar(dX), CVar(dY), "0.0"
			     .UsePickedPoints "False"
			     .LocalCoordinates "False"
			     .Monitor "False"
			     .Create
			End With

        End If

      End Select

    End If

  Loop Until sLine = ""

Finish:
  Close #1

End Sub

Sub ReadStimulationFile(sFileName As String)

  Open sFileName For Input As #1
  On Error GoTo Finish

  Dim sLine As String
  Dim dStart As Double, dStop As Double
  Dim bFrequencyRead As Boolean

  dStart = 1e10
  dStop  = 0
  bFrequencyRead = False

  Do
    Line Input #1, sLine
    sLine = Trim(LCase(sLine))

    If (sLine <> "") Then

        Dim sBlock As String
        sBlock = sLine

        Dim sCommand As String, sWhat As String
        sCommand = Trim(GetItem(sBlock))

        Dim dValue As Double

		Select Case sCommand
	    Case "start"

	      dValue = GetDouble(GetItem(sBlock))
	      dStart = IIf(dValue < dStart, dValue, dStart)

	      sWhat = Trim(GetItem(sBlock))

	      If (sWhat = "stop") Then
	        dValue = GetDouble(GetItem(sBlock))
	      	dStop = IIf(dValue > dStop, dValue, dStop)

	        bFrequencyRead = True
	      End If

        End Select

    End If
  Loop Until sLine = ""

Finish:
  Close #1

If (bFrequencyRead) Then
    Solver.FrequencyRange CVar(dStart), CVar(dStop)
End If

End Sub

Function DoesFileExist(sFileName As String)

  Dim bCouldBeOpened As Boolean
  bCouldBeOpened = False

  On Error GoTo Finish

  Open sFileName For Input As #1
  Close #1
  bCouldBeOpened = True

  Finish:

  DoesFileExist = bCouldBeOpened

End Function

Sub ConvertToDos(sFileName As String)

  On Error GoTo Finish

  Open sFileName For Input As #1
  Dim sLine As String
  Line Input #1, sLine
  Close #1

  If (InStr(sLine, vbLf) <> 0) Then

    Dim nIndex As Long, nEndIndex As Long
    nIndex = 1
    nEndIndex = 1

    Open sFileName For Output As #1

    While (nEndIndex <> 0)

	    nEndIndex = InStr(nIndex, sLine, vbLf)

	    If (nEndIndex = 0) Then
		    Print #1, Mid(sLine, nIndex)
	    Else
		    Print #1, Mid(sLine, nIndex, nEndIndex-nIndex)
		End If

    	nIndex = nEndIndex+1

    Wend

    Close #1

  End If

Finish:

End Sub

Sub Main ()

	Dim doubleTest As Double
	doubleTest = CVar("1.0")
    bReplaceDotByColon = IIf(doubleTest > 5.0, True, False)
	bShowOhmicSheetWarning = False

    Dim sProject                 As String
    Dim sSimplifyAngle           As String
    Dim sSimplifyAdjacentTol     As String
    Dim sSimplifyRadiusTol       As String
    Dim sSimplifyMinPointsArc    As String
    Dim sSimplifyMinPointsCircle As String
    Dim sSimplifyAngleTang       As String
    Dim sSimplifyEdgeLength      As String
    Dim sMinimumMetalThickness   As String
    Dim sUseSheets               As String
    Dim sUseSimplification       As String
    Dim sUseWaveguidePorts       As String

    sProject                 = RestoreGlobalDataValue("Macros\ADS Import\FileName")
    sSimplifyAngle           = RestoreGlobalDataValue("Macros\ADS Import\SimplifyAngle")
    sSimplifyAdjacentTol     = RestoreGlobalDataValue("Macros\ADS Import\SimplifyAdjacentTol")
    sSimplifyRadiusTol       = RestoreGlobalDataValue("Macros\ADS Import\SimplifyRadiusTol")
    sSimplifyMinPointsArc    = RestoreGlobalDataValue("Macros\ADS Import\SimplifyMinPointsArc")
    sSimplifyMinPointsCircle = RestoreGlobalDataValue("Macros\ADS Import\SimplifyMinPointsCircle")
    sSimplifyAngleTang       = RestoreGlobalDataValue("Macros\ADS Import\SimplifyAngleTang")
    sSimplifyEdgeLength      = RestoreGlobalDataValue("Macros\ADS Import\SimplifyEdgeLength")
    sMinimumMetalThickness   = RestoreGlobalDataValue("Macros\ADS Import\MinimumMetalThickness")
    sUseSheets               = RestoreGlobalDataValue("Macros\ADS Import\UseSheets")
    sUseSimplification       = RestoreGlobalDataValue("Macros\ADS Import\UseSimplification")
    sUseWaveguidePorts       = RestoreGlobalDataValue("Macros\ADS Import\UseWaveguidePorts")

    'sProject                 = "D:\examples\ads\test\proj_a"

	Dim nIndex As Integer
	nIndex = InStrRev(sProject, "*")
	If (nIndex <> 0) Then
		Dim sProjectPath3D As String
		sProjectPath3D = getprojectpath("Project") + "\Model\3D\"
		sProject = sProjectPath3D + Mid(sProject, nIndex+1)
	End If

    If (sSimplifyAngle = "") Then   ' if not found, should hapen in old projects.
        sSimplifyAngle = "0.0"
    End If
    If (sSimplifyAdjacentTol = "") Then   ' if not found, should hapen in old projects.
        sSimplifyAdjacentTol = "0.0"
    End If
    If (sSimplifyRadiusTol = "") Then   ' if not found, should hapen in old projects.
        sSimplifyRadiusTol = "0.0"
    End If
    If (sSimplifyMinPointsArc = "") Then   ' if not found, should hapen in old projects.
        sSimplifyMinPointsArc = "0"
    End If
    If (sSimplifyMinPointsCircle = "") Then   ' if not found, should hapen in old projects.
        sSimplifyMinPointsCircle = "0"
    End If
    If (sSimplifyAngleTang = "") Then   ' if not found, should hapen in old projects.
        sSimplifyAngleTang = "0.0"
    End If
    If (sSimplifyEdgeLength = "") Then   ' if not found, should hapen in old projects.
        sSimplifyEdgeLength = "0.0"
    End If
    If (sMinimumMetalThickness = "") Then   ' if not found, should hapen in old projects.
        sMinimumMetalThickness = "0.0"
    End If
    If (sUseSheets = "") Then
        sUseSheets = "False"
    End If
    If (sUseSimplification = "") Then
        sUseSimplification = "False"
    End If
    If (sUseWaveguidePorts = "") Then
        sUseWaveguidePorts = "False"
    End If

    Dim dSimplifyAngle           As Double
    Dim dSimplifyAdjacentTol     As Double
    Dim dSimplifyRadiusTol       As Double
    Dim iSimplifyMinPointsArc    As Integer
    Dim iSimplifyMinPointsCircle As Integer
    Dim dSimplifyAngleTang       As Double
    Dim dSimplifyEdgeLength      As Double
    Dim dMinimumMetalThickness   As Double
    Dim bUseSheets               As Boolean
    Dim bUseWaveguidePorts       As Boolean

    dSimplifyAngle           = Evaluate(sSimplifyAngle)
    dSimplifyAdjacentTol     = Evaluate(sSimplifyAdjacentTol)
    dSimplifyRadiusTol       = Evaluate(sSimplifyRadiusTol)
    iSimplifyMinPointsArc    = Evaluate(sSimplifyMinPointsArc)
    iSimplifyMinPointsCircle = Evaluate(sSimplifyMinPointsCircle)
    dSimplifyAngleTang       = Evaluate(sSimplifyAngleTang)
    dSimplifyEdgeLength      = Evaluate(sSimplifyEdgeLength)
    dMinimumMetalThickness   = Evaluate(sMinimumMetalThickness)
    bUseSheets               = CBool(sUseSheets)
    bUseWaveguidePorts       = CBool(sUseWaveguidePorts)

    If (sProject <> "") Then

        Dim nBackSlashIndex As Integer
        nBackSlashIndex = InStrRev(sProject, "\")

        sProject = Left(sProject, nBackSlashIndex)

	    SetLock True
	   	ScreenUpdating False
	   	SetNoMouseSelection True
	   	LockTree True
	    ResultTree.EnableTreeUpdate False

	    With Units
	     .Geometry "um"
	     .Frequency "ghz"
	     .Time "s"
	    End With

	    Dim dThickness() As Double, dEpsilon() As Double, dMue() As Double, dLossTang() As Double, dLossCond() As Double
	    Dim nPlanes As Integer, nPlaneId() As Integer, dPlaneCoords() As Double, nLayerFrom() As Integer
	    Dim nVias As Integer, nViaId() As Integer, dViaStart() As Double, dViaEnd() As Double
	    Dim nViaStart() As Integer, nViaEnd() As Integer
	    Dim sPlaneTypes() As String, dPlaneThickness() As Double, dPlaneCond() As Double, dPlaneImpedance() As Double
	    Dim sLayerNames() As String, sTopPlane As String, sBottomPlane As String
        Dim bBoundaryOutlineDefined As Boolean
	    ' Wenn diese Funktion als Makro ausgefuehrt wird, dann kann das Projekt den Projektbasisnamen auslesen.

	    ' Datei gelesen ab >=CST Version 2015 SP4
	    ConvertToDos sProject + "emStateFile.xml"
	    ' Dateien des neuen ADS Formates (>=CST Version 2011 SP3).
	    ConvertToDos sProject + "proj.ltd"
	    ConvertToDos sProject + "proj.pin"
	    ConvertToDos sProject + "proj.prt"
		' Dateien des alten ADS Formates.
	    ConvertToDos sProject + "proj.sub"
	    ConvertToDos sProject + "proj.lmp"
		' Dateien fuer alle ADS Formate.
	    ConvertToDos sProject + "proj.sti"
	    ConvertToDos sProject + "proj"
	    ConvertToDos sProject + "proj_a"

		nPlanes = 0
	    nVias   = 0
		bBoundaryOutlineDefined = False

	    If (DoesFileExist(sProject + "proj.ltd")) Then

	      ' New ADS import
	      ' Diese Funktion liest die gesamte Layerinformation ein (dh. Substratlayer und Planelayer).
	      ReadLayersAndPlanes sProject + "proj.ltd", sLayerNames(), dThickness(), dEpsilon(), dMue(), dLossTang(), dLossCond(), sTopPlane, sBottomPlane, nPlanes, nPlaneId(), nLayerFrom(), sPlaneTypes(), dPlaneCond(), dPlaneImpedance(), dPlaneThickness(), nVias, nViaId(), nViaStart(), nViaEnd(), bBoundaryOutlineDefined

	    Else
	      If (DoesFileExist(sProject + "proj.sub")) Then

	      ' Old ADS import
	      ' Diese Funktion liest die gesamte Layerinformation ein und speichert die Dicke, Epsilon, Mue in Arrays ab.
	      ReadLayers sProject + "proj.sub", dThickness, dEpsilon, dMue, dLossTang, dLossCond, sTopPlane, sBottomPlane
	      ReadPlanes sProject + "proj.lmp", nPlanes, nPlaneId(), nLayerFrom(), sPlaneTypes(), dPlaneCond(), dPlaneImpedance(), dPlaneThickness(), nVias, nViaId(), nViaStart(), nViaEnd(), dThickness()

	      Else
	        GoTo IncompleteData:
	      End If
	    End If

	    UpdateMetalThickness bUseSheets, dMinimumMetalThickness, nPlanes, dPlaneThickness()

	    UpdateSubstrateThickness nPlanes, nVias, dPlaneThickness(), dThickness(), nLayerFrom(), sPlaneTypes(), nViaStart(), nViaEnd(), dPlaneCoords(), dViaStart(), dViaEnd()

	    Dim dXL As Double, dXH As Double, dYL As Double, dYH As Double
	    dXL = 1e30
	    dXH = -1e10
	    dYL = 1e30
	    dYH = -1e30

	    Dim sGeometryFileName As String
	    sGeometryFileName = "proj_a"

	    Dim bExists As Boolean
	    bExists = False

	    On Error GoTo Fail:
	      GetAttr(sProject + sGeometryFileName)
		  bExists = True
	    Fail:

	    If (Not bExists) Then sGeometryFileName = "proj"

	    Dim nEntitiesRead As Long
	    nEntitiesRead = 0

	    Dim bPolygonFailed As Boolean
	    bPolygonFailed = False

	    ReadGeometry sProject + sGeometryFileName, sPlaneTypes(), dPlaneThickness(), nPlanes, nPlaneId(), dPlaneCoords(), dPlaneCond(), dPlaneImpedance(), nVias, nViaId(), dViaStart(), dViaEnd(), dXL, dXH, dYL, dYH, nEntitiesRead, bPolygonFailed, sUseSimplification, iSimplifyMinPointsArc, iSimplifyMinPointsCircle, dSimplifyAngle, dSimplifyAdjacentTol, dSimplifyRadiusTol, dSimplifyAngleTang, dSimplifyEdgeLength, dMinimumMetalThickness

      IncompleteData:
      
	    If (nEntitiesRead > 0) Then

		    MergeConductingLayers nPlanes, nPlaneId(), nVias, nViaId()
		    JoinPlanesAndVias nPlanes, nPlaneId(), dPlaneCoords(), dThickness(), nVias, nViaId(), dViaStart(), dViaEnd()

		    StoreStructureBounds dXL, dXH, dYL, dYH
		    StoreBBoxBounds dXL, dXH, dYL, dYH

		    BuildSubstrateLayers dThickness(), dEpsilon(), dMue(), dLossTang(), dLossCond()
		    StoreSubstrateData dThickness()

		    ' Since 2015 SP4 we can read some boundary conditions and bbox space values for FEM simulations in the emStateFile.xml
	  		Dim dSubstrateLateralExtension As Double
		    Dim dSubstrateVerticalExtension As Double
		    Dim sSubstrateBoundaryCond As String

		    Dim bBoundaryDataRead As Boolean
		    bBoundaryDataRead = False

		    If IsCSTVersionOfCurrentBlock_GreaterEqualThan(2015, 4) Then
		    	bBoundaryDataRead = ReadEMStateFile(sProject + "emStateFile.xml", dSubstrateLateralExtension, dSubstrateVerticalExtension, sSubstrateBoundaryCond)
		    End If

			' Now read and create the ports
			Dim bCreateWaveguidePorts As Boolean
			bCreateWaveguidePorts = True

			If (bBoundaryDataRead = True And dSubstrateLateralExtension > 0.0) Then
				bCreateWaveguidePorts = False
			End If

			If (DoesFileExist(sProject + "proj.ltd")) Then

				If IsCSTVersionOfCurrentBlock_GreaterEqualThan(2016, 2) Then

					ReadPinsAndPorts sProject + "proj.pin", sProject + "proj.prt", nPlanes, nPlaneId(), dPlaneCoords(), dPlaneThickness(), dXL, dXH, dYL, dYH, bCreateWaveguidePorts, bUseWaveguidePorts

		    	Else

			  		ReadNewPortFiles sProject + "proj.pin", dXL, dXH, dYL, dYH, nPlanes, nPlaneId(), dPlaneCoords(), dPlaneThickness(), bCreateWaveguidePorts

				End If
		    Else
		    	ReadPortFile sProject + "proj", dXL, dXH, dYL, dYH, nPlanes, nPlaneId(), dPlaneCoords(), dPlaneThickness()
		    End If

			Dim dAddSpace As Double
			dAddSpace = Sqr((dXH-dXL)*(dXH-dXL) + (dYH-dYL)*(dYH-dYL))

			' If the boundary outline is defined in the proj.ltd file then we do not add space to the bbox but consider the dimensions of the boundary outline.
			If (bBoundaryOutlineDefined) Then
				dAddSpace = 0.0
			End If

			Dim dSubstrateHeight As Double
		    dSubstrateHeight = GetSubstrateHeight(dThickness())

		    If (dSubstrateHeight * 5.0 < dAddSpace) Then
	  			dAddSpace = dSubstrateHeight * 5.0
	  		End If

	  		Dim dAddZmin As Double, dAddZmax As Double
	  		Dim sZminBound As String, sZmaxBound As String

	  		dAddZmin = dAddSpace
	  		dAddZmax = dAddSpace

	  		sZminBound = "open"
	  		sZmaxBound = "open"

	  		If (sBottomPlane <> "open") Then
	  		  dAddZmin = 0.0
		      sZminBound = "electric"
	  		End If

	  		If (sTopPlane <> "open") Then
	  		  dAddZmax = 0.0
		      sZmaxBound = "electric"
	  		End If

		    If (bBoundaryDataRead = True) Then

		    	Dim dVerticalExt_Zmin As Double, dVerticalExt_Zmax As Double
		    	dVerticalExt_Zmin = dSubstrateVerticalExtension
		    	dVerticalExt_Zmax = dSubstrateVerticalExtension

		    	If IsCSTVersionOfCurrentBlock_GreaterEqualThan(2016, 2) Then
					If sZminBound = "electric" Then
						dVerticalExt_Zmin = 0.0
					End If
					If sZmaxBound = "electric" Then
						dVerticalExt_Zmax = 0.0
					End If
					If sSubstrateBoundaryCond = "electric" Then
						dSubstrateLateralExtension = 0.0
					End If
				End If

		    	With Background
			     .Type "Normal"
			     .Epsilon "1.0"
                    .Mu("1.0")
			     .XminSpace CVar(dSubstrateLateralExtension)
			     .XmaxSpace CVar(dSubstrateLateralExtension)
			     .YminSpace CVar(dSubstrateLateralExtension)
			     .YmaxSpace CVar(dSubstrateLateralExtension)
			     .ZminSpace CVar(dVerticalExt_Zmin)
			     .ZmaxSpace CVar(dVerticalExt_Zmax)
				End With

		    Else

				With Background
			     .Type "Normal"
			     .Epsilon "1.0"
                    .Mu("1.0")
			     .XminSpace "0.0"
			     .XmaxSpace "0.0"
			     .YminSpace "0.0"
			     .YmaxSpace "0.0"
			     .ZminSpace CVar(dAddZmin)
			     .ZmaxSpace CVar(dAddZmax)
				End With

		    End If

		    If (RestoreGlobalDataValue("Macros\ADS Import\Ports\XMin") = "0") Then dXL = dXL - dAddSpace
		    If (RestoreGlobalDataValue("Macros\ADS Import\Ports\XMax") = "0") Then dXH = dXH + dAddSpace
		    If (RestoreGlobalDataValue("Macros\ADS Import\Ports\YMin") = "0") Then dYL = dYL - dAddSpace
		    If (RestoreGlobalDataValue("Macros\ADS Import\Ports\YMax") = "0") Then dYH = dYH + dAddSpace

			BuildConductingSlotLayers nPlanes, nPlaneId(), dPlaneThickness(), dPlaneCoords(), sPlaneTypes(), dXL, dXH, dYL, dYH

		    StoreBBoxBounds dXL, dXH, dYL, dYH
		    BuildSubstrateBricks

			' set boundary conditions to electric
			If (bBoundaryDataRead = False) Then
				sSubstrateBoundaryCond = "electric"
			End If

			With Boundary
			     .Xmin sSubstrateBoundaryCond
			     .Xmax sSubstrateBoundaryCond
			     .Ymin sSubstrateBoundaryCond
			     .Ymax sSubstrateBoundaryCond
			     .Zmin sZminBound
			     .Zmax sZmaxBound
			     .Xsymmetry "none"
			     .Ysymmetry "none"
			     .Zsymmetry "none"
			End With

			' optimize mesh settings for planar structures

			With Mesh
			     .MergeThinPECLayerFixpoints "True"
			     .RatioLimit "50"
			     .LinesPerWavelength "20"
			     .AutomeshRefineAtPecLines "True", "4"
			End With

			With Solver
			     .CalculationType "TD-S"
			     .StimulationPort "All"
			     .StimulationMode "All"
			     .SteadyStateLimit "-50"
			     .MeshAdaption "False"
			     .AutoNormImpedance "True"
			     .NormingImpedance "50"
			     .CalculateModesOnly "False"
			     .SParaSymmetry "False"
			     .StimulationType "Gaussian"
			     .StoreTDResultsInCache "False"
			     .FullDeembedding "True"
			     .SetSamplesFullDeembedding "5"
			End With

			ReadStimulationFile sProject + "proj.sti"

		Else
			ReportWarningToWindow "The import data is not complete or does not contain any geometric entities."

		End If

		ResultTree.EnableTreeUpdate True
		LockTree False
		SetNoMouseSelection False
		ScreenUpdating True
		SetLock False

		SelectTreeItem "Components"

		'Plot.ZoomToStructure

		If (bPolygonFailed) Then
			ReportWarningToWindow "Not all polygons could be read correctly. The invalid polygons have been stored in the ""failed polygons"" curve."
		End If

		If (bShowOhmicSheetWarning) Then
			ReportWarningToWindow "In ADS/Momentum some parts have been defined as ""impedance"" materials (ohms / square). " + vbCrLf + _
			       "These parts have been imported as material type ""normal""." + vbCrLf + vbCrLf + _
				   "Note: For good conductors the type ""lossy metal"" may improve the simulation efficiency in the transient solver (!T)." + vbCrLf + _
				   "In the tetrahedral frequency domain (!F, general purpose) ""ohmic sheet"" materials may be chosen."
		End If

	End If
End Sub
