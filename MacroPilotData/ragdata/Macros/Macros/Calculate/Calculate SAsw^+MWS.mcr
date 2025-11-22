' Calculate Surface Area in square wavelength (SAsw) for surface mesh

' ================================================================================================
' Copyright 2019-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 10-Dec-2019 ube: First version
' ================================================================================================
Sub Main ()

	FMax = Solver.GetFMax
	FMax1 = FMax * Units.GetFrequencyUnitToSI
	lambdaMax = CStr(Clight/FMax1)

	Area = Mesh.GetSurfaceMeshArea         'in user units
	factor = Units.GetGeometryUnitToSI
	' Returns the factor To convert a geometry value measured In Units of the current project into the Units
	' -> a(SI Unit) = factor * b(project Unit)

	AreaPerLambdaSquare = 0
	If Area > 0 Then
		dLambdainUserUnits = Clight/FMax1/factor
		AreaPerLambdaSquare = Area / ( dLambdainUserUnits * dLambdainUserUnits)
		End If

	Message1 = "FMax = " + CStr(FMax) + " " + Units.GetUnit("Frequency") + vbCrLF
	Message2 = "Wavelength = " + lambdaMax + " m" + vbCrLf + vbCrLf
	Message3 = "Surface area / square wavelength      " + vbCrLf
	Message4 = "SAsw = " + CStr(Format(AreaPerLambdaSquare,"Standard"))
	Message = Message1 + Message2 + Message3 + Message4
	MsgBox(Message)
End Sub
