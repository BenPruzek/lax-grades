Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

' ================================================================================================
' Copyright 2011-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'--------------------------------------------------------------------------------------------
' 07-Jun-2023 set: Add possibility to compute a partial field contribution from a solid selection
' 09-May-2019 fsr: Activated ScriptSettings (needed for PushBrowse2D3DResults)
' 28-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 14-Mar-2012 ube: added rectangular option
' 13-Mar-2012 ube: added more dialogue elements
' 15-Sep-2011 apr,fsr,ube: first version
'--------------------------------------------------------------------------------------------

Public Const sNameDefaultText = "Please enter result name or browse for it."
Dim bFieldStatistics As Boolean

Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, z1 As Double, z2 As Double
Dim nx As Long, ny As Long, nz As Long
Dim dstepsize As Double

Dim sFileREX As String
Dim sFileSCT As String
Dim sFileSLM As String
Dim sTree2D As String
Dim sTree3D As String

Dim aSolidArray_CST() As String, nSolids_CST As Integer

Sub Main ()

	' Activate the StoreScriptSetting / GetScriptSetting functionality. Clear the data in order to
	' provide well defined environment for testing.

	ActivateScriptSettings True
	ClearScriptSettings

	nSolids_CST = 0

	Begin Dialog UserDialog 500,539,"Plot I-Solver Result on 2D Plane",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 20,7,470,168,"Cutting Plane",.GroupBox1
		OptionGroup .GroupCutplane
			OptionButton 40,31,240,14,"Use existing Cutting Plane",.OptionButton1
			OptionButton 40,56,120,14,"Manual Cut at",.OptionButton2
		OptionGroup .GroupXYZ
			OptionButton 160,56,35,14,"x",.OptionButton3
			OptionButton 210,56,35,14,"y",.OptionButton4
			OptionButton 260,56,35,14,"z",.OptionButton5
		TextBox 310,52,160,21,.xyz
		' PushButton 230,294,90,21,"Help",.PushHelp
		CheckBox 70,80,160,14,"Limit rectangular area",.CheckRectangle
		CheckBox 250,80,160,14,"Robust cross section",.Robust
		TextBox 130,112,90,21,.xmin
		TextBox 130,140,90,21,.xmax
		TextBox 250,112,90,21,.ymin
		TextBox 250,140,90,21,.ymax
		TextBox 370,112,90,21,.zmin
		TextBox 370,140,90,21,.zmax
		Text 80,114,40,14,"min",.Text2
		Text 80,144,40,14,"max",.Text3
		Text 165,94,20,14,"x",.Textx
		Text 290,94,20,14,"y",.Texty
		Text 410,94,20,14,"z",.Textz
		GroupBox 20,182,470,105,"Triangulation / Number of Points",.GroupBox4
		Text 250,210,120,14,"Max.Mesh Edge:",.Text6
		TextBox 370,203,90,21,.MaxMesh
		Text 250,236,120,14,"Min.Mesh Edge:",.Text7
		TextBox 370,231,90,21,.MinMesh
		Text 40,209,70,14,"Stepsize:",.Text8
		TextBox 130,203,90,21,.Stepsize
		Text 40,235,70,14,"NPoints ~:",.Text9
		TextBox 130,231,90,21,.NPoints
		PushButton 40,259,150,21,"Estimate",.Estimate
		'PushButton 210,259,150,21,"Preview Triangles",.Preview
		GroupBox 20,294,470,77,"Field Result",.GroupBox2
		TextBox 40,315,430,21,.ResultTreeName
		PushButton 40,343,150,21,"Browse Results...",.PushBrowseResults
		GroupBox 20,378,470,126,"Specials",.GroupBox3
		CheckBox 50,399,400,14,"Additionally calculate Statistics using Field Threshold",.CheckThreshold
		Text 80,424,140,14,"Field threshold value:",.TextThreshold
		TextBox 250,420,210,21,.dThreshold
		CheckBox 50,452,270,14,"Compute partial contribution from solids:",.ComputePartialContrib
		PushButton 330,449,130,21,"Select Solids...",.SelectSolids
		CheckBox 50,480,240,14,"Userdefined Treename Extension:",.CheckUserName
		TextBox 300,476,160,21,.UserName
		PushButton 40,511,150,21,"Calculate Plot",.Calculate
		CancelButton 200,511,90,21
	End Dialog
	Dim dlg As UserDialog
	dlg.Robust = True
	dlg.ResultTreeName = RestoreGlobalDataValue("PlotOnCutPlane\ResultTreeName")
	If dlg.ResultTreeName = "" Then
		dlg.ResultTreeName = sNameDefaultText
	End If

	dlg.Stepsize = Cstr(Mesh.GetMaximumEdgeLength())

	dlg.MaxMesh = Cstr(Mesh.GetMaximumEdgeLength())
	dlg.MinMesh = Cstr(Mesh.GetMinimumEdgeLength())

	Boundary.GetCalculationBox x1,x2,y1,y2,z1,z2
	dlg.xmin = Cstr(x1)
	dlg.xmax = Cstr(x2)
	dlg.ymin = Cstr(y1)
	dlg.ymax = Cstr(y2)
	dlg.zmin = Cstr(z1)
	dlg.zmax = Cstr(z2)

	dlg.xyz = "0.0"
	dlg.dThreshold = "0.0"

	If (Dialog(dlg)=0) Then Exit All ' cancel button was pressed

	StoreGlobalDataValue ("PlotOnCutPlane\ResultTreeName", dlg.ResultTreeName)

	Dim dThreshold As Double
	If dlg.CheckThreshold = 1 Then
		bFieldStatistics = True
		dThreshold = Evaluate(dlg.dThreshold)
	Else
		bFieldStatistics = False
		dThreshold = 0.0
	End If

'	ReportInformationToWindow("CalculateCutplane " + CStr(dThreshold))

	CalculateCutplane(dThreshold, dlg.ComputePartialContrib = 1)

	ActivateScriptSettings False

End Sub


Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean
	If (Action%=1 Or Action%=2) Then
			' Action%=1: The dialog box is initialized
			' Action%=2: The user changes a value or presses a button

		Dim bManualCut As Boolean, bRectangle As Boolean, bRobustCrossSection As Boolean, bComputePartialContrib As Boolean, iXYZ As Integer

		bManualCut = (DlgValue("GroupCutplane")=1)
		bRectangle = (DlgValue("CheckRectangle")=1)
		bRobustCrossSection = (DlgValue("Robust")=1) 
		bComputePartialContrib = (DlgValue("ComputePartialContrib")=1)
		If Not bManualCut Then bRectangle = False ' this prevents strange things further down, if current cutplane is used

		DlgEnable "CheckRectangle", bManualCut

		DlgEnable "GroupXYZ", bManualCut
		DlgEnable "xyz", bManualCut
		DlgEnable "xmin", bManualCut And bRectangle
		DlgEnable "xmax", bManualCut And bRectangle
		DlgEnable "ymin", bManualCut And bRectangle
		DlgEnable "ymax", bManualCut And bRectangle
		DlgEnable "zmin", bManualCut And bRectangle
		DlgEnable "zmax", bManualCut And bRectangle
		DlgEnable "Textx", bManualCut
		DlgEnable "Texty", bManualCut
		DlgEnable "Textz", bManualCut

		DlgEnable "MaxMesh", False
		DlgEnable "MinMesh", False
		DlgEnable "NPoints", False

		DlgEnable "TextThreshold", DlgValue("CheckThreshold")=1
		DlgEnable "dThreshold", DlgValue("CheckThreshold")=1

		DlgEnable "SelectSolids", bComputePartialContrib

		DlgEnable "UserName", DlgValue("CheckUserName")=1

		Dim dCoordinate As Double
		If bManualCut Then
			' manual cutplane
			dCoordinate = Evaluate(DlgText("xyz"))
			With Plot
				.ShowCutplane False
				Select Case DlgValue("GroupXYZ")
				Case 0
					.CutPlaneNormal "x"
					.DefinePlane 1,0,0,dCoordinate,0,0
					DlgEnable "xmin", False
					DlgEnable "xmax", False
					DlgEnable "Textx", False
				Case 1
					.CutPlaneNormal "y"
					.DefinePlane 0,1,0,0,dCoordinate,0
					DlgEnable "ymin", False
					DlgEnable "ymax", False
					DlgEnable "Texty", False
				Case 2
					.CutPlaneNormal "z"
					.DefinePlane 0,0,1,0,0,dCoordinate
					DlgEnable "zmin", False
					DlgEnable "zmax", False
					DlgEnable "Textz", False
				End Select
				.ShowCutplane True
				.Update
			End With
		End If

		Select Case DlgItem
		Case "PushBrowseResults"
			DialogFunction = True       ' Don't close the dialog box.

			Dim s2exclude As String, newname As String
			s2exclude = "Surface Current\"+cExcludeSeperator+"\Port Modes"

			newname = ""
			PushBrowse2D3DResults(newname,s2exclude)
			If newname <> "" Then
				DlgText("ResultTreeName"),newname
			End If
		Case "SelectSolids"
			DialogFunction = True       ' Don't close the dialog box.
			SelectSolids aSolidArray_CST(), nSolids_CST
		Case "Estimate", "Preview", "Calculate"

			DialogFunction = True       ' Don't close the dialog box.

			Boundary.GetCalculationBox x1,x2,y1,y2,z1,z2

			If bRectangle Then
				x1 = Evaluate(DlgText("xmin"))
				x2 = Evaluate(DlgText("xmax"))
				y1 = Evaluate(DlgText("ymin"))
				y2 = Evaluate(DlgText("ymax"))
				z1 = Evaluate(DlgText("zmin"))
				z2 = Evaluate(DlgText("zmax"))
			End If

			dstepsize = Evaluate(DlgText("Stepsize"))

			If dstepsize = 0 Then dstepsize = 1

			nx = 1+CLng((x2-x1)/dstepsize)
			ny = 1+CLng((y2-y1)/dstepsize)
			nz = 1+CLng((z2-z1)/dstepsize)
			If nx<2 Then nx = 2
			If ny<2 Then ny = 2
			If nz<2 Then nz = 2

			Dim idir As Integer
			idir = DlgValue("GroupXYZ")

			If bManualCut Then
				Select Case idir
				Case 0
					DlgText("NPoints"),Cstr(ny*nz)
				Case 1
					DlgText("NPoints"),Cstr(nx*nz)
				Case 2
					DlgText("NPoints"),Cstr(nx*ny)
				End Select
			Else
				Dim nWorstCase As Long
				nWorstCase = nx*ny
				If nWorstCase < ny*nz Then nWorstCase = ny*nz
				If nWorstCase < nx*nz Then nWorstCase = nx*nz
				DlgText("NPoints"),Cstr(nWorstCase)
			End If

			If DlgItem = "Preview" Or DlgItem = "Calculate" Then
			    If (DlgText("ResultTreeName") = "" Or DlgText("ResultTreeName") = sNameDefaultText) Then
					MsgBox "Please select valid field result first.", vbInformation
					DialogFunction = True ' There is an error in the settings -> Don't close the dialog box.
				Else
					If DlgItem = "Calculate" And bComputePartialContrib And nSolids_CST = 0 Then
						MsgBox "Please select a non-empty set of solids or turn off partial contribution calculation.", vbInformation
						DialogFunction = True ' There is an error in the settings -> Don't close the dialog box.
					Else
						Dim sfile3d As String
						sTree3D = "2D/3D Results\"+DlgText("ResultTreeName")
						sfile3d = Resulttree.GetFileFromTreeItem(sTree3D)
						RemoveLastChars(sfile3d,4)  ' remove extension ".mie"

						Dim i As Integer
						i = 0
						Do
							i = i+1
							sFileREX = sfile3d + "_2D-"+Cstr(i)+".rex"
						Loop Until Dir$(sFileREX,vbNormal) = ""   ' loop until new file is found, do never overwrite old results

						sFileSLM = sfile3d + "_2D-"+CStr(i)+"_rex.slm"
						sFileSCT = sfile3d + "_2D-"+CStr(i)+"_rex.sct"

						If DlgValue("CheckUserName")=1 Then
							sTree2D = sTree3D + "_2D-"+DlgText("UserName")
						Else
							sTree2D = sTree3D + "_2D-"+CStr(i)
						End If

						Prepare_SLIM_Mesh(bRectangle,idir,dCoordinate,dstepsize,bRobustCrossSection) ' box coordinates x1,x2,y1, and nx ny nz come through global variables

						If DlgItem = "Calculate" Then
							DialogFunction = False ' dialog will be closed now (otherwise GUI might be blocked...)
						'Else
						'	Plot.ShowCutplane True
						'	Plot.DrawMesh sFileSLM
						End If
					End If
				End If
			End If
		End Select

	End If
End Function


Function Prepare_SLIM_Mesh(bRectangle As Boolean, idir As Integer, dCoordinate As Double, dmesh As Double, dRobustCrossSection As Boolean ) As Double


	Dim ip As Long
	Dim nTriag As Long, iTri As Long
'	Dim myTriangleData As TriangleData
	Dim x As Double, y As Double, z As Double

	If bRectangle Then
		Dim oMesh As Object
		Set oMesh = Result3D("")
		oMesh.InitMesh
		
		' here special rectangular meshing is used
		Dim nu As Long, nv As Long, du As Double, dv As Double, iu As Long, iv As Long
		Select Case idir
		Case 0 ' normal x
			nu = ny
			nv = nz
			du = (y2-y1)/(ny-1)
			dv = (z2-z1)/(nz-1)
			For iv = 0 To nv-1
				For iu = 0 To nu-1
					oMesh.AddNode (dCoordinate , y1 + iu*du , z1 + iv*dv , 10)
				Next iu
			Next iv
		Case 1 ' normal y
			nu = nz
			nv = nx
			du = (z2-z1)/(nz-1)
			dv = (x2-x1)/(nx-1)
			For iv = 0 To nv-1
				For iu = 0 To nu-1
					oMesh.AddNode (x1 + iv*dv , dCoordinate , z1 + iu*du , 10)
				Next iu
			Next iv
		Case 2 ' normal z
			nu = nx
			nv = ny
			du = (x2-x1)/(nx-1)
			dv = (y2-y1)/(ny-1)
			For iv = 0 To nv-1
				For iu = 0 To nu-1
					oMesh.AddNode (x1 + iu*du , y1 + iv*dv , dCoordinate , 10)
				Next iu
			Next iv
		End Select
		For iv = 0 To nv-2
			For iu = 0 To nu-2
				ip = iu + iv*nu
				oMesh.AddTriangle (ip       , ip + 1  , ip + nu , 11)
				oMesh.AddTriangle (ip +1+nu , ip + nu , ip + 1  , 11)
			Next iu
		Next iv
		oMesh.Save(sFileSLM)
	Else
		ScalarPlot2D.RobustCrossSection( dRobustCrossSection )
		' now standard cutplane is triangulated
		nTriag = Plot.TriangulateCutplane(dmesh, sFileSLM)
	End If

End Function


Function CalculateCutplane(dThreshold As Double, bComputePartialContrib As Boolean) As Double
	SelectTreeItem(sTree3D)
	Dim oMesh As Object
	Set oMesh = Result3D(sFileSLM)
	Dim x As Double, y As Double, z As Double
	Plot.ShowCutplane True
	Dim ip As Long, nTriag As Long, nMat As Long, iTri As Long, dmesh As Double
	Dim nNodes As Long, nEdges As Long, nTriangles As Long, nTets As Long
	oMesh.GetMeshInfo(nNodes, nEdges, nTriangles, nTets)
'	ReportInformationToWindow("Cutplane mesh: " + CStr(nNodes) + " nodes, " +CStr(nTriangles)+" triangles")

	If nTriangles<>0 Then

		Dim oField As Object
		Set oField = Result3D("")
		oField.InitSCT(nTriangles, 3, 6, 0) ' nTrianglesFirstSide, nSamplesPerElement, nComponents, nTrianglesSecondSide

		VectorPlot3D.Reset
		Dim uid As Long, iTriangle As Long, iNode As Long

		If bComputePartialContrib Then
			VectorPlot3D.SetISolverSolidNames aSolidArray_CST
		End If

		' here over all nodes is looped to prepare the point list, to be calculated by the solver via "VectorPlot3D.CalculateList"
		For iNode = 0 To nNodes-1
			oMesh.GetNode(iNode, x, y, z, uid)
			VectorPlot3D.AddListItem( x,  y,  z)
		Next iNode

		VectorPlot3D.CalculateList

		Dim FieldOver As Double, FieldUnder As Double, myTriangleData As TriangleData, TotalArea As Double
		Dim FieldMax As Double
		FieldMax = 0
		TotalArea = 0
		FieldOver = 0
		FieldUnder = 0
		For iTriangle = 0 To nTriangles-1
	        Dim n(3) As Long
			oMesh.GetTriangle(iTriangle, n(0), n(1), n(2), uid)
			For iNode = 0 To 2
'				oMesh.GetNode(n(iNode), x, y, z, uid)
				Dim xr As Double, yr As Double, zr As Double, xi As Double, yi As Double, zi As Double
				VectorPlot3D.GetListItem(n(iNode), xr, yr, zr, xi, yi, zi)
				oField.Setxre(iTriangle*3+iNode, xr)
				oField.Setyre(iTriangle*3+iNode, yr)
				oField.Setzre(iTriangle*3+iNode, zr)
				oField.Setxim(iTriangle*3+iNode, xi)
				oField.Setyim(iTriangle*3+iNode, yi)
				oField.Setzim(iTriangle*3+iNode, zi)
				If (Sqr(xr^2+xi^2+yr^2+yi^2+zr^2+zi^2)>FieldMax) Then
					FieldMax = Sqr(xr^2+xi^2+yr^2+yi^2+zr^2+zi^2)
					'ReportInformationToWindow(FieldMax)
				End If
			Next iNode
			myTriangleData = CalculateTriangleData(oMesh, oField, iTriangle)
			If (myTriangleData.FieldMag > dThreshold) Then
				FieldOver = FieldOver + myTriangleData.Area
			Else
				FieldUnder = FieldUnder + myTriangleData.Area
			End If
			TotalArea = TotalArea + myTriangleData.Area
		Next iTriangle
		oField.Save(sFileSCT)
		VectorPlot3D.SaveMetadata( sFileREX, sFileSLM, sFileSCT, "plane" ) 'last parameter is true if field are plotted On plane And Not On srf
		If bFieldStatistics Then
			ReportInformationToWindow("Max. field value: " + CStr(FieldMax))
			ReportInformationToWindow("Above threshold (" + Cstr(dThreshold) + "): " + Cstr(FieldOver/TotalArea*100) + "%")
			ReportInformationToWindow("Below threshold (" + Cstr(dThreshold) + "): " + Cstr(FieldUnder/TotalArea*100) + "%")
		End If

		With Resulttree
			.Name sTree2D
			.Type "EField3D"
			.File sFileREX
			.Add
		End With

		SelectTreeItem (sTree2D)

	End If
End Function


Type TriangleData
	Area As Double
	CentroidX As Double
	CentroidY As Double
	CentroidZ As Double
	FieldMag As Double
End Type


Function CalculateTriangleData(oMesh As Object, oField As Object, iIndex As Long) As TriangleData
	Dim myTriangleData As TriangleData
	Dim uid As Long, iNodeCount As Long
	Dim iNode(2) As Long
	Dim dX(2) As Double, dY(2) As Double, dZ(2) As Double
	Dim dSP As Double, dLa As Double, dLb As Double, dLc As Double
	Dim dCentroidDist(2) As Double

	' Get node numbers for given triangle
	oMesh.GetTriangle(iIndex, iNode(0), iNode(1), iNode(2), uid)
	' Get xyz values for each node
	For iNodeCount = 0 To 2
		oMesh.GetNode(iNode(iNodeCount), dX(iNodeCount), dY(iNodeCount), dZ(iNodeCount), uid)
	Next iNodeCount
	' Calculate edge lengths
	dLa = Sqr((dX(2)-dX(1))^2+(dY(2)-dY(1))^2+(dZ(2)-dZ(1))^2)
	dLb = Sqr((dX(0)-dX(2))^2+(dY(0)-dY(2))^2+(dZ(0)-dZ(2))^2)
	dLc = Sqr((dX(0)-dX(1))^2+(dY(0)-dY(1))^2+(dZ(0)-dZ(1))^2)

	dSP = (dLa+dLb+dLc)/2 ' semiperimeter
	' Heron's formula
	myTriangleData.Area = Sqr(dSP*(dSP-dLa)*(dSP-dLb)*(dSP-dLc))

	' Calculate centroid coordinates of triangle
	myTriangleData.CentroidX = (dX(0) + dX(1) + dX(2))/3
	myTriangleData.CentroidY = (dY(0) + dY(1) + dY(2))/3
	myTriangleData.CentroidZ = (dZ(0) + dZ(1) + dZ(2))/3

	' Interpolate field value at centroid
	myTriangleData.FieldMag = 0
	For iNodeCount = 0 To 2
		dCentroidDist(iNodeCount) = Sqr((dX(iNodeCount)-myTriangleData.CentroidX)^2+(dY(iNodeCount)-myTriangleData.CentroidY)^2+(dZ(iNodeCount)-myTriangleData.CentroidZ)^2)
		myTriangleData.FieldMag = myTriangleData.FieldMag + Sqr(oField.GetXre(iIndex*3+iNodeCount)^2+oField.GetXim(iIndex*3+iNodeCount)^2 + _
																oField.GetYre(iIndex*3+iNodeCount)^2+oField.GetYim(iIndex*3+iNodeCount)^2 + _
																oField.GetZre(iIndex*3+iNodeCount)^2+oField.GetZim(iIndex*3+iNodeCount)^2)/dCentroidDist(iNodeCount)^2
	Next iNodeCount
	myTriangleData.FieldMag = myTriangleData.FieldMag/(1/dCentroidDist(0)^2+1/dCentroidDist(1)^2+1/dCentroidDist(2)^2)

	CalculateTriangleData = myTriangleData
End Function
