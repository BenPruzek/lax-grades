' This macro exports the impedance matrix calculated with the LF-FD-solver 
' into a Z-parameter Touchstone file.
'
' ================================================================================================
'
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
'================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 02-Feb-2017 ckr: initial version
' ================================================================================================

Sub Main ()

	Dim dataThere As Boolean
	Dim treeItem As String
	Dim fileName As String
	Dim baseName As String
	Dim treeFolder As String
	Dim zParameter As Object
	Dim matrixSize As Long
	Dim nFreqs As Long, nPorts As Long
	Dim i As Long, j As Long, k As Long, l As Long
	Dim data() As Double
	Dim fso As Object
	Dim strPath As String
	Dim strRow As String
	Dim strRow2 As String
	Dim boolFlag As Boolean
	Dim strArray() As String
	Dim a As Long, b As Long

	baseName = "LF-FD-ImpedanceMatrix"

	treeFolder = "1D Results\LF Solver (MQS)\Impedance"

	dataThere = SelectTreeItem( treeFolder )

	If dataThere = True Then

		Set fso = CreateObject("Scripting.FileSystemObject")

		treeItem = Resulttree.GetFirstChildName( treeFolder )
		' get nPorts
		matrixSize = 0
		While treeItem <> ""
			matrixSize = matrixSize + 1
			fileName = Resulttree.GetFileFromTreeItem(treeItem)
			Set zParameter = Result1DCOmplex( fileName )
			nFreqs = zParameter.GetN
			treeItem = Resulttree.GetNextItemName(treeItem)
		Wend
		nPorts = CInt( Sqr(matrixSize) )

		' prepare outputfile
		strPath = fso.BuildPath(GetProjectPath("Project"),"Export")
		If Not fso.FolderExists(strPath) Then
		    fso.CreateFolder (strPath)
		End If

		strPath = fso.BuildPath(strPath,"TOUCHSTONE files")
		If Not fso.FolderExists(strPath) Then
		    fso.CreateFolder (strPath)
		End If

		strPath = fso.BuildPath( strPath, baseName + ".z" + CStr( nPorts ) + "p" )
		Set oFile = fso.CreateTextFile(strPath)

		' write header in file
		oFile.WriteLine "! " + CStr(nPorts) + "-port Z-parameter, " + CStr(nFreqs) +" frequency points"
		oFile.WriteLine "# " + Units.GetUnit("Frequency") + " Z RI R 50"
		oFile.WriteLine ""

		' alloc containers
		ReDim data(matrixSize - 1, nFreqs - 1, 2)
		ReDim strArray( matrixSize - 1 )

		' fill data container
		treeItem = Resulttree.GetFirstChildName( treeFolder )
		For i = 0 To matrixSize - 1
			fileName = Resulttree.GetFileFromTreeItem(treeItem)
			Set zParameter = Result1DCOmplex( fileName )
			For j = 0 To nFreqs - 1
				data(i,j,0) = zParameter.GetX(j)
				data(i,j,1) = zParameter.GetYRe(j)/50
				data(i,j,2) = zParameter.GetYIm(j)/50
			Next j
			treeItem = Resulttree.GetNextItemName(treeItem)
		Next i

		' calc layout parameter
		a = CInt( nPorts / 4 )
		b = nPorts - 4*a

		For i = 0 To nFreqs - 1
			For j = 0 To matrixSize - 1
				strArray(j) = " " + CStr( data(j,i,1) ) + " " + CStr( data(j,i,2) )
			Next j
			Select Case nPorts
				Case 1
					' Complete Matrix in one line
					strRow = CStr( data(0,i,0) ) + strArray(0)
					oFile.WriteLine strRow
				Case 2
					' Complete Matrix in one line
					strRow = CStr( data(0,i,0) ) + strArray(0) + strArray(1) + strArray(2) + strArray(3)
					oFile.WriteLine strRow
				Case 3
					' Each row in one line, 3 items per row
					strRow = CStr( data(0,i,0) ) + strArray(0) + strArray(1) + strArray(2)
					oFile.WriteLine strRow
					strRow = strArray(3) + strArray(4) + strArray(5)
					oFile.WriteLine strRow
					strRow = strArray(6) + strArray(7) + strArray(8)
					oFile.WriteLine strRow
				Case 4
					' Each row in one line, 4 items per row
					strRow = CStr( data(0,i,0) ) + strArray(0) + strArray(1) + strArray(2) + strArray(3)
					oFile.WriteLine strRow
					strRow = strArray(4) + strArray(5) + strArray(6) + strArray(7)
					oFile.WriteLine strRow
					strRow = strArray(8) + strArray(9) + strArray(10) + strArray(11)
					oFile.WriteLine strRow
					strRow = strArray(12) + strArray(13) + strArray(14) + strArray(15)
					oFile.WriteLine strRow
				Case Else
					strRow = CStr( data(0,i,0) ) + strArray(0) + strArray(1) + strArray(2) + strArray(3)
					oFile.WriteLine strRow
					If a > 1 Then
						For k = 1 To a
							strRow = strArray( k*4 ) + strArray( k*4 + 1 ) + strArray( k*4 + 2 ) + strArray( k*4 + 3 )
							oFile.WriteLine strRow
						Next k
					End If
					strRow = ""
					For k = 0 To b-1
						strRow = strRow + strArray( a*4 + k )
					Next k
					oFile.WriteLine strRow

					' loop over next columns
					For l = 1 To nPorts - 1
						strRow = strArray( l*nPorts + 0 ) + strArray( l*nPorts + 1 ) + strArray( l*nPorts + 2 ) + strArray( l*nPorts + 3 )
						oFile.WriteLine strRow
						If a > 1 Then
							For k = 1 To a
								strRow = strArray( l*nPorts + k*4 ) + strArray( l*nPorts + k*4 + 1 ) + strArray( l*nPorts + k*4 + 2 ) + strArray( l*nPorts + k*4 + 3 )
								oFile.WriteLine strRow
							Next k
						End If
						strRow = ""
						For k = 0 To b-1
							strRow = strRow + strArray( l*nPorts + a*4 + k )
						Next k
						oFile.WriteLine strRow
					Next l
			End Select
		Next i

		oFile.Close

		Set oFile = Nothing
		Set fso = Nothing

		ReportInformation ("Successfully written TOUCHSTONE file: " + strPath)
	Else
		ReportWarning ("No Impedance Matrix has been found!" + vbNewLine + _
			"Please switch On ""Calculate Impedance Matrix"" in the LF-Solver Setup and recalculate.")
	End If



End Sub
