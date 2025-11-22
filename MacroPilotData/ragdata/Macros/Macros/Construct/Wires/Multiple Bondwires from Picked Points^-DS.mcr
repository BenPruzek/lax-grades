' *Construct / Wires / Multiple Bondwires from Picked Points
' !!! Do not change the line above !!!

' ================================================================================================
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'---------------------------------------------------------------------------------------------
' 30-Dec-2013 ube: added little help text in dialogue
' 01-Mar-2006    : first version
'---------------------------------------------------------------------------------------------

Sub Main ()

 Dim n As Integer
 Dim nPicks As Integer
 Dim pos () As Double

 nPicks = Pick.GetNumberOfPickedPoints
 If nPicks = 0 Then
		MsgBox _
			"No Points are picked - aborting Macro", _
			vbOkOnly + vbCritical, _
			"Multiple Wire"
		Exit All
 End If

 ReDim pos(nPicks,3)

 Dim verify As Integer
 verify = ( nPicks / 2 ) * 2
 If verify <> nPicks Then
 		MsgBox _
 			"Picked points should be pairs - aborting Macro", _
 			vbOkOnly + vbCritical, _
			"Multiple Wire"
		Exit All
 End If


	Begin Dialog UserDialog 400,133,"Generating multiple Bondwires" ' %GRID:10,7,1,1
		Text 30,14,90,14,"Wire Height",.Text1
		TextBox 30,35,90,21,.Wireheight
		Text 200,14,90,14,"Radius",.Text2
		TextBox 200,35,90,21,.Wireradius
		OKButton 20,105,90,21
		CancelButton 120,105,90,21
		Text 20,70,370,28,"Note: The height of the wires is always in z (or w) direction. Please place WCS accordingly prior to using this macro.",.Text3
	End Dialog
 Dim dlg As UserDialog

 dlg.Wireheight = "0"
 dlg.Wireradius = "0"

 If (Dialog(dlg) = 0) Then Exit All

 Dim wierH, wireR As Double
 wireH = Val(dlg.Wireheight)
 wireR = Val(dlg.Wireradius)

 For n = 1 To nPicks STEP 2

    Pick.GetPickpointCoordinates (n, pos(n,0), pos(n,1), pos(n,2))
    Pick.GetPickpointCoordinates (n+1, pos(n+1,0), pos(n+1,1), pos(n+1,2))
 Next n

 Dim wireOffset As Integer
 Dim sContents As String

 Dim nStart As Long, sItemName As String
 nStart = 0

 sItemName = ResultTree.GetFirstChildName("Wires")

 While sItemName <> ""
    Dim sWireName As String
	sWireName = Mid(sItemName, 7)

	If Left(sWireName,6) = "MWIRE_" Then
		Dim nThisWire As Long
		nThisWire = CLng(Mid(sWireName,7))

		If (nThisWire > nStart) Then
		  nStart = nThisWire
		End If

	End If

	sItemName = ResultTree.GetNextItemName(sItemName)
 Wend

 For n = 1 To nPicks STEP 2

   nStart = nStart + 1

   sContents = ""

   wireName = "MWIRE_" + CStr(nStart)

   sContents = sContents + "With Wire" + vbLf
   sContents = sContents + "	.Reset" + vbLf
   sContents = sContents + "	.Name """ + wireName + """" + vbLf
   sContents = sContents + "	.Type ""BondWire""" + vbLf
   sContents = sContents + "	.Height """ + CStr(wireH) + """" + vbLf
   sContents = sContents + "	.Radius """ + CStr(wireR) + """" + vbLf
   sContents = sContents + "	.Point1 """ + CStr(pos(n,0)) + """,""" + CStr(pos(n,1)) + """,""" + CStr(pos(n,2)) + """,""False""" + vbLf
   sContents = sContents + "	.Point2 """ + CStr(pos(n+1,0)) + """,""" + CStr(pos(n+1,1)) + """,""" + CStr(pos(n+1,2)) + """,""False""" + vbLf
   sContents = sContents + "	.Add" + vbLf
   sContents = sContents + "End With" + vbLf
   sContents = sContents + vbLf

   AddToHistory "define bondwire: " + wireName, sContents

 Next n

End Sub

