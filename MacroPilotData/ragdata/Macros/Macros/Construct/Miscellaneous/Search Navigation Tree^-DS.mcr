'#Language "WWB-COM"

' Macro to search in navigation tree. Useful if model is very complex and shapes/groups/results are hard to find.

' ================================================================================================
' Copyright 2014-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------------
' 19-Aug-2014 vgu,msc: first version
'------------------------------------------------------------------------------------------

Option Explicit
Const initialArraySize% = 300     ' Starting size of dynamic array which stores search results
Dim maxSize As Integer            ' Dynamic size of array storing the search results currently allocated
Dim CaseCheck As Boolean          ' True if Case Checking enabled durign search
Dim findString As String          ' Input string to be searched for

Sub Main
Dim dummy(0) As String            ' Initialising ListBox Dialog item to empty string
maxSize = initialArraySize

' Code for creating the User Dialog interface

	Begin Dialog UserDialog 580,567,"Find In Navigation Tree",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 0,0,580,133,"",.GroupBox1
		Text 10,21,190,21,"Search Navigation Tree for :",.Text2
		TextBox 200,21,360,21,.StringFindTextBox
		CheckBox 40,63,180,21,"Components",.ComponentsCheckBox
		GroupBox 0,133,580,434,"",.GroupBox2
		PushButton 40,105,130,21,"Search",.SearchButton
		ListBox 20,147,550,392,dummy(),.ListBox1,1
		Text 40,49,180,14,"Restrict Search to :",.Text1
		CheckBox 40,84,80,14,"Groups",.GroupsCheckBox
		CheckBox 180,105,130,14,"Match Case",.MatchCaseCheckBox
		CancelButton 390,105,130,21, .CancelButtonPushed
		Text 20,546,220,14,"Number of search results found",.CountResults
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub

' This function handles all the clicks made in the user dialog box

Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean

Select Case Action%

    	Case 1 ' Initialisation
            Dim parentString As String
            Dim compareParentString As String              ' Modified parent string for finding substrings
            Dim foundStrings() As String                   ' Array containing search results
            Dim count As Integer                           ' Size of array
            Dim ComponentsCheck As Boolean                 ' Holds value of .ComponentsCheckBox
            Dim GroupsCheck As Boolean                     ' Holds value of .GroupsCheckBox
            Dim selectedString As String                   ' String selected by user from the search results

    	Case 2 ' Value changing or button pressed

    		Select Case DlgItem

    			Case "SearchButton"

    				findString = DlgText("StringFindTextBox")
                    ComponentsCheck = DlgValue("ComponentsCheckBox")
                    GroupsCheck = DlgValue("GroupsCheckBox")
                    CaseCheck = DlgValue("MatchCaseCheckBox")
                    count = 0                                          ' Initialise count to 0
                    ReDim foundStrings(maxSize)
                    DlgListBoxArray "ListBox1", foundStrings
                    DlgText "CountResults","Number of search results found"

                    If ( Not(CaseCheck)) Then
                    	findString = LCase(findString)
                    End If

            If ( ComponentsCheck Or GroupsCheck ) Then          ' If either check box selected search only in those folders

                    If ( ComponentsCheck ) Then
                       	parentString = "Components"

                        compareParentString = caseChange(parentString)

                        If ( InStr( compareParentString , findString ) ) Then
                            addString( parentString, foundStrings, count)
                        End If
                        findSubString( parentString, foundStrings, count)
                    End If


                    If ( GroupsCheck ) Then
                       	parentString = "Groups"

                        compareParentString = caseChange(parentString)

                        If ( InStr( compareParentString , findString ) ) Then
                            addString( parentString, foundStrings, count)
                        End If
                        findSubString( parentString, foundStrings, count )
                     End If

             Else
                    	parentString = "Components"
                        While ( Not (parentString = "") )

                           compareParentString = caseChange(parentString)

                           If ( InStr( compareParentString , findString ) ) Then
                         	addString( parentString, foundStrings, count)
                           End If
                           findSubString( parentString, foundStrings, count )
                           parentString = Resulttree.GetNextItemName(parentString)
                        Wend
              End If

                    DlgListBoxArray "ListBox1", foundStrings                          ' Display search results in dialog box
                    DlgText "CountResults", Cstr( count ) & " search results found"   ' Display number of results found
                    DialogFunc = True

             Case "ListBox1"

                   selectedString = DlgText("ListBox1")
                   SelectTreeItem(selectedString)              ' Select item in navigation tree

             Case "CancelButtonPushed"
                  Exit All
             End Select

       End Select

End Function

' This is a recursive function to find all the strings in the navigation tree containing the findString

 Function findSubString ( parentString$ , foundStrings$(), count% )

   Dim childString As String
   Dim compareChildString As String
   childString = Resulttree.GetFirstChildName(parentString)

   While( Not (childString = "" ) )

   compareChildString = caseChange(childString)

   If ( InStr( Len(parentString)+1 , compareChildString , findString ) ) Then                                  ' Execute if the substring present
   addString( childString, foundStrings, count)
   End If

   findSubString ( childString, foundStrings, count)
   childString = Resulttree.GetNextItemName(childString)

   Wend

 End Function

' This function adds a new string to the search results

 Function addString( newString$, foundStrings$(), count%)

 	 If ( count > maxSize) Then
    	maxSize = 3*maxSize
    	ReDim Preserve foundStrings(maxSize)
     End If
        foundStrings(count) = newString
        count = count + 1

 End Function

' This function returns the input string in lowercase if CaseCheck button is not checked

 Function caseChange(ByVal newString$) As String

  If(CaseCheck) Then
  caseChange = newString
  Else
  caseChange = LCase(newString)
  End If

 End Function


