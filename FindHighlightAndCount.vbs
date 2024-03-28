Sub FindHighlightAndCount()
    Dim searchString As String
    Dim rng As Range
    Dim count As Integer
    
    ' Prompt the user for input using an input box
    searchString = InputBox("Enter the search string:", "Search String Input")
    
    ' Check if the user input is not empty
    If searchString <> "" Then
        ' Initialize counter
        count = 0
        
        ' Set range to search the whole document
        Set rng = ActiveDocument.Content
        
        With rng.Find
            .ClearFormatting
            .Text = searchString
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            Do While .Execute
                ' Increment occurrence count
                count = count + 1
                
                ' Highlight each occurrence
                rng.HighlightColorIndex = wdYellow
                
                ' Concatenate the occurrence count with the search string
                rng.Text = searchString & "(" & count & ")"
                
                ' Move range to the end of the replaced string
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
        
        ' Display total occurrence count in a message box
        MsgBox "Total occurrences found: " & count, vbInformation, "Occurrences Count"
    Else
        ' If user input is empty, display a message
        MsgBox "No search string entered. Please try again.", vbExclamation
    End If
End Sub

