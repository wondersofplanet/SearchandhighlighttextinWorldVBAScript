Sub FindHighlightAndCount()
    Dim searchString As String
    Dim rng As Range
    Dim count As Integer
    Dim highlightColor As Long ' Color for highlighting
    Dim randomColor As Long ' Randomly generated color
    
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
                
                ' Generate a random color with brighter shades
                randomColor = RGB(Int((255 - 150) * Rnd()) + 150, Int((255 - 150) * Rnd()) + 150, Int((255 - 150) * Rnd()) + 150)
                
                ' If it's the first occurrence in this search, assign the random color
                If count = 1 Then
                    highlightColor = randomColor
                End If
                
                ' Highlight each occurrence with the same random color
                rng.Shading.BackgroundPatternColor = highlightColor
                
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

