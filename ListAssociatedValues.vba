Sub ListAssociatedValues()

'Column 1: unique ID associating some rows together
'Column 5: part 1 of desired text from associated rows
'Column 6: part 2 of desired text from associated rows
'Column 7: how many associated rows linked to this row (previously calculated by formula)
'Column 8: whether this is a relevant reference row or not (previously calculated by formula)
'Column 9: new column filled for reference rows, listing the desired texts from each associated rows, comma separated

Dim iRow As Long 'Outer loop finding reference rows
Dim iRow2 As Long 'Inner loop finding associated rows
Dim strInsert As String 'Text to add to column 9 of the reference rows
Dim iFound As Long 'How many of the associated rows have been found in the inner loop for the chosen reference row in the outer loop

'Outer loop of all rows of the sheet
For iRow = 2 To Rows.Count
    'Check if the row is a reference row, and if it has at least 1 associate row
    If (Cells(iRow, 8).Value = "Yes") And (Cells(iRow, 7).Value > 0) Then
        'Initialise the inner loop
        strInsert = ""
        iFound = 0
        'Inner loop of all rows of the sheet
        For iRow2 = 2 To Rows.Count
            'Skip if this the same reference row, we have found an associate row if the unique ID is the same
            If (iRow <> iRow2) And (Cells(iRow, 1).Value = Cells(iRow2, 1).Value) Then
                'Data is dirty, macro crashes when desired text is empty
                If Cells(iRow2, 5).Value <> "" Then
                    'Add the first part of desired text
                    strInsert = strInsert & Cells(iRow2, 5).Value
                    
                    'This data is also dirty, same problem
                    If Cells(iRow2, 6).Value <> "" Then
                        'Add the second part of desired text, with a space between
                        strInsert = strInsert & " " & Cells(iRow2, 6)
                    End If
                    
                    'Add nice comma separation
                    strInsert = strInsert & ", "
                End If
                'Increment how many of the associated rows you have found
                iFound = iFound + 1
                'If you have found all the associated rows for the reference row, you can stop
                If (iFound = Cells(iRow, 7).Value) Then Exit For
            End If
        Next iRow2
        'Check if any text was found
        If strInsert <> "" Then
            'If text was found, then there will be an unused comma separate at the end of the list, remove it
            strInsert = Left(strInsert, Len(strInsert) - 2)
            'Insert the text into the new column for the reference row
            Cells(iRow, 9).Value = strInsert
        End If
    End If
Next iRow
End Sub