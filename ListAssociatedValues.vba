Sub ListAssociatedValues()

    'Takes a structured Excel extract, and adds two columns with data calculated from the other columns and logically associated rows.
    'Strictly speaking this is O(N^2 * MaxAssociatedRows)
    'However in most cases the number of reference rows X <<< N, so in real life real running time is closer to O(N * AverageAssociatedRows)
    'On a dataset of N~=1M, X~=1000, & AverageAssociatedRows~=8 this runs in 8 minutes - slow but liveable
    'For X closer to N, this will probably need to be rewritten with a dictionary tree to at least get to O(N * log N)

    'Constants for which column contains which data
    Const iColumnId As Long = 1  'Unique identifier that associates rows together
    Const iColumnAssociatedCount = 7 'how many associated rows linked to this row, including the reference row in the count
    
    'The contents of these columns from associated rows will be written out by the macro to a new column
    'VBA doesn't support Array type constants
    'Update this for your case
    Dim aColumnsToList(1 To 3) As Long
        aColumnsToList(1) = 8
        aColumnsToList(2) = 9
        aColumnsToList(3) = 10
    
    Const iColumnIsReference As Long = 11 'new column created by the macro and filled with Yes/No : is this a reference row or not?
    Const iColumnAssociatedValues As Long = 12 'new column created by the macro and filled for reference rows, listing the desired texts from each associated rows, comma separated
    
    'If there are duplicate rows to remove, identify them based on the data in these columns being the same.
    'VBA doesn't support Array type constants
    'Update this for your case
    Dim aColumnsDuplicates(1 To 4) As Long
        aColumnsDuplicates(1) = 1
        aColumnsDuplicates(2) = 3
        aColumnsDuplicates(3) = 8
        aColumnsDuplicates(4) = 9
    
    
    Const iRowTitle As Long = 1 'This row is the titles for the columns
    Const iRowDataStart As Long = 2 'This row is the first one containing real data
    'Last row of the data - can't assign Rows.Count to a Const, Excel has 'issues' with strict typing ;)
    Dim iLastRow As String
    iLastRow = Rows.Count
    
    'String constants
    Const sTitle1 As String = "Relavant Row?" 'Title text for the first added column
    Const sTitle2 As String = "Data From Related Rows" 'Title text for the second added column
    Const sYes As String = "Yes" 'True value to put in first added column
    Const sNo As String = "No" 'False value to put in first added column
        
        
        
    'Working variables
    Dim iCurrentRow As Long 'Outer loop finding reference rows
    Dim iTempRow As Long 'Inner loop finding associated rows
    Dim sInsert As String 'Text to add to column 11 of the reference rows
    Dim iFound As Long 'How many of the associated rows have been found in the inner loop for the chosen reference row in the outer loop
    Dim iAssociateRowToFind As Long 'How associated rows do we have to find inner loop for this reference row
    Dim bIsReference As Boolean 'Stores if this is reference row
    Dim aDuplicateCheck() As String 'Stores reference & associated rows for duplicate checking, and thus removal
    Dim bIsDuplicate As Boolean 'Stores if we have found a duplicate row
    Dim vDataColumn As Variant 'To loop columns with the data we want
    Dim sTempDataExtract As String 'Temporary while extracting data we want from an associated row
    Dim sTempDataExtractCombined As String 'Temporary while extracting data we want from an associated row
    
    'Excel is strange - if you update data cell/row by cell/row, it gets extremely slow if the file is big
    'If you do the same thing in an array, and then push the array to range, you don't have this speed issue, so...
    'Using two dimensional arrays, as otherwise Excel treats each element as a column of the first row, so insert works incorrectly
    'With a second dimensional, the first dimension is then treated as rows for insert - the second dimension is never used otherwise
    'This is where we build up what to push to the new Is Reference column
    Dim aColumnsIsReference()
    ReDim aColumnsIsReference(1 To iLastRow, 2)
    'This is where we build up what to push to the new Associated Values column
    Dim aColumnsAssociatedValues()
    ReDim aColumnsAssociatedValues(1 To iLastRow, 2)
    
    'Set the title rows for the new columns
    aColumnsIsReference(iRowTitle, 0) = sTitle1
    aColumnsAssociatedValues(iRowTitle, 0) = sTitle2
    
    'Outer loop of all rows of the sheet
    For iCurrentRow = iRowDataStart To iLastRow
        'Check if this is reference row, and fill the appropriate column
        bIsReference = IsReferenceRow(iCurrentRow)
        If bIsReference Then
            aColumnsIsReference(iCurrentRow, 0) = sYes
        Else
            aColumnsIsReference(iCurrentRow, 0) = sNo
        End If
        
        'If this is a reference row, and if it has at least 1 associate row, then we must fill the associated values column
        If bIsReference And (Cells(iCurrentRow, iColumnAssociatedCount).Value > 1) Then
            'Initialise the inner loop
            sInsert = ""
            iFound = 0
            iAssociateRowToFind = Cells(iCurrentRow, iColumnAssociatedCount).Value - 1
            
            'Initialise duplicate checking
            bIsDuplicate = IsDuplicateRow(iCurrentRow, aDuplicateCheck, aColumnsDuplicates, True)
            
            'Inner loop of all rows of the sheet
            For iTempRow = iRowDataStart To iLastRow
                'Skip if this the same reference row, we have found an associate row if the unique ID is the same
                If (iCurrentRow <> iTempRow) And (Cells(iCurrentRow, iColumnId).Value = Cells(iTempRow, iColumnId).Value) Then
                    'Increment how many of the associated rows you have found
                    iFound = iFound + 1
                    
                    'First check if this a duplicate row for any of the already seen rows in this round - if it is, we can skip
                    If IsDuplicateRow(iTempRow, aDuplicateCheck, aColumnsDuplicates, False) Then
                        'Continue to the next round of loop
                    Else
                        'This is a relevant associated row, so add the values from the required columns to the working string that will later be place in Associated Values column on the reference row
                        'To format nicely, put space between data from columns, and comma between combined data for each associated row
                        sTempDataExtract = ""
                        sTempDataExtractCombined = ""
                        For Each vDataColumn In aColumnsToList
                            'Get the data for this column
                            sTempDataExtract = Cells(iTempRow, vDataColumn).Value
                            'If this cell was empty, can skip to the next cell
                            If sTempDataExtract = "" Then
                                'Continue
                            Else
                                'If the combined data from this row isn't filled yet, put the data from the cell just loaded
                                'If the combined data from this row already has some cell data, append the new cell data with a space between
                                If sTempDataExtractCombined = "" Then
                                    sTempDataExtractCombined = sTempDataExtract
                                Else
                                    sTempDataExtractCombined = sTempDataExtractCombined & " " & sTempDataExtract
                                End If
                            End If
                        Next vDataColumn
                        
                        'If no data was found from this associate row, then we can skip to the next row
                        If sTempDataExtractCombined = "" Then
                            'Continue
                        Else
                            'If the data to insert in the reference row doesn't yet have data from at least one associate row, then directly add the data added for this row
                            'If there is already data, append the new row data with a comma between
                            If sInsert = "" Then
                                sInsert = sTempDataExtractCombined
                            Else
                                sInsert = sInsert & ", " & sTempDataExtractCombined
                            End If
                        End If
                     End If
                End If
                
                'If you have found all the associated rows for the reference row, we can stop searching
                If iFound = iAssociateRowToFind Then Exit For
            Next iTempRow
            
            'Add the extracted data to the array to be pushed later
            aColumnsAssociatedValues(iCurrentRow, 0) = sInsert
        Else
            'If this isn't a reference row that has associated rows, we will push an empty string to the associated values column
            aColumnsAssociatedValues(iCurrentRow, 0) = ""
        End If
        
    Next iCurrentRow
    
    ' Now push the data arrays to the columns we add
    Range(Cells(iRowTitle, iColumnIsReference), Cells(iLastRow, iColumnIsReference)) = aColumnsIsReference
    Range(Cells(iRowTitle, iColumnAssociatedValues), Cells(iLastRow, iColumnAssociatedValues)) = aColumnsAssociatedValues
    
End Sub

Function IsReferenceRow(iRowToCheck As Long) As Boolean
    'Update this function with the test for your case
    'This case: column 8 value is checked, if starts with "WP", then this is a reference row, otherwise it is not
    If Left(Cells(iRowToCheck, 8).Value, 2) = "WP" Then
        IsReferenceRow = True
    Else
        IsReferenceRow = False
    End If
End Function

Function IsDuplicateRow(iRowToCheck As Long, ByRef aSeenRows() As String, ByRef aColumnsDuplicates() As Long, bForceReset As Boolean) As Boolean
    'Checks if a row is a duplicate of one already seen in this pass of associated rows.
    'This function also maintains the array of seen rows - adding rows that are checked OK, and resetting when a duplicate is found.

    'Temporary variables for the check, and array traversals
    Dim sCheckedRowData As String
    Dim vTempSeenRow As Variant
    Dim vTempColumn As Variant
       
    'Get the data from the row you want to test, concatenate directly
    sCheckedRowData = ""
    For Each vTempColumn In aColumnsDuplicates
        sCheckedRowData = sCheckedRowData & Cells(iRowToCheck, vTempColumn).Value
    Next vTempColumn
    
    'If this is the first row (ie your reference row), then we just reset the seen rows, add this row, and return false
    If bForceReset Then
        Erase aSeenRows
        ReDim aSeenRows(1 To 1)
        aSeenRows(1) = sCheckedRowData
        IsDuplicateRow = False
    Else
        'Check for duplicates in the already seen rows.
        For Each vTempSeenRow In aSeenRows
            If sCheckedRowData = vTempSeenRow Then
                'If you find the same data already seen, this is a duplicate, we can stop processing.
                IsDuplicateRow = True
                Exit For
            End If
        Next vTempSeenRow
        
        'If none of the seen rows matched, before return, add the OK row to the seen rows array
        If IsDuplicateRow Then
            'Already set, we are done
        Else
            'No row matched
            ReDim Preserve aSeenRows(1 To (UBound(aSeenRows) + 1))
            aSeenRows(UBound(aSeenRows)) = sCheckedRowData
            IsDuplicateRow = False
        End If
    End If
End Function
