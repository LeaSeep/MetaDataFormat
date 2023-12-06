Sub CompareAndMergeSheets()
    Dim baseSheet As Worksheet
    Dim compareSheet As Worksheet
    Dim baseLastColumn As Long
    Dim compareLastColumn As Long
    Dim baseColumn As Range
    Dim compareColumn As Range
    Dim baseRange As Range
    Dim compareRange As Range
    Dim addedItems As Range
    Dim cell As Range
    Dim potential_match As Range
    
    Dim base_col As Long
    Dim cmp_col As Long
    
    Dim NextFree_base As Long
    ' Set the base sheet (1st sheet in the workbook)
    Set baseSheet = ThisWorkbook.Sheets(1)
    
    ' Loop through each sheet beside the base sheet
    For Each compareSheet In ThisWorkbook.Sheets
        If compareSheet.Index <> baseSheet.Index Then
            ' Find the last column in the base sheet and compare sheet
            baseLastColumn = baseSheet.Cells(6, baseSheet.Columns.Count).End(xlToLeft).Column
            compareLastColumn = compareSheet.Cells(6, compareSheet.Columns.Count).End(xlToLeft).Column
            
            ' Loop through each cell in the 6th row of the compare sheet
            For Each baseColumn In baseSheet.Range(baseSheet.Cells(6, 2), baseSheet.Cells(6, baseLastColumn))
                For Each compareColumn In compareSheet.Range(compareSheet.Cells(6, 2), compareSheet.Cells(6, compareLastColumn))
                    ' Check if the column names match
                    If baseColumn.Value = compareColumn.Value Then
                        ' Find added items in the compare sheet
                    

                        base_col = baseSheet.Rows(6).Find(What:=baseColumn.Value, LookIn:=xlValues, LookAt:=xlWhole).Column
                        cmp_col = compareSheet.Rows(6).Find(What:=compareColumn.Value, LookIn:=xlValues, LookAt:=xlWhole).Column
                        base_col_Letter = Split(Cells(1, base_col).Address, "$")(1)
                        cmp_col_Letter = Split(Cells(1, cmp_col).Address, "$")(1)
                        ' Find last rows in respective colum from 6th row onward
                        'Set nextEmptyCell = baseSheet.Cells(baseSheet.Rows.Count, columnNumber).End(xlUp).Offset(1)
                        
                       NextFree_base = baseSheet.Rows(5).Columns(base_col).End(xlDown).Offset(1, 0).Row
                       NextFree_cmp = compareSheet.Rows(5).Columns(cmp_col).End(xlDown).Offset(1, 0).Row
                        
                        
                            ' Loop through rows 1 to 6
                        For i = 1 To 6
                            ' Compare values in the corresponding rows of col1 and col2
                            If baseSheet.Cells(i, base_col).Value <> compareSheet.Cells(i, cmp_col).Value Then
                                ' If values are different, add the value from col2 to col1 and mark it green
                                ' Prompt the user to make a decision
                                    decision = MsgBox("Keep current value?" & vbCrLf & "Yes: " & baseSheet.Cells(i, base_col).Value & vbCrLf & "No: " & compareSheet.Cells(i, cmp_col).Value, vbYesNo + vbQuestion, "User Decision")
                                    
                                        ' Check the user's decision
                                        If decision = vbYes Then
                                            ' User chose "Yes" (Value A)
                                            baseSheet.Cells(i, base_col).Value = baseSheet.Cells(i, base_col).Value
                                        
                                        Else
                                            ' User chose "No" (Value B)
                                            baseSheet.Cells(i, base_col).Value = compareSheet.Cells(i, cmp_col).Value
                                          
                                        End If
                                'baseSheet.Cells(i, base_col).Value = compareSheet.Cells(i, cmp_col).Value
                                baseSheet.Cells(i, base_col).Interior.Color = RGB(144, 238, 144) ' Green color
                            End If
                        Next i
                        
                        ' get longest last row
                        maxRow = Application.Max(NextFree_base, NextFree_cmp)

                        If maxRow > 7 Then
                            'stuff needs to be added
                            For currentRow = 7 To maxRow
                                Set potential_match = baseSheet.Columns(base_col).Find(What:=compareSheet.Cells(currentRow, cmp_col).Value, LookIn:=xlValues, LookAt:=xlWhole)
                                 If potential_match Is Nothing Then
                                    ' add it
                                    baseSheet.Cells(NextFree_base, base_col).Value = compareSheet.Cells(currentRow, cmp_col).Value
                                baseSheet.Cells(NextFree_base, base_col).Interior.Color = RGB(144, 238, 144) ' Green color
                                NextFree_base = NextFree_base + 1
                                 End If
                            Next currentRow
                        End If

                        
                        ' Remove the column from the compare sheet
                        compareColumn.EntireColumn.Delete
                        
                        Exit For ' Exit the loop once a match is found
                    End If
                Next compareColumn
            Next baseColumn
            
            ' Copy remaining columns in compare sheet to base sheet and mark green
            compareLastColumn_Letter = Split(Cells(1, compareLastColumn).Address, "$")(1)
            
            Set compareRange = compareSheet.Range("B1:" & compareLastColumn_Letter & maxRow)
            compareRange.Copy Destination:=baseSheet.Cells(1, baseLastColumn + 1)
            baseSheet.Range(baseSheet.Columns(baseLastColumn + 1), baseSheet.Columns(baseLastColumn + compareLastColumn)).Interior.Color = RGB(144, 238, 144) ' Green color
        End If
    Next compareSheet
    
    ' Provide a message to the user
    MsgBox "Check all added items in the base sheet for curation.", vbInformation
End Sub



