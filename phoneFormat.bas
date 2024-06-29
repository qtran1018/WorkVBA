Attribute VB_Name = "phoneFormat"

Sub doPhone()
' ----------------------------------------------------------------------------------------
    ' Set vars to find "Phone" address, the address of the cell below "Phone", and get the Column Index of the offset cell.
    
    Dim findPhone
    Dim findPhoneRange
    Dim phoneColumn
    findPhone = Rows(1).Find("Phone").Address
    findPhoneRange = Range(findPhone, findPhone).Offset(1, 0).Address
    phoneColumn = Range(findPhone, findPhone).Offset(1, 0).Column
    
        'Sets the range to the offset cell --> last cell with a value in that column.
    Set phoneRange = Range(findPhoneRange, Cells(Rows.count, phoneColumn).End(xlUp))

    If Not ActiveSheet.Range("F2") = "United States of America" Then
        phoneRange.NumberFormat = "0"
        Exit Sub
    End If

    'Deletes those chars from the cell, sets the cell to the right 10 chars, then formats the cell.
'    For Each Cell In phoneRange
'        Cell.Value = Application.WorksheetFunction.Trim(Cell.Value)
'        Cell.Value = Replace(Cell.Value, " ", "") ' regular space
'        Cell.Value = Replace(Cell.Value, " ", "") ' non-breaking space ALT+255
'        Cell.Value = Replace(Cell.Value, ".", "")
'        Cell.Value = Replace(Cell.Value, "-", "")
'        Cell.Value = Replace(Cell.Value, "(", "")
'        Cell.Value = Replace(Cell.Value, ")", "")
'        Cell.Value = Right(Cell.Value, 10)
'        'cell.NumberFormat = "###-###-####"
'    Next
        phoneRange.NumberFormat = "(###)-###-####"
' ----------------------------------------------------------------------------------------
End Sub
