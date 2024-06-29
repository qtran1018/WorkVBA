Attribute VB_Name = "personalMail"
Sub personalEmail()
    Dim findmail
    Dim findmailRange
    Dim mailColumn
    findmail = Rows(1).Find("Email").Address
    findmailRange = Range(findmail, findmail).Offset(1, 0).Address
    mailColumn = Range(findmail, findmail).Offset(1, 0).Column
    
    'Sets the range to the offset cell --> last cell with a value in that column.
    Set mailRange = Range(findmailRange, Cells(Rows.count, mailColumn).End(xlUp))
    
    Dim count As Integer
    
    For Each Cell In mailRange
        If InStr(1, Cell.Value, "gmail") Then
            Cell.Interior.Color = RGB(255, 51, 51)
            count = count + 1
        End If
        If InStr(1, Cell.Value, "yahoo") Then
            Cell.Interior.Color = RGB(255, 51, 51)
                count = count + 1
        End If
        If InStr(1, Cell.Value, "hotmail") Then
            Cell.Interior.Color = RGB(255, 51, 51)
                count = count + 1
        End If
        If InStr(1, Cell.Value, "me.com") Then
            Cell.Interior.Color = RGB(255, 51, 51)
                count = count + 1
        End If
        If InStr(1, Cell.Value, "aol.com") Then
            Cell.Interior.Color = RGB(255, 51, 51)
                count = count + 1
        End If
    Next
    
    If count > 0 Then
        MsgBox (count & " personal emails.")
    End If
    
End Sub
