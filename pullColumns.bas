Attribute VB_Name = "pullOR"
Option Private Module
Sub pullCorrectColumnsOR()

    Dim lastColumn As Long
'    Dim myLastCellAdd
    
    Dim myLastCell
    Dim MR
    Set myLastCell = Cells(1, 1).SpecialCells(xlLastCell)
'    myLastCellAdd = Cells(myLastCell.Row, myLastCell.Column).Address

    lastColumn = Cells(1, Columns.count).End(xlToLeft).Column
    Set MR = Range("A1", Cells(1, Columns.count).End(xlToLeft))

'----------------------------------------------------------------------------------------
        For Each Cell In MR
            If InStr(1, Cell.Value, "Email") Then
                Worksheets("Sheet1").Range("A1:" & "A" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
        For Each Cell In MR
            If InStr(1, Cell.Value, "First") Then
                Worksheets("Sheet1").Range("B1:" & "B" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
        For Each Cell In MR
            If InStr(1, Cell.Value, "Last") Then
                Worksheets("Sheet1").Range("C1:" & "C" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
             End If
        Next
'----------------------------------------------------------------------------------------
        For Each Cell In MR
            If InStr(1, Cell.Value, "Company") Then
                Worksheets("Sheet1").Range("D1:" & "D" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
        For Each Cell In MR
            If InStr(1, Cell.Value, "Company Size") Then
                Worksheets("Sheet1").Range("E1:" & "E" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Country") Then
                Worksheets("Sheet1").Range("F1:" & "F" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Address 1") Then
                Worksheets("Sheet1").Range("G1:" & "G" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Address 2") Then
                Worksheets("Sheet1").Range("H1:" & "H" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "City") Then
                Worksheets("Sheet1").Range("I1:" & "I" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "State") Then
                Worksheets("Sheet1").Range("J1:" & "J" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
'----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Zip") Then
                Worksheets("Sheet1").Range("K1:" & "K" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Phone") Then
                Worksheets("Sheet1").Range("L1:" & "L" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
  '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Cell") Then
                Worksheets("Sheet1").Range("M1:" & "M" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Industry") Then
                Worksheets("Sheet1").Range("N1:" & "N" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
  '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "Asset Name") Then
                Worksheets("Sheet1").Range("O1:" & "O" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If Cell.Value = ("Registration Date (without time)") Then
                'cell.Value = "Registration Date"
                Worksheets("Sheet1").Range("P1:" & "P" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "JobTitle Function") Then
                Worksheets("Sheet1").Range("Q1:" & "Q" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
            For Each Cell In MR
            If InStr(1, Cell.Value, "JobTitle Position") Then
                Worksheets("Sheet1").Range("R1:" & "R" & myLastCell.Row).Value = Worksheets("Leads").Range(Cell.Address & ":" & Split(Cell.Address, "$")(1) & myLastCell.Row).Value
                Exit For
            End If
        Next
 '----------------------------------------------------------------------------------------
Worksheets("Sheet1").Activate

End Sub

