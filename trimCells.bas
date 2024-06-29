Attribute VB_Name = "trimCells"
Option Private Module
Sub trimmer()

    Dim myLastCellAdd
    Set myLastCell = Cells(1, 1).SpecialCells(xlLastCell)
    myLastCellAdd = Cells(myLastCell.Row, myLastCell.Column).Address
    Set myRange = Range("A1:" & myLastCellAdd)

    For Each Cell In myRange
        Cell.Value = Application.WorksheetFunction.Trim(Cell.Value)
    Next

End Sub
