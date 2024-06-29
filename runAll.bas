Attribute VB_Name = "AAArun"
Sub runEverythingOR()
Attribute runEverythingOR.VB_ProcData.VB_Invoke_Func = "m\n14"
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    
    Call addPageSelectPageOR
'    Call renameColumns 'No need in OneReg pulls
    Call pullCorrectColumnsOR
    Call doPhone
    Call formatAll
'    Call highlightBlank
    
    'MsgBox ("Finished, fix the dates and check the filters.")
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    
    Dim myLastCellAddCopy
    Set myLastCellAddCopy = Cells(1, 1).SpecialCells(xlLastCell)
    myLastCellAddCopy = Cells(myLastCellAddCopy.Row, myLastCellAddCopy.Column).Address
    Dim myRange
    Set myRange = Range("A2:" & myLastCellAddCopy)
    
    myRange.Select
    Selection.Copy

    
    
End Sub
