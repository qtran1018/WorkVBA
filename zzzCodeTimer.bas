Attribute VB_Name = "zzzCodeTimer"
Sub timeCode()

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'*****************************
'CODE HERE
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'Application.DisplayStatusBar = False
Call runEverythingOR
'Application.ScreenUpdating = True
'Application.Calculation = xlCalculationAutomatic
'Application.DisplayStatusBar = True

'*****************************

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub

