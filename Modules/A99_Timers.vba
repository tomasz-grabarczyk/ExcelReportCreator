Sub startMacroShowMessageString(nameOfProcess As String)
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.DisplayStatusBar = True
    Application.StatusBar = "Working on it ... I'm working on " & nameOfProcess
End Sub
Sub startMacroShowMessage(numberOfSecond As Integer)
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.DisplayStatusBar = True
    Application.StatusBar = "Working on it ... It will take about " & numberOfSecond & " seconds ..."
End Sub
Sub stopMacroShowMessage()
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub timerToPaste()
    'start timer
    Dim StartTime As Double
        Dim SecondsElapsed As Double
        StartTime = Timer
    
    'stop timer
    SecondsElapsed = Round(Timer - StartTime, 2)
        MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub
