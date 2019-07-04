Sub StartMacroShowMessageString(nameOfProcess As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.DisplayStatusBar = True
    Application.StatusBar = "Working on it ... " & nameOfProcess
End Sub
Sub StartMacroShowMessage(numberOfSecond As Integer)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.DisplayStatusBar = True
    Application.StatusBar = "Working on it ... It will take about " & numberOfSecond & " seconds ..."
End Sub
Sub StopMacroShowMessage()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub TimerToPaste()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    'start timer
    Dim StartTime As Double
        Dim SecondsElapsed As Double
        StartTime = Timer
    
    'stop timer
    SecondsElapsed = Round(Timer - StartTime, 2)
        MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub
