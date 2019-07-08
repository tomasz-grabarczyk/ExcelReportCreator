Sub SaveToMotherFile()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ActiveWorkbook.Save

    If Len(Dir("C:\Users\" & Environ$("Username") & "\Downloads\!AMS_ARDAGH", vbDirectory)) = 0 Then
        MkDir ("C:\Users\" & Environ$("Username") & "\Downloads\!AMS_ARDAGH")
    End If
    
    ChDir "C:\Users\" & Environ$("Username") & "\Downloads\!AMS_ARDAGH"
    ActiveWorkbook.SaveAs fileName:="AMS_ARDAGH.xlsm"
    
    Sheets("Sheet1").Select

    Range("A2:AZ10000").Copy
    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveWorkbook.Save
    
    Application.Quit
                         
End Sub
Sub DefaultLayout()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Sheets("Sheet1").Select
    ActiveWindow.Zoom = 85
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Sheets("PendingCalculator").Select
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Sheets("Sheet1").Select
    Range("A1").Select
End Sub
