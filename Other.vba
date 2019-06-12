Sub saveWorkbookSubroutine(motherFileName As String, directory As String)
    
    ActiveWorkbook.Save

    ChDir directory
    ActiveWorkbook.SaveAs fileName:=motherFileName
    
    Sheets("Sheet1").Select

    
    Range("A2:AZ10000").Copy
    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveWorkbook.Save
    
    Application.Quit
End Sub
Sub saveWorkbookPopupWindow()
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call saveWorkbookSubroutine( _
                                    "https://atos365-my.sharepoint.com/personal/tomasz_grabarczyk_atos_net/Documents/AMS_ARDAGH/AMS_ARDAGH.xlsm", _
                                    "C:\Users\A702387\OneDrive - Atos\AMS_ARDAGH")
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call saveWorkbookSubroutine( _
                                    "https://atos365-my.sharepoint.com/personal/tomasz_grabarczyk_atos_net/Documents/AMS_ARDAGH/AMS_ARDAGH.xlsm", _
                                    "C:\Users\A700473\Atos\Grabarczyk, Tomasz - AMS_ARDAGH (3)")
    
    End If
End Sub
Sub SetLayout()
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
Sub ScrollMaxUpAndLeft()
    Sheets("Sheet1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
End Sub
