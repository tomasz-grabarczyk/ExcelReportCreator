Sub AllTicketsNotification()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    ListOfAllTickets.Show
End Sub
Sub DailyTransports()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    Call DailyTranportsBackend("export.csv", 45, "A2:A1000", "A2", "E:F,H:H,J:J,M:N,Q:R,T:X")

End Sub
Sub ClearNotifications()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessage(5)
    
    Sheets("ReportCreator").Visible = True
    
    Sheets("ReportCreator").Select
    Columns("J:AV").Delete Shift:=xlToLeft
    Rows("4:3000").Delete Shift:=xlUp
    Range("A2").Select
    
    'Clear conditional formating from ReportCreator sheet
    Cells.FormatConditions.Delete
    
    Sheets("ReportCreator").Visible = False
    
    Call BackToNormal
    
    Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("K1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Call BackToNormal
    Call SortByStatuses
    
    Call StopMacroShowMessage
End Sub
