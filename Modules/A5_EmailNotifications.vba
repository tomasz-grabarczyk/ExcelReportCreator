Sub AllTicketsNotification()
    ListOfAllTickets.Show
End Sub
Sub DailyTransports()

    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    Call DailyTranportsBackend("export.csv", 45, "A2:A1000", "A2", "E:F,H:H,J:J,M:N,Q:R,T:X")

End Sub
Sub ClearNotifications()
    Sheets("ReportCreator").Visible = xlSheetVisible
    
    Sheets("ReportCreator").Select
    Columns("J:AV").Delete Shift:=xlToLeft
    Rows("4:3000").Delete Shift:=xlUp
    Range("A2").Select
    
    Sheets("ReportCreator").Visible = xlSheetVeryHidden
    
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
End Sub
