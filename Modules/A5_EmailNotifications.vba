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
Sub LogisticsReminder()
    Sheets("LogisticsReminder").Visible = True
    
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:X").Select
    Selection.EntireColumn.Hidden = True
    Columns("AB:AD").Select
    Selection.EntireColumn.Hidden = True
    Columns("AG:BG").Select
    Selection.EntireColumn.Hidden = True
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=4, Criteria1:=Array( _
        "Logistic ACE - (PP, PM, QM, MM)", "Logistic ACE - Mass upload/change", _
        "Logistic ACE - WM (APL)", "Logistics"), Operator:=xlFilterValues
        
    ActiveSheet.Range("$A$1:$BG$3323").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", "In Progress", "Pending"), Operator:=xlFilterValues
    
    Call ScrollMaxUpAndLeft
End Sub
