Sub PaintAllTickets()
    Call startMacroShowMessage(3)
    
    Call backToNormal
    Sheets("Sheet1").Select

    'Paint Resolved statuses
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=6, Criteria1:="Resolved"
    Range("A1:BG10000").Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    
    'Paint Cancelled and Closed statuses
    Call backToNormal
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=6, Criteria1:="=Cancelled", Operator:=xlOr, Criteria2:="=Closed"
    Range("A1:BG10000").Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    
    'Paint N/A Consultants
    Call backToNormal
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=5, Criteria1:="N/A"
    Range("A1:BG10000").Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    
    'Paint everything else white
    Call backToNormal
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", "In Progress", "Pending"), Operator:=xlFilterValues
    Range("A1:BG10000").Select
    With Selection.Interior
        .Color = RGB(255, 255, 255)
        .TintAndShade = 0
    End With
    
    Call backToNormal
    Sheets("Sheet1").Select
    Rows(1).Select
    With Selection.Interior
        .Color = RGB(221, 217, 195)
        .TintAndShade = 0
    End With
    
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Sheets("Sheet1").Select
    Range("A1").Select
    
    Call stopMacroShowMessage
End Sub
Sub clearIrrelevantData()
    Call startMacroShowMessage(5)
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$3000").AutoFilter Field:=5, Criteria1:="N/A"
    Range("L2:Q10000,V2:X10000,AA2:AB10000").ClearContents
    Range("A1").Select
    
    Call backToNormal
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Cancelled", "Closed", "Resolved"), Operator:= _
        xlFilterValues
    Range("G2:G10000").ClearContents
    
    Call backToNormal
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", "Cancelled", "Closed", "In Progress", "Resolved"), Operator:= _
        xlFilterValues
    Range("M2:M10000").ClearContents
    
    Call backToNormal
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Cancelled", "Closed", "Resolved"), Operator:= _
        xlFilterValues
    Range("AG2:AX10000").ClearContents
    
    Call backToNormal
    
    Call stopMacroShowMessage
End Sub
Sub ApplyFiltersToStatus()

    Sheets("Sheet1").Select
    
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F2:F10000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Assigned,In Progress,Pending,Resolved,Closed,Cancelled", DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
Sub backToNormal()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.FormatConditions.Delete
    Next ws

    Sheets("Sheet1").Select
    Sheet1.AutoFilterMode = False
    Rows("1:1").AutoFilter
    Columns("A:CC").EntireColumn.Hidden = False
    Range("AE:AE,AI:AI").WrapText = False
    Sheets("Sheet1").Select
    
    'Delete white characters (spaces) from column Incident Number
    [C2:C10000] = [INDEX(TRIM(C2:C10000),)]
    
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Columns("AQ:AV").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:Q").Select
    Selection.EntireColumn.Hidden = True
    Columns("AA:AA").Select
    Selection.EntireColumn.Hidden = True
    Columns("AF:AF").Select
    Selection.EntireColumn.Hidden = True
    Columns("AJ:AK").Select
    Selection.EntireColumn.Hidden = True
    
    Rows("1:1").AutoFilter
    Range("A1").Select
    
   
End Sub