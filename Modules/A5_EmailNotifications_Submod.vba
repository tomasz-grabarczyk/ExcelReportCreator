Sub DailyTranportsBackend(filePath As String, cellRange As String, copyToCell As String, columnsToBeDeleted As String)
    Sheets("DailyTransports").Visible = True
    Sheets("DailyTransports").Select
    Range(cellRange).ClearContents

    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    If Dir(filePath) = "" Then
        MsgBox "Could not find the file: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath)

    Range(columnsToBeDeleted).Select
    Selection.Delete Shift:=xlToLeft
    
    Rows("1:1").Delete Shift:=xlToLeft
    
    Range("A1:V100").Copy
    Windows(thisfile).Activate
    Range(copyToCell).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(filePath).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
    
    Sheets("DailyTransports").Select
    Range("A1").Select
End Sub
Sub SortData()
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("E1:E10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub FormatText()
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    Selection.Font.Bold = True
End Sub
Sub MacroListOfAllTickets(ByVal area As String, Optional ByVal area_second As String, Optional ByVal area_third As String, Optional ByVal area_fourth As String)
    Call StartMacroShowMessage(10)
    
    Call ClearNotifications
    
    Sheets("ReportCreator").Visible = xlSheetVisible
    
    
    Sheets("ReportCreator").Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    
    
    Sheets("Sheet1").Select
    Columns("A:A").EntireColumn.Hidden = True
    Columns("G:G").EntireColumn.Hidden = True
    Columns("I:Y").EntireColumn.Hidden = True
    Columns("AA:AD").EntireColumn.Hidden = True
    Columns("AF:AV").EntireColumn.Hidden = True
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:="Assigned"
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=4, Criteria1:=Array(area, area_second, area_third, area_fourth), Operator:=xlFilterValues
    
    
    Call SortData
    
    
    Range("B2:AW10000").Copy
    Sheets("ReportCreator").Select
    Sheets("ReportCreator").Paste
    

    Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Select
    Range(ActiveCell.Address).Value = "IN PROGRESS"
    Call FormatText
    
    
    Range("A3:I3").Copy
    Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Select
    Range(ActiveCell.Address).PasteSpecial
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:="In Progress"
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=4, Criteria1:=Array(area, area_second, area_third, area_fourth), Operator:=xlFilterValues
    
    Call SortData
    
    Range("B2:AW10000").Copy
    Sheets("ReportCreator").Select
    Sheets("ReportCreator").Paste
    
    
    Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Select
    Range(ActiveCell.Address).Value = "PENDING"
    Call FormatText
    
    
    Range("A3:I3").Copy
    Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Select
    Range(ActiveCell.Address).PasteSpecial
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    
    
    
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=6, Criteria1:="Pending"
    ActiveSheet.Range("$A$1:$AW$10000").AutoFilter Field:=4, Criteria1:=Array(area, area_second, area_third, area_fourth), Operator:=xlFilterValues
    
    Call SortData
    
    Range("B2:AW10000").Copy
    Sheets("ReportCreator").Select
    Sheets("ReportCreator").Paste


    Call BackToNormal
    
    
    Sheets("ReportCreator").Select
    Columns("I:I").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0", Formula2:="=5"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
  
    With ActiveSheet.Range("I1:I300")           ' removes formatting from empty cells
    .SpecialCells(xlCellTypeBlanks).ClearFormats
    End With
    
    Call SortByStatuses
    
    Sheets("ReportCreator").Select
    Range("A1").Select
    
    Call StopMacroShowMessage
End Sub
