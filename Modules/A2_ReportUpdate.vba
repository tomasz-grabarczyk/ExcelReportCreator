Sub UpdateReport()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"

    Call RunVLookUpsUpdatedReport("ArdaghDailyUpdateReport.xls")
    
    Call ScrollMaxUpAndLeft
End Sub
Sub AddNewTicket()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Sheets("Sheet1").Select
    Rows("2:2").Select
    Selection.Copy
    Rows("3:3").Insert Shift:=xlDown
    Range("A2").Select
    Application.CutCopyMode = False
    Range("C2").ClearContents
    Range("K2:O2").ClearContents
End Sub
Sub RemoveFirstRow()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Sheets("Sheet1").Select
    If Range("C2") = "" Then
        Rows("2:2").Delete Shift:=xlUp
    Else
        MsgBox "You are trying to delete row with ticket data!"
    End If
End Sub
Sub ResolvedIntoValues()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F2:F10000"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Resolved", DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim numberOfResolvedTickets As Integer: numberOfResolvedTickets = 1
    
    Dim rng As Range, cell As Range
    Set rng = Range("A1:A1000")
    For Each cell In rng
        If cell.Interior.ThemeColor = xlThemeColorAccent5 Then
            numberOfResolvedTickets = numberOfResolvedTickets + 1
        End If
    Next cell
    
    'Select all column from A to BG
    Range("A2:BG" & numberOfResolvedTickets).Select
    
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
    Application.SendKeys ("{ESC}")
    
    Range("A1").Select
End Sub
