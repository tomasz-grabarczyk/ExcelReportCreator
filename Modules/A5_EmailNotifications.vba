Sub SLACheckLayout()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Call BackToNormal
    
    Sheets("Sheet1").Select
    
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", _
        "In Progress", _
        "Pending"), Operator:=xlFilterValues
    
    Columns("A:A").EntireColumn.Hidden = True
    Columns("G:G").EntireColumn.Hidden = True
    Columns("I:Y").EntireColumn.Hidden = True
    Columns("AA:AD").EntireColumn.Hidden = True
    Columns("AF:AM").EntireColumn.Hidden = True
    Columns("AO:AV").EntireColumn.Hidden = True
    Columns("AY:BG").EntireColumn.Hidden = True
    
    ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=2, Criteria1:="User Service Restoration"
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=5
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=38, Criteria1:=">=" & 11
    
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AX1:AX10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    Sheets("Sheet1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
End Sub
Sub AllTicketsNotification()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    ListOfAllTickets.Show
End Sub
Sub DailyTransports()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    Call DailyTranportsBackend("export.csv", "A2:A1000", "A2", "E:F,H:H,J:J,M:N,Q:R,T:X")

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
Sub TicketResolvingCounter()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 17.07.2019  **********
    
    Call BackToNormal
        
    Sheets("TicketResolving").Visible = True
    Sheets("TicketResolving").Select
    
    Dim DateToday As String, lRow As Long, numberOfTickets As String, arrayOfSAPArea() As String, arrayOfConsultant() As String, TotalTicketNumber() As String
    
    DateToday = Format(Range("A2").Value, "YYYY.MM.DD")
    lRow = Cells(Rows.count, 3).End(xlUp).Row
    
    'Create dynamic array with size of total number of rows for SAP Area, Consultants and Total number of tickets
    ReDim arrayOfSAPArea(lRow)
    ReDim arrayOfConsultant(lRow)
    ReDim TotalTicketNumber(lRow)
    
    'Copy SAP Areas and Consultants to an array
    For i = 1 To lRow - 1
        arrayOfSAPArea(i) = Sheets("TicketResolving").Range("B" & i + 1)
        arrayOfConsultant(i) = Sheets("TicketResolving").Range("C" & i + 1)
    Next i

    For loopNumber = 0 To 2
        For j = 1 To lRow - 1
            Sheets("Sheet1").Select
            ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
            
            If arrayOfSAPArea(j) = "Logistics" Then
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=4, Criteria1:="=*Logistic*"
            Else
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=4, Criteria1:=arrayOfSAPArea(j)
            End If
            
            ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=5, Criteria1:=arrayOfConsultant(j)
            
            If loopNumber = 0 Then
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending"), Operator:=xlFilterValues
            ElseIf loopNumber = 1 Then
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:="Resolved"
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=14, Criteria1:="=*" & DateToday & "*", Operator:=xlAnd
            ElseIf loopNumber = 2 Then
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending", "Resolved"), Operator:=xlFilterValues
                ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=11, Criteria1:="=*" & DateToday & "*", Operator:=xlAnd
            End If
            
            numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))
            TotalTicketNumber(j) = numberOfTickets
        Next j
        
        For k = 1 To lRow
            If loopNumber = 0 Then
                Sheets("TicketResolving").Range("D" & k + 1).Value = TotalTicketNumber(k)
            ElseIf loopNumber = 1 Then
                Sheets("TicketResolving").Range("E" & k + 1).Value = TotalTicketNumber(k)
            ElseIf loopNumber = 2 Then
                Sheets("TicketResolving").Range("F" & k + 1).Value = TotalTicketNumber(k)
            End If
        Next k
        
        Call BackToNormal
    Next loopNumber
    
    Sheets("TicketResolving").Select
    
    MsgBox "Job has finished successfully!"
End Sub
