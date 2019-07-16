Sub CheckDates()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessage(2)
        
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:BG").Select
    Selection.EntireColumn.Hidden = True
    
    Dim checkDatesToCell As Integer: checkDatesToCell = 10000
    
    For colorAssigned = 2 To checkDatesToCell
        If (range("F" & colorAssigned).Value = "Assigned" Or range("F" & colorAssigned).Value = "In Progress" Or range("F" & colorAssigned).Value = "Pending" Or range("F" & colorAssigned).Value = "Resolved") And range("K" & colorAssigned).Value = "" Then
            range("K" & colorAssigned).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            range("C" & colorAssigned).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorAssigned
    
    For colorInProgress = 2 To checkDatesToCell
        If (range("F" & colorInProgress).Value = "In Progress" Or range("F" & colorInProgress).Value = "Pending" Or range("F" & colorInProgress).Value = "Resolved") And range("L" & colorInProgress).Value = "" Then
            range("L" & colorInProgress).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            range("C" & colorInProgress).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorInProgress
    
    For colorPending = 2 To checkDatesToCell
        If range("F" & colorPending).Value = "Pending" And range("M" & colorPending).Value = "" Then
            range("M" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        ElseIf range("F" & colorPending).Value = "Resolved" And range("N" & colorPending).Value = "" Then
            range("N" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorPending
    
    For colorResolved = 2 To checkDatesToCell
        If range("F" & colorResolved).Value = "Resolved" And range("O" & colorResolved).Value = "" Then
        range("O" & colorResolved).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            range("C" & colorResolved).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorResolved
    
    ActiveSheet.range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub SAPSystemCorrectness()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessageString("checking SAP Area correctness ...")
    
    Dim CheckSAPSystemCorrectness As Integer: CheckSAPSystemCorrectness = 10000
    
    For C = 2 To CheckSAPSystemCorrectness
        If Not range("H" & C).Value = "" And Not range("F" & C).Value = "" Then
            If Not (range("H" & C).Value = "BP2" Or _
                    range("H" & C).Value = "ACE" Or _
                    range("H" & C).Value = "BP5" Or _
                    range("H" & C).Value = "HRP" Or _
                    range("H" & C).Value = "RE-FX" Or _
                    range("H" & C).Value = "IFRS") Then
                range("H" & C).Select
                With Selection.Interior
                    .Color = 13260
                    .PatternTintAndShade = 0
                End With
                'Color Incident Numbers
                range("C" & C).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next C
    
    ActiveSheet.range("$A$1:$BH$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub SLACheckLayout()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Call BackToNormal
    
    Sheets("Sheet1").Select
    
    ActiveSheet.range("$A$1:$AP$10000").AutoFilter Field:=6, Criteria1:=Array( _
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
    
    ActiveSheet.range("$A$1:$AP$10000").AutoFilter Field:=5
    ActiveSheet.range("$A$1:$AP$10000").AutoFilter Field:=38, Criteria1:=">=" & 11
    
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        range("AX1:AX10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    range("A1").Select
    Sheets("Sheet1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    range("A1").Select
End Sub
Sub TicketResolvingCounter()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 16.07.2019  **********

    Dim DateToday As String, lRow As Long, numberOfTickets As String, arrayOfSAPArea() As String, arrayOfConsultant() As String, TotalTicketNumber() As String
    
    DateToday = Format(range("A2").Value, "YYYY.MM.DD")
    lRow = Cells(Rows.count, 3).End(xlUp).Row
    
    'Create dynamic array with size of total number of rows for SAP Area, Consultants and Total number of tickets
    ReDim arrayOfSAPArea(lRow)
    ReDim arrayOfConsultant(lRow)
    ReDim TotalTicketNumber(lRow)
    
    'Copy SAP Areas and Consultants to an array
    For i = 1 To lRow - 1
        arrayOfSAPArea(i) = Sheets("TicketResolving").range("B" & i + 1)
        arrayOfConsultant(i) = Sheets("TicketResolving").range("C" & i + 1)
    Next i

    For loopNumber = 0 To 2
        Application.ScreenUpdating = False
        
        For j = 1 To lRow - 1
            Sheets("Sheet1").Select
            ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
            
            If arrayOfSAPArea(j) = "Logistics" Then
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=4, Criteria1:="=*Logistic*"
            Else
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=4, Criteria1:=arrayOfSAPArea(j)
            End If
            
            ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=5, Criteria1:=arrayOfConsultant(j)
            
            If loopNumber = 0 Then
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending"), Operator:=xlFilterValues
            ElseIf loopNumber = 1 Then
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:="Resolved"
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=14, Criteria1:="=*" & DateToday & "*", Operator:=xlAnd
            ElseIf loopNumber = 2 Then
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending", "Resolved"), Operator:=xlFilterValues
                ActiveSheet.range("$A$1:$CA$10000").AutoFilter Field:=11, Criteria1:="=*" & DateToday & "*", Operator:=xlAnd
            End If
            
            numberOfTickets = Application.WorksheetFunction.Subtotal(103, range("A2:A10000"))
            TotalTicketNumber(j) = numberOfTickets
        Next j
        
        For k = 1 To lRow
            If loopNumber = 0 Then
                Sheets("TicketResolving").range("D" & k + 1).Value = TotalTicketNumber(k)
            ElseIf loopNumber = 1 Then
                Sheets("TicketResolving").range("E" & k + 1).Value = TotalTicketNumber(k)
            ElseIf loopNumber = 2 Then
                Sheets("TicketResolving").range("F" & k + 1).Value = TotalTicketNumber(k)
            End If
        Next k
        
        Application.Run ("'AMS_ARDAGH_DD_MACROS.xlam'!BackToNormal")
        
        Application.ScreenUpdating = True
    Next loopNumber
    
    Sheets("TicketResolving").Select
    
End Sub
