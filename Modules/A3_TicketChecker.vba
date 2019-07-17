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
        If (Range("F" & colorAssigned).Value = "Assigned" Or Range("F" & colorAssigned).Value = "In Progress" Or Range("F" & colorAssigned).Value = "Pending" Or Range("F" & colorAssigned).Value = "Resolved") And Range("K" & colorAssigned).Value = "" Then
            Range("K" & colorAssigned).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & colorAssigned).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorAssigned
    
    For colorInProgress = 2 To checkDatesToCell
        If (Range("F" & colorInProgress).Value = "In Progress" Or Range("F" & colorInProgress).Value = "Pending" Or Range("F" & colorInProgress).Value = "Resolved") And Range("L" & colorInProgress).Value = "" Then
            Range("L" & colorInProgress).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & colorInProgress).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorInProgress
    
    For colorPending = 2 To checkDatesToCell
        If Range("F" & colorPending).Value = "Pending" And Range("M" & colorPending).Value = "" Then
            Range("M" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        ElseIf Range("F" & colorPending).Value = "Resolved" And Range("N" & colorPending).Value = "" Then
            Range("N" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorPending
    
    For colorResolved = 2 To checkDatesToCell
        If Range("F" & colorResolved).Value = "Resolved" And Range("O" & colorResolved).Value = "" Then
        Range("O" & colorResolved).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorResolved).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorResolved
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub SAPSystemCorrectness()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessageString("checking SAP Area correctness ...")
    
    Dim CheckSAPSystemCorrectness As Integer: CheckSAPSystemCorrectness = 10000
    
    For C = 2 To CheckSAPSystemCorrectness
        If Not Range("H" & C).Value = "" And Not Range("F" & C).Value = "" Then
            If Not (Range("H" & C).Value = "BP2" Or _
                    Range("H" & C).Value = "ACE" Or _
                    Range("H" & C).Value = "BP5" Or _
                    Range("H" & C).Value = "HRP" Or _
                    Range("H" & C).Value = "RE-FX" Or _
                    Range("H" & C).Value = "IFRS") Then
                Range("H" & C).Select
                With Selection.Interior
                    .Color = 13260
                    .PatternTintAndShade = 0
                End With
                'Color Incident Numbers
                Range("C" & C).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next C
    
    ActiveSheet.Range("$A$1:$BH$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call StopMacroShowMessage
End Sub
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
Sub DeveloperStatusCheck()
    Dim numberOfConsultants As Integer, developers() As String
    
    numberOfConsultants = 0
    
    Sheets("ConsultantList").Visible = True
    Sheets("ConsultantList").Select
    
    Dim lRow As Long: lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Count developers
    For i = 1 To lRow
        If Sheets("ConsultantList").Range("A" & i) = "ABAP" Then
            numberOfConsultants = numberOfConsultants + 1
        End If
    Next i
    
    'Define dynamic array
    ReDim developers(numberOfConsultants)
    
    'Populate dynamic array with developers
    For i = 1 To lRow
        If Sheets("ConsultantList").Range("A" & i) = "ABAP" Then
            developers(i - 1) = Range("B" & i).Value
        End If
    Next i
           
    'Filter by statuses: "Assigned", "In Progress", "Pending", "Resvoled"
    Sheets("ConsultantList").Visible = False
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending", "Resvoled"), Operator:=xlFilterValues
                    
    Dim numberOfTickets As Long, totalNumberOfTickets As Long
    totalNumberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))
    
    'Check if developer doesn't have proper SAP Area ("Development Atos GDC" OR "Development")
    For i = 1 To numberOfConsultants
        ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=5, Criteria1:=Array(developers(i)), Operator:=xlFilterValues
        numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))
        For cellAdd = 1 To totalNumberOfTickets
            If Range("E" & cellAdd) = developers(i) And Not (Range("E" & cellAdd).Offset(0, -1).Value = "Development" Or Range("E" & cellAdd).Offset(0, -1).Value = "Development Atos GDC") Then
                Range("C" & cellAdd).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        Next cellAdd
    Next i
    
    'Check if ticket has In Progress Start Date and is not in Development
    For i = 1 To totalNumberOfTickets
        If Not Range("L" & i).Value = "" And Range("F" & i) = "Assigned" And Not (Range("D" & i).Value = "Development" Or Range("D" & i).Value = "Development Atos GDC") Then
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    
    Call BackToNormal
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
End Sub
Sub ClosedDateCheck()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 16.07.2019  **********

    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Closed", "Cancelled"), Operator:=xlFilterValues
    
    Dim numberOfTickets As Long
    numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))

    For i = 1 To numberOfTickets
        If (Range("F" & i) = "Closed" Or Range("F" & i) = "Cancelled") And (Range("N" & i) = "" Or Range("O" & i) = "") Then
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
End Sub
Sub TicketResolvingCounter()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 17.07.2019  **********

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
        Application.ScreenUpdating = False
        
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
        
        Application.Run ("'AMS_ARDAGH_DD_MACROS.xlam'!BackToNormal")
        
        Application.ScreenUpdating = True
    Next loopNumber
    
    Sheets("TicketResolving").Select
    
End Sub
