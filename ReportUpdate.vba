Sub addNewRowToReport()
    Rows("2:2").Select
    Selection.Copy
    Rows("3:3").Insert Shift:=xlDown
    Range("A2").Select
    Application.CutCopyMode = False
    Range("C2").ClearContents
    Range("K2").ClearContents
End Sub
Sub removeFirstRowFromReport()
    Rows("2:2").Delete Shift:=xlUp
End Sub
Sub CheckDates()
    Call startMacroShowMessage(2)
        
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:BG").Select
    Selection.EntireColumn.Hidden = True
    
    Dim checkDatesToCell As Integer: checkDatesToCell = 10000
    
    For s = 2 To checkDatesToCell
        If (Range("F" & s).Value = "Assigned" Or Range("F" & s).Value = "In Progress" Or Range("F" & s).Value = "Pending" Or Range("F" & s).Value = "Resolved") And Range("K" & s).Value = "" Then
            Range("K" & s).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & s).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next s
    
    For i = 2 To checkDatesToCell
        If (Range("F" & i).Value = "In Progress" Or Range("F" & i).Value = "Pending" Or Range("F" & i).Value = "Resolved") And Range("L" & i).Value = "" Then
            Range("L" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    For j = 2 To checkDatesToCell
        If Range("F" & j).Value = "Pending" And Range("M" & j).Value = "" Then
            Range("M" & j).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & j).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        ElseIf Range("F" & j).Value = "Resolved" And Range("N" & j).Value = "" Then
            Range("N" & j).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & j).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next j
    
    For k = 2 To checkDatesToCell
        If Range("F" & k).Value = "Resolved" And Range("O" & k).Value = "" Then
        Range("O" & k).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & k).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next k
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call stopMacroShowMessage
End Sub
Sub launchReportFile()
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call createReportFile("C:\Users\A702387\Downloads\Atos20DD202d20Ardagh202d20Reports.xls", 45)
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call createReportFile("C:\Users\A700473\Downloads\Atos20DD202d20Ardagh202d20Reports A.xls", 59)
    End If
End Sub
Sub createReportFile(filePath As String, number As Integer)
    Call startMacroShowMessageString("adding new tickets")
    
    Worksheets("NewChecker").Visible = True
    Sheets("NewChecker").Select
    Call clearNewChecker
    
    
    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    trimmedFile = Mid(filePath, 28)
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(trimmedFile).Activate
    Range("A1:H1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Windows(thisfile).Activate
    Sheets("NewChecker").Select
    Range("I3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Call CheckAllTickets
    
    Dim i As Integer
    
    For i = 3 To 100
        Sheets("NewChecker").Select
        Range("AT" & i).Select
    Next i
    
    ActiveWindow.ScrollRow = 1
    
    Windows(trimmedFile).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
    
    'Check if there are any tickets to be added to report
    Sheets("NewChecker").Visible = True
        If Not Range("AQ3").Value = "" Then
            Call RunTicketChecker
        End If
    Sheets("NewChecker").Visible = False
    
    Worksheets("NewChecker").Visible = False
    Sheets("Sheet1").Select
    
    Call stopMacroShowMessage
End Sub
Sub launchVLookUps()
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call runVLookUps("C:\Users\A702387\Downloads\Atos20DD202d20Ardagh202d20Reports.xls", 45)
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call runVLookUps("C:\Users\A700473\Downloads\Atos20DD202d20Ardagh202d20Reports A.xls", 59)
    End If
End Sub
Sub runVLookUps(filePath As String, number As Integer)
    Call startMacroShowMessage(9)
    
    Sheets("VLookUps").Visible = True
    Sheets("VLookUps").Select
    
    Range("A1:K10000").ClearContents
    
    
    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    trimmedFile = Mid(filePath, 28)
    
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(trimmedFile).Activate
    Columns("A:K").Select
    ActiveWorkbook.Worksheets("Sheet 1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet 1").Sort.SortFields.Add2 Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet 1").Sort
        .SetRange Range("A2:K10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("1:3").Delete Shift:=xlUp
    Range("A1:K1").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Windows(thisfile).Activate
    Sheets("VLookUps").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(trimmedFile).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    Columns("A:A").Cut
    Columns("C:C").Insert Shift:=xlToRight
    
    Columns("A:K").EntireColumn.AutoFit
    
    'Replace "  " to " " in: Data from Remedy -> Assignee+:
    Worksheets("VLookUps").Columns("C").Replace _
        What:="  ", Replacement:=" ", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Replace "ł" to "l" in: Data from Remedy -> Assignee+:
    Worksheets("VLookUps").Columns("C").Replace _
        What:="ł", Replacement:="l", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Replace "FICO" to "Fico" in: data from Remedy Model/Version -> Excel Report RESULT Model/Version:
    Worksheets("VLookUps").Columns("G").Replace _
        What:="FICO", Replacement:="Fico", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    Sheets("VLookUps").Visible = False
    Sheets("Sheet1").Select
    Range("A1").Select
    
    Call launchReportFile
    
    Call stopMacroShowMessage
End Sub
Sub CheckAllTickets()
    Worksheets("NewChecker").Visible = True
    
    'Clear columns AP to AX - result of query without blank spaces:
    Range("AP3:AX10000").Select
    Selection.ClearContents
    Range("AP1").Select
    
    'Check if user entered data from Remedy:
    Set myCell = ThisWorkbook.Worksheets("NewChecker").Range("I3")
    If IsEmpty(myCell) Then
        MsgBox "Please enter ticket numbers from Remedy!"
        Range("I3").Select
    Else
    Call backToNormal
    
    'Consultants to filter by:
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AN$10000").AutoFilter Field:=5, Criteria1:="<>N/A"
        
    'Statuses to filter by:
    ActiveSheet.Range("$A$1:$AR$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", "In Progress", "Pending", "Resolved", "="), Operator:=xlFilterValues
        
    'Range variables to check:
    Dim TicketType, incidentNumber, sapArea, consultant, status, StatusReason, Priority, Summary As Range
    
    'Set range to check:
    Set TicketType = Range("B2:B10000")
    Set incidentNumber = Range("C2:C10000")
    Set sapArea = Range("D2:D10000")
    Set consultant = Range("E2:E10000")
    Set status = Range("F2:F10000")
    Set StatusReason = Range("G2:G10000")
    Set Priority = Range("J2:J10000")
    Set Summary = Range("AE2:AE10000")
    
    'Copy Ticket Type (Sheet1) to Excel Report RESULT -> Ticket Type:
    Sheets("Sheet1").Select
    TicketType.Copy
    Sheets("NewChecker").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Copy Ticket Number (Sheet1) to Excel Report RESULT -> Incident Number:
    Sheets("NewChecker").Select
    incidentNumber.Copy
    Sheets("NewChecker").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy SAP Area (Sheet1) to Excel Report RESULT -> SAP Area:
    Sheets("NewChecker").Select
    sapArea.Copy
    Sheets("NewChecker").Select
    Range("G3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy Consultant (Sheet1) to Excel Report RESULT -> Consultant:
    Sheets("NewChecker").Select
    consultant.Copy
    Sheets("NewChecker").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy Status (Sheet1) to Excel Report RESULT -> Status:
    Sheets("NewChecker").Select
    status.Copy
    Sheets("NewChecker").Select
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy Status Reason (Sheet1) to Excel Report RESULT -> Status Reason:
    Sheets("NewChecker").Select
    StatusReason.Copy
    Sheets("NewChecker").Select
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Copy Priority (Sheet1) to Excel Report RESULT -> Priority:
    Sheets("NewChecker").Select
    Priority.Copy
    Sheets("NewChecker").Select
    Range("F3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy Summary (Sheet1) to Excel Report RESULT -> Summary:
    Sheets("NewChecker").Select
    Summary.Copy
    Sheets("NewChecker").Select
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Replace "  " to " " in: Data from Remedy -> Assignee+:
    Worksheets("NewChecker").Columns("K").Replace _
        What:="  ", Replacement:=" ", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Replace "ł" to "l" in: Data from Remedy -> Assignee+:
    Worksheets("NewChecker").Columns("K").Replace _
        What:="ł", Replacement:="l", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Replace "FICO" to "Fico" in: data from Remedy Model/Version -> Excel Report RESULT Model/Version:
    Worksheets("NewChecker").Columns("O").Replace _
        What:="FICO", Replacement:="Fico", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Delete every space in data from Remedy and Excel:
    Worksheets("NewChecker").Columns("A:P").Replace _
        What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Delete black rows from Excel in Remedy NOT FOUND -> Incident ID*+:
    Dim ExcelInRemedyCounter As Integer, ExcelInRemedyIterator As Integer
    ExcelInRemedyCounter = 0
    For ExcelInRemedyIterator = 1 To 10000
        If Cells(ExcelInRemedyIterator, 19).Value <> "" Then
            Cells(ExcelInRemedyCounter + 1, 42).Value = Cells(ExcelInRemedyIterator, 19).Value
            ExcelInRemedyCounter = ExcelInRemedyCounter + 1
       End If
    Next ExcelInRemedyIterator
    
    'Delete black rows from Remedy in Excel NOT FOUND -> Incident ID*+:
    Dim RemedyInExcelCounter As Integer, RemedyInExcelIterator As Integer
    RemedyInExcelCounter = 0
    For RemedyInExcelIterator = 1 To 10000
        If Cells(RemedyInExcelIterator, 20).Value <> "" Then
            Cells(RemedyInExcelCounter + 1, 43).Value = Cells(RemedyInExcelIterator, 20).Value
            RemedyInExcelCounter = RemedyInExcelCounter + 1
        End If
    Next RemedyInExcelIterator
    
    'Delete black rows from Excel Report RESULT -> Incident Type*:
    Dim IncidentTypeCounter As Integer, IncidentTypeIterator As Integer
    IncidentTypeCounter = 0
    For IncidentTypeIterator = 1 To 10000
        If Cells(IncidentTypeIterator, 35).Value <> "" Then
            Cells(IncidentTypeCounter + 1, 44).Value = Cells(IncidentTypeIterator, 35).Value
            IncidentTypeCounter = IncidentTypeCounter + 1
        End If
    Next IncidentTypeIterator
    
    'Delete black rows from Excel Report RESULT -> Assignee+:
    Dim AssigneeCounter As Integer, AssigneeIterator As Integer
    AssigneeCounter = 0
    For AssigneeIterator = 1 To 10000
        If Cells(AssigneeIterator, 36).Value <> "" Then
            Cells(AssigneeCounter + 1, 45).Value = Cells(AssigneeIterator, 36).Value
            AssigneeCounter = AssigneeCounter + 1
        End If
    Next AssigneeIterator
    
    'Delete black rows from Excel Report RESULT -> Status*:
    Dim StatusCounter As Integer, StatusIterator As Integer
    StatusCounter = 0
    For StatusIterator = 1 To 10000
        If Cells(StatusIterator, 37).Value <> "" Then
            Cells(StatusCounter + 1, 46).Value = Cells(StatusIterator, 37).Value
            StatusCounter = StatusCounter + 1
        End If
    Next StatusIterator
    
    'Delete black rows from Excel Report RESULT -> Status_Reason_Hidden:
    Dim StatusReasonCounter As Integer, StatusReasonIterator As Integer
    StatusReasonCounter = 0
    For StatusReasonIterator = 1 To 10000
        If Cells(StatusReasonIterator, 38).Value <> "" Then
            Cells(StatusReasonCounter + 1, 47).Value = Cells(StatusReasonIterator, 38).Value
            StatusReasonCounter = StatusReasonCounter + 1
       End If
    Next StatusReasonIterator
    
    'Delete black rows from Excel Report RESULT -> Priority*:
    Dim PriorityCounter As Integer, PriorityIterator As Integer
    PriorityCounter = 0
    For PriorityIterator = 1 To 10000
        If Cells(PriorityIterator, 39).Value <> "" Then
            Cells(PriorityCounter + 1, 48).Value = Cells(PriorityIterator, 39).Value
            PriorityCounter = PriorityCounter + 1
        End If
    Next PriorityIterator
    
    'Delete black rows from Excel Report RESULT -> Model/Version:
    Dim ModelCounter As Integer, ModelIterator As Integer
    ModelCounter = 0
    For ModelIterator = 1 To 10000
        If Cells(ModelIterator, 40).Value <> "" Then
            Cells(ModelCounter + 1, 49).Value = Cells(ModelIterator, 40).Value
            ModelCounter = ModelCounter + 1
        End If
    Next ModelIterator
    
    'Delete black rows from Excel Report RESULT -> Summary*:
    Dim SummaryCounter As Integer, SummaryIterator As Integer
    SummaryCounter = 0
    For SummaryIterator = 1 To 10000
        If Cells(SummaryIterator, 41).Value <> "" Then
            Cells(SummaryCounter + 1, 50).Value = Cells(SummaryIterator, 41).Value
            SummaryCounter = SummaryCounter + 1
        End If
    Next SummaryIterator
    
    'Unhide hidden columns and rows in Sheet1:
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AP$1293").AutoFilter Field:=6
    ActiveSheet.Range("$A$1:$AP$1293").AutoFilter Field:=5
    Sheets("NewChecker").Select
    Range("AP1").Select
    
    Dim a_newCheckerColumns As Variant
    a_newCheckerColumns = Array("AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX")
    
    'Call addHyperlinks(3, 2000, a_newCheckerColumns)
    End If
End Sub
Sub clearNewChecker()
    'Clear Data from Excel/Data from Remedy/Ticket numbers without blank spaces:
    Range("A3:P10000").ClearContents
    Range("AP3:AX10000").ClearContents
    
    Range("AP3:AX10000").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Underline = xlUnderlineStyleNone
End Sub
Sub hideNewChecker()
    Worksheets("NewChecker").Visible = False
    Sheets("Sheet1").Select
End Sub
