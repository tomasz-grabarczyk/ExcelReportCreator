Sub RunTicketChecker()
    Call startMacroShowMessage(10)

    Worksheets("NewChecker").Visible = True
    
    Sheets("NewChecker").Select
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
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=6
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=5
    Sheets("NewChecker").Select
    Range("AP1").Select
    
    Dim a_newCheckerColumns As Variant
    a_newCheckerColumns = Array("AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX")
    
    'Call addHyperlinks(3, 2000, a_newCheckerColumns)
    End If
    
    Sheets("NewChecker").Visible = True
    
    For i = 3 To 1000
        Sheets("NewChecker").Select
        If Not Range("AQ" & i).Value = "" Then
            Sheets("Sheet1").Select
            Call addNewRowToReport
            Sheets("NewChecker").Select
            Range("AQ" & i).Copy
            Sheets("Sheet1").Select
            Range("C2").PasteSpecial xlPasteValues
            Sheets("NewChecker").Select
            Range("AQ" & i).ClearContents
            Sheets("Sheet1").Select
        End If
    Next i
    Sheets("Sheet1").Select
    
    Sheets("NewChecker").Visible = False

    Call stopMacroShowMessage
End Sub

Sub SLACheckLayout()

    Call backToNormal
    
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
Sub CheckSAPAreaCorrectness()
    Call startMacroShowMessageString("checking SAP Area correctness ...")
    
    Dim CheckSAPAreaCorrectness As Integer: CheckSAPAreaCorrectness = 10000
    
    For c = 2 To CheckSAPAreaCorrectness
        If Not Range("H" & c).Value = "" And Not Range("F" & c).Value = "" Then
            If Not (Range("H" & c).Value = "BP2" Or _
                    Range("H" & c).Value = "ACE" Or _
                    Range("H" & c).Value = "BP5" Or _
                    Range("H" & c).Value = "HRP" Or _
                    Range("H" & c).Value = "RE-FX" Or _
                    Range("H" & c).Value = "IFRS") Then
                Range("H" & c).Select
                With Selection.Interior
                    .Color = 13260
                    .PatternTintAndShade = 0
                End With
                'Color Incident Numbers
                Range("C" & c).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next c
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call stopMacroShowMessage
End Sub
