Sub LaunchCheckerFileUploadBackend(filePath As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessageString("Adding new tickets ...")
    
    Worksheets("NewCheckerUpdated").Visible = True
    Sheets("NewCheckerUpdated").Select
    
    Range("A3:D10000").ClearContents
    Range("J3:K10000").ClearContents
    
    Dim wb As Workbook
    
    'Take name of the file from cell Q18 from sheet PendingCalculator
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(filePath).Activate
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Windows(thisfile).Activate
    Sheets("NewCheckerUpdated").Select
    Range("C3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows(filePath).Activate
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Windows(thisfile).Activate
    Sheets("NewCheckerUpdated").Select
    Range("D3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
    Windows(filePath).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    'Delete file from folder
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
    
    Call CompareAllTickets
    
    Dim lastRowStatus As Long: lastRowStatus = Cells(Rows.Count, 11).End(xlUp).Row
    
    'Compare if any tickets needs to be changed to Closed
    For i = 3 To lastRowStatus
        Sheets("NewCheckerUpdated").Select
        Range("K" & i).Select
    Next i
    
    'Check if there are any tickets to be added to report
    
    Dim lastRowNewTicket As Long: lastRowNewTicket = Cells(Rows.Count, 10).End(xlUp).Row
    
    For addNewTicketToReport = 3 To lastRowNewTicket
        If Not Range("J" & addNewTicketToReport).Value = "" Then
            Call AddNewTicket
            Sheets("NewCheckerUpdated").Select
            Range("J" & addNewTicketToReport).Copy
            Sheets("Sheet1").Select
            Range("C2").PasteSpecial xlPasteValues
            Sheets("NewCheckerUpdated").Select
        End If
    Next addNewTicketToReport
    
    Sheets("NewCheckerUpdated").Visible = False

    Sheets("Sheet1").Select
    
    Call StopMacroShowMessage
End Sub
Sub CompareAllTickets()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    'Consultants to filter by:
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$AN$10000").AutoFilter Field:=5, Criteria1:="<>N/A"
        
    'Statuses to filter by:
    ActiveSheet.Range("$A$1:$AR$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", "In Progress", "Pending", "Resolved", "="), Operator:=xlFilterValues
       
    'Copy Ticket Number (Sheet1) to Excel Report RESULT -> Incident Number:
    Sheets("Sheet1").Select
    Range("C2:C10000").Copy
    Sheets("NewCheckerUpdated").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copy Status (Sheet1) to Excel Report RESULT -> Status:
    Sheets("Sheet1").Select
    Range("F2:F10000").Copy
    Sheets("NewCheckerUpdated").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Delete every space in data from Remedy and Excel:
    Worksheets("NewCheckerUpdated").Columns("A:D").Replace _
        What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=True
    
    'Delete blank rows from Remedy in Excel NOT FOUND -> Incident ID*+:
    Dim RemedyInExcelCounter As Integer, RemedyInExcelIterator As Integer
    RemedyInExcelCounter = 0
    For RemedyInExcelIterator = 1 To 10000
        If Cells(RemedyInExcelIterator, 6).Value <> "" Then
            Cells(RemedyInExcelCounter + 1, 10).Value = Cells(RemedyInExcelIterator, 6).Value
            RemedyInExcelCounter = RemedyInExcelCounter + 1
        End If
    Next RemedyInExcelIterator
    
    'Delete blank rows from Excel Report RESULT -> Status*:
    Dim StatusCounter As Integer, StatusIterator As Integer
    StatusCounter = 0
    For StatusIterator = 1 To 10000
        If Cells(StatusIterator, 9).Value <> "" Then
            Cells(StatusCounter + 1, 11).Value = Cells(StatusIterator, 9).Value
            StatusCounter = StatusCounter + 1
        End If
    Next StatusIterator
    
    Call BackToNormal
    
    Sheets("NewCheckerUpdated").Select
    
End Sub
Sub RunVLookUpsUpdatedReport(filePath As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    If Dir(filePath) = "" Then
            MsgBox "Could not find the file: " & filePath
            Exit Sub
    End If
    
    Call StartMacroShowMessage(9)
    
    Sheets("VLookUps").Visible = True
    Sheets("VLookUps").Select

    Range("A1:K10000").ClearContents
    
    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(filePath).Activate

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
    
    Windows(filePath).Activate
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
    
    Call LaunchCheckerFileUploadBackend(filePath)
    
    Call StopMacroShowMessage
End Sub
