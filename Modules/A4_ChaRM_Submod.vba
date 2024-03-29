Sub HideColumnsForChaRM()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Columns("A:B").Select
        Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
        Selection.EntireColumn.Hidden = True
    Columns("G:AX").Select
        Selection.EntireColumn.Hidden = True
    Columns("BC:BD").Select
        Selection.EntireColumn.Hidden = True
    Columns("BF:BG").Select
        Selection.EntireColumn.Hidden = True

    ActiveSheet.Range("$A$1:$BH$10000").AutoFilter Field:=6, Criteria1:=Array( _
                                                                                "Assigned", _
                                                                                "In Progress", _
                                                                                "Pending" _
                                                                                ), Operator:=xlFilterValues
    Dim ticketNumbersToColor As Long: ticketNumbersToColor = Cells(Rows.count, 3).End(xlUp).Row
    
    For colorIterator = 2 To ticketNumbersToColor
        If Not Range("BA" & colorIterator).Value = "" Or Not Range("BB" & colorIterator).Value = "" Then
            Range("C" & colorIterator).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorIterator
    
    Call FilterIncidentNumbersByColor
    
End Sub
Sub CompareStringsRfC()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    If (ActiveCell.Value = "Created" Or ActiveCell.Value = "In Preparation" Or ActiveCell.Value = "Tech. Specification Request") Then
        If Not (ActiveCell.Offset(0, -45).Value = "Assigned" Or ActiveCell.Offset(0, -45).Value = "In Progress") Then
            ActiveCell.Offset(0, 2).Value = "In Progress"
        End If
    End If
    
    If (ActiveCell.Value = "Business Lead To Sign Off" Or ActiveCell.Value = "IT Bus. Analyst To Sign Off" Or ActiveCell.Value = "To be approved by IT Owner" Or ActiveCell.Value = "To be planned") Then
        If Not (ActiveCell.Offset(0, -45).Value = "Pending") Then
            ActiveCell.Offset(0, 2).Value = "Pending"
        End If
    End If
    
    If (ActiveCell.Value = "Implemented") Then
        If Not (ActiveCell.Offset(0, -45).Value = "Resolved" Or ActiveCell.Offset(0, -45).Value = "Closed") Then
            ActiveCell.Offset(0, 2).Value = "Resolved"
        End If
    End If
    
    If (ActiveCell.Value = "Rejected") Then
        If Not (ActiveCell.Offset(0, -45).Value = "Cancelled") Then
            ActiveCell.Offset(0, 2).Value = "Cancelled"
        End If
    End If
End Sub
Sub CompareStringsCD()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    If (ActiveCell.Value = "Created" Or ActiveCell.Value = "In development" Or ActiveCell.Value = "To be tested in PreProd") Then
        If Not (ActiveCell.Offset(0, -46).Value = "Assigned" Or ActiveCell.Offset(0, -46).Value = "In Progress") Then
            ActiveCell.Offset(0, 2).Value = "In Progress"
        End If
    End If
    
    If (ActiveCell.Value = "To be tested in UAT" Or ActiveCell.Value = "To be confirmed in Prod" Or ActiveCell.Value = "To be imported into Prod") Then
        If Not (ActiveCell.Offset(0, -46).Value = "Pending") Then
            ActiveCell.Offset(0, 2).Value = "Pending"
        End If
    End If
    
    If (ActiveCell.Value = "Completed") Then
        If Not (ActiveCell.Offset(0, -46).Value = "Resolved" Or ActiveCell.Offset(0, -46).Value = "Closed") Then
            ActiveCell.Offset(0, 2).Value = "Resolved"
        End If
    End If
    
    If (ActiveCell.Value = "Withdrawn") Then
        If Not (ActiveCell.Offset(0, -46).Value = "Cancelled") Then
            ActiveCell.Offset(0, 2).Value = "Cancelled"
        End If
    End If
End Sub
Sub launchCharmStatusesCheck(cellRange As String, copyToCell As String, columnsToBeDeleted As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call runCharmStatuses("C:\Users\A702387\Downloads\export.csv", 45, cellRange, copyToCell, columnsToBeDeleted)
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call runCharmStatuses("C:\Users\A700473\Downloads\export.csv", 59, cellRange, copyToCell, columnsToBeDeleted)
    End If
End Sub
Sub convertColumnToNumbers(convertColumn As String, cellRange As String, startFromCell As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
 
    Dim SelectR As Range
    Dim sht As Worksheet
    Dim LastRow As Long
    
    Set sht = Sheets("ChaRM")
    
    LastRow = sht.Cells(sht.Rows.count, convertColumn).End(xlUp).Row
    
    Set SelectR = Sheets("ChaRM").Range(cellRange & LastRow)
    
    SelectR.TextToColumns Destination:=Range(startFromCell), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
End Sub
Sub LoadChaRMDataFromFilesBackend(filePath As String, copyToCell As String, sheetName As String, stopOnColumn As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    If Dir(filePath) = "" Then
        MsgBox "Could not find the file: " & filePath
        Exit Sub
    End If
    
    Dim wb As Workbook
        
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    Sheets(sheetName).Visible = True
    Sheets(sheetName).Select
    Columns("A:" & stopOnColumn).ClearContents
    
    Set wb = Workbooks.Open(filePath)

    Range("A1:" & stopOnColumn & "10000").Copy
    Windows(thisfile).Activate
    Sheets(sheetName).Select
    Range(copyToCell).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            
    Columns("A:" & stopOnColumn).EntireColumn.AutoFit
    
    Range("A1").Select
            
    Windows(filePath).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
 
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
    
    Sheets(sheetName).Select
    
    If sheetName = "RfC" Then
        For i = 0 To 1000
            Dim lastBlankRowRfC As Long
            lastBlankRowRfC = Cells(Rows.count, 27).End(xlUp).Row
            If Not Range("A" & lastBlankRowRfC + 1) = "" Then
                Range("AA2:AD2").Copy
                Range("AA" & lastBlankRowRfC + 1).PasteSpecial xlPasteFormulas
            End If
        Next i
    End If
    
    If sheetName = "CD" Then
        For i = 0 To 1000
            Dim lastBlankRowCD As Long
            lastBlankRowCD = Cells(Rows.count, 23).End(xlUp).Row
            If Not Range("A" & lastBlankRowCD + 1) = "" Then
                Range("W2:Y2").Copy
                Range("W" & lastBlankRowCD + 1).PasteSpecial xlPasteFormulas
            End If
        Next i
    End If
    
    Sheets(sheetName).Visible = False
End Sub
Sub CopyDataFromRfCAndCDToChaRMSheetBackend(sheetName As String, letterOfColumn As String, copyToCell As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
        
    Sheets(sheetName).Select
    Range(letterOfColumn & "2:" & letterOfColumn & "1000").Copy
    Sheets("ChaRM").Select
    Range(copyToCell).PasteSpecial xlPasteValues
End Sub
Sub CopyDataFromRfCAndCDToChaRMSheet()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Sheets("ChaRM").Visible = True
    Sheets("ChaRM").Select
    Range("A2:F1000").ClearContents
    
    Sheets("RfC").Visible = True
    Sheets("CD").Visible = True
    
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("RfC", "T", "A2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("RfC", "E", "B2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("RfC", "U", "C2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("CD", "O", "D2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("CD", "I", "E2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("CD", "Q", "F2")
    
    Call convertColumnToNumbers("C", "C2:C", "C2")
    Call convertColumnToNumbers("F", "F2:F", "F2")
    
    Columns("G:H").Select
    Selection.EntireColumn.Hidden = False
    
    Range("G2:G1000").Copy
    Range("A2").PasteSpecial xlPasteValues
    
    Range("H2:H1000").Copy
    Range("D2").PasteSpecial xlPasteValues
    
    Columns("G:H").Select
    Selection.EntireColumn.Hidden = True
    
    Sheets("RfC").Visible = False
    Sheets("CD").Visible = False
End Sub
Sub RemoveMultipleOccurencesOfTickets()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Dim RfCNumberOfTickets As Long: RfCNumberOfTickets = Cells(Rows.count, 1).End(xlUp).Row
    Dim CDNumberOfTickets As Long: CDNumberOfTickets = Cells(Rows.count, 4).End(xlUp).Row
    
    Dim RfCMultipleTicketNo As Long: RfCMultipleTicketNo = Cells(Rows.count, 9).End(xlUp).Row
    Dim CDMultipleTicketNo As Long: CDMultipleTicketNo = Cells(Rows.count, 11).End(xlUp).Row
    
    For RfCMultipleIterator = 2 To RfCMultipleTicketNo
        For RfCNumberOfTicketsIterator = 2 To RfCNumberOfTickets
            If Range("A" & RfCNumberOfTicketsIterator).Value = Range("I" & RfCMultipleIterator).Value And Range("B" & RfCNumberOfTicketsIterator).Value = Range("J" & RfCMultipleIterator).Value Then
                Range("A" & RfCNumberOfTicketsIterator & ":C" & RfCNumberOfTicketsIterator).ClearContents
            End If
        Next RfCNumberOfTicketsIterator
    Next RfCMultipleIterator
        
    For CDMultipleIterator = 2 To CDMultipleTicketNo
        For CDNumberOfTicketsIterator = 2 To CDNumberOfTickets
            If Range("D" & CDNumberOfTicketsIterator).Value = Range("K" & CDMultipleIterator).Value And Range("E" & CDNumberOfTicketsIterator).Value = Range("L" & CDMultipleIterator).Value Then
                Range("D" & CDNumberOfTicketsIterator & ":F" & CDNumberOfTicketsIterator).ClearContents
            End If
        Next CDNumberOfTicketsIterator
    Next CDMultipleIterator
    
    Sheets("ChaRM").Visible = False
End Sub
