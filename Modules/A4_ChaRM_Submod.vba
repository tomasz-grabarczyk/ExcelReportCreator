Sub CompareStringsRfC()
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
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call runCharmStatuses("C:\Users\A702387\Downloads\export.csv", 45, cellRange, copyToCell, columnsToBeDeleted)
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call runCharmStatuses("C:\Users\A700473\Downloads\export.csv", 59, cellRange, copyToCell, columnsToBeDeleted)
    End If
End Sub
Sub convertColumnToNumbers(convertColumn As String, cellRange As String, startFromCell As String)
 
    Dim SelectR As Range
    Dim sht As Worksheet
    Dim LastRow As Long
    
    Set sht = Sheets("ChaRM")
    
    LastRow = sht.Cells(sht.Rows.Count, convertColumn).End(xlUp).Row
    
    Set SelectR = Sheets("ChaRM").Range(cellRange & LastRow)
    
    SelectR.TextToColumns Destination:=Range(startFromCell), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
End Sub
Sub LoadChaRMDataFromFilesBackend(filePath As String, copyToCell As String, sheetName As String, stopOnColumn As String)
    
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
    
    If sheetName = "ChaRM RfC" Then
        For i = 0 To 1000
            Dim lastBlankRowRfC As Long
            lastBlankRowRfC = Cells(Rows.Count, 27).End(xlUp).Row
            If Not Range("A" & lastBlankRowRfC + 1) = "" Then
                Range("AA2:AD2").Copy
                Range("AA" & lastBlankRowRfC + 1).PasteSpecial xlPasteFormulas
            End If
        Next i
    End If
    
    If sheetName = "ChaRM CD" Then
        For i = 0 To 1000
            Dim lastBlankRowCD As Long
            lastBlankRowCD = Cells(Rows.Count, 23).End(xlUp).Row
            If Not Range("A" & lastBlankRowCD + 1) = "" Then
                Range("W2:Y2").Copy
                Range("W" & lastBlankRowCD + 1).PasteSpecial xlPasteFormulas
            End If
        Next i
    End If
    
    Sheets(sheetName).Visible = False
End Sub
Sub CopyDataFromRfCAndCDToChaRMSheetBackend(sheetName As String, letterOfColumn As String, copyToCell As String)
    Sheets(sheetName).Select
    Range(letterOfColumn & "2:" & letterOfColumn & "1000").Copy
    Sheets("ChaRM").Select
    Range(copyToCell).PasteSpecial xlPasteValues
End Sub
Sub CopyDataFromRfCAndCDToChaRMSheet()
    Sheets("ChaRM").Visible = True
    Sheets("ChaRM").Select
    Range("A2:F1000").ClearContents
    
    Sheets("ChaRM RfC").Visible = True
    Sheets("ChaRM CD").Visible = True
    
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM RfC", "T", "A2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM RfC", "E", "B2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM RfC", "U", "C2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM CD", "O", "D2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM CD", "I", "E2")
    Call CopyDataFromRfCAndCDToChaRMSheetBackend("ChaRM CD", "Q", "F2")
    
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
    
    Sheets("ChaRM RfC").Visible = False
    Sheets("ChaRM CD").Visible = False
End Sub
Sub RemoveMultipleOccurencesOfTickets()
    Dim RfCNumberOfTickets As Long: RfCNumberOfTickets = Cells(Rows.Count, 1).End(xlUp).Row
    Dim CDNumberOfTickets As Long: CDNumberOfTickets = Cells(Rows.Count, 4).End(xlUp).Row
    
    Dim RfCMultipleTicketNo As Long: RfCMultipleTicketNo = Cells(Rows.Count, 9).End(xlUp).Row
    Dim CDMultipleTicketNo As Long: CDMultipleTicketNo = Cells(Rows.Count, 11).End(xlUp).Row
    
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

