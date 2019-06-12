Sub compareStringsRfC()
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
Sub compareStringsCD()
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
Sub checkAllCharmStatuses()
    Call startMacroShowMessage(10)
    
    Sheets("Sheet1").Select
        Range("BA2:BB10000").ClearContents
    
    For rfc = 3 To 10000
        Range("AY" & rfc).Select
    Next rfc
    
    For cd = 3 To 10000
        Range("AZ" & cd).Select
    Next cd
    
    ActiveWindow.ScrollRow = 1
    
    Call stopMacroShowMessage
End Sub
Sub HideColumnsForChaRM()
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

    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=6, Criteria1:=Array( _
                                                                                "Assigned", _
                                                                                "In Progress", _
                                                                                "Pending" _
                                                                                ), Operator:=xlFilterValues
                                                                                
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=57, Criteria1:="<>Status in ChaRM cannot be changed due to upgrade (freeze)."

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
Sub LoadChaRMDataFromFilesBackend(filePath As String, number As Integer, copyToCell As String, sheetName As String, stopOnColumn As String)
        
    Call startMacroShowMessage(3)
    
    Dim wb As Workbook
        
    thisfile = Sheets("PendingCalculator").Range("Q18").Value

    trimmedFile = Mid(filePath, 28)
     
    If Dir(filePath) = "" Then
        MsgBox "Could not find the file: " & filePath
        Exit Sub
    End If
    
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
            
    Windows(trimmedFile).Activate
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
    
    Call stopMacroShowMessage
End Sub
Sub CopyDataFromRfCAndCDToChaRMSheetBackend(sheetName As String, letterOfColumn As String, copyToCell As String)
    Sheets("sheetName").Select
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
    
    Sheets("ChaRM RfC").Visible = False
    Sheets("ChaRM CD").Visible = False
End Sub
Sub LoadChaRMDataFromFiles()
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call LoadChaRMDataFromFilesBackend("C:\Users\A702387\Downloads\rfc.csv", 45, "A1", "ChaRM RfC", "Z")
        Call LoadChaRMDataFromFilesBackend("C:\Users\A702387\Downloads\cd.csv", 45, "A1", "ChaRM CD", "V")
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call LoadChaRMDataFromFilesBackend("C:\Users\A700473\Downloads\rfc.csv", 59, "A1", "ChaRM RfC", "Z")
        Call LoadChaRMDataFromFilesBackend("C:\Users\A700473\Downloads\cd.csv", 59, "A1", "ChaRM CD", "V")
    End If
    
    Call CopyDataFromRfCAndCDToChaRMSheet
    Call ScrollMaxUpAndLeft
End Sub
