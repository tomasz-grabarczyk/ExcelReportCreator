Sub RunLoadDataForDevelopers()
    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    
    Call subLoadDataForDevelopers("DevCounterPrzPaw.xls", "A2")
    Call subLoadDataForDevelopers("DevCounterMicGor.xls", "B2")
    Call subLoadDataForDevelopers("DevCounterAdrKwi.xls", "C2")
    Call subLoadDataForDevelopers("DevCounterGrzZuk.xls", "D2")
    Call subLoadDataForDevelopers("DevCounterJanZat.xls", "E2")
    Call subLoadDataForDevelopers("DevCounterPawZel.xls", "F2")
    
End Sub
Sub subLoadDataForDevelopers(filePath As String, pasteToCell As String)

    Sheets("DeveloperCounterBackend").Select
    
    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(filePath).Activate
    Columns("A:A").Select
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
    Rows("1:2").Delete Shift:=xlUp
    Range("A1:A1000").Copy
    Windows(thisfile).Activate
    Sheets("DeveloperCounterBackend").Select
    Range(pasteToCell).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(filePath).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
End Sub
Sub RemoveBlanks(copyFromColumnNumber As Variant, CopyToColumnNumber As Variant)
    counter = 0
    'Remove blanks from cells 1 to 1000
    For iterator = 2 To 1000
        If Cells(iterator, copyFromColumnNumber).Value <> "" Then
            Cells(counter + 1, CopyToColumnNumber).Value = Cells(iterator, copyFromColumnNumber).Value
            counter = counter + 1
       End If
    Next iterator
End Sub
Sub FindTicketsWithMultipleDevelopers(developerColumn As Variant, copyToColumn As String)
    For clmnNumber = 0 To Application.CountA(developerColumn) - 1
        For i = 2 To 1000
            If Not Range(developerColumn(clmnNumber) & i).Value = "" Then
                Range(copyToColumn & i).Value = Range(developerColumn(clmnNumber) & i).Value
            End If
        Next i
    Next clmnNumber
End Sub
Sub copyToColumn(copyFromColumn As Variant, sheetName As String, columnNumber As Integer, pasteToColumn As Variant)
    For j = 1 To 1000
        Sheets("DeveloperCounterBackend").Select
        If Not Range(copyFromColumn & j).Value = "" Then
            Range(copyFromColumn & j).Copy
            
            Sheets(sheetName).Select
            
            Dim lRow As Long
            lRow = Cells(Rows.Count, columnNumber).End(xlUp).Row
            
            Range(pasteToColumn & lRow + 1).PasteSpecial Paste:=xlPasteValues
            
            If copyFromColumn = "AM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("A1").Value
            ElseIf copyFromColumn = "BM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("B1").Value
            ElseIf copyFromColumn = "CM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("C1").Value
            ElseIf copyFromColumn = "DM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("D1").Value
            ElseIf copyFromColumn = "EM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("E1").Value
            ElseIf copyFromColumn = "FM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("F1").Value
            ElseIf copyFromColumn = "GM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("G1").Value
            ElseIf copyFromColumn = "HM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("H1").Value
            ElseIf copyFromColumn = "IM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("I1").Value
            ElseIf copyFromColumn = "JM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("J1").Value
            ElseIf copyFromColumn = "KM" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = Sheets("DeveloperCounterBackend").Range("K1").Value
            ElseIf copyFromColumn = "AA" And sheetName = "DeveloperCounter" Then
                Range(ActiveCell.Address).Offset(0, -1).Value = "Multiple Developers"
                Range(ActiveCell.Address).Offset(0, -1).Font.ColorIndex = 3
            End If
                
        End If
    Next j
End Sub
Sub RemoveDuplicatesFromTickets(columnLetter As Variant)
    Columns(columnLetter & ":" & columnLetter).Select
    ActiveSheet.Range(columnLetter & "1:" & columnLetter & "1000").RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

