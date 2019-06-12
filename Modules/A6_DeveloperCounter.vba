Sub unhideDeveloperChecker()
    Sheets("DeveloperCounter").Visible = True
    Sheets("DeveloperCounterBackend").Visible = True
End Sub
Sub hideDeveloperCheckers()
    Sheets("DeveloperCounter").Visible = False
    Sheets("DeveloperCounterBackend").Visible = False
End Sub
Sub RunLoadDataForDevelopers()
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterPrzPaw.xls", "A2")
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterMicGor.xls", "B2")
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterAdrKwi.xls", "C2")
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterGrzZuk.xls", "D2")
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterJanZat.xls", "E2")
        Call launchLoadDataForDevelopers("C:\Users\A702387\Downloads\TGDevCounterPawZel.xls", "F2")
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterPrzPaw.xls", "A2")
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterMicGor.xls", "B2")
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterAdrKwi.xls", "C2")
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterGrzZuk.xls", "D2")
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterJanZat.xls", "E2")
        Call launchLoadDataForDevelopers("C:\Users\A700473\Downloads\ARDevCounterPawZel.xls", "F2")
    End If
End Sub
Sub launchLoadDataForDevelopers(filePath As String, pasteToCell As String)
    If (Worksheets("PendingCalculator").Range("Q16") = "Tomasz Grabarczyk") Then
        Call subLoadDataForDevelopers(filePath, 45, pasteToCell)
    ElseIf (Worksheets("PendingCalculator").Range("Q16") = "Adam Rusnak") Then
        Call subLoadDataForDevelopers(filePath, 59, pasteToCell)
    End If
End Sub
Sub subLoadDataForDevelopers(filePath As String, number As Integer, pasteToCell As String)

    Sheets("DeveloperCounterBackend").Select
    
    Dim wb As Workbook
    
    thisfile = Sheets("PendingCalculator").Range("Q18").Value

    trimmedFile = Mid(filePath, 28)
    
    Set wb = Workbooks.Open(filePath)
    
    Windows(trimmedFile).Activate
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
    
    Windows(trimmedFile).Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    If Len(Dir$(filePath)) > 0 Then
        Kill filePath
    End If
End Sub
Sub RemoveBlanks(copyFromColumnNumber As Variant, CopyToColumnNumber As Variant)
    Counter = 0
    'Remove blanks from cells 1 to 1000
    For iterator = 2 To 1000
        If Cells(iterator, copyFromColumnNumber).Value <> "" Then
            Cells(Counter + 1, CopyToColumnNumber).Value = Cells(iterator, copyFromColumnNumber).Value
            Counter = Counter + 1
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
Sub RunDeveloperChecker()
    Call startMacroShowMessage(35)
    
    Sheets("DeveloperCounter").Select
    Range("A2:B1000").ClearContents
    
    Sheets("DeveloperCounterBackend").Select
    Range("A2:K1000").ClearContents
    
    Call RunLoadDataForDevelopers
    
    Dim columnLettersToRemoveDuplicates: columnLettersToRemoveDuplicates = Array("A", _
                                                                                 "B", _
                                                                                 "C", _
                                                                                 "D", _
                                                                                 "E", _
                                                                                 "F", _
                                                                                 "G", _
                                                                                 "H", _
                                                                                 "I", _
                                                                                 "J", _
                                                                                 "K")
                                                                                 
    For RemoveDuplicates = 0 To Application.CountA(columnLettersToRemoveDuplicates) - 1
        Call RemoveDuplicatesFromTickets(columnLettersToRemoveDuplicates(RemoveDuplicates))
    Next RemoveDuplicates

    Dim columnNumbersToRemoveBlanks: columnNumbersToRemoveBlanks = Array(38, 39, _
                                                                         64, 65, _
                                                                         90, 91, _
                                                                         116, 117, _
                                                                         142, 143, _
                                                                         168, 169, _
                                                                         194, 195, _
                                                                         220, 221, _
                                                                         246, 247, _
                                                                         272, 273, _
                                                                         298, 299)
                                                                         
    'Clear columns with tickets without blank spaces
    For clearColumn = 1 To Application.CountA(columnNumbersToRemoveBlanks) - 1 Step 2
        Sheets("DeveloperCounterBackend").Columns(columnNumbersToRemoveBlanks(clearColumn)).ClearContents
    Next clearColumn
    
    Sheets("DeveloperCounterBackend").Columns(27).ClearContents
    
    'Remove blanks from final developer tickets and paste them to column on right
    For removeBlanksFromColumn = 0 To Application.CountA(columnNumbersToRemoveBlanks) - 1 Step 2
        Call RemoveBlanks(columnNumbersToRemoveBlanks(removeBlanksFromColumn), columnNumbersToRemoveBlanks(removeBlanksFromColumn + 1))
    Next removeBlanksFromColumn
    
    Dim clmnNDev02: clmnNDev02 = Array("BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK")
    Dim clmnNDev03: clmnNDev03 = Array("CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK")
    Dim clmnNDev04: clmnNDev04 = Array("DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK")
    Dim clmnNDev05: clmnNDev05 = Array("EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK")
    Dim clmnNDev06: clmnNDev06 = Array("FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK")
    Dim clmnNDev07: clmnNDev07 = Array("GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK")
    Dim clmnNDev08: clmnNDev08 = Array("HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK")
    Dim clmnNDev09: clmnNDev09 = Array("IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK")
    Dim clmnNDev10: clmnNDev10 = Array("JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK")
    Dim clmnNDev11: clmnNDev11 = Array("KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK")
    
    Call FindTicketsWithMultipleDevelopers(clmnNDev02, "BA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev03, "CA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev04, "DA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev05, "EA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev06, "FA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev07, "GA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev08, "HA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev09, "IA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev10, "JA")
    Call FindTicketsWithMultipleDevelopers(clmnNDev11, "KA")
    
    Dim copyToColumnCopyFromColumn: copyToColumnCopyFromColumn = Array("BA", _
                                                                           "CA", _
                                                                           "DA", _
                                                                           "EA", _
                                                                           "FA", _
                                                                           "GA", _
                                                                           "HA", _
                                                                           "IA", _
                                                                           "JA", _
                                                                           "KA")
                                                              
    For copyFromColumn = 0 To Application.CountA(copyToColumnAAcopyFromColumn) - 1
        Call copyToColumn(copyToColumnCopyFromColumn(copyFromColumn), "DeveloperCounterBackend", 27, "AA")
    Next copyFromColumn
    
    'Remove duplicates from tickets with multiple developers
    ActiveSheet.Range("$AA$1:$AA$1000").RemoveDuplicates Columns:=1, Header:= _
        xlNo
    
    Dim copyToDeveloperCounterFromColumn: copyToDeveloperCounterFromColumn = Array("AA", _
                                                                                   "AM", _
                                                                                   "BM", _
                                                                                   "CM", _
                                                                                   "DM", _
                                                                                   "EM", _
                                                                                   "FM", _
                                                                                   "GM", _
                                                                                   "HM", _
                                                                                   "IM", _
                                                                                   "JM", _
                                                                                   "KM")
                                                                                   
    For copyFromColumnToDeveloperCounter = 0 To Application.CountA(copyToDeveloperCounterFromColumn) - 1
        Call copyToColumn(copyToDeveloperCounterFromColumn(copyFromColumnToDeveloperCounter), "DeveloperCounter", 2, "B")
    Next copyFromColumnToDeveloperCounter

    Sheets("DeveloperCounter").Select
    Range("A:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
    Sheets("DeveloperCounter").Select
    
    For i = 0 To 1000
        Dim lastBlankRow As Long
        lastBlankRow = Cells(Rows.Count, 3).End(xlUp).Row
        If Not Range("C" & lastBlankRow + 1).Offset(0, -1) = "" Then
            Range("C2:L2").Copy
            Range("C" & lastBlankRow + 1).PasteSpecial xlPasteFormulas
        End If
    Next i
    
    Range("A2").Select
    Call stopMacroShowMessage

End Sub
