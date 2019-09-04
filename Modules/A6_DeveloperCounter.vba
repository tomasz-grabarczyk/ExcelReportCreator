Sub RunCounter()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Call StartMacroShowMessage(35)
    
    Sheets("DeveloperCounter").Visible = True
    Sheets("DeveloperCounterBackend").Visible = True
    
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
        lastBlankRow = Cells(Rows.count, 3).End(xlUp).Row
        If Not Range("C" & lastBlankRow + 1).Offset(0, -1) = "" Then
            Range("C2:L2").Copy
            Range("C" & lastBlankRow + 1).PasteSpecial xlPasteFormulas
        End If
    Next i
    
    Range("A2").Select
    Sheets("DeveloperCounterBackend").Visible = False
    
    Sheets("MultipleDevelopers").Visible = True
    
    Call StopMacroShowMessage

End Sub
Sub HideCounter()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Sheets("DeveloperCounter").Visible = False
    Sheets("MultipleDevelopers").Visible = False
End Sub
