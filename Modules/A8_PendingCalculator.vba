Sub RunPendings()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 23.09.2019  **********

    Call StartMacroShowMessage(3)
    
    Sheets("PendingCalculator").Select
    Range("A22").Select
    
    Dim myDataObject As DataObject
    Set myDataObject = New DataObject
    myDataObject.GetFromClipboard

    'Check if there is anything in clipboard to be pasted
    If Range("U10") = "Automatic" Then
        If myDataObject.GetFormat(1) = True Then
            ActiveSheet.Paste
        Else
            'If there is nothing is clipboard, show message and exit macro
            MsgBox "There is nothing to be pasted!"
            Call StopMacroShowMessage
            Exit Sub
        End If
    End If
    
    ActiveSheet.Range("$A$21:$E$500").AutoFilter Field:=1, Criteria1:="<>Status has been changed to*"
    ActiveSheet.Range("$A$22:$E$500").ClearContents
    Rows("31:31").AutoFilter
    
    Range("A22:A500").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Dim cell As Range
    
    Dim i As Integer
    Dim cellRangeValue As String
    
    Set rgFound = Range("A22:A500").Find("Status has been changed to Pending")
    Set rgFoundAnything = Range("A22:A500").Find("Status has been changed to*")
    
    If rgFoundAnything Is Nothing Then
        MsgBox "There is nothing to work with!"
        Call ClearPendingCalculator
    ElseIf rgFound Is Nothing Then
        MsgBox "There aren't statuses on Pending!"
        Call ClearPendingCalculator
    Else
        Dim lastRowWithValue As Long
        lastRowWithValue = Cells(Rows.count, 1).End(xlUp).Row
    
        For i = lastRowWithValue To 21 Step -1
            cellRangeValue = Range("A" & i).Address
        
            Range(cellRangeValue).Select
            If (ActiveCell.Value = "Status has been changed to Closed" Or ActiveCell.Value = "Status has been changed to Resolved" Or ActiveCell.Value = "Status has been changed to In Progress" Or ActiveCell.Value = "Status has been changed to Assigned" Or ActiveCell.Value = "") Then
                Range(ActiveCell.Address).EntireRow.Delete
            Else
                Exit For
            End If
        Next i
        
        If Range("A22") = "Status has been changed to Pending" Then
            Call pendingCalculatorCopyTodaysDate
            
            Dim lastRowWithValueAfter As Long
            lastRowWithValueAfter = Cells(Rows.count, 1).End(xlUp).Row
            Range("A" & lastRowWithValueAfter).Select
        End If
            
        
        For checkIfPendingsLeft = 0 To 20
            If ActiveCell.Offset(-2, 0).Value <> "Status has been changed to Pending" Then
                If ActiveCell.Offset(-2, 0).Value = "Status" Then
                    Exit For
                Else
                    ActiveCell.Offset(-2, 0).EntireRow.Delete
                    ActiveCell.Offset(-1, 0).Select
                End If
            Else
                ActiveCell.Offset(-2, 0).Select
            End If
        Next checkIfPendingsLeft
        
        
        Call pendingCalculatorSortDates
    End If
    
    Sheets("PendingCalculator").Range("U10").Value = "Automatic"
    
    Call StopMacroShowMessage
End Sub
Sub ResolutionTime()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Dim FindString As String
    Dim rng As Range
    FindString = Range("U4").Value
    
    If Trim(FindString) <> "" Then
        With Sheets("Sheet1").Range("C:C")
            Set rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not rng Is Nothing Then
                Application.Goto rng, True
            End If
        End With
    End If
    
    ActiveCell.Offset(0, 12).Select
    
    Sheets("PendingCalculator").Select
    Range("Q11").Copy
    Sheets("Sheet1").Select
    Range(ActiveCell.Address).PasteSpecial Paste:=xlPasteValues
    
    Range(ActiveCell.Address).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    
    Sheets("PendingCalculator").Select
    Range("U4").ClearContents
    
    Call ClearPendingCalculator
    
    Sheets("Sheet1").Select
End Sub
Sub ClearPendingCalculator()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Sheets("PendingCalculator").Select
    
    Range("A22:E22").AutoFilter
    
    Range("A22:E1000").ClearContents
    
    Range("A22:E1000").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("B10:C19").ClearContents
    Range("I:J").ClearContents
    Range("G4").ClearContents
    Range("G7").ClearContents
    Range("A22").Select
End Sub
Sub RoundPendingDown()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Range("G7").Value = Application.WorksheetFunction.RoundDown(Range("G4").Value / 10, 0)
    Call PendingTimeCopyToMainSheet
End Sub
