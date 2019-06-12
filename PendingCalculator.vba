Sub deleteUnneededStatues()
    Call startMacroShowMessage(3)
    
    Sheets("PendingCalculator").Select
    Range("A22").Select
    ActiveSheet.Paste
    
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
    ElseIf rgFound Is Nothing Then
        MsgBox "There aren't statuses on Pending!"
    Else
        Dim lastRowWithValue As Long
        lastRowWithValue = Cells(Rows.Count, 1).End(xlUp).Row
    
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
            lastRowWithValueAfter = Cells(Rows.Count, 1).End(xlUp).Row
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
    
    Call stopMacroShowMessage
    
End Sub
Sub pendingCalculatorClear()
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
Sub pendingCalculatorCopyTodaysDate()
    
    Range("C4").Value = Format(Now(), "DD.MM.YYYY HH:MM:SS")
    Range("C4").Replace "-", "/"
    
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown
    Range("B4:C4").Copy
    Range("A22").PasteSpecial Paste:=xlPasteValues
    Range("A22:E500").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Size = 8
        .Bold = False
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

    Call pendingCalculatorSortDates
End Sub

Sub copyAndPasteResolvedDate()
    Dim FindString As String
    Dim rng As Range
    FindString = Range("U4").Value
    
    If Trim(FindString) <> "" Then
        With Sheets("Sheet1").Range("C:C")
            Set rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
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
    
    Call pendingCalculatorClear
    
    Sheets("Sheet1").Select
    
End Sub

Sub copyPendingTime()
    Range("G4").Copy
    Range("G4").Select
End Sub

Sub roundPending()
    Range("G7").Value = Application.WorksheetFunction.RoundDown(Range("G4").Value / 10, 0)
    Range("G7").Copy
End Sub
Sub pendingCalculatorSortDates()
    
    ActiveSheet.Range("$A$21:$E$500").AutoFilter Field:=1, Criteria1:= _
        "Status has been changed to Pending"
        
    Range("B22:B300").Copy
    Range("I10").PasteSpecial Paste:=xlPasteValues
    
    ActiveSheet.Range("$A$21:$E$300").AutoFilter Field:=1, Criteria1:=Array( _
        "Status has been changed to Assigned", _
        "Status has been changed to In Progress", _
        "Status has been changed to Resolved" _
        ), Operator:=xlFilterValues
        
    Range("B22:B100").Copy
    Range("J10").PasteSpecial Paste:=xlPasteValues
    Range("F10:G20").Copy
    Range("L10").PasteSpecial Paste:=xlPasteValues

    ActiveWorkbook.Worksheets("PendingCalculator").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PendingCalculator").Sort.SortFields.Add Key:=Range("L10"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PendingCalculator").Sort
        .SetRange Range("L10:M19")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Copy
    
    Range("B10").PasteSpecial Paste:=xlPasteValues
    Range("I10:J200").ClearContents
    Range("L10:M19").ClearContents
    
    ActiveSheet.Range("$A$21:$E$200").AutoFilter Field:=1, Criteria1:=Array( _
        "Status has been changed to Assigned", _
        "Status has been changed to In Progress", _
        "Status has been changed to Pending", _
        "Status has been changed to Resolved" _
        ), Operator:=xlFilterValues
    
    Range("N4").Copy
    Range("G4").PasteSpecial Paste:=xlPasteValues

    Range("A22:E500").Select
    With Selection
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.Size = 8
        .Font.Bold = False
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Range("A22:E500").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Status has been changed to Pending"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("G4").Select
End Sub
Sub copyTicketNumber(ticketNumber As String)
    Sheets("PendingCalculator").Select
    Range("U4").Value = ticketNumber
    Sheets("NewChecker").Select
End Sub
