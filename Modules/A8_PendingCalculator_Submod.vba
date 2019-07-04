Sub pendingCalculatorCopyTodaysDate()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
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
Sub PendingTimeCopyToMainSheet()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

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
    
    ActiveCell.Select
    Selection.Font.Bold = True
    
    ActiveCell.Offset(0, 37).Select
    
    Sheets("PendingCalculator").Select
    Range("G7").Copy
    Sheets("Sheet1").Select
    Range(ActiveCell.Address).PasteSpecial Paste:=xlPasteValues
    
    Sheets("PendingCalculator").Select
    Range("U4").ClearContents
    
    Call ClearPendingCalculator
    
    Sheets("Sheet1").Select
End Sub
Sub pendingCalculatorSortDates()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
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
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Sheets("PendingCalculator").Select
    Range("U4").Value = ticketNumber
    Sheets("NewChecker").Select
End Sub
