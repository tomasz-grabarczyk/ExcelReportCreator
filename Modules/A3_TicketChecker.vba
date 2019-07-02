Sub CheckDates()
    Call StartMacroShowMessage(2)
        
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:BG").Select
    Selection.EntireColumn.Hidden = True
    
    Dim checkDatesToCell As Integer: checkDatesToCell = 10000
    
    For colorAssigned = 2 To checkDatesToCell
        If (Range("F" & colorAssigned).Value = "Assigned" Or Range("F" & colorAssigned).Value = "In Progress" Or Range("F" & colorAssigned).Value = "Pending" Or Range("F" & colorAssigned).Value = "Resolved") And Range("K" & colorAssigned).Value = "" Then
            Range("K" & colorAssigned).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & colorAssigned).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorAssigned
    
    For colorInProgress = 2 To checkDatesToCell
        If (Range("F" & colorInProgress).Value = "In Progress" Or Range("F" & colorInProgress).Value = "Pending" Or Range("F" & colorInProgress).Value = "Resolved") And Range("L" & colorInProgress).Value = "" Then
            Range("L" & colorInProgress).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & colorInProgress).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorInProgress
    
    For colorPending = 2 To checkDatesToCell
        If Range("F" & colorPending).Value = "Pending" And Range("M" & colorPending).Value = "" Then
            Range("M" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        ElseIf Range("F" & colorPending).Value = "Resolved" And Range("N" & colorPending).Value = "" Then
            Range("N" & colorPending).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorPending).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorPending
    
    For colorResolved = 2 To checkDatesToCell
        If Range("F" & colorResolved).Value = "Resolved" And Range("O" & colorResolved).Value = "" Then
        Range("O" & colorResolved).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            Range("C" & colorResolved).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorResolved
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub SAPAreaCorrectness()
    Call StartMacroShowMessageString("checking SAP Area correctness ...")
    
    Dim CheckSAPAreaCorrectness As Integer: CheckSAPAreaCorrectness = 10000
    
    For C = 2 To CheckSAPAreaCorrectness
        If Not Range("H" & C).Value = "" And Not Range("F" & C).Value = "" Then
            If Not (Range("H" & C).Value = "BP2" Or _
                    Range("H" & C).Value = "ACE" Or _
                    Range("H" & C).Value = "BP5" Or _
                    Range("H" & C).Value = "HRP" Or _
                    Range("H" & C).Value = "RE-FX" Or _
                    Range("H" & C).Value = "IFRS") Then
                Range("H" & C).Select
                With Selection.Interior
                    .Color = 13260
                    .PatternTintAndShade = 0
                End With
                'Color Incident Numbers
                Range("C" & C).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next C
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub SLACheckLayout()

    Call BackToNormal
    
    Sheets("Sheet1").Select
    
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=6, Criteria1:=Array( _
        "Assigned", _
        "In Progress", _
        "Pending"), Operator:=xlFilterValues
    
    Columns("A:A").EntireColumn.Hidden = True
    Columns("G:G").EntireColumn.Hidden = True
    Columns("I:Y").EntireColumn.Hidden = True
    Columns("AA:AD").EntireColumn.Hidden = True
    Columns("AF:AM").EntireColumn.Hidden = True
    Columns("AO:AV").EntireColumn.Hidden = True
    Columns("AY:BG").EntireColumn.Hidden = True
    
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=5
    ActiveSheet.Range("$A$1:$AP$10000").AutoFilter Field:=38, Criteria1:=">=" & 11
    
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AX1:AX10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    Sheets("Sheet1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
End Sub


