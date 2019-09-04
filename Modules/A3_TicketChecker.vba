Sub CheckDates()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call BackToNormal
    
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
Sub SAPSystemCorrectness()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call BackToNormal
        
    Call StartMacroShowMessageString("checking SAP Area correctness ...")
    
    Dim CheckSAPSystemCorrectness As Integer: CheckSAPSystemCorrectness = 10000
    
    For C = 2 To CheckSAPSystemCorrectness
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
    
    ActiveSheet.Range("$A$1:$BH$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("C1").Select
    
    Call StopMacroShowMessage
End Sub
Sub DeveloperStatusCheck()
    
    Call BackToNormal
        
    Dim numberOfConsultants As Integer, developers() As String
    
    numberOfConsultants = 0
    
    Sheets("ConsultantList").Visible = True
    Sheets("ConsultantList").Select
    
    Dim lRow As Long: lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Count developers
    For i = 1 To lRow
        If Sheets("ConsultantList").Range("A" & i) = "ABAP" Then
            numberOfConsultants = numberOfConsultants + 1
        End If
    Next i
    
    'Define dynamic array
    ReDim developers(numberOfConsultants)
    
    'Populate dynamic array with developers
    For i = 1 To lRow
        If Sheets("ConsultantList").Range("A" & i) = "ABAP" Then
            developers(i - 1) = Range("B" & i).Value
        End If
    Next i
           
    'Filter by statuses: "Assigned", "In Progress", "Pending", "Resvoled"
    Sheets("ConsultantList").Visible = False
    Sheets("Sheet1").Select
    ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending", "Resvoled"), Operator:=xlFilterValues
                    
    Dim numberOfTickets As Long, totalNumberOfTickets As Long
    totalNumberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))
    
    'Check if developer doesn't have proper SAP Area ("Development Atos GDC" OR "Development")
    For i = 1 To numberOfConsultants
        ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=5, Criteria1:=Array(developers(i)), Operator:=xlFilterValues
        numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))
        For cellAdd = 1 To totalNumberOfTickets
            If Range("E" & cellAdd) = developers(i) And Not (Range("E" & cellAdd).Offset(0, -1).Value = "Development" Or Range("E" & cellAdd).Offset(0, -1).Value = "Development Atos GDC" Or Range("E" & cellAdd).Offset(0, -1).Value = "Transport Management") Then
                Range("C" & cellAdd).Select
                With Selection.Interior
                    .Color = 16751001
                    .PatternTintAndShade = 0
                End With
            End If
        Next cellAdd
    Next i
    
    Call BackToNormal
    
    ActiveSheet.Range("$A$1:$CA$10000").AutoFilter Field:=6, Criteria1:=Array( _
                    "Assigned", "In Progress", "Pending", "Resvoled"), Operator:=xlFilterValues
                    
    'Check if functional consultant has Development SAP Area
    For i = 1 To totalNumberOfTickets
        If (Range("D" & i).Value = "Development Atos GDC" Or Range("D" & i).Value = "Development") And Not ( _
                                                                                                                 Range("E" & i) = developers(1) Or _
                                                                                                                 Range("E" & i) = developers(2) Or _
                                                                                                                 Range("E" & i) = developers(3) Or _
                                                                                                                 Range("E" & i) = developers(4) Or _
                                                                                                                 Range("E" & i) = developers(5) Or _
                                                                                                                 Range("E" & i) = developers(6) _
        ) Then
          Range("C" & i).Select
             With Selection.Interior
                 .Color = 16751001
                 .PatternTintAndShade = 0
             End With
        End If
    Next i
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
End Sub
Sub ClosedDateCheck()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 02.08.2019  **********

    Call BackToNormal
    
    Dim numberOfTickets As Long
    numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))

    For i = 1 To numberOfTickets
        If (Range("F" & i) = "Resolved" Or Range("F" & i) = "Closed" Or Range("F" & i) = "Cancelled") And (Range("N" & i) = "" Or Range("O" & i) = "") Then
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
End Sub
Sub DiscrepanciesCheck()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 02.08.2019  **********

    Call BackToNormal

    Dim numberOfTickets As Long
    numberOfTickets = Application.WorksheetFunction.Subtotal(103, Range("A2:A10000"))

    'Check if "Priority - J" is present in closed tickets
    For i = 1 To numberOfTickets
        If (Range("F" & i) = "Resolved" Or Range("F" & i) = "Closed" Or Range("F" & i) = "Cancelled") And Range("J" & i) = "" Then
            Range("J" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    'Check if "SLA Resolution Time (days) - AC" is present in closed tickets
    For i = 1 To numberOfTickets
        If (Range("F" & i) = "Resolved" Or Range("F" & i) = "Closed" Or Range("F" & i) = "Cancelled") And Range("AC" & i) = "" Then
            Range("AC" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    'Check if ticket in "ARD SAP AMS" has empty "SAP Area - D"
    For i = 1 To numberOfTickets
        If Range("A" & i) = "ARD SAP AMS" And (Range("D" & i) = "" Or Range("E" & i) = "N/A") Then
            Range("D" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    'Check if ticket has In Progress Start Date and is not in Development
    For i = 1 To numberOfTickets
        If Not Range("L" & i).Value = "" And Range("F" & i) = "Assigned" And Not (Range("D" & i).Value = "Development" Or Range("D" & i).Value = "Development Atos GDC") Then
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    'Check if ticket on Pending status has empty "Status Reason - G"
    For i = 1 To numberOfTickets
        If Range("F" & i) = "Pending" And (Range("G" & i) = "") Then
            Range("G" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    'Check if ticket on Pending status has empty "Reason of Pending Status - AI"
    For i = 1 To numberOfTickets
        If Range("F" & i) = "Pending" And (Range("AI" & i) = "") Then
            Range("AI" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal
    
    'Check if ticket should have Ticket Type changed
    For i = 1 To numberOfTickets
        If Range("F" & i) <> "Closed" And Range("D" & i) = "Monitoring" And (Range("B" & i) = "User Service Restoration") Then
            Range("B" & i).Select
            With Selection.Interior
                .Color = 13260
                .PatternTintAndShade = 0
            End With
            'Color Incident Numbers
            Range("C" & i).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Call BackToNormal


    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
    
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
End Sub
