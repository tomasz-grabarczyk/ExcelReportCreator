Sub CheckChaRMStatuses()
    Dim lRow As Long: lRow = Cells(Rows.Count, 3).End(xlUp).Row
    
    Call StartMacroShowMessage(6)
    
    Call HideColumnsForChaRM
    
    Sheets("Sheet1").Select
        Range("BA2:BB" & lRow).ClearContents
    
    For rfc = 2 To lRow
        Range("AY" & rfc).Select
        Call CompareStringsRfC
    Next rfc
    
    For cd = 2 To lRow
        Range("AZ" & cd).Select
        Call CompareStringsCD
    Next cd

    ActiveWindow.ScrollRow = 1
    
    Call StopMacroShowMessage
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
    Dim ticketNumbersToColor As Long: ticketNumbersToColor = Cells(Rows.Count, 3).End(xlUp).Row
    
    For colorIterator = 2 To ticketNumbersToColor
        If Not Range("BA" & colorIterator).Value = "" Or Not Range("BB" & colorIterator).Value = "" Then
            Range("C" & colorIterator).Select
            With Selection.Interior
                .Color = 16751001
                .PatternTintAndShade = 0
            End With
        End If
    Next colorIterator
    
    Call FilterIncidentNumbersByColor
    
    'Exclude below from filtering
    'ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=57, Criteria1:="<>Status in ChaRM cannot be changed due to upgrade (freeze)."
End Sub
Sub LoadChaRMInformation()
    Call StartMacroShowMessage(6)
    
    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    Call LoadChaRMDataFromFilesBackend("rfc.csv", "A1", "ChaRM RfC", "Z")
    Call LoadChaRMDataFromFilesBackend("cd.csv", "A1", "ChaRM CD", "V")
    
    Call CopyDataFromRfCAndCDToChaRMSheet
    Call RemoveMultipleOccurencesOfTickets
    Call ScrollMaxUpAndLeft
    
    Call StopMacroShowMessage
End Sub
