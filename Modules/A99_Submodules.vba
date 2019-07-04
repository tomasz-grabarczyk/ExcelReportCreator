Sub ScrollMaxUpAndLeft()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Sheets("Sheet1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
End Sub
Sub FindLastRowWithData(columnNumber As Variant)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    Dim lRow As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, columnNumber).End(xlUp).Row
    
    MsgBox "Last Row: " & lRow
End Sub
Sub FilterIncidentNumbersByColor()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ActiveSheet.Range("$A$1:$BG$10000").AutoFilter Field:=3, Criteria1:=RGB(153, 153, 255), Operator:=xlFilterCellColor
End Sub
