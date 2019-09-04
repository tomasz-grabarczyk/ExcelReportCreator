Sub SaveToMotherFileBackend(motherFileName As String, directory As String)
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********

    ActiveWorkbook.Save

    ChDir directory
    ActiveWorkbook.SaveAs fileName:=motherFileName
    
    Sheets("Sheet1").Select

    
    Range("A2:AZ10000").Copy
    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveWorkbook.Save
    
    Application.Quit
End Sub
