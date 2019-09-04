Sub LoadChaRMInformation()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Call StartMacroShowMessage(6)
    
    ChDir "C:\Users\" & Environ$("Username") & "\Downloads"
    Call LoadChaRMDataFromFilesBackend("rfc.csv", "A1", "RfC", "Z")
    Call LoadChaRMDataFromFilesBackend("cd.csv", "A1", "CD", "V")
    
    Call CopyDataFromRfCAndCDToChaRMSheet
    Call RemoveMultipleOccurencesOfTickets
    Call ScrollMaxUpAndLeft
    
    Call StopMacroShowMessage
End Sub
Sub CheckChaRMStatuses()
    '********** Author: Tomasz Grabarczyk **********
    '**********  Last update: 03.07.2019  **********
    
    Dim lRow As Long: lRow = Cells(Rows.count, 3).End(xlUp).Row
    
    Call StartMacroShowMessage(6)
      
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

    Call HideColumnsForChaRM
    
    ActiveWindow.ScrollRow = 1
    
    Call StopMacroShowMessage
End Sub
