Sub SaveToMotherFile()
    
    If Environ$("Username") = "A702387" Then
        Call SaveToMotherFileBackend( _
                                    "https://atos365-my.sharepoint.com/personal/tomasz_grabarczyk_atos_net/Documents/AMS_ARDAGH/AMS_ARDAGH.xlsm", _
                                    "C:\Users\A702387\OneDrive - Atos\AMS_ARDAGH")
    ElseIf Environ$("Username") = "A700473" Then
        Call SaveToMotherFileBackend( _
                                    "https://atos365-my.sharepoint.com/personal/tomasz_grabarczyk_atos_net/Documents/AMS_ARDAGH/AMS_ARDAGH.xlsm", _
                                    "C:\Users\A700473\Atos\Grabarczyk, Tomasz - AMS_ARDAGH (3)")
    
    End If
End Sub
Sub DefaultLayout()
    Sheets("Sheet1").Select
    ActiveWindow.Zoom = 85
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Sheets("PendingCalculator").Select
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Sheets("Sheet1").Select
    Range("A1").Select
End Sub
