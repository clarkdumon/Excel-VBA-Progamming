Attribute VB_Name = "event_handlers"
Private Sub turn_off_events()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Cursor = xlWait
    
End Sub
Private Sub turn_on_events()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub



Private Sub clean_genRaw()

    genRaw.Range("A:ZZ").ClearContents
    genRaw.Cells(1, 7) = "concat"
    genRaw.Cells(1, 8) = "name"
    genRaw.Cells(1, 9) = "status"
    genRaw.Cells(1, 10) = "date"
    genRaw.Cells(1, 11) = "total_duration"
    

End Sub



Private Sub get_time()
Debug.Print Now
End Sub
