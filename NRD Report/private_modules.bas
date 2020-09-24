Attribute VB_Name = "private_modules"
Option Explicit
Public Sub generate_NRD_report()
'turn off event handlers
Application.Run "event_handlers.get_time"
Application.Run "event_handlers.turn_off_events"

'turn data to correct values
Call makevalues
Call generate_raw
Call report_over


'return event handlers back on
Application.Run "event_handlers.turn_on_events"
Application.Run "event_handlers.get_time"
End Sub

Private Sub makevalues()
Dim x As Long: x = 2

Do Until IsEmpty(raw.Cells(x, 1))

raw.Cells(1, 1) = "name"
raw.Cells(1, 2) = "time"
raw.Cells(1, 3) = "reason"
raw.Cells(1, 4) = "reason_duration"
raw.Cells(1, 5) = "date"



    
    raw.Cells(x, 2) = raw.Cells(x, 2).Value
    raw.Cells(x, 4) = raw.Cells(x, 4).Value
    raw.Cells(x, 5) = FormatDateTime(raw.Cells(x, 2), vbShortDate)

x = x + 1: Loop
overbreak.Range("A:Z").ClearContents
overbreak.Cells(1, 1) = "name"
overbreak.Cells(1, 2) = "status"
overbreak.Cells(1, 3) = "date"
overbreak.Cells(1, 4) = "duration"
overbreak.Range("C:C").NumberFormat = "MM/DD/YY"
overbreak.Range("D:D").NumberFormat = "H:MM:SS"
overLunch.Range("A:Z").ClearContents
overLunch.Cells(1, 1) = "name"
overLunch.Cells(1, 2) = "status"
overLunch.Cells(1, 3) = "date"
overLunch.Cells(1, 4) = "duration"
overLunch.Range("C:C").NumberFormat = "MM/DD/YY"
overLunch.Range("D:D").NumberFormat = "H:MM:SS"
overPersonal.Range("A:Z").ClearContents
overPersonal.Cells(1, 1) = "name"
overPersonal.Cells(1, 2) = "status"
overPersonal.Cells(1, 3) = "date"
overPersonal.Cells(1, 4) = "duration"
overPersonal.Range("C:C").NumberFormat = "MM/DD/YY"
overPersonal.Range("D:D").NumberFormat = "H:MM:SS"
overTP.Range("A:Z").ClearContents
overTP.Cells(1, 1) = "name"
overTP.Cells(1, 2) = "status"
overTP.Cells(1, 3) = "date"
overTP.Cells(1, 4) = "duration"
overTP.Range("C:C").NumberFormat = "MM/DD/YY"
overTP.Range("D:D").NumberFormat = "H:MM:SS"



End Sub
Private Sub sort_rawData()

With raw.Sort
    .SortFields.Clear
    .SortFields.Add Key:=raw.Range("A1"), Order:=xlAscending
    .SortFields.Add Key:=Range("E1"), Order:=xlAscending
    .SetRange raw.Range("A:E")
    .Header = xlYes
    .Apply
End With

End Sub

Private Sub generate_raw()

Dim xraw As Long: xraw = 2
Dim graw As Long: graw = 2
Dim gstring As String

Application.Run "event_handlers.clean_genRaw"
genRaw.Range("J:J").NumberFormat = "MM/DD/YY"
Do Until IsEmpty(raw.Cells(xraw, 1))

    gstring = raw.Cells(xraw, 1) & raw.Cells(xraw, 3) & Format(raw.Cells(xraw, 5), "0")
    If find_gstring(gstring) = False Then
        genRaw.Cells(graw, 7) = gstring
        genRaw.Cells(graw, 8) = raw.Cells(xraw, 1)
        genRaw.Cells(graw, 9) = raw.Cells(xraw, 3)
        genRaw.Cells(graw, 10) = raw.Cells(xraw, 5)
        genRaw.Cells(graw, 11) = WorksheetFunction.SumIfs(raw.Range("D:D"), raw.Range("A:A"), genRaw.Cells(graw, 8), raw.Range("C:C"), genRaw.Cells(graw, 9), raw.Range("E:E"), genRaw.Cells(graw, 10))
        graw = graw + 1
    End If
xraw = xraw + 1: Loop

End Sub

Private Sub report_over()

Dim oBreak As Long: oBreak = 2
Dim oLunch As Long: oLunch = 2
Dim oPersonal As Long: oPersonal = 2
Dim otp As Long: otp = 2
Dim graw As Long: graw = 2


Do Until IsEmpty(genRaw.Cells(graw, 7))
Dim status As String: status = genRaw.Cells(graw, 9)
Dim status_d As Double: status_d = genRaw.Cells(graw, 11)
    If exed_th(status, status_d, status_th(status)) = True Then
        Select Case status
            Case "Break"
                overbreak.Cells(oBreak, 1) = genRaw.Cells(graw, 8)
                overbreak.Cells(oBreak, 2) = genRaw.Cells(graw, 9)
                overbreak.Cells(oBreak, 3) = genRaw.Cells(graw, 10)
                overbreak.Cells(oBreak, 4) = genRaw.Cells(graw, 11)
                oBreak = oBreak + 1
            Case "Lunch"
                overLunch.Cells(oLunch, 1) = genRaw.Cells(graw, 8)
                overLunch.Cells(oLunch, 2) = genRaw.Cells(graw, 9)
                overLunch.Cells(oLunch, 3) = genRaw.Cells(graw, 10)
                overLunch.Cells(oLunch, 4) = genRaw.Cells(graw, 11)
                oLunch = oLunch + 1
            Case "Personal"
                overPersonal.Cells(oPersonal, 1) = genRaw.Cells(graw, 8)
                overPersonal.Cells(oPersonal, 2) = genRaw.Cells(graw, 9)
                overPersonal.Cells(oPersonal, 3) = genRaw.Cells(graw, 10)
                overPersonal.Cells(oPersonal, 4) = genRaw.Cells(graw, 11)
                oPersonal = oPersonal + 1
            Case "Ticket-Processing"
                overTP.Cells(otp, 1) = genRaw.Cells(graw, 8)
                overTP.Cells(otp, 2) = genRaw.Cells(graw, 9)
                overTP.Cells(otp, 3) = genRaw.Cells(graw, 10)
                overTP.Cells(otp, 4) = genRaw.Cells(graw, 11)
                otp = otp + 1
        End Select
    End If
graw = graw + 1: Loop


End Sub


