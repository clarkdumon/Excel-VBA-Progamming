Attribute VB_Name = "private_modules"
Option Explicit
Public Sub generate_NRD_report()
'turn off event handlers
Application.Run "event_handlers.turn_off_events"

'turn data to correct values
makevalues




'return event handlers back on
Application.Run "event_handlers.turn_on_events"
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
Application.Run "event_handlers.get_time"
genRaw.Range("J:J").NumberFormat = "MM/DD/YY"
Do Until IsEmpty(raw.Cells(xraw, 1))

    gstring = raw.Cells(xraw, 1) & raw.Cells(xraw, 3) & Format(raw.Cells(xraw, 5), "0")
    If find_gstring(gstring) = False Then
        genRaw.Cells(graw, 7) = gstring
        genRaw.Cells(graw, 8) = raw.Cells(xraw, 1)
        genRaw.Cells(graw, 9) = raw.Cells(xraw, 3)
        genRaw.Cells(graw, 10) = raw.Cells(xraw, 5)
        graw = graw + 1
    End If
xraw = xraw + 1: Loop


Application.Run "event_handlers.get_time"

End Sub


Function find_gstring(findvalue As String) As Boolean
Dim rng As Range
On Error GoTo notFound
Set rng = genRaw.Range("G:G").Find(what:=findvalue)

    If rng.Address = rng.Address Then
        find_gstring = True
        Exit Function
    End If

Exit Function
notFound:
find_gstring = False

End Function
