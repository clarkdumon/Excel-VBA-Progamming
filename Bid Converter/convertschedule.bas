Attribute VB_Name = "convertschedule"
Option Explicit

Global baseweek(6) As Variant
Global base_shiftbid(6) As Variant
Global count_weekdate As Long
Global shift_time As Double
Global shift_rd_1 As String
Global shift_rd_2 As String
Const time_difference As Date = #12:00:00 PM#

Public Sub convert_schedule()
Dim x As Long: x = 3
    clear_dataset
    create_week


Do Until IsEmpty(shiftbid.Cells(x, 1))
    get_schedulebid (x)
    create_schedulebid
    post_schedulebid (x)
    Erase base_shiftbid
x = x + 1: Loop

End Sub
Private Sub post_schedulebid(x As Long)
'post converted schedule to sheets

    For count_weekdate = 0 To 6
        Dim index As Long
        For index = LBound(base_shiftbid) To UBound(base_shiftbid)
            If loc_weekdate(count_weekdate) = StrConv(Format(base_shiftbid(index), "DDD"), vbProperCase) Then
                shiftbid.Cells(x, count_weekdate + 7) = Format(base_shiftbid(index), "H:MM AM/PM")
            End If
        Next index
    Next count_weekdate

End Sub
Private Sub clear_dataset()
'clear dataset in sheets
    shiftbid.Range("G:Z").ClearContents
End Sub

Private Sub create_week()

    For count_weekdate = 0 To 6
        baseweek(count_weekdate) = get_weekdate + count_weekdate
        shiftbid.Cells(2, count_weekdate + 7) = loc_weekdate(baseweek(count_weekdate))
    Next count_weekdate

End Sub

Private Sub create_schedulebid()

    For count_weekdate = 0 To 6
        If loc_weekdate(baseweek(count_weekdate)) <> shift_rd_1 And loc_weekdate(baseweek(count_weekdate)) <> shift_rd_2 Then
            base_shiftbid(count_weekdate) = ((baseweek(count_weekdate) + shift_time) + time_difference)
        End If
    Next count_weekdate

End Sub

Private Sub get_schedulebid(x As Long)
    'get shift value
        shift_time = shiftbid.Cells(x, 3)
    'get rd1 value
        shift_rd_1 = reval_rd(shiftbid.Cells(x, 4))
    'get rd2 value
        shift_rd_2 = reval_rd(shiftbid.Cells(x, 5))
End Sub




' Functions ****
Private Function get_weekdate() As Date
    
    get_weekdate = FormatDateTime(Now - WorksheetFunction.Weekday(Now(), 3), vbShortDate)

End Function

Private Function loc_weekdate(data1) As String
    
    loc_weekdate = StrConv(Format(baseweek(count_weekdate), "DDD"), vbProperCase)

End Function
'Change RD String into 3 letter Char
Private Function reval_rd(data As String) As String
    
    reval_rd = StrConv(Left(data, 3), vbProperCase)

End Function
