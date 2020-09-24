Attribute VB_Name = "function_module"
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

Function status_th(status As String) As Double
    
    Select Case status
        Case "Break"
            status_th = TimeSerial(0, 30, 59)
            Exit Function
        Case "Lunch"
            status_th = TimeSerial(1, 0, 59)
            Exit Function
        Case "Personal"
            status_th = TimeSerial(0, 10, 59)
            Exit Function
        Case "Ticket-Processing"
            status_th = TimeSerial(0, 30, 59)
            Exit Function
    End Select
End Function

Function exed_th(status, status_d, status_th) As Boolean
exed_th = False
    If status = "Break" Or status = "Lunch" Or status = "Personal" Or status = "Ticket-Processing" Then
    
        If status_d >= status_th Then
            exed_th = True
        End If
    End If
End Function
