Option Explicit
Public Sub getData()
Call readRaw
MsgBox "Done!"
End Sub
Private Sub readRaw()
Dim rRow As Long: rRow = 2 'ASDRAW Row Counter
Dim cName As Integer: cName = 3 ' ASDRAW Name Loc
Dim cTimestamp As Integer: cTimestamp = 6 ' ASDRAW Reason Timestamp Loc
Dim cRcode As Integer: cRcode = 7 ' ASDRAW Reason Code Loc
Dim cDuration As Integer: cDuration = 11 ' ASDRAW Duration Loc

Dim dBreak As String: dBreak = "ReasonCode=Break" ' ASDRAW Reason Code Param
Dim dLunch As String: dLunch = "ReasonCode=Lunch" ' ASDRAW Reason Code Param

Dim pRow As Long: pRow = 2 'Data Row Counter
Dim dLunchParam As Date: dLunchParam = TimeSerial(1, 1, 0)
Dim dBreakParam As Date: dBreakParam = TimeSerial(0, 31, 0)
Call dataHeader

Do Until IsEmpty(asdRaw.Cells(rRow, 1))
'Debug.Print (asdRaw.Cells(rRow, cRcode))


    If (asdRaw.Cells(rRow, cRcode) = dBreak And asdRaw.Cells(rRow, cDuration) > dBreakParam) Or (asdRaw.Cells(rRow, cRcode) = dLunch And asdRaw.Cells(rRow, cDuration) > dLunchParam) Then
        data.Cells(pRow, 1) = asdRaw.Cells(rRow, cName) 'paste Name
        data.Cells(pRow, 2) = asdRaw.Cells(rRow, cTimestamp) 'paste Reason TimeStamp
        data.Cells(pRow, 3) = asdRaw.Cells(rRow, cRcode) 'paste Reason Code
        data.Cells(pRow, 4) = asdRaw.Cells(rRow, cDuration) 'paste Reason Duration
        pRow = pRow + 1
    End If

rRow = rRow + 1
Loop

Call clearRaw
End Sub

Private Sub dataHeader()
data.Cells(1, 1).CurrentRegion.ClearContents

        data.Cells(1, 1) = "Name"
        data.Cells(1, 2) = "Reason Timestamp"
        data.Cells(1, 3) = "Reason Code"
        data.Cells(1, 4) = "Reason Duration"

End Sub

Private Sub clearRaw()

asdRaw.Cells(1, 1).CurrentRegion.EntireRow.Delete

End Sub


