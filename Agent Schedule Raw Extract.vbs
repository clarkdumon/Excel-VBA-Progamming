Option Explicit

Private Sub EVENT_OFF()
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.EnableAnimations = False
Application.ScreenUpdating = False
End Sub
Private Sub EVENT_ON()
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.EnableAnimations = True
Application.ScreenUpdating = True
End Sub

Function FUNCFormatter()

RAW.Columns("AB").NumberFormat = "General"
RAW.Columns("AC").NumberFormat = "m/d/yyyy"
RAW.Columns("AD").NumberFormat = "hh:mm AM/PM"
RAW.Columns("AE").NumberFormat = "hh:mm AM/PM"
RAW.Columns("AF").NumberFormat = "m/d/yyyy"
RAW.Columns("AG").NumberFormat = "hh:mm AM/PM"
RAW.Columns("AH").NumberFormat = "hh:mm AM/PM"
RAW.Columns("AI").NumberFormat = "hh:mm AM/PM"


End Function

Sub GetData()


CLEAR_ScheduleAuditFile "A2", "Z" & getLastRow(ThisWorkbook.Name, "RAW", "B")
CLEAR_ScheduleAuditFile "AA2", "AI" & getLastRow(ThisWorkbook.Name, "RAW", "AA")
FUNCFormatter


Dim FilePath As String
Dim FileName As String
FilePath = OpenAgentScheduleFilePath
FileName = getFileName(FilePath)

    If FilePath = "False" Then
        MsgBox "No File Selected"
        Exit Sub
    End If
EVENT_OFF

SUBGetScheduleFile FilePath, FileName
SUBGenerateDataFile

EVENT_ON
MsgBox "Data has been added to RAW, Proceed in adding Schedules"
End Sub

Private Sub SUBGetScheduleFile(FilePath, FileName)

On Error Resume Next
Workbooks(FileName).Activate
Workbooks.Open FileName:=FilePath, ReadOnly:=True
On Error GoTo 0

Dim SchedFile As Worksheet: Set SchedFile = Workbooks(FileName).Sheets(1)
Dim AuditFile As Worksheet: Set AuditFile = Workbooks(ThisWorkbook.Name).Sheets("RAW")

'Copy Data from Schedule Audit File
SchedFile.Range("A1", Cells(getLastRow(FileName, 1, "B"), 14)).Copy

'Paste Values to Schedule Audit File
AuditFile.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'Close Agent Schedule File
Workbooks(FileName).Activate
Application.CutCopyMode = False
Workbooks(FileName).Close savechanges:=False

End Sub


Private Sub SUBGenerateDataFile()
CLEAR_ScheduleAuditFile "AA2", "AI" & getLastRow(ThisWorkbook.Name, "RAW", "AB")

    Dim lastrow As Long: lastrow = getLastRow(ThisWorkbook.Name, "RAW", "B")
    Dim i As Long: i = 4
    
    
    Do While i <> lastrow
        RAW.Cells(i, "AB") = getAirID(RAW.Cells(i, "B"), RAW.Cells(i - 1, "AB"))
        RAW.Cells(i, "AC") = FUNCGetDate(RAW.Cells(i, "C"), RAW.Cells(i - 1, "AC"))
        RAW.Cells(i, "AD") = FUNCGetTime(RAW.Cells(i, "D"))
        RAW.Cells(i, "AE") = FUNCGetTime(RAW.Cells(i, "E"))
        RAW.Cells(i, "AF") = FUNCGetSchedAct(RAW.Cells(i, "G"))
        RAW.Cells(i, "AG") = FUNCGetTime(RAW.Cells(i, "H"))
        RAW.Cells(i, "AH") = FUNCGetTime(RAW.Cells(i, "K"))
        RAW.Cells(i, "AI") = IIf(Not IsEmpty(RAW.Cells(i, "AD")), FUNCGetRealSched(RAW.Cells(i, "AF"), RAW.Cells(i, "AD")), "")
        RAW.Cells(i, "AA") = WorksheetFunction.Concat(RAW.Cells(i, "AB"), RAW.Cells(i, "AC"), IsEmpty(RAW.Cells(i, "AD")))
        RAW.Cells(i, "Z") = RAW.Cells(i, "AB") & RAW.Cells(i, "AC") & RAW.Cells(i, "AF")
    i = i + 1
    Loop

End Sub
    
Function getAirID(cell, prev)

cellcheck = Len(WorksheetFunction.Substitute(cell, "Agent: ", ""))
On Error Resume Next


        If cellcheck = Len(cell) Or IsEmpty(cell) Then
            getAirID = prev
        Else
            Dim x As Variant: x = WorksheetFunction.Substitute(cell, "Agent: ", "")
            Dim y As Variant: y = WorksheetFunction.Find(" ", x, 1)
            getAirID = CLng(Left(x, y))
        End If
        

End Function


'Get Last Row of the
Function getLastRow(FileName, SheetName, ColName As String) As Long
Dim SchedFile As Worksheet: Set SchedFile = Workbooks(FileName).Sheets(SheetName)

getLastRow = SchedFile.Cells(SchedFile.Rows.Count, ColName).End(xlUp).row

End Function


Function CLEAR_ScheduleAuditFile(StartRange As String, EndRange As String)

Dim SAuditFile As Worksheet: Set SAuditFile = ThisWorkbook.Sheets("RAW")


SAuditFile.Range(StartRange, EndRange).ClearContents

End Function

Function getFileName(OpenAgentScheduleFilePath) As String

getFileName = Mid(OpenAgentScheduleFilePath, InStrRev(OpenAgentScheduleFilePath, "\") + 1)

End Function
Function OpenAgentScheduleFilePath() As String

    Dim FilePath As Variant
    

    FilePath = Application.GetOpenFilename("All Files (*.*), *.*")
    

    
    OpenAgentScheduleFilePath = FilePath
    
End Function

Function FUNCGetDate(DATAdate, upperCell)

If IsDate(DATAdate) Then
    FUNCGetDate = DateValue(DATAdate)
    Exit Function
End If
    FUNCGetDate = upperCell
End Function


Function FUNCGetTime(timedata)

        If timedata = "Off" Or timedata = "OFF" Or timedata = "off" Then
            FUNCGetTime = "OFF"
        Else

            On Error GoTo NOTTIME
            FUNCGetTime = TimeValue(timedata)
        End If
Exit Function
NOTTIME:

FUNCGetTime = ""
End Function

Function FUNCGetSchedAct(schedAct)

If IsEmpty(schedAct) Then
Else
FUNCGetSchedAct = schedAct
End If

End Function
Function FUNCGetRealSched(schedAct, StartTime)

Select Case schedAct
Case "FMLA", "PTO"
    FUNCGetRealSched = schedAct
    Exit Function
Case Else
    FUNCGetRealSched = StartTime
    Exit Function
End Select

End Function




