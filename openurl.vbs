Sub OpenURL()

Call crdl_Masterfile

Dim obj As Object

Set obj = CreateObject("WScript.Shell")

obj.Run "chrome.exe --new-tab https://docs.google.com/spreadsheets/d/1pAh3oiqV_VJVUsE1hBVwY4i2zIlTtSc4Tv1v3vdmiNo/export?format=csv&gid=168956116"
Application.Wait (Now + TimeValue("00:00:10"))

obj.SendKeys "^w"


Call crdl_Masterfile

MsgBox "Complete"
End Sub




Private Sub crdl_Masterfile()

On Error Resume Next
Kill "C:\Users\" & Environ("USERNAME") & "\Downloads\TDCX Masterfile 2022 - Active.csv"

On Error GoTo 0

End Sub

