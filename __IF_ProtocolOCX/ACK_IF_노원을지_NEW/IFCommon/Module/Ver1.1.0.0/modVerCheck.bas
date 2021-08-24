Attribute VB_Name = "modVerCheck"
Option Explicit

Public Function fCurVerObject(ByVal sGbn As String, ByVal sMachCd As String) As String

    Dim sBuf$

    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachCd, "v" & sGbn)
            
    If sBuf = "" Then
        MsgBox "현재 Version을 구하려는 개체의 구분자가 틀립니다!!"
    Else
        fCurVerObject = sBuf
    End If
End Function
