Attribute VB_Name = "DMJ0101"
Option Explicit
Function sNow_Date(sopt As String) As String

    Dim sNow   As String
    
    Select Case sopt
    Case "D"
        sNow = Format(Now, "YYYY-MM-DD")
    Case "T"
        sNow = Format(Now, "hh-mm-ss")
    Case "A"
        sNow = Format(Now, "YYYY-MM-DD-hh-mm-ss")
    Case "DS"
        sNow = Format(Now, "YYYYMMDD")
    Case "TS"
        sNow = Format(Now, "hhmmss")
    Case "AS"
        sNow = Format(Now, "YYYYMMDDhhmmss")
    End Select
    
    sNow_Date = sNow
    
End Function
