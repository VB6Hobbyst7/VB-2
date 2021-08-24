Attribute VB_Name = "modHUniCode"
Option Explicit

Public Function LenH(ByVal anystr As String) As Integer
    LenH = LenB(StrConv(anystr, vbFromUnicode))
End Function

Public Function LeftH(ByVal anystr As String, ByVal nPos As Integer) As String
    LeftH = StrConv(LeftB(StrConv(anystr, vbFromUnicode), nPos), vbUnicode)
End Function

Public Function RightH(ByVal anystr As String, ByVal nPos As Integer) As String
    RightH = StrConv(RightB(StrConv(anystr, vbFromUnicode), nPos), vbUnicode)
End Function

Public Function MidH(ByVal anystr As String, ByVal nStartPos As Integer, nSize As Integer) As String
    MidH = StrConv(MidB(StrConv(anystr, vbFromUnicode), nStartPos, nSize), vbUnicode)
End Function

