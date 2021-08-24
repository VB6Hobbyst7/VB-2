Attribute VB_Name = "modUTF8"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetACP Lib "Kernel32" () As Long
Private Const CP_ACP = 0
Private Const CP_UTF8 = 65001

Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String

    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    
    If cpg = -1 Then cpg = GetACP()
    
    pwz = StrPtr(st)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
    
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer))
    AToW = Left$(stBuffer, cwch - 1)
    
End Function

Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String

    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    Dim lpUsedDefaultChar As Long
    
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    
    stBuffer = String$(cwch + 1, vbNullChar)
    
    pwzBuffer = StrPtr(stBuffer)
    
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    
    WToA = Left$(stBuffer, cwch - 1)
    
End Function

Public Function EncodeUTF8(ByVal cnvUni As String)

    If cnvUni = vbNullString Then Exit Function
    
    EncodeUTF8 = WToA(cnvUni, CP_UTF8, 0)
    
End Function

Public Function DecodeUTF8(ByVal cnvUni As String)

    If cnvUni = vbNullString Then Exit Function
    
    cnvUni = WToA(cnvUni, CP_ACP)
    DecodeUTF8 = AToW(cnvUni, CP_UTF8)
    
End Function

Public Function EncodeUTF8_Byte(W As String, UTF8() As Byte) As Long
    On Error GoTo ErrEncode
    
    Dim Bytes&, pwzBuffer&
    Dim stBuffer$
    
    If LenB(W) = 0 Then Exit Function
    ReDim UTF8(LenB(W))     ' + Len(W))
    
    Bytes = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(W), Len(W), UTF8(0), UBound(UTF8), 0, ByVal 0&)
    
    stBuffer = String$(Bytes + 1, vbNullChar)
    
    pwzBuffer = StrPtr(stBuffer)
    
    Bytes = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(W), -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    
    If Bytes = 0 Then
        ReDim Preserve UTF8(0)
    Else
        ReDim Preserve UTF8(Bytes - 1)
    End If
    
    EncodeUTF8_Byte = Bytes

ErrEncode:
    If Err <> 0 Then
    End If
End Function

