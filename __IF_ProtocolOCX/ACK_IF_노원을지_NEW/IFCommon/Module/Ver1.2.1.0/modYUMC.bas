Attribute VB_Name = "modYUMC"
'
'   For 신촌세브란스병원
'
Option Explicit

Type SERVERINFO
    DSN1    As String
    DSN2    As String
    DSN3    As String
    DBGbn   As String
End Type
Public gSvrInfo     As SERVERINFO

Public ADOCN1   As ADODB.Connection
Public ADOCN2   As ADODB.Connection


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Public gsLoginUserNm    As String

Public Function GetAutoRegFlag() As String
    On Error GoTo ErrRtn
   
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.Use")
    GetAutoRegFlag = Trim(sBuf)
        
ErrRtn:
    If Err <> 0 Then
        MsgBox "GetAutoRegFlag - Err(" & Err.Description & ")", vbExclamation
    End If
End Function
Public Function GetAutoRegWDate() As String
    On Error GoTo ErrRtn
   
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.WDate")
    GetAutoRegWDate = Trim(sBuf)
        
ErrRtn:
    If Err <> 0 Then
        MsgBox "GetAutoRegWDate - Err(" & Err.Description & ")", vbExclamation
    End If
End Function

Public Function GetRegSvrHWnd() As Long
    On Error GoTo ErrRtn
   
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.HWnd")
    If Trim(sBuf) <> "" Then
        GetRegSvrHWnd = CLng(sBuf)
    End If
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg "GetRegSvrHWnd - Err(" & Err.Description & ")"
    End If
End Function

Public Sub GetServerInfo_YUMC()
    On Error GoTo ErrRtn
    
    Dim sBuf$
    Dim i%
    
'SERVER1
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Server.DSN1")
    gSvrInfo.DSN1 = sBuf
        
'SERVER2
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Server.DSN2")
    gSvrInfo.DSN2 = sBuf
        
'SERVER3
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Server.DSN3")
    gSvrInfo.DSN3 = sBuf
    
'DB GBN
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Server.DBGbn")
    gSvrInfo.DBGbn = sBuf
    
ErrRtn:
    If Err <> 0 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub
Public Sub SetAutoRegFlag(ByVal sPara As String)
    On Error GoTo ErrRtn
    
    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.Use", sPara)
    
    If bRet = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!", vbInformation
    End If
        
ErrRtn:
    If Err <> 0 Then
        ViewMsg "SetAutoRegFlag - Err(" & Err.Description & ")"
    End If
End Sub
Public Sub SetAutoRegWDate(ByVal sPara As String)
    On Error GoTo ErrRtn
    
    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.WDate", sPara)
    
    If bRet = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!", vbInformation
    End If
        
ErrRtn:
    If Err <> 0 Then
        ViewMsg "SetAutoRegFlag - Err(" & Err.Description & ")"
    End If
End Sub

Public Sub SetRegSvrHWnd(ByVal lHWnd As Long)
    On Error GoTo ErrRtn
    
    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.HWnd", Trim(lHWnd))
    
    If bRet <> True Then
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!", vbInformation
    End If
        
ErrRtn:
    If Err <> 0 Then
        MsgBox "SetRegSvrHWnd - Err(" & Err.Description & ")", vbExclamation
    End If
End Sub


