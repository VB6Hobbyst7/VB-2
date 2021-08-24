VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Tmax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' (1) tmaxreadenv()를 사용하여 환경화일로부터 환경변수를 읽어옴
'=============================================================================
Public Function TUX_READENV(as_envfile As String, as_parm As String) As Integer
    Dim li_ret As Integer
                                                                                                                                                              
    On Error GoTo ErrTrap
    

    li_ret = tmaxreadenv(ByVal as_envfile, ByVal as_parm)
    If li_ret = -1 Then
        MsgBox ("read env file error") & "," & gettperrno()
        TUX_READENV = li_ret
        Exit Function
    End If
    
    Exit Function
ErrTrap:
    Debug.Print Err.Description
    Debug.Assert False

    MsgBox Err.Description, vbExclamation, App.EXEName
End Function

'=============================================================================
' 연결을 한다
'=============================================================================
Public Function TP_INIT() As Integer

    Dim Ret     As Integer
    Dim lsndbuf As Long
    Dim tpinfo  As tpstart_t
    Dim lret    As Long
    Dim iRet    As Integer
    Dim retval  As String
    Dim strptr  As Long
                                                                                     
'''    '-- tpalloc
'''    lsndbuf = tpalloc("TPSTART", "", 0)
'''    If lsndbuf = 0 Then
'''        MsgBox "send buffer allocation error" & gettperrno()
'''        Exit Sub
'''    End If

    '-- Tmax Service Start
    Ret = tpstart(ByVal 0&)
    If Ret = -1 Then
        MsgBox ("tp start error") & "," & gettperrno() '& "," & tpstrerror(gettperrno())
    
        TP_INIT = -1
        End
    End If
    
    TP_INIT = 1

    Exit Function
ErrTrap:
    Debug.Print Err.Description
    Debug.Assert False

    MsgBox Err.Description, vbExclamation, App.EXEName
End Function

'''Public Function TP_INIT(usrname As String, cltname As String) As Integer
'''
'''    Dim ret As Integer
'''    Dim lsndbuf As Long
'''    Dim tpinfo As tpstart_t
'''
'''    lsndbuf = tpalloc("TPSTART", "", 0)
'''    If lsndbuf = 0 Then
'''        TP_INIT_ID = -1
'''        Exit Function
'''    End If
'''
'''    tpinfo.cltname = usrname + Chr(0)
'''    tpinfo.dompwd = cltname + Chr(0)
'''    tpinfo.usrname = "" + Chr(0)
'''    tpinfo.usrpwd = "" + Chr(0)
'''
'''    ret = FilltpstartBuf(lsndbuf, tpinfo)
'''    ret = tpstart(ByVal lsndbuf)
'''    If ret = -1 Then
'''        TP_INIT_ID = -2
'''        Exit Function
'''    End If
'''
'''    Exit Function
'''ErrTrap:
'''    Debug.Print Err.Description
'''    Debug.Assert False
'''
'''    MsgBox Err.Description, vbExclamation, App.EXEName
'''End Function
        
'=============================================================================
' 연결을 끊는다.
'=============================================================================
Public Function TP_TERM() As Integer
    Dim Ret As Integer
                                                                                                                                                              
    On Error GoTo ErrTrap
    
    Ret = tpend()
    If Ret = -1 Then
        MsgBox ("tp end error") & "," & gettperrno()
        TP_TERM = Ret
        Exit Function
    End If

    TP_TERM = Ret
    Exit Function
ErrTrap:
    Debug.Print Err.Description
    Debug.Assert False

    MsgBox Err.Description, vbExclamation, App.EXEName
End Function

'=============================================================================
' 메지시를 불러온다.
'=============================================================================
Public Function get_message() As String
    On Error GoTo ErrTrap

    Call TmaxError(get_message)
    Exit Function
ErrTrap:
    Debug.Print Err.Description
    Debug.Assert False

    MsgBox Err.Description, vbExclamation, App.EXEName
End Function

'=============================================================================
' 서버를 호출한다.
'=============================================================================
Public Function TP_CALL(ByVal svcname As String, ByVal sndbuf As String, rcvbuf As String) As String
    Dim lsndbuf  As Long
    Dim lrcvbuf  As Long
    Dim lrbuflen As Long
    Dim lret     As Long
    Dim ErrMsg   As String
    Dim ll_error As Long
    
     On Error GoTo ErrTrap
     
     
    ' 버퍼를 초기화 한다.
    lsndbuf = tpalloc("STRING", "", 1024000)
    lrcvbuf = tpalloc("STRING", "", 1024000)
    If lsndbuf = 0 Or lrcvbuf = 0 Then
        TP_CALL = "-99"
        Exit Function
    End If
    
    ' 버퍼에 스트링을 카피한다.
    If lstrcpy(ByVal lsndbuf&, ByVal sndbuf) = -1 Then
        TP_CALL = "-98"
        Call tpfree(ByVal lsndbuf)
        Call tpfree(ByVal lrcvbuf)
        Exit Function
    End If
    
    ' 서버를 호출한다.
    lret = tpcall(ByVal svcname, ByVal lsndbuf, ByVal 0, lrcvbuf, lrbuflen, ByVal 0)
    If lret = -1 Then
        ll_error = gettperrno()
        If ll_error = 12 And ll_error <> 11 Then
            TP_CALL = CStr(lret)
            MsgBox ("tpcall error" & ":" & Str(ll_error))
            Call tpfree(ByVal lsndbuf)
            Call tpfree(ByVal lrcvbuf)
            Exit Function
        Else
            lret = 0
        End If
    End If
    
    ' 버퍼에서 텍스트를 추출한다.
    If GETSTR(ByVal lrcvbuf, rcvbuf) = -1 Then
        TP_CALL = "-3"
        Call tpfree(ByVal lsndbuf)
        Call tpfree(ByVal lrcvbuf)
        Exit Function
    End If
    
    TP_CALL = rcvbuf
    Call tpfree(ByVal lsndbuf)
    Call tpfree(ByVal lrcvbuf)
    Exit Function
ErrTrap:
    Debug.Print Err.Description
    Debug.Assert False
    Call tpfree(ByVal lsndbuf)
    Call tpfree(ByVal lrcvbuf)
    MsgBox Err.Description, vbExclamation, App.EXEName
End Function

