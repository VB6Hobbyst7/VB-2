Attribute VB_Name = "Tuxfuct"
Option Explicit

    Global strbuf As tuxbuf
Function GETSTR(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As String * 1024
    Dim Ret As Integer
    Dim err_ret As Integer
    
    wdata = String$(1024, Chr$(32))
    Ret = Fget32(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, 0)
    
    GETSTR = Ret
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fget32", 3)
        Exit Function
    End If
    
    Text = wdata
    'Text = Left(Trim$(wdata), Len(Trim$(wdata)) - 1)
    
End Function

Function GETSTR16(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As String * 1024
    Dim Ret As Integer
    Dim err_ret As Integer
    
    wdata = String$(1024, Chr$(32))
    Ret = Fget(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, 0)
    
    GETSTR16 = Ret
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fget", 3)
        Exit Function
    End If
    
    Text = wdata
    'Text = Left(Trim$(wdata), Len(Trim$(wdata)) - 1)
    
End Function

Function GETCARR(ByVal Fmlptr&, Field As Long, idx As Long, Text As String, Flen As Long)

    Dim wdata As String
    Dim Ret As Long
    Dim getptr As Long
    Dim err_ret As Integer
    
    wdata = String$(Flen, Chr$(32))
    Ret = Fget32(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, Flen&)
'    getptr = CFgetalloc32(ByVal Fmlptr&, ByVal Field, idx, ByVal 6, Flen&)
'    ret = lstrcpy(ByVal wdata$, ByVal getptr&)
    
    GETCARR = Ret
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fget32", 3)
        Exit Function
    End If
    
    Text = wdata
    'Text = Left(Trim$(wdata), Len(Trim$(wdata)) - 1)
    
End Function

Function PUTSTR(ByVal Fmlptr&, Field As Long, idx As Long, Text As String) As Integer

    Dim wdata As String * 1024
    Dim Ret As Integer
    Dim lret As Long
    Dim err_ret As Integer
    
    'wdata = "minam"
    wdata = String$(1024, Chr$(0))
    lret = lstrcpy(ByVal wdata, ByVal Text)
    Ret = Fchg32(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, 0)

    PUTSTR = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fchg32", 3)
    End If
    
End Function

Function PUTSTR16(ByVal Fmlptr&, Field As Long, idx As Long, Text As String) As Integer

    Dim wdata As String * 1024
    Dim Ret As Integer
    Dim lret As Long
    Dim err_ret As Integer
    
    'wdata = "minam"
    wdata = String$(1024, Chr$(0))
    lret = lstrcpy(ByVal wdata, ByVal Text)
    Ret = Fchg(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, 0)

    PUTSTR16 = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fchg", 3)
    End If
    
End Function
    

Function PUTCARR(ByVal Fmlptr&, Field As Long, idx As Long, Text As String, Flen As Long) As Integer

    Dim wdata As String
    Dim Ret As Integer
    Dim lret As Long
    Dim err_ret As Integer
    
    wdata = String$(Flen, Chr$(0))
    lret = lstrcpy(ByVal wdata, ByVal Text)
    Ret = Fchg32(ByVal Fmlptr&, ByVal Field, idx, ByVal wdata$, Flen)

    PUTCARR = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fchg32", 3)
    End If
    
End Function
' *************************************************************************************
'   공통 Error Process 모듈
'
'   Fmlptr  : FML buffer pointer    (long type)
'   service : call 함수명           (string type)
'   msg_type: Error Massage 형태    (integer type)
'             0    -> ATMI 메시지 출력
'             1    -> ATMI 메시지 + STATLIN 메시지 출력
'             2    -> ATMI 메시지 + STATLIN 메시지 + SQL code 출력
'             else -> call 함수명 출력
' *************************************************************************************
Function ErrorMsg(ByVal Fmlptr&, service As String, msg_type As Integer) As Integer

    Dim lret As Long
    Dim Ret As Long
    Dim ErrMsgCl As String
    Dim ErrMsgSv As String
    Dim tp_err_no As Long
    Dim errptr As Long
    Dim Tpurcode As Long
    
    ErrMsgCl = String$(100, Chr$(32))
    ErrMsgSv = String$(100, Chr$(32))
    
    Select Case msg_type
        Case 0              '< TPINIT, ,TPBEGIN, TPALLOC ...>
            ' ATMI error 메세지
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            Ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            Screen.MousePointer = 1         '정상화
            Save_Raw_Data service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1)
        Case 1              '< TPCALL fail ...>
            ' ATMI error 메세지
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            Ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' 서버 메세지 (STATLIN)
            lret = Fvals32(ByVal Fmlptr&, ByVal STATLIN, 0)
            Ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            Screen.MousePointer = 1         '정상화
            Save_Raw_Data service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1))
        Case 2              '< TPCALL fail: SQL error시...>
            ' ATMI error 메세지
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            Ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' 서버 메세지 (STATLIN)
            lret = Fvals32(ByVal Fmlptr&, ByVal STATLIN, 0)
            Ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode값을 받는다.
            Tpurcode = gettpurcode()
            Screen.MousePointer = 1         '정상화
            Save_Raw_Data service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
        Case 3              '< TPCALL fail: SQL error시...>
            ' ATMI error 메세지
            tp_err_no = getFerror32()
            errptr = Fstrerror32(tp_err_no)
            Ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' 서버 메세지 (STATLIN)
            'lret = FVALS32(ByVal Fmlptr&, ByVal STATLIN, 0)
            'ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode값을 받는다.
            'Tpurcode = gettpurcode()
            'Screen.MousePointer = 1         '정상화
           ' MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
             Save_Raw_Data service & " Failed." + ErrMsgCl
        Case 4              '< TPCALL fail: SQL error시...>
            ' ATMI error 메세지
            tp_err_no = getFerror32()
            errptr = Fstrerror32(tp_err_no)
            Ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' 서버 메세지 (STATLIN)
            'lret = FVALS32(ByVal Fmlptr&, ByVal STATLIN, 0)
            'ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode값을 받는다.
            'Tpurcode = gettpurcode()
            'Screen.MousePointer = 1         '정상화
           ' MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
             Save_Raw_Data service & " Failed." + ErrMsgCl
        Case Else           '< FML Function fail시 ...>
            Screen.MousePointer = 1         '정상화
            Save_Raw_Data service & " Failed."
    End Select
    
End Function

' ****************************************************
'   실수형 FML Field data를 string 타입으로 변환하여
'   return
'
'   Fmlptr  : FML buffer pointer    (long type)
'   Field   : FML field name        (string type)
'   idx     : FML occur             (long type)
'   Text    : 리턴될 data
'
' ****************************************************
Function GETDBL(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As Double
    Dim Ret As Integer
    Dim err_ret As Integer
    
    Ret = Fget32(ByVal Fmlptr&, ByVal Field, idx, wdata, 0)
    
    GETDBL = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fget32", 3)
        Exit Function
    End If
    
    Text = Str(wdata)
    
End Function

' ****************************************************
'   정수형 FML Field data를 string 타입으로 변환하여
'   return
'
'   Fmlptr  : FML buffer pointer    (long type)
'   Field   : FML field name        (string type)
'   idx     : FML occur             (long type)
'   Text    : 리턴될 data
'
' ****************************************************
Function GETLONG(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As Long
    Dim Ret As Integer
    Dim err_ret As Integer
    
    Ret = Fget32(ByVal Fmlptr&, ByVal Field, idx, wdata, 0)
    
    GETLONG = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fget32", 3)
        Exit Function
    End If
    
    Text = Str(wdata)
    
End Function


' ****************************************************
'   실수형 FML Field 에 string 으로 전달된 data를
'           전달
'
'   Fmlptr  : FML buffer pointer    (long type)
'   Field   : FML field name        (string type)
'   idx     : FML occur             (long type)
'   Text    : 리턴될 data
'
' ****************************************************
Function PUTDBL(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As Double
    Dim Ret As Integer
    Dim err_ret As Integer
    
    wdata = CDbl(Val(Trim(Text)))
    Ret = Fchg32(ByVal Fmlptr&, ByVal Field, idx, wdata, 0)

    PUTDBL = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fchg32", 3)
    End If

End Function

' ****************************************************
'   정수형 FML Field 에 string 으로 전달된 data를
'           전달
'
'   Fmlptr  : FML buffer pointer    (long type)
'   Field   : FML field name        (string type)
'   idx     : FML occur             (long type)
'   Text    : 리턴될 data
'
' ****************************************************
Function PUTLONG(ByVal Fmlptr&, Field As Long, idx As Long, Text As String)

    Dim wdata As Long
    Dim Ret As Integer
    Dim err_ret As Integer
    
    wdata = CDbl(Val(Trim(Text)))
    Ret = Fchg32(ByVal Fmlptr&, ByVal Field, idx, wdata, 0)

    PUTLONG = Ret
    
    If Ret = -1 Then
        err_ret = ErrorMsg(ByVal Fmlptr&, "Fchg32", 3)
    End If
    
End Function


