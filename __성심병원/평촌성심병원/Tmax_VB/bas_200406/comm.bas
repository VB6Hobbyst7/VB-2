Attribute VB_Name = "comm"
Option Explicit
'***************************************************************
'   Function: Display Error Msg of ATMI Function For FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Service : ATMI API Function
'***************************************************************
Function FdlErrorMsg(service As String) As Integer
                                                                                  
    Dim lret As Long
    Dim ret As Long
    Dim ErrMsg As String
    Dim Errmsgs As String
    Dim fdl_err_no As Long
    Dim errptr As Long
    Dim tpurcode As Long
    Dim msgString As String
                                                                                  
    ' Error Msg of Client Program
    fdl_err_no = getfberrno()
    errptr = fbstrerror(ByVal fdl_err_no&)
                                                                                  
    ErrMsg = Space$(100)
                                                                                  
    ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
                                                                                  
    msgString = service & " Failed." & Chr(13)
    If Len(Trim(ErrMsg)) > 0 Then
        msgString = msgString & Left(ErrMsg, Len(Trim(ErrMsg)) - 1) & Chr(13)
    End If
                                                                                  
    MsgBox msgString
                                                                                  
End Function
'***************************************************************
'   Function: Put The Integer Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   iData   : INTEGER DATA
'***************************************************************
Function PUTINT(ByVal Fdlptr&, Field As String, idx As Long, iData As Long)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, iData, 0)

    PUTINT = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function
'***************************************************************
'   Function: Put The Long Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   iData   : LONG DATA
'***************************************************************
Function PUTLONG(ByVal Fdlptr&, Field As String, idx As Long, lData As Long)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, lData, 0)

    PUTLONG = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function
'***************************************************************
'   Function: Put The Double Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   iData   : DOUBLE DATA
'***************************************************************
Function PUTDOUBLE(ByVal Fdlptr&, Field As String, idx As Long, dData As Double)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, dData, 0)

    PUTDOUBLE = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function
'***************************************************************
'   Function: Put The String Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   text    : STRING DATA
'***************************************************************
Function PUTVAR(ByVal Fdlptr&, Field As String, idx As Long, text As String) As Integer

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByVal text$, 0)
   
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function
'***************************************************************
'   Function: Put The String or Image Data Into String or CARRAY Buffer
'   Fdlptr  : STRING OR CARRAY BUFFER POINTER
'   text    : STRING DATA
'   datalen : DATA LENGTH
'***************************************************************
Function PUTCAR(ByVal Fdlptr&, text As String, datalen As Long)

    Dim ret As Long
    Dim err_ret As Integer
    
    ret = lstrcpyn(ByVal Fdlptr&, ByVal text$, ByVal datalen&)
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("lstrcpyn")
    End If
    
End Function
'***************************************************************
'   Function: Put The Image Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   text    : STRING DATA
'***************************************************************
Function PUTCAR_BA(ByVal Fdlptr&, Field As String, idx As Long, ByteArray() As Byte, datalen As Long) As Integer

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
   
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByteArray(0), ByVal datalen)
 '   ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByVal snddata$, ByVal datalen)
   
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function
'***************************************************************
'   Function: Get The Integer Data
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   iData    : LOCAL VARIABLE
'***************************************************************
Function GETINT(ByVal Fdlptr&, Field As String, idx As Long, iData As Long)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbget_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, iData, 0)
    
    GETINT = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbget_tu")
    End If
    
End Function
'***************************************************************
'   Function: Get The Long Data
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   lData    : LOCAL VARIABLE
'***************************************************************
Function GETLONG(ByVal Fdlptr&, Field As String, idx As Long, lData As Long)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbget_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, lData, 0)
    
    GETLONG = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbget_tu")
    End If
    
End Function
'***************************************************************
'   Function: Get The Double Data
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   dData    : LOCAL VARIABLE
'***************************************************************
Function GETDOUBLE(ByVal Fdlptr&, Field As String, idx As Long, dData As Double)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbget_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, dData, 0)
    
    GETDOUBLE = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbget_tu")
    End If
    
End Function
'***************************************************************
'   Function: Get The String Data From FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   text    : LOCAL VARIABLE
'             No Text property of Control
'***************************************************************
Function GETVAR(ByVal Fdlptr&, Field As String, idx As Long, text As String)

'    Dim data As String * 512
    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    Dim location As Long
    
    text = String$(1024, Chr$(0))

    ret = fbget_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByVal text$, 0)
    
    location = InStr(text, Chr$(0))
    If (location > 0) Then
        text = Trim(Left(text, location - 1))
    End If
    
    GETVAR = ret
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbget_tu")
    End If
    
End Function
'***************************************************************
'   Function: Get The String Data From String Buffer
'   Fdlptr  : STRING BUFFER POINTER
'   text    : LOCAL VARIABLE
'             No Text property of Control
'***************************************************************
Function GETSTR(ByVal Fdlptr&, text As String)

'    Dim data As String * 512
    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    Dim location As Long
    
    text = String$(1024, Chr$(0))

    ret = vb_getstr(ByVal Fdlptr&, ByVal text$)
    
    location = InStr(text, Chr$(0))
    If (location > 0) Then
        text = Trim(Left(text, location - 1))
    End If
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("vb_getstr")
    End If
    
End Function
'***************************************************************
'   Function: Get The Carray Data and DataLen From CARRAY Buffer
'   Fdlptr  : STRING BUFFER POINTER
'   text    : LOCAL VARIABLE
'             No Text property of Control
'   datalen : LOCAL VARIABLE FOR DATA LENGTH
'***************************************************************
Function GETCAR(ByVal Fdlptr&, text As String, datalen As Long)

    Dim ret As Long
    Dim err_ret As Long
  
    text = String$(1024, Chr$(0))

    ret = vb_getcar(ByVal Fdlptr&, ByVal text$, ByVal datalen&)
    
    text = Trim(Left(text, datalen))
    
    If ret = -1 Then
        err_ret = FdlErrorMsg("vb_getcar")
    End If
    
End Function
'***************************************************************
'   Function: Get The Image Data From FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   text    : LOCAL VARIABLE
'             No Text property of Control
'***************************************************************
Function GETCAR_BA(ByVal Fdlptr&, Field As String, idx As Long, ByRef ByteArray() As Byte, datalen As Long)

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    Dim location As Long
    
    ret = fbget_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByteArray(0), datalen)
       
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbget_tu")
    End If
    
End Function
'***************************************************************
'   Function: Display StrErr plus tperrno msg
'***************************************************************
Sub TmaxError(StrErr As String)
    Dim tperrno As Integer
    Dim tpurcode As Integer
    Dim sptr As Long
    Dim tmaxstr$
    Dim ret As Long
    Dim MSG$

    tperrno% = gettperrno()
    tpurcode% = gettpurcode()
    sptr& = tpstrerror(tperrno%)
    tmaxstr$ = String$(100, Chr$(0))
    ret& = lstrcpy(ByVal tmaxstr$, ByVal sptr&)
    MSG$ = StrErr$ + ": " + tmaxstr$
    MsgBox MSG$
End Sub
'***************************************************************
'   Function: Making the start structure
'***************************************************************
Function FilltpstartBuf(sndbufp As Long, startinfop As tpstart_t)
    Static ptr As Long
    Dim ret As Long
    Dim slen As Long
    Dim X  As String
    Const DATALEN = 18
    

    ptr = sndbufp
    ret& = lstrcpyn(ByVal ptr&, ByVal startinfop.usrname, DATALEN)
    
    ptr = ptr + DATALEN
    ret& = lstrcpyn(ByVal ptr&, ByVal startinfop.cltname, DATALEN)
    
    ptr = ptr + DATALEN
    ret& = lstrcpyn(ByVal ptr&, ByVal startinfop.dompwd, DATALEN)

    ptr = ptr + DATALEN
    ret& = lstrcpyn(ByVal ptr&, ByVal startinfop.usrpwd, DATALEN)
    
    ptr = ptr + DATALEN
    X$ = Chr(startinfop.flags) + Chr(0) + Chr(0) + Chr(0)
    slen& = 4
    ret& = lstrcpyn(ByVal ptr&, ByVal X$, ByVal slen&)

End Function
