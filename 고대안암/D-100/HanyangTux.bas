Attribute VB_Name = "HanyangTux"
Option Explicit
Dim g_ret_val As Integer


Public Function GetReceTux(argBarcode As String) As String
    
    
    Dim Ret As Long
    Dim tp_err_no As Integer
    Dim recvlen As tuxbuf
    Dim recvptr As tuxbuf
    Dim lret As Long
    Dim slen As Long
    Dim errptr As Long
    Dim ErrMsg As String
    Dim temp1 As String
    Dim temp2 As String
    Dim err_ret As Integer
    Dim strTuxCmd As String
    Dim i As Integer
    Dim j As Integer
    Dim strOrderExam As String
    
    On Error GoTo Errtux:
    
    GetReceTux = ""
    strOrderExam = ""
    g_ret_val = tuxreadenv(App.Path & "\envfile", "med1")
    If g_ret_val <> 0 Then
         Save_Raw_Data "GetReceTux 환경설정오류"
         End
    End If
    
    '**********************************************
    ' 메모리 할당
    '**********************************************
    strbuf.bufptr = tpalloc("STRING", "", 4096)
    
    If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC 실패. 에러번호 : ")
        Exit Function
    End If
        
    If tpinit(ByVal 0&) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        Ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        Save_Raw_Data "TPINIT GetReceTux 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
        Exit Function
    End If
        
    '********************************************
    ' 서비스 CALL
    '********************************************
    temp1$ = Space$(1024)
    strTuxCmd = "0" & Mid(argBarcode, 1, 10)
    
    
    temp1$ = strTuxCmd
    lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
        
    Ret = tpcall("HAMA0122", ByVal strbuf.bufptr, ByVal 0&, strbuf, recvlen, ByVal 0&)
        
    If Ret = -1 Then
           err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
           Ret = tpabort(0)
           Ret = tpfree(strbuf.bufptr)
           tpterm
           Exit Function
    End If
        
'''    temp2$ = "0         HAMA0122        OK !!!      OK !!!                                                      001610844941" & vbTab & "02417654" & vbTab & "김인수" & vbTab & "M" & vbTab & "20160928" & vbTab & "599" & vbTab & "20160928" & vbTab & "1" & vbTab & "L3142" & vbTab
    temp2$ = Space$(recvlen.bufptr)
    lret = lstrcpy(ByVal temp2$, ByVal strbuf.bufptr)
    
    Save_Raw_Data temp2$
    
    strOrderExam = ""
    
    
    If Mid(temp2$, 1, 1) = "0" Then
        temp2$ = Mid(temp2$, 101)
        i = InStr(1, temp2$, vbTab)
        j = 0
    
        ClearSpread frmInterface.vasTux
        
        While i > 0
            j = j + 1
            
            SetText frmInterface.vasTux, Trim(Mid(temp2$, 1, i - 1)), 1, j
            If j > 8 Then
                Save_Raw_Data Trim(Mid(temp2$, 1, i - 1))
                If strOrderExam = "" Then
                    strOrderExam = "'" & Trim(Mid(temp2$, 1, i - 1)) & "'"
                Else
                    strOrderExam = strOrderExam & ", '" & Trim(Mid(temp2$, 1, i - 1)) & "'"
                End If
            End If
            temp2$ = Mid(temp2$, i + 1)
            i = InStr(1, temp2$, vbTab)
            
        Wend
        
'''        While i > 0
'''            j = j + 1
'''
'''            If j > 2 Then
'''                If strOrderExam = "" Then
'''                    strOrderExam = "'" & Trim(Mid(temp2$, 1, i - 1)) & "'"
'''                Else
'''                    strOrderExam = strOrderExam & ", '" & Trim(Mid(temp2$, 1, i - 1)) & "'"
'''                End If
'''            End If
'''
'''            temp2$ = Mid(temp2$, i + 1)
'''            i = InStr(1, temp2$, vbTab)
'''        Wend

    End If
    
    
    Ret = tpfree(strbuf.bufptr)
    tpterm
    
    If Trim(strOrderExam) <> "" Then
        GetReceTux = strOrderExam
    Else
        GetReceTux = "''"
    End If
    Exit Function
    
Errtux:
    tpterm
    GetReceTux = "''"
    Exit Function
    
End Function


'''- 오더가져오기
'''
'''1.서비스 : HAMAC102
'''2.전송값(자리수) : 구분(1) + 검체년도(2) + 검체번호(7) + 검체번호1(1)
'''                   구분값은 0 (검체번호)
'''3.리턴값(자리수) : 전송값(11) + tab + 검사건수 + tab + tab + 검사코드1 + tab + 검사코드2 ....
'''                   검사코드는 검사건수 만큼 반복
Public Function GetOrderTux(argBarcode As String) As String
    
    
    Dim Ret As Long
    Dim tp_err_no As Integer
    Dim recvlen As tuxbuf
    Dim recvptr As tuxbuf
    Dim lret As Long
    Dim slen As Long
    Dim errptr As Long
    Dim ErrMsg As String
    Dim temp1 As String
    Dim temp2 As String
    Dim err_ret As Integer
    Dim strTuxCmd As String
    Dim i As Integer
    Dim j As Integer
    Dim strOrderExam As String
    
    GetOrderTux = ""
    strOrderExam = ""
    g_ret_val = tuxreadenv(App.Path & "\envfile", "med1")
    If g_ret_val <> 0 Then
         Save_Raw_Data "GetOrderTux 환경설정오류"
         End
    End If
    
    '**********************************************
    ' 메모리 할당
    '**********************************************
    strbuf.bufptr = tpalloc("STRING", "", 4096)
    
    If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC 실패. 에러번호 : ")
        Exit Function
    End If
        
    If tpinit(ByVal 0&) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        Ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        Save_Raw_Data "TPINIT GetOrderTux 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
        Exit Function
    End If
        
    '********************************************
    ' 서비스 CALL
    '********************************************
    temp1$ = Space$(1024)
    strTuxCmd = "0" & Mid(argBarcode, 1, 10)
    
    
    temp1$ = strTuxCmd
    lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
        
    Ret = tpcall("HAMAC102", ByVal strbuf.bufptr, ByVal 0&, strbuf, recvlen, ByVal 0&)
        
    If Ret = -1 Then
           err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
           Ret = tpabort(0)
           Ret = tpfree(strbuf.bufptr)
           tpterm
           Exit Function
    End If
        
            
    temp2$ = Space$(recvlen.bufptr)
    lret = lstrcpy(ByVal temp2$, ByVal strbuf.bufptr)
    
    i = InStr(1, temp2$, vbTab)
    j = 0
    
    If Mid(temp2$, 1, 1) = "0" Then
        While i > 0
            j = j + 1
            
            If j > 2 Then
                If strOrderExam = "" Then
                    strOrderExam = "'" & Trim(Mid(temp2$, 1, i - 1)) & "'"
                Else
                    strOrderExam = strOrderExam & ", '" & Trim(Mid(temp2$, 1, i - 1)) & "'"
                End If
            End If
            
            temp2$ = Mid(temp2$, i + 1)
            i = InStr(1, temp2$, vbTab)
        Wend

    End If
    
    Ret = tpfree(strbuf.bufptr)
    tpterm
    
    If Trim(strOrderExam) <> "" Then
        GetOrderTux = strOrderExam
    Else
        GetOrderTux = "''"
    End If
    

End Function

 
'''- 결과저장하기
'''1.서비스 : HAMAC111
'''2.전송값(자리수) : 장비코드(4) + tab + 사용자(7) + tab + 검사일자(8) + tab + 검체번호(10) + tab + 구분(1) + tab + 검사건수 + tab + 검사결과
'''                   구분값은 0 (검체번호)
'''                   검사결과 : 검사코드 + tab + 결과 + LF
'''                   검사결과는 검사건수 만큼 반복
'''3.리턴값 :
'''answer.sql_code  = Trim(Mid(i_mesg, 1,10))                // Sqlca.sqlcode
'''answer.fun_name  = Trim(Mid(i_mesg,11,16))                // Function Name
'''answer.tbl_name  = Trim(Mid(i_mesg,27,12))                // Table Name
'''answer.msg_desc  = Trim(Mid(i_mesg,39,60))                // Out Message
'''answer.msg_level = Trim(Mid(i_mesg, 99, 1))

Public Function InsertResultTux(argData As String) As String
    
    
    Dim Ret As Long
    Dim tp_err_no As Integer
    Dim recvlen As tuxbuf
    Dim recvptr As tuxbuf
    Dim lret As Long
    Dim slen As Long
    Dim errptr As Long
    Dim ErrMsg As String
    Dim temp1 As String
    Dim temp2 As String
    Dim err_ret As Integer
    Dim strTuxCmd As String
    Dim i As Integer
    Dim j As Integer
    Dim strOrderExam As String
    
    InsertResultTux = ""
    strOrderExam = ""
    g_ret_val = tuxreadenv(App.Path & "\envfile", "med1")
    If g_ret_val <> 0 Then
         Save_Raw_Data "InsertResultTux 환경설정오류"
         End
    End If
    
    '**********************************************
    ' 메모리 할당
    '**********************************************
    strbuf.bufptr = tpalloc("STRING", "", 4096)
    
    If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC 실패. 에러번호 : ")
        Exit Function
    End If
        
    If tpinit(ByVal 0&) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        Ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        Save_Raw_Data "TPINIT InsertResultTux 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
        Exit Function
    End If
        
    '********************************************
    ' 서비스 CALL
    '********************************************
    temp1$ = Space$(1024)
    strTuxCmd = argData
'''    Save_Raw_Data "1"
    
    temp1$ = strTuxCmd
    lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
    
'''    Save_Raw_Data "2"
    
    Ret = tpcall("HAMA0111", ByVal strbuf.bufptr, ByVal 0&, strbuf, recvlen, ByVal 0&)
    
'''    Save_Raw_Data "3"
    
    If Ret = -1 Then
'''           err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
           Save_Raw_Data "HAMAC111" & argData
           Ret = tpabort(0)
           Ret = tpfree(strbuf.bufptr)
           tpterm
           Exit Function
    End If
        
    
    temp2$ = Space$(recvlen.bufptr)
    lret = lstrcpy(ByVal temp2$, ByVal strbuf.bufptr)
    
'''    i = InStr(1, temp2$, vbTab)
'''    j = 0
    
    strOrderExam = Trim(Mid(temp2$, 39, 60))
    
'''    If Mid(temp2$, 1, 1) = "0" Then
'''        While i > 0
'''            j = j + 1
'''
'''            If j > 2 Then
'''                If strOrderExam = "" Then
'''                    strOrderExam = "'" & Trim(Mid(temp2$, 1, i - 1)) & "'"
'''                Else
'''                    strOrderExam = strOrderExam & ", '" & Trim(Mid(temp2$, 1, i - 1)) & "'"
'''                End If
'''            End If
'''
'''            temp2$ = Mid(temp2$, i + 1)
'''            i = InStr(1, temp2$, vbTab)
'''        Wend
'''
'''    End If
    
    Ret = tpfree(strbuf.bufptr)
    tpterm
    
    If Trim(strOrderExam) <> "" Then
        InsertResultTux = strOrderExam
    Else
        InsertResultTux = ""
    End If
    

End Function

'''- login
'''
'''1.서비스 : HAMA0105
'''2.전송값(자리수) : 사번(7) & 패스워드(8) & 일자(8)

Public Function LoginTux(argData As String) As String
    
    Dim Ret As Long
    Dim tp_err_no As Integer
    Dim recvlen As tuxbuf
    Dim recvptr As tuxbuf
    Dim lret As Long
    Dim slen As Long
    Dim errptr As Long
    Dim ErrMsg As String
    Dim temp1 As String
    Dim temp2 As String
    Dim err_ret As Integer
    Dim strTuxCmd As String
    Dim i As Integer
    Dim j As Integer
    Dim strOrderExam As String
    
    LoginTux = ""
    strOrderExam = ""
    g_ret_val = tuxreadenv(App.Path & "\envfile", "med1")
    If g_ret_val <> 0 Then
         Save_Raw_Data "LoginTux 환경설정오류"
         End
    End If
    
    '**********************************************
    ' 메모리 할당
    '**********************************************
    strbuf.bufptr = tpalloc("STRING", "", 4096)
    
    If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC 실패. 에러번호 : ")
        Exit Function
    End If
        
    If tpinit(ByVal 0&) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        Ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        Save_Raw_Data "TPINIT LoginTux 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
        Exit Function
    End If
        
    '********************************************
    ' 서비스 CALL
    '********************************************
    temp1$ = Space$(1024)
    strTuxCmd = argData
    
    
    temp1$ = strTuxCmd
    lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
        
    Ret = tpcall("HAMA0105", ByVal strbuf.bufptr, ByVal 0&, strbuf, recvlen, ByVal 0&)
        
    If Ret = -1 Then
           err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
           Ret = tpabort(0)
           Ret = tpfree(strbuf.bufptr)
           tpterm
           Exit Function
    End If
        
            
    temp2$ = Space$(recvlen.bufptr)
    lret = lstrcpy(ByVal temp2$, ByVal strbuf.bufptr)
    
    i = InStr(1, temp2$, vbTab)
    j = 0
    
    Save_Raw_Data temp2$
    
    If Mid(temp2$, 1, 1) = "0" Then
        While i > 0
            j = j + 1
            
            If j = 2 Then
'''                If strOrderExam = "" Then
                    strOrderExam = Trim(Mid(temp2$, 1, i - 1))
'''                Else
'''                    strOrderExam = strOrderExam & ", '" & Trim(Mid(temp2$, 1, i - 1)) & "'"
'''                End If
            End If
            
            temp2$ = Mid(temp2$, i + 1)
            i = InStr(1, temp2$, vbTab)
        Wend

    End If
    
    Ret = tpfree(strbuf.bufptr)
    tpterm
    
    If Trim(strOrderExam) <> "" Then
        LoginTux = strOrderExam
    Else
        LoginTux = ""
    End If
    

End Function
