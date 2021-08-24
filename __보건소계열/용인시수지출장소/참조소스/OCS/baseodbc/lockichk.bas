Attribute VB_Name = "LockIchk"
Option Explicit

Global GstrLockPtno             As String * 8

Sub IpdOcs_Lock_Delete()
        
    strSQL = " DELETE   FROM   TW_MIS_PMPA.TWIPD_LOCK"
    strSQL = strSQL & " WHERE  Ptno   = '" & Trim(UCase(GstrLockPtno)) & "' "
    strSQL = strSQL & "   AND  Seqno  = 0   "
    strSQL = strSQL & "   AND  GbData = 'I' "

    Result = AdoExecute(strSQL)

End Sub


Function IpdOcs_Lock_Insert() As String

    Dim rs1         As ADODB.Recordset
    
    strSQL = "        SELECT *                        " & vbLf
    strSQL = strSQL & " FROM TW_MIS_PMPA.TWIPD_LOCK   " & vbLf
    strSQL = strSQL & "WHERE Ptno   = " & VarToStr(GstrLockPtno)
    strSQL = strSQL & "  AND Seqno  =  0                    " & vbLf
    strSQL = strSQL & "  AND GbData = 'I'                   " & vbLf
    
    Result = AdoOpenSet(rs1, strSQL)
    
    If rowindicator > 0 Then
        GstrMsgTitle = "전산실 연락 요망 !"
        GstrMsgList = ""
        GstrMsgList = GstrMsgList & "작업자명 : " & AdoGetString(rs1, "UserName", 0) & Chr(13)
        GstrMsgList = GstrMsgList & "작업내용 : " & AdoGetString(rs1, "Remark", 0) & Chr(13)
        GstrMsgList = GstrMsgList & "시작시간 : " & AdoGetString(rs1, "WRTTIME", 0) & Chr(13)
        GstrMsgList = GstrMsgList & Chr(13) & "잠시후에 다시 작업을 하십시요 !!" & Chr(13)
        If Left(AdoGetString(rs1, "Remark", 0), 5) <> "퇴원계산서" Then
        GstrMsgList = GstrMsgList & Chr(13) & "그래도 사용하시겠다면 확인 버튼을 눌러 주십시요.!!"
            If MsgBox(GstrMsgList, vbOKCancel, GstrMsgTitle) = vbOK Then
                Call IpdOcs_Lock_Delete
            Else
                IpdOcs_Lock_Insert = "NO"
                Exit Function
            End If
        Else
            IpdOcs_Lock_Insert = "NO"
            MsgBox "퇴원계산서 발부 중입니다. 조정계에 확인 후 작업하시기 바랍니다.", , "퇴원계산서 LOCK 확인"
            Exit Function
        End If
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    strSQL = " INSERT INTO TW_MIS_PMPA.TWIPD_LOCK(GbData,Ptno,SeqNo,UserName,Remark,WrtTime) " & vbLf
    strSQL = strSQL & " VALUES('I',                     " & vbLf
    strSQL = strSQL & VarToComma(GstrLockPtno)
    strSQL = strSQL & " 0, "
    strSQL = strSQL & VarToComma(GstrPassName)
    strSQL = strSQL & "'Order 입력중입니다', SYSDATE)  " & vbLf
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = -1 Then
        adoConnect.RollbackTrans
        GoSub LOCK_CHECK
        IpdOcs_Lock_Insert = "NO"
    Else
        adoConnect.CommitTrans
        IpdOcs_Lock_Insert = "OK"
    End If

    Exit Function


'/------------------------------------------------------------------------------------------------------/

LOCK_CHECK:             'Locking Process "NO"
        
    strSQL = " SELECT UserName, Remark, WrtTime Jtime "
    strSQL = strSQL & "  FROM TW_MIS_PMPA.TWIPD_LOCK"
    strSQL = strSQL & " WHERE Ptno   = '" & Trim(UCase(GstrLockPtno)) & "' "
    strSQL = strSQL & "   AND GbData = 'I' "
    strSQL = strSQL & "   AND Seqno  = 0   "
    
    Result = AdoOpenSet(rs1, strSQL)
    
    GstrMsgTitle = "주의 !"
    GstrMsgList = ""
    GstrMsgList = GstrMsgList & "작업자명 : " & AdoGetString(rs1, "UserName", 0) & Chr(13)
    GstrMsgList = GstrMsgList & "작업내용 : " & AdoGetString(rs1, "Remark", 0) & Chr(13)
    GstrMsgList = GstrMsgList & "시작시간 : " & AdoGetString(rs1, "Jtime", 0) & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "잠시후에 다시 작업을 하시거나 "
    GstrMsgList = GstrMsgList & Chr(13) & "다른환자에 대한 작업을 하십시오 !"

    MsgBox GstrMsgList, , GstrMsgTitle
    
    rs1.Close
    Set rs1 = Nothing
    
    Return

End Function

