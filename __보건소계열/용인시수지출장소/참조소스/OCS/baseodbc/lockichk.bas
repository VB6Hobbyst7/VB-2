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
        GstrMsgTitle = "����� ���� ��� !"
        GstrMsgList = ""
        GstrMsgList = GstrMsgList & "�۾��ڸ� : " & AdoGetString(rs1, "UserName", 0) & Chr(13)
        GstrMsgList = GstrMsgList & "�۾����� : " & AdoGetString(rs1, "Remark", 0) & Chr(13)
        GstrMsgList = GstrMsgList & "���۽ð� : " & AdoGetString(rs1, "WRTTIME", 0) & Chr(13)
        GstrMsgList = GstrMsgList & Chr(13) & "����Ŀ� �ٽ� �۾��� �Ͻʽÿ� !!" & Chr(13)
        If Left(AdoGetString(rs1, "Remark", 0), 5) <> "�����꼭" Then
        GstrMsgList = GstrMsgList & Chr(13) & "�׷��� ����Ͻðڴٸ� Ȯ�� ��ư�� ���� �ֽʽÿ�.!!"
            If MsgBox(GstrMsgList, vbOKCancel, GstrMsgTitle) = vbOK Then
                Call IpdOcs_Lock_Delete
            Else
                IpdOcs_Lock_Insert = "NO"
                Exit Function
            End If
        Else
            IpdOcs_Lock_Insert = "NO"
            MsgBox "�����꼭 �ߺ� ���Դϴ�. �����迡 Ȯ�� �� �۾��Ͻñ� �ٶ��ϴ�.", , "�����꼭 LOCK Ȯ��"
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
    strSQL = strSQL & "'Order �Է����Դϴ�', SYSDATE)  " & vbLf
    
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
    
    GstrMsgTitle = "���� !"
    GstrMsgList = ""
    GstrMsgList = GstrMsgList & "�۾��ڸ� : " & AdoGetString(rs1, "UserName", 0) & Chr(13)
    GstrMsgList = GstrMsgList & "�۾����� : " & AdoGetString(rs1, "Remark", 0) & Chr(13)
    GstrMsgList = GstrMsgList & "���۽ð� : " & AdoGetString(rs1, "Jtime", 0) & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "����Ŀ� �ٽ� �۾��� �Ͻðų� "
    GstrMsgList = GstrMsgList & Chr(13) & "�ٸ�ȯ�ڿ� ���� �۾��� �Ͻʽÿ� !"

    MsgBox GstrMsgList, , GstrMsgTitle
    
    rs1.Close
    Set rs1 = Nothing
    
    Return

End Function

