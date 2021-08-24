Attribute VB_Name = "Lockchk"
Option Explicit

Global GstrLockPtno             As String * 8
Global GstrLockRemark           As String

Sub OpdOcs_Lock_Delete()

    Dim strSql1             As String
    
    strSql1 = "DELETE FROM TWOCS_OLOCK "
    strSql1 = strSql1 & "WHERE  Ptno   = '" & GstrLockPtno & "'"

    Result = RdoExecute1(strSql1)

End Sub

Function OpdOcs_Lock_Insert(ByVal ArgGbJob As String) As String

    Dim strSql1             As String
    Dim strSql2             As String
    Dim Rs                  As rdoResultset
    Dim strRemark           As String
    Dim strEntTime          As String
    
    GoSub Sql_Stat_SET
    
    RdoDb.BeginTrans
    
    Do
        Result = RdoExecute1("LOCK TABLE TWOCS_OLOCK IN EXCLUSIVE MODE")
    Loop Until Result = 0
    
    Result = RdoOpenSet(Rs, strSql2)

    If Rowindicator > 0 And RdoGetString(Rs, "GbJob", 0) <> ArgGbJob Then   ' ArgGbJob : 1=OCS, 2=SUNAP
        strRemark = RdoGetString(Rs, "Remark", 0)
        strEntTime = RdoGetString(Rs, "EntTime", 0)
        RdoDb.CommitTrans
        GoSub LOCK_CHECK
        OpdOcs_Lock_Insert = "NO"
    Else
        OpdOcs_Lock_Insert = "OK"
        GoSub LOCK_INSERT
        RdoDb.CommitTrans
    End If
    
    RdoCloseSet Rs
    
Exit Function

'/------------------------------------------------------------------------------------

Sql_Stat_SET:           'SQL ���� SET

    strSql2 = "    SELECT Remark, TO_CHAR(EntDate,'YYYY-MM-DD HH24:MI') EntTime, GbJob "
    strSql2 = strSql2 & " FROM TWOCS_OLOCK "
    strSql2 = strSql2 & "WHERE Ptno   = '" & GstrLockPtno & "' "
    strSql2 = strSql2 & "  AND ROWNUM = 1 "
    
    Return

'/------------------------------------------------------------------------------------

LOCK_CHECK:             'Locking Process "NO"
        
    GstrMsgTitle = "����!"
    GstrMsgList = ""
    GstrMsgList = GstrMsgList & "�۾����� : " & strRemark & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "���۽ð� : " & strEntTime & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "����Ŀ� �ٽ� �۾��� �Ͻðų�, "
    GstrMsgList = GstrMsgList & "�ٸ�ȯ�ڿ� ���� �۾��� �Ͻʽÿ�!"

    MsgBox GstrMsgList, vbExclamation, GstrMsgTitle
    
    Return

'/------------------------------------------------------------------------------------

LOCK_INSERT:            'Locking Process "OK"
    
    strSql2 = "DELETE FROM TWOCS_OLOCK "
    strSql2 = strSql2 & "WHERE  Ptno   = '" & GstrLockPtno & "'"
    Result = RdoExecute1(strSql2)

    strSql1 = "INSERT INTO TWOCS_OLOCK (Ptno,Remark,EntDate,GbJob) "
    strSql1 = strSql1 & "VALUES('" & GstrLockPtno & "', '" & GstrLockRemark & "'," _
                      & "SYSDATE,'" & ArgGbJob & "') "
    Result = RdoExecute(strSql1)
    
    Return

End Function

