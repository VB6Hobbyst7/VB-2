Attribute VB_Name = "Lockchk"
Option Explicit

Global GstrLockPtno             As String * 8
Global GstrLockRemark           As String

Sub OpdOcs_Lock_Delete()
    
    adoConnect.BeginTrans
    
    GstrSql = "          DELETE  TWOCS_OLOCK              "
    GstrSql = GstrSql & " WHERE  Ptno   = '" & P.PTNO & "'"
    Result = AdoExecute1(GstrSql)
    
    If Result = -1 Then
        adoConnect.RollbackTrans
        MsgBox "LOCK DELETE ����.����� �������", , "DELETE ����"
    Else
        adoConnect.CommitTrans
    End If
    
End Sub


Function OpdOcs_Lock_Insert() As String

    Dim strSql1             As String
    Dim rs                  As ADODB.Recordset
    
    GoSub Sql_Stat_SET
    
'    Do
'        Result = AdoExecute1("LOCK TABLE TW_MIS_OCS.TWOCS_OLOCK IN EXCLUSIVE MODE")
'    Loop Until Result = 0
'
    Result = AdoOpenSet(rs, strSQL)

    If rowindicator > 0 Then
        GoSub LOCK_CHECK
    Else
        OpdOcs_Lock_Insert = "OK"
        GoSub LOCK_INSERT
    End If
    
    AdoCloseSet rs
    
    Exit Function


'/------------------------------------------------------------------------------------

Sql_Stat_SET:           'SQL ���� SET

    strSQL = "        SELECT  Remark, EntDate                "
    strSQL = strSQL & " FROM TWOCS_OLOCK                     "
    strSQL = strSQL & "WHERE Ptno   = '" & GstrLockPtno & "' "
    strSQL = strSQL & "  AND ROWNUM < 2                      "
    
    Return


'/------------------------------------------------------------------------------------

LOCK_CHECK:             'Locking Process "NO"
        
    GstrMsgTitle = "���� !"
    GstrMsgList = ""
    GstrMsgList = GstrMsgList & "�۾����� : " & AdoGetString(rs, "Remark", 0) & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "���۽ð� : " & AdoGetString(rs, "EntDATE", 0) & Chr(13)
    GstrMsgList = GstrMsgList & Chr(13) & "�۾������� ��� ����ϼ̴� ���� �ƴϸ�  "
    GstrMsgList = GstrMsgList & "�ٸ�ȯ�ڿ� ���� �۾��� �Ͻʽÿ� !!  " & Chr(13) & Chr(13)
    GstrMsgList = GstrMsgList & "�׷��� ����Ͻðڴٸ� Ȯ�� ��ư�� ���� �ֽʽÿ�."

    If MsgBox(GstrMsgList, vbOKCancel, GstrMsgTitle) = vbOK Then
        GoSub LOCK_DELETE
        OpdOcs_Lock_Insert = "OK"
        GoSub LOCK_INSERT
    Else
        OpdOcs_Lock_Insert = "NO"
    End If
    
    Return

'/------------------------------------------------------------------------------------
LOCK_DELETE:
    
    adoConnect.BeginTrans
    
    strSql1 = "DELETE TWOCS_OLOCK WHERE Ptno = '" & GstrLockPtno & "'"
    Result = AdoExecute(strSql1)
    
    If Result = -1 Then
        adoConnect.RollbackTrans
        MsgBox "LOCK DELETE ����.����� �������", , "DELETE ����"
    Else
        adoConnect.CommitTrans
    End If
    
    Return

'/------------------------------------------------------------------------------------

LOCK_INSERT:            'Locking Process "OK"
    
    adoConnect.BeginTrans
    
    strSql1 = "INSERT INTO TWOCS_OLOCK (Ptno,Remark,EntDate) "
    strSql1 = strSql1 & "VALUES('" & GstrLockPtno & "', '" & GstrLockRemark & "',SYSDATE) "
    Result = AdoExecute(strSql1)
    If Result = -1 Then
        adoConnect.RollbackTrans
        MsgBox "LOCK INSERT ����.����� �������", , "INSERT ����"
    Else
        adoConnect.CommitTrans
    End If
    
    Return
    
End Function

