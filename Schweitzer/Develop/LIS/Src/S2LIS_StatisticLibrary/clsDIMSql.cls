VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIMSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*--------------------------------------------------------
'*  ��� : �ش������� �ٹ�����, ���������� ����
'*  Parameter
'*      pEmpId : ����ID
'*  Return : �ٹ�����, ��������
'*--------------------------------------------------------
Public Function GetEmpInfo(ByVal pEmpId As String) As String
    Dim SQL As String
    
    SQL = "SELECT shiftcd, remain FROM " & T_COM006 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    
    GetEmpInfo = SQL
End Function

'*--------------------------------------------------------
'*  ��� : �ش����� ���������� ����
'*  Parameter
'*      pDt : ��¥
'*  Return : ����ID, �����̸�, �ð�, ����(���, ����, ���...)
'*--------------------------------------------------------
Public Function GetEmpFlagInfo(ByVal pDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    
    SQL1 = "SELECT A.empid, B.empnm, A.tm, A.flag "
    SQL1 = SQL1 & "FROM " & T_COM006 & " B, " & T_COM013 & " A "
    SQL1 = SQL1 & "WHERE" & DBW("A.dt=", pDt)
    SQL1 = SQL1 & "   AND A.empid=B.empid "
    SQL1 = SQL1 & "GROUP BY A.empid, B.empnm, A.tm, A.flag"
    
    SQL2 = "SELECT A.empid, B.empnm, '' AS tm, A.flag "
    SQL2 = SQL2 & "FROM " & T_COM006 & " B, " & T_COM014 & " A "
    SQL2 = SQL2 & "WHERE" & DBW("A.dt=", pDt)
    SQL2 = SQL2 & "   AND A.empid=B.empid "
    SQL2 = SQL2 & "GROUP BY A.empid, B.empnm, '', A.flag"
    
    SQL3 = "SELECT A.empid, B.empnm, '' AS tm, A.flag "
    SQL3 = SQL3 & "FROM " & T_COM006 & " B, " & T_COM015 & " A "
    SQL3 = SQL3 & "WHERE '" & pDt & "' BETWEEN fromdt AND todt"
    SQL3 = SQL3 & "     AND A.empid=B.empid "
    SQL3 = SQL3 & "GROUP BY A.empid, B.empnm, '', A.flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " UNION ALL " & SQL3 & " ORDER BY empid, flag ASC"
    GetEmpFlagInfo = SQL1
End Function

'*--------------------------------------------------------
'*  ��� : �ش������� ����ٹ� �ϼ��� ����(������� ����)
'*  Parameter
'*      pEmpId : ����ID
'*      pFromDt : ������
'*      pToDt : ������
'*      pFromTm : �ٹ����۽ð�
'*      pToTm : �ٹ�����ð�
'*  Return : ����ٹ� �ϼ�
'*--------------------------------------------------------
Public Function GetNormalWorkCnt01(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                                   ByVal pFromTm As String, ByVal pToTm As String) As Integer
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT Count(empid) AS count FROM " & T_COM013 & " A "
    SQL = SQL & "WHERE" & DBW("A.empid=", pEmpId)
    SQL = SQL & "   AND A.dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "   AND A.flag='1'"
    SQL = SQL & "   AND A.tm<='" & pFromTm & "'"
    SQL = SQL & "   AND '" & pToTm & "' <= (SELECT B.tm FROM " & T_COM013 & " B "
    SQL = SQL & "                           WHERE" & DBW("B.empid=", pEmpId)
    SQL = SQL & "                               AND B.dt=A.dt"
    SQL = SQL & "                               AND B.flag='3')"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetNormalWorkCnt01 = CInt(RS.Fields("count").Value & "")
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "����"
End Function

'*--------------------------------------------------------
'*  ��� : �ش������� ����Ͽ� ����ٹ��� �ߴ��� Ȯ��
'*  Parameter
'*      pEmpId : ����ID
'*      pDt : ��¥(�����)
'*      pFromTm : �ٹ����۽ð�
'*      pToTm : �ٹ�����ð�
'*  Return : True / False
'*--------------------------------------------------------
Public Function GetNormalWorkCnt02(ByVal pEmpId As String, ByVal pDt As String, _
                                   ByVal pFromTm As String, ByVal pToTm As String) As Boolean
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT A.empid FROM " & T_COM013 & " A "
    SQL = SQL & "WHERE" & DBW("A.empid=", pEmpId)
    SQL = SQL & "   AND" & DBW("A.dt=", pDt)
    SQL = SQL & "   AND A.flag='1'"
    SQL = SQL & "   AND A.tm<='" & pFromTm & "'"
    SQL = SQL & "   AND '" & pToTm & "' <= (SELECT B.tm FROM " & T_COM013 & " B "
    SQL = SQL & "                           WHERE" & DBW("B.empid=", pEmpId)
    SQL = SQL & "                               AND B.dt=A.dt"
    SQL = SQL & "                               AND B.flag='3')"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetNormalWorkCnt02 = True
    Else
        GetNormalWorkCnt02 = False
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "����"
End Function

'*--------------------------------------------------------
'*  ��� : �˻��Ⱓ���� ����, �ް��� �ִ��� ����
'*  Parameter
'*      pEmpId : ����ID
'*      pFromDt : ������
'*      pToDt : ������
'*      pFlag : ����(9), �ް�(10) Flag
'*  Return : ������, ������
'*--------------------------------------------------------
Public Function GetHoliInfo(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                            ByVal pFlag As String) As String
    Dim SQL As String
    
    SQL = "SELECT fromdt, todt FROM " & T_COM015 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId) & " "
    SQL = SQL & "   AND ((fromdt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL = SQL & "           OR todt BETWEEN '" & pFromDt & "' AND '" & pToDt & "')"
    SQL = SQL & "           OR ('" & pFromDt & "' BETWEEN fromdt AND todt"
    SQL = SQL & "           OR '" & pToDt & "' BETWEEN fromdt AND todt))"
    SQL = SQL & "   AND" & DBW("flag=", pFlag) & " "
    SQL = SQL & "ORDER BY fromdt ASC"
    
    GetHoliInfo = SQL
End Function

'*--------------------------------------------------------
'*  ��� : �ش������� �Ⱓ�� ��踦 ����
'*  Parameter
'*      pEmpId : ����ID
'*      pFromDt : ������
'*      pToDt : ������
'*  Return : �� Flag�� ���
'*--------------------------------------------------------
Public Function GetStaticsEmp(ByVal pEmpId As String, ByVal pFromDt As String, _
                                            ByVal pToDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    
    SQL1 = "SELECT flag, Count(flag) as count "
    SQL1 = SQL1 & "FROM " & T_COM013 & " "
    SQL1 = SQL1 & "WHERE" & DBW("empid=", pEmpId)
    SQL1 = SQL1 & "     AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL1 = SQL1 & "GROUP BY flag"
    
    SQL2 = "SELECT flag, Count(flag) as count "
    SQL2 = SQL2 & "FROM " & T_COM014 & " "
    SQL2 = SQL2 & "WHERE" & DBW("empid=", pEmpId)
    SQL2 = SQL2 & "     AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    SQL2 = SQL2 & "GROUP BY flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " ORDER BY flag ASC"
    GetStaticsEmp = SQL1
End Function

'*--------------------------------------------------------
'*  ��� : ��ü������ �Ⱓ�� ��踦 ����
'*  Parameter
'*      pFromDt : ������
'*      pToDt : ������
'*  Return : ����ID, �����̸�, ����(���, ����, ���� ���..)
'*--------------------------------------------------------
Public Function GetStaticsEmpAll(ByVal pFromDt As String, ByVal pToDt As String) As String
    Dim SQL1 As String
    Dim SQL2 As String
    
    SQL1 = "SELECT B.empid, B.empnm, A.flag, Count(A.flag) AS count "
    SQL1 = SQL1 & "FROM " & T_COM006 & " B, " & T_COM013 & " A "
    SQL1 = SQL1 & "WHERE dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL1 = SQL1 & "      AND A.empid=B.empid "
    SQL1 = SQL1 & "GROUP BY B.empid, B.empnm, A.flag"
    
    SQL2 = "SELECT B.empid, B.empnm, A.flag, Count(A.flag) AS count "
    SQL2 = SQL2 & "FROM " & T_COM006 & " B, " & T_COM014 & " A "
    SQL2 = SQL2 & "WHERE dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL2 = SQL2 & "      AND A.empid=B.empid "
    SQL2 = SQL2 & "GROUP BY B.empid, B.empnm, A.flag"
    
    SQL1 = SQL1 & " UNION ALL " & SQL2 & " ORDER BY empid, flag ASC"
    GetStaticsEmpAll = SQL1
End Function

'*--------------------------------------------------------
'*  ��� : �ش������� ������¥�� ����
'*  Parameter
'*      pEmpId : ����ID
'*      pFromDt : ������
'*      pToDt : ������
'*  Return : ������, �ð�
'*--------------------------------------------------------
Public Function GetLateDt(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String) _
                            As String
    Dim SQL As String
    
    SQL = "SELECT dt, tm FROM " & T_COM013 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    SQL = SQL & "       AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "       AND flag='2' "
    SQL = SQL & "ORDER BY dt ASC"
    
    GetLateDt = SQL
End Function

'*--------------------------------------------------------
'*  ��� : �ش������� ����, ����, ����, ����, ��ٳ�¥�� ����
'*  Parameter
'*      pEmpId : ����ID
'*      pFromDt : ������
'*      pToDt : ������
'*      pFlag :����(4), ����(5), ����(6), ����(7), ���(8) Flag
'*  Return : ����, ����, ����, ����, �������
'*--------------------------------------------------------
Public Function GetDate(ByVal pEmpId As String, ByVal pFromDt As String, ByVal pToDt As String, _
                            ByVal pFlag As String) As String
    Dim SQL As String
    
    SQL = "SELECT dt FROM " & T_COM014 & " "
    SQL = SQL & "WHERE" & DBW("empid=", pEmpId)
    SQL = SQL & "       AND dt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' "
    SQL = SQL & "       AND" & DBW("flag=", pFlag)
    SQL = SQL & "ORDER BY dt ASC"
    
    GetDate = SQL
End Function

'*--------------------------------------------------------
'*  ��� : �ش�Ⱓ�� ������ �ϼ��� ����
'*  Parameter
'*      PFromDt : ������
'*      pToDt : ������
'*  Return : ������ �ϼ�
'*--------------------------------------------------------
Public Function GetHolidayCount(ByVal pFromDt As String, ByVal pToDt As String) As Integer
    Dim RS As New Recordset
    Dim SQL As String
    
On Error GoTo ErrorMsg:
    SQL = "SELECT Count(holidt) AS count FROM " & T_COM017 & " "
    SQL = SQL & "WHERE holidt BETWEEN '" & pFromDt & "' AND '" & pToDt & "'"
    
    RS.Open SQL, dbconn
    If Not (RS.BOF Or RS.EOF) Then
        GetHolidayCount = CInt(RS.Fields("count").Value & "")
    End If
    
    Set RS = Nothing
    Exit Function
    
ErrorMsg:
    MsgBox Err.Description, vbOKOnly + vbCritical, "����"
End Function