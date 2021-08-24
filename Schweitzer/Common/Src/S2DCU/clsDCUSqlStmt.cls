VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDCUSqlStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+--------------------------------------------------------------------------------------+
'|  1. cls ��   : clsDCUSqlStmt
'|  2. ��  ��   : �������� ���,����,����
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.11.6
'|
'|  CopyRight(C) 2002 Pomis
'+--------------------------------------------------------------------------------------+
Option Explicit

Private mvarWorkFg As String     '��������(0), ��������(1)

Public Property Get WorkFg() As String
    WorkFg = mvarWorkFg
End Property

Public Property Let WorkFg(ByVal vNewValue As String)
    mvarWorkFg = vNewValue
End Property
'
'Public Function Get_EMPNM(ByVal EID As String) As Recordset
'    Dim strSQL As String
'
'    '�����̸��� ��������.
'    strSQL = "SELECT   empnm FROM " & T_COM006 & _
'             " WHERE " & DBW("empid", EID, 2)
'
'
'
'    Set Get_EMPNM = New Recordset
'    Get_EMPNM.RsOpen , strSQL
'
'    If Get_EMPNM.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

'Cos_hos�� �ִ� ���� ����ϱ⶧���� ����
'Public Function GetCOM006EmpNm(ByVal vEmpId As String) As String
''Get_EMPNM �̶� �м��� �ٲ�
'    Dim Rs As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT empnm FROM " & T_COM006 & _
'             " WHERE " & DBW("empid", vEmpId, 2)
'
'    Set Rs = New Recordset
'    On Error GoTo ErrTrap
'
'    Rs.Open strSQL, dbconn
'
'    If Rs.EOF Then
'        GetCOM006EmpNm = ""
'    Else
'        GetCOM006EmpNm = Rs.Fields("empnm").Value & ""
'    End If
'
'ErrTrap:
'    Set Rs = Nothing
'    MsgBox Err.Description, vbExclamation
'End Function

'Public Function Get_COM006(ByVal EID As String) As Recordset
'    Dim strSQL As String
'
'    '�����̸��� ��������.
'    strSQL = "SELECT " & F_EMPNM & " as empnm FROM " & T_HIS005 & _
'             " WHERE " & DBW(F_EMPID, EID, 2)
'
'
'
'    Set Get_COM006 = New Recordset
'    Get_COM006.RsOpen , strSQL
'
'    If Get_COM006.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

Public Function GetHIS005EmpNm(ByVal vEmpId As String) As String
'Get_COM006 �̶� �м��� �ٲ�
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT " & F_EMPNM & " as empnm FROM " & T_HIS005 & _
             " WHERE " & DBW(F_EMPID, vEmpId, 2)
                     
    Set Rs = New Recordset
    On Error GoTo ErrTrap
    
    Rs.Open strSQL, dbconn
    
    If Rs.EOF Then
        GetHIS005EmpNm = ""
    Else
        GetHIS005EmpNm = Rs.Fields("empnm").Value & ""
    End If
    Exit Function
    
ErrTrap:
    Set Rs = Nothing
    MsgBox Err.Description, vbExclamation
End Function

'Public Function Get_COM011() As Recordset
'    Dim strSQL As String
'
'    '�����̸��� ��������.
'    strSQL = "SELECT max(seq)  FROM " & T_COM011
'
'    Set Get_COM011 = New Recordset
'    Get_COM011.RsOpen , strSQL
'
'    If Get_COM011.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

Public Function GetCOM011MaxSeq() As Long
'Get_COM011 �̶� �м��� �ٲ�
'Com011 ���̺����� �ƽ� ������ ��������
    Dim Rs As Recordset
    Dim lngSeq As Long
    
    Dim strSQL As String
    
    strSQL = "select max(seq) as maxseq from " & T_COM011
    
    Set Rs = New Recordset
    
    On Error GoTo ErrTrap
    
    Rs.Open strSQL, dbconn
    
    If Rs.EOF Then
        lngSeq = 1
    Else
        lngSeq = Val(Rs.Fields("maxseq").Value & "") + 1
    End If
    
    GetCOM011MaxSeq = lngSeq
    
ErrTrap:
    Set Rs = Nothing
End Function

'Public Function Get1_COM011(ByVal DeptFg As String) As Recordset
'    Dim strSQL As String
'    Dim sQuery As String
'
'
'    sQuery = "and workfg='" & WorkFg & "'"
'
'    '�μ�üũ���ϰ� ������ ��������...
'    If DeptFg = "LIS" Then
'        strSQL = "SELECT * FROM " & T_COM011 & _
'                 " where limitday >= '" & Format(DbConn.GetSysDate, "yyyymmdd") & "' and lisfg='1' " ' & _
'    ElseIf DeptFg = "APS" Then
'        strSQL = "SELECT * FROM " & T_COM011 & _
'                 " where limitday >= '" & Format(DbConn.GetSysDate, "yyyymmdd") & "' and apsfg='1' " ' & _
'    Else
'        strSQL = "SELECT * FROM " & T_COM011 & _
'                 " where limitday >= '" & Format(DbConn.GetSysDate, "yyyymmdd") & "' and bbsfg='1' " ' & _
'    End If
'
'    Set Get1_COM011 = New Recordset
'    Get1_COM011.RsOpen , strSQL
'
'    If Get1_COM011.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

Public Function GetSQLCOM011ByDeptFg(ByVal DeptFg As String) As String
'Get1_COM011�̶� �м��� �ٲ�
    Dim strSQL As String
    
    If DeptFg = "LIS" Then
        strSQL = "SELECT * FROM " & T_COM011 & _
                 " where limitday >= '" & Format(GetSystemDate, "yyyymmdd") & "' and lisfg='1' "
    ElseIf DeptFg = "APS" Then
        strSQL = "SELECT * FROM " & T_COM011 & _
                 " where limitday >= '" & Format(GetSystemDate, "yyyymmdd") & "' and apsfg='1' "
    ElseIf DeptFg = "BBS" Then
        strSQL = "SELECT * FROM " & T_COM011 & _
                 " where limitday >= '" & Format(GetSystemDate, "yyyymmdd") & "' and bbsfg='1' "
    End If
    
    GetSQLCOM011ByDeptFg = strSQL
End Function

'Public Function Get2_COM011(ByVal lngSeq As String, ByVal strDt As String) As Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT * FROM " & T_COM011 & _
'             " where " & DBW("seq=", lngSeq) & " and " & DBW("inputday=", strDt)
'
'    Set Get2_COM011 = New Recordset
'    Get2_COM011.RsOpen , strSQL
'
'    If Get2_COM011.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

Public Function GetSQLCOM011BySeq(ByVal lngSeq As String, ByVal strDt As String) As String
'Get2_COM011 �̶�� �м��� �ٲ�
    Dim strSQL As String
    
    strSQL = "SELECT * FROM " & T_COM011 & _
             " where " & DBW("seq=", lngSeq) & " and " & DBW("inputday=", strDt)
    GetSQLCOM011BySeq = strSQL
End Function

Public Function GetSQLCOM011(ByVal lngSeq As String, ByVal strDt As String) As String
    GetSQLCOM011 = "SELECT * FROM " & T_COM011 & _
                   " where " & DBW("seq=", lngSeq) & " and " & DBW("inputday=", strDt)
End Function

Public Function GetSQLCOM011ByDeptFg2(ByVal DeptFg As String) As String
'Get3_COM011�� �̸��� �м��� �ٲ�.

    Dim strSQL As String
    Dim sQuery As String
    
    sQuery = "and workfg='" & WorkFg & "'"

    '�μ�üũ���ϰ� ������ ��������...
    If DeptFg = "LIS" Then

        If WorkFg = "0" Then
            strSQL = "SELECT * FROM " & T_COM011 & " where  lisfg='1' and workfg='" & WorkFg & "'"
        ElseIf WorkFg = "1" Then
            strSQL = "SELECT * FROM " & T_COM011 & " where workfg='" & WorkFg & "'"
        End If
    ElseIf DeptFg = "APS" Then
        strSQL = "SELECT * FROM " & T_COM011 & " where  apsfg='1'" & sQuery
        
    Else
        strSQL = "SELECT * FROM " & T_COM011 & " where  bbsfg='1'" & sQuery
    End If
    
    strSQL = strSQL & " AND limitday >= '" & Format(GetSystemDate, "yyyymmdd") & "'"
    

    strSQL = strSQL & " ORDER BY limitday DESC"

    GetSQLCOM011ByDeptFg2 = strSQL
End Function

'Public Function Get3_COM011(ByVal DeptFg As String) As Recordset
'    Dim strSQL As String
'    Dim sQuery As String
'
'    sQuery = "and workfg='" & WorkFg & "'"
'
'
'    '�μ�üũ���ϰ� ������ ��������...
'    If DeptFg = "LIS" Then
'
'        If WorkFg = "0" Then
'            strSQL = "SELECT * FROM " & T_COM011 & " where  lisfg='1' and workfg='" & WorkFg & "'"
'        ElseIf WorkFg = "1" Then
'            strSQL = "SELECT * FROM " & T_COM011 & " where workfg='" & WorkFg & "'"
'        End If
'    ElseIf DeptFg = "APS" Then
'        strSQL = "SELECT * FROM " & T_COM011 & " where  apsfg='1'" & sQuery
'
'    Else
'        strSQL = "SELECT * FROM " & T_COM011 & " where  bbsfg='1'" & sQuery
'    End If
'
'    strSQL = strSQL & " AND limitday >= '" & Format(DbConn.GetSysDate, "yyyymmdd") & "'"
'
'
'    strSQL = strSQL & " ORDER BY limitday DESC"
'
'    Set Get3_COM011 = New Recordset
'    Get3_COM011.RsOpen , strSQL
'
'    If Get3_COM011.DBerror = True Then
'        MsgBox DbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'End Function

Public Function Insert_COM011(ByVal lngSeq As Long, ByVal IDt As String, ByVal LDt As String, ByVal txtTitle As String, _
                ByVal TxtNote As String, ByVal TxtUser As String, ByVal Lfg As String, ByVal Afg As String, ByVal Bfg As String) As Boolean

    Dim strSQL As String

    '��� ��Ű��
    strSQL = " INSERT INTO " & T_COM011 & "(seq,inputday,limitday,title,note,users,lisfg,apsfg,bbsfg,workfg) " & _
             " VALUES(" & DBV("seq", lngSeq) & "," & DBV("inputday", IDt) & " , " & DBV("limitday", LDt) & " , " & _
             DBV("title", txtTitle) & " , " & DBV("note", TxtNote) & " , " & _
             DBV("users", TxtUser) & " , " & DBV("lisfg", Lfg) & " , " & DBV("apsfg", Afg) & " , " & DBV("bbsfg", Bfg) & ", " & _
             DBV("workfg", mvarWorkFg) & " ) "

On Error GoTo Insert_COM011_Error
    dbconn.BeginTrans
    dbconn.Execute strSQL
    dbconn.CommitTrans
    Insert_COM011 = True
    Exit Function

Insert_COM011_Error:
    dbconn.RollbackTrans
    Insert_COM011 = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Del_COM011(ByVal lngSeq As Long, ByVal IDt As String) As Boolean
    Dim strSQL As String

    '��� ��Ű��
    strSQL = " delete from " & T_COM011 & _
             " where " & DBW("seq=", lngSeq) & _
             " and " & DBW("inputday=", IDt)

On Error GoTo Del_COM011_Error
    dbconn.BeginTrans
    dbconn.Execute strSQL
    dbconn.CommitTrans
    Del_COM011 = True
    Exit Function

Del_COM011_Error:
    dbconn.RollbackTrans
    Del_COM011 = False
    MsgBox Err.Description, "�������"
End Function