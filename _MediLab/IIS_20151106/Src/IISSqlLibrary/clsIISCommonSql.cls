VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCommonSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCommonSql.cls (���ֿ�������)
'   �ۼ���  : �̻��
'   ��  ��  : EXE���� ����ϴ� ����
'   �ۼ���  : 2004-10-05
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : ����� ���� ��ȸ����
'   �μ� :
'       1.pLoginId : �α��� ���̵�
'-----------------------------------------------------------------------------'
Public Function SelectEmpInfo(ByVal pLoginId As String) As String
    Dim SQL As String
    
    SQL = " SELECT a.empid AS EMPID, a.empnm AS EMPNM, b.loginpass AS LOGINPASS" & _
          " FROM " & THIS006 & " a, " & THIS005 & " b" & _
          " WHERE b.loginid='" & pLoginId & "'" & _
          "     AND a.empid=b.empid"
          
    SelectEmpInfo = SQL
End Function