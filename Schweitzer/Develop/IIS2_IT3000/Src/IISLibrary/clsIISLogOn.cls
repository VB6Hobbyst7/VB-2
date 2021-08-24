VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISLogOn.cls
'   �ۼ���  : �̻��
'   ��  ��  : �α��� Ŭ����
'   �ۼ���  : 2003-12-07
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - ���� ���̵�
'   - ���� �̸�
'   - �α��� ���̵�
'   - �α��� ���

'## �޼ҵ�
'   - ����� ���� ��ȸ

Option Explicit

Private mSQL        As clsIISCommonSql

Private mEmpId      As String       '�������̵�
Private mEmpNm      As String       '�����̸�
Private mLoginId    As String       '�α��� ���̵�
Private mLoginPass  As String       '�α��� ��й�ȣ

Public Property Get EmpId() As String
    EmpId = mEmpId
End Property

Public Property Get EmpNm() As String
    EmpNm = mEmpNm
End Property

Public Property Get LoginId() As String
    LoginId = mLoginId
End Property

Public Property Get LoginPass() As String
    LoginPass = mLoginPass
End Property

'-----------------------------------------------------------------------------'
'   ��� : ����� ������ �ε�
'   �μ� :
'       1. pLoginId : �α��� ���̵�
'-----------------------------------------------------------------------------'
Public Function GetEmpInfo(ByVal pLoginId As String) As Boolean
    Dim Rs      As ADODB.Recordset
    Dim SQL     As String
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEmpInfo(pLoginId), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetEmpInfo = False
    Else
        mEmpId = Rs.Fields("EMPID").Value
        mEmpNm = Rs.Fields("EMPNM").Value
        mLoginId = pLoginId
        mLoginPass = Rs.Fields("LOGINPASS").Value
        GetEmpInfo = True
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsIISLogOn", "GetEmpInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISCommonSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub