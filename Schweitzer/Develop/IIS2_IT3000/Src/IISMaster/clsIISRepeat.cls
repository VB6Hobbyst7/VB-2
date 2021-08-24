VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISRepeat.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : �ٺ�ó�� ���� Ŭ����
'   �ۼ���  : 2004-02-23
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - �ٺ�ó�� �ڵ�
'   - �ٺ�ó�� �ڵ��
'   - ��ü�ڵ�
'   - ��ü��

'## �޼ҵ�
'   - �ٺ�ó�� �߰�
'   - �ٺ�ó�� ����
'   - �ٺ�ó�� ����
'   - ��� �ٺ�ó�� ��ȸ
'   - �ٺ�ó�� �ڵ忡 ���� �̸� ��ȸ(�˻��׸��� �׷��ڵ�, �󼼸��ڵ� �̾�� �Ѵ�.)
'   - �ش� ��ü�� ���� �̸� ��ȸ(��ü�ڵ�� �ٺ�ó�� �ڵ忡 ��ϵ� ��ü�̾�� �Ѵ�.)
'   - TIIS101���� �׷��ڵ�, �󼼸��ڵ� ��ȸ
'   - �ش� �ٺ�ó�� �ڵ忡 ��ϵ� ����ü ��ȸ

Option Explicit

Private mSQL        As clsIISMasterSql

Private mRepeatCd   As String       '�ٺ�ó�� �ڵ�
Private mRepeatNm   As String       '�ٺ�ó�� �ڵ��
Private mSpcCd      As String       '��ü�ڵ�
Private mSpcNm      As String       '��ü�ڵ��

Public Property Get RepeatCd() As String
    RepeatCd = mRepeatCd
End Property

Public Property Let RepeatCd(ByVal vData As String)
    mRepeatCd = vData
End Property

Public Property Get RepeatNm() As String
    RepeatNm = mRepeatNm
End Property

Public Property Let RepeatNm(ByVal vData As String)
    mRepeatNm = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �ٺ�ó�� �ڵ� �߰�
'   �μ� : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertRepeatCd(mRepeatCd, mSpcCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "AddRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ٺ�ó�� �ڵ� ����
'   �μ� : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdateRepeatCd(mRepeatCd, mSpcCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "ModifyRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ٺ�ó�� �ڵ� ����
'   �μ� : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelRepeatCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteRepeatCd(mRepeatCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelRepeatCd = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISRepeat", "ModifyRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ٺ�ó�� ����Ʈ ��ȸ
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetRepeatCd() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectRepeatCd, , adCmdText)
    Set GetRepeatCd = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetRepeatCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ٺ�ó�� �ڵ�� ��ȸ
'   �μ� :
'       1.pRepeatCd : �ٺ�ó�� �ڵ�
'   ��ȯ :
'       1.�ش� �ٺ�ó�� �ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetRepeatNm(ByVal pRepeatCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectRepeatNm(pRepeatCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetRepeatNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetRepeatNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��ü �ڵ�� ��ȸ (�ٺ�ó�� �ڵ忡 ��ϵ� ��ü�ڵ��߿��� ��ȸ)
'   �μ� :
'       1.pRepeatCd : �ٺ�ó�� �ڵ�
'       2.pSpcCd    : ��ü�ڵ�
'   ��ȯ :
'       1.�ش� ��ü�ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetSpcNm(ByVal pRepeatCd As String, ByVal pSpcCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectSpcNmX(pRepeatCd, pSpcCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetSpcNm = Rs.Fields("SPCNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISRepeat", "GetSpcNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : TIIS101���� �׷��ڵ�, �󼼸��ڵ常 ��ȸ
'-----------------------------------------------------------------------------'
Public Function GetRepeatCdBySql() As String
    GetRepeatCdBySql = mSQL.SelectRepeatCdx
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��� ��ü�ڵ� ��ȸ
'   �μ� :
'       1.pRepeatCd : �ٺ�ó�� �ڵ�
'-----------------------------------------------------------------------------'
Public Function GetSpcCdBySql(ByVal pRepeatCd As String) As String
    GetSpcCdBySql = mSQL.SelectSpcCdX(pRepeatCd)
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub