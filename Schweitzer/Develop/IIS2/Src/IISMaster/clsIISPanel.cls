VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISPanel.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : �׷�ó�� ���� Ŭ����
'   �ۼ���  : 2004-02-21
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - �׷��ڵ�
'   - �׷��ڵ��
'   - �׷��ڵ� SEQ
'   - ���ڵ�
'   - ���ڵ��

'## �޼ҵ�
'   - �׷��ڵ� ����
'   - �׷��ڵ� ����
'   - �׷��ڵ� ����
'   - �׷��ڵ� �����׸� ��� ����
'   - �׷��ڵ� ����Ʈ ��ȸ
'   - ���׸� ��ȸ
'   - �׷��ڵ��� �̸���ȸ
'   - ���ڵ��� �̸���ȸ
'   - �˻��ڵ� �����Ϳ��� �׷��ڵ常 ��ȸ
'   - �˻��ڵ� �����Ϳ��� �����׸�, �󼼾����� ��ȸ

Option Explicit

Private mSQL        As clsIISMasterSql

Private mParentCd   As String       '�׷��ڵ�
Private mParentNm   As String       '�׷��ڵ��
Private mSeq        As String       '�׷��ڵ� SEQ
Private mChildCd    As String       '���ڵ�
Private mChildNm    As String       '���ڵ��

Public Property Get ParentCd() As String
    ParentCd = mParentCd
End Property

Public Property Let ParentCd(ByVal vData As String)
    mParentCd = vData
End Property

Public Property Get ParentNm() As String
    ParentNm = mParentNm
End Property

Public Property Let ParentNm(ByVal vData As String)
    mParentNm = vData
End Property

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

Public Property Get ChildCd() As String
    ChildCd = mChildCd
End Property

Public Property Let ChildCd(ByVal vData As String)
    mChildCd = vData
End Property

Public Property Get ChildNm() As String
    ChildNm = mChildNm
End Property

Public Property Let ChildNm(ByVal vData As String)
    ChildNm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �׷��ڵ� �߰�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertPanelCd(mParentCd, mSeq, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "AddPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �׷��ڵ� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.UpdatePanelCd(mParentCd, mSeq, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "ModifyPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �׷��ڵ� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelPanel() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeletePanelCd(mParentCd, mSeq), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelPanel = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "DelPanel", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �׷��ڵ� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelPanelAll() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeletePanelCdAll(mParentCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelPanelAll = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISPanel", "DelPanelAll", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : Panel ��ǥ�ڵ� ����Ʈ ��ȸ
'   ��ȯ : Panel ��ǥ�ڵ� Recordset
'-----------------------------------------------------------------------------'
Public Function GetPanelList() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelList, , adCmdText)
    Set GetPanelList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : Panel ���ڵ� ����Ʈ ��ȸ
'   ��ȯ : Panel ���ڵ� Recordset
'-----------------------------------------------------------------------------'
Public Function GetPanelChildList(ByVal pParentCd As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelChildList(pParentCd), , adCmdText)
    Set GetPanelChildList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �׷��ڵ��� �ڵ�� ��ȸ
'   �μ� :
'       1.pParentCd : �׷��ڵ�
'   ��ȯ :
'       1.�ش� �׷��ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetPanelNm(ByVal pParentCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelNm(pParentCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetPanelNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetPanelNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��� �ڵ�� ��ȸ
'   �μ� :
'       1.pChildCd : ���ڵ�
'   ��ȯ :
'       1.�ش� ���ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetChildNm(ByVal pChildCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectPanelChildNm(pChildCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    GetChildNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISPanel", "GetChildNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : TIIS101���� �׷��ڵ� ����Ʈ�� ��ȸ
'   ��ȯ : ��ȸ ����
'-----------------------------------------------------------------------------'
Public Function GetPanelListBySql() As String
    GetPanelListBySql = mSQL.SelectPanelParentCd
End Function

'-----------------------------------------------------------------------------'
'   ��� : TIIS101���� �����׸�, �󼼸��ڵ� ����Ʈ�� ��ȸ
'   ��ȯ : ��ȸ ����
'-----------------------------------------------------------------------------'
Public Function GetChildListBySql() As String
    GetChildListBySql = mSQL.SelectPanelChildCd
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub