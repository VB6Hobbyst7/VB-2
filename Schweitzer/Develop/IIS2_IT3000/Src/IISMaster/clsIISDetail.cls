VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISDetail.cls
'   �ۼ���  : �̻��
'   ��  ��  : ���׸� ���� Ŭ����
'   �ۼ���  : 2004-02-20
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mParentCd   As String       '���ڵ�
Private mParentCdNm As String       '���ڵ��
Private mChildCd    As String       '���ڵ�
Private mChildCdNm  As String       '���ڵ��

Public Property Get ParentCd() As String
    ParentCd = mParentCd
End Property

Public Property Let ParentCd(ByVal vData As String)
    mParentCd = vData
End Property

Public Property Get ParentCdNm() As String
    ParentCdNm = mParentCdNm
End Property

Public Property Let ParentCdNm(ByVal vData As String)
    mParentCdNm = vData
End Property

Public Property Get ChildCd() As String
    ChildCd = mChildCd
End Property

Public Property Let ChildCd(ByVal vData As String)
    mChildCd = vData
End Property

Public Property Get ChildCdNm() As String
    ChildCdNm = mChildCdNm
End Property

Public Property Let ChildCdNm(ByVal vData As String)
    mChildCdNm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���׸� �߰�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddDetail() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.InsertDetailCd(mParentCd, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddDetail = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "AddDetail", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���׸� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelDetail() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteDetailCd(mParentCd, mChildCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelDetail = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "DelDetail", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���׸��� ���ڵ忡 ��ϵ� ���ڵ带 ��λ���
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelDetailAll() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteDetailCdall(mParentCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelDetailAll = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISDetail", "DelDetailAll", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���ڵ� ����Ʈ ��ȸ
'   ��ȯ : ���ڵ� ����Ʈ Recordset
'-----------------------------------------------------------------------------'
Public Function GetParentList() As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailParentCd, , adCmdText)
    Set GetParentList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetParentList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���ڵ� ����Ʈ ��ȸ
'   �μ� :
'       1.pParentCd : ���ڵ�
'   ��ȯ : ���ڵ� ����Ʈ Recordset
'-----------------------------------------------------------------------------'
Public Function GetChildList(ByVal pParentCd As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailChildCd(pParentCd), , adCmdText)
    Set GetChildList = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ带 ��ȸ�Ͽ� �����ϸ� �ڵ�� ��ȯ
'   �μ� :
'       1.pParentCd : ���ڵ�
'   ��ȯ :
'       1.�ش� ���ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetParentNm(ByVal pParentCd As String) As String
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDetailParentNm(pParentCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine

    GetParentNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetParentNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ带 ��ȸ�Ͽ� �����ϸ� �ڵ�� ��ȯ
'   �μ� :
'       1.pChildCd : ���ڵ�
'   ��ȯ :
'       1.�ش� ���ڵ尡 �����ϸ� �ڵ�� ��ȯ
'       2.�������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetChildNm(ByVal pChildCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectDatailChildNm(pChildCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine

    GetChildNm = Rs.Fields("TESTNM").Value
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISDetail", "GetChildNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : TIIS101���� ���ڵ� ����Ʈ�� ��ȸ
'   ��ȯ : ��ȸ ����
'-----------------------------------------------------------------------------'
Public Function GetParentListSql() As String
    GetParentListSql = mSQL.SelectDetailParentCdX
End Function

'-----------------------------------------------------------------------------'
'   ��� : TIIS101���� ���ڵ� ����Ʈ�� ��ȸ
'   ��ȯ : ��ȸ ����
'-----------------------------------------------------------------------------'
Public Function GetChildListSql() As String
    GetChildListSql = mSQL.SelectDetailChildCdX
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub