VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntNm.cls
'   �ۼ���  : �̻��
'   ��  ��  : ��� �˻��׸� Ŭ����
'   �ۼ���  : 2004-04-29
'   ��  ��  :
'       1. 1.1.6: �̻��(2005-05-11)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mEqpCd      As String   '����ڵ�
Private mIntNm      As String   '�˻��
Private mIntBase    As String   '������ �˻��
Private mTestCd     As String   '�˻��ڵ�
Private mFrVal      As String   'BorderLine From Value
Private mToVal      As String   'BorderLine To Value

'## 1.1.6: �̻��(2005-05-11)
'   - ���������� S2Lab702 ���̺� ���ؿ� ���߱� ���� ������� �߰�
Private mEqpSeq     As Long     '������
Private mPrtOrd     As Long     '��¼���

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get IntNm() As String
    IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As String)
    mIntNm = vData
End Property

Public Property Get IntBase() As String
    IntBase = mIntBase
End Property

Public Property Let IntBase(ByVal vData As String)
    mIntBase = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get FrVal() As String
    FrVal = mFrVal
End Property

Public Property Let FrVal(ByVal vData As String)
    mFrVal = vData
End Property

Public Property Get ToVal() As String
    ToVal = mToVal
End Property

Public Property Let ToVal(ByVal vData As String)
    mToVal = vData
End Property

'## 1.1.6: �̻��(2005-05-11)
'   - ���������� S2Lab702 ���̺� ���ؿ� ���߱� ���� ������Ƽ �߰�
Public Property Get EqpSeq() As Long
    EqpSeq = mEqpSeq
End Property

Public Property Let EqpSeq(ByVal vData As Long)
    mEqpSeq = vData
End Property

Public Property Get PrtOrd() As Long
    PrtOrd = mPrtOrd
End Property

Public Property Let PrtOrd(ByVal vData As Long)
    mPrtOrd = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �˻�� �Է�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddIntNm() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.InsertIntNm(mEqpCd, mIntNm, mIntBase, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻�� �Է�(����������) - 1.1.6: �̻��(2005-05-11)
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddIntNmX() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.InsertIntNmX(mEqpCd, mEqpSeq, mIntBase, mIntNm, mPrtOrd, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻�� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyIntNm() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.UpdateIntNm(mEqpCd, mIntNm, mIntBase, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "ModifyIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻�� ����(����������) - 1.1.6: �̻��(2005-05-11)
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyIntNmX() As Boolean
    Dim SQL As String
    
On Error GoTo Errors
    DbCon.BeginTrans
    SQL = mSQL.UpdateIntNmX(mEqpCd, mEqpSeq, mIntBase, mIntNm, mPrtOrd, mFrVal, mToVal)
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "ModifyIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻�� ����
'          �ش� �˻���� ��ϵ� ��� �˻��ڵ� ������ �˻�� ����
'   �μ� :
'       - pEqpCd : ����ڵ�
'       - pIntNm : �˻��
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelIntNm(ByVal pEqpCd As String, ByVal pIntNm As String) As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    
    DbCon.Execute mSQL.DeleteIntTestCdAll(pEqpCd, pIntNm), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.DeleteIntNm(pEqpCd, pIntNm), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelIntNm = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelIntNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻�� ����(����������) - 1.1.6: �̻��(2005-05-11)
'          �ش� �˻���� ��ϵ� ��� �˻��ڵ� ������ �˻�� ����
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpSeq : ������ �˻�� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelIntNmX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    
    DbCon.Execute mSQL.DeleteIntTestCdAllX(pEqpCd, pEqpSeq), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.DeleteIntNmX(pEqpCd, pEqpSeq), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelIntNmX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelIntNmX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ����� ��� �˻�� ��ȸ
'   �μ� :
'       - pEqpCd : ����ڵ�
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetIntNms(ByVal pEqpCd As String) As ADODB.Recordset
    Dim Rs  As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntNms(pEqpCd), , adCmdText)
    Set GetIntNms = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetIntNms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �˻��ڵ� �Է� (������ ������ �Է�)
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddTestCd() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.InsertIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �˻��ڵ� �Է� (������ ������ �Է�, ����������)
'          - 1.1.6: �̻��(2005-05-11)
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddTestCdX() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.Execute mSQL.InsertIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTestCdX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "AddTestCdX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �˻��ڵ� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelTestCd() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCd(mEqpCd, mIntNm, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �˻��ڵ� ����(����������) - 1.1.6: �̻��(2005-05-11)
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelTestCdX() As Boolean
On Error GoTo Errors
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteIntTestCdX(mEqpCd, mEqpSeq, mTestCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTestCdX = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISIntNm", "DelTestCdX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� ��ϵ� ��� �˻��ڵ� ��ȸ
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pIntNm  : �˻��
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCds(ByVal pEqpCd As String, ByVal pIntNm As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntTestCds(pEqpCd, pIntNm), , adCmdText)
    Set GetTestCds = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCds", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� ��ϵ� ��� �˻��ڵ� ��ȸ(����������) - 1.1.6: �̻��(2005-05-11)
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpSeq : ������ �˻�� ����
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCdsX(ByVal pEqpCd As String, ByVal pEqpSeq As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectIntTestCdsX(pEqpCd, pEqpSeq), , adCmdText)
    Set GetTestCdsX = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCdsX", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ带 �������� �˻�� ��ȸ
'   �μ� :
'       - pEqpCd  : ����ڵ�
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCdIntNms(ByVal pEqpCd As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectTestCdIntNms(pEqpCd), , adCmdText)
    Set GetTestCdIntNms = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCdIntNms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ����ڵ��� ���� ��ȸ
'   �μ� :
'       - pEqpCd : ����ڵ�
'   ��ȯ : �����ϴ� ����̸� ���� ��ȯ, �������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetEqpNm(ByVal pEqpCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEqpCd(pEqpCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetEqpNm = Rs.Fields("EQPNM").Value
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetEqpNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �Ϲ������ ������� ����ڵ�, ���� ��ȸ
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetUsingEqp() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectUsingEqp, , adCmdText)
    Set GetUsingEqp = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetUsingEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� �����Ϳ��� �����ڵ�, �����׸� ��ȸ
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetTestCd() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectForIntTestCd, , adCmdText)
    Set GetTestCd = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش�˻��ڵ��� �˻�� ��ȸ(����, �����ڵ忡���� ��ȸ)
'   �μ� :
'       - pTestCd : �˻��ڵ�
'   ��ȯ : �����ϸ� �̸� �˻�� ��ȯ, �������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetTestNm(ByVal pTestCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectForIntTestNm(pTestCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetTestNm = Rs.Fields("TESTNM").Value
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISIntNm", "GetTestNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub