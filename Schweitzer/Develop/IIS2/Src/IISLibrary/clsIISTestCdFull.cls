VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCdFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTestCdFull.cls
'   �ۼ���  : �̻��
'   ��  ��  : �˻��ڵ�(��ü) Ŭ����
'   �ۼ���  : 2004-01-19
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mTestCd     As String       '�˻��ڵ�
Private mApplyDt    As String       '������
Private mExpireDt   As String       '�����
Private mTestNm5    As String       '�˻��(5�ڸ�)
Private mTestNm10   As String       '�˻��(10�ڸ�)
Private mTestNm     As String       '�˻��(��ü)
Private mWorkarea   As String       'Workarea
Private mRstType    As String       '�������(Null: �Ϲ�, F: Free, N: Numeric, A: Alpha)
Private mRstDiv     As String       'Alternative, Require ����(A: Alternative, R: Require)
Private mPanelFg    As String       'ó�汸��(Null: �����׸�, G: �׷��׸�, D: ���׸�)
Private mDetailFg   As String       'Detail �׸񿩺�(Null: ����, *:���׸� ���ڵ�)
Private mRptSeq     As Long         '��¼���

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get Applydt() As String
    Applydt = mApplyDt
End Property

Public Property Let Applydt(ByVal vData As String)
    mApplyDt = vData
End Property

Public Property Get ExpireDt() As String
    ExpireDt = mExpireDt
End Property

Public Property Let ExpireDt(ByVal vData As String)
    mExpireDt = vData
End Property

Public Property Get TestNm5() As String
    TestNm5 = mTestNm5
End Property

Public Property Let TestNm5(ByVal vData As String)
    mTestNm5 = vData
End Property

Public Property Get TestNm10() As String
    TestNm10 = mTestNm10
End Property

Public Property Let TestNm10(ByVal vData As String)
    mTestNm10 = vData
End Property

Public Property Get TestNm() As String
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vData As String)
    mTestNm = vData
End Property

Public Property Get Workarea() As String
    Workarea = mWorkarea
End Property

Public Property Let Workarea(ByVal vData As String)
    mWorkarea = vData
End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
End Property

Public Property Get RstDiv() As String
    RstDiv = mRstDiv
End Property

Public Property Let RstDiv(ByVal vData As String)
    mRstDiv = vData
End Property

Public Property Get PanelFg() As String
    PanelFg = mPanelFg
End Property

Public Property Let PanelFg(ByVal vData As String)
    mPanelFg = vData
End Property

Public Property Get DetailFg() As String
    DetailFg = mDetailFg
End Property

Public Property Let DetailFg(ByVal vData As String)
    mDetailFg = vData
End Property

Public Property Get RptSeq() As Long
    RptSeq = mRptSeq
End Property

Public Property Let RptSeq(ByVal vData As Long)
    mRptSeq = vData
End Property

'------------------------------------------------------------------------------'
'   ��� : �˻��ڵ� �߰�
'   ��ȯ : True(����), False(����)
'------------------------------------------------------------------------------'
Public Function AddTestCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertTestCd(mTestCd, mApplyDt, mExpireDt, mTestNm5, mTestNm10, mTestNm, _
                mWorkarea, mRstType, mRstDiv, mPanelFg, mDetailFg, mRptSeq)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTestCdFull", "AddTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'------------------------------------------------------------------------------'
'   ��� : �˻��ڵ� ����
'   ��ȯ : True(����), False(����)
'------------------------------------------------------------------------------'
Public Function ModifyTestCd() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateTestCd(mTestCd, mApplyDt, mExpireDt, mTestNm5, mTestNm10, mTestNm, _
                mWorkarea, mRstType, mRstDiv, mPanelFg, mDetailFg, mRptSeq)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTestCdFull", "ModifyTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'------------------------------------------------------------------------------'
'   ��� : �˻��ڵ� ����
'   ��ȯ : True(����), False(����)
'------------------------------------------------------------------------------'
Public Function DelTestCd() As Boolean
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteTestCd(mTestCd, mApplyDt), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelTestCd = True
    Exit Function
    
Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISTestCdFull", "DelTestCd", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub