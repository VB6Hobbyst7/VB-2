VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISEqpMaster.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : �˻���� ����
'   �ۼ���  : 2004-03-02
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - ����ڵ�
'   - ����
'   - ��񱸺�
'   - �𵨸�(Serial-No)
'   - �ǹ��ڵ�
'   - �������
'   - ����ȸ���ڵ�
'   - �����µ�(��)
'   - �����µ�(��)
'   - �µ�����
'   - ���

'## �޼ҵ�
'   - �ǹ����� ��ȸ
'   - ����ȸ�� ���� ��ȸ
'   - ��񸮽�Ʈ ��ȸ
'   - �˻���� �Է�
'   - �˻���� ����
'   - �˻���� ����

Option Explicit

Private mSQL        As clsIISMasterSql

Private mEqpCd      As String       '����ڵ�
Private mEqpNm      As String       '����
Private mEqpDiv     As String       '��񱸺�(E:�Ϲ����, C:�µ��������)
Private mModelNm    As String       '�𵨸�, Serial-No
Private mPurchDt    As String       '��������
Private mLocationCd As String       '�ǹ��ڵ�
Private mInUseFg    As String       '�������(1:���, 0:������)
Private mVandCd     As String       '�ǹ��ڵ�
Private mTempLow    As Single       '�����µ�(��)
Private mTempHigh   As Single       '�����µ�(��)
Private mTempScale  As String       '�µ�����(C:����, F:ȭ��)
Private mRemark     As String       '���

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get EqpNm() As String
    EqpNm = mEqpNm
End Property

Public Property Let EqpNm(ByVal vData As String)
    mEqpNm = vData
End Property

Public Property Get EqpDiv() As String
    EqpDiv = mEqpDiv
End Property

Public Property Let EqpDiv(ByVal vData As String)
    mEqpDiv = vData
End Property

Public Property Get ModelNm() As String
    ModelNm = mModelNm
End Property

Public Property Let ModelNm(ByVal vData As String)
    mModelNm = vData
End Property

Public Property Get PurchDt() As String
    PurchDt = mPurchDt
End Property

Public Property Let PurchDt(ByVal vData As String)
    mPurchDt = vData
End Property

Public Property Get LocationCd() As String
    LocationCd = mLocationCd
End Property

Public Property Let LocationCd(ByVal vData As String)
    mLocationCd = vData
End Property

Public Property Get InUseFg() As String
    InUseFg = mInUseFg
End Property

Public Property Let InUseFg(ByVal vData As String)
    mInUseFg = vData
End Property

Public Property Get VandCd() As String
    VandCd = mVandCd
End Property

Public Property Let VandCd(ByVal vData As String)
    mVandCd = vData
End Property

Public Property Get TempLow() As Single
    TempLow = mTempLow
End Property

Public Property Let TempLow(ByVal vData As Single)
    mTempLow = vData
End Property

Public Property Get TempHigh() As Single
    TempHigh = mTempHigh
End Property

Public Property Let TempHigh(ByVal vData As Single)
    mTempHigh = vData
End Property

Public Property Get TempScale() As String
    TempScale = mTempScale
End Property

Public Property Let TempScale(ByVal vData As String)
    mTempScale = vData
End Property

Public Property Get Remark() As String
    Remark = mRemark
End Property

Public Property Let Remark(ByVal vData As String)
    mRemark = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �ǹ����� ��ȸ
'   ��ȯ : �ǹ����� Recordset (�ǹ��ڵ�, �ǹ���)
'-----------------------------------------------------------------------------'
Public Function GetLocations() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectLocations, , adCmdText)
    Set GetLocations = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetLocations", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ����ȸ�� ������ȸ
'   ��ȯ : ����ȸ�� ���� Recordset (����ȸ���ڵ�, ����ȸ���)
'-----------------------------------------------------------------------------'
Public Function GetVands() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectVands, , adCmdText)
    Set GetVands = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetVands", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� ��ȸ
'   ��ȯ : �˻���� Recordset
'-----------------------------------------------------------------------------'
Public Function GetEqps() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors:
    Set Rs = DbCon.Execute(mSQL.SelectEqps, , adCmdText)
    Set GetEqps = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISEqpMaster", "GetEqps", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �Է�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddEqp() As Boolean
    Dim SQL As String
    
    SQL = mSQL.InsertEqp(mEqpCd, mEqpNm, mEqpDiv, mModelNm, mPurchDt, mLocationCd, mInUseFg, _
                mVandCd, mTempLow, mTempHigh, mTempScale, mRemark)
                
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    AddEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "AddEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyEqp() As Boolean
    Dim SQL As String
    
    SQL = mSQL.UpdateEqp(mEqpCd, mEqpNm, mEqpDiv, mModelNm, mPurchDt, mLocationCd, mInUseFg, _
                mVandCd, mTempLow, mTempHigh, mTempScale, mRemark)
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute SQL, , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    ModifyEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "ModifyEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DelEqp() As Boolean
    Dim SQL As String
    
On Error GoTo Errors:
    DbCon.BeginTrans
    DbCon.Execute mSQL.DeleteEqp(mEqpCd), , adCmdText + adExecuteNoRecords
    DbCon.CommitTrans
    DelEqp = True
    Exit Function

Errors:
    DbCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISEqpMaster", "DelEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub