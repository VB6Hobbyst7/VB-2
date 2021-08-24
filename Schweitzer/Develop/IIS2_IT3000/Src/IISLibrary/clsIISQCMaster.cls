VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQCMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISQCMaster.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : QC ������ Ŭ����
'   �ۼ���  : 2004-05-21
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISQcSql

Private mCtrlCd     As String   'Control Code
Private mLevelCd    As String   'Level Code
Private mLotNo      As String   'Lot No
Private mTestCd     As String   '�˻��ڵ�
Private mMeanVal    As Single   '���
Private mSDVal      As Single   'ǥ������
Private mAvalVal    As Long     '��ȿ����
Private mRefCd      As String   'Alpha��� ����ġ
Private mUnit       As String   '�������
Private mCVVal      As Single   '����
Private mMinVal     As Single   '�ּҰ�
Private mMaxVal     As Single   '�ִ밪
Private mWMSet      As String   'Multi-Rule Setting

Public Property Get CtrlCd() As String
    CtrlCd = mCtrlCd
End Property

Public Property Let CtrlCd(ByVal vData As String)
    mCtrlCd = vData
End Property

Public Property Get LevelCd() As String
    LevelCd = mLevelCd
End Property

Public Property Let LevelCd(ByVal vData As String)
    mLevelCd = vData
End Property

Public Property Get LotNo() As String
    LotNo = mLotNo
End Property

Public Property Let LotNo(ByVal vData As String)
    mLotNo = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get MeanVal() As Single
    MeanVal = mMeanVal
End Property

Public Property Let MeanVal(ByVal vData As Single)
    mMeanVal = vData
End Property

Public Property Get SDVal() As Single
    SDVal = mSDVal
End Property

Public Property Let SDVal(ByVal vData As Single)
    mSDVal = vData
End Property

Public Property Get AvalVal() As Long
    AvalVal = mAvalVal
End Property

Public Property Let AvalVal(ByVal vData As Long)
    mAvalVal = vData
End Property

Public Property Get RefCd() As String
    RefCd = mRefCd
End Property

Public Property Let RefCd(ByVal vData As String)
    mRefCd = vData
End Property

Public Property Get Unit() As String
    Unit = mUnit
End Property

Public Property Let Unit(ByVal vData As String)
    mUnit = vData
End Property

Public Property Get CVVal() As Single
    CVVal = mCVVal
End Property

Public Property Let CVVal(ByVal vData As Single)
    mCVVal = vData
End Property

Public Property Get MinVal() As Single
    MinVal = mMinVal
End Property

Public Property Let MinVal(ByVal vData As Single)
    mMinVal = vData
End Property

Public Property Get MaxVal() As Single
    MaxVal = mMaxVal
End Property

Public Property Let MaxVal(ByVal vData As Single)
    mMaxVal = vData
End Property

Public Property Get WMSet() As String
    WMSet = mWMSet
End Property

Public Property Let WMSet(ByVal vData As String)
    mWMSet = vData
End Property

Private Sub Class_Initialize()
    Set mSQL = New clsIISQcSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub