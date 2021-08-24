VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntOrder.cls
'   �ۼ���  : ������
'   ��  ��  : MultiReader ��� �������� Ŭ����
'   �ۼ���  : 2010-05-31
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '�������� Ŭ����
Private mBarNo      As String           '���ڵ��ȣ
Private mNoOrder    As Boolean          '�������� �������� (True:������, False:����)
Private mItems()    As String           '������ �˻��׸��
Private mCount      As Long             '������ �� �˻��׸� ����
Private mSendCnt    As Long             '������ �˻��׸� ����
Private mPtId       As String           '������ ȯ��ID
Private mSex        As String           '������ ����
Private mStatFg     As String           '������ ��������

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Let Count(ByVal vData As Long)
    mCount = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    SendCnt = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get StatFg() As String
    StatFg = mStatFg
End Property

Public Property Let StatFg(ByVal vData As String)
    mStatFg = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ��ȸ
'-----------------------------------------------------------------------------'
Public Sub GetItems()
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strTemp     As String

    ReDim mItems(mAccInfo.SendCnt)
    
    If mAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In mAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
            If strIntBase <> strTemp Then
                mItems(mCount) = "^^^" & strIntBase
                strTemp = strIntBase
                mCount = mCount + 1
            End If
        Next
        Set objResult = Nothing
        
        mPtId = mAccInfo.PtId
        mPtId = mAccInfo.Sex
        mStatFg = IIf(mAccInfo.StatFg = "1", "S", "")
    ElseIf mAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = mGetP(objQCResult.IntNm.IntBase, 1, "|")
            If strIntBase <> strTemp Then
                mItems(mCount) = "^^^" & strIntBase
                strTemp = strIntBase
                mCount = mCount + 1
            End If
        Next
        Set objQCResult = Nothing
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �˻��׸� ��ȸ
'   ��ȯ : �˻��׸�
'-----------------------------------------------------------------------------'
Public Function GetItem() As String
    Dim strIntBase  As String
    
    GetItem = mItems(mSendCnt)
    mSendCnt = mSendCnt + 1
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    Set mAccInfo = Nothing
    mBarNo = ""
    mNoOrder = False
    mCount = 0
    mSendCnt = 0
    mPtId = ""
    mSex = ""
    mStatFg = ""
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub
