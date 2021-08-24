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
'   ��  ��  : Hitachi 7180 ��� �������� Ŭ����
'   �ۼ���  : 2008-07-09
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '���ڵ��ȣ
Private mSampleData As String   'Sample Data
Private mFu         As String   'Function Character
Private mDiskNo     As String   'Disk No
Private mPos        As String   'Position No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SampleData() As String
    SampleData = mSampleData
End Property

Public Property Let SampleData(ByVal vData As String)
    mSampleData = vData
End Property

Public Property Get Fu() As String
    Fu = mFu
End Property

Public Property Let Fu(ByVal vData As String)
    mFu = vData
End Property

Public Property Get DiskNo() As String
    DiskNo = mDiskNo
End Property

Public Property Let DiskNo(ByVal vData As String)
    mDiskNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   �μ� :
'       - pAccInfo : �������� ��ü
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim lngIntBase  As Long             '������ �˻��
    Dim strItems    As String           '������ �˻��׸�
    Dim blnISE      As Boolean          'Na, K, Cl �˻翩��

    '## �˻��׸� ���ڿ� ����
    strItems = String$(88, "0")
    If pAccInfo.QcFg = "0" Then
        For Each objResult In pAccInfo.Results
            lngIntBase = CLng(objResult.IntNm.IntBase)
            
            '## ����׸�: 93~100
            If lngIntBase >= 93 And lngIntBase <= 100 Then GoTo Skip1
            
            '## Na, K, Cl �˻翩�� Check
            If lngIntBase = 87 Or lngIntBase = 88 Or lngIntBase = 89 Then
                blnISE = True
            Else
                Mid$(strItems, lngIntBase, 1) = "1"
            End If
Skip1:
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then
        For Each objQCResult In pAccInfo.QCResults
            lngIntBase = CLng(objQCResult.IntNm.IntBase)
            
            '## ����׸�: 93~100
            If lngIntBase >= 93 And lngIntBase <= 100 Then GoTo Skip2
            
            '## Na, K, Cl �˻翩�� Check
            If lngIntBase = 87 Or lngIntBase = 88 Or lngIntBase = 89 Then
                blnISE = True
            Else
                Mid$(strItems, lngIntBase, 1) = "1"
            End If
Skip2:
        Next
        Set objQCResult = Nothing
    End If
    
    '## Na, K, Cl �˻翩�� Check
    If blnISE Then
        Mid$(strItems, 87, 1) = "1"
    End If

    GetOrder = STX & ";" & mFu & mSampleData & " 88" & strItems & "000000" & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ ������ ���ڿ� ��ȸ
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    GetNoOrder = STX & ";" & mFu & mSampleData & Space(15) & " 88" & String$(94, "0") & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mSampleData = ""
    mFu = ""
    mDiskNo = ""
    mPos = ""
End Sub