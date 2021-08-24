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
'   �ۼ���  : �̻��
'   ��  ��  : TBA-200FR �������� Ŭ����
'   �ۼ���  : 2005-02-04
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo   As String   '���ڵ��ȣ
Private mDiskNo  As String   'Disk No
Private mPos     As String   'Position
Private mOrder   As String   '�������� ���ڿ�
Private mSendCnt As Long     '���� ������ ȸ��

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
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

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '������� Ŭ����
    Dim objQCResult   As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase    As String           '������ �˻��
    Dim strItems      As String           '�۽��� �˻��׸� ���ڿ�

    '## �˻��׸� ���ڿ� ����
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(Format$(objResult.IntNm.IntBase, "#"), "@@@@") & "1"
            strItems = strItems & strIntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(Format$(objQCResult.IntNm.IntBase, "#"), "@@@@") & "1"
            strItems = strItems & strIntBase
        Next
        Set objQCResult = Nothing
    End If

    mOrder = STX & "O " & mBarNo & mDiskNo & mPos & "  1" & strItems & ETB & ETB & ETX
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ ���� ���ڿ� ��ȸ
'   ��ȯ : ���������� ������ ���� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    mOrder = STX & "O " & mBarNo & mDiskNo & mPos & ETB & ETB & ETX
    GetNoOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mDiskNo = ""
    mPos = ""
    mOrder = ""
    mSendCnt = 0
End Sub
