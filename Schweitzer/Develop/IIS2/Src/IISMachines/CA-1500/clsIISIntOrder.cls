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
'   ��  ��  : CA-1500 �������� Ŭ����
'   �ۼ���  : 2005-07-25
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '���ڵ��ȣ
Private mRackNo     As String           'Rack No
Private mPos        As String           'Tube Position
Private mNoOrder    As Boolean          '�������� ��������(True:������, False:����)
Private mOrder      As String           'Order Information Text Format
Private mSendCnt    As String           '���� ����ȸ��

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
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
    Dim strStatFg     As String           '�������� (R:Routin, E:Stat)
    Dim strTemp       As String

    '## �˻��׸� ���ڿ� ����
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space(6)

            If strIntBase <> strTemp Then
                strItems = strItems & strIntBase
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space(6)

            If strIntBase <> strTemp Then
                strItems = strItems & strIntBase
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If

    strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")

    mOrder = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mRackNo & mPos & mBarNo & _
             "B" & Space(15) & strItems & ETX
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ ���� ���ڿ� ��ȸ
'   ��ȯ : ���������� ������ ���� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    mOrder = STX & "S2210101U" & Space(6) & Space(4) & mRackNo & mPos & mBarNo & "B" & _
             Space(15) & ETX
    
    GetNoOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRackNo = ""
    mPos = ""
    mNoOrder = True
    mOrder = ""
    mSendCnt = 0
End Sub