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
'   ��  ��  : XT-1800i �������� Ŭ����
'   �ۼ���  : 2005-02-03
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '���ڵ��ȣ
Private mRackNo     As String           'Rack No
Private mPos        As String           'Tube Position
Private mNoOrder    As Boolean          '�������� ��������(True:������, False:����)
Private mOrder1     As String           'Analysis Informat Format 1
Private mOrder2     As String           'Analysis Informat Format 2
Private mSendCnt    As Long             '�������� ����

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

Public Property Get Order1() As String
    Order1 = mOrder1
End Property

Public Property Let Order1(ByVal vData As String)
    mOrder1 = vData
End Property

Public Property Get Order2() As String
    Order2 = mOrder2
End Property

Public Property Let Order2(ByVal vData As String)
    Order2 = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ�1 ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�1
'-----------------------------------------------------------------------------'
Public Function GetOrder1(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '������� Ŭ����
    Dim objQCResult   As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase    As String           '������ �˻��
    Dim strItems      As String           '�۽��� �˻��׸� ���ڿ�
    Dim strTemp       As String

    strItems = String$(33, "0")
    If pAccInfo.QcFg = "0" Then     '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = mGetP(objQCResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objQCResult = Nothing
    End If
    
    mOrder1 = STX & "S11" & Space(8) & "000" & mBarNo & "00" & mRackNo & mPos & "1" & _
              Space(145) & String$(18, "0") & strItems & String$(17, "0") & ETX
    GetOrder1 = mOrder1
    mNoOrder = False
    Debug.Print "Len(Order1) = " & Len(mOrder1)
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ�2 ��ȸ
'   ��ȯ : �������� ���ڿ�2
'-----------------------------------------------------------------------------'
Public Function GetOrder2() As String
    Dim strInfoCd   As String   '�۽��� Informat Status

    strInfoCd = IIf(NoOrder, "0", "1")
    
    mOrder2 = STX & "S2" & strInfoCd & Space(8) & "000" & mBarNo & "00" & mRackNo & mPos & "1" & _
              Space(116) & String(97, "0") & ETX
    GetOrder2 = mOrder2
    Debug.Print "Len(Order2) = " & Len(mOrder2)
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ ���� ���ڿ�1 ��ȸ
'   ��ȯ : ���������� ������ ���� ���ڿ�1
'-----------------------------------------------------------------------------'
Public Function GetNoOrder1() As String
    mOrder1 = STX & "S11" & Space(8) & "000" & mBarNo & "00" & mRackNo & mPos & "1" & _
              Space(145) & String$(51, "0") & String$(17, "0") & ETX
    GetNoOrder1 = mOrder1
    mNoOrder = True
    Debug.Print "Len(GetNoOrder1) = " & Len(mOrder1)
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRackNo = ""
    mPos = ""
    mNoOrder = True
    mOrder1 = ""
    mOrder2 = ""
    mSendCnt = 0
End Sub