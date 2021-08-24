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
'   ���ϸ�  : clsIISBNII.cls
'   �ۼ���  : �̻��
'   ��  ��  : BN II �������� Ŭ����
'   �ۼ���  : 2005-01-27
'   ��  ��  :
'       1. 1.0.3: �̻��(2005-02-07)
'       2. 1.0.5: �̻��(2005-02-21)
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '���ڵ��ȣ
Private mOrder      As String           '�������� ���ڿ�

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
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
    Dim strBarNo      As String           '�۽��� ���ڵ��ȣ
    Dim strPtId       As String           '�۽��� ȯ��ID
    Dim strSex        As String           '�۽��� ����
    Dim strStatFg     As String           '�۽��� ��������
    Dim strTemp       As String
    
    '## 1.0.3: �̻��(2005-02-07)
    '   - �Ϲݰ�ü �϶��� ȯ���� ����, QC��ü�϶��� Space(1)�� �����ϵ��� ����
    '## 1.0.5: �̻��(2005-02-21)
    '   - �Ϲݰ�ü �϶��� ����������, QC��ü�϶��� �׻� �������� �����ϵ��� ����
    '## �˻��׸� ���ڿ� ����
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            If strTemp <> strIntBase Then
                strItems = strItems & strIntBase & "0"
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
        strSex = pAccInfo.Sex
        strStatFg = pAccInfo.StatFg
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            If strTemp <> strIntBase Then
                strItems = strItems & strIntBase & "0"
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
        strSex = Space(1)
        strStatFg = "0"
    End If
    
    strBarNo = Format$(mBarNo, "!" & String$(29, "@"))
    strPtId = Format$(pAccInfo.PtId, "!" & String$(30, "@"))
    mOrder = STX & "J" & strBarNo & Space(8) & strSex & strStatFg & Space(2) & "1" & _
             Space(40) & strPtId & Space(27) & strItems & Space(1) & vbCr & ETX & ACK
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mOrder = ""
End Sub