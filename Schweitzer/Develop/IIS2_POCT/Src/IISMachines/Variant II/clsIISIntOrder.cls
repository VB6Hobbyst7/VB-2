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
'   ��  ��  : Variant II ��� �������� Ŭ����
'   �ۼ���  : 2004-10-13
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '�������� Ŭ����
Private mBarNo      As String           '���ڵ��ȣ
Private mOrder      As String           '���� ������ ��������
Private mNoOrder    As Boolean          '�������� �������� (True:������, False:����)
Private mIsSending  As Boolean          '���� ������ ���� ���ڿ��� ��������

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

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

Public Property Get IsSending() As Boolean
    IsSending = mIsSending
End Property

Public Property Let IsSending(ByVal vData As Boolean)
    mIsSending = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim objIntNm    As clsIISIntNm      '��� �˻��׸� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '������ �˻�� ���ڿ�
    Dim strTemp     As String

    '## �������ڿ� ����: ^^^1\^^^2
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            Set objIntNm = objResult.IntNm
            strIntBase = objIntNm.IntBase
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                    strTemp = strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                    strTemp = strIntBase
                End If
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            strIntBase = objIntNm.IntBase
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                    strTemp = strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                    strTemp = strIntBase
                End If
            End If
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing

    GetOrder = strItems
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
    mOrder = ""
    mNoOrder = False
    mIsSending = False
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub