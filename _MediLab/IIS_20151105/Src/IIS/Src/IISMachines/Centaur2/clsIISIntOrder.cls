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
'   ��  ��  : Centaur �������� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo   '�������� Ŭ����

Private mBarNo      As String   '���ڵ��ȣ
Private mNoOrder    As String   '�������� �������� (True:������, False:����)
Private mOrder      As String   '���� ������ ��������

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

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '�۽��� ������ �˻�� ���ڿ�
    Dim strTemp     As String

    '## 1.0.1: ���� �˻��׸� ���ؼ��� �ѹ��� ������ �����ϵ��� ����
    If mAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strTemp <> strIntBase Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                End If
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strTemp <> strIntBase Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                End If
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If

    GetOrder = strItems
End Function

'-----------------------------------------------------------------------------'
'   ��� : ȯ��ID ��ȸ
'   ��ȯ :
'       - ȯ��ID
'       - -1: �������� ��ü�� Null�϶�
'-----------------------------------------------------------------------------'
Public Function GetPtId() As String
    If Not (mAccInfo Is Nothing) Then
        GetPtId = mAccInfo.PtId
    Else
        GetPtId = "-1"
    End If
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
    mBarNo = ""
    mNoOrder = False
    mOrder = ""
    
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub