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
'   ��  ��  : VIDAS ��� �������� Ŭ����
'   �ۼ���  : 2004-08-23
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '���ڵ��ȣ
Private mPtId       As String   'ȯ��Id
Private mRow        As Long     '���� ���ڵ��ȣ�� Row No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal vData As Long)
    mRow = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ - 1.1.0: �̻��(2005-06-30)
'   �μ� :
'       - pAccInfo : �������� ��ü
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByRef pAccInfo As clsIISAccInfo) As String
'    Dim objResult   As clsIISResult     '������� ��ü
'    Dim objQCResult As clsIISQCResult   'QC������� ��ü
'    Dim strIntBase  As String           '������ �˻��
'    Dim strOrder    As String           '�������� ���ڿ�
'
'    If pAccInfo.QcFg = "0" Then     '## �Ϲݰ�ü
'        For Each objResult In pAccInfo.Results
'            If objResult.SndChk = False Then
'                strIntBase = objResult.IntNm.IntBase
'                strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
'                objResult.SndChk = True
'                Exit For
'            End If
'        Next
'        Set objResult = Nothing
'    ElseIf pAccInfo.QcFg = "1" Then '## QC��ü
'        For Each objQCResult In pAccInfo.QCResults
'            If objResult.SndChk = False Then
'                strIntBase = objQCResult.IntNm.IntBase
'                strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
'                objQCResult.SndChk = True
'                Exit For
'            End If
'        Next
'        Set objQCResult = Nothing
'    End If
'
'    strOrder = Chr(30) & "mtmpr|pi" & pAccInfo.PtId & "||si|ci" & mBarNo & "|" & _
'               "rt" & strIntBase & "|" & Chr(29)
'    GetOrder = STX & strOrder & GetChkSum(strOrder) & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ - 1.1.0: �̻��(2005-06-30)
'   �μ� :
'       - pIntBase : ������ �˻��
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrderX(ByVal pIntBase As String) As String
    Dim strOrder As String      '�������� ���ڿ�
    
    strOrder = Chr(30) & "mtmpr|pi" & mPtId & "||si|ci" & mBarNo & "|" & "rt" & pIntBase & "|" & Chr(29)
    GetOrderX = STX & strOrder & GetChkSum(strOrder) & ETX
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
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mPtId = ""
    mRow = 0
End Sub