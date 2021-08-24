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
'   ��  ��  : Dimension RXL ��� �������� Ŭ����
'   �ۼ���  : 2004-07-14
'   ��  ��  :
'       1. 1.0.1: �̻��(2004-12-16)
'          - Cardiac QC(MMB,MYO,CTNI)�� ��쿡�� QC1,2,3�� ����ؼ� IISQCNORMAL
'            ��ü���� �߰�!
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String   '���ڵ��ȣ

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : No Request Message ��ȸ
'   ��ȯ : No Request Message
'-----------------------------------------------------------------------------'
Public Function GetNoRequest() As String
    GetNoRequest = STX & "N6A" & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Wait Poll Message ��ȸ
'   ��ȯ : Wait Poll Message
'-----------------------------------------------------------------------------'
Public Function GetWaitPoll() As String
    GetWaitPoll = STX & "W73" & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Result Acceptance Message ��ȸ
'   ��ȯ : Result Acceptance Message
'-----------------------------------------------------------------------------'
Public Function GetResultAccept() As String
    GetResultAccept = STX & "MAE2" & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '������� Ŭ����
    Dim objQCResult   As clsIISQCResult   'QC������� Ŭ����
    Dim objIntNm      As clsIISIntNm      '��� �˻��׸� Ŭ����
    Dim strIntBase    As String           '������ �˻��׸�
    Dim strSampleType As String           '������ ��ü����
    Dim strOrder      As String           '�������� ���ڿ�
    
    '## OSMO �˻�� NA, GLU, BUN�˻����� ���� �˻��̹Ƿ� ���������� �������� ����
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            Set objIntNm = objResult.IntNm
            If objIntNm.IntBase = "OSMO" Then
                pAccInfo.SendCnt = pAccInfo.SendCnt - 1
            Else
                strIntBase = strIntBase & objIntNm.IntBase & ""
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            
            '## QC�� HA1C�� ��ü������ "W"�� �ؾ���
            Select Case objIntNm.IntBase
                Case "OSMO"
                    pAccInfo.SendCnt = pAccInfo.SendCnt - 1
                Case "HA1C"
                    pAccInfo.SpcNm = "B"
                    strIntBase = strIntBase & objIntNm.IntBase & ""
                Case Else
                    strIntBase = strIntBase & objIntNm.IntBase & ""
            End Select
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing

    '## 1.0.1: �̻��(2004-12-16)
    '   - IISQCNORMAL ��ü���� �߰�
    '## ��ü���� ��ȸ
    Select Case Mid$(pAccInfo.SpcNm, 1, IISSPCLEN)
        Case IISSPCSERUM:   strSampleType = "1"
        Case IISSPCPLASMA:  strSampleType = "2"
        Case IISSPCURINE:   strSampleType = "3"
        Case IISSPCCSF:     strSampleType = "4"
        Case IISSPCBLOOD:   strSampleType = "W"
        Case IISQCLOW:      strSampleType = "5"
        Case IISQCNORMAL:   strSampleType = "6"
        Case IISQCHIGH:     strSampleType = "7"
        Case Else:          strSampleType = "1"
    End Select

    strOrder = "D00A" & pAccInfo.PtId & "" & mBarNo & "" & strSampleType & "01**1" & _
               pAccInfo.SendCnt & "" & strIntBase

    GetOrder = STX & strOrder & GetChkSum(strOrder) & ETX
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
End Sub
