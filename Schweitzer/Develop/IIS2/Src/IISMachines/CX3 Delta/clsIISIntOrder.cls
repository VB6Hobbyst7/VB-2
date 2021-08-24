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
'   ��  ��  : CX3 Delta ��� �������� Ŭ����
'   �ۼ���  : 2004-10-14
'   ��  ��  :
'       1. 1.0.1: �̻��(2004-12-14)
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '�������� Ŭ����
Private mDeviceId   As String           'Device ID
Private mBarNo      As String           '���ڵ��ȣ
Private mSendCnt    As String           '������ ��ü����

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

Public Property Get DeviceId() As String
    DeviceId = mDeviceId
End Property

Public Property Let DeviceId(ByVal vData As String)
    mDeviceId = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '�۽��� �˻��׸� ���ڿ�
    Dim strTestTp   As String           '�۽��� Test Type(CO:Control, RO:Rountine, SC:STAT Control, ST:STAT)
    Dim strSpcTp    As String           '�۽��� Sample Type(SE:Serum, SF:Spinal Fluid, UR:Urine, TU:Timed Urine, PL:Plasma)
    Dim strPtId     As String           '�۽��� ȯ��ID
    Dim strAge      As String           '�۽��� Age
    Dim strSex      As String           '�۽��� Sex
    Dim strOutput   As String           '�۽��� ������
    Dim strSendCnt  As Long
    
    '## �˻��׸� ���ڿ� ����
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strItems = strItems & "," & Format$(strIntBase, "!@@@@") & ",0"
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strItems = strItems & "," & Format$(strIntBase, "!@@@@") & ",0"
        Next
        Set objQCResult = Nothing
    End If
    
    '## Test Type, Sample Type ��ȸ
    strTestTp = IIf(mAccInfo.StatFg = "1", "ST", "RO")
    Select Case Mid$(mAccInfo.SpcNm, 1, IISSPCLEN)
        Case IISSPCSERUM:   strSpcTp = "SE"
        Case IISSPCURINE:   strSpcTp = "UR"
        Case IISSPCPLASMA:  strSpcTp = "PL"
        Case IISSPCFLUID:   strSpcTp = "SF"
        Case IISSPCCSF:     strSpcTp = "SF"
        Case Else:          strSpcTp = "SE"
    End Select
    
    '## 1.0.1: �̻��(2004-12-14)
    '   - QC������ ��� ����, ������ ���� ����!
    '## ȯ��ID, ���� ��ȸ
    strPtId = Format$(mAccInfo.PtId, "!" & String$(12, "@"))
    If mAccInfo.QcFg = "0" Then
        strAge = Format$(Mid$(mAccInfo.Age, 1, Len(mAccInfo.Age) - 1), "000")
        strSex = mAccInfo.Sex
    ElseIf mAccInfo.QcFg = "1" Then
        strAge = "000"
        strSex = "M"
    End If
    
    strOutput = "[" & mDeviceId & ",701,01,00,00,0," & strTestTp & "," & strSpcTp & "," & mBarNo & _
                "," & Space(20) & "," & Space(25) & "," & Space(25) & "," & Space(18) & "," & _
                Space(15) & "," & Space(1) & "," & strPtId & "," & Space(18) & "," & Space(6) & _
                "," & Space(4) & "," & Space(20) & "," & strAge & ",5," & Space(6) & "," & strSex & _
                "," & Space(25) & "," & Space(7) & "," & Space(4) & "," & Space(4) & "," & Space(6) & _
                "," & Format$(mAccInfo.SendCnt, "000") & strItems & "]"
    strOutput = strOutput & GetChkSum(strOutput) & vbCrLf
    GetOrder = strOutput
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
        lngChkSum = lngChkSum + Asc(Mid$(pMsg, i, 1))
    Next i
    
    lngChkSum = lngChkSum Mod 256
    lngChkSum = 256 - lngChkSum
    
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
End Sub