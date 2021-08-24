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
'   ��  ��  : Coulter LH750 ��� �������� Ŭ����
'   �ۼ���  : 2005-05-15
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '���ڵ��ȣ
Private mBlkCnt     As Long     'Block Count
Private mOrder      As String   '���� ������ �˻�����

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get BlkCnt() As Long
    BlkCnt = mBlkCnt
End Property

Public Property Let BlkCnt(ByVal vData As Long)
    mBlkCnt = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
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
    Dim strIntBase  As String           '��� �˻��׸�
    Dim strItems    As String           '������ �˻��׸�
    Dim blnCBC      As Boolean          'CBC �˻翩��
    Dim blnDiff     As Boolean          'Diff �˻翩��
    Dim blnRetic    As Boolean          'Retic �˻翩��
    Dim strOutput   As String           '������ Data

    '## �˻��׸� ���ڿ� ����
    If pAccInfo.QcFg = "0" Then
        For Each objResult In pAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 2, "|")
            Select Case strIntBase
                Case "C": blnCBC = True
                Case "D": blnDiff = True
                Case "R": blnRetic = True
            End Select
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = mGetP(objResult.IntNm.IntBase, 2, "|")
            Select Case strIntBase
                Case "C": blnCBC = True
                Case "D": blnDiff = True
                Case "R": blnRetic = True
            End Select
        Next
        Set objQCResult = Nothing
    End If
    
    If blnCBC = True And blnDiff = False And blnRetic = False Then      '## CBC
        strItems = "CBC"
    ElseIf blnCBC = True And blnDiff = True And blnRetic = False Then   '## CBC+Diff
        strItems = "DIFF"
    ElseIf blnCBC = True And blnDiff = False And blnRetic = True Then   '## CBC+Retic
        strItems = "CR"
    ElseIf blnCBC = True And blnDiff = True And blnRetic = True Then    '## CBC+Diff+Retic
        strItems = "CDR"
    ElseIf blnCBC = False And blnDiff = False And blnRetic = True Then  '## Retic
        strItems = "CR"
    End If
    
    strOutput = SOH & "02" & vbCrLf & _
                "WLAD" & vbCrLf & _
                "ID " & pAccInfo.PtId & vbCrLf & _
                "TS " & strItems & "," & mBarNo & "," & vbCrLf
    
    strOutput = strOutput & String$(256 - Len(strOutput), " ")
    strOutput = STX & "01" & strOutput & GetCalCRC16(strOutput) & ETX
    mOrder = strOutput
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : CRC16-CCITT ����Լ�
'   �μ� :
'       - pMsg : �����
'   ��ȯ : CRC Value
'-----------------------------------------------------------------------------'
Public Function GetCalCRC16(ByVal pMsg As String) As String
    Dim strMSB  As String
    Dim strLSB  As String
    Dim CRCMSB  As Long
    Dim CRCLSB  As Long
    Dim X       As Long
    Dim i       As Long
    
    CRCMSB = &HFF
    CRCLSB = &HFF
    
    For i = 1 To Len(pMsg)
        X = (Asc(Mid(pMsg, i, 1)) Xor CRCMSB) And &HFF
        X = X Xor Fix(X / 16)
        CRCMSB = CRCLSB Xor Fix(X / 8) Xor Fix(X * 16)
        CRCLSB = X Xor (X * 32)
    Next i
    
    CRCMSB = (CRCMSB Xor &HFF) And &HFF
    CRCLSB = (CRCLSB Xor &HFF) And &HFF
    
    If Len(Hex(CRCMSB)) = 1 Then
        strMSB = "0" & Hex(CRCMSB)
    Else
        strMSB = Hex(CRCMSB)
    End If
    
    If Len(Hex(CRCLSB)) = 1 Then
        strLSB = "0" & Hex(CRCLSB)
    Else
        strLSB = CStr(Hex(CRCLSB))
    End If
    
    GetCalCRC16 = strMSB & strLSB
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mBlkCnt = 0
    mOrder = ""
End Sub