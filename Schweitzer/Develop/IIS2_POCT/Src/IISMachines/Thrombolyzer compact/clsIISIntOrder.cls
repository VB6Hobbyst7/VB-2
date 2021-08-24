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
'   ��  ��  : Thrombolyzer compact ��� �������� Ŭ����
'   �ۼ���  : 2004-10-16
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Const RS As String = ""    'Record Separator
Private Const GS As String = ""    'Group Separator

Private mBarNo  As String   '���ڵ��ȣ
Private mPos    As String   'Tube Position

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '�۽��� �˻��׸�
    Dim strOutput   As String           '�۽��� ������
    Dim lngCount    As Long             '�۽��� �˻��׸� ����
    Dim strTemp     As String
    Dim i           As Long
    
    '## ������ �˻��
    '   PT(Time):1T, PT(%):1%, PT(INR):1I
    '   aPTT    :2T
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                strItems = strItems & "|" & Format$(strIntBase, "@@")
                lngCount = lngCount + 1
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                strItems = strItems & "|" & Format$(strIntBase, "@@")
                lngCount = lngCount + 1
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If
    
    '## �ִ� 9���׸���� ���۰���, 9�� �̸��϶��� "99"�� ����
    For i = 1 To 9 - lngCount
        strItems = strItems & "|99"
    Next i
    
    strOutput = "0|" & Format$(mPos, "@@") & "|" & Format$(mBarNo, "!" & String$(30, "@")) & strItems
    strOutput = RS & strOutput & vbCrLf & GS
    strOutput = STX & vbCrLf & strOutput & GetChkSum(strOutput) & vbCrLf
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ �������� ���ڿ� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strOutput   As String   '�۽��� ������
    
    strOutput = "1|" & Format$(mPos, "@@") & "|" & Format$(mBarNo, "!" & String$(30, "@")) & _
                "|99|99|99|99|99|99|99|99|99"
    strOutput = RS & strOutput & vbCrLf & GS
    strOutput = STX & vbCrLf & strOutput & GetChkSum(strOutput) & vbCrLf
    GetNoOrder = strOutput
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
    mPos = ""
End Sub