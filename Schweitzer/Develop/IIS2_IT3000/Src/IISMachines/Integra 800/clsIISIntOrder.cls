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
'   ��  ��  : Integra 800 ��� �������� Ŭ����
'   �ۼ���  : 2005-07-01
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo As String

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
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
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '������ �˻��׸�
    Dim strStat     As String           '��������(A:Routin, S:Stat)
    Dim strOutput   As String           '�۽��� ������
    
    If pAccInfo.QcFg = "0" Then     '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = "55 " & Format$(objResult.IntNm.IntBase, "000") & mEndChar
            strItems = strItems & strIntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = "55 " & Format$(objQCResult.IntNm.IntBase, "000") & mEndChar
            strItems = strItems & strIntBase
        Next
        Set objQCResult = Nothing
    End If
    
    strStat = IIf(pAccInfo.StatFg = "1", "S", "A")
    strOutput = SOH & mEndChar & _
                "09 " & Format$(mEqpNm, "!" & String$(16, "@")) & " 10" & mEndChar & _
                STX & mEndChar & _
                "53 " & Format$(mBarNo, "!" & String(15, "@")) & " 00/00/0000" & mEndChar & _
                "54 00000 00 " & strStat & mEndChar & strItems & _
                ETX & mEndChar & _
                EOT & mEndChar
    
              ' "54 000 00 " & strStat & mEndChar & strItems & _

    
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub