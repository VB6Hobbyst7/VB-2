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
'   ��  ��  : ENERGIUM �������� Ŭ����
'   �ۼ���  : 2014-07-28
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '���ڵ��ȣ

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '������� ��ü
    Dim objQCResult   As clsIISQCResult   'QC������� ��ü
    Dim strIntBase()  As String           '������ �˻��
    Dim strItems      As String           '�۽��� �˻��׸� ���ڿ�
    Dim strOutput     As String           '�۽��� ���ڿ�
    Dim intCnt        As Integer
    Dim intCnt2       As Integer
    
    intCnt = 0
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            intCnt = intCnt + 1
            ReDim Preserve strIntBase(intCnt)
            If Trim(objResult.IntNm.IntBase) <> "" Then
                strIntBase(intCnt) = mGetP(objResult.IntNm.IntBase, 1, "|")
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            ReDim Preserve strIntBase(intCnt)
            strIntBase(intCnt) = mGetP(objResult.IntNm.IntBase, 1, "|")
        Next
        Set objQCResult = Nothing
    End If
    
    'OBR||531455510078||TEST|||20140720003156||||A||||||||||||||||||
    For intCnt2 = 1 To intCnt
        strOutput = strOutput & strIntBase(intCnt2) & "|"
    Next
    
    If strOutput <> "" Then
        strOutput = Mid(strOutput, 1, Len(strOutput) - 1)
    End If
    
    GetOrder = strOutput

End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub