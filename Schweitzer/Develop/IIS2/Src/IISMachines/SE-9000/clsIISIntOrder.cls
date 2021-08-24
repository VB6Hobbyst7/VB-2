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
'   ��  ��  : SE-9000 ��� �������� Ŭ����
'   �ۼ���  : 2004-10-13
'   ��  ��  :
'       1. 1.0.1: �̻��(2004-12-07)
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String   '���ڵ��ȣ
Private mRackNo As String   'Rack No
Private mPos    As String   'Tube Position

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '������ �˻�� ���ڿ�
    Dim strBarNo    As String           '�۽��� ���ڵ��ȣ
    Dim strPtId     As String           '�۽��� ȯ��ID
    Dim strSex      As String           '�۽��� ����
    Dim strOutput   As String           '�۽��� ������

    strItems = String$(24, "0")
    If pAccInfo.QcFg = "0" Then     '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = mGetP(objQCResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objQCResult = Nothing
    End If
    
    '## 1.0.1: �̻��(2004-12-07)
    '   - ȯ��ID�� ������ ��� "0"�����ʴ� ��������
    strBarNo = Format$(mBarNo, String$(13, "0"))
    strPtId = pAccInfo.PtId
    If IsNumeric(strPtId) Then
        strPtId = Format$(strPtId, String$(13, "0"))
    Else
        strPtId = String$(13 - Len(strPtId), "0") & strPtId
    End If
    strSex = IIf(UCase$(pAccInfo.Sex) = "M", 1, 2)
    
    strOutput = STX & "S1        " & strBarNo & mRackNo & mPos & "1" & strPtId & Space(25) & _
                strSex & Space(71) & strItems & Space(7) & ETX
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ������ �������ڿ��� ��ȸ
'   ��ȯ : ���������� ������ �������ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strBarNo    As String   '�۽��� ���ڵ��ȣ
    Dim strOutput   As String   '�۽��� ������
    
    strBarNo = Format$(mBarNo, String$(13, "0"))
    
    strOutput = STX & "S0        " & strBarNo & mRackNo & mPos & "1" & Space(110) & _
                String$(24, "0") & Space(7) & ETX
    Debug.Print "Len(strOutput)=" & Len(strOutput)
    GetNoOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRackNo = ""
    mPos = ""
End Sub