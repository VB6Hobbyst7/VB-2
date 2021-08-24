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
'   ��  ��  : LPIA-NV7 ��� �������� Ŭ����
'   �ۼ���  : 2004-10-15
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '���ڵ��ȣ
Private mMsgType    As String   'Message Type (1: Routine, 4:Stat)
Private mWorkNo     As Long     'Worksheet No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get MsgType() As String
    MsgType = mMsgType
End Property

Public Property Let MsgType(ByVal vData As String)
    mMsgType = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult       As clsIISResult     '������� Ŭ����
    Dim objQCResult     As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase      As String       '������ �˻��
    Dim strItems        As String       '�۽��� �˻��׸�
    Dim strWorkNo       As String       '�۽��� WorkNo
    Dim strMsgType      As String       '�۽��� Message Type(2:Routin, 3:Stat)
    Dim strOutput       As String       '�۽��� ������
    Dim lngSendCnt      As Long         '�۽��� �˻��׸� ����
    
    '## �۽��� �˻��׸� ��ȸ
    If pAccInfo.QcFg = "0" Then     '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(Format$(objResult.IntNm.IntBase, "###"), "@@@")
            strItems = strItems & strIntBase
            lngSendCnt = lngSendCnt + 1
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(Format$(objQCResult.IntNm.IntBase, "###"), "@@@")
            strItems = strItems & strIntBase
            lngSendCnt = lngSendCnt + 1
        Next
        Set objQCResult = Nothing
    End If
    
    '## �ִ� 16������ ���۰����ϰ� ���ºκ��� ��������!
    strItems = strItems & Space(3 * (16 - lngSendCnt))
    
    '## Worksheet No, Message Type
    strMsgType = IIf(mMsgType = "1", "2", "3")
    If strMsgType = "3" Then
        strWorkNo = Space(4)
    Else
        mWorkNo = mWorkNo + 1
        strWorkNo = Format$(mWorkNo, "0000")
    End If
    
    strOutput = STX & strMsgType & strWorkNo & Format$(mBarNo, "!" & String$(13, "@")) & _
                strItems & ETX

    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� �������� ���ڿ� ��ȸ
'   ��ȯ : ���������� �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strWorkNo   As String   '�۽��� Worksheet No
    Dim strMsgType  As String   '�۽��� Message Type(2:Routin, 3:Stat)
    Dim strItems    As String   '�۽��� �˻��׸�
    Dim strOutput   As String   '�۽��� ������
    
    '## Worksheet No, Message Type
    strItems = Space(3 * 16)
    strMsgType = IIf(mMsgType = "1", "2", "3")
    If strMsgType = "3" Then
        strWorkNo = Space(4)
    Else
        mWorkNo = mWorkNo + 1
        strWorkNo = Format$(mWorkNo, "0000")
    End If
    
    strOutput = STX & strMsgType & strWorkNo & Format$(mBarNo, "!" & String$(13, "@")) & _
                strItems & ETX
    
    GetNoOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : Control Information ���ڿ� ��ȸ
'-----------------------------------------------------------------------------'
Public Function GetControlInfo() As String
    Dim strOutput As String
    
    strOutput = STX & "900" & ETX
    GetControlInfo = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mMsgType = ""
End Sub
