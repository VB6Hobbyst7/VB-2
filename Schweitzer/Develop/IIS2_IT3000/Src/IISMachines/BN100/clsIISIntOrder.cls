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
'   ��  ��  : BN100 ��� �������� Ŭ����
'   �ۼ���  : 2004-07-19
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo As String    '���ڵ��ȣ
Private mSeq   As String    '������ ������ Spread�� Row

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByRef pAccinfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻�� ���ڿ�
    Dim strOrder    As String           '�������� ���ڿ�
    Dim strTemp1    As String
    Dim strTemp2    As String
    
    If pAccinfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccinfo.Results
            strTemp1 = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            
            If strTemp1 <> strTemp2 Then
                strIntBase = strIntBase & Format$(strTemp1, "00")
                strTemp2 = strTemp1
            Else
                pAccinfo.SendCnt = pAccinfo.SendCnt - 1
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccinfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccinfo.QCResults
            strTemp1 = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            
            If strTemp1 <> strTemp2 Then
                strIntBase = strIntBase & Format$(strTemp1, "00")
                strTemp2 = strTemp1
            Else
                pAccinfo.SendCnt = pAccinfo.SendCnt - 1
            End If
        Next
        Set objQCResult = Nothing
    End If
    strIntBase = strIntBase & String$(40 - Len(strIntBase), Space(1))
    
    strOrder = STX & "J" & mBarNo & String$(29 - SPCLEN, Space(1)) & Space(8) & _
               strIntBase & Space(1) & vbCr & ETX
    GetOrder = strOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : Request Result Message ���ڿ� ��ȸ
'   �μ� :
'       - �������� Ŭ����
'   ��ȯ :
'       - Request Result Message
'       - ���̻� �䱸�� �˻��׸��� ������쿡�� "-1" ��ȯ
'-----------------------------------------------------------------------------'
Public Function GetReqResult(ByVal pAccinfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻�� ���ڿ�
    Dim strReqMsg   As String           'Request Result Message
    Dim strTemp1    As String
    Dim strTemp2    As String
    
    '## ������ۿ䱸�� �˻��׸��� �Ѱ��� �䱸�ؾ� �Ѵ�
    If pAccinfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccinfo.Results
            If objResult.SndChk = False Then
                strTemp1 = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
                
                If strTemp1 <> strTemp2 Then
                    strIntBase = Format$(strTemp1, "00")
                    strTemp2 = strTemp1
                End If
                objResult.SndChk = True
                Exit For
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccinfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccinfo.QCResults
            If objQCResult.SndChk = False Then
                strTemp1 = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
                
                If strTemp1 <> strTemp2 Then
                    strIntBase = Format$(strTemp1, "00")
                    strTemp2 = strTemp1
                End If
                objQCResult.SndChk = True
                Exit For
            End If
        Next
        Set objQCResult = Nothing
    End If
    
    If strIntBase = "" Then
        strReqMsg = "-1"
    Else
        strReqMsg = STX & "D" & CStr(mBarNo) & String$(29 - SPCLEN, Space(1)) & Space(1) & _
                    strIntBase & Space(1) & vbCr & ETX
    End If
    GetReqResult = strReqMsg
End Function

'-----------------------------------------------------------------------------'
'   ��� : Positive Acknowledge Message ��ȸ
'   ��ȯ : Positive Acknowledge Message
'-----------------------------------------------------------------------------'
Public Function GetPositiveMsg() As String
    GetPositiveMsg = STX & "A " & vbCr & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Positive Acknowledge Message ��ȸ
'   ��ȯ : Positive Acknowledge Message
'-----------------------------------------------------------------------------'
Public Function GetNegativeMsg() As String
    GetNegativeMsg = STX & "E " & vbCr & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mSeq = ""
End Sub