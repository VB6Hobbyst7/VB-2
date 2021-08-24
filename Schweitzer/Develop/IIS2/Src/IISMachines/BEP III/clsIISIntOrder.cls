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
'   ���ϸ�  : clsIISBEPIII.cls
'   �ۼ���  : �̻��
'   ��  ��  : BEP III �������� Ŭ����
'   �ۼ���  : 2005-01-27
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
    Dim strIntBase    As String           '������ �˻��
    Dim strItems      As String           '�۽��� �˻��׸� ���ڿ�
    Dim strOutput     As String           '�۽��� ���ڿ�

    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(mGetP(objResult.IntNm.IntBase, 1, "|"), "!" & String$(6, "@"))
            If strItems = "" Then
                strItems = strIntBase
            Else
                strItems = strItems & ";" & strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(mGetP(objQCResult.IntNm.IntBase, 1, "|"), "!" & String$(6, "@"))
            If strItems = "" Then
                strItems = strIntBase
            Else
                strItems = strItems & ";" & strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If
    
    strOutput = pAccInfo.PtId & ";" & mBarNo & ";" & strItems
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub