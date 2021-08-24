VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCA7000.cls
'   �ۼ���  : �̻��
'   ��  ��  : CA-7000 ����� ǥ�� Ŭ����
'   �ۼ���  : 2005-02-18
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## ��� �ɼ���ȸ
    Call GetOption(pEqpKey)

    With frmIISCA7000
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : IIS.ini���Ͽ��� ��� �ɼ� ��ȸ
'   �μ� :
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub GetOption(ByVal pEqpKey As String)
    Dim strFileNm   As String   'INI���� ���+���ϸ�
    
    strFileNm = IniPath & "IIS.ini"
    OPTFORMAT = mReadINI(strFileNm, pEqpKey, "FORMAT")
End Sub

Private Sub Class_Terminate()
    Unload frmIISCA7000
End Sub
