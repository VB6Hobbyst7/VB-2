VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISDPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISDPC.cls
'   �ۼ���  : ������
'   ��  ��  : DPC ����� ǥ�� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mDPC As frmIISDPC

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## ��� �ɼ���ȸ
    Call GetOption(pEqpKey)
    
    If mDPC Is Nothing Then Set mDPC = New frmIISDPC
    With mDPC
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
    OPTVERSION = mReadINI(strFileNm, pEqpKey, "VERSION")
End Sub

Private Sub Class_Terminate()
    Unload mDPC
End Sub
