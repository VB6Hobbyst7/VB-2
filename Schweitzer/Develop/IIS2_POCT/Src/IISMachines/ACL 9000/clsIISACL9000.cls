VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISACL9000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISACL9000.cls
'   �ۼ���  : �̻��
'   ��  ��  : ACL 9000 ����� ǥ�� Ŭ����
'   �ۼ���  : 2005-02-23
'   ��  ��  :
'       1. 1.0.1: �̻��(2005-06-21)
'-----------------------------------------------------------------------------'

Option Explicit

Private mACL9000 As frmIISACL9000

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.1: �̻��(2005-06-24)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mACL9000 Is Nothing Then Set mACL9000 = New frmIISACL9000
    With mACL9000
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.1: �̻��(2005-06-24)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    Unload mACL9000
    Set mACL9000 = Nothing
End Sub
