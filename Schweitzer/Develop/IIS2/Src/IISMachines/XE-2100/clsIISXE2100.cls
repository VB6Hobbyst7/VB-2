VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISXE2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISXE2100.cls
'   �ۼ���  : �̻��
'   ��  ��  : XE-2100 ����� ǥ�� Ŭ����
'   �ۼ���  : 2005-02-15
'   ��  ��  :
'       1. 1.0.1: �̻��(2005-06-24)
'-----------------------------------------------------------------------------'

Option Explicit

Private mXE2100 As frmIISXE2100

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.1: �̻��(2005-06-24)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mXE2100 Is Nothing Then Set mXE2100 = New frmIISXE2100
    With mXE2100
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
    
    Unload mXE2100
    Set mXE2100 = Nothing
End Sub
