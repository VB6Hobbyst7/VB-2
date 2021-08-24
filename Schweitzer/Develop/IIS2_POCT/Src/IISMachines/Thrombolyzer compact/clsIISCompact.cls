VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCompact.cls
'   �ۼ���  : �̻��
'   ��  ��  : Thrombolyzer compact ����� ǥ�� Ŭ����
'   �ۼ���  : 2004-10-16
'   ��  ��  :
'       1. 1.0.3: �̻��(2005-03-03)
'          - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
'-----------------------------------------------------------------------------'

Option Explicit

Private mCompact As frmIISCompact

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.3: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mCompact Is Nothing Then Set mCompact = New frmIISCompact
    With mCompact
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.3: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    Unload mCompact
    Set mCompact = Nothing
End Sub
