VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISChorusTRIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISiQ200.cls
'   �ۼ���  : ������
'   ��  ��  : iQ200 ����� ǥ�� Ŭ����
'   �ۼ���  : 2007-10-26
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mChorusTRIO As frmIISChorusTRIO

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.2: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mChorusTRIO Is Nothing Then Set mChorusTRIO = New frmIISChorusTRIO
    With mChorusTRIO
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.2: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    Unload mChorusTRIO
    Set mChorusTRIO = Nothing
End Sub

