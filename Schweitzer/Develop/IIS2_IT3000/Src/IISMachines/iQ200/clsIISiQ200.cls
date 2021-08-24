VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISiQ200"
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

Private miQ200 As frmIISiQ200

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.2: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If miQ200 Is Nothing Then Set miQ200 = New frmIISiQ200
    With miQ200
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
    
    Unload miQ200
    Set miQ200 = Nothing
End Sub

