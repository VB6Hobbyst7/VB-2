VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISUriscanPro3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISUriscanPro.cls
'   �ۼ���  : �̻��
'   ��  ��  : Uriscan Pro ����� ǥ�� Ŭ����
'   �ۼ���  : 2004-06-02
'   ��  ��  :
'       1. 1.0.2: �̻��(2005-03-03)
'          - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
'-----------------------------------------------------------------------------'

Option Explicit

Private mUriscan As frmIISUriscanPro3

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.2: �̻��(2005-03-03)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mUriscan Is Nothing Then Set mUriscan = New frmIISUriscanPro3
    With mUriscan
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
    
    Unload mUriscan
    Set mUriscan = Nothing
End Sub

