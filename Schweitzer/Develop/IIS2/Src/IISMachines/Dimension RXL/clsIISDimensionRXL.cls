VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISDimensionRXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISDimensionRXL.cls
'   �ۼ���  : �̻��
'   ��  ��  : Dimension RXL ����� ǥ�� Ŭ����
'   �ۼ���  : 2004-06-21
'   ��  ��  :
'       1. 1.0.2: ������(2007-10-19)
'          - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
'-----------------------------------------------------------------------------'

Option Explicit

Private mDRXLM As frmIISDimensionRXL

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.2: ������(2007-10-19)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    
    If mDRXLM Is Nothing Then Set mDRXLM = New frmIISDimensionRXL
    With mDRXLM
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.2: ������(2007-10-19)
    '   - �� PC���� ������� 2�� ����Ҽ� �ֵ��� ����
    Unload mDRXLM
    Set mDRXLM = Nothing
End Sub
