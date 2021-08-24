VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISKX21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISKX21.cls
'   �ۼ���  : �̻��
'   ��  ��  : KX-21 ����� ǥ�� Ŭ����
'   �ۼ���  : 2004-06-29
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mKX21 As frmIISKX21

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    If mKX21 Is Nothing Then Set mKX21 = New frmIISKX21
    With mKX21
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    Unload mKX21
    Set mKX21 = Nothing
End Sub
