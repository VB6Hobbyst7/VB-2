VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISErrorShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISErrorShow.cls
'   �ۼ���  : �̻��
'   ��  ��  : ������������ ǥ���ϴ� Ŭ����
'   �ۼ���  : 2004-05-11
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : ������������ ǥ��
'   �μ� :
'       - pIntErrors : �������̽� ���� �÷���
'-----------------------------------------------------------------------------'
Public Sub ShowErrors(ByRef pIntErrors As clsIISIntErrors)
    With frmIISIntErrors
        .IntErrors = pIntErrors
        .Show vbModal
    End With
End Sub

Private Sub Class_Terminate()
    Unload frmIISIntErrors
End Sub