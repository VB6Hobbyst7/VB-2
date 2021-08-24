VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIT3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIT3000.cls
'   �ۼ���  : ������
'   ��  ��  : IT3000 ����� ǥ�� Ŭ����
'   �ۼ���  : 2014-07-28
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mIT3000 As frmIISIT3000

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## ��� �ɼ���ȸ
    Call GetOption(pEqpKey)

    If mIT3000 Is Nothing Then Set mIT3000 = New frmIISIT3000
    
    With frmIISIT3000
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
    mOrderPath = mReadINI(strFileNm, pEqpKey, "OrderPath")
    mResultPath = mReadINI(strFileNm, pEqpKey, "ResultPath")
    mOrderFileNm = mReadINI(strFileNm, pEqpKey, "OrderFileNm")
    mResultFileNm = mReadINI(strFileNm, pEqpKey, "ResultFileNm")
    mOrderRefresh = mReadINI(strFileNm, pEqpKey, "OrderRefresh")
    mResultRefresh = mReadINI(strFileNm, pEqpKey, "ResultRefresh")
    mDB = mReadINI(strFileNm, pEqpKey, "DB")
    mUID = mReadINI(strFileNm, pEqpKey, "UID")
    mPW = mReadINI(strFileNm, pEqpKey, "PWD")

End Sub

Private Sub Class_Terminate()
    Unload frmIISIT3000
    Set mIT3000 = Nothing
End Sub
