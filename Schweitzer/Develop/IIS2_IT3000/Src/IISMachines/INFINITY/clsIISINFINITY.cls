VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISINFINITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISINFINITY.cls
'   �ۼ���  : ������
'   ��  ��  : IT3000 ����� ǥ�� Ŭ����
'   �ۼ���  : 2021-07-05
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mINFINITY As frmIISINFINITY

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## ��� �ɼ���ȸ
    Call GetOption(pEqpKey)

    If mINFINITY Is Nothing Then Set mINFINITY = New frmIISINFINITY
    
    With frmIISINFINITY
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
    mBackUpPath = mReadINI(strFileNm, pEqpKey, "BackUpPath")
    mOrderFileNm = mReadINI(strFileNm, pEqpKey, "OrderFileNm")
    mResultFileNm = mReadINI(strFileNm, pEqpKey, "ResultFileNm")
    mOrderRefresh = mReadINI(strFileNm, pEqpKey, "OrderRefresh")
    mResultRefresh = mReadINI(strFileNm, pEqpKey, "ResultRefresh")
    mDB = mReadINI(strFileNm, pEqpKey, "DB")
    mUID = mReadINI(strFileNm, pEqpKey, "UID")
    mPW = mReadINI(strFileNm, pEqpKey, "PWD")

End Sub

Private Sub Class_Terminate()
    Unload frmIISINFINITY
    Set mINFINITY = Nothing
End Sub
