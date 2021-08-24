VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISLASCHST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISLASCHST.cls
'   �ۼ���  : �̻��
'   ��  ��  : LASC-HST ����� ǥ�� Ŭ����
'   �ۼ���  : 2005-09-15
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mLASC As frmIISLASCHST

'-----------------------------------------------------------------------------'
'   ��� : ������� ǥ��
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pEqpKey : ���Ű
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## ��� �ɼ���ȸ
    Call GetOption(pEqpKey)
    
    If mLASC Is Nothing Then Set mLASC = New frmIISLASCHST
    With mLASC
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

On Error Resume Next
    mPort = mReadINI(strFileNm, pEqpKey, "Port")
    mBaudRate = mReadINI(strFileNm, pEqpKey, "BaudRate")
    mDataBit = mReadINI(strFileNm, pEqpKey, "DataBit")
    mStopBit = mReadINI(strFileNm, pEqpKey, "StopBit")
    mParityBit = mReadINI(strFileNm, pEqpKey, "ParityBit")
    mInterval = mReadINI(strFileNm, pEqpKey, "Interval")
End Sub

Private Sub Class_Terminate()
    Unload mLASC
    Set mLASC = Nothing
End Sub
