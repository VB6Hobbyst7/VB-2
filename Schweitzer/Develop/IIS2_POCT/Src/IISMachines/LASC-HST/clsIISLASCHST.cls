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
'   파일명  : clsIISLASCHST.cls
'   작성자  : 이상대
'   내  용  : LASC-HST 장비폼 표시 클래스
'   작성일  : 2005-09-15
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mLASC As frmIISLASCHST

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 장비별 옵션조회
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
'   기능 : IIS.ini파일에서 장비별 옵션 조회
'   인수 :
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub GetOption(ByVal pEqpKey As String)
    Dim strFileNm   As String   'INI파일 경로+파일명

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

