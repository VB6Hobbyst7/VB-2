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
'   파일명  : clsIISINFINITY.cls
'   작성자  : 오세원
'   내  용  : IT3000 장비폼 표시 클래스
'   작성일  : 2021-07-05
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mINFINITY As frmIISINFINITY

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 장비별 옵션조회
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
'   기능 : IIS.ini파일에서 장비별 옵션 조회
'   인수 :
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub GetOption(ByVal pEqpKey As String)
    Dim strFileNm   As String   'INI파일 경로+파일명

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

