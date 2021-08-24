VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISModularDPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISModularDPE.cls
'   작성자  : 이상대
'   내  용  : Modular DPE 장비폼 표시 클래스
'   작성일  : 2005-06-27
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mModularDPE As frmIISModularDPE

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 장비별 옵션조회
    Call GetOption(pEqpKey)
    
    If mModularDPE Is Nothing Then Set mModularDPE = New frmIISModularDPE
    With mModularDPE
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
    OPTVERSION = mReadINI(strFileNm, pEqpKey, "VERSION")
End Sub

Private Sub Class_Terminate()
    Unload mModularDPE
End Sub

