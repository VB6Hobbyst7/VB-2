VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISXE2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISXE2100.cls
'   작성자  : 이상대
'   내  용  : XE-2100 장비폼 표시 클래스
'   작성일  : 2005-02-15
'   버  전  :
'       1. 1.0.1: 이상대(2005-06-24)
'-----------------------------------------------------------------------------'

Option Explicit

Private mXE2100 As frmIISXE2100

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.1: 이상대(2005-06-24)
    '   - 한 PC에서 같은장비를 2대 사용할수 있도록 수정
    
    If mXE2100 Is Nothing Then Set mXE2100 = New frmIISXE2100
    With mXE2100
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.1: 이상대(2005-06-24)
    '   - 한 PC에서 같은장비를 2대 사용할수 있도록 수정
    
    Unload mXE2100
    Set mXE2100 = Nothing
End Sub

