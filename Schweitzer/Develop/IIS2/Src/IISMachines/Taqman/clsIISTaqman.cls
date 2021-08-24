VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTaqman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTaqman.cls
'   작성자  : 오세원
'   내  용  : Taqman 장비폼 표시 클래스
'   작성일  : 2014-12-29
'   버  전  :
'       1. 1.0.0: 오세원(2014-12-29)
'-----------------------------------------------------------------------------'

Option Explicit

Private mTaqman As frmIISTaqman

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    
    If mTaqman Is Nothing Then Set mTaqman = New frmIISTaqman
    With mTaqman
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    Unload mTaqman
    Set mTaqman = Nothing
End Sub


