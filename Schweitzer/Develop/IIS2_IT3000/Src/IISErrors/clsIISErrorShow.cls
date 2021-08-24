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
'   파일명  : clsIISErrorShow.cls
'   작성자  : 이상대
'   내  용  : 에러정보폼을 표시하는 클래스
'   작성일  : 2004-05-11
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   기능 : 에러정보폼을 표시
'   인수 :
'       - pIntErrors : 인터페이스 에러 컬렉션
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
