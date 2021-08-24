VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISiQ200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISiQ200.cls
'   작성자  : 오세원
'   내  용  : iQ200 장비폼 표시 클래스
'   작성일  : 2007-10-26
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private miQ200 As frmIISiQ200

'-----------------------------------------------------------------------------'
'   기능 : 장비폼을 표시
'   인수 :
'       - pEqpCd  : 장비코드
'       - pEqpKey : 장비키
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pEqpCd As String, ByVal pEqpKey As String)
    '## 1.0.2: 이상대(2005-03-03)
    '   - 한 PC에서 같은장비를 2대 사용할수 있도록 수정
    
    If miQ200 Is Nothing Then Set miQ200 = New frmIISiQ200
    With miQ200
        .EqpCd = pEqpCd
        .EqpKey = pEqpKey
        .MDIActiveX.WindowState = ccMaximize
        .MDIActiveX.ZOrder 0
        .Show
    End With
End Sub

Private Sub Class_Terminate()
    '## 1.0.2: 이상대(2005-03-03)
    '   - 한 PC에서 같은장비를 2대 사용할수 있도록 수정
    
    Unload miQ200
    Set miQ200 = Nothing
End Sub


