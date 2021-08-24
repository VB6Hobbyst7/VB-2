VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISInstrShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISInstrShow.cls (성모자애병원)
'   작성자  : 이상대
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2005-03-11
'   버  전  :
'       1. 1.1.0: 이상대(2005-05-08)
'          - Hitachi 7180 장비추가
'       2. 1.2.0: 이상대(2005-05-17)
'          - Coulter LH750 장비추가
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISHitachi7600  'Hitachi 7600
Private mE002 As clsIISHitachi7180  'Hitachi 7180
Private mE003 As clsIISLH750        'Coulter LH750

'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "Hitachi 7600"
            If mE001 Is Nothing Then Set mE001 = New clsIISHitachi7600
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi 7180"
            If mE002 Is Nothing Then Set mE002 = New clsIISHitachi7180
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "LH750"
            If mE003 Is Nothing Then Set mE003 = New clsIISLH750
            Call mE003.ShowForm(pEqpCd, pEqpKey)
    End Select
End Sub

Private Sub Class_Terminate()
    Set mE001 = Nothing
    Set mE002 = Nothing
    Set mE003 = Nothing
End Sub

