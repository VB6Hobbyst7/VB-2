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
'   파일명  : clsIISInstrShow.cls (연천군보건의료원)
'   작성자  :
'   내  용  : 해당 장비폼을 표시하는 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISActDiff5
Private mE002 As clsIISAdvia120
Private mE003 As clsIISArchitect
Private mE004 As clsIISAU680
Private mE005 As clsIISRapidLab744
Private mE006 As clsIISUriscanPro



'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey

        Case "ActDiff5"
            If mE001 Is Nothing Then Set mE001 = New clsIISActDiff5
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "Advia2120"
            If mE002 Is Nothing Then Set mE002 = New clsIISAdvia120
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "Architect"
            If mE003 Is Nothing Then Set mE003 = New clsIISArchitect
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "AU680"
            If mE004 Is Nothing Then Set mE004 = New clsIISAU680
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab744"
            If mE005 Is Nothing Then Set mE005 = New clsIISRapidLab744
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "UrinscanPro"
            If mE006 Is Nothing Then Set mE006 = New clsIISUriscanPro
            Call mE006.ShowForm(pEqpCd, pEqpKey)

    End Select
End Sub

Private Sub Class_Terminate()
    Set mE001 = Nothing
    Set mE002 = Nothing
    Set mE003 = Nothing
    Set mE004 = Nothing
    Set mE005 = Nothing
    Set mE006 = Nothing
End Sub

