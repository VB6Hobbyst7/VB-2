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
'   파일명  : clsIISInstrShow.cls (단국대병원)
'   작성자  : 오세원
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2006-06-23
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISXE2100         'XE-2100
Private mE002 As clsIISXS1000iASTM         'XS-1000i
Private mE003 As clsIISStago          'STA Compact CT
Private mE004 As clsIISRapidLab348    'RapidLab 348
Private mE005 As clsIISRapidLab744    'RapidLab 744
Private mE006 As clsIISHitachi7180    'Hitachi 7180
Private mE007 As clsIISVitros250      'Vitros 250
Private mE008 As clsIISCentaur        'Centaur
Private mE009 As clsIISAxsym          'Axsym
Private mE010 As clsIISUriscanPro     'UrinScanPro II
Private mE012 As clsIISPFA100         'PFA100
Private mE013 As clsIISMultiReader    'MultiReader



'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "XE-2100"
            If mE001 Is Nothing Then Set mE001 = New clsIISXE2100
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "XS-1000i"
            If mE002 Is Nothing Then Set mE002 = New clsIISXS1000iASTM
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "Stago"
            If mE003 Is Nothing Then Set mE003 = New clsIISStago
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab348"
            If mE004 Is Nothing Then Set mE004 = New clsIISRapidLab348
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab744"
            If mE005 Is Nothing Then Set mE005 = New clsIISRapidLab744
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7180"
            If mE006 Is Nothing Then Set mE006 = New clsIISHitachi7180
            Call mE006.ShowForm(pEqpCd, pEqpKey)
        Case "Vitros250"
            If mE007 Is Nothing Then Set mE007 = New clsIISVitros250
            Call mE007.ShowForm(pEqpCd, pEqpKey)
        Case "Centaur"
            If mE008 Is Nothing Then Set mE008 = New clsIISCentaur
            Call mE008.ShowForm(pEqpCd, pEqpKey)
        Case "Axsym"
            If mE009 Is Nothing Then Set mE009 = New clsIISAxsym
            Call mE009.ShowForm(pEqpCd, pEqpKey)
        Case "UrinScanPro"
            If mE010 Is Nothing Then Set mE010 = New clsIISUriscanPro
            Call mE010.ShowForm(pEqpCd, pEqpKey)
        Case "PFA100"
            If mE012 Is Nothing Then Set mE012 = New clsIISPFA100
            Call mE012.ShowForm(pEqpCd, pEqpKey)
        Case "MultiReader"
            If mE013 Is Nothing Then Set mE013 = New clsIISMultiReader
            Call mE013.ShowForm(pEqpCd, pEqpKey)
    End Select
End Sub

Private Sub Class_Terminate()
    Set mE001 = Nothing
    Set mE002 = Nothing
    Set mE003 = Nothing
    Set mE004 = Nothing
    Set mE005 = Nothing
    Set mE006 = Nothing
    Set mE007 = Nothing
    Set mE008 = Nothing
    Set mE009 = Nothing
    Set mE010 = Nothing
    Set mE012 = Nothing
    Set mE013 = Nothing
End Sub
