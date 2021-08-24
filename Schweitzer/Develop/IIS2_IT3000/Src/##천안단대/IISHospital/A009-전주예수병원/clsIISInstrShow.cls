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
'   파일명  : clsIISInstrShow.cls (전주예수병원)
'   작성자  : 이상대
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2004-10-05
'   버  전  :
'       1. 1.1.0: 이상대(2004-11-18)
'          - RapidLab 850, 860 장비추가
'       2. 1.3.0: 이상대(2005-01-31)
'          - Vitek II 장비추가
'       3. 1.4.0: 이상대(2005-03-04)
'          - Thrombolyzer RackRotor 장비추가
'       4. 1.5.0: 이상대(2005-07-22)
'          - XT-1800i, CA-500 장비추가
'       5. 1.6.0: 이상대(2005-07-25)
'          - CA-1500 장비추가
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISUriscanPro   'Uriscan Pro+ 1
Private mE002 As clsIISUriscanPro   'Uriscan Pro+ 2
Private mE003 As clsIISDimensionRXL 'Dimension RXL
Private mE004 As clsIISRapidLab865  'RapidLab 865
Private mE005 As clsIISStks         'Stks-1
Private mE006 As clsIISStks         'Stks-2
Private mE007 As clsIISHitachi7600  'Hitachi 7600
Private mE008 As clsIISAxsym        'Axsym
Private mE009 As clsIISSE9000       'SE-9000
Private mE010 As clsIISVariantII    'Variant II
Private mE011 As clsIISCX3Delta     'CX3 Delta
Private mE012 As clsIISLpiaNv7      'LPIA-NV7
Private mE013 As clsIISCompact      'Thrombolyzer compact
Private mE014 As clsIISVitek        'Vitek
Private mE015 As clsIISRapidLab850  'RapidLab 850
Private mE016 As clsIISRapidLab860  'RapidLab 860
Private mE023 As clsIISVitekII      'Vitek II
Private mE024 As clsIISRackRotor    'Thrombolyzer RackRotor
Private mE025 As clsIISXT1800i      'XT-1800i
Private mE026 As clsIISCA500        'CA-500
Private mE027 As clsIISCA1500       'CA-1500

'** By M.G.Choi 2006.01.06
'   응급검사실 CA-1500 추가
Private mE028 As clsIISCA1500       'CA-1500ER

'** By M.G.Choi 2006.06.14
'   응급검사실 D-RXLM 추가
Private mE031 As clsIISDimensionRXL 'D-RXLM

'** By S.W.Oh 2006.08.02
'   검사실 Architect 추가
Private mE032 As clsIISArchitect  'Architect

'** By S.W.Oh 2007.03.27
'   검사실 TEST 1 추가
Private mE033 As clsIISTEST1   'ESR (TEST 1)

'** By S.W.Oh 2007.05.09
'   검사실 PCX 추가
Private mE034 As clsIISPCX   'PCX

'** By S.W.Oh 2007.06.11
'   검사실 Gem3000 추가
Private mE035 As clsIISGem3000   'Gem3000


'** By S.W.Oh 2007.10.19
'   응급검사실 D-RXLM 추가
Private mE039 As clsIISDimensionRXL 'D-RXLM2

'** By S.W.Oh 2008.02.18
'   Sysmex XE-2100 추가(ASTM)
Private mE040 As clsIISXE2100ASTM 'XE2100

'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "Uriscan Pro-1"
            If mE001 Is Nothing Then Set mE001 = New clsIISUriscanPro
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "Uriscan Pro-2"
            If mE002 Is Nothing Then Set mE002 = New clsIISUriscanPro
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "Dimension RXL"
            If mE003 Is Nothing Then Set mE003 = New clsIISDimensionRXL
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab 865"
            If mE004 Is Nothing Then Set mE004 = New clsIISRapidLab865
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "Stks-1"
            If mE005 Is Nothing Then Set mE005 = New clsIISStks
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "Stks-2"
            If mE006 Is Nothing Then Set mE006 = New clsIISStks
            Call mE006.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi 7600"
            If mE007 Is Nothing Then Set mE007 = New clsIISHitachi7600
            Call mE007.ShowForm(pEqpCd, pEqpKey)
        Case "Axsym"
            If mE008 Is Nothing Then Set mE008 = New clsIISAxsym
            Call mE008.ShowForm(pEqpCd, pEqpKey)
        Case "SE-9000"
            If mE009 Is Nothing Then Set mE009 = New clsIISSE9000
            Call mE009.ShowForm(pEqpCd, pEqpKey)
        Case "Variant II"
            If mE010 Is Nothing Then Set mE010 = New clsIISVariantII
            Call mE010.ShowForm(pEqpCd, pEqpKey)
        Case "CX3 Delta"
            If mE011 Is Nothing Then Set mE011 = New clsIISCX3Delta
            Call mE011.ShowForm(pEqpCd, pEqpKey)
        Case "LPIA-NV7"
            If mE012 Is Nothing Then Set mE012 = New clsIISLpiaNv7
            Call mE012.ShowForm(pEqpCd, pEqpKey)
        Case "Thrombolyzer compact"
            If mE013 Is Nothing Then Set mE013 = New clsIISCompact
            Call mE013.ShowForm(pEqpCd, pEqpKey)
        Case "Vitek"
            If mE014 Is Nothing Then Set mE014 = New clsIISVitek
            Call mE014.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab 850"
            If mE015 Is Nothing Then Set mE015 = New clsIISRapidLab850
            Call mE015.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab 860"
            If mE016 Is Nothing Then Set mE016 = New clsIISRapidLab860
            Call mE016.ShowForm(pEqpCd, pEqpKey)
        Case "Vitek II"
            If mE023 Is Nothing Then Set mE023 = New clsIISVitekII
            Call mE023.ShowForm(pEqpCd, pEqpKey)
        Case "Thrombolyzer RackRotor"
            If mE024 Is Nothing Then Set mE024 = New clsIISRackRotor
            Call mE024.ShowForm(pEqpCd, pEqpKey)
        Case "XT-1800i"
            If mE025 Is Nothing Then Set mE025 = New clsIISXT1800i
            Call mE025.ShowForm(pEqpCd, pEqpKey)
        Case "CA-500"
            If mE026 Is Nothing Then Set mE026 = New clsIISCA500
            Call mE026.ShowForm(pEqpCd, pEqpKey)
        Case "CA-1500"
            If mE027 Is Nothing Then Set mE027 = New clsIISCA1500
            Call mE027.ShowForm(pEqpCd, pEqpKey)
        Case "CA-1500ER"
            If mE028 Is Nothing Then Set mE028 = New clsIISCA1500
            Call mE028.ShowForm(pEqpCd, pEqpKey)
        Case "D-RXLM"
            If mE031 Is Nothing Then Set mE031 = New clsIISDimensionRXL
            Call mE031.ShowForm(pEqpCd, pEqpKey)
        Case "Architect"
            If mE032 Is Nothing Then Set mE032 = New clsIISArchitect
            Call mE032.ShowForm(pEqpCd, pEqpKey)
        Case "ESR"
            If mE033 Is Nothing Then Set mE033 = New clsIISTEST1
            Call mE033.ShowForm(pEqpCd, pEqpKey)
        Case "PCX1"
            If mE034 Is Nothing Then Set mE034 = New clsIISPCX
            Call mE034.ShowForm(pEqpCd, pEqpKey)
        Case "Gem3000"
            If mE035 Is Nothing Then Set mE035 = New clsIISGem3000
            Call mE035.ShowForm(pEqpCd, pEqpKey)
        Case "D-RXLM2"
            If mE039 Is Nothing Then Set mE039 = New clsIISDimensionRXL
            Call mE039.ShowForm(pEqpCd, pEqpKey)
        Case "XE2100"
            If mE040 Is Nothing Then Set mE040 = New clsIISXE2100ASTM
            Call mE040.ShowForm(pEqpCd, pEqpKey)
    
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
    Set mE011 = Nothing
    Set mE012 = Nothing
    Set mE013 = Nothing
    Set mE014 = Nothing
    Set mE015 = Nothing
    Set mE016 = Nothing
    Set mE023 = Nothing
    Set mE024 = Nothing
    Set mE025 = Nothing
    Set mE026 = Nothing
    Set mE027 = Nothing
    Set mE028 = Nothing
    Set mE031 = Nothing
    Set mE032 = Nothing
    Set mE033 = Nothing
    Set mE034 = Nothing
    Set mE039 = Nothing
    Set mE040 = Nothing
End Sub
