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
'   작성자  :
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2004-10-05
'   버  전  :
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
Private mE037 As clsIISGem3000   'Gem3000

'** By S.W.Oh 2007.10.19
'   응급검사실 D-RXLM 추가
Private mE039 As clsIISDimensionRXL 'D-RXLM2

'** By S.W.Oh 2008.02.18
'   Sysmex XE-2100 추가(ASTM)
Private mE040 As clsIISXE2100ASTM 'XE2100

Private mE042 As clsIISDXC800        'DXC800

'** By S.W.Oh 2010.06.29
Private mE048 As clsIISNS1000  'NS1000

'** By MAJESTIC 2011.01.28
'   Sysmex XT-2000i 추가(ASTM)
Private mE051 As clsIISXT2000i       'XT-2000i
Private mE055 As clsIISChorusTRIO    'ChorusTRIO
Private mE056 As clsIISRapidLab348_1 'RapidLab348-1
Private mE057 As clsIISRapidLab348_2 'RapidLab348_2
Private mE058 As clsIISRapidLab348_3 'RapidLab348_3
Private mE059 As clsIISRapidLab348_4 'RapidLab348_4
'** By MAJESTIC 2012.04.13
'   Hitachi7600/Hitachi7189 추가
Private mE061 As clsIISHitachi7600_1
Private mE062 As clsIISHitachi7180
Private mE063 As clsIISHitachi7180_1
Private mE064 As clsIISMultiReader
Private mE067 As clsIISIT1000_1
Private mE075 As clsIISGemini

'** By S.W.Oh 2013.07.18
Private mE076 As clsIISVitek3       'Vitek 3
Private mE077 As clsIISVitek4       'Vitek 4

'** By S.W.Oh 2013.12.11
'Private mE078 As clsIISUF1000i     'UF1000i
Private mE078 As clsIISUF1000i
Private mE079 As clsIISCoa1        'STA-R 1
Private mE080 As clsIISCoaER       'STA-R 2

'** By S.W.Oh 2014.01.21
Private mE081 As clsIISCobasE411    'CobasE411

'** By S.W.Oh 2014.01.23
Private mE082 As clsIISCobasE601    'CobasE601

'** By S.W.Oh 2014.07.30
Private mE086 As clsIISGem3500_1    'GEM3500-1
Private mE087 As clsIISGem3500_2    'GEM3500-2

'** By S.W.Oh 2014.12.29
Private mE089 As clsIISTaqman       'Taqman
Private mE090 As clsIISTriage       'Triage
Private mE091 As clsIISSofia        'Sofia

'==> 092,093 XN에서 사용

'** By S.W.Oh 2015.01.28
Private mE094 As clsIISSofia2       'Sofia2

'** By S.W.Oh 2015.05.28
Private mE095 As clsIISGeneXpert    'GeneXpert

'** By S.W.Oh 2015.08.25
Private mE096 As clsIISB4C          'B4C

'** By S.W.Oh 2015.11.11
Private mE097  As clsIISGreenCross  '녹십자연동

'** By S.W.Oh 2016.03.10
Private mE098  As clsIISNSPRIME

'** By S.W.Oh 2016.09.06
Private mE088  As clsIISHitachi7600_2

'** By S.W.Oh 2016.11.29
Private mE105  As clsIISCobasC513

'** By S.W.Oh 2016.12.19
Private mE106  As clsIISSofia3       'Sofia3

'** By S.W.Oh 2016.12.22
Private mE107  As clsIISMEDIWISS     'MEDIWISS

'** By S.W.Oh 2016.12.22
Private mE108  As clsIISCFX96        'CFX96

'** By S.W.Oh 2017.04.05
Private mE109  As clsIISIH500        'IH500

'** By S.W.Oh 2017.04.25
Private mE110  As clsIISUriscanPro3
Private mE111  As clsIISAlloStation

'** By S.W.Oh 2017.09.18
Private mE112  As clsIISCFX96_RV16

'** By S.W.Oh 2017.11.16
Private mE113  As clsIISEPOC_1
Private mE114  As clsIISEPOC_2

'** By S.W.Oh 2018.01.08
Private mE115  As clsIISLiaison

'** By majestic 2018.02.08
Private mE116 As clsIISCobasE411_Disk    'CobasE411_Disk

'** By S.W.Oh 2018.04.12
Private mE117  As clsIISMGIT960

'** By S.W.Oh 2018.11.15
Private mE118  As clsIISUF5000

'** By S.W.Oh 2019.01.07
Private mE119  As clsIISGem5000

'** By S.W.Oh 2019.09.18
Private mE120  As clsIISGem5000_2

'** By S.W.Oh 2019.10.24
Private mE121  As clsIISOCSensorM

'** By S.W.Oh 2019.10.24
Private mE122  As clsIISAlinity

'** By S.W.Oh 2020.10.19
Private mE123  As clsIISGem5000_OP

'** By S.W.Oh 2020.10.28
Private mE124  As clsIISLABOSPECT_1
Private mE125  As clsIISLABOSPECT_2
Private mE126  As clsIISLABOSPECT_ER

'** By S.W.Oh 2021.01.11
Private mE127  As clsIISPATHFAST


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
            If mE037 Is Nothing Then Set mE037 = New clsIISGem3000
            Call mE037.ShowForm(pEqpCd, pEqpKey)
        Case "D-RXLM2"
            If mE039 Is Nothing Then Set mE039 = New clsIISDimensionRXL
            Call mE039.ShowForm(pEqpCd, pEqpKey)
        Case "XE2100"
            If mE040 Is Nothing Then Set mE040 = New clsIISXE2100ASTM
            Call mE040.ShowForm(pEqpCd, pEqpKey)
        Case "DXC800"
            If mE042 Is Nothing Then Set mE042 = New clsIISDXC800
            Call mE042.ShowForm(pEqpCd, pEqpKey)
        Case "NS1000"
            If mE048 Is Nothing Then Set mE048 = New clsIISNS1000
            Call mE048.ShowForm(pEqpCd, pEqpKey)
        Case "ChorusTRIO"
            If mE055 Is Nothing Then Set mE055 = New clsIISChorusTRIO
            Call mE055.ShowForm(pEqpCd, pEqpKey)
        Case "XT-2000i"
            If mE051 Is Nothing Then Set mE051 = New clsIISXT2000i
            Call mE051.ShowForm(pEqpCd, pEqpKey)
            'RapidLab348-1
        Case "RapidLab348-1"
            If mE056 Is Nothing Then Set mE056 = New clsIISRapidLab348_1
            Call mE056.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab348-2"
            If mE057 Is Nothing Then Set mE057 = New clsIISRapidLab348_2
            Call mE057.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab348-3"
            If mE058 Is Nothing Then Set mE058 = New clsIISRapidLab348_3
            Call mE058.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab348-4"
            If mE059 Is Nothing Then Set mE059 = New clsIISRapidLab348_4
            Call mE059.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7600-1"
            If mE061 Is Nothing Then Set mE061 = New clsIISHitachi7600_1
            Call mE061.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7180"
            If mE062 Is Nothing Then Set mE062 = New clsIISHitachi7180
            Call mE062.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7180-1"
            If mE063 Is Nothing Then Set mE063 = New clsIISHitachi7180_1
            Call mE063.ShowForm(pEqpCd, pEqpKey)
        Case "MultiReader"
            If mE064 Is Nothing Then Set mE064 = New clsIISMultiReader
            Call mE064.ShowForm(pEqpCd, pEqpKey)
        Case "IT1000"
            If mE067 Is Nothing Then Set mE067 = New clsIISIT1000_1
            Call mE067.ShowForm(pEqpCd, pEqpKey)
        Case "Gemini"
            If mE075 Is Nothing Then Set mE075 = New clsIISGemini
            Call mE075.ShowForm(pEqpCd, pEqpKey)
        '** By S.W.Oh 2013.07.18
        Case "Vitek 3"
            If mE076 Is Nothing Then Set mE076 = New clsIISVitek3
            Call mE076.ShowForm(pEqpCd, pEqpKey)
        Case "Vitek 4"
            If mE077 Is Nothing Then Set mE077 = New clsIISVitek4
            Call mE077.ShowForm(pEqpCd, pEqpKey)
        '== CHO N.S. 2013.12.11 ==
        Case "UF1000i"
            If mE078 Is Nothing Then Set mE078 = New clsIISUF1000i
            Call mE078.ShowForm(pEqpCd, pEqpKey)
        Case "Coa1"
            If mE079 Is Nothing Then Set mE079 = New clsIISCoa1
            Call mE079.ShowForm(pEqpCd, pEqpKey)
        Case "CoaER"
            If mE080 Is Nothing Then Set mE080 = New clsIISCoaER
            Call mE080.ShowForm(pEqpCd, pEqpKey)
        Case "CobasE411"
            If mE081 Is Nothing Then Set mE081 = New clsIISCobasE411
            Call mE081.ShowForm(pEqpCd, pEqpKey)
        Case "CobasE601"
            If mE082 Is Nothing Then Set mE082 = New clsIISCobasE601
            Call mE082.ShowForm(pEqpCd, pEqpKey)
        Case "GEM3500_1"
            If mE086 Is Nothing Then Set mE086 = New clsIISGem3500_1
            Call mE086.ShowForm(pEqpCd, pEqpKey)
        Case "GEM3500_2"
            If mE087 Is Nothing Then Set mE087 = New clsIISGem3500_2
            Call mE087.ShowForm(pEqpCd, pEqpKey)
        Case "Taqman"
            If mE089 Is Nothing Then Set mE089 = New clsIISTaqman
            Call mE089.ShowForm(pEqpCd, pEqpKey)
        Case "Triage"
            If mE090 Is Nothing Then Set mE090 = New clsIISTriage
            Call mE090.ShowForm(pEqpCd, pEqpKey)
        Case "Sofia"
            If mE091 Is Nothing Then Set mE091 = New clsIISSofia
            Call mE091.ShowForm(pEqpCd, pEqpKey)
        Case "Sofia2"
            If mE094 Is Nothing Then Set mE094 = New clsIISSofia2
            Call mE094.ShowForm(pEqpCd, pEqpKey)
        Case "GeneXpert"
            If mE095 Is Nothing Then Set mE095 = New clsIISGeneXpert
            Call mE095.ShowForm(pEqpCd, pEqpKey)
        Case "B4C"
            If mE096 Is Nothing Then Set mE096 = New clsIISB4C
            Call mE096.ShowForm(pEqpCd, pEqpKey)
        Case "GreenCross"
            If mE097 Is Nothing Then Set mE097 = New clsIISGreenCross
            Call mE097.ShowForm(pEqpCd, pEqpKey)
        Case "NSPRIME"
            If mE098 Is Nothing Then Set mE098 = New clsIISNSPRIME
            Call mE098.ShowForm(pEqpCd, pEqpKey)
        Case "Hitachi7600-2"
            If mE088 Is Nothing Then Set mE088 = New clsIISHitachi7600_2
            Call mE088.ShowForm(pEqpCd, pEqpKey)
        Case "CobasC513"
            If mE105 Is Nothing Then Set mE105 = New clsIISCobasC513
            Call mE105.ShowForm(pEqpCd, pEqpKey)
        Case "Sofia3"
            If mE106 Is Nothing Then Set mE106 = New clsIISSofia3
            Call mE106.ShowForm(pEqpCd, pEqpKey)
        Case "MEDIWISS"
            If mE107 Is Nothing Then Set mE107 = New clsIISMEDIWISS
            Call mE107.ShowForm(pEqpCd, pEqpKey)
        Case "CFX96"
            If mE108 Is Nothing Then Set mE108 = New clsIISCFX96
            Call mE108.ShowForm(pEqpCd, pEqpKey)
        Case "IH500"
            If mE109 Is Nothing Then Set mE109 = New clsIISIH500
            Call mE109.ShowForm(pEqpCd, pEqpKey)
        Case "Uriscan Pro-3"
            If mE110 Is Nothing Then Set mE110 = New clsIISUriscanPro3
            Call mE110.ShowForm(pEqpCd, pEqpKey)
        Case "AlloStation"
            If mE111 Is Nothing Then Set mE111 = New clsIISAlloStation
            Call mE111.ShowForm(pEqpCd, pEqpKey)
        Case "CFX96_RV16"
            If mE112 Is Nothing Then Set mE112 = New clsIISCFX96_RV16
            Call mE112.ShowForm(pEqpCd, pEqpKey)
        Case "EPOC_1"
            If mE113 Is Nothing Then Set mE113 = New clsIISEPOC_1
            Call mE113.ShowForm(pEqpCd, pEqpKey)
        Case "EPOC_2"
            If mE114 Is Nothing Then Set mE114 = New clsIISEPOC_2
            Call mE114.ShowForm(pEqpCd, pEqpKey)
        Case "LIAISON"
            If mE115 Is Nothing Then Set mE115 = New clsIISLiaison
            Call mE115.ShowForm(pEqpCd, pEqpKey)
        Case "CobasE411_Disk"
            If mE116 Is Nothing Then Set mE116 = New clsIISCobasE411_Disk
            Call mE116.ShowForm(pEqpCd, pEqpKey)
        Case "MGIT960"
            If mE117 Is Nothing Then Set mE117 = New clsIISMGIT960
            Call mE117.ShowForm(pEqpCd, pEqpKey)
        Case "UF5000"
            If mE118 Is Nothing Then Set mE118 = New clsIISUF5000
            Call mE118.ShowForm(pEqpCd, pEqpKey)
        Case "GEM5000"
            If mE119 Is Nothing Then Set mE119 = New clsIISGem5000
            Call mE119.ShowForm(pEqpCd, pEqpKey)
        Case "GEM5000_2"
            If mE120 Is Nothing Then Set mE120 = New clsIISGem5000_2
            Call mE120.ShowForm(pEqpCd, pEqpKey)
        Case "OCSENSORM"
            If mE121 Is Nothing Then Set mE121 = New clsIISOCSensorM
            Call mE121.ShowForm(pEqpCd, pEqpKey)
        Case "ALINITY"
            If mE122 Is Nothing Then Set mE122 = New clsIISAlinity
            Call mE122.ShowForm(pEqpCd, pEqpKey)
        Case "GEM5000_OP"
            If mE123 Is Nothing Then Set mE123 = New clsIISGem5000_OP
            Call mE123.ShowForm(pEqpCd, pEqpKey)
        Case "LABOSPECT_1"
            If mE124 Is Nothing Then Set mE124 = New clsIISLABOSPECT_1
            Call mE124.ShowForm(pEqpCd, pEqpKey)
        Case "LABOSPECT_2"
            If mE125 Is Nothing Then Set mE125 = New clsIISLABOSPECT_2
            Call mE125.ShowForm(pEqpCd, pEqpKey)
        Case "LABOSPECT_ER"
            If mE126 Is Nothing Then Set mE126 = New clsIISLABOSPECT_ER
            Call mE126.ShowForm(pEqpCd, pEqpKey)
        Case "PATHFAST"
            If mE127 Is Nothing Then Set mE127 = New clsIISPATHFAST
            Call mE127.ShowForm(pEqpCd, pEqpKey)
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
    Set mE042 = Nothing
    Set mE048 = Nothing
    Set mE055 = Nothing
    Set mE051 = Nothing
    Set mE056 = Nothing
    Set mE057 = Nothing
    Set mE058 = Nothing
    Set mE059 = Nothing
    Set mE061 = Nothing
    Set mE062 = Nothing
    Set mE063 = Nothing
    Set mE064 = Nothing
    Set mE067 = Nothing
    Set mE075 = Nothing
    Set mE076 = Nothing
    Set mE077 = Nothing
    Set mE078 = Nothing
    Set mE079 = Nothing
    Set mE080 = Nothing
    Set mE081 = Nothing
    Set mE082 = Nothing
    Set mE086 = Nothing
    Set mE089 = Nothing
    Set mE090 = Nothing
    Set mE091 = Nothing
    Set mE094 = Nothing
    Set mE095 = Nothing
    Set mE096 = Nothing
    Set mE097 = Nothing
    Set mE098 = Nothing
    Set mE088 = Nothing
    Set mE105 = Nothing
    Set mE106 = Nothing
    Set mE107 = Nothing
    Set mE108 = Nothing
    Set mE109 = Nothing
    Set mE110 = Nothing
    Set mE111 = Nothing
    Set mE112 = Nothing
    Set mE113 = Nothing
    Set mE114 = Nothing
    Set mE115 = Nothing
    Set mE116 = Nothing
    Set mE117 = Nothing
    Set mE118 = Nothing
    Set mE119 = Nothing
    Set mE120 = Nothing
    Set mE121 = Nothing
    Set mE122 = Nothing
    Set mE123 = Nothing
    Set mE124 = Nothing
    Set mE125 = Nothing
    Set mE126 = Nothing
    Set mE127 = Nothing
End Sub

