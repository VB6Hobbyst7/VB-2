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
'   작성자  : 이상대
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2006-06-23
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISLASCHST      'LASC-HST
Private mE002 As clsIISACL9000      'ACL9000
Private mE003 As clsIISTEST1        'TEST-1
Private mE004 As clsIISModularDPE   'Modular DP
Private mE005 As clsIISCX3Delta     'CX3 Delta
Private mE006 As clsIISBEPIII       'BEP III
Private mE007 As clsIISBNII         'BN II
Private mE008 As clsIISAxsym        'Axsym
Private mE009 As clsIISElecsys2010  'Elecsys 2010
Private mE010 As clsIISVidas        'Vidas
Private mE011 As clsIISVariantII    'Variant II
Private mE012 As clsIISXE2100       'XE-2100(응급)
Private mE013 As clsIISKX21         'KX-21
Private mE014 As clsIISACL9000      'ACL9000(응급)
Private mE015 As clsIISModularDPE   'Modular P
Private mE016 As clsIISIntegra700   'Integra 700
Private mE017 As clsIISABL555       'ABL 555
Private mE018 As clsIISABL520       'ABL 520
Private mE019 As clsIISABL50        'ABL 50
Private mE020 As clsIISMiditron     'Miditron
Private mE021 As clsIISMiditron     'Miditorn(응급)
Private mE022 As clsIISVitek        'Vitek
Private mE023 As clsIISOCSensorII   'OC-SensorII
Private mE024 As clsIISDimensionRXL 'Dimension RXL
Private mE026 As clsIISIntegra800   'Integra 800
Private mE027 As clsIISABL830       'ABL 800
Private mE028 As clsIISABL835       'ABL 800
Private mE029 As clsIISBEP2000      'BEP 2000
Private mE030 As clsIISiQ200        'iQ200
Private mE031 As clsIISGem3000      'Gem3000

'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "XE-2100"
            If mE001 Is Nothing Then Set mE001 = New clsIISLASCHST
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "ACL9000"
            If mE002 Is Nothing Then Set mE002 = New clsIISACL9000
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "TEST-1"
            If mE003 Is Nothing Then Set mE003 = New clsIISTEST1
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "Modular DP"
            If mE004 Is Nothing Then Set mE004 = New clsIISModularDPE
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "CX3 Delta"
            If mE005 Is Nothing Then Set mE005 = New clsIISCX3Delta
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "BEP III"
            If mE006 Is Nothing Then Set mE006 = New clsIISBEPIII
            Call mE006.ShowForm(pEqpCd, pEqpKey)
        Case "BN II"
            If mE007 Is Nothing Then Set mE007 = New clsIISBNII
            Call mE007.ShowForm(pEqpCd, pEqpKey)
        Case "Axsym"
            If mE008 Is Nothing Then Set mE008 = New clsIISAxsym
            Call mE008.ShowForm(pEqpCd, pEqpKey)
        Case "Elecsys 2010"
            If mE009 Is Nothing Then Set mE009 = New clsIISElecsys2010
            Call mE009.ShowForm(pEqpCd, pEqpKey)
        Case "Vidas"
            If mE010 Is Nothing Then Set mE010 = New clsIISVidas
            Call mE010.ShowForm(pEqpCd, pEqpKey)
        Case "Variant II"
            If mE011 Is Nothing Then Set mE011 = New clsIISVariantII
            Call mE011.ShowForm(pEqpCd, pEqpKey)
        Case "XE-2100(응급)"
            If mE012 Is Nothing Then Set mE012 = New clsIISXE2100
            Call mE012.ShowForm(pEqpCd, pEqpKey)
        Case "KX-21"
            If mE013 Is Nothing Then Set mE013 = New clsIISKX21
            Call mE013.ShowForm(pEqpCd, pEqpKey)
        Case "ACL9000(응급)"
            If mE014 Is Nothing Then Set mE014 = New clsIISACL9000
            Call mE014.ShowForm(pEqpCd, pEqpKey)
        Case "Modular P"
            If mE015 Is Nothing Then Set mE015 = New clsIISModularDPE
            Call mE015.ShowForm(pEqpCd, pEqpKey)
        Case "Integra 700"
            If mE016 Is Nothing Then Set mE016 = New clsIISIntegra700
            Call mE016.ShowForm(pEqpCd, pEqpKey)
        Case "ABL 555"
            If mE017 Is Nothing Then Set mE017 = New clsIISABL555
            Call mE017.ShowForm(pEqpCd, pEqpKey)
        Case "ABL 520"
            If mE018 Is Nothing Then Set mE018 = New clsIISABL520
            Call mE018.ShowForm(pEqpCd, pEqpKey)
        Case "ABL 50"
            If mE019 Is Nothing Then Set mE019 = New clsIISABL50
            Call mE019.ShowForm(pEqpCd, pEqpKey)
        Case "Miditron"
            If mE020 Is Nothing Then Set mE020 = New clsIISMiditron
            Call mE020.ShowForm(pEqpCd, pEqpKey)
        Case "Miditron(응급)"
            If mE021 Is Nothing Then Set mE021 = New clsIISMiditron
            Call mE021.ShowForm(pEqpCd, pEqpKey)
        Case "Vitek"
            If mE022 Is Nothing Then Set mE022 = New clsIISVitek
            Call mE022.ShowForm(pEqpCd, pEqpKey)
        Case "OC-SensorII"
            If mE023 Is Nothing Then Set mE023 = New clsIISOCSensorII
            Call mE023.ShowForm(pEqpCd, pEqpKey)
        Case "Dimension RXL"
            If mE024 Is Nothing Then Set mE024 = New clsIISDimensionRXL
            Call mE024.ShowForm(pEqpCd, pEqpKey)
        Case "Integra 800"
            If mE026 Is Nothing Then Set mE026 = New clsIISIntegra800
            Call mE026.ShowForm(pEqpCd, pEqpKey)
        Case "ABL 830"
            If mE027 Is Nothing Then Set mE027 = New clsIISABL830
            Call mE027.ShowForm(pEqpCd, pEqpKey)
        Case "ABL 835"
            If mE028 Is Nothing Then Set mE028 = New clsIISABL835
            Call mE028.ShowForm(pEqpCd, pEqpKey)
        Case "BEP2000"
            If mE029 Is Nothing Then Set mE029 = New clsIISBEP2000
            Call mE029.ShowForm(pEqpCd, pEqpKey)
        Case "iQ200"
            If mE030 Is Nothing Then Set mE030 = New clsIISiQ200
            Call mE030.ShowForm(pEqpCd, pEqpKey)
        Case "GEM3000"
            If mE031 Is Nothing Then Set mE031 = New clsIISGem3000
            Call mE031.ShowForm(pEqpCd, pEqpKey)
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
    Set mE017 = Nothing
    Set mE018 = Nothing
    Set mE019 = Nothing
    Set mE020 = Nothing
    Set mE021 = Nothing
    Set mE022 = Nothing
    Set mE023 = Nothing
    Set mE024 = Nothing
    Set mE026 = Nothing
    Set mE027 = Nothing
    Set mE028 = Nothing
    Set mE029 = Nothing
    Set mE030 = Nothing
    Set mE031 = Nothing
End Sub
