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
'   파일명  : clsIISInstrShow.cls (경주동국대병원)
'   작성자  : 이상대
'   내  용  : 해당 장비폼을 표시하는 클래스
'   작성일  : 2004-05-28
'   버  전  :
'       1. 1.1.0: 이상대(2005-01-27)
'          - BN II장비 추가
'       2. 1.2.0: 이상대(2005-02-23)
'          - ACL 9000 장비추가
'-----------------------------------------------------------------------------'

Option Explicit

Private mE001 As clsIISABL555       'ABL555
Private mE002 As clsIISUriscanPro   'Uriscan Pro+
Private mE003 As clsIISAdvia120     'ADVIA 120-1
Private mE004 As clsIISAU1000       'AU1000
Private mE005 As clsIISRapidLab850  'RapidLab 850
Private mE006 As clsIISDimensionAR  'Dimension AR
Private mE007 As clsIISDimensionRXL 'Dimension RXL
Private mE008 As clsIISElecsys1010  'Elecsys 1010
Private mE009 As clsIISBN100        'BN100
Private mE010 As clsIISAdvia120     'ADVIA 120-2
Private mE011 As clsIISCentaur      'ADVIA Centaur
Private mE012 As clsIISElecsys2010  'Elecsys 2010
Private mE013 As clsIISVidas        'VIDAS
Private mE014 As clsIISACL6000      'ACL6000
Private mE015 As clsIISACL100       'ACL100
Private mE016 As clsIISVitek        'Vitek
Private mE017 As clsIISBNII         'BN II
Private mE018 As clsIISACL9000      'ACL9000
Private mE019 As clsIISABL835       'ABL835
Private mE026 As clsIISExicycler    'Exicycler
Private mE801 As clsIISTEST1        'Test1

'------------------------------------------------------------------------------'
'   기능 : 해당장비폼을 표시
'   인수 :
'       1.pEqpCd  : 장비코드
'       2.pEqpKey : 장비키
'------------------------------------------------------------------------------'
Public Sub InstrShow(ByVal pEqpCd As String, ByVal pEqpKey As String)
    Select Case pEqpKey
        Case "ABL555"
            If mE001 Is Nothing Then Set mE001 = New clsIISABL555
            Call mE001.ShowForm(pEqpCd, pEqpKey)
        Case "Uriscan Pro+"
            If mE002 Is Nothing Then Set mE002 = New clsIISUriscanPro
            Call mE002.ShowForm(pEqpCd, pEqpKey)
        Case "ADVIA 120-1"
            If mE003 Is Nothing Then Set mE003 = New clsIISAdvia120
            Call mE003.ShowForm(pEqpCd, pEqpKey)
        Case "AU1000"
            If mE004 Is Nothing Then Set mE004 = New clsIISAU1000
            Call mE004.ShowForm(pEqpCd, pEqpKey)
        Case "RapidLab 850"
            If mE005 Is Nothing Then Set mE005 = New clsIISRapidLab850
            Call mE005.ShowForm(pEqpCd, pEqpKey)
        Case "Dimension AR"
            If mE006 Is Nothing Then Set mE006 = New clsIISDimensionAR
            Call mE006.ShowForm(pEqpCd, pEqpKey)
        Case "Dimension RXL"
            If mE007 Is Nothing Then Set mE007 = New clsIISDimensionRXL
            Call mE007.ShowForm(pEqpCd, pEqpKey)
        Case "Elecsys 1010"
            If mE008 Is Nothing Then Set mE008 = New clsIISElecsys1010
            Call mE008.ShowForm(pEqpCd, pEqpKey)
        Case "BN100"
            If mE009 Is Nothing Then Set mE009 = New clsIISBN100
            Call mE009.ShowForm(pEqpCd, pEqpKey)
        Case "ADVIA 120-2"
            If mE010 Is Nothing Then Set mE010 = New clsIISAdvia120
            Call mE010.ShowForm(pEqpCd, pEqpKey)
        Case "ADVIA Centaur"
            If mE011 Is Nothing Then Set mE011 = New clsIISCentaur
            Call mE011.ShowForm(pEqpCd, pEqpKey)
        Case "Elecsys 2010"
            If mE012 Is Nothing Then Set mE012 = New clsIISElecsys2010
            Call mE012.ShowForm(pEqpCd, pEqpKey)
        Case "VIDAS"
            If mE013 Is Nothing Then Set mE013 = New clsIISVidas
            Call mE013.ShowForm(pEqpCd, pEqpKey)
        Case "ACL6000"
            If mE014 Is Nothing Then Set mE014 = New clsIISACL6000
            Call mE014.ShowForm(pEqpCd, pEqpKey)
        Case "ACL100"
            If mE015 Is Nothing Then Set mE015 = New clsIISACL100
            Call mE015.ShowForm(pEqpCd, pEqpKey)
        Case "Vitek"
            If mE016 Is Nothing Then Set mE016 = New clsIISVitek
            Call mE016.ShowForm(pEqpCd, pEqpKey)
        Case "BN II"
            If mE017 Is Nothing Then Set mE017 = New clsIISBNII
            Call mE017.ShowForm(pEqpCd, pEqpKey)
        Case "ACL9000"
            If mE018 Is Nothing Then Set mE018 = New clsIISACL9000
            Call mE018.ShowForm(pEqpCd, pEqpKey)
        Case "ABL835"
            If mE019 Is Nothing Then Set mE019 = New clsIISABL835
            Call mE019.ShowForm(pEqpCd, pEqpKey)
        Case "Exicycler"
            If mE026 Is Nothing Then Set mE026 = New clsIISExicycler
            Call mE026.ShowForm(pEqpCd, pEqpKey)
        Case "TEST1"
            If mE801 Is Nothing Then Set mE801 = New clsIISTEST1
            Call mE801.ShowForm(pEqpCd, pEqpKey)
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
    Set mE026 = Nothing
    Set mE801 = Nothing
End Sub

