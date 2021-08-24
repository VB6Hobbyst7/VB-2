VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsMICIISDrug.cls
'   작성자  : 오세원
'   내  용  : Vitek 장비 항생제결과 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mDrugCd As String   '항생제코드
Private mDrugNm As String   '항생제명
Private mVolumn As String   '함량
Private mRstCd  As String   '결과코드

Public Property Get DrugCd() As String
    DrugCd = mDrugCd
End Property

Public Property Let DrugCd(ByVal vData As String)
    mDrugCd = vData
End Property

Public Property Get DrugNm() As String
    DrugNm = mDrugNm
End Property

Public Property Let DrugNm(ByVal vData As String)
    mDrugNm = vData
End Property

Public Property Get Volumn() As String
    Volumn = mVolumn
End Property

Public Property Let Volumn(ByVal vData As String)
    mVolumn = vData
End Property

Public Property Get RstCd() As String
    RstCd = mRstCd
End Property

Public Property Let RstCd(ByVal vData As String)
    mRstCd = vData
End Property

