VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICWSBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMICWSBody.cls

'   내  용  : 미생물 WorkSheet Body 클래스
'   작성일  : 2004-09-15
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mWsCd       As String       'WorkSheet Code
Private mWsUnit     As String       'WorkSheet Unit
Private mWorkarea   As String       'Workarea
Private mAccDt      As String       '접수일자
Private mAccSeq     As Long         '접수순번
Private mSelFg      As String       '? (1:WorkSheet출력, 2:최종)
Private mGenFg      As String       'Growth유무(1:No Growth, 2:Growth)
Private mScFg       As String       '결과유형(S:일반감수성, M:MIC감수성, G:Gram Stain, M:AFB Culture, F:Fungus Culture, N:Fungus Stain, B:AFB Stain, A:AFB Sensitivity)
Private mRemark     As String       'Remark

Public Property Get WsCd() As String
    WsCd = mWsCd
End Property

Public Property Let WsCd(ByVal vData As String)
    mWsCd = vData
End Property

Public Property Get WsUnit() As String
    WsUnit = mWsUnit
End Property

Public Property Let WsUnit(ByVal vData As String)
    mWsUnit = vData
End Property

Public Property Get Workarea() As String
    Workarea = mWorkarea
End Property

Public Property Let Workarea(ByVal vData As String)
    mWorkarea = vData
End Property

Public Property Get AccDt() As String
    AccDt = mAccDt
End Property

Public Property Let AccDt(ByVal vData As String)
    mAccDt = vData
End Property

Public Property Get AccSeq() As Long
    AccSeq = mAccSeq
End Property

Public Property Let AccSeq(ByVal vData As Long)
    mAccSeq = vData
End Property

Public Property Get SelFg() As String
    SelFg = mSelFg
End Property

Public Property Let SelFg(ByVal vData As String)
    mSelFg = vData
End Property

Public Property Get GenFg() As String
    GenFg = mGenFg
End Property

Public Property Let GenFg(ByVal vData As String)
    mGenFg = vData
End Property

Public Property Get ScFg() As String
    ScFg = mScFg
End Property

Public Property Let ScFg(ByVal vData As String)
    mScFg = vData
End Property

Public Property Get Remark() As String
    Remark = mRemark
End Property

Public Property Let Remark(ByVal vData As String)
    mRemark = vData
End Property
