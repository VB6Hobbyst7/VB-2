VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mErrCnt      As Long         '알람건수
Private mErrSeq      As String       '알람코드
Private mErrDesc     As String       '알람내용


Public Property Get Count() As String
    Count = mErrCnt
End Property

Public Property Let Count(ByVal vData As String)
    mErrCnt = vData
End Property

Public Property Get SeqNo() As String
    SeqNo = mErrSeq
End Property

Public Property Let SeqNo(ByVal vData As String)
    mErrSeq = vData
End Property

Public Property Get Desc() As String
    Desc = mErrDesc
End Property

Public Property Let Desc(ByVal vData As String)
    mErrDesc = vData
End Property



Private Sub Class_Initialize()
'    Set mIntResults = New clsIISIntResults
'    Set mDrugs = New clsIISMICDrugs
End Sub

Private Sub Class_Terminate()
'    Set mDrugs = Nothing
'    Set mIntResults = Nothing
End Sub




