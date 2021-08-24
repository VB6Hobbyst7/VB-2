VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRstHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarTestCd As String
Private mvarVerifyDt As String
Private mvarVerifyTm As String
Private mvarResult As Double

Public Property Get TestCd() As String
    TestCd = mvarTestCd
End Property
Public Property Let TestCd(ByVal vNewValue As String)
    mvarTestCd = vNewValue
End Property

Public Property Get VerifyDt() As String
    VerifyDt = mvarVerifyDt
End Property
Public Property Let VerifyDt(ByVal vNewValue As String)
    mvarVerifyDt = vNewValue
End Property

Public Property Get VerifyTm() As String
    VerifyTm = mvarVerifyTm
End Property
Public Property Let VerifyTm(ByVal vNewValue As String)
    mvarVerifyTm = vNewValue
End Property

Public Property Get Result() As Double
    Result = mvarResult
End Property
Public Property Let Result(ByVal vNewValue As Double)
    mvarResult = vNewValue
End Property


