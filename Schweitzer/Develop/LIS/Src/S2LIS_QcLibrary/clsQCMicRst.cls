VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQCMicRst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarRstCd As String
Private mvarRstString As String
Private mvarRstCnt As Integer

Public Property Get RstCd() As String
    RstCd = mvarRstCd
End Property

Public Property Let RstCd(ByVal mNewValue As String)
    mvarRstCd = mNewValue
End Property

Public Property Get RstString() As String
    RstString = mvarRstString
End Property

Public Property Let RstString(ByVal mNewValue As String)
    mvarRstString = mNewValue
End Property

Public Property Get RstCnt() As Integer
    RstCnt = mvarRstCnt
End Property

Public Property Let RstCnt(ByVal mNewValue As Integer)
    mvarRstCnt = mNewValue
End Property


