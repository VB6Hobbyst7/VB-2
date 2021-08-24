Attribute VB_Name = "modCX"
'
'   CX 계열 전용 모듈
'
Option Explicit

Public piAckEtx     As Integer
Public pbContension As Boolean
Public psNakBuf     As String

Type CXINFO
    CURINDEX    As Integer
    BARCODE(7)  As String
End Type
