Attribute VB_Name = "modLX"
'
'   LX 계열 전용 모듈
'
Option Explicit

Public piAckEtx     As Integer
Public pbContension As Boolean
Public psNakBuf     As String

Type CXINFO
    CURINDEX    As Integer
    BARCODE(4)  As String       'Rack당 Pos - LX:4개, CX:7개
End Type
