Attribute VB_Name = "modCX"
'
'   CX �迭 ���� ���
'
Option Explicit

Public piAckEtx     As Integer
Public pbContension As Boolean
Public psNakBuf     As String

Type CXINFO
    CURINDEX    As Integer
    BARCODE(7)  As String
End Type
