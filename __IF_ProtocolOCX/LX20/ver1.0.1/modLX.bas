Attribute VB_Name = "modLX"
'
'   LX �迭 ���� ���
'
Option Explicit

Public piAckEtx     As Integer
Public pbContension As Boolean
Public psNakBuf     As String

Type CXINFO
    CURINDEX    As Integer
    BARCODE(4)  As String       'Rack�� Pos - LX:4��, CX:7��
End Type
