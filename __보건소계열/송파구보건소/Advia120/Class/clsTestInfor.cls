VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private AdoCn               As ADODB.Connection
Private AdoRS               As ADODB.Recordset
Private Call_Procedure      As String

Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

Private Sub Class_Initialize()
    Call DBErrClear
End Sub

Private Sub Class_Terminate()
    Set AdoCn = Nothing
End Sub

