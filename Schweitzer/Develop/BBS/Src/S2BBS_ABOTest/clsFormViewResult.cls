VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormViewResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents frm As frmViewResult
Attribute frm.VB_VarHelpID = -1
Public Event FormClose()

Public Sub FormLoad(parentObject As Object)
    If frm Is Nothing Then Set frm = New frmViewResult
    Call SetParent(frm.hWnd, parentObject.hWnd)
    frm.Show
    frm.ZOrder
End Sub

Private Sub Class_Terminate()
    If Not (frm Is Nothing) Then Unload frm: Set frm = Nothing
End Sub

Private Sub frm_FormClose()
    RaiseEvent FormClose
End Sub
