VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICSCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents fL208    As frm208Infection

Private MySQL As New clsInfectiontSQL

Public Event LastFormUnload()


Public Sub InfectionButtonClick(ByVal pKey As String, Optional picForm As Object)
    
    On Error GoTo ErrFormLoad
    
    Select Case pKey
        Case "LIS208"
'            If fL208 Is Nothing Then Set fL208 = New frm208Infection
'            Call SetParent(fL208.hwnd, picForm.hwnd)
'            fL208.WindowState = 2
'            fL208.Show
'            fL208.ZOrder
    End Select
    
    Exit Sub
    
ErrFormLoad:
    MsgBox Err.Description, vbCritical, "프로그램 로드"
End Sub

Private Sub fL208_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
