VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisQCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'QC New Version ¿ë
Private WithEvents fQC01 As frm3011QCControlMaster_N
Attribute fQC01.VB_VarHelpID = -1
Private WithEvents fQC02 As frm301QCMaster_N
Attribute fQC02.VB_VarHelpID = -1
Private WithEvents fQC03 As frm312QCSchedule_N
Attribute fQC03.VB_VarHelpID = -1
Private WithEvents fQC04 As frm310QCReprint_N
Attribute fQC04.VB_VarHelpID = -1
Private WithEvents fQC05 As frm309QCOrder_N
Attribute fQC05.VB_VarHelpID = -1
Private WithEvents fQC06 As frm311QCResultEntry_N
Attribute fQC06.VB_VarHelpID = -1
Private WithEvents fQC07 As frm302QCReview_N
Attribute fQC07.VB_VarHelpID = -1
'Private WithEvents fQC08 As frm330Calculation_N
'Private WithEvents fQC10 As frm303QCCalibration_N
'Private WithEvents fQC11 As frm305QCRefrigerator_N
'Private WithEvents fQC12 As frm331EquipHistory
Private WithEvents fQC13 As frm302QCReview_N_ALL
Attribute fQC13.VB_VarHelpID = -1

Public Event LastFormUnload()

Public Property Let IsDeveloper(ByVal vData As Boolean)
    gIsDeveloper = vData
End Property

Public Property Let EmpId(ByVal vData As String)
    gEmpId = vData
End Property

Private Sub Class_Terminate()
'QC New Version ¿ë
    If Not (fQC01 Is Nothing) Then Unload fQC01: Set fQC01 = Nothing
    If Not (fQC02 Is Nothing) Then Unload fQC02: Set fQC02 = Nothing
    If Not (fQC03 Is Nothing) Then Unload fQC03: Set fQC03 = Nothing
    If Not (fQC04 Is Nothing) Then Unload fQC04: Set fQC04 = Nothing
    If Not (fQC05 Is Nothing) Then Unload fQC05: Set fQC05 = Nothing
    If Not (fQC06 Is Nothing) Then Unload fQC06: Set fQC06 = Nothing
    If Not (fQC07 Is Nothing) Then Unload fQC07: Set fQC07 = Nothing
'    If Not (fQC08 Is Nothing) Then Unload fQC08: Set fQC08 = Nothing
'    If Not (fQC10 Is Nothing) Then Unload fQC10: Set fQC10 = Nothing
'    If Not (fQC11 Is Nothing) Then Unload fQC11: Set fQC11 = Nothing
'    If Not (fQC12 Is Nothing) Then Unload fQC12: Set fQC12 = Nothing
    If Not (fQC13 Is Nothing) Then Unload fQC13: Set fQC13 = Nothing
End Sub

Public Sub QCButtonClick(ByVal pKey As String, Optional picForm As Object)
    Dim frm As Form
    Dim blnExist As Boolean
    
    On Error GoTo ErrFormLoad
    
    Select Case pKey
'QC NEW VERSION ¿ë...
'---------------------------------------------------------------------------
        Case "QC01"
            If fQC01 Is Nothing Then Set fQC01 = frm3011QCControlMaster_N
            Call SetParent(fQC01.hwnd, picForm.hwnd)
            fQC01.WindowState = 2
            fQC01.Show
            fQC01.ZOrder
        Case "QC02"
            If fQC02 Is Nothing Then Set fQC02 = frm301QCMaster_N
            Call SetParent(fQC02.hwnd, picForm.hwnd)
            fQC02.WindowState = 2
            fQC02.Show
            fQC02.ZOrder
        Case "QC03"
            If fQC03 Is Nothing Then Set fQC03 = frm312QCSchedule_N
            Call SetParent(fQC03.hwnd, picForm.hwnd)
            fQC03.WindowState = 2
            fQC03.Show
            fQC03.ZOrder
        Case "QC04"
            If fQC04 Is Nothing Then Set fQC04 = frm310QCReprint_N
            Call SetParent(fQC04.hwnd, picForm.hwnd)
            fQC04.WindowState = 2
            fQC04.Show
            fQC04.ZOrder
        Case "QC05"
            If fQC05 Is Nothing Then Set fQC05 = frm309QCOrder_N
            Call SetParent(fQC05.hwnd, picForm.hwnd)
            fQC05.WindowState = 2
            fQC05.Show
            fQC05.ZOrder
        Case "QC06"
            If fQC06 Is Nothing Then Set fQC06 = frm311QCResultEntry_N
            Call SetParent(fQC06.hwnd, picForm.hwnd)
            fQC06.WindowState = 2
            fQC06.Show
            fQC06.ZOrder
        Case "QC07"
            If fQC07 Is Nothing Then Set fQC07 = frm302QCReview_N
            Call SetParent(fQC07.hwnd, picForm.hwnd)
            fQC07.WindowState = 2
            fQC07.Show
            fQC07.ZOrder
        Case "QC08"
'            If fQC08 Is Nothing Then Set fQC08 = frm330Calculation_N
'            Call SetParent(fQC08.hwnd, picForm.hwnd)
'            fQC08.WindowState = 2
'            fQC08.Show
'            fQC08.ZOrder
        Case "QC10"
'            If fQC10 Is Nothing Then Set fQC10 = frm303QCCalibration_N
'            Call SetParent(fQC10.hwnd, picForm.hwnd)
'            fQC10.WindowState = 2
'            fQC10.Show
'            fQC10.ZOrder
        Case "QC11"
'            If fQC11 Is Nothing Then Set fQC11 = frm305QCRefrigerator_N
'            Call SetParent(fQC11.hwnd, picForm.hwnd)
'            fQC11.WindowState = 2
'            fQC11.Show
'            fQC11.ZOrder
        Case "QC12"
'            If fQC12 Is Nothing Then Set fQC12 = frm331EquipHistory
'            Call SetParent(fQC12.hwnd, picForm.hwnd)
'            fQC12.WindowState = 2
'            fQC12.Show
'            fQC12.ZOrder
        Case "QC13"
            If fQC13 Is Nothing Then Set fQC13 = frm302QCReview_N_ALL
            Call SetParent(fQC13.hwnd, picForm.hwnd)
            fQC13.WindowState = 2
            fQC13.Show
            fQC13.ZOrder
'--------------------------------------------------------------------
    End Select
    Exit Sub
    
ErrFormLoad:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub fQC01_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC02_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC03_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC04_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC05_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC06_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC07_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC08_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC09_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC10_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC11_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC12_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fQC13_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
