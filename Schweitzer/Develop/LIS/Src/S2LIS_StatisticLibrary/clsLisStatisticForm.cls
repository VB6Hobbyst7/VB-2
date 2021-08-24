VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisStatisticForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents fL4511   As frm451_N
Attribute fL4511.VB_VarHelpID = -1
Private WithEvents fL451    As frm451AccCnt
Attribute fL451.VB_VarHelpID = -1
Private WithEvents fL452    As frm452TurnAroundTime
Attribute fL452.VB_VarHelpID = -1
'Private WithEvents fL453    As frm453WorkLoad
Private WithEvents fL454    As frm454AbnormalList
Attribute fL454.VB_VarHelpID = -1
Private WithEvents fL455    As frm455AnalysisList
Attribute fL455.VB_VarHelpID = -1
Private WithEvents fL456    As frm456SuscTrand
Attribute fL456.VB_VarHelpID = -1
Private WithEvents fL457    As frm457Epidemic
Attribute fL457.VB_VarHelpID = -1
Private WithEvents fL458    As frm458Infection
Attribute fL458.VB_VarHelpID = -1
Private WithEvents fL459    As frm459MAccCnt
Attribute fL459.VB_VarHelpID = -1
Private WithEvents fL460    As frm460ItemCnt
Attribute fL460.VB_VarHelpID = -1
'Private WithEvents fL461    As frm461BldCultureCnt
Private WithEvents fL462    As frm462CaseStudy
Attribute fL462.VB_VarHelpID = -1
Private WithEvents fL463    As frm463EmmaList
Attribute fL463.VB_VarHelpID = -1
'Private WithEvents fL464    As frm464Infect
'Private WithEvents fL465    As frm465ImageCnt
'Private WithEvents fL466    As frm466WorkUnit
Private WithEvents fL467    As frm467TestTat
Attribute fL467.VB_VarHelpID = -1
Private WithEvents fL468    As frm500MonthTAT
Attribute fL468.VB_VarHelpID = -1
Private WithEvents fL469    As frm501AbList
Attribute fL469.VB_VarHelpID = -1
Public Event LastFormUnload()

Public Property Let IsDeveloper(ByVal vData As Boolean)
    gIsDeveloper = vData
End Property

Public Property Let EmpId(ByVal vData As String)
    gEmpId = vData
End Property

Private Sub Class_Terminate()
    If Not (fL4511 Is Nothing) Then Unload fL4511: Set fL4511 = Nothing
    If Not (fL451 Is Nothing) Then Unload fL451: Set fL451 = Nothing
    If Not (fL452 Is Nothing) Then Unload fL452: Set fL452 = Nothing
'    If Not (fL453 Is Nothing) Then Unload fL453: Set fL453 = Nothing
    If Not (fL454 Is Nothing) Then Unload fL454: Set fL454 = Nothing
    If Not (fL455 Is Nothing) Then Unload fL455: Set fL455 = Nothing
    If Not (fL456 Is Nothing) Then Unload fL456: Set fL456 = Nothing
    If Not (fL457 Is Nothing) Then Unload fL457: Set fL457 = Nothing
    If Not (fL458 Is Nothing) Then Unload fL458: Set fL458 = Nothing
    If Not (fL459 Is Nothing) Then Unload fL459: Set fL459 = Nothing
    If Not (fL460 Is Nothing) Then Unload fL460: Set fL460 = Nothing
'    If Not (fL461 Is Nothing) Then Unload fL461: Set fL461 = Nothing
    If Not (fL462 Is Nothing) Then Unload fL462: Set fL462 = Nothing
    If Not (fL463 Is Nothing) Then Unload fL463: Set fL463 = Nothing
'    If Not (fL464 Is Nothing) Then Unload fL464: Set fL464 = Nothing
'    If Not (fL465 Is Nothing) Then Unload fL465: Set fL465 = Nothing
'    If Not (fL466 Is Nothing) Then Unload fL466: Set fL466 = Nothing
    If Not (fL467 Is Nothing) Then Unload fL467: Set fL467 = Nothing
    If Not (fL468 Is Nothing) Then Unload fL468: Set fL468 = Nothing
End Sub


Public Sub StatisticButtonClick(ByVal pKey As String, Optional picForm As Object)
    
    
    On Error GoTo ErrFormLoad
    
    Select Case pKey
        Case "LIS801"
            If fL4511 Is Nothing Then Set fL4511 = New frm451_N
            Call SetParent(fL4511.hWnd, picForm.hWnd)
            fL4511.WindowState = 2
            fL4511.Show
            fL4511.ZOrder
        Case "LIS802"
            If fL452 Is Nothing Then Set fL452 = New frm452TurnAroundTime
            Call SetParent(fL452.hWnd, picForm.hWnd)
            fL452.WindowState = 2
            fL452.Show
            fL452.ZOrder
        Case "LIS803"
'            If fL464 Is Nothing Then Set fL464 = New frm464Infect
'            Call SetParent(fL464.hWnd, picForm.hWnd)
'            fL464.WindowState = 2
'            fL464.Show
'            fL464.ZOrder

        Case "LIS804"
            If fL454 Is Nothing Then Set fL454 = New frm454AbnormalList
            Call SetParent(fL454.hWnd, picForm.hWnd)
            fL454.WindowState = 2
            fL454.Show
            fL454.ZOrder
        Case "LIS805"
            If fL455 Is Nothing Then Set fL455 = New frm455AnalysisList
            Call SetParent(fL455.hWnd, picForm.hWnd)
            fL455.WindowState = 2
            fL455.Show
            fL455.ZOrder
        Case "LIS806"
            If fL456 Is Nothing Then Set fL456 = New frm456SuscTrand
            Call SetParent(fL456.hWnd, picForm.hWnd)
            fL456.WindowState = 2
            fL456.Show
            fL456.ZOrder
        Case "LIS807"
'            If fL453 Is Nothing Then Set fL453 = New frm453WorkLoad
'            Call SetParent(fL453.hWnd, picForm.hWnd)
'            fL453.WindowState = 2
'            fL453.Show
'            fL453.ZOrder
        Case "LIS808"
            If fL460 Is Nothing Then Set fL460 = New frm460ItemCnt
            Call SetParent(fL460.hWnd, picForm.hWnd)
            fL460.WindowState = 2
            fL460.Show
            fL460.ZOrder
        Case "LIS809"
'            If fL461 Is Nothing Then Set fL461 = New frm461BldCultureCnt
'            Call SetParent(fL461.hWnd, picForm.hWnd)
'            fL461.WindowState = 2
'            fL461.Show
'            fL461.ZOrder
        Case "LIS810"
            If fL459 Is Nothing Then Set fL459 = New frm459MAccCnt
            Call SetParent(fL459.hWnd, picForm.hWnd)
            fL459.WindowState = 2
            fL459.Show
            fL459.ZOrder
        Case "LIS811"
            If fL462 Is Nothing Then Set fL462 = New frm462CaseStudy
            Call SetParent(fL462.hWnd, picForm.hWnd)
            fL462.WindowState = 2
            fL462.Show
            fL462.ZOrder
        Case "LIS812"
'            If fL463 Is Nothing Then Set fL463 = New frm463Statis
'            Call SetParent(fL463.hWnd, picForm.hWnd)
'            fL463.WindowState = 2
'            fL463.Show
'            fL463.ZOrder
' 수정 : 온승호 2011.01.15
' 수정사유 : emma list 추가
            If fL463 Is Nothing Then Set fL463 = New frm463EmmaList
            Call SetParent(fL463.hWnd, picForm.hWnd)
            fL463.WindowState = 2
            fL463.Show
            fL463.ZOrder

        Case "LIS813"
            If fL451 Is Nothing Then Set fL451 = New frm451AccCnt
            Call SetParent(fL451.hWnd, picForm.hWnd)
            fL451.WindowState = 2
            fL451.Show
            fL451.ZOrder
        Case "LIS814"
'            If fL465 Is Nothing Then Set fL465 = New frm465ImageCnt
'            Call SetParent(fL465.hWnd, picForm.hWnd)
'            fL465.WindowState = 2
'            fL465.Show
'            fL465.ZOrder
        Case "LIS815"
'            If fL466 Is Nothing Then Set fL466 = New frm466WorkUnit
'            Call SetParent(fL466.hWnd, picForm.hWnd)
'            fL466.WindowState = 2
'            fL466.Show
'            fL466.ZOrder
        Case "LIS816"
            If fL467 Is Nothing Then Set fL467 = New frm467TestTat
            Call SetParent(fL467.hWnd, picForm.hWnd)
            fL467.WindowState = 2
            fL467.Show
            fL467.ZOrder
        Case "LIS817"
            If fL468 Is Nothing Then Set fL468 = New frm500MonthTAT
            Call SetParent(fL468.hWnd, picForm.hWnd)
            fL468.WindowState = 2
            fL468.Show
            fL468.ZOrder
        Case "LIS818"
            If fL469 Is Nothing Then Set fL469 = New frm501AbList
            Call SetParent(fL469.hWnd, picForm.hWnd)
            fL469.WindowState = 2
            fL469.Show
            fL469.ZOrder
    End Select
    Exit Sub
    
ErrFormLoad:
    MsgBox Err.Description, vbCritical, "프로그램 로드"
End Sub

Private Sub fL4511_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL451_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL452_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL453_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL454_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL455_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL456_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL457_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL458_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL459_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL460_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL461_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL462_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL463_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
Private Sub fL464_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL465_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL466_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL467_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL468_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

