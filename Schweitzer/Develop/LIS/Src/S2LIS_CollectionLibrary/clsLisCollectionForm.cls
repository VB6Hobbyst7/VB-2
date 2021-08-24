VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisCollectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents fL154 As frm154NurCol
Attribute fL154.VB_VarHelpID = -1
Private WithEvents fL159 As frm159BatchBarReprint
Attribute fL159.VB_VarHelpID = -1
Private WithEvents fL160 As frm160WardBarReprint
Attribute fL160.VB_VarHelpID = -1
Private WithEvents fL161 As frm161WardCollect
Attribute fL161.VB_VarHelpID = -1
Private WithEvents fL164 As frm164BatchCol
Attribute fL164.VB_VarHelpID = -1
Private WithEvents fL165 As frm165OutCol
Attribute fL165.VB_VarHelpID = -1
Private WithEvents fL166 As frm166OgyCollect
Attribute fL166.VB_VarHelpID = -1

Private WithEvents fL167 As frm167CollectionM
Attribute fL167.VB_VarHelpID = -1

Private WithEvents fL168 As frm168POCTCol
Attribute fL168.VB_VarHelpID = -1

Private WithEvents fL169 As frm265BarPrint
Attribute fL169.VB_VarHelpID = -1

Public Event LastFormUnload()
Public Event ThisFormUnload()

Public Property Let IsDeveloper(ByVal vData As Boolean)
    gIsDeveloper = vData
End Property

Public Property Let EmpId(ByVal vData As String)
    gEmpId = vData
End Property

Public Property Let DeptCd(ByVal vData As String)
    gDeptCd = vData
End Property

Public Property Let WardId(ByVal vData As String)
    gWardId = vData
End Property
Public Property Let WardNm(ByVal vData As String)
    gWardNm = vData
End Property

Public Property Let UsingInWardMenu(ByVal vData As Boolean)
    gUsingInWardMenu = vData
End Property

Private Sub Class_Terminate()
    If Not (fL154 Is Nothing) Then Unload fL154: Set fL154 = Nothing
    If Not (fL159 Is Nothing) Then Unload fL159: Set fL159 = Nothing
    If Not (fL160 Is Nothing) Then Unload fL160: Set fL160 = Nothing
    If Not (fL161 Is Nothing) Then Unload fL161: Set fL161 = Nothing
    If Not (fL164 Is Nothing) Then Unload fL164: Set fL164 = Nothing
    If Not (fL165 Is Nothing) Then Unload fL165: Set fL165 = Nothing
    If Not (fL166 Is Nothing) Then Unload fL166: Set fL166 = Nothing
    If Not (fL167 Is Nothing) Then Unload fL167: Set fL167 = Nothing
    If Not (fL168 Is Nothing) Then Unload fL168: Set fL168 = Nothing
    If Not (fL169 Is Nothing) Then Unload fL169: Set fL169 = Nothing
End Sub

Public Sub CollectionButtonClick(ByVal pKey As String, Optional picForm As Object)
    
    
    On Error GoTo ErrFormLoad
    
    Select Case pKey
        Case "LIS204"
            If fL154 Is Nothing Then Set fL154 = New frm154NurCol
            Call SetParent(fL154.hWnd, picForm.hWnd)
            fL154.WindowState = 2
            fL154.Show
            fL154.ZOrder
        Case "LIS206"
            If fL165 Is Nothing Then Set fL165 = New frm165OutCol
            Call SetParent(fL165.hWnd, picForm.hWnd)
            fL165.WindowState = 2
            fL165.Show
            fL165.ZOrder
        Case "LIS212"
            If fL159 Is Nothing Then Set fL159 = New frm159BatchBarReprint
            Call SetParent(fL159.hWnd, picForm.hWnd)
            fL159.WindowState = 2
            fL159.Show
            fL159.ZOrder
        Case "LIS213"
            If fL160 Is Nothing Then Set fL160 = New frm160WardBarReprint
            Call SetParent(fL160.hWnd, picForm.hWnd)
            fL160.WindowState = 2
            fL160.Show
            fL160.ZOrder
        Case "LIS214"
            If fL161 Is Nothing Then Set fL161 = New frm161WardCollect
            Call SetParent(fL161.hWnd, picForm.hWnd)
            fL161.WindowState = 2
            fL161.Show
            fL161.ZOrder
        Case "LIS217"
            If fL164 Is Nothing Then Set fL164 = New frm164BatchCol
            Call SetParent(fL164.hWnd, picForm.hWnd)
            fL164.WindowState = 2
            fL164.Show
            fL164.ZOrder
        Case "LIS218"
            If fL165 Is Nothing Then Set fL165 = New frm165OutCol
            Call SetParent(fL165.hWnd, picForm.hWnd)
            fL165.WindowState = 2
            fL165.Show
            fL165.ZOrder
        Case "LIS220"
            If fL166 Is Nothing Then Set fL166 = New frm166OgyCollect
            Call SetParent(fL166.hWnd, picForm.hWnd)
            fL166.WindowState = 2
            fL166.Show
            fL166.ZOrder
        Case "LIS221"
            If fL167 Is Nothing Then Set fL167 = New frm167CollectionM
            Call SetParent(fL167.hWnd, picForm.hWnd)
            fL167.WindowState = 2
            fL167.Show
            fL167.ZOrder
        Case "LIS222"
            If fL168 Is Nothing Then Set fL168 = New frm168POCTCol
            Call SetParent(fL168.hWnd, picForm.hWnd)
            fL168.WindowState = 2
            fL168.Show
            fL168.ZOrder
        Case "LIS223"
            If fL169 Is Nothing Then Set fL169 = New frm265BarPrint
            Call SetParent(fL169.hWnd, picForm.hWnd)
            fL169.WindowState = 2
            fL169.Show
            fL169.ZOrder
    End Select
    Exit Sub
    
ErrFormLoad:
    MsgBox Err.Description, vbCritical, "프로그램 로드"
End Sub


Private Sub fL154_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL159_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
'
Private Sub fL160_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL160_ThisFormUnload()
    RaiseEvent ThisFormUnload
End Sub

Private Sub fL161_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL165_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL164_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL166_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
Private Sub fL167_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
Public Sub LoadOutCollection(ByVal Ptid As String, ByVal OrdDt As String)

    If fL165 Is Nothing Then Set fL165 = New frm165OutCol
    
    fL165.Ptid = Ptid:    fL165.OrdDt = OrdDt
    fL165.AccListDisplayer

End Sub
