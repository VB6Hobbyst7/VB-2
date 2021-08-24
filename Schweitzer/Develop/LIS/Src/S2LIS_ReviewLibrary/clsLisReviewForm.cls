VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents fL401 As frm401ResultView
Attribute fL401.VB_VarHelpID = -1
Private WithEvents fL412 As frm401ResultView_WM
Attribute fL412.VB_VarHelpID = -1
Private WithEvents fL402 As frm402Cumulative
Attribute fL402.VB_VarHelpID = -1
'Private WithEvents fL403 As frm403SelReview
'Private WithEvents fL404 As frm404AllResult
'Private WithEvents fL407 As frm407VerifyList

Private WithEvents fL408 As frm408AccResult
Attribute fL408.VB_VarHelpID = -1
'Private WithEvents fL409 As frm409MedReport
'Private WithEvents fL410 As frm410PastResult

'Private WithEvents fL411 As frm411CumResult_New

Private WithEvents fL514 As frm4NewResultView
Attribute fL514.VB_VarHelpID = -1
' majestic 2018-07-30 시행율/양성율 추가
Private WithEvents fL515 As frm515CVR
Attribute fL515.VB_VarHelpID = -1
Private WithEvents fL516 As frm516Ab
Attribute fL516.VB_VarHelpID = -1

'Append By Legends 2003/09/23
'---
'Private WithEvents fL401N As frm401ResultView_N
'---
Private mvarPTid As String
Private mvarFromDate As String
Private mvarToDate As String
Private mvarDeptCd As String
Private mvarWardId As String
Private mvarEmpId As String

Private objParentForm As Object

Public Event LastFormUnload()
Public Event ThisFormUnload()
Public Event ListSelected()
Public Event MouseMove()

Public Function SetReviewForm(ByVal vData As Object)
    Set objParentForm = vData
End Function

Public Property Let UsingInWardMenu(ByVal vData As Boolean)
    gUsingInWardMenu = vData
End Property

Public Property Let IsDeveloper(ByVal vData As Boolean)
    gIsDeveloper = vData
End Property

Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
    gEmpId = vData
    mvarDeptCd = GetDeptCdByEmp(gEmpId)
    gDeptCd = mvarDeptCd
End Property

Public Property Let ToDate(ByVal vData As String)
    mvarToDate = vData
End Property

Public Property Let FromDate(ByVal vData As String)
    mvarFromDate = vData
End Property

Public Property Let PTid(ByVal vData As String)
    mvarPTid = vData
End Property

Public Property Get PTid() As String
    PTid = gPatientId
End Property
'DeptCd
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
    ObjSysInfo.DeptCd = vData
End Property
'WardId
Public Property Let WardId(ByVal vData As String)
    mvarWardId = vData
   ' ObjSysInfo.WardId = vData
End Property


Private Sub Class_Initialize()
    mvarPTid = ""
End Sub

Private Sub Class_Terminate()
    If Not (fL401 Is Nothing) Then Unload fL401: Set fL401 = Nothing
    If Not (fL402 Is Nothing) Then Unload fL402: Set fL402 = Nothing
'    If Not (fL403 Is Nothing) Then Unload fL403: Set fL403 = Nothing
'    If Not (fL407 Is Nothing) Then Unload fL407: Set fL407 = Nothing
'    If Not (fL404 Is Nothing) Then Unload fL404: Set fL404 = Nothing
    
    If Not (fL408 Is Nothing) Then Unload fL408: Set fL408 = Nothing
'    If Not (fL409 Is Nothing) Then Unload fL409: Set fL409 = Nothing
'    If Not (fL410 Is Nothing) Then Unload fL410: Set fL410 = Nothing
    
'    If Not (fL411 Is Nothing) Then Unload fL411: Set fL411 = Nothing
'    If Not (fL401N Is Nothing) Then Unload fL401N: Set fL401N = Nothing
    If Not (fL514 Is Nothing) Then Unload fL514: Set fL514 = Nothing
   
End Sub

' 2009.01.09 양성현 환자ID 파라메터 추가

Public Sub ReviewButtonClick(ByVal pKey As String, Optional picForm As Object, Optional varPTid As String)
' Public Sub ReviewButtonClick(ByVal pKey As String, Optional picForm As Object)
    
    Dim strTypeCd As String
    
    On Error GoTo ErrFormLoad
    
    'Set objParentForm = picForm
    
    Select Case pKey
' 2009.03.12 양성현 접수를 통해서 올때
        Case "LIS155A"
            If fL401 Is Nothing Then Set fL401 = New frm401ResultView
            Call SetParent(fL401.hwnd, picForm.hwnd)
            If gUsingInWardMenu = False Then
                picForm.WindowState = 2
                fL401.WindowState = 2
            End If
            fL401.DeptCd = mvarDeptCd
            fL401.DeptCd = mvarDeptCd
            fL401.txtPtid = varPTid
            Call fL401.Call_PtId_KeyPress
           
            fL401.Show
            Call fL401.cmdRefresh_Click
            fL401.ZOrder


'' 2009.01.09 양성현 환자ID 파라메터 추가
'            fL401.accPTid (varPTid)
''            fL401.Show
'            fL401.ZOrder
'' 2009.05.20 양성현 결과조회 후 이전환자ID로 돌아가는 현상제거
''            strTypeCd = "A1"
'' 2009.03.12 양성현 접수를 통해서 올때
''            Call SetParameters(strTypeCd, fL401)
''            Set objParentForm = picForm
'' 2009.05.20 양성현 -------------------------

        Case "LIS155B"
            If fL401 Is Nothing Then Set fL401 = New S2LIS_ReviewLib.frm401ResultView
            Call SetParent(fL401.hwnd, picForm.hwnd)
            If gUsingInWardMenu = False Then
                Load (fL401)
                picForm.WindowState = 0
                fL401.WindowState = 0
            End If
            fL401.DeptCd = mvarDeptCd
            fL401.txtPtid = varPTid
            Call fL401.Call_PtId_KeyPress

            fL401.Show
            Call fL401.cmdRefresh_Click
            fL401.ZOrder
            strTypeCd = "A1"
'            Call SetParameters(strTypeCd, fL401)
            Set objParentForm = picForm
        
        Case "LIS501", "APS301"
            If fL401 Is Nothing Then Set fL401 = New frm401ResultView
            Call SetParent(fL401.hwnd, picForm.hwnd)
            If gUsingInWardMenu = False Then
                picForm.WindowState = 2
                fL401.WindowState = 2
            End If
            fL401.DeptCd = mvarDeptCd

            fL401.Show
            fL401.ZOrder
            strTypeCd = "A1"
            Call SetParameters(strTypeCd, fL401)
            Set objParentForm = picForm
        Case "LIS501_N"
            If fL412 Is Nothing Then Set fL412 = New frm401ResultView_WM
            Call SetParent(fL412.hwnd, picForm.hwnd)
            If gUsingInWardMenu = False Then
                picForm.WindowState = 2
                fL412.WindowState = 2
            End If
            fL412.DeptCd = mvarDeptCd

            fL412.Show
            fL412.ZOrder
            strTypeCd = "A1"
            Call SetParameters(strTypeCd, fL412)
            Set objParentForm = picForm
        Case "LIS501N"
'            If fL401N Is Nothing Then Set fL401N = New frm401ResultView_N
'            Call SetParent(fL401N.hWnd, picForm.hWnd)
'            If gUsingInWardMenu = False Then
'                picForm.WindowState = 2
'                fL401N.WindowState = 2
'            End If
'            fL401N.DeptCd = mvarDeptCd
'            fL401N.Show
'            fL401N.ZOrder
'            strTypeCd = "A1"
'            Call SetParameters(strTypeCd, fL401N)
'            Set objParentForm = picForm
        Case "LIS502"
            If fL402 Is Nothing Then Set fL402 = New frm402Cumulative
            Call SetParent(fL402.hwnd, picForm.hwnd)
            If gUsingInWardMenu = False Then
                picForm.WindowState = 2
                fL402.WindowState = 2
                Call SetParameters(strTypeCd, fL402)
            End If
            fL402.DeptCd = mvarDeptCd
            fL402.Show
            fL402.ZOrder
            strTypeCd = "A4"
            Call SetParameters(strTypeCd, fL402)
            Set objParentForm = picForm
        Case "LIS503"
'            If fL403 Is Nothing Then Set fL403 = New frm403SelReview
'            Call SetParent(fL403.hWnd, picForm.hWnd)
'            fL403.WindowState = 2
'            fL403.Show
'            fL403.ZOrder
        Case "LIS504"
'            If fL404 Is Nothing Then Set fL404 = New frm404AllResult
'            Call SetParent(fL404.hWnd, picForm.hWnd)
'            If gUsingInWardMenu = False Then
'                picForm.WindowState = 2
'                fL404.WindowState = 2
'            End If
'
'            fL404.DeptCd = mvarDeptCd
'            fL404.Show
'            fL404.ZOrder
'            strTypeCd = "A1"
'            Call SetParameters(strTypeCd, fL404)
'            Set objParentForm = picForm
        Case "LIS505"
'            If fL407 Is Nothing Then Set fL407 = New frm407VerifyList
'            Call SetParent(fL407.hWnd, picForm.hWnd)
'            fL407.WindowState = 2
'            fL407.DeptCd = mvarDeptCd
'            fL407.Show
'            fL407.ZOrder
        Case "LIS507"
            If fL408 Is Nothing Then Set fL408 = New frm408AccResult
            Call SetParent(fL408.hwnd, picForm.hwnd)
            fL408.WindowState = 2
            fL408.Show
            fL408.ZOrder
        Case "LIS508"
'            If fL409 Is Nothing Then Set fL409 = New frm409MedReport
'            Call SetParent(fL409.hWnd, picForm.hWnd)
'            fL409.WindowState = 2
'            'fL408.DeptCd = mvarDeptCd
'            'fL408.EmpId = mvarEmpId
'            fL409.Show
'            fL409.ZOrder
        '과거결과 조회
        Case "LIS509"
'            If fL410 Is Nothing Then Set fL410 = New frm410PastResult
'            Call SetParent(fL410.hWnd, picForm.hWnd)
'            fL410.WindowState = 2
'            'fL408.DeptCd = mvarDeptCd
'            'fL408.EmpId = mvarEmpId
'            fL410.Show
'            fL410.ZOrder
        Case "LIS510"
'            If fL411 Is Nothing Then Set fL411 = New frm411CumResult_New
'            Call SetParent(fL411.hWnd, picForm.hWnd)
'            fL411.WindowState = 2
'            fL411.Show
        Case "LIS514"
            If fL514 Is Nothing Then Set fL514 = New frm4NewResultView
            Call SetParent(fL514.hwnd, picForm.hwnd)
            fL514.WindowState = 2
            fL514.Show
' majestic 2018-07-30 CVR시행율 추가
        Case "LIS515"
            If fL515 Is Nothing Then Set fL515 = New frm515CVR
            Call SetParent(fL515.hwnd, picForm.hwnd)
            fL515.WindowState = 2
            fL515.Show
' majestic 2018-07-30 양성율 추가
        Case "LIS516"
            If fL516 Is Nothing Then Set fL516 = New frm516Ab
            Call SetParent(fL516.hwnd, picForm.hwnd)
            fL516.WindowState = 2
            fL516.Show
    End Select
    Exit Sub
  
    
ErrFormLoad:
    Screen.MousePointer = vbDefault
    Resume Next
    'MsgBox Err.Description, vbCritical, "프로그램 로드"
End Sub

Public Sub SetParameters(ByVal pType As String, ByRef MyForm As Object)
    
    Screen.MousePointer = vbArrowHourglass
     
    If mvarPTid = "" Then mvarPTid = gPatientId
    
    If pType Like "A*" Then
        If mvarPTid <> "" Then
            
            MyForm.txtPtid.Text = mvarPTid
            Call MyForm.Call_PtId_KeyPress
            MyForm.txtPtid.SetFocus
'           DoEvents
            If Not MyForm.PtFg Then
                Screen.MousePointer = vbDefault
'                mvarPtid = ""
                Exit Sub
            End If
        End If
        
        
        If pType = "A2" Or pType = "A4" Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
'
        If pType = "A1" Then
            If mvarFromDate <> "" Then MyForm.dtpFromDate.Value = Format(mvarFromDate, CS_DateMask): DoEvents
        End If
       
        If mvarToDate <> "" Then
             MyForm.dtpToDate.Value = Format(mvarToDate, CS_DateMask)
                If mvarPTid <> "" Then
                   Call MyForm.Call_ToDate_LostFocus
                End If
        End If
        
        If mvarPTid <> "" Then Call MyForm.Call_ToDate_LostFocus

    
    ElseIf pType = "B1" Then
        MyForm.txtWardID.Text = mvarDeptCd  '병동ID
        DoEvents
        Call MyForm.Call_WardId_KeyPress
        MyForm.dtpToTime.Value = Now
        DoEvents
        Call MyForm.Call_cmdGetOrders_click
    ElseIf pType = "B2" Then
        MyForm.txtWardID.Text = mvarDeptCd  '병동ID
        DoEvents
        Call MyForm.Call_WardId_KeyPress
        MyForm.dtpToTime.Value = GetSystemDate
        DoEvents
        Call MyForm.Call_dtpToTime_Validate
        MyForm.cboColTm.ListIndex = MyForm.cboColTm.ListCount - 1
        DoEvents
        Call MyForm.Call_cmdGetOrders_click
    End If
    Screen.MousePointer = vbDefault
    mvarPTid = ""
    Exit Sub

End Sub

Private Sub fL401_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL401_ThisFormUnload()
    RaiseEvent ThisFormUnload
End Sub

'Private Sub fL401N_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
'
'Private Sub fL401N_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub

Private Sub fL402_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL402_ThisFormUnload()
    RaiseEvent ThisFormUnload
End Sub

'Private Sub fL403_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
'
'Private Sub fL403_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub

'Private Sub fL404_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
'
'Private Sub fL404_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub

Private Sub fL407_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub

Private Sub fL408_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
'Private Sub fL409_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
'Private Sub fL410_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub
'Private Sub fL410_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
'Private Sub fL411_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub
'Private Sub fL411_LastFormUnload()
'    RaiseEvent LastFormUnload
'End Sub
Private Sub fL514_LastFormUnload()
    RaiseEvent LastFormUnload
End Sub
'
'Private Sub fL407_ListSelected(ByVal SelPtId As String, ByVal SelFrDt As String, ByVal SelToDt As String)
'
'    If fL404 Is Nothing Then Set fL404 = New frm404AllResult
'
'    gPatientId = SelPtId
'    objParentForm.Show
'
'    With fL404
'        Call SetParent(.hWnd, objParentForm.hWnd)
'        .WindowState = 2
'        .DeptCd = mvarDeptCd
'        .Show
'        .ZOrder 0
'        Call SetParameters("A1", fL404)
'
'        If .QueryFg Then
'            MsgBox "지금 다른 환자의 결과를 검색하고 있습니다. 잠시 후에 다시 조회하세요", vbInformation, "메세지"
'            Exit Sub
'        End If
'
'        RaiseEvent ListSelected
'        DoEvents
'
'
'        .txtPtid.Text = SelPtId
'        DoEvents
'        .dtpFromDate.Value = SelFrDt
'        .dtpToDate.Value = SelToDt
'        .txtPtid.SetFocus
'        DoEvents
'        SendKeys "{TAB}"
'
''        Call .Call_ToDate_LostFocus
'    End With
'
'End Sub
'
'Private Sub fL407_MouseMove()
'    RaiseEvent MouseMove
'End Sub
'
'Private Sub fL407_ThisFormUnload()
'    RaiseEvent ThisFormUnload
'End Sub

Private Function GetDeptCdByEmp(ByVal strEmpId As String) As String

    Dim objSQL As New clsLISSqlStatement
    Dim objRs As Recordset
    Dim strSQL As String
    
    strSQL = GetSQLDoct(strEmpId)  'objSQL.SqlHIS007CodeList(strEmpId)
    Set objRs = New Recordset
    objRs.Open strSQL, DBConn
    
    If Not objRs.EOF Then
        GetDeptCdByEmp = objRs.Fields("deptcd").Value & ""
    Else
        GetDeptCdByEmp = ""
    End If
    objRs.Close
    Set objRs = Nothing
    
End Function
