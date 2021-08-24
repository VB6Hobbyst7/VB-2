VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private WithEvents fL157 As frm157BarReprint
Attribute fL157.VB_VarHelpID = -1

Private WithEvents fL410 As frm410PWardColList
Attribute fL410.VB_VarHelpID = -1
Private WithEvents fL411 As frm411PCollectionList
Attribute fL411.VB_VarHelpID = -1
Private WithEvents fL412 As frm412PWorkListG
Attribute fL412.VB_VarHelpID = -1
Private WithEvents fL413 As frm413PWorkListM
Attribute fL413.VB_VarHelpID = -1
Private WithEvents fL414 As frm414PrintEWS
Attribute fL414.VB_VarHelpID = -1
Private WithEvents fL415 As frm415PAccessionList
Attribute fL415.VB_VarHelpID = -1
Private WithEvents fL416 As frm416PUnverifiedList
Attribute fL416.VB_VarHelpID = -1
Private WithEvents fL417 As frm417PModifiedList
Attribute fL417.VB_VarHelpID = -1
Private WithEvents fL418 As frm418BatchReport
Attribute fL418.VB_VarHelpID = -1

Private WithEvents fL419 As frm419PColList
Attribute fL419.VB_VarHelpID = -1
Private WithEvents fL420 As frm420RIAColList
Attribute fL420.VB_VarHelpID = -1
Private WithEvents fL421 As frm421TubercleReport
Attribute fL421.VB_VarHelpID = -1
Private WithEvents fL422 As frm422RiPrint
Attribute fL422.VB_VarHelpID = -1
Private WithEvents fL423 As frm423RICnt
Attribute fL423.VB_VarHelpID = -1

Private WithEvents fL424 As frm424SpcStatics
Attribute fL424.VB_VarHelpID = -1
Private WithEvents fL425 As frm425ModifyCnt
Attribute fL425.VB_VarHelpID = -1
Private WithEvents fL426 As frm426ImageReport
Attribute fL426.VB_VarHelpID = -1

Private WithEvents fL427 As frm427ColList   '외래 채혈 리스트
Attribute fL427.VB_VarHelpID = -1
Private WithEvents fL428 As frm428SpcStatics
Attribute fL428.VB_VarHelpID = -1

Private WithEvents fL160 As frm160WardBarReprint
Attribute fL160.VB_VarHelpID = -1

Private CurNode As Node

Public Event ThisFormClose(ByVal IsLast As Boolean)

Public Property Let IsDeveloper(ByVal vData As Boolean)
    gIsDeveloper = vData
End Property

Public Property Let EmpId(ByVal vData As String)
    gEmpId = vData
End Property

Public Property Let UsingInWardMenu(ByVal vData As Boolean)
    gUsingInWardMenu = vData
End Property
Public Property Let WardID(ByVal vData As String)
    gWardid = vData
End Property
Public Property Let WardNm(ByVal vData As String)
    gWardNm = vData
End Property

Public Sub ReportTreeviewLoad(tvwMenu As Object)

    With tvwMenu
        
        .Nodes.Clear
       
        Call .Nodes.Add(, , "P0", "재발행", 1)
        Call .Nodes.Add("P0", tvwChild, "R001", "Barcode Label 재출력", 2)
'        Call .Nodes.Add("P0", tvwChild, "R002", "건물별 채혈 리스트", 2)
'        Call .Nodes.Add("P0", tvwChild, "R003", "병동별 채혈 리스트", 2)
        Call .Nodes.Add("P0", tvwChild, "R002", "병동별 채혈 리스트", 2)
'        Call .Nodes.Add("P0", tvwChild, "R003", "건물별 채혈 리스트", 2)
        
        Call .Nodes.Add("P0", tvwChild, "R004", "일반검사 Work List", 2)
        Call .Nodes.Add("P0", tvwChild, "R005", "기타검사 Work List", 2)
        Call .Nodes.Add("P0", tvwChild, "R006", "미생물 Work List", 2)
'        Call .Nodes.Add("P0", tvwChild, "R007", "RI 채혈리스트", 2)
        
'        If  = "05" Then    '성바오로병원만 적용
'            Call .Nodes.Add("P0", tvwChild, "R008", "외래채혈 리스트", 2)
'        End If
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("P0", tvwNext, "P1", "일괄 출력", 1)
        Call .Nodes.Add("P1", tvwChild, "R101", "접수 대장", 2)
        Call .Nodes.Add("P1", tvwChild, "R102", "미확인 리스트", 2)
        Call .Nodes.Add("P1", tvwChild, "R103", "수정/Remark입력자 리스트", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("P1", tvwNext, "P2", "결과 보고서", 1)
        Call .Nodes.Add("P2", tvwChild, "R201", "결과보고서 일괄 출력", 2)
        
        If P_ImageSystem = True Then
            Call .Nodes.Add("P2", tvwChild, "R204", "이미지 일괄 출력", 2)
        End If
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
'        Select Case
'            Case "05"
'                Call .Nodes.Add("P2", tvwChild, "R203", "핵의학 검사건수", 2)
'                Call .Nodes.Add("P2", tvwChild, "R200", "핵의학 결과지출력", 2)
                Call .Nodes.Add("P2", tvwChild, "R202", "항산성균 감수성결과 일괄출력", 2)
'            Case "02"
                Call .Nodes.Add("P2", tvwNext, "P3", "통계", 1)
                Call .Nodes.Add("P3", tvwChild, "R301", "검체총건수 통계", 2)
                Call .Nodes.Add("P3", tvwChild, "R303", "상태별건수 통계", 2)
                Call .Nodes.Add("P3", tvwChild, "R302", "수정사유 통계", 2)
'        End Select

        
        
'        Call .Nodes.Add("P2", tvwChild, "R202", "일반검사 결과 보고서", 2)
'        Call .Nodes.Add("P2", tvwChild, "R203", "특수검사 결과 보고서", 2)
'        Call .Nodes.Add("P2", tvwChild, "R204", "미생물 결과 보고서", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
      
        .BorderStyle = vbFixedSingle
    
    End With
   
    tvwMenu.Nodes(1).Selected = True
    Set CurNode = tvwMenu.Nodes(1)

End Sub

Public Sub ReportTreeviewNodeClick(ByVal Rkey As String, ByVal RName As String, Optional ByVal picForm As Object)

On Error GoTo ReportTreeviewNodeClick_error

    Dim strFrmNm As String

    gParentWhnd = picForm.hWnd
    
    Select Case Rkey
        Case "R001":
            strFrmNm = "frm157BarReprint"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL157 Is Nothing Then Set fL157 = New frm157BarReprint: frmCount = frmCount + 1
            Call SetParent(fL157.hWnd, picForm.hWnd)
            fL157.WindowState = 2
            fL157.Show
            fL157.ZOrder
            DoEvents
        Case "R002":
           ' If fL410 Is Nothing Then Set fL410 = New frm410PWardColList: frmCount = frmCount + 1
            strFrmNm = "frm419PColList"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL419 Is Nothing Then Set fL419 = New frm419PColList: frmCount = frmCount + 1
            Call SetParent(fL419.hWnd, picForm.hWnd)
            fL419.WindowState = 2
            fL419.Show
            fL419.ZOrder
            DoEvents
        Case "R003":
            strFrmNm = "frm411PCollectionList"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL411 Is Nothing Then Set fL411 = New frm411PCollectionList: frmCount = frmCount + 1
            Call SetParent(fL411.hWnd, picForm.hWnd)
            fL411.WindowState = 2
            fL411.Show
            fL411.ZOrder
            DoEvents
        Case "R004":
            strFrmNm = "frm412PWorkListG"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL412 Is Nothing Then Set fL412 = New frm412PWorkListG: frmCount = frmCount + 1
            Call SetParent(fL412.hWnd, picForm.hWnd)
            fL412.WindowState = 2
            fL412.Show
            fL412.ZOrder
            DoEvents
        Case "R005":
            strFrmNm = "frm414PrintEWS"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL414 Is Nothing Then Set fL414 = New frm414PrintEWS: frmCount = frmCount + 1
            Call SetParent(fL414.hWnd, picForm.hWnd)
            fL414.WindowState = 2
            fL414.Show
            fL414.ZOrder
            DoEvents
        Case "R006":
            strFrmNm = "frm413PWorkListM"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL413 Is Nothing Then Set fL413 = New frm413PWorkListM: frmCount = frmCount + 1
            Call SetParent(fL413.hWnd, picForm.hWnd)
            fL413.WindowState = 2
            fL413.Show
            fL413.ZOrder
            DoEvents
        
        Case "R007":
            strFrmNm = "frm420RIAColList"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL420 Is Nothing Then Set fL420 = New frm420RIAColList: frmCount = frmCount + 1
            Call SetParent(fL420.hWnd, picForm.hWnd)
            fL420.WindowState = 2
            fL420.Show
            fL420.ZOrder
            DoEvents
        Case "R008"
            strFrmNm = "frm427ColList"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL427 Is Nothing Then Set fL427 = New frm427ColList: frmCount = frmCount + 1
            Call SetParent(fL427.hWnd, picForm.hWnd)
            fL427.WindowState = 2
            fL427.Show
            fL427.ZOrder
            DoEvents
            
        Case "R101":
            strFrmNm = "frm415PAccessionList"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL415 Is Nothing Then Set fL415 = New frm415PAccessionList: frmCount = frmCount + 1
            Call SetParent(fL415.hWnd, picForm.hWnd)
            fL415.WindowState = 2
            fL415.Show
            fL415.ZOrder
            DoEvents
        Case "R102":
            strFrmNm = "frm416PUnverifiedList"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL416 Is Nothing Then Set fL416 = New frm416PUnverifiedList: frmCount = frmCount + 1
            Call SetParent(fL416.hWnd, picForm.hWnd)
            fL416.WindowState = 2
            fL416.Show
            fL416.ZOrder
            DoEvents
        Case "R103":
            strFrmNm = "frm417PModifiedList"
            If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
            If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            If fL417 Is Nothing Then Set fL417 = New frm417PModifiedList: frmCount = frmCount + 1
            Call SetParent(fL417.hWnd, picForm.hWnd)
            fL417.WindowState = 2
            fL417.Show
            fL417.ZOrder
            DoEvents
        Case "R200":
            strFrmNm = "frm422RiPrint"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL422 Is Nothing Then Set fL422 = New frm422RiPrint: frmCount = frmCount + 1
            Call SetParent(fL422.hWnd, picForm.hWnd)
            fL422.WindowState = 2
            fL422.Show
            fL422.ZOrder
            DoEvents
            
            
        Case "R201":
            strFrmNm = "frm418BatchReport"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL418 Is Nothing Then Set fL418 = New frm418BatchReport: frmCount = frmCount + 1
            Call SetParent(fL418.hWnd, picForm.hWnd)
            fL418.WindowState = 2
            fL418.Show
            fL418.ZOrder
            DoEvents
        Case "R202":
            strFrmNm = "frm421TubercleReport"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL421 Is Nothing Then Set fL421 = New frm421TubercleReport: frmCount = frmCount + 1
            Call SetParent(fL421.hWnd, picForm.hWnd)
            fL421.WindowState = 2
            fL421.Show
            fL421.ZOrder
            DoEvents
        
        '핵의학 검사건수 조회
        Case "R203":
            strFrmNm = "frm423RICnt"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL423 Is Nothing Then Set fL423 = New frm423RICnt: frmCount = frmCount + 1
            Call SetParent(fL423.hWnd, picForm.hWnd)
            fL423.WindowState = 2
            fL423.Show
            fL423.ZOrder
            DoEvents
        
        
        '핵의학 검사건수 조회
        Case "R204":
            strFrmNm = "frmImageReport"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL426 Is Nothing Then Set fL426 = New frm426ImageReport: frmCount = frmCount + 1
            Call SetParent(fL426.hWnd, picForm.hWnd)
            fL426.WindowState = 2
            fL426.Show
            fL426.ZOrder
            DoEvents
        
        Case "R301":
        '검체총건수통계
            strFrmNm = "frm424SpcStatics"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL424 Is Nothing Then Set fL424 = New frm424SpcStatics: frmCount = frmCount + 1
            Call SetParent(fL424.hWnd, picForm.hWnd)
            fL424.WindowState = 2
            fL424.Show
            fL424.ZOrder
            DoEvents
        Case "R302":
        '수정사유통계
            strFrmNm = "frm425ModifyCnt"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL425 Is Nothing Then Set fL425 = New frm425ModifyCnt: frmCount = frmCount + 1
            Call SetParent(fL425.hWnd, picForm.hWnd)
            fL425.WindowState = 2
            fL425.Show
            fL425.ZOrder
            DoEvents
            
        Case "R303":
        '검체잔건수통계
            strFrmNm = "frm428SpcStatics"
            If Not gUsingInWardMenu Then
                If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
                If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied
            End If
            If fL428 Is Nothing Then Set fL428 = New frm428SpcStatics: frmCount = frmCount + 1
            Call SetParent(fL428.hWnd, picForm.hWnd)
            fL428.WindowState = 2
            fL428.Show
            fL428.ZOrder
            DoEvents
    End Select
    
    Exit Sub
    
PermissionDenied:
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
    Exit Sub
'
ReportTreeviewNodeClick_error:
    MsgBox Err.Description, vbCritical, "프로그램 로드"
    
End Sub

Private Sub Class_Terminate()
    
    If Not (fL157 Is Nothing) Then Unload fL157: Set fL157 = Nothing
    If Not (fL410 Is Nothing) Then Unload fL410: Set fL410 = Nothing
    If Not (fL411 Is Nothing) Then Unload fL411: Set fL411 = Nothing
    If Not (fL412 Is Nothing) Then Unload fL412: Set fL412 = Nothing
    If Not (fL413 Is Nothing) Then Unload fL413: Set fL413 = Nothing
    
    If Not (fL414 Is Nothing) Then Unload fL414: Set fL414 = Nothing
    If Not (fL415 Is Nothing) Then Unload fL415: Set fL415 = Nothing
    If Not (fL416 Is Nothing) Then Unload fL416: Set fL416 = Nothing
    If Not (fL417 Is Nothing) Then Unload fL417: Set fL417 = Nothing
    If Not (fL418 Is Nothing) Then Unload fL418: Set fL418 = Nothing
    If Not (fL419 Is Nothing) Then Unload fL419: Set fL419 = Nothing
    If Not (fL420 Is Nothing) Then Unload fL420: Set fL420 = Nothing
    If Not (fL421 Is Nothing) Then Unload fL421: Set fL421 = Nothing
    If Not (fL422 Is Nothing) Then Unload fL422: Set fL422 = Nothing
    If Not (fL423 Is Nothing) Then Unload fL423: Set fL423 = Nothing
    If Not (fL424 Is Nothing) Then Unload fL424: Set fL424 = Nothing
    If Not (fL425 Is Nothing) Then Unload fL425: Set fL425 = Nothing
    If Not (fL426 Is Nothing) Then Unload fL426: Set fL426 = Nothing
    If Not (fL428 Is Nothing) Then Unload fL428: Set fL428 = Nothing
End Sub

Private Sub fL157_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL157 Is Nothing) Then Unload fL157: Set fL157 = Nothing
End Sub

Private Sub fL410_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL410 Is Nothing) Then Unload fL410: Set fL410 = Nothing
End Sub

Private Sub fL411_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL411 Is Nothing) Then Unload fL411: Set fL411 = Nothing
End Sub

Private Sub fL412_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL412 Is Nothing) Then Unload fL412: Set fL412 = Nothing
End Sub

Private Sub fL413_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL413 Is Nothing) Then Unload fL413: Set fL413 = Nothing
End Sub

Private Sub fL414_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL414 Is Nothing) Then Unload fL414: Set fL414 = Nothing
End Sub

Private Sub fL415_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL415 Is Nothing) Then Unload fL415: Set fL415 = Nothing
End Sub

Private Sub fL417_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL417 Is Nothing) Then Unload fL417: Set fL417 = Nothing
End Sub

Private Sub fL418_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL418 Is Nothing) Then Unload fL418: Set fL418 = Nothing
End Sub

Private Sub fL419_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL419 Is Nothing) Then Unload fL419: Set fL419 = Nothing
End Sub

Private Sub fL420_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL420 Is Nothing) Then Unload fL420: Set fL420 = Nothing
End Sub

Private Sub fL421_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL421 Is Nothing) Then Unload fL421: Set fL421 = Nothing
End Sub
Private Sub fL422_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL422 Is Nothing) Then Unload fL422: Set fL422 = Nothing
End Sub

Private Sub fL423_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL423 Is Nothing) Then Unload fL423: Set fL423 = Nothing
End Sub
Private Sub fL424_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL424 Is Nothing) Then Unload fL424: Set fL424 = Nothing
End Sub
Private Sub fL425_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL425 Is Nothing) Then Unload fL425: Set fL425 = Nothing
End Sub

Private Sub fL426_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL426 Is Nothing) Then Unload fL426: Set fL426 = Nothing
End Sub
Private Sub fL428_FormClose()
    frmCount = frmCount - 1
    RaiseEvent ThisFormClose(frmCount = 0)
    If Not (fL428 Is Nothing) Then Unload fL428: Set fL428 = Nothing
End Sub

