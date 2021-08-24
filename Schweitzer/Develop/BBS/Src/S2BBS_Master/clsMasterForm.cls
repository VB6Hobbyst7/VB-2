VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMasterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private fB800 As frmBBS800
Private fB801 As frmBBS801
Private fB802 As frmBBS802
Private fB803 As frmBBS803
Private fB804 As frmBBSProject
Private fB811 As frmBBS811
Private fB812 As frmBBS812
Private fB813 As frmBBS813
Private fB814 As frmBBS814
Private fB815 As frmBBS815
Private fB816 As frmBBS816
Private fB817 As frmBBS817
Private fB818 As frmBBS818
Private fB819 As frmBBS819
Private fB820 As frmBBS820
Private fB821 As frmBBS821
Private fB822 As frmBBS822
Private fB823 As frmBBS823
Private fB824 As frmBBS824
'Private fB825 As frmBBS825
'Private fB826 As frmBBS826
'Private fB827 As frmBBS827
Private fB828 As frmBBS828

Private fB861B As frmBBS861
Private fB861C As frmBBS861
Private fB861D As frmBBS861
Private fB861E As frmBBS861
Private fB861F As frmBBS861
Private fB861G As frmBBS861
Private fB861H As frmBBS861
Private fB861I As frmBBS861
Private fB861J As frmBBS861
Private fB861K As frmBBS861
Private fB861L As frmBBS861
Private fB861M As frmBBS861
Private fB861N As frmBBS861
Private fB861O As frmBBS861
Private fB861P As frmBBS861
Private fB861Q As frmBBS861
Private fB861R As frmBBS861
Private fB861S As frmBBS861

Public Sub MasterTreeviewLoad(tvwMenu As Object)
    Dim ii As Integer
    ii = 1
    
    With tvwMenu
        .Nodes.Clear
        If ObjMyUser.IsDeveloper Then
            Call .Nodes.Add(, , "M0", "개발자메뉴", 1)
            Call .Nodes.Add("M0", tvwChild, "B804", "프로젝트 옵션설정", 2)
            Call .Nodes.Add("M0", tvwChild, "B800", "개발자관리 TOOL", 2) 'frmCTControl
            Call .Nodes.Add("M0", tvwChild, "B801", "임시환자번호 범위(DONOR)", 2)
            Call .Nodes.Add("M0", tvwChild, "B802", "검사건물 설정", 2)
'            Call .Nodes.Add("M0", tvwChild, "B803", "수가계산내역 조회", 2)
            Call .Nodes.Add("M0", tvwNext, "M1", "혈액은행 마스터", 1)
        Else
            Call .Nodes.Add(, , "M1", "혈액은행 마스터", 1)
        End If
        
        Call .Nodes.Add("M" & ii, tvwChild, "M" & ii + 1, "수혈처방", 1)
        ii = ii + 1
        Call .Nodes.Add("M" & ii, tvwChild, "B811", "수혈처방 마스터", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B812", "X-Matching Step 관리", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B819", "수혈전(후)검사항목", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861I", "수혈사유 마스터", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B823", "접수기준(년,월,일)", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861L", "접수취소사유", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861J", "Filter 수가코드", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("M" & ii, tvwNext, "M" & ii + 1, "수혈검체", 1)
        ii = ii + 1
        Call .Nodes.Add("M" & ii, tvwChild, "B814", "검체보관일수 관리", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B815", "검체보관장소 마스터", 2)
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("M" & ii, tvwNext, "M" & ii + 1, "혈액관리(입출고,반환,폐기등)", 1)
        ii = ii + 1
        Call .Nodes.Add("M" & ii, tvwChild, "B861G", "혈액원 마스터", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B822", "혈액제제 마스터 ", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861M", "혈액제제 통계그룹 ", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B828", "XM결과 필수입력 제제", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B813", "혈액반환가능시간", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861H", "혈액폐기사유 마스터", 2)
        
        Call .Nodes.Add("M" & ii, tvwChild, "B861P", "수혈부작용 마스터 ", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861O", "혈액원 제제관리", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861Q", "혈액원 혈액형관리", 2)
        Call .Nodes.Add("M" & ii, tvwChild, "B861R", "혈액원 용량관리", 2)
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        If DonorUserFg = True Then
            Call .Nodes.Add("M" & ii, tvwNext, "M" & ii + 1, "헌혈자", 1)
            ii = ii + 1
            Call .Nodes.Add("M" & ii, tvwChild, "B816", "헌혈자 검사 항목 마스터", 2)
            Call .Nodes.Add("M" & ii, tvwChild, "B818", "문진내역 마스터", 2)
            Call .Nodes.Add("M" & ii, tvwChild, "B861F", "직업 마스터", 2)
            Call .Nodes.Add("M" & ii, tvwChild, "B861N", "부적격 사유", 2)
            tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        End If
        
        
        Call .Nodes.Add("M" & ii, tvwNext, "M" & ii + 1, "ABO검사", 1)
        ii = ii + 1
        Call .Nodes.Add("M" & ii, tvwChild, "B821", "ABO검사항목", 2)
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        If DonorUserFg = True Then
            Call .Nodes.Add("M" & ii, tvwNext, "M" & ii + 1, "출력양식", 1)
            ii = ii + 1
            Call .Nodes.Add("M" & ii, tvwChild, "B824", "헌혈실적 통지서", 2)
'            Call .Nodes.Add("M" & ii, tvwChild, "B825", "부적격 혈액이송", 2)
'            Call .Nodes.Add("M" & ii, tvwChild, "B861S", "혈액바코드양식", 2)
            tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        End If
        .BorderStyle = vbFixedSingle
    End With
End Sub

Public Sub MasterTreeviewNodeClick(ByVal key As String, Optional ByVal picForm As Object)

On Error GoTo MasterTreeviewNodeClick_error

    Select Case key
        Case "B800":
            If fB800 Is Nothing Then Set fB800 = New frmBBS800
            Call SetParent(fB800.hWnd, picForm.hWnd)
            fB800.Show
            fB800.ZOrder
        Case "B801":
            If fB801 Is Nothing Then Set fB801 = New frmBBS801
            Call SetParent(fB801.hWnd, picForm.hWnd)
            fB801.Show
            fB801.ZOrder
        Case "B802":
            If fB802 Is Nothing Then Set fB802 = New frmBBS802
            Call SetParent(fB802.hWnd, picForm.hWnd)
            fB802.Show
            fB802.ZOrder
        Case "B803":
            If fB803 Is Nothing Then Set fB803 = New frmBBS803
            Call SetParent(fB803.hWnd, picForm.hWnd)
            fB803.Show
            fB803.ZOrder
        Case "B804":
'            Dim strTmp As String
''            If fB804 Is Nothing Then
'                strTmp = InputBox("개발자 암호를 입력하세요", "사용자 옵션설정")
'
'                If strTmp <> "pomis0085" Then
'                    MsgBox "개발자 Password가 일치하지 않으므로 사용하실수 없습니다.", vbOKOnly + vbCritical
'                    Exit Sub
'                End If
'            Else
'            End If
            
            If fB804 Is Nothing Then Set fB804 = New frmBBSProject
            Call SetParent(fB804.hWnd, picForm.hWnd)
            fB804.Show
            fB804.ZOrder
        
        Case "B811":
            If fB811 Is Nothing Then Set fB811 = New frmBBS811
            Call SetParent(fB811.hWnd, picForm.hWnd)
            fB811.Show
            fB811.ZOrder
        Case "B812":
            If fB812 Is Nothing Then Set fB812 = New frmBBS812
            Call SetParent(fB812.hWnd, picForm.hWnd)
            fB812.Show
            fB812.ZOrder
        Case "B813":
            If fB813 Is Nothing Then Set fB813 = New frmBBS813
            Call SetParent(fB813.hWnd, picForm.hWnd)
            fB813.Show
            fB813.ZOrder
        Case "B814":
            If fB814 Is Nothing Then Set fB814 = New frmBBS814
            Call SetParent(fB814.hWnd, picForm.hWnd)
            fB814.Show
            fB814.ZOrder
        Case "B815":
            If fB815 Is Nothing Then Set fB815 = New frmBBS815
            Call SetParent(fB815.hWnd, picForm.hWnd)
            fB815.Show
            fB815.ZOrder
        Case "B816":
            If fB816 Is Nothing Then Set fB816 = New frmBBS816
            Call SetParent(fB816.hWnd, picForm.hWnd)
            fB816.Show
            fB816.ZOrder
        Case "B817":
            If fB817 Is Nothing Then Set fB817 = New frmBBS817
            Call SetParent(fB817.hWnd, picForm.hWnd)
            fB817.Show
            fB817.ZOrder
        Case "B818":
            If fB818 Is Nothing Then Set fB818 = New frmBBS818
            Call SetParent(fB818.hWnd, picForm.hWnd)
            fB818.Show
            fB818.ZOrder
        Case "B819":
            If fB819 Is Nothing Then Set fB819 = New frmBBS819
            Call SetParent(fB819.hWnd, picForm.hWnd)
            fB819.Show
            fB819.ZOrder
        Case "B820":
            If fB820 Is Nothing Then Set fB820 = New frmBBS820
            Call SetParent(fB820.hWnd, picForm.hWnd)
            fB820.Show
            fB820.ZOrder
        Case "B821":
            If fB821 Is Nothing Then Set fB821 = New frmBBS821
            Call SetParent(fB821.hWnd, picForm.hWnd)
            fB821.Show
            fB821.ZOrder
        Case "B822":
            If fB822 Is Nothing Then Set fB822 = New frmBBS822
            Call SetParent(fB822.hWnd, picForm.hWnd)
            fB822.Show
            fB822.ZOrder
        Case "B823":
            If fB823 Is Nothing Then Set fB823 = New frmBBS823
            Call SetParent(fB823.hWnd, picForm.hWnd)
            fB823.Show
            fB823.ZOrder
        Case "B824":
            If fB824 Is Nothing Then Set fB824 = New frmBBS824
            Call SetParent(fB824.hWnd, picForm.hWnd)
            fB824.Show
            fB824.ZOrder
'        Case "B825":
'            If fB825 Is Nothing Then Set fB825 = New frmBBS825
'            Call SetParent(fB825.hWnd, picForm.hWnd)
'            fB825.Show
'            fB825.ZOrder
'        Case "B826":
'            If fB826 Is Nothing Then Set fB826 = New frmBBS826
'            Call SetParent(fB826.hWnd, picForm.hWnd)
'            fB826.Show
'            fB826.ZOrder
'        Case "B827":
'            If fB827 Is Nothing Then Set fB827 = New frmBBS827
'            Call SetParent(fB827.hWnd, picForm.hWnd)
'            fB827.Show
'            fB827.ZOrder
        Case "B828":
            If fB828 Is Nothing Then Set fB828 = New frmBBS828
            Call SetParent(fB828.hWnd, picForm.hWnd)
            fB828.Show
            fB828.ZOrder
        Case "B861B":
            If fB861B Is Nothing Then Set fB861B = New frmBBS861
            Call SetParent(fB861B.hWnd, picForm.hWnd)
            fB861B.CDINDEX = BC2_COMPO
            fB861B.Show
            fB861B.ZOrder
        Case "B861C":
            If fB861C Is Nothing Then Set fB861C = New frmBBS861
            Call SetParent(fB861C.hWnd, picForm.hWnd)
            fB861C.CDINDEX = BC2_BRANCH
            fB861C.Show:
            fB861C.ZOrder
        Case "B861D":
            If fB861D Is Nothing Then Set fB861D = New frmBBS861
            Call SetParent(fB861D.hWnd, picForm.hWnd)
            fB861D.CDINDEX = BC2_LOCAL
            fB861D.Show
            fB861D.ZOrder
        Case "B861E":
            If fB861E Is Nothing Then Set fB861E = New frmBBS861
            Call SetParent(fB861E.hWnd, picForm.hWnd)
            fB861E.CDINDEX = BC2_KIT
            fB861E.Show
            fB861E.ZOrder
        Case "B861F":
            If fB861F Is Nothing Then Set fB861F = New frmBBS861
            Call SetParent(fB861F.hWnd, picForm.hWnd)
            fB861F.CDINDEX = BC2_JOB
            fB861F.Show
            fB861F.ZOrder
        Case "B861G":
            If fB861G Is Nothing Then Set fB861G = New frmBBS861
            Call SetParent(fB861G.hWnd, picForm.hWnd)
            fB861G.CDINDEX = BC2_BLD_SRC
            fB861G.Show
            fB861G.ZOrder
        Case "B861H":
            If fB861H Is Nothing Then Set fB861H = New frmBBS861
            Call SetParent(fB861H.hWnd, picForm.hWnd)
            fB861H.CDINDEX = BC2_EXP_RESON
            fB861H.Show
            fB861H.ZOrder
        Case "B861I":
            If fB861I Is Nothing Then Set fB861I = New frmBBS861
            Call SetParent(fB861I.hWnd, picForm.hWnd)
            fB861I.CDINDEX = BC2_Trans_Reason
            fB861I.Show
            fB861I.ZOrder
        Case "B861J":
            If fB861J Is Nothing Then Set fB861J = New frmBBS861
            Call SetParent(fB861J.hWnd, picForm.hWnd)
            fB861J.CDINDEX = BC2_MATERIAL
            fB861J.Show
            fB861J.ZOrder
        Case "B861K":
            If fB861K Is Nothing Then Set fB861K = New frmBBS861
            Call SetParent(fB861K.hWnd, picForm.hWnd)
            fB861K.CDINDEX = BC2_REACTION_TEST
            fB861K.Show
            fB861K.ZOrder
        Case "B861L":
            If fB861L Is Nothing Then Set fB861L = New frmBBS861
            Call SetParent(fB861L.hWnd, picForm.hWnd)
            fB861L.CDINDEX = BC2_CANCELRSN
            fB861L.Show
            fB861L.ZOrder
        Case "B861M":
            If fB861M Is Nothing Then Set fB861M = New frmBBS861
            Call SetParent(fB861M.hWnd, picForm.hWnd)
            fB861M.CDINDEX = BC2_COMPO_GROUP
            fB861M.Show
            fB861M.ZOrder
        Case "B861N":
            If fB861N Is Nothing Then Set fB861N = New frmBBS861
            Call SetParent(fB861N.hWnd, picForm.hWnd)
            fB861N.CDINDEX = BC2_DONOR_REASON
            fB861N.Show
            fB861N.ZOrder
        Case "B861O":
            If fB861O Is Nothing Then Set fB861O = New frmBBS861
            Call SetParent(fB861O.hWnd, picForm.hWnd)
            fB861O.CDINDEX = BC2_RC_COMPO
            fB861O.Show
            fB861O.ZOrder
        Case "B861Q":
            If fB861Q Is Nothing Then Set fB861Q = New frmBBS861
            Call SetParent(fB861Q.hWnd, picForm.hWnd)
            fB861Q.CDINDEX = BC2_RC_ABO
            fB861Q.Show
            fB861Q.ZOrder
        Case "B861R":
            If fB861R Is Nothing Then Set fB861R = New frmBBS861
            Call SetParent(fB861R.hWnd, picForm.hWnd)
            fB861R.CDINDEX = BC2_RC_VOL
            fB861R.Show
            fB861R.ZOrder
        Case "B861S":
            If fB861S Is Nothing Then Set fB861S = New frmBBS861
            Call SetParent(fB861S.hWnd, picForm.hWnd)
            fB861S.CDINDEX = BC2_BLOOD_BAR
            fB861S.Show
            fB861S.ZOrder
        Case "B861P"
            If fB861P Is Nothing Then Set fB861P = New frmBBS861
            Call SetParent(fB861P.hWnd, picForm.hWnd)
            fB861P.CDINDEX = BC2_REACTION_REASON
            fB861P.Show
            fB861P.ZOrder
    End Select
    
    Exit Sub
    
MasterTreeviewNodeClick_error:
    MsgBox Err.Description, vbCritical, "프로그램 로드"
End Sub

Public Sub MasterClose()

End Sub

Private Sub Class_Terminate()
    If Not (fB800 Is Nothing) Then Unload fB800: Set fB800 = Nothing
    If Not (fB801 Is Nothing) Then Unload fB801: Set fB801 = Nothing
    If Not (fB802 Is Nothing) Then Unload fB802: Set fB802 = Nothing
    If Not (fB803 Is Nothing) Then Unload fB803: Set fB803 = Nothing
    
    If Not (fB804 Is Nothing) Then Unload fB804: Set fB804 = Nothing
    
    If Not (fB811 Is Nothing) Then Unload fB811: Set fB811 = Nothing
    If Not (fB812 Is Nothing) Then Unload fB812: Set fB812 = Nothing
    If Not (fB813 Is Nothing) Then Unload fB813: Set fB813 = Nothing
    If Not (fB814 Is Nothing) Then Unload fB814: Set fB814 = Nothing
    If Not (fB815 Is Nothing) Then Unload fB815: Set fB815 = Nothing
    If Not (fB816 Is Nothing) Then Unload fB816: Set fB816 = Nothing
    If Not (fB817 Is Nothing) Then Unload fB817: Set fB817 = Nothing
    If Not (fB818 Is Nothing) Then Unload fB818: Set fB818 = Nothing
    If Not (fB819 Is Nothing) Then Unload fB819: Set fB819 = Nothing
    If Not (fB820 Is Nothing) Then Unload fB820: Set fB820 = Nothing
    If Not (fB821 Is Nothing) Then Unload fB821: Set fB821 = Nothing
    If Not (fB822 Is Nothing) Then Unload fB822: Set fB822 = Nothing
    If Not (fB823 Is Nothing) Then Unload fB823: Set fB823 = Nothing
    If Not (fB824 Is Nothing) Then Unload fB824: Set fB824 = Nothing
'    If Not (fB825 Is Nothing) Then Unload fB825: Set fB825 = Nothing
'    If Not (fB826 Is Nothing) Then Unload fB826: Set fB826 = Nothing
'    If Not (fB827 Is Nothing) Then Unload fB827: Set fB827 = Nothing
    If Not (fB828 Is Nothing) Then Unload fB828: Set fB828 = Nothing
     
    If Not (fB861B Is Nothing) Then Unload fB861B: Set fB861B = Nothing
    If Not (fB861C Is Nothing) Then Unload fB861C: Set fB861C = Nothing
    If Not (fB861D Is Nothing) Then Unload fB861D: Set fB861D = Nothing
    If Not (fB861E Is Nothing) Then Unload fB861E: Set fB861E = Nothing
    If Not (fB861F Is Nothing) Then Unload fB861F: Set fB861F = Nothing
    If Not (fB861G Is Nothing) Then Unload fB861G: Set fB861G = Nothing
    If Not (fB861H Is Nothing) Then Unload fB861H: Set fB861H = Nothing
    If Not (fB861I Is Nothing) Then Unload fB861I: Set fB861I = Nothing
    If Not (fB861J Is Nothing) Then Unload fB861J: Set fB861J = Nothing
    If Not (fB861K Is Nothing) Then Unload fB861K: Set fB861K = Nothing
    If Not (fB861L Is Nothing) Then Unload fB861L: Set fB861L = Nothing
    If Not (fB861M Is Nothing) Then Unload fB861M: Set fB861M = Nothing
    If Not (fB861N Is Nothing) Then Unload fB861N: Set fB861N = Nothing
    If Not (fB861O Is Nothing) Then Unload fB861O: Set fB861O = Nothing
    If Not (fB861P Is Nothing) Then Unload fB861P: Set fB861P = Nothing
    If Not (fB861Q Is Nothing) Then Unload fB861Q: Set fB861Q = Nothing
    If Not (fB861R Is Nothing) Then Unload fB861R: Set fB861R = Nothing
End Sub
