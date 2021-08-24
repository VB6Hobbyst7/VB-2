VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStaticsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private fB911 As frmBBS911
Private fB912 As frmBBS912
Private fB913 As frmBBS913
Private fB914 As frmBBS914
Private fB915 As frmBBS915
Private fB916 As frmBBS916
Private fB917 As frmBBS917
Private fB919 As frmBBS919
Private fB920 As frmBBS920
Private fB921 As frmBBS921
Private fB922 As frmBBS922
Private fB923 As frmBBS923
Private fB924 As frmBBS924
Private fB925 As frmBBS925
Private fB961 As frmBBS961
Private fB962 As frmBBS962
Private fB963 As frmBBS963
Private fB964 As frmBBS964
Private fB965 As frmBBS965

Public Sub StaticsTreeviewLoad(tvwMenu As Object)
    Dim ii As Integer
    ii = 1
    
    With tvwMenu
        .Nodes.Clear
        Call .Nodes.ADD(, , "B" & ii, "마감", 1)
'        Call .Nodes.ADD("B" & ii, tvwChild, "BBS925", "혈액일마감", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS924", "월별혈액입출고현황", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS964", "진료과/병동별 수혈통계", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS962", "월중혈액수불통계", 2)
        .Nodes(.Nodes.Count).EnsureVisible
        
        If DonorUserFg = True Then
            Call .Nodes.ADD("B" & ii, tvwNext, "B" & ii + 1, "혈액원 리포트", 1)
            ii = ii + 1
'            Call .Nodes.ADD("B" & ii, tvwChild, "BBS920", "부적격혈액이송", 2)
            Call .Nodes.ADD("B" & ii, tvwChild, "BBS921", "헌혈실적통지서", 2)
'            Call .Nodes.ADD("B" & ii, tvwChild, "BBS922", "헌혈경력조회서", 2)
'            Call .Nodes.ADD("B" & ii, tvwChild, "BBS923", "헌혈자대장", 2)
'            Call .Nodes.ADD("B" & ii, tvwChild, "BBS965", "DM발송", 2)
            
            .Nodes(.Nodes.Count).EnsureVisible
        End If
        
        Call .Nodes.ADD("B" & ii, tvwNext, "B" & ii + 1, "통계", 1)
        ii = ii + 1
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS913", "혈액일보", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS914", "C-T Ratio", 2)
'        Call .Nodes.ADD("B" & ii, tvwChild, "BBS915", "폐기사유별 건수", 2)
        
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS961", "혈액출고/폐기대장", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS912", "혈액반납량/사용량", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS916", "혈액제제별 사용현황", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS919", "혈액제제별 입/출고현황", 2) '전주에서 추가됨
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS917", "방사선조사 혈액제제", 2)
        Call .Nodes.ADD("B" & ii, tvwChild, "BBS925", "혈액은행 검사통계", 2) '전주에서 추가됨
'        Call .Nodes.ADD("B" & ii, tvwChild, "BBS963", "수혈부작용 건수", 2)
        
        
'        If TransReactionUsed = True Then
'            Call .Nodes.ADD("B" & ii, tvwChild, "BBS963", "수혈부작용 건수", 2)
'        End If
        
        .Nodes(.Nodes.Count).EnsureVisible
        .BorderStyle = vbFixedSingle
    End With
End Sub
'
Public Sub StaticsTreeviewNodeClick(ByVal key As String, picForm As Object)

On Error GoTo StaticsTreeviewNodeClick_error

    Select Case key
        Case "BBS911":
            If fB911 Is Nothing Then Set fB911 = New frmBBS911
            Call SetParent(fB911.hWnd, picForm.hWnd)
            fB911.Show:
            fB911.ZOrder
        Case "BBS912":
            If fB912 Is Nothing Then Set fB912 = New frmBBS912
            Call SetParent(fB912.hWnd, picForm.hWnd)
            fB912.Show:
            fB912.ZOrder
        Case "BBS913":
            If fB913 Is Nothing Then Set fB913 = New frmBBS913
            Call SetParent(fB913.hWnd, picForm.hWnd)
            fB913.Show:
            fB913.ZOrder
        Case "BBS914":
            If fB914 Is Nothing Then Set fB914 = New frmBBS914
            Call SetParent(fB914.hWnd, picForm.hWnd)
            fB914.Show:
            fB914.ZOrder
        Case "BBS915":
            If fB915 Is Nothing Then Set fB915 = New frmBBS915
            Call SetParent(fB915.hWnd, picForm.hWnd)
            fB915.Show:
            fB915.ZOrder
        Case "BBS916":
            If fB916 Is Nothing Then Set fB916 = New frmBBS916
            Call SetParent(fB916.hWnd, picForm.hWnd)
            fB916.Show:
            fB916.ZOrder
        Case "BBS917":
            If fB917 Is Nothing Then Set fB917 = New frmBBS917
            Call SetParent(fB917.hWnd, picForm.hWnd)
            fB917.Show:
            fB917.ZOrder
        Case "BBS919":
            If fB919 Is Nothing Then Set fB919 = New frmBBS919
            Call SetParent(fB919.hWnd, picForm.hWnd)
            fB919.Show:
            fB919.ZOrder
        Case "BBS920":
            If fB920 Is Nothing Then Set fB920 = New frmBBS920
            Call SetParent(fB920.hWnd, picForm.hWnd)
            fB920.Show:
            fB920.ZOrder
        Case "BBS921":
            If fB921 Is Nothing Then Set fB921 = New frmBBS921
            Call SetParent(fB921.hWnd, picForm.hWnd)
            fB921.Show:
            fB921.ZOrder
        Case "BBS922":
            If fB922 Is Nothing Then Set fB922 = New frmBBS922
            Call SetParent(fB922.hWnd, picForm.hWnd)
            fB922.Show:
            fB922.ZOrder
        Case "BBS923":
            If fB923 Is Nothing Then Set fB923 = New frmBBS923
            Call SetParent(fB923.hWnd, picForm.hWnd)
            fB923.Show:
            fB923.ZOrder
        Case "BBS924":
            If fB924 Is Nothing Then Set fB924 = New frmBBS924
            Call SetParent(fB924.hWnd, picForm.hWnd)
            fB924.Show:
            fB924.ZOrder
        Case "BBS925":
            If fB925 Is Nothing Then Set fB925 = New frmBBS925
            Call SetParent(fB925.hWnd, picForm.hWnd)
            fB925.Show:
            fB925.ZOrder
        Case "BBS961":
            If fB961 Is Nothing Then Set fB961 = New frmBBS961
            Call SetParent(fB961.hWnd, picForm.hWnd)
            fB961.Show:
            fB961.ZOrder
        Case "BBS962":
            If fB962 Is Nothing Then Set fB962 = New frmBBS962
            Call SetParent(fB962.hWnd, picForm.hWnd)
            fB962.Show:
            fB962.ZOrder
        Case "BBS963":
            If fB963 Is Nothing Then Set fB963 = New frmBBS963
            Call SetParent(fB963.hWnd, picForm.hWnd)
            fB963.Show:
            fB963.ZOrder
        Case "BBS964":
            If fB964 Is Nothing Then Set fB964 = New frmBBS964
            Call SetParent(fB964.hWnd, picForm.hWnd)
            fB964.Show:
            fB964.ZOrder
        Case "BBS965":
            If fB965 Is Nothing Then Set fB965 = New frmBBS965
            Call SetParent(fB965.hWnd, picForm.hWnd)
            fB965.Show:
            fB965.ZOrder
    End Select
    
    Exit Sub
    
StaticsTreeviewNodeClick_error:
    MsgBox Err.Description, vbCritical, "프로그램 로딩"
End Sub


Private Sub Class_Terminate()
    If Not (fB911 Is Nothing) Then Unload fB911: Set fB911 = Nothing
    If Not (fB912 Is Nothing) Then Unload fB912: Set fB912 = Nothing
    If Not (fB913 Is Nothing) Then Unload fB913: Set fB913 = Nothing
    If Not (fB914 Is Nothing) Then Unload fB914: Set fB914 = Nothing
    If Not (fB915 Is Nothing) Then Unload fB915: Set fB915 = Nothing
    If Not (fB916 Is Nothing) Then Unload fB916: Set fB916 = Nothing
    If Not (fB917 Is Nothing) Then Unload fB917: Set fB917 = Nothing
    If Not (fB919 Is Nothing) Then Unload fB919: Set fB919 = Nothing
    
    If Not (fB920 Is Nothing) Then Unload fB920: Set fB920 = Nothing
    If Not (fB921 Is Nothing) Then Unload fB921: Set fB921 = Nothing
    If Not (fB922 Is Nothing) Then Unload fB922: Set fB922 = Nothing
    If Not (fB923 Is Nothing) Then Unload fB923: Set fB923 = Nothing
    If Not (fB924 Is Nothing) Then Unload fB924: Set fB924 = Nothing
    If Not (fB925 Is Nothing) Then Unload fB925: Set fB925 = Nothing
    
    If Not (fB961 Is Nothing) Then Unload fB961: Set fB961 = Nothing
    If Not (fB962 Is Nothing) Then Unload fB962: Set fB962 = Nothing
    If Not (fB963 Is Nothing) Then Unload fB963: Set fB963 = Nothing
    If Not (fB964 Is Nothing) Then Unload fB964: Set fB964 = Nothing
    If Not (fB965 Is Nothing) Then Unload fB965: Set fB965 = Nothing
End Sub
