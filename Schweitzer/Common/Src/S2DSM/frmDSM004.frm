VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDSM004 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 관리"
   ClientHeight    =   6300
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8160
   Icon            =   "frmDSM004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8160
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.ImageList imgList 
      Left            =   105
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSM004.frx":06EA
            Key             =   "등록"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSM004.frx":0A06
            Key             =   "그룹등록"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSM004.frx":0C22
            Key             =   "편집"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSM004.frx":0F3E
            Key             =   "삭제"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDSM004.frx":125A
            Key             =   "종료"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwGroup 
      Height          =   2310
      Left            =   75
      TabIndex        =   1
      Top             =   3930
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   4075
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776183
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "그룹명"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "그룹설명"
         Object.Width           =   3201
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Developer"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Supervisor"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Manager"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "End User"
         Object.Width           =   1905
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   3090
      Left            =   75
      TabIndex        =   0
      Top             =   825
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5450
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776191
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "로그인ID"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "직원아이디"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "직원명"
         Object.Width           =   1931
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "부서"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "비밀번호"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "그룹아이디"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "그룹명"
         Object.Width           =   3863
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "그룹권한"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "사용자 설명"
         Object.Width           =   5980
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarUser 
      Align           =   1  '위 맞춤
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1402
      ButtonWidth     =   1561
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "새 사용자"
            Key             =   "newuser"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "새 그룹"
            Key             =   "newgroup"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "편집"
            Key             =   "edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "삭제"
            Key             =   "delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "종료"
            Key             =   "exit"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmDSM004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private clsRef As New clsDSMUserInfo

'Private Sub Form_Activate()
'    Query_lvwUser
'End Sub

Private Sub Form_Load()
'    Query_lvwGroup
    Call LoadUser
    Call LoadGroup
End Sub

Private Sub LoadUser()
    Dim Rs As Recordset
    Dim strSQL As String
    Dim itmX As ListItem
    
    strSQL = " select c.loginid,a.empid, a.empnm, a.deptcd, c.loginpass, b.groupid, b.groupnm, b.userfg, c.logindesc " & _
             " from " & T_COM006 & " a, " & T_COM008 & " b, " & T_COM010 & " c " & _
             " where a.empid = c.empid " & _
             " and c.groupid = b.groupid " & " order by a.empnm "
    Set Rs = New Recordset
    
    Rs.Open strSQL, DBConn
    
    lvwUser.ListItems.clear
    Do Until Rs.EOF
        Set itmX = lvwUser.ListItems.Add()
        itmX.Text = Rs.Fields("loginid").Value & ""
        itmX.SubItems(1) = Rs.Fields("empid").Value & ""
        itmX.SubItems(2) = Rs.Fields("empnm").Value & ""
        itmX.SubItems(3) = Rs.Fields("deptcd").Value & ""
        itmX.SubItems(4) = Rs.Fields("loginpass").Value & ""
        itmX.SubItems(5) = Rs.Fields("groupid").Value & ""
        itmX.SubItems(6) = Rs.Fields("groupnm").Value & ""
        itmX.SubItems(7) = Rs.Fields("userfg").Value & ""
        itmX.SubItems(8) = Rs.Fields("logindesc").Value & ""
        
        Rs.MoveNext
    Loop
    Set Rs = Nothing
End Sub

Private Sub LoadGroup()
    Dim Rs As Recordset
    Dim strSQL As String
    Dim itmX As ListItem
    
    strSQL = " select groupid, groupnm, groupdesc, userfg from " & T_COM008 & " order by groupnm "
    
    Set Rs = New Recordset
    
    Rs.Open strSQL, DBConn
    
    lvwGroup.ListItems.clear
    Do Until Rs.EOF
        Set itmX = lvwGroup.ListItems.Add()
        itmX.Text = Rs.Fields("groupnm").Value & ""
        itmX.SubItems(1) = Rs.Fields("groupdesc").Value & ""
        
        Select Case Rs.Fields("userfg").Value & ""
            Case "D"
                itmX.SubItems(2) = "YES"
            Case "S"
                itmX.SubItems(3) = "YES"
            Case "M"
                itmX.SubItems(4) = "YES"
            Case "E"
                itmX.SubItems(5) = "YES"
        End Select
        
        Rs.MoveNext
    Loop
    Set Rs = Nothing
End Sub

'Private Sub Query_lvwUser()
'    Dim strSQL As String
''LoginId , EmpId, Dept, passwd, passwd, groupcd, Desc
'
'    strSQL = " select c.loginid, a.empnm, a.deptcd, c.loginpass, b.groupid, b.groupnm, c.logindesc " & _
'             " from " & T_COM006 & " a, " & T_COM008 & " b, " & T_COM010 & " c " & _
'             " where a.empid = c.empid " & _
'             " and c.groupid = b.groupid "
'
'    clsRef.Lvw_Set lvwUser, strSQL, 3
'End Sub

'Private Sub Query_lvwGroup()
'    Dim Rs As New Recordset
'    Dim LvwItm As Object
'    Dim strSQL As String
'
'    On Error GoTo ErrlvwGroup
'
'    lvwGroup.ListItems.clear
'
'    strSQL = " select groupid, groupnm, groupdesc, userfg from " & T_COM008
'
'    Rs.Open strSQL, DBConn
'
'    While Rs.EOF = False
'          Set LvwItm = lvwGroup.ListItems.Add()
'
'          With LvwItm
'               .Text = IIf(IsNull(Rs.Fields("groupnm").Value) = True, "", "" & Rs.Fields("groupnm").Value)
'               .SubItems(1) = IIf(IsNull(Rs.Fields("groupdesc").Value) = True, "", "" & Rs.Fields("groupdesc").Value)
'
'               If IsNull(Rs.Fields("userfg").Value) = False Or "" & Rs.Fields("userfg").Value <> "" Then
'                  If "" & Rs.Fields("userfg").Value = "M" Then .SubItems(2) = "Yes"
'                  If "" & Rs.Fields("userfg").Value = "D" Then .SubItems(3) = "Yes"
'                  If "" & Rs.Fields("userfg").Value = "S" Then .SubItems(4) = "Yes"
'               End If
'
'               Rs.MoveNext
'          End With
'    Wend
'
'    Set Rs = Nothing
'    Exit Sub
'ErrlvwGroup:
'    Set Rs = Nothing
'    MsgBox Err.Description, vbCritical
'
'End Sub

'Private Sub lvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Static i As Integer
'
'    With lvwGroup
'         .SortKey = ColumnHeader.Index - 1
'         .SortOrder = IIf(i = 0, lvwAscending, lvwDescending)
'         .Sorted = True
'    End With
'
'    i = (i + 1) Mod 2
'End Sub
'
'Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Static i As Integer
'
'    With lvwUser
'         .SortKey = ColumnHeader.Index - 1
'         .SortOrder = IIf(i = 0, lvwAscending, lvwDescending)
'         .Sorted = True
'    End With
'
'    i = (i + 1) Mod 2
'End Sub

'Private Sub lvwUser_DblClick()
'    mnuEdit_Click
'End Sub

'Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    GblUser = Item.Text
'End Sub

'Private Sub mnuDelete_Click()
'    Dim strMsg As String
'    Dim Message As String
'
'    If GblUser <> "" Then
'       strMsg = "[" & GblUser & "]" & " " & _
'                "ID 삭제를 요청하셨습니다. 한 번 삭제된 정보는 복구 할 수 없습니다. 계속 하시겠습니까?"
'       Message = MsgBox(strMsg, vbCritical + vbOKCancel, "삭제")
'       If Message = vbCancel Then Exit Sub
'       clsRef.COM010_Delete GblUser
'       lvwItem_Remove
'    End If
'
'End Sub

Private Sub lvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'소트
    Static blnToggle() As Boolean
    Static blnFirst As Boolean
    Dim i As Long
    
    If blnFirst = False Then
        ReDim blnToggle(lvwGroup.ColumnHeaders.Count - 1)
        blnFirst = True
    End If
    
    '▲▼
    
    For i = 1 To lvwGroup.ColumnHeaders.Count
        lvwGroup.ColumnHeaders(i).Text = Trim(Replace(lvwGroup.ColumnHeaders(i).Text, "▲", ""))
        lvwGroup.ColumnHeaders(i).Text = Trim(Replace(lvwGroup.ColumnHeaders(i).Text, "▼", ""))
    Next
    
    
    With lvwGroup
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(blnToggle(ColumnHeader.Index - 1), lvwDescending, lvwAscending)
        .Sorted = True
        
        ColumnHeader.Text = ColumnHeader.Text & " " & IIf(.SortOrder = lvwAscending, "▲", "▼")
        
        blnToggle(ColumnHeader.Index - 1) = IIf(blnToggle(ColumnHeader.Index - 1), False, True)
    End With
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'소트
    Static blnToggle() As Boolean
    Static blnFirst As Boolean
    Dim i As Long
    
    If blnFirst = False Then
        ReDim blnToggle(lvwUser.ColumnHeaders.Count - 1)
        blnFirst = True
    End If
    
    '▲▼
    
    For i = 1 To lvwUser.ColumnHeaders.Count
        lvwUser.ColumnHeaders(i).Text = Trim(Replace(lvwUser.ColumnHeaders(i).Text, "▲", ""))
        lvwUser.ColumnHeaders(i).Text = Trim(Replace(lvwUser.ColumnHeaders(i).Text, "▼", ""))
    Next
    
    
    With lvwUser
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(blnToggle(ColumnHeader.Index - 1), lvwDescending, lvwAscending)
        .Sorted = True
        
        ColumnHeader.Text = ColumnHeader.Text & " " & IIf(.SortOrder = lvwAscending, "▲", "▼")
        
        blnToggle(ColumnHeader.Index - 1) = IIf(blnToggle(ColumnHeader.Index - 1), False, True)
    End With
End Sub

Private Sub lvwUser_DblClick()
    Call EditUser
End Sub

Private Sub lvwUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPop As clsPopupMenu
    Dim strSQL As String
    
    If lvwUser.ListItems.Count = 0 Then Exit Sub
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    
    If Button = vbRightButton Then
        Set objPop = New clsPopupMenu
        
        With objPop
            .AddMenu 1, "삭제"
            .AddMenu 2, "-", eSEPARATOR
            .AddMenu 3, "편집"
            
            .PopupMenus Me.hwnd
            
            If .MenuID = 1 Then
                Call DeleteUser
            ElseIf .MenuID = 3 Then
                Call EditUser
            End If
        End With
        
        Set objPop = Nothing
    End If
End Sub

Private Sub DeleteUser()
'삭제권한 체크
    Dim strLoginId As String
    Dim strSQL As String

    If lvwUser.ListItems.Count = 0 Then Exit Sub
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    
    If ObjMyUser.IsManager Then
        If (lvwUser.SelectedItem.SubItems(7) = "D") Or (lvwUser.SelectedItem.SubItems(7) = "S") Then
            MsgBox "삭제할 권한이 없습니다.", vbExclamation
            Exit Sub
        End If
    ElseIf ObjMyUser.IsSupervisor Then
        If lvwUser.SelectedItem.SubItems(7) = "D" Then
            MsgBox "삭제할 권한이 없습니다.", vbExclamation
            Exit Sub
        End If
    End If
    
    strLoginId = lvwUser.SelectedItem.Text

    If MsgBox("사용자를 삭제하시겠습니까?", vbExclamation + vbDefaultButton2 + vbYesNo) = vbYes Then
        On Error GoTo ErrTrap
        
        DBConn.BeginTrans
        strSQL = " delete " & T_COM010 & _
                 " where " & DBW("loginid=", strLoginId)
        
        DBConn.Execute strSQL
        DBConn.CommitTrans
        
        MsgBox "정상적으로 처리되었습니다.", vbInformation
        lvwUser.ListItems.Remove lvwUser.SelectedItem.Index
        
        GoTo Skip
ErrTrap:
        DBConn.RollbackTrans
        MsgBox "처리도중 오류가 발생하였습니다.", vbExclamation
Skip:
    End If
End Sub

Private Sub EditUser()
'편집 가능 권한 체크

    If lvwUser.ListItems.Count = 0 Then Exit Sub
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    
    If ObjMyUser.IsManager Then
        If (lvwUser.SelectedItem.SubItems(7) = "D") Or (lvwUser.SelectedItem.SubItems(7) = "S") Then
            MsgBox "사용자 설정을 변경할 권한이 없습니다.", vbExclamation
            Exit Sub
        End If
    ElseIf ObjMyUser.IsSupervisor Then
        If lvwUser.SelectedItem.SubItems(7) = "D" Then
            MsgBox "사용자 설정을 변경할 권한이 없습니다.", vbExclamation
            Exit Sub
        End If
    End If
    
    With frmDSM005
        .EditFg = True
        .LoginId = lvwUser.SelectedItem.Text
        .EmpId = lvwUser.SelectedItem.SubItems(1)
        .DeptCd = lvwUser.SelectedItem.SubItems(3)
        .LogInPass = lvwUser.SelectedItem.SubItems(4)
        .GroupID = lvwUser.SelectedItem.SubItems(5)
        .LoginDesc = lvwUser.SelectedItem.SubItems(8)
        .Show vbModal, Me
    End With
    
    Call LoadUser
    Call LoadGroup
End Sub

'Private Sub lvwItem_Remove()
'    clsRef.Lvw_Item_Remove lvwUser, GblUser
'End Sub

'Private Sub mnuEdit_Click()
'    GblEdit = True
'    frmDSM005.UserId = GblUser
'    frmDSM005.ChangePwd = False
'    frmDSM005.Show vbModal, Me
'End Sub

'Private Sub mnuExit_Click()
'    Unload Me
'End Sub

'Private Sub mnuNewGroup_Click()
'
'End Sub
'
'Private Sub mnuNewUser_Click()
'    GblEdit = False
'    frmDSM005.Show vbModal, Me
'End Sub

Private Sub tbarUser_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strMsg As String
    Dim Message As String
    
    Select Case Button.Key
        Case "newuser"
            frmDSM005.EditFg = False
            frmDSM005.Show vbModal, Me
             Call LoadUser
             Call LoadGroup
        Case "newgroup"
             frmDSM003_N.Show vbModal, Me
             Call LoadUser
             Call LoadGroup
        Case "edit"
'             GblEdit = True
'             frmDSM005.UserId = GblUser
'             frmDSM005.ChangePwd = False
'LoginId , EmpId, Dept, passwd, passwd, groupcd, Desc
            Call EditUser
        Case "delete"
            Call DeleteUser
            
'             If GblUser <> "" Then
'                strMsg = "[" & GblUser & "]" & " " & _
'                         "ID 삭제를 요청하셨습니다. 한 번 삭제된 정보는 복구 할 수 없습니다. 계속 하시겠습니까?"
'                Message = MsgBox(strMsg, vbCritical + vbOKCancel, "삭제")
'                If Message = vbCancel Then Exit Sub
''                clsRef.COM010_Delete GblUser
''                lvwItem_Remove
'             End If
        Case "exit"
             Unload Me
    End Select
End Sub
