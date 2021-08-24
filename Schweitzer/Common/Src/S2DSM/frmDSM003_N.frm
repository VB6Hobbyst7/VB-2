VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDSM003_N 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "그룹등록 및 권한등록"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   105
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈  그룹 관리"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Top             =   90
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   529
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈  권한 관리"
      LeftGab         =   100
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   6270
      Left            =   120
      TabIndex        =   3
      Top             =   315
      Width           =   3570
      Begin VB.TextBox txtGroupID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "그룹명"
         Top             =   240
         Width           =   2205
      End
      Begin VB.TextBox txtGroupNm 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "그룹명"
         Top             =   600
         Width           =   2205
      End
      Begin VB.TextBox txtGroupDesc 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "그룹설명"
         Top             =   960
         Width           =   2205
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   765
         Left            =   105
         TabIndex        =   10
         Top             =   1245
         Width           =   3345
         Begin VB.OptionButton optUserFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Manager"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   11
            Tag             =   "M"
            Top             =   450
            Width           =   1215
         End
         Begin VB.OptionButton optUserFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "De&veloper"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   12
            Tag             =   "D"
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton optUserFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Su&pervisor"
            Height          =   180
            Index           =   2
            Left            =   1650
            TabIndex        =   13
            Tag             =   "S"
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton optUserFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&End User"
            Height          =   180
            Index           =   0
            Left            =   1650
            TabIndex        =   14
            Tag             =   "E"
            Top             =   450
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   450
         Left            =   105
         TabIndex        =   6
         Top             =   1965
         Width           =   3345
         Begin VB.CheckBox chkDeptFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "혈액 은행"
            Height          =   255
            Index           =   1
            Left            =   1155
            TabIndex        =   9
            Tag             =   "B"
            Top             =   165
            Width           =   1080
         End
         Begin VB.CheckBox chkDeptFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "진단 병리"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Tag             =   "A"
            Top             =   165
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CheckBox chkDeptFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "임상 병리"
            Height          =   255
            Index           =   2
            Left            =   2220
            TabIndex        =   7
            Tag             =   "L"
            Top             =   165
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView lvwGroup 
         Height          =   3705
         Left            =   90
         TabIndex        =   19
         Top             =   2460
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   6535
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "그룹 ID"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "그룹명"
            Object.Width           =   3545
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "그룹설명"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "권한"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "APS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "BBS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LIS"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "그  룹  명 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   105
         TabIndex        =   17
         Tag             =   "105"
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "그  룹 ID  : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Tag             =   "105"
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "그룹  설명 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   105
         TabIndex        =   15
         Tag             =   "105"
         Top             =   1020
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   6270
      Left            =   3840
      TabIndex        =   18
      Top             =   300
      Width           =   6975
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00EEEBED&
         Caption         =   "삭제"
         Height          =   405
         Left            =   3420
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   5775
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00EEEBED&
         Caption         =   "닫기(&X)"
         Height          =   405
         Left            =   5640
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   5775
         Width           =   1050
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00EEEBED&
         Caption         =   "화면지움"
         Height          =   405
         Left            =   4530
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   5775
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00EEEBED&
         Caption         =   "저장(&S)"
         Height          =   405
         Left            =   2310
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   5775
         Width           =   1050
      End
      Begin FPSpread.vaSpread tblForm 
         Height          =   5565
         Left            =   60
         TabIndex        =   23
         Top             =   150
         Width           =   6840
         _Version        =   196608
         _ExtentX        =   12065
         _ExtentY        =   9816
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   7
         MaxRows         =   50
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14411494
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmDSM003_N.frx":0000
      End
   End
End
Attribute VB_Name = "frmDSM003_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dev. by Legends
'2003/09/29

Private Sub chkDeptFg_Click(Index As Integer)
    If Screen.ActiveControl.Name <> chkDeptFg(Index).Name Then Exit Sub
    
    If Trim(txtGroupID.Text) = "" Then Exit Sub
    
    tblForm.MaxRows = 0
    tblForm.MaxRows = 23
    Call medClearTable(tblForm)
    
    Call LoadForm
End Sub

Private Sub cmdClear_Click()
    txtGroupID.Text = ""
    Call InitForm
End Sub

Private Sub cmdDelete_Click()
    Dim strMsg As VbMsgBoxResult
    Dim strSQL1 As String
    Dim strSQL2 As String
    
    If Trim(txtGroupID.Text) = "" Then Exit Sub
    
    strMsg = MsgBox("선택한 그룹내역 및 등록된 권한내역을 삭제합니다. 삭제하시겠습니까?" & vbNewLine & vbNewLine & "주)이 작업의 수행으로 프로그램 권한과 관련한 오류가 발생할 수도 있습니다.", vbExclamation + vbYesNo)
    If strMsg = vbNo Then Exit Sub
    
    strSQL1 = "delete " & T_COM008 & " where  " & DBW("groupid =", Trim(txtGroupID.Text))
    strSQL2 = "delete " & T_COM009 & " where  " & DBW("groupid =", Trim(txtGroupID.Text))
    
On Error GoTo ErrTrap
    DBConn.BeginTrans
    DBConn.Execute strSQL1
    DBConn.Execute strSQL2
    DBConn.CommitTrans
    
    txtGroupID.Text = ""
    Call InitForm
    Call LoadGroup
    Exit Sub
    
ErrTrap:
    DBConn.RollbackTrans
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function SaveGroup() As Boolean
'com008에 넣는거
    Dim strSQL1 As String
    Dim strSQL2 As String
    
    If ChkValidation = False Then GoTo ErrTrap
    
    strSQL1 = "delete " & T_COM008 & " where  " & DBW("groupid =", Trim(txtGroupID.Text))
    strSQL2 = "Insert Into " & T_COM008 & " (groupid,groupnm,groupdesc,userfg,apsfg,bbsfg,lisfg) Values(" & _
              DBV("groupid", Trim(txtGroupID.Text), 1) & DBV("groupnm", Trim(txtGroupNm.Text), 1) & _
              DBV("groupdesc", Trim(txtGroupDesc.Text), 1) & _
              DBV("userfg", GetUserFg, 1) & DBV("apsfg", chkDeptFg(0).Value, 1) & DBV("bbsfg", chkDeptFg(1).Value, 1) & DBV("lisfg", chkDeptFg(2).Value) & ")"
    
On Error GoTo ErrTrap
'    DBConn.BeginTrans
    DBConn.Execute strSQL1
    DBConn.Execute strSQL2
'    DBConn.CommitTrans
    SaveGroup = True
    
    Exit Function
    
ErrTrap:
'    DBConn.RollbackTrans
    SaveGroup = False
End Function

Private Function ChkValidation() As Boolean
    Dim blnChk As Boolean
    Dim i As Long
    
    ChkValidation = False

    If Trim(txtGroupID.Text) = "" Then
        MsgBox "그룹 ID를 입력하세요.", vbExclamation
        Exit Function
    End If
    
    blnChk = False
    For i = optUserFg.lbound To optUserFg.UBound
        If optUserFg(i).Value Then
            blnChk = True
        End If
    Next
    
    If blnChk = False Then
        MsgBox "그룹의 권한을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    blnChk = False
    For i = chkDeptFg.lbound To chkDeptFg.UBound
        If chkDeptFg(i).Value = 1 Then
            blnChk = True
        End If
    Next
    
    If blnChk = False Then
        MsgBox "그룹이 참조하는 업부구분을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    ChkValidation = True
End Function

Private Function GetUserFg() As String
    Dim i As Long
    
    For i = optUserFg.lbound To optUserFg.UBound
        If optUserFg(i).Value Then
            GetUserFg = optUserFg(i).Tag
            Exit For
        End If
    Next
End Function

Private Function SaveForm() As Boolean
'com009에 넣는거.
    Dim arySQL() As String
    Dim i As Long
    Dim strDeptFg As String
    Dim strFormId As String
    Dim strReadFg As String
    Dim strWriteFg As String
    Dim strPrintFg As String
    
    ReDim arySQL(0)
    
    arySQL(0) = "delete " & T_COM009 & " where  " & DBW("groupid =", Trim(txtGroupID.Text))
    
    For i = 1 To tblForm.DataRowCnt
        ReDim Preserve arySQL(UBound(arySQL) + 1)
        
        With tblForm
            .Row = i
            .Col = 1: strDeptFg = Mid(.Value, 1, 1)
            .Col = 6: strFormId = .Value
            'CellType이 chkeckbox 인 경우에만 등록
            .Col = 3: strReadFg = IIf(.CellType = CellTypeCheckBox, .Value, "")
            .Col = 4: strWriteFg = IIf(.CellType = CellTypeCheckBox, .Value, "")
            .Col = 5: strPrintFg = IIf(.CellType = CellTypeCheckBox, .Value, "")
        
            arySQL(UBound(arySQL)) = " insert into " & T_COM009 & " (groupid,deptfg,formid,readfg,writefg,printfg) Values(" & _
                                     DBV("groupid", Trim(txtGroupID.Text), 1) & DBV("deptfg", strDeptFg, 1) & DBV("formid", strFormId, 1) & _
                                     DBV("readfg", strReadFg, 1) & DBV("writefg", strWriteFg, 1) & DBV("printfg", strPrintFg) & " )"
        End With
    Next
    
On Error GoTo ErrTrap
'    DBConn.BeginTrans
    For i = LBound(arySQL) To UBound(arySQL)
        If arySQL(i) <> "" Then
            DBConn.Execute arySQL(i)
        End If
    Next
'    DBConn.CommitTrans
    SaveForm = True
    Exit Function
    
ErrTrap:
'    DBConn.RollbackTrans
    SaveForm = False
End Function

Private Sub cmdSave_Click()
    Dim strMsg As VbMsgBoxResult
        
    If CheckValidation = False Then Exit Sub
    
    strMsg = MsgBox("그룹등록 및 권한등록 작업을 수행합니다. 저장하시겠습니까?" & vbNewLine & vbNewLine & "주)이 작업은 기존의 데이터를 모두 초기화하고 새로운 값으로 대치합니다.", vbExclamation + vbYesNo)
    If strMsg = vbNo Then Exit Sub

On Error GoTo ErrTrap
        
    DBConn.BeginTrans
        
    If SaveGroup = False Then GoTo ErrTrap
    If SaveForm = False Then GoTo ErrTrap
    
    DBConn.CommitTrans
    
    Call LoadGroup
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    
    Exit Sub
    
ErrTrap:
    DBConn.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = True
    
    If lvwGroup.ListItems.Count = 0 Then Exit Function
    If lvwGroup.SelectedItem Is Nothing Then Exit Function
        
    CheckValidation = False
        
    If ObjMyUser.IsManager Then
        '개발자, 수퍼바이저 인경우 하위 그룹으로 이동 못하도록
        If optUserFg(3).Value Or optUserFg(0).Value Then
            MsgBox "그룹권한을 변경할 권한이 없습니다.", vbExclamation
            Exit Function
        End If
    ElseIf ObjMyUser.IsSupervisor Then
        '개발자인 경우 하위그룹으로 이동 못하도록
        If optUserFg(2).Value Or optUserFg(3).Value Or optUserFg(0).Value Then
            MsgBox "그룹권한을 변경할 권한이 없습니다.", vbExclamation
            Exit Function
        End If
    End If
    
    CheckValidation = True
End Function

Private Sub Form_Activate()
    Call LoadGroup
End Sub

Private Sub Form_Load()
    txtGroupID.Text = ""
    Call InitForm
    lvwGroup.ListItems.clear
End Sub

Private Sub InitForm()
    Dim i As Long
    
'    txtGroupID.Text = ""
    txtGroupNm.Text = ""
    txtGroupDesc.Text = ""
    
    For i = optUserFg.lbound To optUserFg.UBound
        optUserFg(i).Value = False
    Next
    
    For i = chkDeptFg.lbound To chkDeptFg.UBound
        chkDeptFg(i).Value = 0
    Next
    
    '로긴사용자의 권한에 따라 선택할 수 있는 그룹을 지정한다.
    
    If ObjMyUser.IsManager Then
        optUserFg(1).Enabled = False
        optUserFg(2).Enabled = False
    ElseIf ObjMyUser.IsSupervisor Then
        optUserFg(1).Enabled = False
    End If
    
    tblForm.MaxRows = 0
    tblForm.MaxRows = 23
    Call medClearTable(tblForm)
End Sub

Private Sub LoadGroup()
    Dim Rs As Recordset
    Dim objSQL As clsDSMSqlStmt
    Dim strSQL As String
    Dim itmX As ListItem
    
    Set objSQL = New clsDSMSqlStmt
    Set Rs = New Recordset
    
    strSQL = "select * from " & T_COM008 & " order by groupid "
    
    Rs.Open strSQL, DBConn
    
    lvwGroup.ListItems.clear
    Do Until Rs.EOF
    
        Set itmX = lvwGroup.ListItems.Add()
        itmX.Text = Rs.Fields("groupid").Value & ""
        itmX.SubItems(1) = Rs.Fields("groupnm").Value & ""
        itmX.SubItems(2) = Rs.Fields("groupdesc").Value & ""
        itmX.SubItems(3) = Rs.Fields("userfg").Value & ""
        itmX.SubItems(4) = Rs.Fields("apsfg").Value & ""
        itmX.SubItems(5) = Rs.Fields("bbsfg").Value & ""
        itmX.SubItems(6) = Rs.Fields("lisfg").Value & ""
        
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub

Private Sub lvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
'폼별권한을 보여준다.
    
    If Item Is Nothing Then Exit Sub
    
    txtGroupID.Text = Item.Text
    txtGroupNm.Text = Item.SubItems(1)
    txtGroupDesc.Text = Item.SubItems(2)
    
    Select Case Item.SubItems(3)
        Case "D"
            optUserFg(1).Value = True
        Case "S"
            optUserFg(2).Value = True
        Case "M"
            optUserFg(3).Value = True
        Case "E"
            optUserFg(0).Value = True
    End Select
    
    chkDeptFg(0).Value = Val(Item.SubItems(4))
    chkDeptFg(1).Value = Val(Item.SubItems(5))
    chkDeptFg(2).Value = Val(Item.SubItems(6))
    
    tblForm.MaxRows = 0
    tblForm.MaxRows = 23
    Call medClearTable(tblForm)
    
    Call LoadForm
    
'선택한 권한별로 삭제권한을 준다.
    If ObjMyUser.IsManager Then
        If (Item.SubItems(3) = "M") Or (Item.SubItems(3) = "E") Then
            cmdDelete.Enabled = True
            cmdSave.Enabled = True
        Else
            cmdDelete.Enabled = False
            cmdSave.Enabled = False
        End If
    ElseIf ObjMyUser.IsSupervisor Then
        If (Item.SubItems(3) = "S") Or (Item.SubItems(3) = "M") Or (Item.SubItems(3) = "E") Then
            cmdDelete.Enabled = True
            cmdSave.Enabled = True
        Else
            cmdDelete.Enabled = False
            cmdSave.Enabled = False
        End If
    ElseIf ObjMyUser.IsDeveloper Then
        cmdDelete.Enabled = True
        cmdSave.Enabled = True
    End If
End Sub

Private Sub LoadForm()
    Dim Rs As Recordset
    Dim strSQL As String
    Dim Row As Long
    Dim strKey As String
    Dim strDept As String
    
    strDept = GetDept
    
    If strDept = "" Then Exit Sub
            
    strSQL = " select a.deptfg,a.formid,a.formnm, " & _
        " a.readfg as readuse, a.writefg as writeuse,a.printfg as printuse , " & _
        " '' as readval, '' as writeval, '' as printval " & _
        " from " & T_COM007 & " a " & _
        " where a.deptfg in (" & GetDept & ")" & _
        " union " & _
        " select b.deptfg,b.formid,'' as formnm, " & _
        " '' as readuse, '' as writeuse, '' as printuse, " & _
        " b.readfg as readval, b.writefg as writeval, b.printfg as printval " & _
        " from " & T_COM009 & " b " & _
        " where b.deptfg in (" & strDept & ")" & _
        " and " & DBW("groupid=", Trim(txtGroupID.Text))
       
    Set Rs = New Recordset
    
    On Error GoTo Nodata
    
    Rs.Open strSQL, DBConn
    
    tblForm.ReDraw = False
    
    Do Until Rs.EOF
        
        With tblForm
            If strKey <> Rs.Fields("formid").Value & "" Then
                Row = Row + 1
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
                .Row = Row
                .Col = 1: .Value = IIf(Rs.Fields("deptfg").Value & "" = "A", "APS", IIf(Rs.Fields("deptfg").Value & "" = "B", "BBS", "LIS"))
                .Col = 2: .Value = Rs.Fields("formnm").Value & ""
                .Col = 3: .CellType = IIf(Rs.Fields("readuse").Value & "" = "1", CellTypeCheckBox, CellTypeStaticText)
                          .TypeHAlign = TypeHAlignCenter
                .Col = 4: .CellType = IIf(Rs.Fields("writeuse").Value & "" = "1", CellTypeCheckBox, CellTypeStaticText)
                          .TypeHAlign = TypeHAlignCenter
                .Col = 5: .CellType = IIf(Rs.Fields("printuse").Value & "" = "1", CellTypeCheckBox, CellTypeStaticText)
                          .TypeHAlign = TypeHAlignCenter
                .Col = 6: .Value = Rs.Fields("formid").Value & ""
            Else
                .Col = 3: .Value = Rs.Fields("readval").Value & ""
                .Col = 4: .Value = Rs.Fields("writeval").Value & ""
                .Col = 5: .Value = Rs.Fields("printval").Value & ""
            End If
            
            strKey = Rs.Fields("formid").Value & ""
        End With
        
        
        Rs.MoveNext
    Loop
    
    tblForm.ReDraw = True
    
Nodata:
    Set Rs = Nothing

End Sub

Private Function GetDept()
    Dim i As Long
    Dim strTmp As String
    
    For i = chkDeptFg.lbound To chkDeptFg.UBound
        If chkDeptFg(i).Value = 1 Then
            strTmp = strTmp & "'" & chkDeptFg(i).Tag & "',"
        End If
    Next
    If strTmp = "" Then Exit Function
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    
    GetDept = strTmp
End Function

Private Sub tblForm_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Long
    Static lngToggle(1 To 3) As Long
    
    If Col < 3 Then Exit Sub
    If Row > 0 Then Exit Sub
    
    lngToggle(Col - 2) = (lngToggle(Col - 2) + 1) Mod 2
    
    With tblForm
        .Col = Col
        For i = 1 To .DataRowCnt
            .Row = i
            If .CellType = CellTypeCheckBox Then
                .Value = lngToggle(Col - 2)
            End If
        Next
    End With
End Sub

Private Sub txtGroupID_Change()
    If txtGroupNm.Text <> "" Then Call InitForm
End Sub
