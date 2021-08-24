VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPassmgr 
   Caption         =   "사용자관리(Password)"
   ClientHeight    =   6810
   ClientLeft      =   315
   ClientTop       =   1140
   ClientWidth     =   11085
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11085
   Begin FPSpreadADO.fpSpread ssBasPass 
      Height          =   5775
      Left            =   3330
      TabIndex        =   15
      Top             =   360
      Width           =   7620
      _Version        =   196608
      _ExtentX        =   13441
      _ExtentY        =   10186
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   15
      MaxRows         =   499
      ScrollBars      =   2
      SpreadDesigner  =   "frmPassmgr.frx":0000
      Appearance      =   1
      ScrollBarTrack  =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   10186
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbSubPart 
         Height          =   300
         Left            =   1230
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   2790
         Width           =   1695
      End
      Begin VB.ComboBox cmbRank 
         Height          =   300
         ItemData        =   "frmPassmgr.frx":3F5F
         Left            =   1230
         List            =   "frmPassmgr.frx":3F78
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   2475
         Width           =   1695
      End
      Begin VB.ComboBox cmbGrade 
         Height          =   300
         ItemData        =   "frmPassmgr.frx":3F91
         Left            =   1230
         List            =   "frmPassmgr.frx":3FA1
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox cmbClass 
         Height          =   300
         ItemData        =   "frmPassmgr.frx":3FC2
         Left            =   1230
         List            =   "frmPassmgr.frx":3FCC
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   1845
         Width           =   1695
      End
      Begin Threed.SSCommand cmdQry 
         Height          =   555
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "                 조회확인 ▶"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "frmPassmgr.frx":3FDF
      End
      Begin VB.TextBox txtUserId 
         Height          =   300
         Left            =   1230
         MaxLength       =   6
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtUsername 
         Height          =   300
         Left            =   1230
         TabIndex        =   2
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   1155
         Left            =   2160
         TabIndex        =   6
         Top             =   4275
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   2037
         _StockProps     =   78
         Caption         =   "화면정리"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmPassmgr.frx":48B9
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   1155
         Left            =   1260
         TabIndex        =   7
         Top             =   4275
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   2037
         _StockProps     =   78
         Caption         =   "삭제확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmPassmgr.frx":604B
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   1155
         Left            =   300
         TabIndex        =   4
         Top             =   4275
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   2037
         _StockProps     =   78
         Caption         =   "입력확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmPassmgr.frx":6925
      End
      Begin VB.Label Label7 
         Caption         =   "Part"
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   2835
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "Rank"
         Height          =   330
         Left            =   180
         TabIndex        =   21
         Top             =   2565
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "관리구분"
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   2250
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "작업범위"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1935
         Width           =   915
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   555
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   180
         Width           =   2835
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   1155
         Index           =   2
         Left            =   2220
         TabIndex        =   13
         Top             =   4335
         Width           =   795
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   1155
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   4335
         Width           =   795
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   1155
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   4335
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "UserID"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "사용자 이름"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "비밀번호"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1620
         Width           =   990
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmPassmgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then
            Me.Controls(i).Text = ""
        ElseIf TypeOf Me.Controls(i) Is ComboBox Then
            If Me.Controls(i).Style = vbComboDropdownList Then
                Me.Controls(i).ListIndex = -1
            Else
                Me.Controls(i).Text = ""
            End If
        ElseIf TypeOf Me.Controls(i) Is fpSpread Then
            Me.Controls(i).Row = 1
            Me.Controls(i).Row2 = Me.Controls(i).DataRowCnt
            Me.Controls(i).Col = 1
            Me.Controls(i).Col2 = Me.Controls(i).DataColCnt
            Me.Controls(i).BlockMode = True
            Me.Controls(i).Text = ""
            Me.Controls(i).BlockMode = False
        End If
    Next
    
    mdiMain.stbMain.Panels(1).Text = ""
    

End Sub

Private Sub cmdDel_Click()
    
    If Trim(txtUserId.Text) = "" Then
        MsgBox "삭제할 UserID 가 없습니다!.", vbInformation + vbOKOnly, "Please Data Check..."
        Exit Sub
    End If

    sMsg = Trim$(txtUserId.Text) & " = " & Trim(txtUsername.Text) & vbCrLf & "의 Data 를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbYesNo + vbQuestion, "삭제 확인 Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBas_Pass"
    strSql = strSql & " WHERE  idNumber = '" & Trim(txtUserId.Text) & "'"
    strSql = strSql & " AND    DeptCode = 'CP'"
    
    If adoExec(strSql) = True Then
        mdiMain.stbMain.Panels(1).Text = "Data를 삭제하였습니다!.."
        GoSub Left_Clear
    Else
        mdiMain.stbMain.Panels(1).Text = "Data 삭제시 오류가 났습니다!..."
    End If
    
    Call cmdQry_Click
    Exit Sub
    
    
Left_Clear:
    txtUserId.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    cmbClass.ListIndex = -1
    cmbGrade.ListIndex = -1
    cmbRank.ListIndex = -1
    cmbSubPart.ListIndex = -1
    Return

    
    
End Sub

Private Sub cmdInsert_Click()
    Dim sToisaGb        As String * 1
    
    
    If Trim(txtUserId.Text) = "" Then
        MsgBox "입력할 자료가 없습니다!", vbOKOnly, "입력Data Not Found"
        Exit Sub
    End If
    
        
    strSql = ""
    strSql = strSql & " SELECT *                     "
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Pass"
    strSql = strSql & " WHERE  idNumber  =  '" & Trim(txtUserId.Text) & "'"
    strSql = strSql & " AND    DeptCode  =  'CP'"

    If False = adoSetOpen(strSql, adoSet) Then
        GoSub Password_Insert_Sub
    Else
        Call adoSetClose(adoSet)
        GoSub Password_Update_Sub
    End If
    
    
    Exit Sub
    


Password_Insert_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TW_MIS_PMPA.TWBAS_Pass"
    strSql = strSql & "       (Programid, Programname, idNumber, Name, Password,"
    strSql = strSql & "        Class,     Grade,       Sabun,    Part, "
    strSql = strSql & "        DeptCode,  SubClass,    Rank,     SubPart )"
    strSql = strSql & " VALUES(' ', "
    strSql = strSql & "        ' ', "
    strSql = strSql & "        '" & Trim(txtUserId.Text) & "',"
    strSql = strSql & "        '" & Trim(txtUsername.Text) & "',"
    strSql = strSql & "        '" & Trim(txtPassword.Text) & "',"
    strSql = strSql & "        '" & Trim(cmbClass.Text) & "',"
    strSql = strSql & "        '" & Trim(cmbGrade.Text) & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        'CP',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        '" & Trim(cmbRank.Text) & "',"
    strSql = strSql & "        '" & Left(Trim(cmbSubPart.Text), 2) & "')"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        mdiMain.stbMain.Panels(1).Text = "Data 가 신규 입력 되었습니다!.."
        GoSub Left_Clear
        Call cmdQry_Click
    Else
        adoConnect.RollbackTrans
        mdiMain.stbMain.Panels(1).Text = "신규 입력시 오류가 났습니다!..."
    End If
    Return
    
Password_Update_Sub:
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_PMPA.TWBAS_Pass"
    strSql = strSql & " SET    Name        =  '" & Trim(txtUsername.Text) & "',"
    strSql = strSql & "        Password    =  '" & Trim(txtPassword.Text) & "',"
    strSql = strSql & "        Class       =  '" & Trim(cmbClass.Text) & "',"
    strSql = strSql & "        Grade       =  '" & Trim(cmbGrade.Text) & "',"
    strSql = strSql & "        Rank        =  '" & Trim(cmbRank.Text) & "',"
    strSql = strSql & "        SubPart     =  '" & Left(Trim(cmbSubPart.Text), 2) & "'"
    strSql = strSql & " WHERE  idNumber    =  '" & Trim(txtUserId.Text) & "'"
    strSql = strSql & " AND    Deptcode    =  'CP'"
    
    If adoExec(strSql) = True Then
        mdiMain.stbMain.Panels(1).Text = "Data 가 수정 되었습니다!.."
        GoSub Left_Clear
        Call cmdQry_Click
    Else
        mdiMain.stbMain.Panels(1).Text = "Data 수정시 오류가 났습니다!..."
    End If
    
    Return
    
    
Left_Clear:
    txtUserId.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    cmbClass.ListIndex = -1
    cmbGrade.ListIndex = -1
    cmbRank.ListIndex = -1
    cmbSubPart.ListIndex = -1
    Return
End Sub

Private Sub cmdQry_Click()
    
    mdiMain.stbMain.Panels(1).Text = ""
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.ROWID RwID, b.Codenm "
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Pass    a,"
    strSql = strSql & "        TWEXAM_Specode b            "
    strSql = strSql & " WHERE  a.DeptCode  = 'CP'"
    strSql = strSql & " AND    b.Codegu(+) = '12'"
    strSql = strSql & " AND    a.Subpart   = b.Codeky(+)"
    strSql = strSql & " Order  By  a.Rank"
    strSql = strSql & " ORDER BY a.GRADE"
    
    ssBasPass.ReDraw = False
    ssBasPass.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssBasPass.MaxRows = adoSet.RecordCount
    ssBasPass.RowHeight(-1) = 11.5
    
    Do Until adoSet.EOF
        ssBasPass.Row = ssBasPass.DataRowCnt + 1
        ssBasPass.Col = 2:  ssBasPass.Text = adoSet.Fields("RwID").Value & ""
        ssBasPass.Col = 3:  ssBasPass.Text = adoSet.Fields("ProgramID").Value & ""
        ssBasPass.Col = 4:  ssBasPass.Text = adoSet.Fields("Programname").Value & ""
        ssBasPass.Col = 5:  ssBasPass.Text = adoSet.Fields("IdNumber").Value & ""
        ssBasPass.Col = 6:  ssBasPass.Text = adoSet.Fields("Name").Value & ""
        ssBasPass.Col = 7:  ssBasPass.Text = adoSet.Fields("Password").Value & ""
        ssBasPass.Col = 8:  ssBasPass.Text = adoSet.Fields("Class").Value & ""
        ssBasPass.Col = 9:  ssBasPass.Text = adoSet.Fields("Grade").Value & ""
        ssBasPass.Col = 10:  ssBasPass.Text = adoSet.Fields("Sabun").Value & ""
        ssBasPass.Col = 11: ssBasPass.Text = adoSet.Fields("Part").Value & ""
        ssBasPass.Col = 12: ssBasPass.Text = adoSet.Fields("deptcode").Value & ""
        ssBasPass.Col = 13: ssBasPass.Text = adoSet.Fields("SubClass").Value & ""
        ssBasPass.Col = 14: ssBasPass.Text = adoSet.Fields("Rank").Value & ""
        ssBasPass.Col = 15: ssBasPass.Text = adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    ssBasPass.ReDraw = True
    ssBasPass.SetFocus
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Form_Load()
    Me.Top = 1
    Me.Left = 1
    Me.Height = 7000
    Me.Width = 11200
    
    
    GoSub SlipNo_Set            '임상병리과 내 근무 Part Code Get
    
    Exit Sub
    
SlipNo_Set:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  CODEGU = '12'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSubPart.AddItem Trim(adoSet.Fields("Codeky").Value & "") & " " & _
                           Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    cmbSubPart.AddItem " "
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub mnuQuit_Click()
    
    Unload Me
    
End Sub


Private Sub ssBasPass_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Col = 1 Then
        ssBasPass.Row = Row
        ssBasPass.Col = 5: txtUserId.Text = Trim(ssBasPass.Text)
        Call txtUserId_KeyDown(vbKeyReturn, 0)
    End If

End Sub

Private Sub ssBasPass_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        ssBasPass.Row = 1
        ssBasPass.Col = 4
        ssBasPass.Row2 = ssBasPass.DataRowCnt
        ssBasPass.Col2 = ssBasPass.DataColCnt
        ssBasPass.SortBy = SS_SORT_BY_ROW
        ssBasPass.SortKey(1) = Col
        ssBasPass.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        ssBasPass.Action = SS_ACTION_SORT
    End If
    

End Sub



Private Sub TxtPassWord_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    
End Sub


Public Sub txtUserId_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub GET_Password_Data
    End If
    
    Exit Sub
    

GET_Password_Data:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Pass"
    strSql = strSql & " WHERE  DeptCode = 'CP'"
    strSql = strSql & " AND    idNumber = '" & txtUserId.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        txtUsername.Text = ""
        txtPassword.Text = ""
        cmbClass.ListIndex = -1
        cmbGrade.ListIndex = -1
        cmbRank.ListIndex = -1
        cmbSubPart.ListIndex = -1
        Return
    End If
    txtUsername.Text = adoSet.Fields("Name").Value & ""
    txtPassword.Text = adoSet.Fields("Password").Value & ""
    Call SetComboBox(cmbClass, Trim(adoSet.Fields("Class").Value & ""))
    Call SetComboBox(cmbGrade, Trim(adoSet.Fields("Grade").Value & ""))
    Call SetComboBox(cmbRank, Trim(adoSet.Fields("Rank").Value & ""))
    
    Dim sTmp        As String
    sTmp = Trim(adoSet.Fields("SubPart").Value & "")
    
    For i = 0 To cmbSubPart.ListCount - 1
        If Left(Trim(cmbSubPart.List(i)), 2) = Trim(sTmp) Then
            cmbSubPart.ListIndex = i
            Exit For
        End If
    Next
    
    
    Return
    
End Sub

Private Sub txtUsername_GotFocus()
    txtUsername.SelStart = 0
    txtUsername.SelLength = Len(txtUsername.Text)
    
End Sub
