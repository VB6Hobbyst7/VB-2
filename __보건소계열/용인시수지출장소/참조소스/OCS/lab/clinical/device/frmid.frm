VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmIdPass 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 번호 & 비밀 번호 입력ㆍ 변경"
   ClientHeight    =   2148
   ClientLeft      =   1836
   ClientTop       =   1980
   ClientWidth     =   7620
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   11.4
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2148
   ScaleWidth      =   7620
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2160
      Top             =   4350
   End
   Begin Threed.SSPanel Panel 
      Height          =   2100
      Index           =   0
      Left            =   15
      TabIndex        =   8
      Top             =   2175
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   3704
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.6
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.TextBox txtPrePass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   3855
         MaxLength       =   8
         TabIndex        =   5
         Top             =   345
         Width           =   1995
      End
      Begin VB.TextBox txtNewPass 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   3855
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   885
         Width           =   1995
      End
      Begin VB.TextBox txtNewOK 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   3855
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1440
         Width           =   1995
      End
      Begin Threed.SSPanel Panel 
         Height          =   420
         Index           =   3
         Left            =   1710
         TabIndex        =   9
         Top             =   315
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   741
         _StockProps     =   15
         Caption         =   "이전 비 밀 번 호"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel Panel 
         Height          =   420
         Index           =   4
         Left            =   1710
         TabIndex        =   10
         Top             =   1395
         Width           =   1920
         _Version        =   65536
         _ExtentX        =   3387
         _ExtentY        =   741
         _StockProps     =   15
         Caption         =   "새 비밀번호 확인"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel Panel 
         Height          =   465
         Index           =   5
         Left            =   1710
         TabIndex        =   11
         Top             =   825
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "새   비 밀 번 호"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin Threed.SSPanel Panel 
         Height          =   1905
         Index           =   6
         Left            =   6060
         TabIndex        =   12
         Top             =   90
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   3360
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSCommand cmdPCancel 
            Height          =   945
            Left            =   30
            TabIndex        =   14
            Top             =   945
            Width           =   1380
            _Version        =   65536
            _ExtentX        =   2434
            _ExtentY        =   1667
            _StockProps     =   78
            Caption         =   "변경취소[&X]"
            ForeColor       =   8388736
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.6
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            MouseIcon       =   "frmid.frx":0000
            Picture         =   "frmid.frx":0452
         End
         Begin Threed.SSCommand cmdPOk 
            Height          =   945
            Left            =   30
            TabIndex        =   13
            Top             =   15
            Width           =   1380
            _Version        =   65536
            _ExtentX        =   2434
            _ExtentY        =   1667
            _StockProps     =   78
            Caption         =   "변경완료[&T]"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.6
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Picture         =   "frmid.frx":08A4
         End
      End
      Begin VB.Image Image2 
         Height          =   384
         Left            =   576
         Picture         =   "frmid.frx":0BBE
         Top             =   360
         Width           =   384
      End
   End
   Begin Threed.SSPanel Panel 
      Height          =   2100
      Index           =   7
      Left            =   15
      TabIndex        =   15
      Top             =   0
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   3704
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.4
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin Threed.SSPanel PanChanging 
         Height          =   750
         Left            =   2130
         TabIndex        =   18
         Top             =   660
         Visible         =   0   'False
         Width           =   5340
         _Version        =   65536
         _ExtentX        =   9419
         _ExtentY        =   1323
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Image Image4 
            Height          =   675
            Left            =   30
            Picture         =   "frmid.frx":1000
            Stretch         =   -1  'True
            Top             =   30
            Width           =   510
         End
         Begin VB.Image Image3 
            Height          =   675
            Left            =   540
            Picture         =   "frmid.frx":2416
            Stretch         =   -1  'True
            Top             =   30
            Width           =   4755
         End
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   3855
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1470
         Width           =   1995
      End
      Begin VB.TextBox txtIdnumber 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3855
         MaxLength       =   6
         TabIndex        =   0
         Top             =   945
         Width           =   1995
      End
      Begin Threed.SSPanel Panel 
         Height          =   285
         Index           =   1
         Left            =   2325
         TabIndex        =   16
         Top             =   975
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "사용자 번호"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel Panel 
         Height          =   270
         Index           =   2
         Left            =   2325
         TabIndex        =   17
         Top             =   1500
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   476
         _StockProps     =   15
         Caption         =   "비 밀 번 호"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.4
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   510
         Left            =   6210
         TabIndex        =   3
         Top             =   780
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2223
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "취 소[&C]"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   510
         Left            =   6210
         TabIndex        =   2
         Top             =   135
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2223
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "확 인[&O]"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand cmdReplace 
         Height          =   510
         Left            =   6210
         TabIndex        =   4
         Top             =   1455
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2223
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "변 경[&R]"
         ForeColor       =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.6
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin VB.Image Image1 
         Height          =   504
         Left            =   2088
         Picture         =   "frmid.frx":E900
         Top             =   192
         Width           =   3240
      End
      Begin VB.Image Img 
         BorderStyle     =   1  '단일 고정
         Height          =   1830
         Left            =   135
         Picture         =   "frmid.frx":16E7A
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1830
      End
   End
End
Attribute VB_Name = "FrmIdPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTag              As String
Dim strSqlDef           As String
Dim strConnect          As String
Dim strGrade            As String
Dim nPassCount          As Integer

Private Sub cmdPCancel_Click()

    strTag = "T"
    Call Control_Change

    PanChanging.Visible = False

    Me.Height = 2490
    Panel(0).Enabled = False
    Call Text_Clear
    txtPassword.SetFocus
    
End Sub

Private Sub cmdPOk_Click()

    If txtPrePass.Text = "" Or txtNewPass.Text = "" Or txtNewOK.Text = "" Then
        MsgBox "Data를 입력해 주십시요!", vbExclamation, "알림"
        txtPrePass.SetFocus
        Exit Sub
    End If

    
    strSqlDef = "UPDATE TWBAS_PASS "
    strSqlDef = strSqlDef & " SET    PassWord = '" & UCase(Trim(txtNewOK.Text)) & "' "
    strSqlDef = strSqlDef & " WHERE  IDNumber = '" & Trim(txtIdnumber) & "' "
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSqlDef)
    
    If Rowindicator = 0 Then
        MsgBox "사용자ID, PassWord와 일치하는 Data가 없습니다!" _
            & Chr(13) + Chr(10) & "ID와 PassWord를 확인 후 다시 변경하세요." _
            , vbInformation, "알림"
        Call cmdPCancel_Click
        Exit Sub
    End If
    
'    If Result = -1 Then
    If Result Then
        adoConnect.CommitTrans
        MsgBox "비밀 번호를 변경 하였습니다.", vbInformation, "알림"
    Else
        adoConnect.RollbackTrans
        MsgBox "비밀 번호를 변경 실패하였습니다.", vbCritical, "알림"
    End If
    
    strTag = "T"
    Call Control_Change

    PanChanging.Visible = False

    Me.Height = 2490
    Panel(0).Enabled = False
    txtPassword.Text = UCase(Trim(txtNewOK.Text))
    Call Text_Clear
    txtPassword.SetFocus
    
End Sub

Private Sub cmdPOk_GotFocus()

    If Trim(txtNewOK.Text) = "" Then Exit Sub
    
    If UCase(Trim(txtNewPass.Text)) <> UCase(Trim(txtNewOK.Text)) Then
        MsgBox "새 비밀 번호와 확인 비밀 번호가 틀립니다!" & Chr(13) + Chr(10) & _
        "확인 후 다시 입력하여 주십시요.", vbExclamation, "알림"
        txtNewOK.SetFocus
    End If

End Sub

Private Sub cmdReplace_Click()

    If Trim(txtIdnumber.Text) = "" Then
        MsgBox "사용자 번호를 입력 해 주세요!", vbExclamation, "알림"
        txtIdnumber.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtIdnumber.Text) Then
        MsgBox "사용자 번호는 숫자만 가능합니다!", vbExclamation, "확인"
        txtIdnumber.SetFocus
        Exit Sub
    End If
    
    PanChanging.Visible = True
    
    strTag = "F"
    Call Control_Change
    
    FrmIdPass.Height = 4660
    Panel(0).Enabled = True
    
    If txtPassword.Text <> "" Then
        txtPrePass.Text = Trim(txtPassword.Text)
    End If
    
    txtPrePass.SetFocus
    
End Sub

Private Sub Timer1_Timer()

    'Static n        As Integer
    
    'Img.Picture = Img1(Trim(Str(n Mod 31))).Picture
    
    'n = n + 1
    'If n > 30000 Then n = 0
    'DoEvents
    
End Sub

Private Sub PassWordCheck()

    Dim Rs                  As ADODB.Recordset
    
    strConnect = "OK"
    
    If Not IsNumeric(txtIdnumber.Text) Then txtIdnumber.Text = "0"
    
    GoSub Read_Part
    GoSub Read_Grade
    
    Exit Sub
    

'/--------------------------------------------------------------------------------------------/

Read_Part:

    strSqlDef = ""
    strSqlDef = strSqlDef & "SELECT Name, PassWord, Class, SubClass, Grade, Part, SubPart, DeptCode, Rank "
    strSqlDef = strSqlDef & "  FROM TWBAS_PASS "
    strSqlDef = strSqlDef & " WHERE ProgramID = ' ' "
    strSqlDef = strSqlDef & "   AND IDnumber  = '" & Trim(txtIdnumber.Text) & "'"
    
    Result = AdoOpenSet(Rs, strSqlDef)
    
'    If Result <> -1 And Rowindicator <> 0 Then
    If Result = True And Rowindicator <> 0 Then
        GstrPassWord = AdoGetString(Rs, "PassWord", 0)
        GstrPassName = AdoGetString(Rs, "Name", 0)
'        GstrPassRank = AdoGetString(rs, "Rank", 0)
        GstrPassClass = AdoGetString(Rs, "Class", 0)
'        GstrSubClass = AdoGetString(rs, "SubClass", 0)
        GstrPassGrade = AdoGetString(Rs, "Grade", 0)
        GstrPassPart = AdoGetString(Rs, "Part", 0)
'        GstrSubPart = AdoGetString(rs, "SubPart", 0)
        GstrPassDept = AdoGetString(Rs, "DeptCode", 0)
        GstrIdnumber = Format(txtIdnumber.Text, "000000")
        GstrPassIDnumber = Format(txtIdnumber.Text, "000000")
        
        GsExDate = Dual_Date_Get("yyyy-MM-dd")
        
        If txtPassword.Text <> GstrPassWord Then strConnect = "NO2"
'        If GstrPassDept <> "AP" And GstrPassClass <> "ALL" Then strConnect = "NO9"  '추가
    Else
        strConnect = "NO1"
    End If
    
    AdoCloseSet Rs
               
    Return
    

'/--------------------------------------------------------------------------------------------/

Read_Grade:

    If strConnect <> "OK" Then Return
    If Trim(GstrPassProgramID) < "0" Or Trim(GstrPassProgramID) > "z" Then Return
    
    strSqlDef = ""
    strSqlDef = strSqlDef & "SELECT Name, PassWord, Class, SubClass, Grade, Part, SubPart, DeptCode, Rank "
    strSqlDef = strSqlDef & "  FROM TWBAS_PASS "
    strSqlDef = strSqlDef & " WHERE ProgramID = '" & Trim(GstrPassProgramID) & "' "
    strSqlDef = strSqlDef & "   AND IDnumber  = '" & Trim(txtIdnumber.Text) & "'"
    Result = AdoOpenSet(Rs, strSqlDef)
    
    If Result = False Or Rowindicator = 0 Then strConnect = "NO3"
    
    AdoCloseSet Rs

    Return

End Sub

Private Sub CmdCancel_Click()
    
    Call DbAdoDisConnect
    End

End Sub

Private Sub CmdOK_Click()

    If Trim(txtIdnumber.Text) = "" Then
        MsgBox "사용자 번호를 입력 해 주세요!", vbExclamation, "알림"
        txtIdnumber.SetFocus
        Exit Sub
    End If
    
    nPassCount = nPassCount + 1
    
    Call PassWordCheck
    
    If Left(strConnect, 2) = "NO" Then
        If nPassCount > 3 Then
            MsgBox "ID 와 Password를 확인후에 다시 시작하십시요", 48, "경고"
            Call DbAdoDisConnect
            End
        Else
            Select Case Trim(strConnect)
                Case "NO1":     MsgBox "사용자 번호가 틀림니다 !", 48, "확인요망":  txtIdnumber.SetFocus
                Case "NO2":     MsgBox "Password가 틀림니다 !", 48, "확인요망":     txtPassword.SetFocus
                Case "NO9":     MsgBox "해부병리과만 사용할 수 있습니다.", 48, "확인요망":     txtPassword.SetFocus
                Case Else:      MsgBox "이 Program을 사용하실 권한이 없습니다 !", 48, "확인요망"
            End Select
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Call Form_Position
    Me.Height = 2490
    Panel(0).Enabled = False
    
End Sub

Private Sub TxtIdnumber_GotFocus()
    
    txtIdnumber.SelStart = 0
    txtIdnumber.SelLength = Len(txtIdnumber.Text)

End Sub

Private Sub TxtIdnumber_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtNewOK_GotFocus()

    txtNewOK.SelStart = 0
    txtNewOK.SelLength = Len(txtNewOK.Text)

End Sub

Private Sub txtNewOK_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub txtNewPass_GotFocus()

    txtNewPass.SelStart = 0
    txtNewPass.SelLength = Len(txtNewPass.Text)

End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub txtPassword_GotFocus()
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 13) Then
        KeyAscii = 0
        DoEvents
        SendKeys "{TAB}"
        DoEvents
    End If

End Sub

Private Sub txtPassword_LostFocus()
    
    txtPassword.Text = UCase(txtPassword.Text)

End Sub

Private Sub txtprepass_GotFocus()

    txtPrePass.SelStart = 0
    txtPrePass.SelLength = Len(txtPrePass.Text)

End Sub

Private Sub txtprepass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub Form_Position()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 200

End Sub

Private Sub Text_Clear()

    txtPrePass.Text = ""
    txtNewPass.Text = ""
    txtNewOK.Text = ""
    
End Sub

Private Sub Control_Change()

    Select Case strTag
        Case "F"
            txtIdnumber.Enabled = False
            txtPassword.Enabled = False
            CmdOK.Enabled = False
            CmdCancel.Enabled = False
            cmdReplace.Enabled = False
        Case "T"
            txtIdnumber.Enabled = True
            txtPassword.Enabled = True
            CmdOK.Enabled = True
            CmdCancel.Enabled = True
            cmdReplace.Enabled = True
    End Select
    
End Sub
