VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0601 
   Caption         =   "기초자료 - USER"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "FGB0601.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11280
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   3735
      Left            =   390
      OleObjectBlob   =   "FGB0601.frx":030A
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2850
      Width           =   8475
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   390
      TabIndex        =   10
      Top             =   420
      Width           =   8460
      Begin VB.OptionButton optSchOpt 
         BackColor       =   &H00008080&
         Caption         =   "전체내용 한번에 보기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   1
         Left            =   5010
         TabIndex        =   29
         Top             =   1890
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.OptionButton optSchOpt 
         BackColor       =   &H000040C0&
         Caption         =   "계단식 보기(웹게시판처럼)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   28
         Top             =   1890
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.TextBox txtUserCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "01210630"
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txtUserNm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   10  '한글 
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "중앙대학교용산병원 임상병리과"
         Top             =   810
         Width           =   2220
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "kdkdlkd"
         Top             =   1275
         Width           =   2235
      End
      Begin Threed.SSPanel Panel3D7 
         Height          =   375
         Left            =   270
         TabIndex        =   11
         Top             =   810
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "사용자 이름"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   375
         Left            =   270
         TabIndex        =   12
         Top             =   345
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "사용자 코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   375
         Left            =   270
         TabIndex        =   13
         Top             =   1275
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Password"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   65535
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1455
         Left            =   4440
         TabIndex        =   14
         Top             =   270
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   2566
         _StockProps     =   14
         Caption         =   "사용자의 Default 항목 설정"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin VB.TextBox txtSlipCd 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   8  '영문
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "H02"
            Top             =   405
            Width           =   570
         End
         Begin VB.TextBox txtSpcCd 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   8  '영문
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "H02"
            Top             =   885
            Width           =   570
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   390
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "SLIP 코드"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   870
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "검체 코드"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSCommand cmdButtonSpc 
            Height          =   330
            Left            =   2490
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   885
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
            Picture         =   "FGB0601.frx":0755
         End
         Begin Threed.SSCommand cmdButtonSlip 
            Height          =   330
            Left            =   2490
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   405
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
            Picture         =   "FGB0601.frx":0877
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   375
         Left            =   270
         TabIndex        =   27
         Top             =   1830
         Visible         =   0   'False
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "조회 옵션"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   65535
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   9510
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2790
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "삭제 F4"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0601.frx":0999
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   9510
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "조회 F3"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0601.frx":1273
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   9510
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3810
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "종료 ESC"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0601.frx":1B4D
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   9510
      TabIndex        =   5
      Top             =   750
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "등록 F2"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0601.frx":2427
   End
   Begin Threed.SSPanel pnlPWD 
      Height          =   3135
      Left            =   4080
      TabIndex        =   17
      Top             =   2430
      Width           =   5115
      _Version        =   65536
      _ExtentX        =   9022
      _ExtentY        =   5530
      _StockProps     =   15
      BackColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin Threed.SSCommand cmdConfirm 
         Height          =   675
         Left            =   900
         TabIndex        =   23
         Top             =   1980
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1191
         _StockProps     =   78
         Caption         =   "확 인 (&C)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtNowPWD 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   2040
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   22
         Text            =   "88888888888888888888"
         Top             =   1320
         Width           =   2205
      End
      Begin Threed.SSPanel pnlPrePWD 
         Height          =   375
         Left            =   870
         TabIndex        =   20
         Top             =   840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "이전 암호"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtPrePWD 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   2040
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   18
         Text            =   "88888888888888888888"
         Top             =   840
         Width           =   2205
      End
      Begin Threed.SSPanel pnlNowPWD 
         Height          =   375
         Left            =   870
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "새  암호"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   675
         Left            =   2580
         TabIndex        =   24
         Top             =   1980
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1191
         _StockProps     =   78
         Caption         =   "취 소 (&X)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   450
         Width           =   3405
      End
   End
End
Attribute VB_Name = "FGB0601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iEditUserMode%
Dim iCurRow%
Dim iHlpClick%
Dim iSpdClick1%
Dim iSpdClick2%
Dim iExistSlipCd%
Dim iExistSpcCd%

Private Sub AddUser()
    Dim CUser As DCB0101
    Dim sSchOpt$
    
    Set CUser = New DCB0101
        
    If optSchOpt(0).Value = True Then
        sSchOpt = "0"
    ElseIf optSchOpt(1).Value = True Then
        sSchOpt = "1"
    Else
        sSchOpt = "0"
    End If
        
    CUser.Add_USER txtUserCd, txtPassword, txtUserNm, txtSlipCd, txtSpcCd, sSchOpt
    
    If CUser.AdoErrNum = 0 Then
        ViewMsg "등록 작업이 성공적으로 수행되었습니다..."
        
        With spdBaseCode
            .MaxRows = .MaxRows + 1
            Call .SetText(1, .MaxRows, txtUserCd & "")
            Call .SetText(2, .MaxRows, txtUserNm & "")
            Call .SetText(3, .MaxRows, txtPassword & "")
            Call .SetText(4, .MaxRows, txtSlipCd & "")
            Call .SetText(5, .MaxRows, txtSpcCd & "")
            Call .SetText(6, .MaxRows, sSchOpt & "")
            
            optSchOpt(CInt(sSchOpt)).Value = True
        End With
    End If
    
    Set CUser = Nothing
End Sub

Private Sub CompareSlipCd()
    Dim CPart As DCB0101
    Dim i%
    
    Set CPart = New DCB0101
        
    If txtSlipCd = "" Then
        Exit Sub
    End If
    
    CPart.Get_PART Left$(txtSlipCd, 1), Right$(txtSlipCd, 2)
    
    i = CPart.CurItemCnt
    
    If i = 0 Then
        MsgBox "기초자료에 존재하지 않는 SLIP 코드입니다!!" & vbCrLf & _
               "만약 이 SLIP 코드를 사용하려면 기초자료의 PART 설정부분을 바꿔야 합니다."
        Call Txt_Highlight(txtSlipCd)
        Set CPart = Nothing
        iExistSlipCd = 0    '슬립코드가 존재안하면 0
        Exit Sub
    ElseIf i = 1 Then
        If iHlpClick = 1 Then
            Call Txt_Highlight(txtSlipCd)
        Else
            txtSpcCd.SetFocus
        End If
        
        Set CPart = Nothing
        
        iHlpClick = 0
        iExistSlipCd = 1    '슬립코드가 존재하면 1
    ElseIf i > 1 Then
        MsgBox "코드설정에 오류가 있습니다!!"
        Call Txt_Highlight(txtSlipCd)
        Set CPart = Nothing
        Exit Sub
    End If
    
    Txt_Highlight txtSpcCd
End Sub

Private Sub CompareSpcCd()
    Dim CSpecimen As DCB0101
    Dim i%
    
    Set CSpecimen = New DCB0101
    
    CSpecimen.Get_SPC txtSpcCd
    
    i = CSpecimen.CurItemCnt
    
    If i = 0 Then
        MsgBox "기초자료에 존재하지 않는 검체 코드입니다!!" & vbCrLf & _
               "만약 이 검체 코드를 사용하려면 기초자료의 SPECIMEN 설정부분을 바꿔야 합니다."
        Call Txt_Highlight(txtSpcCd)
        Set CSpecimen = Nothing
        iExistSpcCd = 0     '검체코드가 존재안하면 0
        Exit Sub
    ElseIf i = 1 Then
        Set CSpecimen = Nothing
        iExistSpcCd = 1     '검체코드가 존재하면 1
    ElseIf i > 1 Then
        MsgBox "코드설정에 오류가 있습니다!!"
        Set CSpecimen = Nothing
        Exit Sub
    End If
End Sub

Private Sub ComparePwd()
    
End Sub
Private Sub DisplayInit()
    txtUserCd = ""
    txtUserNm = ""
    txtPassword = ""
    txtSlipCd = ""
    txtSpcCd = ""
    
    Me.KeyPreview = True
    
    pnlPWD.Visible = False
    
    'SpreadBackColor Option
    iSpdBackColorOption = 2
    
    With spdBaseCode
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .EditModePermanent = True
        .Protect = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Col2 = 6
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
    End With
    
End Sub

Private Sub cmdButtonSlip_Click()
    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    j = CPart.CurItemCnt
    
    iHlpClick = 1
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtSlipCd.hwnd
    
    FSB0101.Left = 2700
    FSB0101.Top = 1400
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdButtonSpc_Click()
   Dim i%
    Dim j%
    Dim CSpecimen As DCB0101
    Dim sTot01$
    Dim sTot02$
    
    Set CSpecimen = New DCB0101
    CSpecimen.Get_SPC
    j = CSpecimen.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CSpecimen
        sTot01 = .TotField01
        sTot02 = .TotField02
    End With
    
    Set CSpecimen = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtSpcCd.hwnd
    
    FSB0101.Left = 2700
    FSB0101.Top = 1750
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    cmdReg.Enabled = True
    cmdSearch.Enabled = True
    cmdDelete.Enabled = True
    
    txtUserCd.Enabled = True
    txtUserNm.Enabled = True
    txtPassword.Enabled = True
    txtSlipCd.Enabled = True
    txtSpcCd.Enabled = True
    
    spdBaseCode.Enabled = True
    
    pnlPWD.Visible = False
    
    iEditUserMode = 0
End Sub

Private Sub cmdConfirm_Click()
    Dim CUser As DCB0101
    Dim vPrev, vUserCd
    Dim sSchOpt$, sCurUserCd$
    Dim i%
    
    If iEditUserMode = 1 Then
        '암호가 바뀌지 않은 경우
        
        Call spdBaseCode.GetText(3, iCurRow, vPrev)
        
        If CStr(vPrev) = txtPrePWD Then
        Else
            MsgBox "암호가 일치하지 않습니다!!"
            Call Txt_Highlight(txtPrePWD)
            Exit Sub
        End If
    ElseIf iEditUserMode = 2 Then
        '암호가 바뀐 경우
        
        Call spdBaseCode.GetText(3, iCurRow, vPrev)
        
        If CStr(vPrev) = txtPrePWD And txtPassword = txtNowPWD Then
        Else
            MsgBox "암호가 정확하지 않습니다!!"
            txtPrePWD.SetFocus
            Exit Sub
        End If
        
    Else    'Delete Mode
        iCurRow = 0
        
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vUserCd)
            
            If CStr(vUserCd) = txtUserCd Then
                iCurRow = i
                sCurUserCd = txtUserCd
                Exit For
            End If
        Next
                
        If iCurRow = 0 Then
            MsgBox "입력된 사용자 코드와 일치하는 데이터가 없습니다!!"
            Exit Sub
        End If
            
        Call spdBaseCode.GetText(3, iCurRow, vPrev)
        
        If CStr(vPrev) = txtPrePWD Then
        Else
            MsgBox "암호가 일치하지 않습니다!!"
            txtPrePWD.SetFocus
            Exit Sub
        End If
        
        Set CUser = New DCB0101
            
        CUser.Delete_USER txtUserCd
        
        If CUser.AdoErrNum = 0 Then
            With spdBaseCode
                .Row = iCurRow
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
        
            txtUserCd = ""
            txtUserNm = ""
            txtPassword = ""
            txtSlipCd = ""
            txtSpcCd = ""
            optSchOpt(0).Value = False
            optSchOpt(1).Value = False
            
            If sCurUserCd = fCurUserCd Then
            
                MsgBox "현재 프로그램을 사용중인 사용자정보가 모두 지워져 Re-LogIn 해야 합니다!!" & _
                       "다른번호로 LogIn 하십시요!!"
                
                Call HidePrevFrm
                
                Load FSB0301
                FSB0301.Show vbModal
            End If
        End If
        
        Set CUser = Nothing
    
        pnlPWD.Visible = False
        
        cmdReg.Enabled = True
        cmdSearch.Enabled = True
        cmdDelete.Enabled = True
        
        txtUserCd.Enabled = True
        txtUserNm.Enabled = True
        txtPassword.Enabled = True
        txtSlipCd.Enabled = True
        txtSpcCd.Enabled = True
        spdBaseCode.Enabled = True
        
        Exit Sub
    End If
    
    Set CUser = New DCB0101
    
    If optSchOpt(0).Value = True Then
        sSchOpt = "0"
    ElseIf optSchOpt(1).Value = True Then
        sSchOpt = "1"
    End If
    
    CUser.Edit_USER txtUserCd, txtPassword, txtUserNm, txtSlipCd, txtSpcCd, sSchOpt
    
    If CUser.AdoErrNum = 0 Then
        ViewMsg "변경작업이 이루어 졌습니다..."
        
        With spdBaseCode
            For i = 1 To .MaxRows
                Call .GetText(1, i, vUserCd)
                
                If CStr(vUserCd) = txtUserCd Then
                    Call .SetText(1, i, txtUserCd & "")
                    Call .SetText(2, i, txtUserNm & "")
                    Call .SetText(3, i, txtPassword & "")
                    Call .SetText(4, i, txtSlipCd & "")
                    Call .SetText(5, i, txtSpcCd & "")
                    Call .SetText(6, i, sSchOpt & "")
                    
                    Exit For
                End If
            Next
        End With
        
        pnlPWD.Visible = False
        
        cmdReg.Enabled = True
        cmdSearch.Enabled = True
        cmdDelete.Enabled = True
        
        txtUserCd.Enabled = True
        txtUserNm.Enabled = True
        txtPassword.Enabled = True
        txtSlipCd.Enabled = True
        txtSpcCd.Enabled = True
        spdBaseCode.Enabled = True
        
        If CStr(vUserCd) = fCurUserCd Then
            
            MsgBox "현재 프로그램을 사용중인 사용자정보가 바뀌어 Re-LogIn 해야 합니다!!"
            
            Call HidePrevFrm
            
            Load FSB0301
            FSB0301.Show vbModal
        End If
    End If
    
    Set CUser = Nothing
    
End Sub

Private Sub cmdDelete_Click()
    Dim i%, j%
    Dim vUserCd
    Dim iRetVal%
    
    If txtUserCd = "" Then
        Exit Sub
    Else
        j = 0
        
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vUserCd)
            
            If CStr(vUserCd) = txtUserCd Then
                j = i
                Exit For
            End If
        Next
                
        If j = 0 Then
            MsgBox "입력된 사용자 코드와 일치하는 데이터가 없습니다!!"
            Exit Sub
        End If
                
        iRetVal = MsgBox("사용자 코드 : " & txtUserCd & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "사용자 코드 삭제 확인")
        
        If iRetVal = 1 Then
            iEditUserMode = 3   'Delete Mode
            
            pnlPWD.Visible = True
            txtPrePWD = ""
            
            lblMsg = "암호 확인을 해주십시요!!"
            pnlPrePWD = "암  호"
            
            pnlNowPWD.Visible = False
            txtNowPWD.Visible = False
            
            txtPrePWD.SetFocus
        Else
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Dim vUserCd
    Dim vUserNm
    Dim vPwd
    Dim vSlipCd
    Dim vSpcCd
    Dim vSchOpt
    Dim i%
    Dim bExist As Boolean
    
    bExist = False
    
    If Len(txtUserCd) = 0 Or Len(txtUserNm) = 0 Or Len(txtPassword) = 0 Or Len(txtSlipCd) = 0 Or Len(txtSpcCd) = 0 Then
        MsgBox "사용자 코드, 사용자 이름, Password, SLIP 코드, 검체 코드에 모두 입력이 되어야 합니다!!"
        Exit Sub
    End If
    
    If Len(txtSlipCd) = 3 And Len(txtSpcCd) = 3 And iExistSlipCd = 1 And iExistSpcCd = 1 Then
    Else
        ViewMsg "존재하지 않는 SLIP 코드이거나 검체코드입니다..."
        Exit Sub
    End If
    
    With spdBaseCode
        For i = 1 To .MaxRows
            Call .GetText(1, i, vUserCd)
            Call .GetText(2, i, vUserNm)
            Call .GetText(3, i, vPwd)
            Call .GetText(4, i, vSlipCd)
            Call .GetText(5, i, vSpcCd)
            Call .GetText(6, i, vSchOpt)
            
            If CStr(vUserCd) = txtUserCd And CStr(vUserNm) = txtUserNm And _
               CStr(vPwd) = txtPassword And CStr(vSlipCd) = txtSlipCd And _
               CStr(vSpcCd) = txtSpcCd And optSchOpt(CInt(vSchOpt)).Value = True Then
                
                MsgBox "이미 존재하는 데이터입니다!!"
                Exit Sub
            ElseIf CStr(vUserCd) = txtUserCd Then
                
                bExist = True
                iCurRow = i
                Exit For
            End If
        Next
    End With
    
    If bExist = False Then
        Call AddUser
    ElseIf bExist = True Then
        If txtPassword = vPwd Then
            '암호가 바뀌지 않은 경우
            iEditUserMode = 1
            
            pnlPWD.Visible = True
            txtPrePWD = ""
            
            lblMsg = "암호 확인을 해주십시요!!"
            pnlPrePWD = "암  호"
            
            pnlNowPWD.Visible = False
            txtNowPWD.Visible = False
            
            txtPrePWD.SetFocus
        Else
            '암호가 바뀐경우
            iEditUserMode = 2
            
            pnlPWD.Visible = True
            pnlPrePWD.Visible = True
            pnlNowPWD.Visible = True
            txtNowPWD.Visible = True
            
            txtPrePWD = ""
            txtNowPWD = ""
            
            lblMsg = "이전암호와 새암호를 입력하세요!!"
            pnlPrePWD = "이전 암호"
            pnlNowPWD = "새  암호"
            
            txtPrePWD.SetFocus
        End If
        
        cmdReg.Enabled = False
        cmdSearch.Enabled = False
        cmdDelete.Enabled = False
        
        txtUserCd.Enabled = False
        txtUserNm.Enabled = False
        txtPassword.Enabled = False
        txtSlipCd.Enabled = False
        txtSpcCd.Enabled = False
        spdBaseCode.Enabled = False
    End If
    
End Sub

Private Sub cmdSearch_Click()
    Dim CUser As DCB0101
    Dim j%, i%
    Dim sField01$, sField02$, sField03$, sField04$, sField05$, sField06$
    
    Set CUser = New DCB0101
    
    CUser.Get_USER
    
    i = CUser.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 사용자가 등록되어 있지 않습니다!!"
        Set CUser = Nothing
        Exit Sub
    End If
    
    'USERCD
    'PASSWORD
    'USERNM
    'SLIPCD
    'SPECIMENCD
    
    sField01 = CUser.TotField01
    sField02 = CUser.TotField02
    sField03 = CUser.TotField03
    sField04 = CUser.TotField04
    sField05 = CUser.TotField05
    sField06 = CUser.TotField06
    
    For j = 1 To i
        spdBaseCode.MaxRows = j
        Call spdBaseCode.SetText(1, j, GetByOne(sField01, sField01) & "")
        Call spdBaseCode.SetText(2, j, GetByOne(sField03, sField03) & "")
        Call spdBaseCode.SetText(3, j, GetByOne(sField02, sField02) & "")
        Call spdBaseCode.SetText(4, j, GetByOne(sField04, sField04) & "")
        Call spdBaseCode.SetText(5, j, GetByOne(sField05, sField05) & "")
        Call spdBaseCode.SetText(6, j, GetByOne(sField06, sField06) & "")
    Next
        
    Set CUser = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'Case vbKeyF1:        Call cmdButtonPart_Click
        Case vbKeyF2:        Call cmdReg_Click
        Case vbKeyF3:        Call cmdSearch_Click
        Case vbKeyF4:        Call cmdDelete_Click
        Case vbKeyEscape:    Call cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim iDigUserCd%
    
    If Me.StartUpPosition = 2 Then
    Else
        Me.Left = 250
        Me.Top = 10
        Me.Width = 11400
        Me.Height = 7500
    End If
    
    iDigUserCd = fDigUserCd
    
    txtUserCd.MaxLength = iDigUserCd
    
    '초기화 - Module Variable
    iEditUserMode = 0
    iHlpClick = 0
    iSpdClick1 = 0
    iSpdClick2 = 0
    iExistSlipCd = 0
    iExistSpcCd = 0
    
    Call DisplayInit
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vField01
    Dim vField02
    Dim vField03
    Dim vField04
    Dim vField05
    Dim vField06
    
    If Row = 0 Then
        Exit Sub
    End If
    
    iSpdClick1 = 1
    iSpdClick2 = 1
    
    iExistSlipCd = 1
    iExistSpcCd = 1
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    Call spdBaseCode.GetText(3, Row, vField03)
    Call spdBaseCode.GetText(4, Row, vField04)
    Call spdBaseCode.GetText(5, Row, vField05)
    Call spdBaseCode.GetText(6, Row, vField06)
    
    txtUserCd = CStr(vField01)
    txtUserNm = CStr(vField02)
    txtPassword = CStr(vField03)
    txtSlipCd = CStr(vField04)
    txtSpcCd = CStr(vField05)
    
    optSchOpt(CInt(vField06)).Value = True
    
End Sub

Private Sub txtNowPWD_Click()
    Call Txt_Highlight(txtNowPWD)
End Sub

Private Sub txtNowPWD_GotFocus()
    Call Txt_Highlight(txtNowPWD)
End Sub

Private Sub txtNowPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdConfirm.SetFocus
    End If
End Sub

Private Sub txtPassword_Click()
    Call Txt_Highlight(txtPassword)
End Sub

Private Sub txtPassword_GotFocus()
    Call Txt_Highlight(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSlipCd.SetFocus
    End If
End Sub

Private Sub txtPrePWD_Click()
    Call Txt_Highlight(txtPrePWD)
End Sub

Private Sub txtPrePWD_GotFocus()
    Call Txt_Highlight(txtPrePWD)
End Sub

Private Sub txtPrePWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If txtNowPWD.Visible = True Then
            txtNowPWD.SetFocus
        Else
            cmdConfirm.SetFocus
        End If
    End If
End Sub

Private Sub txtSlipCd_Change()
    If Len(txtSlipCd) = txtSlipCd.MaxLength Then
        If iSpdClick1 = 1 Then
        Else
            Call CompareSlipCd
        End If
    End If
    
    iSpdClick1 = 0
End Sub

Private Sub txtSlipCd_Click()
    Call Txt_Highlight(txtSlipCd)
End Sub

Private Sub txtSlipCd_GotFocus()
    Call Txt_Highlight(txtSlipCd)
End Sub

Private Sub txtSlipCd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonSlip_Click
    End Select
End Sub

Private Sub txtSlipCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSpcCd.SetFocus
    End If
End Sub

Private Sub txtSpcCd_Change()
    If Len(txtSpcCd) = txtSpcCd.MaxLength Then
        If iSpdClick2 = 1 Then
        Else
            Call CompareSpcCd
        End If
    End If
    
    iSpdClick2 = 0
End Sub

Private Sub txtSpcCd_Click()
    Call Txt_Highlight(txtSpcCd)
End Sub

Private Sub txtSpcCd_GotFocus()
    Call Txt_Highlight(txtSpcCd)
End Sub

Private Sub txtSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonSpc_Click
    End Select
End Sub

Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtUserCd.SetFocus
    End If
End Sub

Private Sub txtSpcCd_LostFocus()
    If Len(txtSpcCd) < txtSpcCd.MaxLength Then
        txtSpcCd = Format$(txtSpcCd, "000")
    End If
End Sub

Private Sub txtUserCd_Change()
    If Len(txtUserCd) = txtUserCd.MaxLength Then
        txtUserNm.SetFocus
    End If
End Sub

Private Sub txtUserCd_Click()
    Call Txt_Highlight(txtUserCd)
End Sub

Private Sub txtUserCd_GotFocus()
    Call Txt_Highlight(txtUserCd)
End Sub

Private Sub txtUserCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtUserNm.SetFocus
    End If
End Sub

Private Sub txtUserNm_Click()
    Call Txt_Highlight(txtUserNm)
End Sub

Private Sub txtUserNm_GotFocus()
    Call Txt_Highlight(txtUserNm)
End Sub

Private Sub txtUserNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPassword.SetFocus
    End If
End Sub
