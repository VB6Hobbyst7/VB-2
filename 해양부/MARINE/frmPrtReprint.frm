VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPrtReprint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "라벨 재출력"
   ClientHeight    =   12885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12885
   ScaleWidth      =   22275
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   10635
      Left            =   90
      TabIndex        =   3
      Top             =   1980
      Width           =   22035
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3105
         Left            =   13320
         TabIndex        =   37
         Top             =   4050
         Visible         =   0   'False
         Width           =   6555
         Begin VB.ComboBox cboSlittingNo 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "frmPrtReprint.frx":0000
            Left            =   3210
            List            =   "frmPrtReprint.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   1230
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.ComboBox cboSchProd 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1830
            Style           =   2  '드롭다운 목록
            TabIndex        =   38
            Top             =   870
            Visible         =   0   'False
            Width           =   2685
         End
         Begin MSCommLib.MSComm comEqp 
            Left            =   2250
            Top             =   180
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1710
            Top             =   300
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   9
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":0004
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":059E
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":0B38
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":10D2
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":1964
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":1ABE
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":1C18
                  Key             =   "NOF"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":1D72
                  Key             =   "ON"
                  Object.Tag             =   "OFF"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPrtReprint.frx":264C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "▶ Slitting 작업번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   780
            TabIndex        =   41
            Top             =   1350
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   "▶ 제품명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   780
            TabIndex        =   40
            Top             =   930
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin FPSpread.vaSpread spdPrtReprint 
         Height          =   10335
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   21735
         _Version        =   393216
         _ExtentX        =   38338
         _ExtentY        =   18230
         _StockProps     =   64
         ColsFrozen      =   7
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   25
         MaxRows         =   20
         RestrictCols    =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtReprint.frx":2F26
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   22065
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   630
         TabIndex        =   8
         Top             =   120
         Width           =   7065
         Begin VB.Frame Frame5 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   1620
            TabIndex        =   18
            Top             =   990
            Visible         =   0   'False
            Width           =   5175
            Begin VB.CommandButton cmdFilter 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "필터"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3750
               Style           =   1  '그래픽
               TabIndex        =   25
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txtPToNo 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2820
               MaxLength       =   3
               TabIndex        =   20
               Text            =   "143"
               Top             =   120
               Width           =   480
            End
            Begin VB.TextBox txtPFrNo 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1710
               MaxLength       =   3
               TabIndex        =   19
               Text            =   "101"
               Top             =   120
               Width           =   540
            End
            Begin VB.Label Label7 
               BackStyle       =   0  '투명
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   2580
               TabIndex        =   24
               Top             =   150
               Width           =   255
            End
            Begin VB.Label Label3 
               BackStyle       =   0  '투명
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1470
               TabIndex        =   23
               Top             =   180
               Width           =   255
            End
            Begin VB.Label Label6 
               Alignment       =   2  '가운데 맞춤
               BackStyle       =   0  '투명
               Caption         =   "~"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   2280
               TabIndex        =   22
               Top             =   210
               Width           =   195
            End
            Begin VB.Label Label2 
               BackStyle       =   0  '투명
               Caption         =   "▶ P No"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   21
               Top             =   150
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00E0E0E0&
            Caption         =   "화면정리"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5700
            Style           =   1  '그래픽
            TabIndex        =   16
            ToolTipText     =   "현재화면을 모두 지웁니다"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00E0E0E0&
            Caption         =   "조회"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4560
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   210
            Width           =   1095
         End
         Begin VB.ComboBox cboProd 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   600
            Width           =   2685
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   375
            Left            =   1650
            TabIndex        =   9
            Top             =   180
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   137232384
            CurrentDate     =   43884
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "▶ 제품명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   12
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "▶ 생산일자 "
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   10
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1605
         Left            =   8400
         TabIndex        =   5
         Top             =   120
         Width           =   9435
         Begin VB.Frame Frame7 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1650
            TabIndex        =   31
            Top             =   990
            Width           =   6015
            Begin VB.OptionButton optICEPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "모두 출력"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   36
               Top             =   180
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optICEPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "전용라벨출력"
               Height          =   225
               Index           =   3
               Left            =   4320
               TabIndex        =   34
               Top             =   180
               Width           =   1485
            End
            Begin VB.OptionButton optICEPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "메인출력"
               Height          =   225
               Index           =   1
               Left            =   1470
               TabIndex        =   33
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton optICEPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "자재코드출력"
               Height          =   225
               Index           =   2
               Left            =   2730
               TabIndex        =   32
               Top             =   180
               Width           =   1395
            End
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   1650
            TabIndex        =   27
            Top             =   540
            Width           =   6015
            Begin VB.OptionButton optPPPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "모두 출력"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   0
               Left            =   180
               TabIndex        =   35
               Top             =   120
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optPPPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "측면 출력 (408G는 내부용 라벨)"
               Height          =   285
               Index           =   2
               Left            =   2730
               TabIndex        =   29
               Top             =   120
               Width           =   3255
            End
            Begin VB.OptionButton optPPPrint 
               BackColor       =   &H00FFFFFF&
               Caption         =   "메인 출력"
               Height          =   285
               Index           =   1
               Left            =   1470
               TabIndex        =   28
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdPrint 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "출력"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   8010
            Style           =   1  '그래픽
            TabIndex        =   17
            Top             =   540
            Width           =   1095
         End
         Begin VB.ComboBox cboLabelType 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "frmPrtReprint.frx":4244
            Left            =   1650
            List            =   "frmPrtReprint.frx":4246
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label Label13 
            BackStyle       =   0  '투명
            Caption         =   "▶ ICE Box"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   30
            Top             =   1170
            Width           =   1155
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "▶ PP Box"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   26
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "▶ 구분"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   7
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   19950
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   " 출   력   구   분"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   7980
         TabIndex        =   14
         Top             =   210
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   " 결   과   조   회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   210
         TabIndex        =   13
         Top             =   210
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   19200
         Picture         =   "frmPrtReprint.frx":4248
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '투명
         Caption         =   "Com1 연결성공"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   19590
         TabIndex        =   4
         Top             =   480
         Width           =   1965
      End
   End
End
Attribute VB_Name = "frmPrtReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmPrtReprint.frm
'   작성자  : 오세원
'   내  용  : 라벨 재출력
'   작성일  : 2020-03-03
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cboLabelType_Click()

    'ppBOX ,ICEBOX   ETX 로 구분하여 나눈다

End Sub

Private Sub cboProd_Click()
    Dim strCompCd    As String
    
'    txtProdCd.Text = Trim(mGetP(cboProd.Text, 2, "|"))
'    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))
    
'    Call GetComp_CodeName(txtProd.Text)
    
'    spdRegOrderDetail.MaxRows = 0
    
End Sub

'
'Private Sub cmdAdd_Click()
'    Dim pAdoRS      As ADODB.Recordset
'    Dim intRow      As Integer
'    Dim intNum      As Integer
'    Dim intMaxNum   As Integer
'
'    With spdRegOrderDetail
'        .MaxRows = .MaxRows + 1
'        Call SetText(spdRegOrderDetail, dtpProdOrderDt.Value, .MaxRows, 1)
'        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 2)
'        Call SetText(spdRegOrderDetail, txtProdCd.Text, .MaxRows, 3)
'        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 4)
'        Call SetText(spdRegOrderDetail, CStr(.MaxRows), .MaxRows, 5)
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 6)
'        .Row = .MaxRows
'        .Col = 6
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 300
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'
''        Call SetText(spdRegOrderDetail, "P" & CStr(.MaxRows), .MaxRows, 7)
''        .Row = .MaxRows
''        .Col = 7
''        .CellType = CellTypeEdit
''        .TypeMaxEditLen = 2
''        .TypeHAlign = TypeHAlignCenter
''        .TypeVAlign = TypeVAlignCenter
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 7)
'        .Row = .MaxRows
'        .Col = 7
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 4
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 8)
'        .Row = .MaxRows
'        .Col = 8
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 4
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'    End With
'
'
'
'
'End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReprint.MaxRows = 0

    dtpDate.Value = Now

    cboSlittingNo.Clear
    For i = 1 To 10
        cboSlittingNo.AddItem CStr(i)
    Next
    cboSlittingNo.ListIndex = 0
    
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub



'제품 리스트 가져오기(조회용)
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset

    Set pAdoRS = New ADODB.Recordset

    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)

    cboProd.Clear

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboProd.AddItem "전체" & Space(50) & "|전체"

        Do Until pAdoRS.EOF
            cboProd.AddItem pAdoRS.Fields("PROD_PRT_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop

        If pAdoRS.RecordCount > 0 Then
            cboProd.ListIndex = 0
        End If
    End If

    pAdoRS.Close

End Sub
    
    
''제품 리스트 가져오기(등록용)
'Private Sub GetProdList_CodeName_Reg(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
'
'    cboProdCd.Clear
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until pAdoRS.EOF
'            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboProdCd.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub
    
'' 작업지시서 리스트 가져옴
Private Sub GetPrintList(ByVal pOrderDate As String, ByVal pProdCd As String, ByVal pLabelType As String, ByVal pSltNo As String, ByVal pPFrNo As String, ByVal pPToNo As String)

    Dim strLabelType    As String

    Set AdoRs = Get_PrintList(pOrderDate, pProdCd, pLabelType, pSltNo, pPFrNo, pPToNo)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        With spdPrtReprint
            
            Do Until AdoRs.EOF
                .MaxRows = .MaxRows + 1
                '.FontSize = 8
                '.FontBold = False
                '.RowHeight(.MaxRows) = 12
                Call SetText(spdPrtReprint, "0", .MaxRows, 1)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReprint, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_PRT_NAME").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReprint, AdoRs.Fields("COMP_VIEW").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_REEL_BAR").Value & "", .MaxRows, 6)
                
                Select Case AdoRs.Fields("PROD_CD").Value & ""
                Case "P0001", "P0002":                      Call SetText(spdPrtReprint, Mid(AdoRs.Fields("PROD_REEL_BAR").Value & "", 15, 3), .MaxRows, 7)
                Case "P0003":                               Call SetText(spdPrtReprint, Mid(AdoRs.Fields("PROD_REEL_BAR").Value & "", 19, 4), .MaxRows, 7)
                Case "P0004", "P0005", "P0008", "P0009":    Call SetText(spdPrtReprint, Mid(AdoRs.Fields("PROD_REEL_BAR").Value & "", 25, 3), .MaxRows, 7)
                Case "P0006", "P0007", "P0010":             Call SetText(spdPrtReprint, Mid(AdoRs.Fields("PROD_REEL_BAR").Value & "", 13, 3), .MaxRows, 7)
                End Select
                'Call SetText(spdPrtReprint, AdoRs.Fields("PNO").Value & "", .MaxRows, 8)
                
                
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_PP_BAR").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_PP_BAR_IN").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_ICE_BAR").Value & "", .MaxRows, 10)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_ICE_BAR_IN").Value & "", .MaxRows, 11)
                'Call SetText(spdPrtReprint, AdoRs.Fields("USER_NAME").Value & "", .MaxRows, 13)
                
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_ID_R").Value & "", .MaxRows, 12)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_DT_R").Value & "", .MaxRows, 13)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_ID_P").Value & "", .MaxRows, 14)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_DT_P").Value & "", .MaxRows, 15)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_ID_I").Value & "", .MaxRows, 16)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_DT_I").Value & "", .MaxRows, 17)
                
                Call SetText(spdPrtReprint, AdoRs.Fields("REEL_PRT_VAL").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReprint, AdoRs.Fields("PP_PRT_VAL").Value & "", .MaxRows, 19)
                Call SetText(spdPrtReprint, AdoRs.Fields("ICE_PRT_VAL").Value & "", .MaxRows, 20)
                
                AdoRs.MoveNext
            Loop
            AdoRs.Close
        End With
        
    End If

End Sub


Private Sub cmdFilter_Click()
    
    Dim i As Integer
    
    With spdPrtReprint
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = 7
            If txtPFrNo.Text <= .Text And txtPToNo >= .Text Then
                'cboSchProd.AddItem .Text
            Else
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1
            End If
'            strFilter = .Text
        Next
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim i           As Integer
    Dim j           As Integer
    
    Dim strPrtData  As String
    Dim strICEData  As Variant
    Dim intCnt      As Integer
    
'                Call SetText(spdPrtReprint, AdoRs.Fields("REEL_PRT_VAL").Value & "", .MaxRows, 12)
'                Call SetText(spdPrtReprint, AdoRs.Fields("PP_PRT_VAL").Value & "", .MaxRows, 13)
'                Call SetText(spdPrtReprint, AdoRs.Fields("ICE_PRT_VAL").Value & "", .MaxRows, 14)
    
    With spdPrtReprint
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Value = "1" Then
                Select Case UCase(Mid(cboLabelType.Text, 1, 1))
                    Case "R"
                            strPrtData = GetText(spdPrtReprint, i, 18)
                            comEqp.Output = strPrtData
                    Case "P"
                            strPrtData = GetText(spdPrtReprint, i, 19)
                            
                            If optPPPrint(0).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    comEqp.Output = Mid(strPrtData, 1, InStr(strPrtData, ETX) - 1)
                                    comEqp.Output = Mid(strPrtData, InStr(strPrtData, ETX) + 1)
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            ElseIf optPPPrint(1).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    comEqp.Output = Mid(strPrtData, 1, InStr(strPrtData, ETX) - 1)
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            ElseIf optPPPrint(2).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    comEqp.Output = Mid(strPrtData, InStr(strPrtData, ETX) + 1)
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            End If
                            
                    Case "I"
                            strPrtData = GetText(spdPrtReprint, i, 20)
                            strICEData = Split(strPrtData, ETX)
                            
                            If optICEPrint(0).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    For intCnt = 0 To UBound(strICEData)
                                        If intCnt = 0 Then
                                            For j = 1 To 3
                                                comEqp.Output = CStr(strICEData(intCnt))
                                            Next
                                        Else
                                            comEqp.Output = CStr(strICEData(intCnt))
                                        End If
                                    Next
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            ElseIf optICEPrint(1).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    For j = 1 To 3
                                        comEqp.Output = CStr(strICEData(0))
                                    Next
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            ElseIf optICEPrint(2).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    comEqp.Output = CStr(strICEData(1))
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            ElseIf optICEPrint(3).Value = True Then
                                If InStr(strPrtData, ETX) > 0 Then
                                    comEqp.Output = CStr(strICEData(2))
                                Else
                                    comEqp.Output = strPrtData
                                End If
                            End If

                End Select
                .Row = i
                .Col = 1
                .Value = "0"
                .BackColor = vbYellow
            End If
        Next
    End With
    
End Sub

Private Sub cmdSearch_Click()
    Dim strDate         As String
    Dim strProdCd       As String
    Dim strLabelType    As String
    Dim strSltNo        As String
    Dim strFrNo         As String
    Dim strToNo         As String
    Dim i               As Integer
    Dim strFilter       As String
    
    spdPrtReprint.MaxRows = 0
    strDate = Format(dtpDate.Value, "yyyymmdd")
    strProdCd = mGetP(cboProd.Text, 2, "|")
    strLabelType = Mid(cboLabelType, 1, 1)
    strSltNo = cboSlittingNo.Text
    strFrNo = txtPFrNo.Text
    strToNo = txtPToNo.Text
    
    
    ' 포장코드 리스트 가져오기
    Call GetPrintList(strDate, strProdCd, strLabelType, strSltNo, strFrNo, strToNo)

    
'    cboSchProd.Clear
'    With spdPrtReprint
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 4
'            If strFilter <> .Text Then
'                cboSchProd.AddItem .Text
'            End If
'            strFilter = .Text
'        Next
'    End With
    
End Sub


Private Sub Form_Load()

    Unload frmPrtReel
    Unload frmPrtPP
    Unload frmPrtICE
    
    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
    '고객사 리스트 가져오기
'    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    ' 포장코드 리스트 가져오기
'    Call GetPackList
    
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle
    
'    If frmPrtReel.comEqp.PortOpen = True Then
'        frmPrtReel.comEqp.PortOpen = False
'    End If
'    If frmPrtPP.comEqp.PortOpen = True Then
'        frmPrtPP.comEqp.PortOpen = False
'    End If
'    If frmPrtICE.comEqp.PortOpen = True Then
'        frmPrtICE.comEqp.PortOpen = False
'    End If

    comEqp.CommPort = gComm.COMPORT
    comEqp.RTSEnable = gComm.RTSEnable
    comEqp.DTREnable = gComm.DTREnable
    comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            
            Resume Next
        Else
            
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & "Public Sub OpenCommunication()" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If


End Sub

'Private Sub GetPackList()
'    Dim pAdoRS      As ADODB.Recordset
'    Dim strPackInfo As String
'
'    Set pAdoRS = Get_PackList
'
'    cboPackCd.Clear
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until pAdoRS.EOF
'            ' PACK_CAT_WIDTH,PACK_PRO_WIDTH,PACK_PRO_LENGTH
'            strPackInfo = pAdoRS.Fields("PACK_CORE").Value & "x" & pAdoRS.Fields("PACK_DIA").Value & " " & pAdoRS.Fields("PACK_CAT_WIDTH").Value & " " & pAdoRS.Fields("PACK_PRO_WIDTH").Value
'
'            cboPackCd.AddItem pAdoRS.Fields("PACK_NAME").Value & Space(3) & strPackInfo & Space(20) & "|" & pAdoRS.Fields("PACK_CD").Value & Space(3)
'            pAdoRS.MoveNext
'        Loop
'
'    End If
'
'    If pAdoRS.RecordCount > 0 Then
'        cboPackCd.ListIndex = 0
'    End If
'
'    pAdoRS.Close
'
'End Sub



Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_Name(pCompCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function

'-- 상단 고객사리스트 가져오기
'Private Sub GetCompList_CodeName()
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_CompList_CodeName
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        cboCompCd.Clear
'
'        Do Until pAdoRS.EOF
'            cboCompCd.AddItem pAdoRS.Fields("COMP_VIEW").Value & Space(1) & ":" & Space(1) & pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
'
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboCompCd.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub

'-- 제품선택했을때 해당 고객사 가져오기
'Private Sub GetComp_CodeName(ByVal pProdCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'    Dim i           As Integer
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_Comp_CodeName(pProdCd)
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        txtProdLen.Text = ""
'
'        Do Until pAdoRS.EOF
'            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
'            For i = 0 To cboCompCd.ListCount
'                If pAdoRS.Fields("COMP_CD").Value & "" = mGetP(cboCompCd.List(i), 2, "|") Then
'                    cboCompCd.ListIndex = i
'                    Exit For
'                End If
'            Next
'            pAdoRS.MoveNext
'        Loop
'
'        pAdoRS.Close
'
'    End If
'
'
'End Sub


'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdPrtReprint
        Call SetText(spdPrtReprint, "선택", 0, 1):                  .ColWidth(1) = 4
        Call SetText(spdPrtReprint, "생산LotNo", 0, 2):             .ColWidth(2) = 15
        Call SetText(spdPrtReprint, "생산일자", 0, 3):              .ColWidth(3) = 10
        Call SetText(spdPrtReprint, "제품명", 0, 4):                .ColWidth(4) = 15
        Call SetText(spdPrtReprint, "고객사", 0, 5):                .ColWidth(5) = 6
        Call SetText(spdPrtReprint, "Reel바코드", 0, 6):            .ColWidth(6) = 20
        Call SetText(spdPrtReprint, "P No", 0, 7):                  .ColWidth(7) = 5
        Call SetText(spdPrtReprint, "PP BOX 바코드", 0, 8):         .ColWidth(8) = 20
        Call SetText(spdPrtReprint, "PP BOX 내부바코드", 0, 9):     .ColWidth(9) = 20
        Call SetText(spdPrtReprint, "ICE BOX 바코드", 0, 10):       .ColWidth(10) = 20
        Call SetText(spdPrtReprint, "ICE BOX 내부바코드", 0, 11):   .ColWidth(11) = 20
        Call SetText(spdPrtReprint, "R 출력자", 0, 12):          .ColWidth(12) = 8
        Call SetText(spdPrtReprint, "R 출력시간", 0, 13):        .ColWidth(13) = 10
        Call SetText(spdPrtReprint, "P 출력자", 0, 14):        .ColWidth(14) = 8
        Call SetText(spdPrtReprint, "P 출력시간", 0, 15):      .ColWidth(15) = 10
        Call SetText(spdPrtReprint, "I 출력자", 0, 16):       .ColWidth(16) = 8
        Call SetText(spdPrtReprint, "I 출력시간", 0, 17):      .ColWidth(17) = 10
'        .ColWidth(12) = 0
'        .ColWidth(13) = 0
'        .ColWidth(14) = 0
'        .ColWidth(15) = 0
'        .ColWidth(16) = 0
'        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .MaxRows = 0
    End With
    
    dtpDate.Value = Now
    cboSlittingNo.Clear
        cboSlittingNo.AddItem "전체"
    For i = 1 To 10
        cboSlittingNo.AddItem CStr(i)
    Next
    cboSlittingNo.ListIndex = 0
    
    cboLabelType.Clear
    cboLabelType.AddItem "Reel"
    cboLabelType.AddItem "PP Box"
    cboLabelType.AddItem "ICE Box"
    cboLabelType.ListIndex = 0
    
    txtPFrNo.Text = ""
    txtPToNo.Text = ""
    
    gSORT = 0

End Sub

'Private Sub spdRegOrder_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i               As Integer
'    Dim strDate         As String
'    Dim strLotNo        As String
'    Dim strProdPosNo    As String
'    Dim strProdCd       As String
'    Dim strSltNo        As String
'
'    If Row = 0 Then
'        Call SetSpreadSort(spdRegOrder)
'        Exit Sub
'    End If
'
'    spdRegOrderDetail.MaxRows = 0
'
'    txtLotNo.Text = GetText(spdRegOrder, Row, 1)
'    strDate = GetText(spdRegOrder, Row, 2)
'    dtpProdOrderDt.Value = Format(strDate, "####-##-##")
'    'cboProdPosNo.Text = GetText(spdRegOrder, Row, 3)
''    strProdPosNo = cboProdPosNo.Text
'
'    For i = 0 To cboCompCd.ListCount
'        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegOrder, Row, 4) Then
'            cboProdCd.ListIndex = i
'            strProdCd = Trim(mGetP(cboProdCd.List(i), 2, "|"))
'            Exit For
'        End If
'    Next
'    For i = 0 To cboPackCd.ListCount
'        If Mid(cboPackCd.List(i), 1, 2) = GetText(spdRegOrder, Row, 6) Then
'            cboPackCd.ListIndex = i
'            Exit For
'        End If
'    Next
'    txtOrderMemo.Text = GetText(spdRegOrder, Row, 7)
'    txtProdLen.Text = GetText(spdRegOrder, Row, 9)
'    cboSlittingNo.Text = GetText(spdRegOrder, Row, 10)
'    strSltNo = cboSlittingNo.Text
'    txtReelQTY.Text = GetText(spdRegOrder, Row, 11)
'
'    '고객사
'    For i = 0 To cboCompCd.ListCount
'        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = Trim(mGetP(GetText(spdRegOrder, Row, 12), 2, "|")) Then
'            cboCompCd.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    Call GetOrderDetail(strDate, strProdCd, strSltNo)
'
'
'End Sub


' 작업지시서 리스트 가져옴 'strDate, cboProdPosNo.Text, cboProdCd.Text, cboSlittingNo.Text
'Private Sub GetOrderDetail(ByVal pDate As String, ByVal pProCd As String, ByVal pSltNo As String)
'
'    Set AdoRs = Get_OrderDetail(pDate, pProCd, pSltNo)
'
'    If AdoRs Is Nothing Then
'        '등록된 정보 없음
'    Else
'        spdRegOrderDetail.MaxRows = 0
'
'        Do Until AdoRs.EOF
'            With spdRegOrderDetail
'                .MaxRows = .MaxRows + 1
'
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
'                'Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 2)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 3)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 4)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SEQ_NO").Value & "", .MaxRows, 5)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_INFO").Value & "", .MaxRows, 6)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_F").Value & "", .MaxRows, 7)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_T").Value & "", .MaxRows, 8)
'
'            End With
'
'            AdoRs.MoveNext
'        Loop
'        AdoRs.Close
'    End If
'
'End Sub


Private Sub spdPrtReprint_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row = 0 Then
        If Col = 1 Then
            If GetText(spdPrtReprint, 1, 1) = "1" Then
                For i = 1 To spdPrtReprint.DataRowCnt
                    Call SetText(spdPrtReprint, "0", i, 1)
                Next
            Else
                For i = 1 To spdPrtReprint.DataRowCnt
                    Call SetText(spdPrtReprint, "1", i, 1)
                Next
            End If
        Else
            '-- 정렬 추가
            Call SetSpreadSort(spdPrtReprint, 0)
        End If
        Exit Sub
    End If

End Sub
