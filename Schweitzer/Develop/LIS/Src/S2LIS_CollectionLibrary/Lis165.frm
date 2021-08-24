VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm165OutCol 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14670
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      Height          =   3750
      Left            =   1845
      ScaleHeight     =   3690
      ScaleWidth      =   10350
      TabIndex        =   118
      Top             =   2115
      Visible         =   0   'False
      Width           =   10410
      Begin VB.CommandButton Command3 
         Caption         =   "종료"
         Height          =   600
         Left            =   8505
         TabIndex        =   119
         Top             =   2925
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         Caption         =   "HIV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2625
         Left            =   270
         TabIndex        =   120
         Top             =   180
         Width           =   9735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "감염관리"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7860
      Left            =   5460
      TabIndex        =   123
      Top             =   1170
      Width           =   7005
      Begin VB.CommandButton Command1 
         Caption         =   "종 료"
         Height          =   495
         Left            =   5250
         TabIndex        =   167
         Top             =   7245
         Width           =   1665
      End
      Begin VB.Frame Frame12 
         Caption         =   "특이소견"
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   165
         Top             =   5790
         Width           =   6795
         Begin RichTextLib.RichTextBox RichText 
            Height          =   540
            Left            =   150
            TabIndex        =   166
            Top             =   300
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   953
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Lis165.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Drug Allergy"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   90
         TabIndex        =   161
         Top             =   4605
         Width           =   6795
         Begin VB.TextBox txtDrug 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   164
            Text            =   "Text1"
            Top             =   570
            Width           =   6465
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Penicillin"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   180
            TabIndex        =   163
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RadioContrast"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   1575
            TabIndex        =   162
            Top             =   225
            Width           =   1650
         End
      End
      Begin VB.Frame Frame10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         TabIndex        =   155
         Top             =   870
         Width           =   6795
         Begin VB.CheckBox Check1 
            Caption         =   "신종감염병"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   165
            TabIndex        =   160
            Top             =   510
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "홍역"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   3855
            TabIndex        =   159
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "수두"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2565
            TabIndex        =   158
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tb"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1335
            TabIndex        =   157
            Top             =   210
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "AFB"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   156
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   154
         Text            =   "Caution 수정은 감염관리실에 요청하여 주십시요."
         Top             =   6825
         Width           =   6795
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   90
         TabIndex        =   147
         Top             =   1680
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   180
            TabIndex        =   153
            Top             =   510
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VDRL"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   1305
            TabIndex        =   152
            Top             =   225
            Width           =   900
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HBsAg"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   2565
            TabIndex        =   151
            Top             =   225
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HIV"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   180
            TabIndex        =   150
            Top             =   225
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HCV"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   3870
            TabIndex        =   149
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HBc IgM"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   5205
            TabIndex        =   148
            Top             =   225
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   90
         TabIndex        =   132
         Top             =   2505
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   3855
            TabIndex        =   146
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CJD"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   2595
            TabIndex        =   145
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRSA"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   1560
            TabIndex        =   144
            Top             =   900
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRE"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   135
            TabIndex        =   143
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rotavirus"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   135
            TabIndex        =   142
            Top             =   585
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HAVIgM"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   135
            TabIndex        =   141
            Top             =   240
            Width           =   1380
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRSA"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   1560
            TabIndex        =   140
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRE"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   2595
            TabIndex        =   139
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "C.diffic"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   3855
            TabIndex        =   138
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRAB(IRAB)"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   5205
            TabIndex        =   137
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Caption         =   "옴"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1560
            TabIndex        =   136
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "이"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   2595
            TabIndex        =   135
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "장티푸스"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   3855
            TabIndex        =   134
            Top             =   585
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "세균성이질"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   5205
            TabIndex        =   133
            Top             =   585
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   825
         Left            =   90
         TabIndex        =   124
         Top             =   3750
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "유행성이하선염"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   33
            Left            =   4410
            TabIndex        =   131
            Top             =   225
            Width           =   1875
         End
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   32
            Left            =   3420
            TabIndex        =   130
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "수막구균수막염"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   31
            Left            =   1575
            TabIndex        =   129
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Caption         =   "백일해"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   30
            Left            =   135
            TabIndex        =   128
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "인플루엔자"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   2790
            TabIndex        =   127
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "신종플루"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   135
            TabIndex        =   126
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "풍진"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   1575
            TabIndex        =   125
            Top             =   225
            Width           =   1335
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   180
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종기록일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   19
         Left            =   3720
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종기록자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWDt 
         Height          =   300
         Left            =   5040
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   180
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWNm 
         Height          =   300
         Left            =   5040
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   510
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdNameP 
      BackColor       =   &H00E0E0E0&
      Caption         =   "네임지발행(&P)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5130
      Style           =   1  '그래픽
      TabIndex        =   117
      Tag             =   "0"
      Top             =   8550
      Width           =   1320
   End
   Begin VB.CommandButton cmdNotice 
      Caption         =   "전달사항"
      Height          =   285
      Left            =   4350
      TabIndex        =   112
      Top             =   270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "전달사항(&N)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6480
      Style           =   1  '그래픽
      TabIndex        =   109
      Tag             =   "0"
      Top             =   8550
      Width           =   1320
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "외래채혈 전달사항"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4260
      Left            =   5490
      TabIndex        =   99
      Top             =   1200
      Width           =   7005
      Begin VB.CommandButton cmdNoDel 
         Caption         =   "삭 제"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3195
         TabIndex        =   108
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNoSave 
         Caption         =   "저 장"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4455
         TabIndex        =   107
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         Caption         =   "전달사항"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Left            =   90
         TabIndex        =   101
         Top             =   870
         Width           =   6840
         Begin VB.TextBox txtPtId1 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1275
            MaxLength       =   10
            TabIndex        =   113
            Top             =   225
            Width           =   1410
         End
         Begin RichTextLib.RichTextBox RichRemark 
            Height          =   1920
            Left            =   135
            TabIndex        =   102
            Top             =   630
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   3387
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Lis165.frx":007F
         End
         Begin MedControls1.LisLabel lblPtNm1 
            Height          =   330
            Left            =   4710
            TabIndex        =   114
            Top             =   210
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
            BackColor       =   15857140
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   ""
            Appearance      =   0
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   16
            Left            =   120
            TabIndex        =   115
            Top             =   210
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "환자   ID"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   17
            Left            =   3660
            TabIndex        =   116
            Top             =   210
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "성      명"
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdNoExit 
         Caption         =   "종 료"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5700
         TabIndex        =   100
         Top             =   3600
         Width           =   1215
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   14
         Left            =   2850
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   180
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종기록일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   15
         Left            =   2850
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종기록자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblNtDt 
         Height          =   300
         Left            =   4170
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblNtID 
         Height          =   300
         Left            =   4170
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   510
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblNtNm 
         Height          =   300
         Left            =   5550
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   510
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblNtTm 
         Height          =   300
         Left            =   5550
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8655
      Left            =   0
      TabIndex        =   84
      Top             =   -160
      Width           =   4095
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   645
         Left            =   0
         TabIndex        =   88
         Tag             =   "136"
         Top             =   960
         Width           =   4065
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "주민등록번호"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2000
            TabIndex        =   96
            Tag             =   "15305"
            Top             =   340
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.TextBox txtSearchKey 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            MaxLength       =   13
            TabIndex        =   91
            Top             =   240
            Width           =   1830
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2500
            TabIndex        =   90
            Tag             =   "15305"
            Top             =   120
            Width           =   825
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   2000
            TabIndex        =   89
            Tag             =   "15304"
            Top             =   120
            Width           =   510
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   250
            Index           =   1
            Left            =   3350
            Shape           =   4  '둥근 사각형
            Top             =   140
            Width           =   555
         End
         Begin VB.Label lblReset 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "Reset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   200
            Left            =   3360
            MouseIcon       =   "Lis165.frx":010E
            MousePointer    =   99  '사용자 정의
            TabIndex        =   92
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   480
         Left            =   40
         TabIndex        =   85
         Tag             =   "136"
         Top             =   480
         Width           =   4035
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈 대상자만"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   87
            Tag             =   "15304"
            Top             =   165
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체 환자"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2040
            TabIndex        =   86
            Tag             =   "15305"
            Top             =   165
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6915
         Left            =   45
         TabIndex        =   93
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   12197
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16775406
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   285
         Left            =   75
         TabIndex        =   94
         Top             =   195
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         BackColor       =   8388608
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
         Caption         =   "환자검색"
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread slvwPtList 
         Height          =   900
         Left            =   60
         TabIndex        =   95
         Top             =   180
         Width           =   4035
         _Version        =   196608
         _ExtentX        =   7117
         _ExtentY        =   1588
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   20
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "Lis165.frx":09D8
         UserResize      =   1
         ScrollBarTrack  =   3
      End
   End
   Begin VB.CommandButton cmdRRmkVisible 
      BackColor       =   &H00E0E0E0&
      Caption         =   "맞춤채혈등록 (&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5775
      Style           =   1  '그래픽
      TabIndex        =   35
      Tag             =   "0"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame capFrame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4150
      TabIndex        =   13
      Tag             =   "104"
      Top             =   285
      Width           =   10455
      Begin MedControls1.LisLabel lblJumin 
         Height          =   330
         Left            =   4710
         TabIndex        =   122
         Top             =   840
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   21
         Left            =   3690
         TabIndex        =   121
         Top             =   870
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "주민번호"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdCaution 
         BackColor       =   &H008080FF&
         Caption         =   "Caution"
         Height          =   345
         Left            =   2700
         MaskColor       =   &H8000000F&
         Style           =   1  '그래픽
         TabIndex        =   98
         Top             =   -120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRRmk 
         BackColor       =   &H00F7F3F8&
         Caption         =   "맞춤채혈확인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2745
         Picture         =   "Lis165.frx":0E00
         Style           =   1  '그래픽
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton cmdContent 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기존채혈내역"
         Height          =   345
         Left            =   8310
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   1185
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   17
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txtReceptNo 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   16
         Top             =   555
         Width           =   1410
      End
      Begin VB.ComboBox cboOrdDate 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   210
         Width           =   2055
      End
      Begin VB.TextBox txtMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   14
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   1620
         Width           =   8250
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   4740
         TabIndex        =   18
         Top             =   210
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   300
         Left            =   7575
         TabIndex        =   19
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   1575
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         BackColor       =   16249848
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "◈ Remark"
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   1200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
         BackColor       =   15988216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   60
         Top             =   210
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "환자   ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReceptNo 
         Height          =   315
         Left            =   150
         TabIndex        =   61
         Top             =   540
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "영수증번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   3690
         TabIndex        =   62
         Top             =   210
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성      명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   3690
         TabIndex        =   63
         Top             =   540
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   6390
         TabIndex        =   64
         Top             =   210
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "처 방 일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   6390
         TabIndex        =   65
         Top             =   540
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "진 료 과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   20
         Left            =   150
         TabIndex        =   66
         Top             =   1200
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "상 병 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColID 
         Height          =   300
         Left            =   7575
         TabIndex        =   69
         Top             =   870
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   13
         Left            =   6390
         TabIndex        =   70
         Top             =   870
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "전문채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSang 
         Height          =   315
         Index           =   14
         Left            =   150
         TabIndex        =   82
         Top             =   855
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   255
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "감염주의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDiseaseSang 
         Height          =   315
         Left            =   1320
         TabIndex        =   83
         Top             =   885
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4740
         TabIndex        =   24
         Top             =   555
         Width           =   480
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5445
         TabIndex        =   23
         Top             =   555
         Width           =   480
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6150
         TabIndex        =   22
         Top             =   585
         Width           =   60
      End
      Begin VB.Label lblOrdDtCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   9675
         TabIndex        =   21
         Top             =   315
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7F3F8&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   690
         Index           =   0
         Left            =   165
         Top             =   1560
         Width           =   9420
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Caption         =   "        /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4725
         TabIndex        =   25
         Top             =   540
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00CDC481&
      Caption         =   "채혈+접수(&A)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   7830
      Style           =   1  '그래픽
      TabIndex        =   81
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblRIMsg 
      Height          =   465
      Left            =   1530
      TabIndex        =   80
      Top             =   8550
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   820
      BackColor       =   14411494
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "핵의학 처방이 존재합니다."
      Appearance      =   0
   End
   Begin VB.CommandButton cmdRA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "핵의학(&R)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9165
      Style           =   1  '그래픽
      TabIndex        =   79
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CheckBox chkPay 
      BackColor       =   &H00800000&
      Caption         =   "전체처방조회(미수납포함)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5850
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   3345
   End
   Begin MedControls1.LisLabel lblBar 
      Height          =   315
      Left            =   4150
      TabIndex        =   1
      Top             =   2580
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "검체 채취 리스트"
      LeftGab         =   100
   End
   Begin VB.Frame fraOrder 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   4150
      TabIndex        =   5
      Top             =   2820
      Width           =   10470
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5070
         Left            =   0
         TabIndex        =   12
         Tag             =   "10114"
         Top             =   480
         Width           =   10365
         _Version        =   196608
         _ExtentX        =   18283
         _ExtentY        =   8943
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14936810
         GridColor       =   14737632
         MaxCols         =   37
         MaxRows         =   5
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis165.frx":138A
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   6
         VisibleRows     =   5
      End
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2145
         ScaleHeight     =   300
         ScaleWidth      =   3690
         TabIndex        =   9
         Top             =   180
         Width           =   3690
         Begin VB.Shape shp1 
            BackColor       =   &H00553755&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   1
            Left            =   135
            Shape           =   3  '원형
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblLIS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임상병리"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   420
            TabIndex        =   11
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00496835&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   2
            Left            =   1185
            Shape           =   3  '원형
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblBBS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "혈액은행"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   1485
            TabIndex        =   10
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   0
            Left            =   0
            Shape           =   4  '둥근 사각형
            Top             =   0
            Width           =   2385
         End
      End
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체 선택(&A)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Width           =   1470
      End
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈시간변경 : "
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   300
         Left            =   6465
         TabIndex        =   6
         Top             =   150
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtpColDtTm 
         Height          =   300
         Left            =   8400
         TabIndex        =   8
         Top             =   135
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "yyyy-MM-dd H:mm"
         Format          =   84344835
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "채   혈 (&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   10560
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종 료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   285
      Left            =   4155
      TabIndex        =   0
      Top             =   45
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   503
      BackColor       =   8388608
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
      Caption         =   "환자 기본정보"
      LeftGab         =   100
   End
   Begin VB.Frame fraContent 
      BackColor       =   &H00DBE6E6&
      Height          =   3975
      Left            =   4425
      TabIndex        =   29
      Top             =   1605
      Visible         =   0   'False
      Width           =   9585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00DBE6E6&
         Height          =   3210
         Left            =   2910
         ScaleHeight     =   3150
         ScaleWidth      =   6555
         TabIndex        =   75
         Top             =   195
         Width           =   6615
         Begin MedControls1.LisLabel LisLabel7 
            Height          =   300
            Index           =   1
            Left            =   45
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   45
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   529
            BackColor       =   8388608
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
            Caption         =   "접수번호별 검사항목 리스트"
            Appearance      =   0
            LeftGab         =   200
         End
         Begin FPSpread.vaSpread tblLabno 
            Height          =   2745
            Left            =   45
            TabIndex        =   77
            Tag             =   "10114"
            Top             =   375
            Width           =   6465
            _Version        =   196608
            _ExtentX        =   11404
            _ExtentY        =   4842
            _StockProps     =   64
            AutoCalc        =   0   'False
            AutoClipboard   =   0   'False
            BackColorStyle  =   1
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormulaSync     =   0   'False
            MaxCols         =   8
            MoveActiveOnFocus=   0   'False
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "Lis165.frx":30B1
            StartingColNumber=   2
            VirtualRows     =   24
            VisibleCols     =   5
            VisibleRows     =   500
         End
      End
      Begin VB.PictureBox fraLabNo 
         BackColor       =   &H00DBE6E6&
         Height          =   3210
         Left            =   105
         ScaleHeight     =   3150
         ScaleWidth      =   2730
         TabIndex        =   72
         Top             =   195
         Width           =   2790
         Begin MedControls1.LisLabel LisLabel7 
            Height          =   300
            Index           =   0
            Left            =   30
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   30
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   529
            BackColor       =   8388608
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
            Caption         =   "접수번호리스트"
            Appearance      =   0
            LeftGab         =   200
         End
         Begin MSComctlLib.ListView lvwLabNo 
            Height          =   2745
            Left            =   15
            TabIndex        =   74
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4842
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "접수번호"
               Object.Width           =   2824
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "상태"
               Object.Width           =   1766
            EndProperty
         End
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8190
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   3405
         Width           =   1320
      End
   End
   Begin MedControls1.LisLabel lblRIFlag 
      Height          =   375
      Left            =   4170
      TabIndex        =   78
      Top             =   8175
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BackColor       =   15857140
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.Frame fraPtRmk 
      BackColor       =   &H00F1F5F4&
      Height          =   5280
      Left            =   4425
      TabIndex        =   31
      Top             =   870
      Visible         =   0   'False
      Width           =   9570
      Begin VB.ComboBox cboRmk 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   68
         Text            =   "Combo1"
         Top             =   2175
         Width           =   5400
      End
      Begin VB.CommandButton cmdRSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6840
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   4710
         Width           =   1320
      End
      Begin VB.TextBox txtRMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4095
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   44
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   2850
         Width           =   5370
      End
      Begin VB.TextBox txtRTitle 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2520
         Width           =   5370
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   40
         Tag             =   "PtID"
         Top             =   1515
         Width           =   315
      End
      Begin MedControls1.LisLabel lblRDeptNm 
         Height          =   330
         Left            =   5940
         TabIndex        =   38
         Top             =   1185
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   37
         Tag             =   "PtID"
         Top             =   1185
         Width           =   315
      End
      Begin VB.TextBox txtRDept 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1185
         Width           =   1515
      End
      Begin VB.CommandButton cmdRClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8175
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   4710
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   330
         Index           =   2
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   582
         BackColor       =   8388608
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
         Caption         =   "환자 특이사항 등록 및 조회"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin MSComCtl2.DTPicker dtpEntDt 
         Height          =   330
         Left            =   4095
         TabIndex        =   34
         Top             =   855
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   84344833
         CurrentDate     =   37505
      End
      Begin MedControls1.LisLabel lblRColNm 
         Height          =   315
         Left            =   5940
         TabIndex        =   41
         Top             =   1530
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRPtnm 
         Height          =   315
         Left            =   5940
         TabIndex        =   42
         Top             =   525
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblREntNm 
         Height          =   315
         Left            =   5940
         TabIndex        =   45
         Top             =   855
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblRData 
         Height          =   4125
         Left            =   60
         TabIndex        =   46
         Top             =   525
         Width           =   2685
         _Version        =   196608
         _ExtentX        =   4736
         _ExtentY        =   7276
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   9
         MaxRows         =   17
         OperationMode   =   2
         ScrollBars      =   0
         SelectBlockOptions=   0
         ShadowColor     =   15663103
         SpreadDesigner  =   "Lis165.frx":58BF
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F1F5F4&
         Height          =   435
         Left            =   4095
         TabIndex        =   47
         Top             =   1740
         Width           =   1845
         Begin VB.OptionButton optExp 
            BackColor       =   &H00F1F5F4&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   49
            Top             =   165
            Width           =   795
         End
         Begin VB.OptionButton optExp 
            BackColor       =   &H00F1F5F4&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   48
            Top             =   165
            Width           =   960
         End
      End
      Begin MedControls1.LisLabel lblRPtid 
         Height          =   315
         Left            =   4095
         TabIndex        =   51
         Top             =   525
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRSeq 
         Height          =   300
         Left            =   5955
         TabIndex        =   52
         Top             =   1860
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   2805
         TabIndex        =   53
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "환자정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   2805
         TabIndex        =   54
         Top             =   855
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "등록정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   2805
         TabIndex        =   55
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "맞춤채혈부서"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   2805
         TabIndex        =   56
         Top             =   1515
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "맞춤 채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   2805
         TabIndex        =   57
         Top             =   1845
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "폐기여부"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2805
         TabIndex        =   58
         Top             =   2520
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Title"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1800
         Index           =   10
         Left            =   2805
         TabIndex        =   59
         Top             =   2850
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3175
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   2805
         TabIndex        =   67
         Top             =   2175
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Templete"
         Appearance      =   0
      End
      Begin VB.TextBox txtRColid 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1515
         Width           =   1515
      End
      Begin VB.Shape Shape2 
         Height          =   4185
         Index           =   1
         Left            =   30
         Top             =   495
         Width           =   2745
      End
      Begin VB.Shape Shape2 
         Height          =   4185
         Index           =   0
         Left            =   2760
         Top             =   495
         Width           =   6735
      End
   End
   Begin MedControls1.LisLabel lblDiseaseSang_New 
      Height          =   315
      Left            =   45
      TabIndex        =   97
      Top             =   8610
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
End
Attribute VB_Name = "frm165OutCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private objPatient  As clsPatient
Private objSQL      As clsLISSqlCollection
Private objCollect  As clsLISCollectioin

Private blnOrdFg    As Boolean
Private blnSelAllFg As Boolean
Private blnMsgFg    As Boolean
Private blnClearFg  As Boolean

Private lngBackEven As Long
Private lngBackOdd  As Long
Private lngForeEven As Long
Private lngForeOdd  As Long

Private strDeptCd   As String

Private blnInitFg   As Boolean
Private mvarPtID    As String
Private mvarOrddt   As String

Private strWrkDiv   As String

Private Const lngMaxRows = 20
Private Const lngRowHeight = 12

Public Event LastFormUnload()

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

Private blnTest    As Boolean

Public Property Let Ptid(ByVal vData As String)
    mvarPtID = vData
End Property

Public Property Let OrdDt(ByVal vData As String)
    mvarOrddt = vData
End Property

Private Sub cboOrdDate_Click()
   
    '** 예수병원 추가 변수========
    Dim Message     As String
    Dim tmpDate     As String
    Dim strTemp     As String
    '=============================
       
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If cboOrdDate.ListIndex > -1 Then
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        strTemp = objSQL.RI_Collection_Flag(txtPtId.Text, tmpDate)
        
        lblRIFlag.Caption = medGetP(strTemp, 1, COL_DIV)
        strWrkDiv = medGetP(strTemp, 2, COL_DIV)
        If lblRIFlag.Caption = "Y" Then
            lblRIMsg.Caption = "핵의학 처방이 있습니다."
        Else
            lblRIMsg.Caption = ""
        End If
    Else
        lblRIFlag.Caption = "": strWrkDiv = ""
        lblRIMsg.Caption = ""
    End If
    '=====================================================================================
    
    Call DisplayOrder
    Call cmdNotice_Click
    
    If blnOrdFg Then
        cmdSave(0).Enabled = True
        cmdSave(1).Enabled = True
        tblOrdSheet.SetFocus
    Else
        If cboOrdDate.ListCount > 1 Then
            With tblOrdSheet
                .Row = 1: .Row2 = .MaxRows
                .Col = 1: .Col2 = .MaxCols
                .BlockMode = True
                .Action = ActionClearText
                .BlockMode = False
            End With
            cboOrdDate.SetFocus
            
            '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
            '=====================================================================================
            If lblRIFlag.Caption = "Y" Then
                '-- 핵의학 채취화면 Call 여부
                Message = MsgBox("핵의학 검사가 존재합니다. 계속진행하시겠습니까?", vbExclamation + vbYesNo, "채취확인")
                
'                If Message = vbYes Then
'                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                    Else
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        '사용자ID : ObjSysInfo.EmpId
'                    End If
'                End If

'                If Message = vbYes Then
'                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    Else
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    End If
'
'                End If


'                If Message = vbYes Then
'                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    Else
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    End If
'
'                End If

                If Message = vbYes Then
                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
                    Else
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
                    End If
                End If

            End If
            '=====================================================================================
        Else
            '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
            '=====================================================================================
            If lblRIFlag.Caption = "Y" Then
                '-- 핵의학 채취화면 Call 여부
                Message = MsgBox("핵의학 검사가 존재합니다. 계속진행하시겠습니까?", vbExclamation + vbYesNo, "채취확인")
                
'                If Message = vbYes Then
'                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    Else
'                        '** 핵의학 화면 Call
'                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'                        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & Trim(lblDeptNm.Caption) & " " & Trim(Format(cboOrdDate.Text, "yyyy-mm-dd")) & " " & "1", vbNormalFocus)
'                    End If
                    
'                End If

                If Message = vbYes Then
                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
                    Else
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
                    End If
                End If
            End If
            '=====================================================================================
            
            Call ClearRtn
'            txtPtId.Text = ""
'            txtPtId.SetFocus
            
        End If
    End If

End Sub

Private Sub cboRmk_Click()
    txtRMesg.Text = ""
    If cboRmk.ListIndex <> 0 Then
        txtRMesg.Text = medGetP(cboRmk.Text, 1, vbTab)
    End If
End Sub


Private Sub chkPay_Click()
    If txtPtId.Text <> "" Then
        Call txtPtId_LostFocus
    End If
End Sub

Private Sub chkSelAll_Click()
    Dim ii  As Integer
    
    blnSelAllFg = True
    With tblOrdSheet
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .Value = chkSelAll.Value
        .BlockMode = False
        If P_PayDtUsed Then
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = enCOLLIST.tcPAYDT
                If .Value = "" Then
                    .Col = enCOLLIST.tcCHECK: .Value = 0
                End If
            Next
        End If
    End With
    
    
    blnSelAllFg = False
   
End Sub

Private Sub cmdCaution_Click()
    Dim SQL As String
    Dim iCnt As Integer

    Set AdoCn_ORACLE = New ADODB.Connection

    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
'        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("Persist Security Info") = True
        
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        
'        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    SQL = ""
    SQL = SQL + "SELECT AFBYN,                                     "
    SQL = SQL + "       TBYN,                                      "
    SQL = SQL + "       SUDUYN,                                    "
    SQL = SQL + "       HONGYN,                                    "
    SQL = SQL + "       HIVYN,                                     "
    SQL = SQL + "       VDRLYN,                                    "
    SQL = SQL + "       HBSAGYN,                                   "
    SQL = SQL + "       HCVYN,                                     "
    SQL = SQL + "       HBCYN,                                     "
    SQL = SQL + "       HAVYN,                                     "
    SQL = SQL + "       MRSAYN,                                    "
    SQL = SQL + "       VREYN,                                     "
    SQL = SQL + "       CDIFFIYN,                                  "
    SQL = SQL + "       FUNGUSYN,                                  "
    SQL = SQL + "       ROTAYN,                                    "
    SQL = SQL + "       OHMYN,                                     "
    SQL = SQL + "       EEEYN,                                     "
    SQL = SQL + "       JANGTIYN,                                  "
    SQL = SQL + "       EEEJILYN,                                  "
    SQL = SQL + "       NEWFLUYN,                                  "
    SQL = SQL + "       PUNGYN,                                    "
    SQL = SQL + "       PENICILN,                                  "
    SQL = SQL + "       INFLUYN,                                    "
    SQL = SQL + "       NEWINFECYN,                                 "
    SQL = SQL + "       BETCYN,                                     "
    SQL = SQL + "       CREYN,                                      "
    SQL = SQL + "       VRSAYN,                                     "
    SQL = SQL + "       CJDYN,                                      "
    SQL = SQL + "       CETCYN,                                     "
    SQL = SQL + "       PERYN,                                      "
    SQL = SQL + "       MENYN,                                      "
    SQL = SQL + "       DETCYN,                                     "
    SQL = SQL + "       MUMPSYN,                                    "
    SQL = SQL + "       RADCONT,                                   "
    SQL = SQL + "       DRUGALGY,                                  "
    SQL = SQL + "       OTHERRMK,                                  "
    SQL = SQL + "       PATNO,                                     "
    SQL = SQL + "       SEQ,                                       "
    SQL = SQL + "       TO_CHAR(EDITDATE,'YYYYMMDD') AS EDITDATE,                      "
    SQL = SQL + "       EDITID,                                                        "
    SQL = SQL + "       FN_USERNAME_SELECT(EDITID) AS EDITNM                          "
    SQL = SQL + "  FROM MDCAUTNT                                                       "
    SQL = SQL + " WHERE PATNO = '" & Trim(txtPtId.Text) & "'                                             "
    SQL = SQL + "   AND SEQ = (SELECT MAX(SEQ) FROM MDCAUTNT WHERE PATNO = '" & Trim(txtPtId.Text) & "') "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open SQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            For iCnt = 0 To 33
                If .Fields(iCnt).Value = "Y" Then
                    Check1(iCnt).Value = 1
                Else
                    Check1(iCnt).Value = 0
                End If
            Next
            
'            '2014-01-26 인플루엔자 삽입
'            If .Fields("INFLUYN").Value = "Y" Then
'                Check1(23).Value = 1
'            Else
'                Check1(23).Value = 0
'            End If
'            '2014-01-26 인플루엔자 삽입
            
            lblWDt.Caption = Format(.Fields("EDITDATE").Value & "", "####-##-##")
            lblWNm.Caption = .Fields("EDITNM").Value & ""
            txtDrug.Text = .Fields("DRUGALGY").Value & ""
            RichText.Text = .Fields("OTHERRMK").Value & ""
            
            Frame4.Visible = True
            If Check1(4).Value = 1 Then
                Picture1.Visible = True
            Else
                Picture1.Visible = False
            End If
        Else
            Frame4.Visible = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    txtPtId.Text = ""
    txtReceptNo.Text = ""
    txtPtId.SetFocus
End Sub

Private Sub cmdContent_Click()
    Call LastCollectFg(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
    Set objPatient = Nothing
    Set objSQL = Nothing
    Set objCollect = Nothing
End Sub

Private Sub cboOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cboOrdDate_Click
End Sub

Private Sub cmdNameP_Click()
    Dim tmpDate     As String
    Dim OrdTmp      As String
    
    tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
    tmpDate = Format(tmpDate, "####-##-##")
    
    With tblOrdSheet
        .Row = 1: .Col = 37
        OrdTmp = Format(.Text, "####-##-##")
'        OrdTmp = Format(OrdTmp, CS_DateDbFormat)
    End With
    
    
    Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & "*" & " " & "*" & " " & "nurLABEL4" & " " & "0" & " " & "*" & " " & tmpDate & " " & "*", vbNormalFocus)
End Sub

Private Sub cmdNoDel_Click()
    Dim SSQL As String
        
    On Error GoTo Error_Jump
    
    If Trim(txtPtId1.Text) = "" Then
        Call MsgBox("환자번호가 없습니다.", vbInformation, "채혈 전달사항")
        Exit Sub
    End If
'    txtPtId = "12345678"
    
           SSQL = "DELETE                                      "
    SSQL = SSQL + "  FROM S2COM103                             "
    SSQL = SSQL + " WHERE PATID = '" & Format(Trim(txtPtId1.Text), "00000000") & "' "
     
    DBConn.BeginTrans
    
    DBConn.Execute SSQL
      
    lblNtDt.Caption = ""
    lblNtTm.Caption = ""
    lblNtID.Caption = ""
    lblNtNm.Caption = ""
    RichRemark.Text = ""
    txtPtId1 = ""
    lblPtNm1.Caption = ""
    
    Frame8.Visible = False
    
    DBConn.CommitTrans
    
    Exit Sub
    
Error_Jump:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmdNoExit_Click()
    txtPtId1 = ""
    lblPtNm1.Caption = ""
    Frame8.Visible = False
End Sub

Private Sub cmdNoSave_Click()
    Dim SSQL As String
    Dim Rs   As Recordset
        
    On Error GoTo Error_Jump
    
    If Trim(txtPtId1.Text) = "" Then
        Call MsgBox("환자번호가 없습니다.", vbInformation, "채혈 전달사항")
        Exit Sub
    End If
    
'    txtPtId = "12345678"
    
           SSQL = "SELECT *                                           "
    SSQL = SSQL + "  FROM S2COM103                             "
    SSQL = SSQL + " WHERE PATID = '" & Format(Trim(txtPtId1.Text), "00000000") & "' "

    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
      
    DBConn.BeginTrans
       
    With Rs
        If .RecordCount > 0 Then
            SSQL = "UPDATE s2com103 SET "
            SSQL = SSQL & " NOTICEREMARK = '" & Trim(RichRemark.Text) & "',           "
            SSQL = SSQL & " NOTICEID = '" & Trim(lblNtID.Caption) & "',               "
            SSQL = SSQL & " NOTICEDT = '" & Format(GetSystemDate, "YYYYMMDD") & "', "
            SSQL = SSQL & " NOTICETM = '" & Format(GetSystemDate, "HHmmss") & "'    "
            SSQL = SSQL & " WHERE PATID = '" & Format(Trim(txtPtId1.Text), "00000000") & "'                "
        Else
            SSQL = "Insert into s2com103"
            SSQL = SSQL & " (noticedt, noticetm, patid, noticeid, noticeremark)"
            SSQL = SSQL & " values("
            SSQL = SSQL & " '" & Replace(lblNtDt.Caption, "-", "") & "', "
            SSQL = SSQL & " '" & Replace(lblNtTm.Caption, ":", "") & "', "
            SSQL = SSQL & " '" & Format(Trim(txtPtId1.Text), "00000000") & "', "
            SSQL = SSQL & " '" & Trim(lblNtID.Caption) & "', "
            SSQL = SSQL & " '" & Trim(RichRemark.Text) & "') "
        End If
        DBConn.Execute SSQL
        .Close
    End With
      
    lblNtDt.Caption = ""
    lblNtTm.Caption = ""
    lblNtID.Caption = ""
    lblNtNm.Caption = ""
    RichRemark.Text = ""
    txtPtId1 = ""
    lblPtNm1.Caption = ""
    
    Frame8.Visible = False
    DBConn.CommitTrans
    Exit Sub
    
Error_Jump:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmdNotice_Click()
    Dim SSQL As String
    Dim iCnt As Integer
    Dim Rs As Recordset
    Dim strPtid As String
    
    If txtPtId1.Text = "" Then
        strPtid = Trim(txtPtId.Text)
    Else
        strPtid = Trim(txtPtId1.Text)
    End If
'    txtPtId.Text = "12345678"
    
           SSQL = "SELECT *                                           "
    SSQL = SSQL + "  FROM S2COM103                             "
    SSQL = SSQL + " WHERE PATID = '" & Format(Trim(strPtid), "00000000") & "' "

    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
      
    With Rs
        If .RecordCount > 0 Then
            lblNtID.Caption = Trim(.Fields("NOTICEID"))
            lblNtNm.Caption = GetEmpNm(.Fields("NOTICEID"))
            lblNtDt.Caption = Format(.Fields("NOTICEDT"), "####-##-##")
            lblNtTm.Caption = Format(.Fields("NOTICETM"), "##:##:##")
            RichRemark.Text = "" & .Fields("NOTICEREMARK")
            Frame8.Visible = True
        Else
            Frame8.Visible = False
        End If
        .Close
    End With

End Sub

Private Sub cmdOK_Click()
    fraContent.Visible = False
End Sub

Private Sub cmdRA_Click()
    Dim tmpDate     As String
    
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If lblRIFlag.Caption = "Y" Then
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        
'        If strWrkDiv = "3" Or strWrkDiv = "4" Then
'            '** 핵의학 화면 Call
'            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'        Else
'            '** 핵의학 화면 Call
'            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'            '사용자ID : ObjSysInfo.EmpId
'        End If

        If strWrkDiv = "3" Or strWrkDiv = "4" Then
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
        Else
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
        End If
      
        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
    End If
    '=====================================================================================
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Dim objPrgBar       As clsProgress
    Dim objDIC          As clsDictionary
    Dim BBSColSuccess   As Boolean
    Dim LISColSuccess   As Boolean
    
    Dim iCheckOrder     As Integer

    Dim ii              As Integer
    
'    If blnTest = True Then
'       MsgBox "즉시 검체보관 할 검체입니다.", vbInformation, "검체보관확인"
'    End If
    
    If CollectionTargetChk = False Then
'       MsgBox "채혈할 항목을 선택하세요..", vbInformation, "항목선택"
       tblOrdSheet.SetFocus
       Exit Sub
    End If
    
    cmdSave(0).Enabled = False
    cmdSave(1).Enabled = False
        
    iCheckOrder = objCollect.CheckSameOrder(tblOrdSheet, 1)     '중복처방 Check
    If iCheckOrder > 0 Then GoTo OrdCheck1
    
    Call MouseRunning
     
    Set objPrgBar = New clsProgress
    With objPrgBar
        .Container = Me
        .Left = lblBar.Left + 5
        .Top = lblBar.Top + 5
        .Width = lblBar.Width - 10
        .Height = lblBar.Height - 10
        .Message = "선택된 검사항목에 대해 채혈처리중입니다..."
        .Max = 90
        .Value = 10
        DoEvents
    End With

    DoEvents

    '----------------------------------------------------------
    '업무별 구분을 위해서 업무별로 불럭을 구분한다.(2001/06/08)
    '----------------------------------------------------------

    Set objDIC = New clsDictionary
    objDIC.Clear
    objDIC.FieldInialize "orddiv", "first,last,coldt,coltm"
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = enCOLLIST.tcORDDIV
            Select Case .Value
                Case BBS_ORDDIV
                    If objDIC.Exists(.Value) Then
                        objDIC.KeyChange BBS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        .Col = enCOLLIST.tcREQDTTM
                        objDIC.AddNew BBS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & _
                                      Format(.Text, "yyyymmdd") & COL_DIV & Format(.Text, "HHmm")
                    End If
                Case LIS_ORDDIV
                    If objDIC.Exists(.Value) Then
                        objDIC.KeyChange LIS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        objDIC.AddNew LIS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
            End Select
        Next
        objDIC.MoveFirst
        Do Until objDIC.EOF
            If objDIC.Fields("last") = "" Then
                objDIC.Fields("last") = objDIC.Fields("first")
            End If
            objDIC.MoveNext
        Loop
    End With
    With objDIC
        .MoveFirst
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case LIS_ORDDIV: iCheckOrder = objCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
            End Select
            If iCheckOrder > 0 Then GoTo OrdCheck2
            .MoveNext
        Loop
    End With
    
    '-------------------------------------------------------------
    '업무별로 채혈을 수행한다(혈액은행은 지정검체 체크가 필요없음)
    '-------------------------------------------------------------
    With objDIC
        .MoveFirst
        BBSColSuccess = True:  LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV: BBSColSuccess = CollectForBBS(.Fields("first"), .Fields("last"), _
                                                                    Format(GetSystemDate, "yyyymmdd"), _
                                                                    Format(GetSystemDate, "HHmmss"), objPrgBar)
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS(.Fields("first"), .Fields("last"), objPrgBar, Index)
            End Select
            .MoveNext
        Loop
    End With
    
    If Not BBSColSuccess And LISColSuccess Then
        Set objPrgBar = Nothing
        MsgBox "채혈처리중 오류가 발생했습니다 !!" & vbCrLf & _
               "재실행하신 후 오류가 계속되면 전산실 혹은 임상병리과로 연락바랍니다.", _
               vbCritical, "오류"
    End If
    
    cmdSave(0).Enabled = True
    cmdSave(1).Enabled = True
    
    Call MouseDefault
    Set objPrgBar = Nothing
    Set objDIC = Nothing
ExitPos:
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If lblRIFlag.Caption = "Y" Then
        Dim tmpDate     As String
        
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        
'        If strWrkDiv = "3" Or strWrkDiv = "4" Then
'            '** 핵의학 화면 Call
'            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'        Else
'            '** 핵의학 화면 Call
'            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
'            '사용자ID : ObjSysInfo.EmpId
'        End If
        
        If strWrkDiv = "3" Or strWrkDiv = "4" Then
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
        Else
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "lisLABEL" & " " & "O" & " " & "" & " " & "" & " " & "", vbNormalFocus)
        End If
        
        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
        
        '-- 핵의학 채취화면 Call 여부
'        Message = MsgBox("핵의학 검사가 존재합니다. 계속진행하시겠습니까?", vbExclamation + vbYesNo, "채취확인")
'
'        If Message = vbYes Then
'            '** 핵의학 화면 Call
'            Call Shell("C:\Schweitzer\DEVELOP\LIS\Bin\Collect.exe", vbMaximizedFocus)
'
'        End If
    End If
    '=====================================================================================
    
    '-- Clear Modify By M.G.Choi 200.06.28
    Call cboOrdDate_Click
    
    '-- 채혈 후 무조건 Clear By 온승호 2010.07.06
    If tblOrdSheet.DataRowCnt < 1 Then
        Call cmdClear_Click
    End If
    
    cmdSave(0).Enabled = True
    cmdSave(1).Enabled = True
    
    txtPtId.SetFocus
    Set objDIC = Nothing
    
    Exit Sub

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "중복처방입니다. 확인하고 다시 채혈하십시오.", vbExclamation, "중복처방"
    cmdSave(0).Enabled = True
    cmdSave(1).Enabled = True
    tblOrdSheet.SetFocus
    
    Exit Sub

OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "지정검체 정보가 없습니다. 전산실 혹은 임상병리과로 연락하세요.", vbInformation + vbOKOnly, "오류"
    cmdSave(0).Enabled = True
    cmdSave(1).Enabled = True
    tblOrdSheet.SetFocus
    Set objDIC = Nothing
    Exit Sub

End Sub


'** 혈액은행 채혈루틴
Private Function CollectForBBS(ByVal FRowCnt As Integer, ByVal LRowCnt As Integer, _
                                ByVal ColDt As String, ByVal ColTm As String, _
                                ByRef objProgress As Object) As Boolean

    
    Dim dicBBS      As clsDictionary
    Dim objBar      As clsDictionary
    Dim objCollect  As clsBBSCollection
    
    Dim tmpClipData As String
    Dim strStatFg   As String
    
    Dim tmpTotData  As Variant
    Dim tmpRowData  As Variant
    
    Dim lngColCnt   As Integer
    
    Dim i           As Long
    
    lngColCnt = 0
    
    Set objCollect = New clsBBSCollection
    
    If objCollect.Blood_Existence(txtPtId.Text, Format(GetSystemDate, "yyyymmdd"), Format(GetSystemDate, "hhmmss")) = False Then
        If objCollect.SetAccessCheck(txtPtId.Text) = True Then
           '검체가 이미 존재하는 경우
           CollectForBBS = objCollect.SetWardAccess(txtPtId.Text, enBussDiv.BussDiv_OutPatient, Format(GetSystemDate, "yyyymmdd"), _
                                    Format(GetSystemDate, "hhmmss"), ObjSysInfo.EmpId)
                
            Set objCollect = Nothing
            Exit Function
        Else
            GoTo AutoCollect
        End If
    End If
    
AutoCollect:

    Set dicBBS = New clsDictionary
    Set objBar = New clsDictionary
'    Set objCollect = New clsBBSCollection
    
    With tblOrdSheet
        .Col = 1:                  .Col2 = .MaxCols
        .Row = FRowCnt:            .Row2 = LRowCnt:                         .BlockMode = True
        tmpClipData = .ClipValue:  tmpTotData = Split(tmpClipData, vbCrLf): .BlockMode = False
        
        .Col = 7: strStatFg = IIf(Trim(.Value) = "Y", "1", "0")
        
        For i = 0 To UBound(tmpTotData) - 1

            tmpRowData = Split(tmpTotData(i), vbTab)
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            If tmpRowData(0) = 0 Then GoTo Skip       '선택여부
          
            lngColCnt = lngColCnt + 1
            
            '혈액은행-----------------------------------------------------------------------------
                
                dicBBS.Clear
                dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
                dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, ColDt, ColTm, _
                              gEmpId, enBussDiv.BussDiv_OutPatient, ObjSysInfo.BuildingCd, "", strStatFg), COL_DIV)

Skip:
       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForBBS = True
        Exit Function
    End If
          
    CollectForBBS = objCollect.Set_Collect(dicBBS, , objProgress)
    
    If CollectForBBS Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
        '바코드 출력
            BarCodePrintForBBS objBar
        Else
            If objCollect.CheckCol Then
                MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
            Else
                MsgBox "수혈처방 검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
            End If
        End If
        
        If objCollect.Spc72Chk Then
            MsgBox "해당 환자는 72시간내에 채혈한 검체가 존재합니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
    End If
    
    Set objBar = Nothing
    Set dicBBS = Nothing
    Set objCollect = Nothing
'    Set objProgress = Nothing
End Function

Private Function CollectForLIS(ByVal FRowCnt As Long, _
                               ByVal LRowCnt As Long, _
                               ByRef objProgress As Object, _
                               Optional pINx As Integer = 0) As Boolean
    Dim tmpData()   As String
    Dim i           As Integer
    Dim SelCount    As Integer
    Dim CollectCnt  As Integer
    
    Dim ColSuccess  As Boolean
    
    CollectCnt = 0
    Call objCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        For i = FRowCnt To LRowCnt
            
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            
            .Row = i
            
            .Col = enCOLLIST.tcCHECK
            If .Value <> 1 Then GoTo Skip

            CollectCnt = CollectCnt + 1
            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
            .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .Value        'ReqColDate

            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value        'TestDiv
            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value        'MultiFg
            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value        'SpcGrp
' 2009.01.14 양성현 미래처방의 처리를 위해 tcORDDATE를 BEDINDT로 사용함.
'            .Col = enCOLLIST.tcORDDT:  tmpData(9) = .Value        'OrdDt
            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value        'OrdDt
            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value       'OrdNo
            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value       'OrdSeq
            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value       'OrdCd
            .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .Value       '진료과
            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value       '처방의
            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value       '주치의
            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value       '검사 약어명
            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value       '라벨출력장수
            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value       'LabDiv
            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value       '검체약어명
            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value       '미생물접수번호범위
            
            Call objCollect.AddLabCollect(tmpData)

Skip:
        Next
    End With

    If CollectCnt = 0 Then
        CollectForLIS = True
        Exit Function
    End If

    With objCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)  '검체년도
        tmpData(1) = objPatient.Ptid                            '환자ID
        
'' 2008.10.24 전염성 보균자일경우 바코드에 별을 붙이는 기능 추가.
'
'        If Len(lblDiseaseSang.Caption) > 5 Then
'            tmpData(2) = "*" & Trim(objPatient.PtNm)
'        Else
'            tmpData(2) = objPatient.PtNm
'        End If

' 2011.07.06 전염성 보균자일경우 바코드에 별을 붙이는 기능 수정

        If Len(lblDiseaseSang_New.Caption) > 0 Then
            tmpData(2) = "*" & Trim(objPatient.PtNm)
        Else
            tmpData(2) = objPatient.PtNm
        End If
        
        
        tmpData(3) = objPatient.Sex                             '성별
        If IsDate(Format(objPatient.Dob, CS_DateLongMask)) Then                         '환자일령
            tmpData(4) = DateDiff("y", Format(objPatient.Dob, CS_DateLongMask), GetSystemDate)
        Else
            tmpData(4) = Mid(objPatient.Dob, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = ""                                          '입원일
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)  '입력일
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)  '입력시간
        tmpData(8) = ObjSysInfo.EmpId                            '입력자
        tmpData(9) = ""                                          '원접수번호
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) '채혈일
        .ColTm = Format(GetSystemDate, "HHMMSS")             '채혈일
        tmpData(11) = ObjSysInfo.EmpId                           '채혈자
        tmpData(12) = ""                                         '병동ID
        tmpData(13) = ""                                         '병실ID
        tmpData(14) = ""                                         '침상ID
        tmpData(15) = ""                                         '침상ID
        tmpData(16) = ObjSysInfo.BuildingCd                      '** 채혈이 수행되는 건물코드
        
        

        Call .SetColData(tmpData)
    End With
    ' 채혈 수행
    ColSuccess = objCollect.DoCollection(objProgress)
    
    '** 접수수행:장비PCx(혈당측정) 도입에 따라 채혈-접수 루틴 필요 (외래채혈실에서 바로 결과등록 하기 위함)
    '   추가 By M.G.Choi 2007.04.02
    '---------------------------------------------------------------------------------------------------------------------
    If ColSuccess = True And pINx = 1 Then
        objProgress.Message = "접수 Procedure를 수행하고 있습니다."
        Dim objAccess   As New clsLISAccession
        
        With objCollect
            If .CollectDone Then
                Dim pWorkArea As String
                Dim pAccDt As String
                Dim pAccSeq As Long
                
                For i = 1 To .ColCount
                    objProgress.Message = "접수 Procedure를 수행하고 있습니다. (" & CStr(i) & "/" & CStr(.ColCount) & ")"
                    Call .GetLabNumbers(i, pWorkArea, pAccDt, pAccSeq)
                    ColSuccess = objAccess.DoAccession_New(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
                    If Not ColSuccess Then Exit For
                    If objProgress.Value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
                    objProgress.Value = objProgress.Value + 1
                    DoEvents
                Next
            End If
        End With
        Set objAccess = Nothing
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    If Not ColSuccess Then
        Set objProgress = Nothing
        MsgBox "채혈처리중 오류가 발생했습니다 !!"
        MouseDefault  '0
        CollectForLIS = False
        Exit Function
    End If
    CollectForLIS = True

End Function

Private Sub BarCodePrintForBBS(objDIC As clsDictionary)
    Dim objSQL      As New clsBBSCollection
    Dim objBar      As New clsBarcode
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strW_Dept   As String
    Dim strAccSeq   As String         'SpcYy-SpcNo 형태의 검체번호
    Dim strHosilid  As String
    Dim strStatFg   As String
    
    Set objSQL = New clsBBSCollection
    Set objBar = New clsBarcode
    
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    strW_Dept = strDeptCd
    
    objDIC.MoveFirst
    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
'        strColDt = Mid(medGetP(objDIC.GetString, 4, COL_DIV), 1, 4)
        strColDt = Format(Mid(medGetP(objDIC.GetString, 4, COL_DIV), 5), "0#/0#")
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        strColTm = Format(strColTm, "0#:##")
        
        '검체번호 출력 : 2001.2.8 추가
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '
        objBar.Label_PrintOut BBSName, "XM", "", strAccSeq, strSpcNo, strPtid, _
                              strPtnm, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
                              "", 1
        objDIC.MoveNext
    Loop
    Set objBar = Nothing
    Set objSQL = Nothing

End Sub

Private Sub Command1_Click()
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame4.Visible = False
End Sub

Private Sub Command2_Click()
'    If txtPtId.Text = "" Then
        lblNtID.Caption = Trim(ObjSysInfo.EmpId)
        lblNtNm.Caption = GetEmpNm(Trim(ObjSysInfo.EmpId))
        lblNtDt.Caption = Format(GetSystemDate, "YYYY-MM-DD")
        lblNtTm.Caption = Format(GetSystemDate, "HH:mm:ss")
        RichRemark.Text = ""
'    Else
        Call cmdNotice_Click
'    End If

    Frame8.Visible = True
End Sub

Private Sub Command3_Click()
    Picture2.Visible = False
End Sub



Private Sub Form_Activate()
    If blnInitFg = False Then
        txtReceptNo.Visible = P_UseReceptForSearch
        lblReceptNo.Visible = P_UseReceptForSearch
    
        blnSelAllFg = False
        blnMsgFg = False
        blnClearFg = True
        optSort(2).Value = True
        
        Call ClearRtn
' 08.09.25 양성현 생년월일 대신 집주소로 변겅
'        medInitLvwHead lvwPtList, "환자ID,환자성명,주민등록번호,생년월일,성별/나이,연락처", "100,100,300,300,100,800"
        medInitLvwHead lvwPtList, "환자ID,성 명,주민등록번호,주 소,성별/나이,연락처", "200,100,700,5000,100,700"
        
        Call tblOrdSheet.GetOddEvenRowColor(lngBackOdd, lngForeOdd, lngBackEven, lngForeEven)
    End If
    
    blnInitFg = True
    txtPtId.SetFocus

End Sub

Private Sub Form_Load()
    
    blnInitFg = False

    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame4.Visible = False
    chkPay.Value = 1
    
    Frame8.Visible = False
    lblNtDt.Caption = ""
    lblNtID.Caption = ""
    RichRemark.Text = ""
    blnTest = False
    
'    cmdSave.Caption = "채 혈(&S)"
    chkPay.Visible = True
    Set objPatient = New clsPatient
    Set objSQL = New clsLISSqlCollection
    Set objCollect = New clsLISCollectioin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    mvarPtID = "": mvarOrddt = ""
    Set objMyList = Nothing
    Set objPatient = Nothing
    Set objSQL = Nothing
    Set objCollect = Nothing
End Sub

Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub

Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static lngOrder As Long
    With lvwPtList
        lngOrder = (lngOrder + 1) Mod 2
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = Choose(lngOrder + 1, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '환자정보 Display
    If Item = "" Then Exit Sub
    DoEvents
    With Item
        txtPtId.Text = Trim(.Text)                '환자ID
        txtPtId.SetFocus
        Call txtPtId_KeyPress(vbKeyReturn)
    End With
    
End Sub

Private Sub optOption_Click(Index As Integer)
    lvwPtList.ListItems.Clear
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    
    Dim ButtonValue As Variant
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim i           As Integer
    
    If blnSelAllFg Then Exit Sub
    
    blnSelAllFg = True
    
    With tblOrdSheet
      
        If Not blnOrdFg Then Exit Sub
        If Col <> 1 Then Exit Sub
      
        .Row = Row
        .Col = Col:   ButtonValue = .Value
' 2009.02.19 양성현 일단 왜래접수의 경우이니까 Pay가 안되었다면 무조건 선택이 안된다.

        If P_PayDtUsed Then  ' And Mid(ObjSysInfo.projectid, 1, 1) = LIS_ORDDIV Then
            .Col = enCOLLIST.tcPAYDT
            If .Value = "" Then
                .Col = Col
                .Value = Val(Val(.Value + 1) Mod 2)
                If .Value = 0 Then GoTo Skip
                Exit Sub
            End If
        End If
        
        If .Value = 0 Then GoTo Skip
        
        .Col = 9:   SvOrdDt = .Value
        .Col = 10:  SvOrdNo = .Value
        
        For i = 1 To .DataRowCnt
            If i <> Row Then
                .Row = i
                .Col = 9
                If .Value = SvOrdDt Then
                    .Col = 10
                    If .Value = SvOrdNo Then
                        .Col = 1
                        If .Value <> ButtonValue Then .Value = ButtonValue
                    End If
                End If
            End If
        Next
    End With
    
Skip:
    blnSelAllFg = False

End Sub

Private Sub LastCollectFg(Optional ByVal blnClick As Boolean = False)
    Dim Rs   As Recordset
    Dim SSQL As String
    Dim itmX As ListItem
    
    cmdContent.Visible = False
    
    SSQL = " SELECT workarea,accdt,accseq,coldt,coltm,stscd,testdiv FROM " & T_LAB201 & _
           " WHERE " & _
               DBW("ptid=", txtPtId.Text) & _
           " AND " & DBW("stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("coldt>=", Format(DateAdd("d", -30, GetSystemDate), "YYYYMMDD")) & _
           " AND " & DBW("coldt<=", Format(GetSystemDate, "YYYYMMDD")) & _
           " ORDER BY coldt desc,coltm desc"
           
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        cmdContent.Visible = True
        If Not blnClick Then
            Set Rs = Nothing
            Set itmX = Nothing
            Exit Sub
        End If
        lvwLabNo.ListItems.Clear
        Do Until Rs.EOF
            Set itmX = lvwLabNo.ListItems.Add(, , Rs.Fields("workarea").Value & "" & "-" & _
                                                  Rs.Fields("accdt").Value & "" & "-" & _
                                                  Rs.Fields("accseq").Value & "")
            Select Case Rs.Fields("stscd").Value & ""
                Case enStsCd.StsCd_LIS_Order:       itmX.SubItems(1) = STS_LIS_Order
                Case enStsCd.StsCd_LIS_Collection:  itmX.SubItems(1) = STS_LIS_HaveSpc
                Case enStsCd.StsCd_LIS_Accession:   itmX.SubItems(1) = STS_LIS_Access
                Case enStsCd.StsCd_LIS_InProcess:   itmX.SubItems(1) = STS_LIS_Worksheet
                Case enStsCd.StsCd_LIS_MidRst:
                                                If Rs.Fields("testdiv").Value & "" <> enTestDiv.TST_MicTest Then
                                                    itmX.SubItems(1) = STS_LIS_Partial
                                                Else
                                                    itmX.SubItems(1) = STS_LIS_MidRst
                                                End If
                Case enStsCd.StsCd_LIS_FinRst:
                                                If Rs.Fields("testdiv").Value & "" <> enTestDiv.TST_MicTest Then
                                                    itmX.SubItems(1) = STS_LIS_Verify
                                                Else
                                                    itmX.SubItems(1) = STS_LIS_FinRst
                                                End If
                Case enStsCd.StsCd_LIS_Modify:      itmX.SubItems(1) = STS_LIS_Modify
            End Select
            Rs.MoveNext
        Loop
                        '
        fraContent.Visible = True
        fraContent.ZOrder
        Set itmX = lvwLabNo.SelectedItem
        
        Call lvwLabNo_ItemClick(itmX)
    End If
    Set Rs = Nothing
    Set itmX = Nothing
End Sub
Private Sub lvwLabNo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sWorkarea   As String
    Dim sAccdt      As String
    Dim sAccSeq     As String
    Dim sTestcd     As String
    Dim strtmp      As String
    Dim SSQL        As String
    Dim Rs          As Recordset
    
    DoEvents

    Call medClearTable(tblLabno)
    
    strtmp = Item.Text
    
    sWorkarea = medGetP(strtmp, 1, "-")
    sAccdt = medGetP(strtmp, 2, "-")
    sAccSeq = medGetP(strtmp, 3, "-")
    
    SSQL = " SELECT a.orddt,b.coldt,b.coltm,b.rcvdt,b.rcvtm,b.vfydt,b.vfytm,c.abbrnm10,d.field3 as spcnm" & _
           " FROM " & T_LAB102 & " a," & T_LAB201 & " b," & T_LAB001 & " c," & T_LAB032 & " d" & _
           " WHERE " & _
                     DBW("b.workarea=", sWorkarea) & _
           " AND " & DBW("b.accdt=", sAccdt) & _
           " AND " & DBW("b.accseq=", sAccSeq) & _
           " AND " & DBW("d.cdindex=", LC3_Specimen) & _
           " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
           " AND c.testcd=a.ordcd" & _
           " AND c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
           "                   WHERE testcd = c.testcd ) " & _
           " AND d.cdval1=a.spccd"
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        With tblLabno
            Do Until Rs.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1:   .Value = Format(Rs.Fields("orddt").Value & "", "####-##-##")
                .Col = 2:   .Value = Rs.Fields("abbrnm10").Value & ""
                .Col = 3:   .Value = Rs.Fields("spcnm").Value & ""
                .Col = 4:   .Value = Format(Rs.Fields("coldt").Value & "", "####-##-##")
                            If Rs.Fields("coltm").Value & "" <> "" Then
                                .Value = .Value & " " & Format(Mid(Rs.Fields("coltm").Value & "", 1, 4), "0#:##")
                            End If
                .Col = 5:
                            If Rs.Fields("rcvdt").Value & "" <> "" Then
                                .Value = Format(Rs.Fields("rcvdt").Value & "", "####-##-##")
                            End If
                            If Rs.Fields("rcvtm").Value & "" <> "" Then
                                .Value = .Value & " " & Format(Mid(Rs.Fields("rcvtm").Value & "", 1, 4), "0#:##")
                            End If
                .Col = 6:
                            If Rs.Fields("vfydt").Value & "" <> "" Then
                                .Value = Format(Rs.Fields("vfydt").Value & "", "####-##-##")
                            End If
                            If Rs.Fields("vfytm").Value & "" <> "" Then
                                .Value = .Value & " " & Format(Mid(Rs.Fields("vfytm").Value & "", 1, 4), "0#:##")
                            End If
                Rs.MoveNext
            Loop
        End With
    End If
    Set Rs = Nothing
End Sub

Private Sub txtPtId_LostFocus()
    Dim lngCnt As Long
    Dim lblDiseaseSangT As String
    Dim lblDiseaseSangT_New As String
    
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
   
    Call ClearRtn
    With objPatient
        If IsNumeric(txtPtId.Text) Then
            txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
        End If
        
        Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
        
        If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm         '성명
            lblSex.Caption = .SEXNM         '성별
            lblAge.Caption = .Age           '연령
            lblAgeDiv.Caption = .AGEDIV     '나이단위
            lblDeptNm.Caption = .DeptNm     '진료과
            lblJumin.Caption = Mid(.SSN, 1, 6) & "-*******"
            blnClearFg = False
            
            If chkPay.Value = 1 Then
                lngCnt = objCollect.LoadOrderDate(txtPtId.Text, cboOrdDate)
            Else
                lngCnt = objCollect.LoadOrderDateYesPay(txtPtId.Text, cboOrdDate)
            End If
            
            Call cmdCaution_Click
            
            If lngCnt <= 0 Then
                blnMsgFg = True
                txtPtId1.Text = txtPtId.Text
                lblPtNm1.Caption = lblPtNm.Caption
                Call cmdNotice_Click
                MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다", vbExclamation, "외래채혈"
                txtPtId.Text = ""
                lblJumin.Caption = ""
                txtPtId.SetFocus
                blnMsgFg = False
                Call txtPtId_GotFocus
                Exit Sub
            Else
                Call LastCollectFg
                Call GetPtRmkVisibleTrueFalse
                
                If tblOrdSheet.DataRowCnt <> 0 Then
                    lblOrdDtCnt.Caption = CStr(lngCnt)
                    cboOrdDate.ListIndex = 0
                End If
                txtPtId1.Text = txtPtId.Text
                lblPtNm1.Caption = lblPtNm.Caption
            End If
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.Ptid = txtPtId.Text
            lblDisease.Caption = objDisease.Disease

'========================================================
' 08.10.24. 양성현 전주 예수병원 요청사항.
'========================================================
            
            lblDiseaseSangT = objDisease.DiseaseSang
            If Len(lblDiseaseSangT) > 5 Then
                lblSang(14).Visible = True
                lblSang(14).Refresh
                lblDiseaseSang.Visible = True
                lblDiseaseSang.Caption = lblDiseaseSangT
                
'            lblDiseaseSang.ForeColor = 1
            End If
' -------------------------------------------------------
            
'========================================================
' 2011.07.06. 온승호 전주 예수병원 요청사항.
'========================================================
            
            lblDiseaseSangT_New = objDisease.DiseaseSang_New
            If Len(lblDiseaseSangT_New) > 0 Then
                lblDiseaseSang_New.Caption = lblDiseaseSangT_New
            End If
' -------------------------------------------------------
            
            Set objDisease = Nothing
'            Call cmdCaution_Click
        Else
            blnMsgFg = True
            txtPtId.Text = ""
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbExclamation, "외래채혈"
            txtPtId.SetFocus
            blnMsgFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With

End Sub

Private Sub txtPtId1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lblPtNm1.Caption = GetPtNm(Format(Trim(txtPtId1.Text), "00000000"))
        Call cmdNotice_Click
    End If
End Sub

Private Sub txtReceptNo_Change()
   
    If Not blnClearFg Then
       txtPtId.Text = ""
    End If

End Sub

Private Sub txtReceptNo_GotFocus()
   
    With txtReceptNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtReceptNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtReceptNo_LostFocus()
   
    If txtReceptNo.Text = "" Then Exit Sub
    
    Dim tmpRs As Recordset
    
    Set tmpRs = New Recordset
    tmpRs.Open objSQL.SqlLoadOrderDate(txtReceptNo.Text, 1), DBConn
    
    If tmpRs.EOF Then
        MsgBox "해당 영수증번호는 존재하지 않습니다.", vbExclamation, "외래채혈"
        txtReceptNo.SetFocus
        GoTo NoData
    End If
    
    txtPtId.Text = "" & tmpRs.Fields("PtId").Value
    
    Call txtPtId_LostFocus
    
    cboOrdDate.Clear
    cboOrdDate.AddItem Format("" & tmpRs.Fields("OrdDt").Value, CS_DateMask)
    cboOrdDate.ListIndex = 0
    
    Call cboOrdDate_Click
    
NoData:
    Set tmpRs = Nothing
   
End Sub



'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call LoadPatient
       
End Sub

Private Sub LoadPatient()
'검색조건에 해당되는 환자를 조회한다.
'optOption의 값 체크
    Dim objPtInfo1  As clsPatient
'    Dim objPtInfo2  As clsPtInformation
    Dim Rs          As Recordset
    Dim itmX        As ListItem
    Dim strSQL      As String

    If Trim(txtSearchKey.Text) = "" Then
        
        Exit Sub
    End If
'    If optSort(2).Value Then
'        optOption(0).Value = False
'        optOption(1).Value = True
'    End If
    If optOption(0).Value And Not optSort(2).Value Then
        If IsNumeric(txtSearchKey.Text) Then
            txtSearchKey.Text = Format(txtSearchKey.Text, P_PatientIdFormat)
        End If
    End If
    
    
    If optOption(0).Value Then  '채혈대상에서 검색
        Set objPtInfo1 = New clsPatient  '  clsHosComSQLStmt

        If optSort(0).Value Then strSQL = objPtInfo1.GetSQLPtNt("3", txtSearchKey.Text)
        If optSort(1).Value Then strSQL = objPtInfo1.GetSQLPtNt("4", txtSearchKey.Text)
        If optSort(2).Value Then strSQL = objPtInfo1.GetSQLPtNt("7", txtSearchKey.Text)

'        If optSort(0).Value Then    'ID로 검색
'            strSQL = objPtInfo1.SqlPtntSearch(IIf(optSort(0).Value, 1, 2) + 2, txtSearchKey.Text)
'        Else    '환자명으로 검색
'            strSQL = objPtInfo1.SqlPtntSearch(IIf(optSort(0).Value, 1, 2) + 2, Trim(txtSearchKey.Text))
'        End If
    Else
'        Set objPtInfo2 = New clsPtInformation
        Set objPtInfo1 = New clsPatient
        
        If optSort(0).Value Then strSQL = objPtInfo1.GetSQLPtNt("1", txtSearchKey.Text)
        If optSort(1).Value Then strSQL = objPtInfo1.GetSQLPtNt("2", txtSearchKey.Text)
        
        '20131216 주민번호 6자리로 조회하도록 적용
        'If optSort(2).Value Then strSQL = objPtInfo1.GetSQLPtNt("0", txtSearchKey.Text)
        If optSort(2).Value Then strSQL = objPtInfo1.GetSQLPtNt("0", txtSearchKey.Text)

'        If optSort(0).Value Then    'ID로 검색
'            strSQL = objPtInfo2.GetPtInfo(Trim(txtSearchKey.Text), True)
'        Else    '환자명으로 검색
'            strSQL = objPtInfo2.GetPtInfo(txtSearchKey.Text, False)
'        End If
    End If

On Error GoTo NoData
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    lvwPtList.ListItems.Clear
    
    If Rs.EOF Then MsgBox "검색조건에 맞는 자료가 없습니다.", vbExclamation: GoTo NoData
            
    Do Until Rs.EOF
        Set itmX = lvwPtList.ListItems.Add()
        
        itmX.Text = Rs.Fields("ptid").Value & ""
        itmX.SubItems(1) = Rs.Fields("ptnm").Value & ""
        itmX.SubItems(2) = Rs.Fields("SSN").Value & ""

' 08.09.25 양성현 생년월일 대신 집주소로 변겅
'        itmX.SubItems(3) = Format(Rs.Fields("DOB").Value & "", CS_DateLongMask)
       
        itmX.SubItems(3) = Format(Rs.Fields("address").Value & "", CS_DateLongMask)
        Dim strYear As String
        
        If IsNull(Rs.Fields("ssn").Value & "") Then
            itmX.SubItems(4) = "XX"
        Else
'            If Len(Rs.Fields("ssn").Value & "") = 13 Then
'                itmX.SubItems(4) = IIf((Mid(Rs.Fields("ssn").Value & "", 7, 1) Mod 2) = 1, "남", "여")
'                If IsNumeric(Mid(Rs.Fields("ssn").Value & "", 1, 6)) Then
'                    itmX.SubItems(3) = Format(Mid(Rs.Fields("ssn").Value & "", 1, 6), "0#/##/##")
'                End If
'
'                If IsDate(itmX.SubItems(3)) Then
'                    itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), GetSystemDate)
'                Else
'                    itmX.SubItems(4) = itmX.SubItems(4) & " / " & "0"
'                End If
'            End If
'
            strYear = ""
        End If
        
        If Rs.Fields("hptelno").Value & "" <> "" Then
            itmX.SubItems(5) = Rs.Fields("hptelno").Value & ""
        Else
            itmX.SubItems(5) = Rs.Fields("telno").Value & ""
        End If
'        If lvwPtList.ListItems.Count = 1000 Then Exit Do
        Rs.MoveNext
    Loop
        
NoData:
    Set Rs = Nothing
    Set objPtInfo1 = Nothing
'    Set objPtInfo2 = Nothing
End Sub

'% 정렬 기준 선택
Private Sub optSort_Click(Index As Integer)
    If txtSearchKey.Text <> "" Then
       Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    If blnInitFg Then txtSearchKey.SetFocus
End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not blnClearFg Then Call ClearRtn
End Sub
'% 환자 ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = vbKeyReturn Then
        
        If txtPtId.Text <> "" Then
            cboOrdDate.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    Else
        If txtReceptNo.Text <> "" Then txtReceptNo.Text = ""
    End If
End Sub

'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
    Dim Rs          As Recordset
    Dim i           As Integer
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim strDoctNm   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim strOrdDiv   As String
    Dim strPayDt     As String
   
On Error GoTo NoData
     ' 처방내역 검색
    tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
    tmpTime = Format(Now, CS_TimeDbFormat)
    
    txtMesg.Text = ""
    DoEvents

    strOrdDiv = "W"
    strPayDt = IIf(chkPay.Value = "1", "", "1")
    
    blnTest = False
    
    SqlStmt = objSQL.SqlReadWardOrder(txtPtId.Text, tmpDate, tmpTime, , enBussDiv.BussDiv_OutPatient, strPayDt, strOrdDiv)
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        '-- 원본 =============================================
'        MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
        '=====================================================
        
        '-- 예수병원 변경 루틴 ===============================
        If lblRIFlag.Caption <> "Y" Then
'            MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
        End If
        '=====================================================
        
        MouseDefault
        blnOrdFg = False
        GoTo NoData
    End If
    
    With tblOrdSheet
       
        .ReDraw = False
        .MaxRows = 0
        If Rs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
            .Row = Rs.RecordCount + 1
            .Row2 = lngMaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = Rs.RecordCount   '데이타 건수
        End If

        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
             
        For i = 1 To Rs.RecordCount
            .Row = i
'            strDoctNm = GetDoctName(RS.Fields("orddoct").Value & "")
            strDoctNm = GetDoctNm(Rs.Fields("orddoct").Value & "") 'Trim(Rs.Fields("DoctNm").Value & "")
            If SvOrdDt <> Trim("" & Rs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & Rs.Fields("OrdDt").Value, CS_DateShortMask)    '처방일
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & Rs.Fields("OrdNo").Value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctNm                               '처방의
                SvOrdDt = Trim("" & Rs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)    '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)    '검체
                SvOrdDoct = strDoctNm                           '처방의
            End If
            If SvOrdNo <> Trim("" & Rs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & Rs.Fields("OrdNo").Value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctNm                               '처방의
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)    '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)    '검체
                SvOrdDoct = strDoctNm                            '처방의
            End If
            
            If SvSpcNm <> Trim("" & Rs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '검체
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)

            End If
            If SvOrdDoct <> strDoctNm Then
                .Col = enCOLLIST.tcDOCTNM: .Text = strDoctNm  '처방의
                SvOrdDoct = Trim(.Text)
            End If

'            Dim objBld As clsBasisData
            Dim strBld As String
            
            Select Case Rs.Fields("orddiv").Value & ""

                Case BBS_ORDDIV:
                    .Col = enCOLLIST.tcSTATFG: .Value = Trim("" & Rs.Fields("StatFg").Value)     '응급여부  --> 위에서 처리...
                    .Col = enCOLLIST.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
                    
'                    Set objBld = Nothing
'                    Set objBld = New clsBasisData
                    strBld = GetBuildNm(.Value)
'                    Set objBld = Nothing
                    
'                    If ObjLISComCode.Building.Exists(.Value) Then
'                        ObjLISComCode.Building.KeyChange (.Value)
'                    End If
                    .Col = enCOLLIST.tcBUILDNM: .Value = strBld 'ObjLISComCode.Building.Fields("buildnm")
                
                Case LIS_ORDDIV:
                    .Col = enCOLLIST.tcBUILDCD:  .Text = ObjSysInfo.BuildingCd
                    .Col = enCOLLIST.tcBUILDNM:  .Text = ObjSysInfo.BuildingNm
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(Rs.Fields("StatFg").Value & "")
            End Select
          
DataSet:
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & Rs.Fields("TestNm").Value)     '처방명
            Select Case Rs.Fields("orddiv").Value & ""
                Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '약간녹색   &H00845584&보라색
                Case LIS_ORDDIV: .ForeColor = &H553755
            End Select
            .Col = enCOLLIST.tcSTATFG:
                .Text = IIf("" & Rs.Fields("StatFg").Value = "1", "Y", "") '응급여부
                .ForeColor = DCM_Red                                '빨간색
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & Rs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format(IIf(Len("" & Rs.Fields("ReqTm").Value) = 4, "" & Rs.Fields("ReqTm").Value & "00", "" & Rs.Fields("ReqTm").Value), CS_TimeLongMask) '희망채취일시
            
            '처방일과 희망채혈일시가 다른경우 표시
'--- 2009.04.13 양성현 미래희망채혈일시를 Red로 수정- 강성수선생 요청사항
'--- 2012.03.12 온승호 조건을 현시점으로 변경 미래희망채혈일시를 Red로 수정- 강성수선생 요청사항
'            If Trim("" & Rs.Fields("OrdDt").Value < Trim("" & Rs.Fields("ReqDt").Value)) Then
            Dim tmpSysDate As String
            
            tmpSysDate = Format(GetSystemDate, CS_DateDbFormat)
            tmpSysDate = Replace(tmpSysDate, "-", "")
            
            If Trim(tmpSysDate < Trim("" & Rs.Fields("ReqDt").Value)) Then
                .ForeColor = DCM_Red
            Else
                If Trim("" & Rs.Fields("OrdDt").Value <> Trim("" & Rs.Fields("ReqDt").Value)) Then
                    .ForeColor = DCM_Blue
    '                .ForeColor = DCM_LightRed
    '--- 2009.03.01 양성현 희망채혈일시를 Blue로 수정- 강성수선생 요청사항
                Else
                    .ForeColor = DCM_Blue
                End If
            End If
'2009.01.14 양성현 미래처방의 처리를 위해BEDINDT로 수정 tcBedInDT
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & Rs.Fields("OrdDt").Value)
            .Col = enCOLLIST.tcBedInDT: .Text = Trim("" & Rs.Fields("bedindt").Value)
            .ForeColor = DCM_Blue                                '빨간색

            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & Rs.Fields("OrdNo").Value)      '처방번호
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & Rs.Fields("OrdSeq").Value)     '처방Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & Rs.Fields("OrdCd").Value)      '검사코드
'2015.09.15 온승호 보관방법이 즉각냉동인 항목 표기 추가
'            If .Text = "C3650" Then
'                .Col = 4
'                .ForeColor = DCM_Red
'                blnTest = True
'            Else
'                .Col = 4
'                .ForeColor = DCM_Black
'            End If
            
            Dim strLabDiv As String
            strLabDiv = GetLabDiv(.Text)
            
'            Call ObjLISComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = strLabDiv 'ObjLISComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & Rs.Fields("SpcCd").Value)      '검체코드
            
            Dim strLabRng As String
            Dim strSpcAbbr As String
            
            Call GetSpcInfo(.Text, strSpcAbbr, strLabRng)
            
'            Call ObjLISComCode.LisSpc.KeyChange(.Text)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & Rs.Fields("spcnm5").Value)         '검체약어명
            .Col = enCOLLIST.tcLABRANGE: .Text = strLabRng 'ObjLISComCode.LisSpc.Fields("labrange")    '미생물접수번호범위

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & Rs.Fields("WorkArea").Value)  'WorkArea
            
                
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & Rs.Fields("StoreCd").Value)   '보관코드
            '2015.09.15 온승호 보관방법이 즉각냉동인 항목 표기 추가
            
            If .Text = "I" Or .Text = "K" Or .Text = "L" Or .Text = "M" Or .Text = "N" Or .Text = "O" Or .Text = "P" Or .Text = "Q" Or .Text = "Y" Or .Text = "H" Then
                .Col = 4
                .ForeColor = DCM_Red
                blnTest = True
            Else
                .Col = 4
                .ForeColor = DCM_Black
            End If
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & Rs.Fields("TestDiv").Value)   '검사구분
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & Rs.Fields("MultiFg").Value)   '복수검체여부
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & Rs.Fields("SpcGrp").Value)    '검체군
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & Rs.Fields("OrdDoct").Value)   '처방의
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & Rs.Fields("MajDoct").Value)   '주치의
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & Rs.Fields("DeptCd").Value)    '진료과
                                         '진료과명
                                         'If .Text <> "" And lblDeptNm.Caption = "" Then
                                         If .Text <> "" Then
                                            strDeptCd = .Text
                                            
'                                            Set objBld = Nothing
'                                            Set objBld = New clsBasisData
                                            strBld = GetDeptNm(.Text)
'                                            Set objBld = Nothing
                                            
                                            lblDeptNm.Caption = strBld
'                                            If ObjLISComCode.DeptCd.Exists(.Text) Then
'                                                ObjLISComCode.DeptCd.KeyChange (.Text)
'                                                lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                                            End If
                                         End If
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & Rs.Fields("AbbrNm5").Value)    '약어명
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & Rs.Fields("LabelCnt").Value)   '라벨출력장수
            
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & Rs.Fields("PayDt").Value)   '영수증번호
'--- 2009.03.01 양성현 수납일시를 Black으로 수정- 강성수선생 요청사항
                                        .ForeColor = vbBlack

            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & Rs.Fields("WardId").Value)     '병동
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & Rs.Fields("hosilid").Value)     '병실
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & Rs.Fields("roomid").Value)      '병상

            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & Rs.Fields("FzFg").Value)       '동결절편
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & Rs.Fields("OrdDiv").Value)     '처방구분
            
            '진료부서 Remark
            If Trim("" & Rs.Fields("Mesg").Value) <> "" Then
                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & Rs.Fields("OrdNo").Value), "##") & " - "
                txtMesg.Text = txtMesg.Text & Trim("" & Rs.Fields("TestNm").Value) & vbCrLf
                txtMesg.Text = txtMesg.Text & Trim("" & Rs.Fields("Mesg").Value) & vbCrLf
            End If

            Rs.MoveNext
        Next

        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
       
    End With
    blnOrdFg = True
    fraOrder.Enabled = True
    
    If blnTest = True Then
       MsgBox "즉시 검체보관 할 검체입니다.", vbInformation, "검체보관확인"
    End If
    
NoData:
    Call MouseDefault
    Set Rs = Nothing
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
        GetLabDiv = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
        vSpcAbbr = Rs.Fields("spcabbr").Value & ""
        vLabRng = Rs.Fields("labrange").Value & ""
    End If
    Set Rs = Nothing
End Sub

Private Sub ClearRtn()
    lblDiseaseSang.Visible = False
    lblSang(14).Visible = False
    lblDisease.Caption = ""
    lblDiseaseSang.Caption = ""
    lblDiseaseSang_New.Caption = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDeptNm.Caption = ""
    lblOrdDtCnt.Caption = ""
    lblColID.Caption = ""
    txtMesg.Text = ""
    chkSelAll.Value = 0
    cboOrdDate.Clear
    fraOrder.Enabled = False
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    cmdSave(0).Enabled = False
    cmdSave(1).Enabled = False
    blnOrdFg = False
    blnMsgFg = False
    blnClearFg = True
    cmdContent.Visible = False
    fraContent.Visible = False
    fraPtRmk.Visible = False
    cmdRRmk.Visible = False
    lblRIFlag.Caption = ""
    lblRIMsg.Caption = ""
    
    strWrkDiv = ""
    
    Call PtRmkClear
    Set objPatient = Nothing
    Set objPatient = New clsPatient
'    Set objPatient.objDB = dbconn
    Set objCollect = Nothing
    Set objCollect = New clsLISCollectioin
    
End Sub
Public Sub Call_cmdClear_Click()
    Call cmdClear_Click
End Sub

Private Function CollectionTargetChk() As Boolean
    Dim ii      As Integer
    Dim tmpDate As String
    Dim strtmp  As String
    Dim strMsg  As String
    
    'tmpDate = Format(GetSystemDate, CS_DateDbFormat) & " " & Format(GetSystemDate, CS_TimeDbFormat)
    '2013-09-11 PSK 날짜 + 시간까지 Checking....
    tmpDate = Format(GetSystemDate, CS_DateDbFormat)
    tmpDate = Replace(Replace(Replace(tmpDate, "-", ""), ":", ""), " ", "")
    
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcCHECK
            If .Value = 1 Then
                CollectionTargetChk = True
                Exit For
            End If
        Next
        
        If CollectionTargetChk = False Then
            MsgBox "채혈할 항목을 선택하세요..", vbInformation, "항목선택"
            Exit Function
        End If
        
        Dim strMsgBox As String
        strMsgBox = " 진료일자가 " & strtmp & " 인 처방은 미래처방이므로 채혈이 불가능합니다." & vbCrLf & _
                    " 당일 채혈을 해야한다면 원무과에 진료일자를 당일로 변경요청 하신 후 다시 조회하여 처리 하셔야 합니다."
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcORDDIV
            If .Value = LIS_ORDDIV Then
' 퇴원환자의 미래채혈일시인 경우의 처리
                .Col = enCOLLIST.tcBedInDT: strtmp = .Value
                strtmp = Replace(Replace(Replace(strtmp, "-", ""), ":", ""), " ", "")
                'If tmpDate < strtmp Then
                If Left(tmpDate, 8) < Left(strtmp, 8) Then '2013-09-11 PSK 순수날짜만 비교한다.
                    strMsg = MsgBox(strMsgBox, vbYesNo + vbInformation, "Info")
                    If strMsg = vbNo Then
                        CollectionTargetChk = False
                        .Col = enCOLLIST.tcCHECK: .Value = 0
                    End If
                    Exit For
                End If

' 희망채혈일시가 미래인 경우의 처리
                .Col = enCOLLIST.tcREQDTTM: strtmp = .Value
                strtmp = Replace(Replace(Replace(strtmp, "-", ""), ":", ""), " ", "")
                'If tmpDate < strtmp Then
                If Left(tmpDate, 8) < Left(strtmp, 8) Then '2013-09-11 PSK 순수날짜만 비교한다.
                    strMsg = MsgBox(strMsgBox, vbYesNo + vbInformation, "Info")
                    If strMsg = vbNo Then CollectionTargetChk = False
                    
                    Exit For
                End If
            
            
            End If
        Next
        
    End With
    
End Function

Private Sub tblordersheet()
    With tblOrdSheet
        .SortBy = SortByRow
        .SortKey(1) = enCOLLIST.tcORDDIV
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Action = ActionSort
    End With
End Sub

Public Function AccListDisplayer()
    If mvarPtID = "" Then Exit Function
    blnClearFg = False
    txtPtId.SetFocus
    txtPtId.Text = mvarPtID
    cboOrdDate.AddItem Format(mvarOrddt, CS_DateLongFormat)
    
    cboOrdDate.ListIndex = medComboFind(cboOrdDate, Format(mvarOrddt, CS_DateLongFormat))

End Function

'환자별 특이사항등록 조회에 관련된 부분입니다.
'Table: s2ptmesg
'20040203

Private Sub cmdRSave_Click()
    Call InsertPtRmk
    Call QueryPtRmk
End Sub
Private Sub cmdRClose_Click()
    fraPtRmk.Visible = False
End Sub

Private Sub cmdRRmk_Click()
    Call RmkScreenSetting
    fraPtRmk.Visible = True: cmdRSave.Visible = False
    fraPtRmk.ZOrder
    Call QueryPtRmk
End Sub

Private Sub cmdRRmkVisible_Click()
    If txtPtId.Text = "" Then Exit Sub
    Call QueryPtRmk
    
    fraPtRmk.Visible = True:  cmdRSave.Visible = True
    fraPtRmk.ZOrder
    Call RmkScreenSetting(False)
    Call PtRmkClear
End Sub

Private Sub RmkScreenSetting(Optional ByVal blnSave As Boolean = True)
    LisLabel4(9).Visible = True:  lblRPtid.Visible = True: lblRPtnm.Visible = True
    LisLabel4(0).Visible = True:  dtpEntDt.Visible = True: lblREntNm.Visible = True
    
    LisLabel4(1).Visible = True:  txtRDept.Visible = True: cmdHelpList(0).Visible = True: lblRDeptNm.Visible = True
    LisLabel4(2).Visible = True:  txtRColid.Visible = True: cmdHelpList(1).Visible = True: lblRColNm.Visible = True
    
    LisLabel4(3).Visible = True:  Frame2.Visible = True:
    
    LisLabel4(6).Visible = True:  cboRmk.Visible = True
    LisLabel4(4).Visible = True:  txtRTitle.Visible = True
    LisLabel4(10).Visible = True: txtRMesg.Visible = True

    LisLabel4(1).Top = 1185:  txtRDept.Top = 1185:  cmdHelpList(0).Top = 1185: lblRDeptNm.Top = 1185
    LisLabel4(1).Enabled = True:  txtRDept.Enabled = True:  cmdHelpList(0).Enabled = True: lblRDeptNm.Enabled = True
    
    LisLabel4(2).Top = 1515:  txtRColid.Top = 1515: cmdHelpList(1).Top = 1515: lblRColNm.Top = 1515
    LisLabel4(2).Enabled = True:  txtRColid.Enabled = True: cmdHelpList(1).Enabled = True: lblRColNm.Enabled = True
    
    LisLabel4(6).Top = 2175:  cboRmk.Top = 2175
    LisLabel4(4).Top = 2520:  txtRTitle.Top = 2520
    LisLabel4(10).Top = 2855: txtRMesg.Top = 2855
    LisLabel4(6).Enabled = True:  cboRmk.Enabled = True
    LisLabel4(4).Enabled = True:  txtRTitle.Enabled = True
    LisLabel4(10).Enabled = True: txtRMesg.Enabled = True
    LisLabel4(10).Height = 1800: txtRMesg.Height = 1800
    If blnSave = True Then
        LisLabel4(9).Visible = False:  lblRPtid.Visible = False: lblRPtnm.Visible = False
        LisLabel4(0).Visible = False:  dtpEntDt.Visible = False: lblREntNm.Visible = False
        LisLabel4(3).Visible = False:  Frame2.Visible = False:
    
    
        LisLabel4(1).Enabled = False:  txtRDept.Enabled = False:  cmdHelpList(0).Enabled = False: lblRDeptNm.Enabled = False
        LisLabel4(2).Enabled = False:  txtRColid.Enabled = False: cmdHelpList(1).Enabled = False: lblRColNm.Enabled = False
        LisLabel4(6).Enabled = False:  cboRmk.Enabled = False
        LisLabel4(4).Enabled = False:  txtRTitle.Enabled = False
        LisLabel4(10).Enabled = False: txtRMesg.Enabled = False
        LisLabel4(1).Top = 525:  txtRDept.Top = 525:  cmdHelpList(0).Top = 525: lblRDeptNm.Top = 525
        LisLabel4(2).Top = 855:  txtRColid.Top = 855: cmdHelpList(1).Top = 855: lblRColNm.Top = 855
        
        LisLabel4(6).Top = 1185:  cboRmk.Top = 1185
        LisLabel4(4).Top = 1515:  txtRTitle.Top = 1515
        LisLabel4(10).Top = 1845: txtRMesg.Top = 1845
        LisLabel4(10).Height = 2805: txtRMesg.Height = 2805
    End If

End Sub


Private Sub PtRmkClear()
    lblRPtid.Caption = "":  lblRPtnm.Caption = "":  lblRDeptNm.Caption = "":    lblREntNm.Caption = ""
    lblRColNm.Caption = "": lblRSeq.Caption = ""
    txtRMesg.Text = "":     txtRTitle.Text = "":    txtRColid.Text = "":        txtRDept.Text = ""
    optExp(1).Value = True
    dtpEntDt.Value = GetSystemDate
    lblRPtid.Caption = txtPtId.Text: lblRPtnm.Caption = lblPtNm.Caption
    lblREntNm.Caption = ObjSysInfo.EmpNm
    cmdRSave.Caption = "저장"
    If cboRmk.ListCount > 0 Then cboRmk.ListIndex = 0
    If fraPtRmk.Visible = True And txtRDept.Enabled Then txtRDept.SetFocus

End Sub
Private Sub dtpEntDt_Click()
    txtRDept.SetFocus
End Sub

Private Sub txtRColid_KeyPress(KeyAscii As Integer)
'    Dim objEmp As clsBasisData
    
    If KeyAscii = vbKeyReturn Then
'        Set objEmp = New clsBasisData
        
        lblRColNm.Caption = GetEmpNm(UCase(txtRColid.Text)) 'GetEmpname(UCase(txtRColid.Text))
        If lblRColNm.Caption <> "" Then
            txtRTitle.SetFocus
        Else
            lblRColNm.Caption = ""
            txtRColid.Text = "": txtRColid.SetFocus
        End If
        
'        Set objEmp = Nothing
    End If
End Sub

Private Sub txtRDept_KeyPress(KeyAscii As Integer)
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    If KeyAscii = vbKeyReturn Then
'        Set objDept = New clsBasisData
        strDept = GetDeptNm(txtRDept.Text)
'        Set objDept = Nothing
        
        If strDept <> "" Then
            lblRDeptNm.Caption = strDept
            txtRColid.SetFocus
        Else
            lblRDeptNm.Caption = "": txtRDept.Text = "": txtRDept.SetFocus
        End If
'        If ObjLISComCode.DeptCd.Exists(UCase(txtRDept.Text)) Then
'            ObjLISComCode.DeptCd.KeyChange UCase(txtRDept.Text)
'            lblRDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'            txtRColid.SetFocus
'        Else
'            lblRDeptNm.Caption = "": txtRDept.Text = "": txtRDept.SetFocus
'        End If
    End If
End Sub

Private Sub txtRTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtRMesg.SetFocus
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objEmp As clsBasisData
    
    Set objMyList = New clsPopUpList
'    Set objEmp = New clsBasisData
    
    With objMyList
        .Connection = DBConn
        .HideToolTipText = True
        
        Select Case Index
            Case 1
                 .Connection = DBConn
                 .FormCaption = "사용자조회"
                 .Tag = "ColID"
                 .ColumnHeaderText = "사번;성명"
                 Call .LoadPopUp(GetSQLEmpList) ', fraPtRmk.Top + cmdHelpList(Index).Top, fraPtRmk.Left + cmdHelpList(Index).Left)
            Case 0
                 .FormCaption = "진료과 조회"
                 .Tag = "DeptCd"
                 .ColumnHeaderText = "진료과코드;진료과명"
                 Call .LoadPopUp(GetSQLDeptList) ', fraPtRmk.Top + cmdHelpList(Index).Top, fraPtRmk.Left + cmdHelpList(Index).Left)
        End Select
    End With
'    Set objEmp = Nothing
    Set objMyList = Nothing
End Sub
Private Sub objMyList_SendCode(ByVal SelString As String)
    Select Case objMyList.Tag
        Case "DeptCd"
            txtRDept.Text = Trim(medGetP(SelString, 1, ";"))
            lblRDeptNm.Caption = Trim(medGetP(SelString, 2, ";"))
        Case "ColID"
            txtRColid.Text = Trim(medGetP(SelString, 1, ";"))
            lblRColNm.Caption = Trim(medGetP(SelString, 2, ";"))
    End Select
End Sub
Private Sub tblRData_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii  As Integer
    
    cmdRSave.Caption = "저장"
    If Row < 1 Then Exit Sub
    Call PtRmkClear
    With tblRData
        .Row = Row
        .Col = 1: If .Value = "" Then Exit Sub
        .Col = 1: dtpEntDt.Value = CDate(.Value)
        .Col = 2: txtRTitle.Text = .Value
        .Col = 3: lblREntNm.Caption = medGetP(.Value, 1, COL_DIV)
        .Col = 4: txtRColid.Text = medGetP(.Value, 1, COL_DIV): lblRColNm.Caption = medGetP(.Value, 2, COL_DIV)
        .Col = 5: txtRDept.Text = medGetP(.Value, 1, COL_DIV):  lblRDeptNm.Caption = medGetP(.Value, 2, COL_DIV)
        .Col = 6: optExp(0).Value = IIf(.Value = "1", False, True)
        .Col = 7: txtRMesg.Text = .Value
        .Col = 8: lblRSeq.Caption = .Value
        .Col = 9:
            For ii = 0 To cboRmk.ListCount
                If .Value = Trim(medGetP(cboRmk.List(ii), 2, vbTab)) Then
                    cboRmk.ListIndex = ii
                    Exit For
                End If
            Next
    End With
    cmdRSave.Caption = "수정"
    
End Sub
Private Sub QueryPtRmk()
    Dim Rs   As Recordset
    Dim SSQL As String
    Dim ii   As Integer
    
    Call GetcboRmkCd
    Call medClearTable(tblRData)
    On Error GoTo Error_Jump
    Set Rs = New Recordset
    
    SSQL = GetPtRmkSQL
    
    Rs.Open SSQL, DBConn
    If Not Rs.EOF Then
        With tblRData
            Do Until Rs.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = Format(Rs.Fields("entdt").Value & "", "####-##-##")
                .Col = 2: .Value = Rs.Fields("title").Value & ""
                .Col = 3: .Value = Rs.Fields("empid").Value & ""
                .Col = 4: .Value = Rs.Fields("colid").Value & ""
                .Col = 5: .Value = Rs.Fields("deptcd").Value & ""
                .Col = 6: .Value = Rs.Fields("expfg").Value & ""
                .Col = 7: .Value = Rs.Fields("mesg").Value & ""
                .Col = 8: .Value = Rs.Fields("seq").Value & ""
                .Col = 9: .Value = Rs.Fields("rmkcd").Value & ""
                Rs.MoveNext
            Loop
        End With
    
    
    End If
    Call tblRData_Click(1, 1)
    
Error_Jump:
    Set Rs = Nothing
End Sub
Private Sub InsertPtRmk()
    Dim SSQL        As String
    Dim strMaxSEQ   As String
    Dim strExpFg    As String
    Dim strExpDt    As String
    Dim strRmkCD    As String
    
    
    strMaxSEQ = GetPtRmkMaxSeq
    If cboRmk.ListCount = 0 Then
        strRmkCD = ""
    Else
        strRmkCD = Trim(medGetP(cboRmk.Text, 2, vbTab))
    End If
    If strMaxSEQ = "" Then Exit Sub
    strExpFg = "0"
    
    If optExp(1).Value Then
        strExpFg = "1"
        strExpDt = Format(GetSystemDate, "YYYYMMDD")
    End If
    
    On Error GoTo Error_Jump
    DBConn.BeginTrans
    
    If lblRSeq.Caption = "" Then
        SSQL = " insert into " & T_LAB902 & " ( ptid,seq,entdt,empid,colid,deptcd,expfg,expdt,rmkcd,title,mesg) values(" & _
                DBV("ptid", lblRPtid.Caption, 1) & DBV("seq", strMaxSEQ, 1) & DBV("entdt", Format(dtpEntDt.Value, "YYYYMMDD"), 1) & _
                DBV("empid", lblREntNm.Caption, 1) & DBV("colid", txtRColid.Text & COL_DIV & lblRColNm.Caption, 1) & DBV("deptcd", txtRDept.Text & COL_DIV & lblRDeptNm.Caption, 1) & _
                DBV("expfg", strExpFg, 1) & DBV("expdt", strExpDt, 1) & DBV("rmkcd", strRmkCD, 1) & DBV("title", txtRTitle.Text, 1) & DBV("mesg", txtRMesg) & ")"
    Else
        SSQL = " update  " & T_LAB902 & " set " & _
                DBW("entdt", Format(dtpEntDt.Value, "YYYYMMDD"), 3) & DBW("empid", ObjSysInfo.EmpId & COL_DIV & lblREntNm.Caption, 3) & _
                DBW("colid", txtRColid.Text & COL_DIV & lblRColNm.Caption, 3) & DBW("deptcd", txtRDept.Text & COL_DIV & lblRDeptNm.Caption, 3) & _
                DBW("expfg", strExpFg, 3) & DBW("expdt", strExpDt, 3) & DBW("rmkcd", strRmkCD, 3) & DBW("title", txtRTitle.Text, 3) & DBW("mesg", txtRMesg, 2) & _
               " WHERE " & _
                DBW("ptid=", lblRPtid.Caption) & " AND " & DBW("seq=", lblRSeq.Caption)
    End If
    
    DBConn.Execute SSQL
    
    DBConn.CommitTrans
    Exit Sub
    
Error_Jump:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub
Private Function GetPtRmkMaxSeq() As String
    Dim Rs      As Recordset
    Dim SSQL    As String
    
    On Error GoTo Error_Jump
    
    SSQL = " SELECT max(seq) as maxseq FROM  " & T_LAB902 & "" & _
           " WHERE " & DBW("ptid", lblRPtid.Caption, 2)
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        GetPtRmkMaxSeq = Val(Rs.Fields("maxseq").Value & "") + 1
    Else
        GetPtRmkMaxSeq = 1
    End If
    
Error_Jump:
    Set Rs = Nothing
    
End Function
Private Sub GetPtRmkVisibleTrueFalse()
    Dim SSQL As String
    Dim Rs   As Recordset
'    Dim objEmp As clsBasisData
    
    lblColID.Caption = ""
    
    cmdRRmk.Visible = False
    On Error GoTo Error_Jump
    SSQL = GetPtRmkSQL
    
'    Set objEmp = New clsBasisData
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        lblColID.Caption = GetEmpNm(medGetP(Rs.Fields("colid").Value & "", 1, COL_DIV)) 'GetEmpname(medGetP(Rs.Fields("colid").Value & "", 1, COL_DIV))

        cmdRRmk.Visible = True
    End If
Error_Jump:
    Set Rs = Nothing
'    Set objEmp = Nothing

End Sub
Private Function GetPtRmkSQL() As String
    Dim SSQL As String
    
    SSQL = " SELECT * FROM  " & T_LAB902 & " " & _
           " WHERE " & DBW("ptid=", txtPtId.Text)
           
    GetPtRmkSQL = SSQL & " ORDER BY entdt desc"
End Function

Private Sub GetcboRmkCd()
    Dim SSQL    As String
    Dim Rs      As Recordset
    
    SSQL = " SELECT cdval1,text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex=", LC4_AccessComment)
    
    cboRmk.Clear
    cboRmk.AddItem "[ Remark Templete ]"
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        Do Until Rs.EOF
            cboRmk.AddItem Rs.Fields("text1").Value & "" & vbTab & Rs.Fields("cdval1").Value & ""
            Rs.MoveNext
        Loop
        cboRmk.ListIndex = 1
    End If
    Set Rs = Nothing
End Sub

 
