VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm155Accession 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검체접수"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
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
      Height          =   8580
      Left            =   3750
      TabIndex        =   53
      Top             =   420
      Width           =   7005
      Begin VB.Frame Frame3 
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
         TabIndex        =   99
         Top             =   3750
         Width           =   6795
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
            Left            =   150
            TabIndex        =   101
            Top             =   480
            Width           =   1125
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
            Left            =   150
            TabIndex        =   100
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "종 료"
         Height          =   495
         Left            =   5250
         TabIndex        =   98
         Top             =   8055
         Width           =   1665
      End
      Begin VB.Frame Frame12 
         Caption         =   "특이소견"
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   96
         Top             =   6600
         Width           =   6795
         Begin RichTextLib.RichTextBox RichText 
            Height          =   540
            Left            =   150
            TabIndex        =   97
            Top             =   300
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   953
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"Lis155.frx":0000
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
         TabIndex        =   92
         Top             =   5415
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
         TabIndex        =   86
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
         TabIndex        =   85
         Text            =   "Caution 수정은 감염관리실에 요청하여 주십시요."
         Top             =   7635
         Width           =   6795
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   90
         TabIndex        =   78
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
            TabIndex        =   84
            Top             =   510
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VDRL"
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
            Height          =   225
            Index           =   5
            Left            =   4080
            TabIndex        =   83
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
            Left            =   1305
            TabIndex        =   82
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
            TabIndex        =   81
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
            Left            =   2580
            TabIndex        =   80
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HBc IgM"
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
            Height          =   225
            Index           =   8
            Left            =   5205
            TabIndex        =   79
            Top             =   225
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   90
         TabIndex        =   62
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
            Index           =   36
            Left            =   5205
            TabIndex        =   77
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRPA"
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
            Index           =   35
            Left            =   5205
            TabIndex        =   76
            Top             =   600
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CPE"
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
            Index           =   34
            Left            =   3855
            TabIndex        =   75
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
            Left            =   1545
            TabIndex        =   74
            Top             =   900
            Width           =   1005
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
            Left            =   150
            TabIndex        =   73
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
            Left            =   2610
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   585
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
            Caption         =   "A형간염"
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRAB(CRAB)"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   3855
            TabIndex        =   67
            Top             =   240
            Width           =   1245
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   63
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   825
         Left            =   90
         TabIndex        =   54
         Top             =   4560
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
            Top             =   225
            Width           =   1155
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   102
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
         TabIndex        =   103
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
         TabIndex        =   104
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
         TabIndex        =   105
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      Height          =   3750
      Left            =   2250
      ScaleHeight     =   3690
      ScaleWidth      =   10350
      TabIndex        =   31
      Top             =   1935
      Visible         =   0   'False
      Width           =   10410
      Begin VB.CommandButton Command2 
         Caption         =   "종료"
         Height          =   600
         Left            =   8505
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   180
         Width           =   9735
      End
   End
   Begin VB.Frame fraMulti 
      BackColor       =   &H00DBE6E6&
      Height          =   6660
      Left            =   7575
      TabIndex        =   46
      Top             =   1785
      Width           =   6885
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Clear List"
         Height          =   330
         Left            =   165
         Style           =   1  '그래픽
         TabIndex        =   48
         Top             =   210
         Width           =   1185
      End
      Begin VB.CommandButton cmdClearRow 
         BackColor       =   &H00EDE2ED&
         Caption         =   "Clear Row"
         Height          =   330
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   47
         Top             =   210
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstAccList 
         Height          =   5940
         Left            =   150
         TabIndex        =   49
         Top             =   555
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15857140
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "총                건,  오류                 건"
         Height          =   225
         Left            =   4155
         TabIndex        =   52
         Top             =   330
         Width           =   2520
      End
      Begin VB.Label lblTotCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "150"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4410
         TabIndex        =   51
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblErrCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "150"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5835
         TabIndex        =   50
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
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
      TabIndex        =   15
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
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
      TabIndex        =   14
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   75
      ScaleHeight     =   390
      ScaleWidth      =   7440
      TabIndex        =   11
      Top             =   1890
      Width           =   7440
      Begin MedControls1.LisLabel lblMsg 
         Height          =   360
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   635
         BackColor       =   13434879
         ForeColor       =   5584725
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
         Caption         =   "이미 접수된 검체입니다 !!"
         LeftGab         =   100
      End
   End
   Begin VB.Frame fraReceive 
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
      Height          =   1620
      Left            =   75
      TabIndex        =   0
      Tag             =   "15502"
      Top             =   2205
      Width           =   7455
      Begin VB.CommandButton cmdCaution 
         BackColor       =   &H008080FF&
         Caption         =   "Caution"
         Height          =   345
         Left            =   3930
         MaskColor       =   &H8000000F&
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdOrderView 
         BackColor       =   &H00F4F0F2&
         Caption         =   "처방별조회(&C)"
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
         Left            =   5880
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   120
         Width           =   1500
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   1425
         TabIndex        =   1
         Top             =   510
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   556
         BackColor       =   13752531
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
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   870
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   556
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel lblPtId 
         Height          =   300
         Left            =   1425
         TabIndex        =   3
         Top             =   180
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel lblWard 
         Height          =   315
         Left            =   1410
         TabIndex        =   9
         Top             =   1215
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   556
         BackColor       =   13752531
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
         Index           =   5
         Left            =   225
         TabIndex        =   17
         Top             =   165
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
         Left            =   225
         TabIndex        =   18
         Top             =   510
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
         Caption         =   "성      명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   19
         Top             =   870
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
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   225
         TabIndex        =   20
         Top             =   1215
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
         Caption         =   "병     동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBLoodType 
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         BackColor       =   13752531
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel lblBLoodD 
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   3100
         _ExtentX        =   5477
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
         Caption         =   " 혈 액 형       검 사 일 시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBLoodDT 
         Height          =   315
         Left            =   5200
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         BackColor       =   13752531
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
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   13
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
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
      Caption         =   "접수대상검체 확인"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   16
      Top             =   3825
      Width           =   7440
      _ExtentX        =   13123
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
      Caption         =   "병동 선택"
      LeftGab         =   100
   End
   Begin VB.Frame Frame2 
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
      Height          =   4995
      Left            =   75
      TabIndex        =   4
      Top             =   4065
      Width           =   7440
      Begin VB.TextBox txtMesg 
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
         Height          =   1200
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   25
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   3690
         Width           =   7050
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   2445
         Left            =   180
         TabIndex        =   8
         Top             =   855
         Width           =   6960
         _Version        =   196608
         _ExtentX        =   12277
         _ExtentY        =   4313
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   14737632
         MaxCols         =   4
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis155.frx":007F
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   1470
         TabIndex        =   5
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel lblStoreNm 
         Height          =   315
         Left            =   4725
         TabIndex        =   6
         Top             =   510
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel lblLabNo 
         Height          =   315
         Left            =   1470
         TabIndex        =   10
         Top             =   165
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         BackColor       =   13752531
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
         Index           =   1
         Left            =   195
         TabIndex        =   21
         Top             =   165
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   195
         TabIndex        =   22
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "검     체"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   3450
         TabIndex        =   23
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "검체보관방법"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   195
         TabIndex        =   24
         Top             =   3315
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7F3F8&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   0
         Left            =   210
         Top             =   3660
         Width           =   7110
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "응 급"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3780
         TabIndex        =   7
         Tag             =   "105"
         Top             =   210
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpStat 
         BackColor       =   &H000000FF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   360
         Left            =   3435
         Top             =   150
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame fraInput 
      BackColor       =   &H00DBE6E6&
      Height          =   1605
      Left            =   60
      TabIndex        =   34
      Top             =   180
      Width           =   14385
      Begin VB.CheckBox chkReader 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Barcode Reader 사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2670
         TabIndex        =   43
         Tag             =   "15501"
         Top             =   255
         Width           =   2535
      End
      Begin VB.TextBox txtBarcode 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4020
         TabIndex        =   42
         Top             =   615
         Width           =   2910
      End
      Begin VB.PictureBox picLabNo 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   4020
         ScaleHeight     =   330
         ScaleWidth      =   2865
         TabIndex        =   38
         Top             =   1095
         Width           =   2925
         Begin VB.TextBox txtAccNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2115
            TabIndex        =   41
            Top             =   60
            Width           =   705
         End
         Begin VB.TextBox txtAccDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   6
            TabIndex        =   40
            Top             =   60
            Width           =   1080
         End
         Begin VB.TextBox txtWorkArea 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   15
            MaxLength       =   2
            TabIndex        =   39
            Top             =   60
            Width           =   600
         End
         Begin VB.Line Line1 
            X1              =   660
            X2              =   810
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line2 
            X1              =   1950
            X2              =   2100
            Y1              =   180
            Y2              =   180
         End
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일괄접수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   495
         TabIndex        =   37
         Top             =   405
         Width           =   1260
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "개별접수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   495
         TabIndex        =   36
         Top             =   1065
         Width           =   1260
      End
      Begin VB.CommandButton cmdExecute 
         BackColor       =   &H00F4F0F2&
         Caption         =   "일괄접수 실행(&S)"
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
         Left            =   7485
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   1035
         Width           =   1740
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   3
         Left            =   2655
         TabIndex        =   44
         Top             =   615
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Caption         =   "Barcode"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   4
         Left            =   2655
         TabIndex        =   45
         Top             =   1095
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Caption         =   "접수 번호"
         Appearance      =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   2220
         X2              =   2220
         Y1              =   225
         Y2              =   1600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   2205
         X2              =   2205
         Y1              =   225
         Y2              =   1600
      End
   End
End
Attribute VB_Name = "frm155Accession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Click()

Private WithEvents fL401 As S2LIS_ReviewLib.clsLisReviewForm
Attribute fL401.VB_VarHelpID = -1

Private tmpAccDt As String
Private objMySql As New clsLISSqlAccession
Private blnExeFg As Boolean

Private Const CS_AccSuccess = "정상"
Private Const lngMaxRows = 9
Private Const lngRowHeight = 12.5

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

'% 바코드 리더기를 사용할 것인지에 대한 여부
Private Sub chkReader_Click()
    Call ClearRtn
    If chkReader.Value = 1 Then
        txtBarcode.Locked = False
        txtBarcode.BackColor = DCM_White    '&H80000005
        picLabNo.Enabled = False
        picLabNo.BackColor = DCM_LightGray
        txtWorkArea.BackColor = DCM_LightGray
        txtAccDt.BackColor = DCM_LightGray
        txtAccNo.BackColor = DCM_LightGray
        txtBarcode.SetFocus
    Else
        txtBarcode.Locked = True
        txtBarcode.BackColor = DCM_LightGray
        picLabNo.Enabled = True
        picLabNo.BackColor = DCM_White
        txtWorkArea.BackColor = DCM_White
        txtAccDt.BackColor = DCM_White
        txtAccNo.BackColor = DCM_White
        txtWorkArea.SetFocus
    End If
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
    SQL = SQL + "       CPETCYN,                                    "
    SQL = SQL + "       CPEYN,                                    "
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
    SQL = SQL + " WHERE PATNO = '" & Trim(lblPtId.Caption) & "'                                             "
    SQL = SQL + "   AND SEQ = (SELECT MAX(SEQ) FROM MDCAUTNT WHERE PATNO = '" & Trim(lblPtId.Caption) & "') "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open SQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            For iCnt = 0 To 36
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
            
            Frame1.Visible = True
            If Check1(4).Value = 1 Then
                Picture1.Visible = True
            Else
                Picture1.Visible = False
            End If
        Else
            Frame1.Visible = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing
End Sub

Private Sub cmdClear_Click()
    Dim intTemp As Integer
    
    '## 5.1.7: 이상대(2005-05-31)
    '   - 바코드만 리딩하고 접수하지 않은 검체가 있을경우 확인 메시지를 출력하도록 수정
    If lstAccList.ListItems.Count > 0 Then
        If lstAccList.ListItems(1).SubItems(2) = "" Then
            intTemp = MsgBox("접수되지 않은 검체가 있습니다. 화면을 지울까요?", vbYesNo + vbDefaultButton2 + vbQuestion, "확인")
            If intTemp = vbNo Then Exit Sub
        End If
    End If
    
    Call ClearRtn
    Call cmdReset_Click
    optOption(0).Value = True
    txtWorkArea.Text = ""
    txtAccDt.Text = ""
    txtAccNo.Text = ""
    txtBarcode.Text = ""
    If chkReader.Value = 1 Then
        txtBarcode.SetFocus
    Else
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub cmdClearRow_Click()

    Dim i As Long
    
    For i = lstAccList.ListItems.Count To 1 Step -1
        If lstAccList.ListItems(i).Checked Then
            lstAccList.ListItems.Remove (i)
        End If
    Next
    For i = 1 To lstAccList.ListItems.Count
        lstAccList.ListItems(i).Text = CStr(i)
    Next

End Sub

Private Sub cmdExecute_Click()
    
    If lstAccList.ListItems.Count <= 0 Then
        MsgBox "바코드가 한 건도 리딩되지 않았습니다.", vbExclamation, "검체접수"
        Exit Sub
    End If
    
    optOption(0).Enabled = False
    optOption(1).Enabled = False
    txtBarcode.Enabled = False
    txtBarcode.BackColor = DCM_LightGray
    cmdExecute.Enabled = False
    cmdExit.Enabled = False
    cmdClear.Enabled = False
    'fraInput.Enabled = False
    
    Dim blnAccFg As Boolean
    Dim i As Integer
        
    blnExeFg = True
    For i = 1 To lstAccList.ListItems.Count
        txtBarcode.Text = lstAccList.ListItems(i).SubItems(1)
        
        Call ClearRtn
        
        blnAccFg = DisplayOrder(0, i)
        If blnAccFg Then Call DoAccession(i)
        
        txtBarcode.Text = ""
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
        DoEvents
    Next
    
'    For i = lstAccList.ListItems.Count To 1 Step -1
'        If lstAccList.ListItems(i).SubItems(2) = CS_AccSuccess Then '"정상"
'            lstAccList.ListItems.Remove (i)
'        Else
'            lblErrCnt.Caption = Val(lblErrCnt.Caption) + 1
'        End If
'        DoEvents
'    Next
    blnExeFg = False
    
    cmdExit.Enabled = True
    cmdClear.Enabled = True

End Sub

'% 종료
Private Sub cmdExit_Click()
    Set objMySql = Nothing
    Unload Me
    Set frm155Accession = Nothing
End Sub

Private Sub cmdOrderView_Click()
' 2008.12.17. 양성현 작업중입니다.
' 2009.01.09 양성현 환자ID 파라메터 추가
    Dim i As Integer
    Dim pFrmName As String
    If Len(lblPtId.Caption) < 2 Then GoTo End2Stop

'    Dim cxxx  As S2LIS_ReviewLib.clsLISResultReview
    pFrmName = "frm401ResultView"
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    medMain.lblSubMenu.Caption = "처방결과조회" 'medGetP(Button.Tag, 1, "(")
    
    
'   gPatientId = lblPtId.Caption
'  s2lis_reviewlib.PtId = lblPtId.Caption
    
'    gUsingInWardMenu = True
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.PtId = lblPtId.Caption
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
End2Stop:
End Sub

Private Sub cmdReset_Click()
    lstAccList.ListItems.Clear
    'fraInput.Enabled = True
    optOption(0).Enabled = True
    optOption(1).Enabled = True
    txtBarcode.Enabled = True
    cmdExecute.Enabled = True
    fraInput.Enabled = True
    txtBarcode.BackColor = DCM_White
    lblMsg.Caption = ""
    lblTotCnt.Caption = ""
    lblErrCnt.Caption = ""
    txtBarcode.SetFocus
End Sub

Private Sub Command1_Click()
    lblWDt.Caption = ""
    lblWNm.Caption = ""
'    txtVival.Text = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame1.Visible = False
End Sub

Private Sub Command2_Click()
    Picture2.Visible = False
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

'% 폼 로드
Private Sub Form_Load()
    Me.Show
    chkReader.Value = 1
    medInitLvwHead lstAccList, "Seq,검체번호,Message,SeqNo", "-1000,-300,1000,300"

    optOption(0).Value = True
    
    lblWDt.Caption = ""
    lblWNm.Caption = ""
'    txtVival.Text = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame1.Visible = False
    
    Call cmdReset_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub optOption_Click(Index As Integer)
    cmdExecute.Enabled = optOption(0).Value
    If optOption(0).Value Then
        chkReader.Value = 1
        chkReader.Enabled = False
        fraMulti.Enabled = True
    Else
        chkReader.Enabled = True
        fraMulti.Enabled = False
    End If
    If chkReader.Value = 1 Then
        txtBarcode.SetFocus
    Else
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With tblOrdSheet
        .Row = Row: .Col = 4
        txtMesg.Text = .Value
    End With
End Sub

Private Sub txtAccDt_Change()
    Dim strDt As String
    Dim strYYYY As String
    
    strDt = Mid(txtAccDt.Text, 1, 2) & "-01-01"
    strYYYY = Format(strDt, "yyyy")
    tmpAccDt = strYYYY & Mid(txtAccDt.Text, 3)
    
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then
       If txtAccNo.Enabled Then txtAccNo.SetFocus
    End If
End Sub

Private Sub txtAccDt_GotFocus()
    With txtAccDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    If txtAccDt.Text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccNo.SetFocus
End Sub

Private Sub txtAccNo_GotFocus()
    With txtAccNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 접수번호를 입력했을 경우
Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = vbKeyReturn Then
    
        If txtAccNo.Text = "" Then Exit Sub
           
        Call ClearRtn
        
        Dim blnAccFg As Boolean
        blnAccFg = DisplayOrder(1)
        If blnAccFg Then Call DoAccession
        
        txtBarcode.Text = ""
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
        If chkReader.Value = 1 Then
            txtBarcode.SetFocus
        Else
            txtWorkArea.SetFocus
        End If
    
    End If

End Sub

'% 리더기로 검체번호를 리딩했을 경우...
Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call medClearTable(tblOrdSheet)
        If txtBarcode.Text = "" Then Exit Sub
         Dim blnAccFg As Boolean
        If optOption(0).Value Then
            
            '일괄접수
            If blnExeFg Then Exit Sub
            blnAccFg = DisplayOrder(0)
            'lstAccList.ListItems.Add , , txtBarcode.Text
            lstAccList.ListItems.Add , , lstAccList.ListItems.Count + 1
            lstAccList.ListItems(lstAccList.ListItems.Count).SubItems(1) = txtBarcode.Text
            lblTotCnt.Caption = lstAccList.ListItems.Count
            lstAccList.SetFocus
            SendKeys "^{END}"
            DoEvents
            txtBarcode.Text = ""
            txtBarcode.SetFocus
        
        Else
            
            '개별접수
            Call ClearRtn
           
            blnAccFg = DisplayOrder(0)
            
            If blnAccFg Then Call DoAccession
            
            txtBarcode.Text = ""
            txtWorkArea.Text = ""
            txtAccDt.Text = ""
            txtAccNo.Text = ""
            If chkReader.Value = 1 Then
                txtBarcode.SetFocus
            Else
                txtWorkArea.SetFocus
            End If
        End If
    End If

End Sub


'% 접수번호 또는 검체번호를 기준으로 발생한 검사내역을 검색한다.
Private Function DisplayOrder(ByVal QueryOption As Integer, Optional ByVal ii As Integer) As Boolean

    Dim objRs, bRs As Recordset
    Dim tmpSQL As String
    Dim tmpBarcode As String
    Dim tmpBLood    As String
    Dim i As Long

    
    If QueryOption = 1 Then
        tmpSQL = objMySql.SqlOrdersForAccess(1, txtWorkArea.Text, tmpAccDt, txtAccNo.Text)
    Else
        '** 원본 ============================================================
        'tmpBarcode = CStr(Mid(txtBarcode.Text, 1, Len(txtBarcode.Text) - 1))
        '====================================================================
        
        '** 예수병원 ========================================================
        tmpBarcode = CStr(Mid(txtBarcode.Text, 1, Len(txtBarcode.Text)))
        '====================================================================
        
        tmpSQL = objMySql.SqlOrdersForAccess(2, Mid(tmpBarcode, 1, P_SpcYyLength), Val(Mid(tmpBarcode, P_SpcYyLength + 1)))
    End If
    
    Set objRs = New Recordset
    objRs.Open tmpSQL, DBConn
    
    If objRs.EOF Then
        DisplayOrder = False
        lblMsg.Caption = "해당 데이타가 없습니다 !!"
        If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "해당 데이타가 없습니다 !!"
        'MsgBox "해당 데이타가 없습니다 !!", vbOKOnly + vbExclamation, "Message"
        GoTo NoData
    End If
    
    txtWorkArea.Text = "" & objRs.Fields("WorkArea").Value
    txtAccDt.Text = Mid("" & objRs.Fields("AccDt").Value, 3)
    txtAccNo.Text = "" & objRs.Fields("AccSeq").Value
    
    lblLabNo.Caption = "" & objRs.Fields("WorkArea").Value & "-" & _
                        Mid(objRs.Fields("AccDt").Value, 3) & "-" & _
                        objRs.Fields("AccSeq").Value
    lblPtId.Caption = "" & objRs.Fields("PtId").Value
    
    ' Caution 즉시 호출
'    Call cmdCaution_Click
    
' 2009.08.20 양성현 검체접수시 이전 혈액검사의 혈액형을 가져와서 보여줌.
    tmpBLood = objMySql.SqlPtBloodType("" & objRs.Fields("PtId").Value)
    Set bRs = New Recordset
    bRs.Open tmpBLood, DBConn
    If Not bRs.EOF Then
        lblBLoodType.Caption = "" & bRs.Fields("blood").Value
        lblBLoodDT.Caption = Format("" & bRs.Fields("vfydtm").Value, "####-##-## ##:##:##")
    End If
    Set bRs = Nothing
    
    '감염관리
    Call ICSPatientMark(objRs.Fields("ptid").Value & "", enICSNum.LIS_ALL)
    
    
    lblPtNm.Caption = "" & objRs.Fields("PtNm").Value
    lblDeptNm.Caption = "" & objRs.Fields("DeptNm").Value
    lblWard.Caption = "" & objRs.Fields("Location").Value
    If objRs.Fields("StatFg").Value = "1" Then
        shpStat.Visible = True
        lblStat.Visible = True
    Else
        shpStat.Visible = False
        lblStat.Visible = False
    End If
    lblSpcNm.Caption = "" & objRs.Fields("SpcNm").Value

' 08.10.23. 양성현 검체 보관방법 명칭으로 변경

' 2015.09.17 온승호 검체보관방법에 따른 팝업 표기
    lblStoreNm.Caption = "" & objRs.Fields("storenm").Value
    'OT냉동 차광즉시냉동 차광냉장 OT냉장
    If lblStoreNm.Caption = "차광즉시냉동1" Or lblStoreNm.Caption = "차광즉시냉동2" Or lblStoreNm.Caption = "차광냉장" Or lblStoreNm.Caption = "즉시냉동4" Or lblStoreNm.Caption = "OT냉장" Or _
       lblStoreNm.Caption = "즉시냉동1" Or lblStoreNm.Caption = "즉시냉동2" Or lblStoreNm.Caption = "즉시냉동3" Or lblStoreNm.Caption = "즉시냉동" Or lblStoreNm.Caption = "OT냉동" Then
        lblStoreNm.ForeColor = vbRed
    Else
        lblStoreNm.ForeColor = vbBlack
    End If
    
    If objRs.Fields("StsCd").Value >= enStsCd.StsCd_LIS_Accession Then
        DisplayOrder = False
        lblMsg.Caption = "이미 접수된 검체입니다 !!"
        If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "이미 접수된 검체입니다(" & lblLabNo.Caption & ")"
        'MsgBox "이미 접수된 검체입니다 !!", vbOKOnly + vbExclamation, "Message"
    Else
        lblMsg.Caption = ""
        DisplayOrder = True
    End If
    
    With tblOrdSheet
        If objRs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
        Else
            .MaxRows = objRs.RecordCount
        End If
        For i = 1 To objRs.RecordCount
           .Row = i
           .Col = 1: .Value = objRs.Fields("OrdDt").Value & ""
           .Col = 2: .Value = objRs.Fields("TestNm").Value & ""
                     .ForeColor = DCM_LightBlue        '약간 파란색
           .Col = 3: .Value = objRs.Fields("OrdCd").Value & ""
           .Col = 4: .Value = objRs.Fields("mesg").Value & ""
           objRs.MoveNext
        Next
        .RowHeight(-1) = lngRowHeight
    End With
    Call tblOrdSheet_Click(1, 1)
    
    If lblStoreNm.Caption = "차광즉시냉동1" Or lblStoreNm.Caption = "차광즉시냉동2" Or lblStoreNm.Caption = "차광냉장" Or lblStoreNm.Caption = "즉시냉동4" Or lblStoreNm.Caption = "OT냉장" Or _
       lblStoreNm.Caption = "즉시냉동1" Or lblStoreNm.Caption = "즉시냉동2" Or lblStoreNm.Caption = "즉시냉동3" Or lblStoreNm.Caption = "즉시냉동" Or lblStoreNm.Caption = "OT냉동" Then
        MsgBox lblStoreNm.Caption & "을 요하는 검체입니다.", vbInformation, "검체보관확인"
    End If
NoData:
    Set objRs = Nothing

End Function

'% 접수Procedure를 수행한다.
Private Sub DoAccession(Optional ByVal ii As Integer = 0)

    Dim objAccess  As New clsLISAccession
    Dim blnSuccess As Boolean
    Dim strRcvDt   As String
    Dim strRcvTm   As String
      
    '데이타베이스의 날짜/시간으로 System Date/Time을 셋팅...
    Date = GetSystemDate
    strRcvDt = Format(GetSystemDate, "yyyymmdd")
    Time = GetSystemDate
    strRcvTm = Format(GetSystemDate, "hhmmss")
      
    MouseRunning  '13
      
    With objAccess
'        Call .SetDatabase(DbConn)
        blnSuccess = .DoAccession_New(txtWorkArea.Text, tmpAccDt, txtAccNo.Text, ObjMyUser.EmpId)
        If blnSuccess Then
            'lblMsg.Caption = "정상적으로 접수되었습니다 !!"
            If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "정상"
            
            '리스트뷰에 구해온번호 넣어놓구...
            '-----------------------------------------------------------------
            '접수가 정상적으로 수행되면 WorkArea 별 접수 순번을 번호 정보 테이블에 부여한다.
            '-- Parameter (WorkArea, 접수일자(RcvDt))
            If ii > 0 Then
                lblMsg.Caption = "정상적으로 접수되었습니다 !!"
                lstAccList.ListItems(ii).SubItems(3) = GetSeqNo(txtWorkArea.Text, strRcvDt, tmpAccDt, txtAccNo.Text)
            Else
                lblMsg.Caption = "정상적으로 접수되었습니다 !!" & "  일련번호:" & GetSeqNo(txtWorkArea.Text, strRcvDt, tmpAccDt, txtAccNo.Text)
            End If
            '-----------------------------------------------------------------
        Else
            lblMsg.Caption = "오류 발생 !! (" & lblLabNo.Caption & ")"
            If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "오류 발생 !!"
        End If
    End With
    
    '-- 업무나열서 순번(Rack 위치 확인 위한 일련번호)
    ' - 추가작업 By M.G.Choi
'    Call RackNo_Seq_Insert(txtWorkArea.Text, tmpAccDt, txtAccNo.Text, strRcvDt, strRcvTm)
    
    Set objAccess = Nothing
    
    MouseDefault
    
End Sub

'Private Sub RackNo_Seq_Insert(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccNo As String, _
'                              ByVal pRcvDt As String, ByVal pRcvTm As String)
'    Dim ObjAccess As New clsLISSqlAccession
'    Dim strSQL    As String
'
'    strSQL = ObjAccess.SqlRackNoInsert(pWorkArea, pAccDt, pAccNo, pRcvDt, pRcvTm, "")
'
'    Call dbconn.Execute(strSQL)
'
'    Set ObjAccess = Nothing
'
'End Sub
'
'% 워크애리어별 순번을 부여한다.
Private Function GetSeqNo(ByVal pWorkArea As String, ByRef pRcvDt As String, _
                          ByVal pAccDt As String, ByVal pAccSeq As String) As String

    Dim objSeq As clsLISSqlCollection
    Dim tmpRs As New Recordset
    Dim tmpSQL As String
    Dim LabDiv As String
    Dim tmpStr As String
    Dim tmpRng1 As Integer, tmpRng2 As Integer
    Dim tmpSpcGrp As String
    Dim tmpAccNo  As String

    GetSeqNo = 0

    tmpRng1 = 1
    tmpRng2 = 9999
    tmpSpcGrp = "0"

    Set objSeq = New clsLISSqlCollection
    
    tmpAccNo = pWorkArea & pAccDt & pAccSeq
    
    tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, tmpSpcGrp, 4, tmpAccNo)
    
    On Error GoTo Err_Trap

    DBConn.BeginTrans
    DBConn.Execute tmpSQL   'Lock 걸림

    '// Sql문장 생성
    tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, tmpSpcGrp, 5, tmpAccNo)
    '//
    tmpRs.Open tmpSQL, DBConn
    
    If tmpRs.EOF Then
        GetSeqNo = tmpRng1
        tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, GetSeqNo, 2, tmpAccNo, GetSeqNo)
    Else
        GetSeqNo = Val(tmpRs.Fields("Seq").Value & "")
        If GetSeqNo < tmpRng1 Then
            GetSeqNo = tmpRng1
        Else
            GetSeqNo = GetSeqNo + 1
        End If
        If GetSeqNo > tmpRng2 Then
            MainFrm.stsBar.Panels(2).Text = "접수순번의 범의 (" & tmpRng1 & "-" & tmpRng2 & ")를 벗어났읍니다. : " & GetSeqNo
            GoTo Err_Trap
        End If
        tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, GetSeqNo, 2, tmpAccNo, GetSeqNo)
    End If
    Set tmpRs = Nothing

    DBConn.Execute tmpSQL
    DBConn.CommitTrans

    Exit Function

Err_Trap:
    DBConn.RollbackTrans
    Set tmpRs = Nothing
    Set objSeq = Nothing
    GetSeqNo = 0
    Exit Function

End Function


'% Clear 루틴
Sub ClearRtn()
    txtMesg.Text = ""
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblDeptNm.Caption = ""
    lblWard.Caption = ""
    lblSpcNm.Caption = ""
    lblStoreNm.Caption = ""
    lblLabNo.Caption = ""
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    lblMsg.Caption = ""
    
     Call ICSPatientMark
End Sub

Private Sub txtWorkArea_Change()
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then
        If txtAccDt.Enabled Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtWorkArea_GotFocus()
    With txtWorkArea
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)
    Call ClearRtn
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If txtWorkArea = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccDt.SetFocus
End Sub
