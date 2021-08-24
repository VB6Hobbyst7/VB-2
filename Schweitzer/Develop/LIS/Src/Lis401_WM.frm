VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm401ResultView_WM 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9180
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   14715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis401_WM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
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
      Left            =   4650
      TabIndex        =   87
      Top             =   1110
      Width           =   7005
      Begin VB.CommandButton Command1 
         Caption         =   "종 료"
         Height          =   495
         Left            =   5250
         TabIndex        =   120
         Top             =   7245
         Width           =   1665
      End
      Begin VB.Frame Frame12 
         Caption         =   "특이소견"
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   118
         Top             =   5790
         Width           =   6795
         Begin RichTextLib.RichTextBox RichText 
            Height          =   540
            Left            =   150
            TabIndex        =   119
            Top             =   300
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   953
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"Lis401_WM.frx":08CA
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
         Height          =   1095
         Left            =   90
         TabIndex        =   114
         Top             =   4605
         Width           =   6795
         Begin VB.TextBox txtDrug 
            Height          =   315
            Left            =   180
            TabIndex        =   117
            Text            =   "Text1"
            Top             =   570
            Width           =   6465
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Penicillin"
            Height          =   225
            Index           =   21
            Left            =   180
            TabIndex        =   116
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RadioContrast"
            Height          =   225
            Index           =   22
            Left            =   1575
            TabIndex        =   115
            Top             =   225
            Width           =   1650
         End
      End
      Begin VB.Frame Frame10 
         Enabled         =   0   'False
         Height          =   795
         Left            =   90
         TabIndex        =   109
         Top             =   870
         Width           =   6795
         Begin VB.CommandButton cmdBMP1 
            BackColor       =   &H008080FF&
            Caption         =   "주의지침"
            Height          =   315
            Left            =   5820
            MaskColor       =   &H8000000F&
            Style           =   1  '그래픽
            TabIndex        =   136
            Top             =   390
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "신종감염병"
            Height          =   225
            Index           =   24
            Left            =   165
            TabIndex        =   126
            Top             =   510
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "홍역"
            Height          =   195
            Index           =   3
            Left            =   3855
            TabIndex        =   113
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "수두"
            Height          =   195
            Index           =   2
            Left            =   2565
            TabIndex        =   112
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tb"
            Height          =   225
            Index           =   1
            Left            =   1335
            TabIndex        =   111
            Top             =   210
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "AFB"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   110
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
         TabIndex        =   108
         Text            =   "Caution 수정은 감염관리실에 요청하여 주십시요."
         Top             =   6825
         Width           =   6795
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   90
         TabIndex        =   102
         Top             =   1680
         Width           =   6810
         Begin VB.CommandButton cmdBMP2 
            BackColor       =   &H008080FF&
            Caption         =   "주의지침"
            Height          =   315
            Left            =   5820
            MaskColor       =   &H8000000F&
            Style           =   1  '그래픽
            TabIndex        =   138
            Top             =   390
            Width           =   915
         End
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            Height          =   195
            Index           =   25
            Left            =   180
            TabIndex        =   127
            Top             =   510
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VDRL"
            Height          =   225
            Index           =   5
            Left            =   1305
            TabIndex        =   107
            Top             =   225
            Width           =   900
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HBsAg"
            Height          =   225
            Index           =   6
            Left            =   2565
            TabIndex        =   106
            Top             =   225
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HIV"
            Height          =   225
            Index           =   4
            Left            =   180
            TabIndex        =   105
            Top             =   225
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HCV"
            Height          =   195
            Index           =   7
            Left            =   3870
            TabIndex        =   104
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HBc IgM"
            Height          =   225
            Index           =   8
            Left            =   1260
            TabIndex        =   103
            Top             =   510
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   90
         TabIndex        =   91
         Top             =   2505
         Width           =   6810
         Begin VB.CommandButton cmdBMP3 
            BackColor       =   &H008080FF&
            Caption         =   "주의지침"
            Height          =   315
            Left            =   5820
            MaskColor       =   &H8000000F&
            Style           =   1  '그래픽
            TabIndex        =   137
            Top             =   810
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            Height          =   225
            Index           =   29
            Left            =   3855
            TabIndex        =   131
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CJD"
            Height          =   225
            Index           =   28
            Left            =   2595
            TabIndex        =   130
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRSA"
            Height          =   225
            Index           =   27
            Left            =   1560
            TabIndex        =   129
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRE"
            Height          =   225
            Index           =   26
            Left            =   135
            TabIndex        =   128
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rotavirus"
            Height          =   225
            Index           =   14
            Left            =   135
            TabIndex        =   101
            Top             =   585
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HAVIgM"
            Height          =   225
            Index           =   9
            Left            =   135
            TabIndex        =   100
            Top             =   240
            Width           =   1380
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRSA"
            Height          =   225
            Index           =   10
            Left            =   1560
            TabIndex        =   99
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRE"
            Height          =   225
            Index           =   11
            Left            =   2595
            TabIndex        =   98
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "C.diffic"
            Height          =   225
            Index           =   12
            Left            =   3855
            TabIndex        =   97
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRAB(IRAB)"
            Height          =   225
            Index           =   13
            Left            =   5205
            TabIndex        =   96
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "옴"
            Height          =   225
            Index           =   15
            Left            =   1560
            TabIndex        =   95
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "이"
            Height          =   225
            Index           =   16
            Left            =   2595
            TabIndex        =   94
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "장티푸스"
            Height          =   225
            Index           =   17
            Left            =   3855
            TabIndex        =   93
            Top             =   585
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "세균성이질"
            Height          =   225
            Index           =   18
            Left            =   5205
            TabIndex        =   92
            Top             =   585
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   825
         Left            =   90
         TabIndex        =   88
         Top             =   3750
         Width           =   6810
         Begin VB.CommandButton cmdBMP4 
            BackColor       =   &H008080FF&
            Caption         =   "주의지침"
            Height          =   315
            Left            =   5820
            MaskColor       =   &H8000000F&
            Style           =   1  '그래픽
            TabIndex        =   139
            Top             =   420
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "유행성이하선염"
            Height          =   225
            Index           =   33
            Left            =   4410
            TabIndex        =   135
            Top             =   225
            Width           =   1875
         End
         Begin VB.CheckBox Check1 
            Caption         =   "기타"
            Height          =   225
            Index           =   32
            Left            =   3420
            TabIndex        =   134
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "수막구균수막염"
            Height          =   225
            Index           =   31
            Left            =   1575
            TabIndex        =   133
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Caption         =   "백일해"
            Height          =   225
            Index           =   30
            Left            =   135
            TabIndex        =   132
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "인플루엔자"
            Height          =   225
            Index           =   23
            Left            =   2790
            TabIndex        =   125
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "신종플루"
            Height          =   225
            Index           =   19
            Left            =   135
            TabIndex        =   90
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "풍진"
            Height          =   225
            Index           =   20
            Left            =   1575
            TabIndex        =   89
            Top             =   225
            Width           =   1335
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   121
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
         TabIndex        =   122
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
         TabIndex        =   123
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
         TabIndex        =   124
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   3750
      Left            =   1980
      ScaleHeight     =   3690
      ScaleWidth      =   10350
      TabIndex        =   84
      Top             =   2250
      Visible         =   0   'False
      Width           =   10410
      Begin VB.CommandButton Command2 
         Caption         =   "종료"
         Height          =   600
         Left            =   8505
         TabIndex        =   85
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
         TabIndex        =   86
         Top             =   270
         Width           =   9735
      End
   End
   Begin VB.PictureBox picOrder 
      BackColor       =   &H00F3F5F8&
      Height          =   795
      Left            =   -15
      ScaleHeight     =   735
      ScaleWidth      =   7935
      TabIndex        =   46
      Top             =   1560
      Width           =   7995
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00CCFFFF&
         Caption         =   "Re&fresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6780
         MaskColor       =   &H00C0FFFF&
         Style           =   1  '그래픽
         TabIndex        =   57
         Tag             =   "128"
         Top             =   195
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F3F5F8&
         BorderStyle     =   0  '없음
         Height          =   240
         Left            =   3645
         TabIndex        =   53
         Top             =   30
         Width           =   2940
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00F3F5F8&
            Caption         =   "보고일"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -30
            TabIndex        =   56
            Tag             =   "15305"
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00F3F5F8&
            Caption         =   "접수일"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   870
            TabIndex        =   55
            Tag             =   "15304"
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00F3F5F8&
            Caption         =   "처방일"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1785
            TabIndex        =   54
            Tag             =   "15305"
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00F3F5F8&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   3525
         TabIndex        =   47
         Top             =   75
         Width           =   3525
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00F4FDF5&
            Caption         =   "혈액은행"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   345
            Index           =   2
            Left            =   2550
            Style           =   1  '그래픽
            TabIndex        =   51
            Top             =   210
            Width           =   840
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00FFF2EE&
            Caption         =   "임상병리"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   345
            Index           =   0
            Left            =   855
            Style           =   1  '그래픽
            TabIndex        =   50
            Top             =   210
            Width           =   840
         End
         Begin VB.CheckBox ChkDivAll 
            BackColor       =   &H00F3F5F8&
            Caption         =   "전체"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C56152&
            Height          =   225
            Left            =   60
            TabIndex        =   49
            Top             =   300
            Width           =   780
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00EFFCFC&
            Caption         =   "미생물"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   345
            Index           =   3
            Left            =   1695
            Style           =   1  '그래픽
            TabIndex        =   48
            Top             =   210
            Width           =   840
         End
         Begin VB.Label lblOrders 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Orders"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   15
            TabIndex        =   52
            Tag             =   "155"
            Top             =   -30
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   300
         Left            =   3660
         TabIndex        =   58
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24248323
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   300
         Left            =   5220
         TabIndex        =   59
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24248323
         CurrentDate     =   36328
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   225
         Left            =   5055
         TabIndex        =   60
         Tag             =   "40110"
         Top             =   270
         Width           =   105
      End
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6600
      Left            =   15
      TabIndex        =   62
      Top             =   2400
      Width           =   7980
      _Version        =   196608
      _ExtentX        =   14076
      _ExtentY        =   11642
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
      GrayAreaBackColor=   14411494
      MaxCols         =   35
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "Lis401_WM.frx":0949
      TextTip         =   4
   End
   Begin VB.PictureBox picPtList 
      Align           =   3  '왼쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00D7E6E6&
      DragMode        =   1  '자동
      Height          =   7620
      Left            =   0
      ScaleHeight     =   7560
      ScaleWidth      =   4185
      TabIndex        =   0
      Top             =   1560
      Width           =   4245
      Begin VB.Frame fraSearch 
         BackColor       =   &H00D7E6E6&
         Height          =   645
         Left            =   0
         TabIndex        =   3
         Tag             =   "136"
         Top             =   600
         Width           =   4200
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "테"
            Top             =   240
            Width           =   1830
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00D7E6E6&
            Caption         =   "&Name"
            Height          =   255
            Index           =   1
            Left            =   2505
            TabIndex        =   5
            Tag             =   "15305"
            Top             =   285
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00D7E6E6&
            Caption         =   "&ID"
            Height          =   240
            Index           =   0
            Left            =   1995
            TabIndex        =   4
            Tag             =   "15304"
            Top             =   300
            Width           =   495
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   285
            Index           =   1
            Left            =   3465
            Shape           =   4  '둥근 사각형
            Top             =   255
            Width           =   675
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
            Height          =   225
            Left            =   3570
            MouseIcon       =   "Lis401_WM.frx":60D3
            MousePointer    =   99  '사용자 정의
            TabIndex        =   7
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.CheckBox chkVerified 
         BackColor       =   &H00D7E6E6&
         Caption         =   "금일 결과보고 대상만 검색"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   225
         Left            =   1635
         TabIndex        =   2
         Top             =   375
         Width           =   2460
      End
      Begin VB.CheckBox chkAllWard 
         BackColor       =   &H00D7E6E6&
         Caption         =   "전체병동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   1
         Top             =   75
         Width           =   1035
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6300
         Left            =   15
         TabIndex        =   8
         Top             =   1230
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   11113
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16643054
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblWardId 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "병동선택"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   180
         Left            =   2730
         MouseIcon       =   "Lis401_WM.frx":63DD
         MousePointer    =   99  '사용자 정의
         TabIndex        =   9
         ToolTipText     =   "Click하시면 마감시간을 수정할 수 있습니다."
         Top             =   90
         Width           =   1320
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         FillColor       =   &H00E8F7F7&
         FillStyle       =   0  '단색
         Height          =   270
         Left            =   2685
         Shape           =   4  '둥근 사각형
         Top             =   45
         Width           =   1395
      End
      Begin VB.Label lblPtList 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Patient List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Tag             =   "106"
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.PictureBox picResult 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3F5F8&
      Height          =   660
      Left            =   8010
      ScaleHeight     =   600
      ScaleWidth      =   6720
      TabIndex        =   16
      Top             =   1560
      Width           =   6780
      Begin VB.CheckBox chkSize 
         BackColor       =   &H00F3F5F8&
         Caption         =   "참고치보기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3420
         TabIndex        =   63
         Top             =   345
         Width           =   1305
      End
      Begin VB.CheckBox chkRstCmt 
         BackColor       =   &H00F3F5F8&
         Caption         =   "Result   Comment"
         ForeColor       =   &H00553755&
         Height          =   255
         Left            =   4785
         TabIndex        =   18
         Tag             =   "40204"
         Top             =   330
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.CheckBox chkSamCmt 
         BackColor       =   &H00F3F5F8&
         Caption         =   "Sample Comment"
         ForeColor       =   &H00553755&
         Height          =   255
         Left            =   4785
         TabIndex        =   17
         Tag             =   "40205"
         Top             =   75
         Value           =   1  '확인
         Width           =   1815
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   300
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   450
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
         Caption         =   "검체 "
         Appearance      =   0
      End
      Begin VB.Label lblResults 
         BackStyle       =   0  '투명
         Caption         =   "Results  -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Tag             =   "19908"
         Top             =   45
         Width           =   990
      End
      Begin VB.Label lblWorkArea 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "Chemistry"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DF6A3E&
         Height          =   225
         Left            =   1305
         TabIndex        =   20
         Top             =   60
         Width           =   1110
      End
      Begin VB.Label lblSpecimenNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Serum"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   840
         TabIndex        =   19
         Top             =   345
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00DFE3E8&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   0
      ScaleHeight     =   1560
      ScaleWidth      =   14715
      TabIndex        =   22
      Top             =   0
      Width           =   14715
      Begin VB.CommandButton cmdCaution 
         BackColor       =   &H008080FF&
         Caption         =   "Caution"
         Height          =   315
         Left            =   2670
         MaskColor       =   &H8000000F&
         Style           =   1  '그래픽
         TabIndex        =   83
         Top             =   90
         Width           =   885
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   270
         Index           =   12
         Left            =   9570
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   476
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
         Caption         =   "처방일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   13
         Left            =   9570
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   14
         Left            =   9570
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   645
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   15
         Left            =   9570
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   930
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   11
         Left            =   6600
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   930
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "보 고 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   10
         Left            =   6600
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   645
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "접 수 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   9
         Left            =   6600
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "채 혈 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   270
         Index           =   8
         Left            =   6600
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   476
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
         Caption         =   "처 방 의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   5
         Left            =   3630
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   645
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "입 원 일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   4
         Left            =   3630
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "병     실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   270
         Index           =   3
         Left            =   3630
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   476
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
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   27
         Top             =   135
         Width           =   1245
      End
      Begin VB.CheckBox chkPtList 
         BackColor       =   &H00DFE3E8&
         Caption         =   "환자검색 리스트(&S)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Tag             =   "40101"
         Top             =   1245
         Width           =   2460
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
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
         Height          =   420
         Left            =   12990
         Style           =   1  '그래픽
         TabIndex        =   25
         Tag             =   "128"
         Top             =   1065
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
         Height          =   420
         Left            =   12990
         Style           =   1  '그래픽
         TabIndex        =   24
         Tag             =   "40102"
         Top             =   615
         Width           =   1320
      End
      Begin VB.PictureBox picESign 
         Height          =   500
         Left            =   12975
         ScaleHeight     =   435
         ScaleWidth      =   1140
         TabIndex        =   23
         Top             =   45
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   1350
         TabIndex        =   28
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
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
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblReceiverNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   29
         Top             =   630
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCollectorNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   30
         Top             =   345
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCollectDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   31
         Top             =   345
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifierNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   32
         Top             =   915
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   33
         Top             =   60
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblReceiveDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   34
         Top             =   645
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifyDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   35
         Top             =   930
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   36
         Top             =   60
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   255
         Left            =   4620
         TabIndex        =   37
         Top             =   360
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblBedoutDt 
         Height          =   255
         Left            =   4620
         TabIndex        =   38
         Top             =   930
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblBedinDt 
         Height          =   255
         Left            =   4620
         TabIndex        =   39
         Top             =   645
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   270
         Left            =   4620
         TabIndex        =   40
         Top             =   60
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   476
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   210
         Left            =   4620
         TabIndex        =   41
         Top             =   1245
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   370
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
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   450
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
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
         Caption         =   "성  명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   780
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
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
         Height          =   255
         Index           =   6
         Left            =   3630
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   930
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "퇴 원 일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   7
         Left            =   3630
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1215
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   17
         Left            =   9555
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1215
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "확 인 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblLisDoctNm 
         Height          =   255
         Left            =   10545
         TabIndex        =   82
         Top             =   1215
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Label lblSex 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "여자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1455
         TabIndex        =   44
         Top             =   855
         Width           =   360
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   42
         Top             =   840
         Width           =   120
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   43
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00F3F5F8&
         Caption         =   "            /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         TabIndex        =   45
         Top             =   780
         Width           =   2010
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   285
         Left            =   1350
         Shape           =   4  '둥근 사각형
         Top             =   120
         Width           =   1290
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         FillColor       =   &H00DFE3E8&
         FillStyle       =   0  '단색
         Height          =   1545
         Left            =   45
         Shape           =   4  '둥근 사각형
         Top             =   0
         Width           =   14640
      End
   End
   Begin VB.PictureBox picFootNote 
      Appearance      =   0  '평면
      BackColor       =   &H00EFFEFE&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   8010
      ScaleHeight     =   930
      ScaleWidth      =   6660
      TabIndex        =   13
      Top             =   8025
      Width           =   6690
      Begin RichTextLib.RichTextBox txtSamCmt 
         Height          =   960
         Left            =   75
         TabIndex        =   14
         Top             =   30
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   1693
         _Version        =   393217
         BackColor       =   15728382
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis401_WM.frx":66E7
         MouseIcon       =   "Lis401_WM.frx":678C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picRstText 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFF7&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   8010
      ScaleHeight     =   1515
      ScaleWidth      =   6660
      TabIndex        =   11
      Top             =   6465
      Width           =   6690
      Begin RichTextLib.RichTextBox txtRstCmt 
         Height          =   1440
         Left            =   75
         TabIndex        =   12
         Top             =   45
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   2540
         _Version        =   393217
         BackColor       =   16777207
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis401_WM.frx":68EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   4230
      Left            =   8010
      TabIndex        =   15
      Top             =   2205
      Width           =   6690
      _Version        =   196608
      _ExtentX        =   11800
      _ExtentY        =   7461
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
      BackColorStyle  =   3
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis401_WM.frx":6993
      UnitType        =   0
      UserResize      =   0
      VisibleCols     =   8
      VisibleRows     =   22
      TextTip         =   4
   End
   Begin RichTextLib.RichTextBox rtfResult 
      Height          =   7470
      Left            =   7995
      TabIndex        =   61
      Top             =   1530
      Visible         =   0   'False
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   13176
      _Version        =   393217
      BackColor       =   16777207
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      TextRTF         =   $"Lis401_WM.frx":85E3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm401ResultView_WM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objTB     As frmTBReport
Attribute objTB.VB_VarHelpID = -1

Private objPatient  As clsPatient              '환자 클래스
Private objSQL      As clsLISSqlReview       'Sql문 클래스

Private OldOrdDiv   As String
Private aryMesg()   As String
Private svWorkArea  As String
Private svAccDt     As String
Private svAccSeq    As String
Private SvTestCd    As String
Private mvarDeptCd  As String
Private ClearFg     As Boolean
Private OrderFg     As Boolean
Private ResultFg    As Boolean
Private MsgFg       As Boolean
Private PtFg         As Boolean
Private QueryFg      As Boolean

Private OldRow      As Long
Private Const lngMaxRows = 29
Private Const lngRowHeight = 11.5

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuPrint As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_PRT& = 1

Public Event LastFormUnload()
Public Event ThisFormUnload()

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Private Sub chkAllWard_Click()
    If chkAllWard.Value = 0 Then
        chkVerified.Enabled = True
        If lblWardId.Caption <> "" Then
            Call PtList_Display(lblWardId.Caption)
        Else
            lvwPtList.ListItems.Clear
        End If
    Else
        chkVerified.Enabled = False
        chkVerified.Value = 1
        
        Call PtList_Display(lblWardId.Caption)
    End If
End Sub

Private Sub ChkDivAll_Click()
    If ChkDivAll.Value = 1 Then
        optOrdDiv(0).Value = False
        optOrdDiv(2).Value = False
        optOrdDiv(3).Value = False
    Else
        optOrdDiv(0).Value = True
    End If
End Sub

Private Sub chkSize_Click()
    Dim strTmp As String
    
    With tblResult
        .Row = -1: .Col = 1
        .BlockMode = True: strTmp = .Clip: .BlockMode = False
        If chkSize.Value = 1 Then
            .ColWidth(3) = 10
        Else
            .Row = -1: .Col = 3
            .ColWidth(3) = 30: .ColWidth(2) = 30: .ColWidth(2) = 15
        End If
        .Row = -1: .Col = 1
        .BlockMode = True: .Clip = strTmp:
        .AllowCellOverflow = True
        .BlockMode = False
        .Refresh
    End With

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
            
            Frame2.Visible = True
            If Check1(4).Value = 1 Then
                Picture1.Visible = True
            Else
                Picture1.Visible = False
            End If
        Else
            Frame2.Visible = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing

End Sub

Private Sub Command1_Click()
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame2.Visible = False
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub lblWardId_Click()
'    Dim objWard As clsBasisData
'
'    Set objWard = New clsBasisData
    Set objMyList = New clsPopUpList
    
    lblWardId.Caption = "    "
    With objMyList
        .Connection = DBConn
        .FormCaption = "병동 조회": .ColumnHeaderText = "병동코드;병동명"
        Call .LoadPopUp(GetSQLWardList) ', 1640, 10550)  ', ObjLISComCode.WardId)
        If .SelectedString <> "" Then
            lblWardId.Caption = Trim(medGetP(.SelectedString, 1, ";"))
            lblWardId.Tag = "1"
            Call PtList_Display(lblWardId.Caption)
            mvarDeptCd = lblWardId.Caption
            chkVerified.Enabled = True
            If chkVerified.Value = 1 Then Call txtSearchKey_KeyPress(vbKeyReturn)
        End If
    End With
    Set objMyList = Nothing
'    Set objWard = Nothing
End Sub

Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer
        
    '-- 정렬
    With lvwPtList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

'Private Sub mnuPrint_Click()
'    Dim MyReport    As clsBatchReport
'    Dim RS          As Recordset
'
'    Dim pWorkArea   As String
'    Dim pAccDt      As String
'    Dim pAccSeq     As String
'    Dim lngseq      As Integer
'
'    Set MyReport = New clsBatchReport
'
'    With tblOrdSheet
'        .Row = OldRow
'        .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value
'        .Col = enREVIEW1.tcACCDT:    pAccDt = .Value
'        .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value
'    End With
'
'    Set RS = New Recordset
'    RS.Open objSQL.SqlMultiTest(pWorkArea, pAccDt, Val(pAccSeq)), dbconn
'
'    If Not RS.EOF Then
'        While Not RS.EOF
'            lngseq = lngseq + 1
'            pWorkArea = "" & RS.Fields("WorkArea").Value
'            pAccDt = "" & RS.Fields("AccDt").Value
'            pAccSeq = "" & RS.Fields("AccSeq").Value
'            Set MyReport = Nothing
'            Set MyReport = New clsBatchReport
'            MyReport.PtId = txtPtId.Text
'            MyReport.PtNm = lblPtNm.Caption
'            MyReport.PtSex = lblSex.Caption
'            MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
'            MyReport.DeptNm = lblDeptNm.Caption
'            MyReport.VfyDt = lblVerifyDt.Caption
'            MyReport.VfyNM = lblVerifierNm.Caption
'            MyReport.ICD = lblDisease.Caption
'            Call MyReport.MicSensiReport(pWorkArea, pAccDt, pAccSeq, picESign)
'            RS.MoveNext
'        Wend
'        Set RS = Nothing
'    Else
'        MyReport.PtId = txtPtId.Text
'        MyReport.PtNm = lblPtNm.Caption
'        MyReport.PtSex = lblSex.Caption
'        MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
'        MyReport.DeptNm = lblDeptNm.Caption
'        MyReport.VfyDt = lblVerifyDt.Caption
'        MyReport.VfyNM = lblVerifierNm.Caption
'        MyReport.ICD = lblDisease.Caption
'        Call MyReport.MicSensiReport(pWorkArea, pAccDt, pAccSeq, picESign)
'    End If
'    Set RS = Nothing
'    Set MyReport = Nothing
'
'End Sub

'-- 해당병동의 재원환자 리스트 Display
Private Sub PtList_Display(Optional pWardID As String = "")
    Dim RS        As Recordset
    Dim strSQL    As String
    Dim strOutDt  As String
    Dim strPtId   As String
    Dim Jumin1    As String
    Dim Jumin2    As String
    Dim LvwItem   As ListItem
    
    strOutDt = Format(GetSystemDate, "yyyymmdd")
    
    If chkAllWard.Value = 1 Then
        If chkVerified = 1 Then
            strSQL = objSQL.SqlWard_PtList(pWardID, strOutDt, True, strOutDt)
        Else
            strSQL = objSQL.SqlWard_PtList(pWardID, strOutDt, True)
        End If
    Else
        If chkVerified = 1 Then
            strSQL = objSQL.SqlWard_PtList(pWardID, strOutDt, False, strOutDt)
        Else
            strSQL = objSQL.SqlWard_PtList(pWardID, strOutDt, False)
        End If
    End If
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    Me.MousePointer = vbHourglass
    
    lvwPtList.ListItems.Clear
    If RS.BOF = False Then
        With lvwPtList
            
            Do Until RS.EOF = True
                Set LvwItem = .ListItems.Add()
                strPtId = Trim(RS.Fields("hospno").Value) & ""
                '-- 환자정보 Set
                Call objPatient.GETPatient(strPtId)
                With objPatient
                     LvwItem.Text = .PTid
                     LvwItem.SubItems(1) = .PtNm
                     LvwItem.SubItems(2) = .WardId & "-" & .RoomId
                     If .SSN <> "" Then
                         Jumin1 = Mid(.SSN, 1, 6)
                         Jumin2 = Mid(.SSN, 7, 7)
                         LvwItem.SubItems(5) = Jumin1 & "-" & Jumin2
                     End If
                     LvwItem.SubItems(3) = .Dob
                     LvwItem.SubItems(4) = .Sex & "/" & .Age
                     LvwItem.SubItems(6) = .ADDR ' .Addr1 & " " & .Addr2
                End With
                
                RS.MoveNext
            Loop
        End With
    End If
    
    Me.MousePointer = vbDefault
    Set RS = Nothing
End Sub

'% 환자리스트 Display 여부
Private Sub chkPtList_Click()

On Error GoTo Errors
    If chkPtList.Value = 1 Then
        lblWardId.Caption = mvarDeptCd
        picPtList.Visible = True
        picPtList.Width = 4290
        picOrder.Left = picPtList.Width
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picPtList.Width + picOrder.Width
        tblResult.Left = picResult.Left
        picRstText.Left = picResult.Left
        picFootNote.Left = picResult.Left
        txtSearchKey.SetFocus
    ElseIf chkPtList.Value = 0 Then
        picPtList.Visible = False
        picOrder.Left = 0
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picOrder.Width
        tblResult.Left = picResult.Left
        picRstText.Left = picResult.Left
        picFootNote.Left = picResult.Left
    End If
Errors:

End Sub

'% 텍스트 결과내역 박스 Display 여부
Private Sub chkRstCmt_Click()
    If chkRstCmt.Value = 1 And picRstText.Visible = False Then
        picRstText.Visible = True
        tblResult.Height = tblResult.Height - picRstText.Height
    ElseIf chkRstCmt.Value = 0 And picRstText.Visible = True Then
        picRstText.Visible = False
        tblResult.Height = tblResult.Height + picRstText.Height
    End If
End Sub

'% 풋노트, 검체리마크 박스 Display 여부
Private Sub chkSamCmt_Click()
    If chkSamCmt.Value = 1 And picFootNote.Visible = False Then
        picFootNote.Visible = True
        tblResult.Height = tblResult.Height - picFootNote.Height
        picRstText.Top = picRstText.Top - picFootNote.Height
    ElseIf chkSamCmt.Value = 0 And picFootNote.Visible = True Then
        picFootNote.Visible = False
        tblResult.Height = tblResult.Height + picFootNote.Height
        picRstText.Top = picRstText.Top + picFootNote.Height
    End If
End Sub

Private Sub chkVerified_Click()
    Call PtList_Display(lblWardId.Caption)
End Sub

Private Sub cmdClear_Click()
    cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
    Call ClearRtn
    txtPtId.Text = ""
On Error GoTo Errors
    txtPtId.SetFocus
Errors:

End Sub

'%종료
Private Sub cmdExit_Click()
    Unload Me
    Set objSQL = Nothing
    Set objTB = Nothing
    Set objMyList = Nothing
    Set objPatient = Nothing
    RaiseEvent ThisFormUnload
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Public Sub cmdRefresh_Click()
    OldRow = 0
    OrderFg = False
    Call dtpToDate_KeyDown(vbKeyReturn, 0)
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
End Sub

'% 조회기간 입력 (From Date)
Private Sub dtpFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errors
    If KeyCode = vbKeyReturn Then dtpToDate.SetFocus
Errors:
End Sub


Private Sub dtpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Errors

    If KeyCode = vbKeyReturn Then
        If Format(dtpToDate.Value, CS_DateDbFormat) < Format(dtpFromDate.Value, CS_DateDbFormat) Then
            MsgBox "기간 입력 오류입니다. 날짜를 조정하십시오..", vbExclamation, "입력오류"
            dtpFromDate.SetFocus
            Exit Sub
        End If
        '% 처방조회
        cmdRefresh.Enabled = False
        dtpFromDate.Enabled = False
        dtpToDate.Enabled = False
        
        Call FieldClear
        Call DisplayOrders
        
        ResultFg = False
        cmdRefresh.Enabled = True
        dtpFromDate.Enabled = True
        dtpToDate.Enabled = True
        
        If OrderFg Then
            tblOrdSheet.SetFocus
        Else
            dtpFromDate.SetFocus
        End If
    End If
    Exit Sub
    
Errors:
    Resume Next
End Sub

'% 환자ID, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
Private Sub DisplayOrders()
    Dim objStatus   As jProgressBar.clsProgress
    Dim RS          As Recordset
    Dim tmpRs       As Recordset
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvDoctNm    As String
    Dim SvSpcNm     As String
    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim strStsCd    As String
    Dim strStsNm    As String

    Dim strOrdDiv   As String
    Dim strTestDiv  As String
    Dim strUnit     As String
    Dim strSelDiv   As String
    Dim strTestcd   As String
    Dim strKeyFld   As String
    
    Dim I           As Integer
    Dim ColCnt      As Integer
    Dim RecordCnt   As Integer
    
    Dim lngColor    As Long
    Dim iBtnFg      As Long
    Dim pForeColor   As Long
    
   
    QueryFg = True
    
    Call TableClear
    Call ResultClear

    Me.Enabled = False
    Call MouseRunning  '13
   
    DoEvents

    Set RS = New Recordset
    Set tmpRs = New Recordset
    Set objStatus = New jProgressBar.clsProgress
    
    With objStatus
        .Container = Me
        .Height = 280
        .Width = tblOrdSheet.Width
        .Left = tblOrdSheet.Left
        .Top = tblOrdSheet.Top
        .Message = lblPtNm.Caption & " 님의 처방내역을 검색중입니다..."
        .Max = 100
        .Value = 1
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = tblOrdSheet.Width
'        .XPos = tblOrdSheet.Left
'        .YPos = tblOrdSheet.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = lblPtNm.Caption & " 님의 처방내역을 검색중입니다..."
'        .Value = 1
    End With
    
    chkSize.Value = 0
    DoEvents
    
    '처방일/접수일/보고일 기준
    
    If optQueryKey(0).Value = True Then
        strKeyFld = "rcvdt"
    ElseIf optQueryKey(1).Value = True Then
        strKeyFld = "orddt"
    Else
        strKeyFld = "examdt"
    End If
        
    'pooh 수정  0-전체, 1-임상, 2-해부, 4-혈액
    If ChkDivAll.Value = 1 Then
        strSelDiv = "0"
    Else
        If optOrdDiv(0).Value = True Then
            strSelDiv = "1"
        ElseIf optOrdDiv(3).Value = True Then   '미생물만
            strSelDiv = "7"
        Else
            strSelDiv = "4"
        End If
    End If
        
    SqlStmt = objSQL.SqlQueryOrders_WM(txtPtId.Text, strKeyFld, Format(dtpFromDate.Value, CS_DateDbFormat), _
                                   Format(dtpToDate.Value, CS_DateDbFormat), strSelDiv)
    
On Error GoTo Errors
    
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo Errors
    
    objStatus.Max = 100
    objStatus.Value = objStatus.Max
   
    SvOrdDt = "": SvOrdNo = "": SvDoctNm = "": SvSpcNm = ""
    Erase aryMesg
   
    With tblOrdSheet
        .MaxRows = 0
        RecordCnt = 0
      
        Do Until tmpRs.EOF
            RecordCnt = RecordCnt + 1
            ReDim Preserve aryMesg(RecordCnt)   'Message Array ...
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDate").Value) Then
                .Col = enREVIEW1.tcORDDT:   .Value = Trim("" & tmpRs.Fields("OrdDate").Value)    '처방일
                .Col = enREVIEW1.tcORDNO:   .Value = Trim("" & tmpRs.Fields("OrdNo").Value)      '처방번호
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     '처방의
                SvOrdDt = Trim("" & tmpRs.Fields("OrdDate").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvDoctNm = Trim("" & tmpRs.Fields("DoctNm").Value)
            End If
            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
                .Col = enREVIEW1.tcORDNO:   .Value = Trim("" & tmpRs.Fields("OrdNo").Value)      '처방번호
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     '처방의
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvDoctNm = Trim("" & tmpRs.Fields("DoctNm").Value)
            End If
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            End If
            If SvDoctNm <> Trim("" & tmpRs.Fields("DoctNm").Value) Then
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     '처방의
                SvDoctNm = Trim("" & tmpRs.Fields("DoctNm").Value)
            End If
            
            .Col = enREVIEW1.tcTESTNM:    .Value = Trim("" & tmpRs.Fields("TestNm").Value)        '검사명
            .Col = enREVIEW1.tcSTATFG:    .Value = Choose(Val("" & tmpRs.Fields("StatFg").Value) + 1, " ", "Y")     '응급여부
            .Col = enREVIEW1.tcRCVDT:     .Value = "" & Format(Format(tmpRs.Fields("RcvDt"), CS_DateMask), "YY/MM/DD") & " " & _
                                                                      tmpRs.Fields("RcvTm")                   '접수일시
            If "" & tmpRs.Fields("OrdDtm") = "" Then
                .Col = enREVIEW1.tcORDDATE:   .Value = "" & Format(Format(tmpRs.Fields("OrdDt"), CS_DateMask), "YY/MM/DD") & " " & _
                                                                          tmpRs.Fields("OrdTm")                   '처방일시
            Else
                .Col = enREVIEW1.tcORDDATE:   .Value = "" & Format(Format(tmpRs.Fields("OrdDtm"), CS_DateMask), "YY/MM/DD") & " " & _
                                                                          tmpRs.Fields("OrdTm")                   '처방일시
            End If
            .Col = enREVIEW1.tcORDDOCT:   .Value = Trim("" & tmpRs.Fields("DoctNm").Value)        '처방의
            .Col = enREVIEW1.tcSPCNAME:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
            .Col = enREVIEW1.tcORDNUM:    .Value = Trim("" & tmpRs.Fields("OrdNo").Value)         '처방번호
            .Col = enREVIEW1.tcWORKAREA:  .Value = Trim("" & tmpRs.Fields("WorkArea").Value): pWorkArea = .Value    'WorkArea
            .Col = enREVIEW1.tcACCDT:     .Value = Trim("" & tmpRs.Fields("AccDt").Value):    pAccDt = .Value       'AccDt
            .Col = enREVIEW1.tcACCSEQ:    .Value = Trim("" & tmpRs.Fields("AccSeq").Value):   pAccSeq = .Value      'AccSeq
            .Col = enREVIEW1.tcVFYNM:     .Value = Trim("" & tmpRs.Fields("ExamNm").Value)
            .Col = enREVIEW1.tcVFYDATE:   .Value = Trim("" & tmpRs.Fields("ExamDt").Value) & " " & _
                                                   Trim("" & tmpRs.Fields("ExamTm").Value)                          '보고일시
            .Col = enREVIEW1.tcTESTCD:    .Value = Trim("" & tmpRs.Fields("OrdCd").Value)         '처방코드
            .Col = enREVIEW1.tcSPCCD:     .Value = Trim("" & tmpRs.Fields("SpcCd").Value)                           '검체코드
            .Col = enREVIEW1.tcSPCYY:     .Value = Trim("" & tmpRs.Fields("SpcYy").Value)                           '검체년도
            .Col = enREVIEW1.tcSPCNO:     .Value = Trim("" & tmpRs.Fields("SpcNo").Value)                           '검체번호
            .Col = enREVIEW1.tcORDDIV:    .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)                          '처방구분
            .Col = enREVIEW1.tcUNITQTY:   .Value = Trim("" & tmpRs.Fields("UnitQty").Value): strUnit = .Value       '수혈수량
            .Col = enREVIEW1.tcREQDATE:   .Value = Trim("" & tmpRs.Fields("ReqDt").Value)         '수혈예정일
            .Col = enREVIEW1.tcREQTIME:   .Value = Trim("" & tmpRs.Fields("ReqTm").Value)         '수혈예정시간
            .Col = enREVIEW1.tcWARDID:    .Value = Trim("" & tmpRs.Fields("WardId").Value)        '병동
            .Col = enREVIEW1.tcHOSILID:   .Value = Trim("" & tmpRs.Fields("HosilId").Value)       '호실
            .Col = 31:   .Value = Trim("" & tmpRs.Fields("PanelFg").Value)       '그룹여부
            .Col = 32:   .Value = Trim("" & tmpRs.Fields("TestDiv").Value)
         
            strOrdDiv = Trim("" & tmpRs.Fields("OrdDiv").Value)
            strStsCd = Trim("" & tmpRs.Fields("StsCd").Value)
            strTestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)
            strTestcd = Trim("" & tmpRs.Fields("OrdCd").Value)
            
            .Col = enREVIEW1.tcSTSCD:     .Value = strStsCd     'Status
            
            .Col = enREVIEW1.tcSTSNM:
            
            'D/C여부
            If tmpRs.Fields("DcFg") = "1" Then .Value = .Value & "*"
           
            If P_ImageSystem = True Then
                Set RS = Nothing
                Set RS = New Recordset
                RS.Open objSQL.SqlGetImageData(pWorkArea, pAccDt, pAccSeq, strTestcd, ""), DBConn
                
                If RS.RecordCount > 0 Then
                    .Col = 34: .Value = 1
                Else
                    .Col = 34: .Value = 0
                End If
                Set RS = Nothing
            Else
                .Col = 34: .Value = 0
            End If
            '진료과 Remark(Message)
            
            '혈액처방 순번추가
            .Col = 35:  .Value = Trim("" & tmpRs.Fields("RSTSEQ").Value)
            
            aryMesg(RecordCnt) = "" & tmpRs.Fields("Mesg")
            tmpRs.MoveNext
        Loop
        
        Set tmpRs = Nothing
        For I = 1 To .MaxRows
            .Row = I
            .Col = enREVIEW1.tcWORKAREA:  pWorkArea = .Value    'WorkArea
            .Col = enREVIEW1.tcACCDT:     pAccDt = .Value       'AccDt
            .Col = enREVIEW1.tcACCSEQ:    pAccSeq = .Value      'AccSeq
            .Col = enREVIEW1.tcORDDIV:    strOrdDiv = .Value
            .Col = enREVIEW1.tcSTSCD:     strStsCd = .Value
            .Col = enREVIEW1.tcUNITQTY:   strUnit = .Value
            .Col = 32:                    strTestDiv = .Value
            
            .Col = enREVIEW1.tcSTSNM:
            If strOrdDiv = BBS_ORDDIV Then
                Select Case strStsCd
                    Case "0": .Value = "처방" & .Value
                    Case "1": .Value = "채취" & .Value
                    Case "2": .Value = "접수" & .Value
                    Case Else
                            If pWorkArea <> "" Then
                                .Value = BBS_STATUS(pWorkArea, pAccDt, pAccSeq, strUnit) & .Value
                            End If
                End Select
            Else
                Call GetOrderStatus(strOrdDiv, strStsCd, strTestDiv, _
                                    strStsNm, lngColor, pForeColor, iBtnFg, pWorkArea, pAccDt, pAccSeq, strUnit)
                .Value = strStsNm & .Value
                .ForeColor = pForeColor
                .Col = enREVIEW1.tcTAT   '검사소요시간버튼
                If iBtnFg = 1 Then
                    .CellType = CellTypeButton
                    .TypeButtonText = CS_QuestionMark   '"?"
                    .TypeButtonColor = DCM_LightGray     '회색
                Else
                    .CellType = CellTypeStaticText
                    .Text = ""
                End If
            
            End If
        Next
        
        If .MaxRows < lngMaxRows Then .MaxRows = lngMaxRows
        
        .RowHeight(-1) = lngRowHeight
        .Col = 1: .Row = 1: .Action = ActionActiveCell
      
       .ReDraw = True
    End With
    
On Error GoTo Err_Trap1
    
Errors:
    Set tmpRs = Nothing
    Set RS = Nothing
    
Err_Trap1:
    Set objStatus = Nothing
    
    ClearFg = False
    OrderFg = True
    OldRow = 0
   
    MouseDefault
    Me.Enabled = True
    QueryFg = False
'    tblOrdSheet.SetFocus
    
    If RecordCnt = 0 And txtPtId.Text <> "" Then
        MsgBox "이 환자는 입력하신 기간동안에 발생한 처방이 없습니다.", vbInformation, "결과조회"
        OrderFg = False
    End If
    
'    txtPtId.SetFocus

End Sub

Private Function BBS_STATUS(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal unitqty As String) As String
    Dim strTmp As String
    Dim lngA   As Long 'Assign
    Dim lngAC  As Long 'Assign 취소
    Dim lngD   As Long '출고
    Dim lngR   As Long '반환
    Dim lngE   As Long '폐기
    Dim lngT As Long
    Dim lngM As Long
    
' STS_NM_ORDER  '처방
' STS_NM_COLLECT  '채혈
' STS_NM_ACCESS  '접수
' STS_NM_INPROGRESS  '진행중
' STS_NM_REQUEST  '요청
' STS_NM_DONE  '완료
' STS_NM_END  '종결
    
    strTmp = objSQL.GetDeliveryCnt(WorkArea, AccDt, AccSeq)
    If strTmp <> "" Then
        lngA = medGetP(strTmp, 1, COL_DIV)
        lngAC = medGetP(strTmp, 2, COL_DIV)
        lngD = medGetP(strTmp, 3, COL_DIV)
        lngR = medGetP(strTmp, 4, COL_DIV)
        lngE = medGetP(strTmp, 5, COL_DIV)
        
        lngT = lngA - lngAC '실제 Assign된 량 모두 Assign되었으면 완료
        lngM = lngD - (lngR + lngE) '출고된 수량-(반환된 수량+폐기된 수량)'실제 출고량
        
        '출고는 하나도 안하고 어싸인만 했다가 모두 어싸인 취소하면 접수상태로 롤백...
        
        'vUnitqty : 처방수량
        '처방수량만큼 Assign이 되었으면 완료, 아니면 검사중
        If unitqty = lngT Then
            If lngD >= 1 Then '출고 액션이 한번이라도 된 경우
                If lngM >= 1 Then '실제 출고가 한건 이상인 경우
                    BBS_STATUS = "완료"
                Else
                    BBS_STATUS = "진행중"
                End If
            Else '출고가 하나도 안된 경우
                BBS_STATUS = "완료"
            End If
        Else
            BBS_STATUS = "진행중"
        End If
        
        If unitqty = lngM Then
            BBS_STATUS = "종결"
        End If
        
'        If unitqty = lngA - lngAC And unitqty = lngD - lngR Then
'            BBS_STATUS = "완결"
'        ElseIf lngA - lngAC = lngD - lngR Then
'            BBS_STATUS = "대기"
'        ElseIf unitqty >= lngA - lngAC And lngA - lngAC > lngD - lngR Then
'            BBS_STATUS = "준비"
'        Else
'            BBS_STATUS = "검사중"
'        End If
    Else
'        BBS_STATUS = "검사중"
    End If
    
End Function
Private Sub GetOrderStatus(ByVal pOrdDiv As String, ByVal pStsCd As String, _
                           ByVal pTestDiv As String, ByRef pStsNm As String, _
                           ByRef pStsColor As Long, ByRef ForeColor As Long, ByRef pBttnFg As Long, _
                           ByVal WorkArea As String, ByVal AccDt As String, _
                           ByVal AccSeq As String, ByVal unitqty As String)
' 08.11.07 양성현 결과조회 화면에서 상태값의 색깔 조정
    Select Case Trim(pStsCd)
        Case enStsCd.StsCd_LIS_Order:
             pStsNm = STS_LIS_Order:     pStsColor = DCM_Gray: pBttnFg = 1 '회색
            ForeColor = DCM_Green                                         ' 처방
        Case enStsCd.StsCd_LIS_Collection:
             pStsNm = STS_LIS_HaveSpc:   pStsColor = DCM_Gray: pBttnFg = 1 '회색
            ForeColor = DCM_Red                                    ' 채취
        Case enStsCd.StsCd_LIS_Accession:
                pBttnFg = 1
                pStsNm = STS_LIS_Access:    pStsColor = DCM_Gray: pBttnFg = 1 '회색
                ForeColor = DCM_Blue                                       ' 접수
        Case enStsCd.StsCd_LIS_InProcess:
             pStsNm = STS_LIS_Worksheet: pStsColor = DCM_Gray: pBttnFg = 1 '회색
            ForeColor = DCM_Title_Blue                                     ' 검사중
        Case enStsCd.StsCd_LIS_MidRst:
                pBttnFg = 1
                If pOrdDiv = APS_ORDDIV Then
                    pStsNm = STS_LIS_Reading:   pStsColor = DCM_Gray           '회색
                Else
                    If pTestDiv = TST_MicTest Then            '미생물검사
                        pStsNm = STS_LIS_MidRst:    pStsColor = DCM_Black      '검정색
                    Else: pStsNm = STS_LIS_Partial: pStsColor = DCM_Black      '검정색
                    End If
                End If
                ForeColor = DCM_Brown                                   ' 부분결과
        Case enStsCd.StsCd_LIS_FinRst:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_MidRst, _
                      IIf(pTestDiv = TST_MicTest, STS_LIS_FinRst, STS_LIS_Verify))
            ForeColor = DCM_Black                                      ' 결과
        Case enStsCd.StsCd_LIS_Modify:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_Verify, STS_LIS_Modify)
            ForeColor = DCM_MidGray                                    ' 수정
        Case "7":
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = STS_LIS_Modify
             ForeColor = DCM_Black                                     ' 최종
    End Select

End Sub

' 2009.01.09 양성현 환자ID 파라메터 추가

Public Sub accPTid(ByVal PTid As String)
'    gPatientId = PtId
    txtPtId.Text = PTid
    gUsingInWardMenu = True
    Call txtPtId_LostFocus
    gUsingInWardMenu = False

'    Call Form_Activate
    Call DisplayOrders

'    Call cmdRefresh_Click
    OldRow = 0
'    OrderFg = False

'    Call dtpToDate_KeyDown(vbKeyReturn, 0)
    Me.Show
    txtPtId.SetFocus

End Sub

Private Sub Form_Activate()
    If Trim(gPatientId) <> "" Then txtPtId.Text = gPatientId
    MsgFg = False
    Call chkPtList_Click
    
'On Error GoTo Err_Trap
'    If Trim(gPatientId) <> "" Then txtPtId.Text = gPatientId
'    txtPtId.SetFocus
'Err_Trap:
    If Trim(txtPtId.Text) <> "" Then SendKeys "{TAB}"
'    txtPtId.SetFocus
End Sub


Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub


Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo Err_Trap
    
    If Item.Text = "" Then Exit Sub
    txtPtId.SetFocus
    DoEvents
    txtPtId.Text = Item.Text
    Call txtPtId_KeyPress(vbKeyReturn)
    Exit Sub
    
Err_Trap:
    Resume Next

End Sub


Private Sub objPop_Click(ByVal vMenuID As Long)
    
    '** 결과지 출력 추가 By M.G.Choi 2008.01.29
    Select Case vMenuID
        Case MENU_PRT
           Dim MyData       As clsResults
           Dim MyReport     As clsResultReport
           Dim strLastRst   As String
        
           Dim I            As Integer
           
           Set MyData = New clsResults
           Set MyReport = New clsResultReport
           
           Screen.MousePointer = vbArrowHourglass
            
           With tblResult
                For I = 1 To .DataRowCnt
                    .Row = I
                    
                    MyData.ORDDT = medGetP(lblOrdDt.Caption, 1, " ")  '처방일
                    MyData.SpcNm = lblSpecimenNm.Caption   '검체명
                    MyData.ColDtTm = lblCollectDt.Caption              '채혈일
                    
                    .Col = 2: MyData.TestNm = .Value        '검사명
                    
                    MyData.VfyDt = lblVerifyDt.Caption        '보고일
                    
                    .Col = 3:   MyData.RstCd = .Value       '결과
                    .Col = 4:   MyData.RstUnit = .Value     '단위
                    .Col = 5: MyData.HLDiv = .Value       'High/Low
                    .Col = 6:   MyData.DPDiv = .Value       'Delta/Panic
                    .Col = 8:
                    If Trim(.Value) = "" Then
                        .Col = 12: MyData.RefRng = .Value     '참고치
                    Else
                        MyData.RefRng = .Value     '참고치
                    End If
                    
                    With tblOrdSheet
                        .Row = OldRow
                        .Col = enREVIEW1.tcWORKAREA: MyData.WorkArea = .Value
                        .Col = enREVIEW1.tcACCDT:    MyData.AccDt = .Value
                        .Col = enREVIEW1.tcACCSEQ:   MyData.AccSeq = .Value
                    End With
                    
                    .Col = 20:
                    
                    strLastRst = .Value         '최근결과
                    .Col = 21:
                               If Trim(strLastRst) <> "" Then
                                  MyData.LastRst = strLastRst & " (" & Mid(.Value, 4, 5) & ")"
                               Else
                                  MyData.LastRst = strLastRst
                               End If
                    .Col = 27: MyData.TestCd = .Value       '검사코드
                    .Col = 28: MyData.SpcCd = .Value        '검체코드
                    .Col = 30: MyData.TestDiv = .Value      'TestDiv
                    .Col = 32: MyData.OrdDate = .Value
                    .Col = 33: MyData.SpcName = .Value
                    .Col = 35: MyData.FootNoteFg = .Value   'footnotefg
                    .Col = 36: MyData.RmkCd = .Value        'Remark 코드
                    .Col = 37: MyData.SenFg = .Value        '감수성여부
                    .Col = 48: MyData.DetailFg = .Value
                    Call MyReport.Add(MyData)
Skip:
                Next
           End With
           If MyReport.Count <> 0 Then
                MyReport.PTid = txtPtId.Text
                MyReport.PtNm = lblPtNm.Caption
                MyReport.PtSex = lblSex.Caption
                MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
                MyReport.FromDt = Format(dtpFromDate.Value, CS_DateLongFormat)
                MyReport.TODT = Format(dtpToDate.Value, CS_DateLongFormat)
                MyReport.DeptCd = lblDeptNm.Caption
                If Trim(lblDeptNm.Caption) = "" Then MyReport.DeptCd = medGetP(lblLocation.Caption, 1, "-")
                MyReport.VfyDt = medGetP(lblVerifyDt.Caption, 1, " ")
                Call MyReport.Print_Report
            End If
            Screen.MousePointer = vbDefault
                    
            Set MyData = Nothing
            Set MyReport = Nothing
            
    End Select

    '** 원본 ----------------------------------------------------------------------------------------
'    Select Case vMenuID
'        Case MENU_PRT
'            Dim MyReport    As clsBatchReport
'            Dim RS          As Recordset
'
'            Dim pWorkArea   As String
'            Dim pAccDt      As String
'            Dim pAccSeq     As String
'            Dim lngseq      As Integer
'
'            Set MyReport = New clsBatchReport
'
'            With tblOrdSheet
'                .Row = OldRow
'                .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value
'                .Col = enREVIEW1.tcACCDT:    pAccDt = .Value
'                .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value
'            End With
'
'            Set RS = New Recordset
'            RS.Open objSQL.SqlMultiTest(pWorkArea, pAccDt, Val(pAccSeq)), DBConn
'
'            If Not RS.EOF Then
'                While Not RS.EOF
'                    lngseq = lngseq + 1
'                    pWorkArea = "" & RS.Fields("WorkArea").Value
'                    pAccDt = "" & RS.Fields("AccDt").Value
'                    pAccSeq = "" & RS.Fields("AccSeq").Value
'                    Set MyReport = Nothing
'                    Set MyReport = New clsBatchReport
'                    MyReport.PtId = txtPtId.Text
'                    MyReport.PtNm = lblPtNm.Caption
'                    MyReport.PtSex = lblSex.Caption
'                    MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
'                    MyReport.DeptNm = lblDeptNm.Caption
'                    MyReport.VfyDt = lblVerifyDt.Caption
'                    MyReport.VfyNM = lblVerifierNm.Caption
'                    MyReport.ICD = lblDisease.Caption
'                    Call MyReport.MicSensiReport(pWorkArea, pAccDt, pAccSeq, picESign)
''                    Call MyReport.ReportForOnePatient(txtPtId.Text,format(lblVerifyDt.Caption,"yyyymmdd"),format(lblVerifyDt.Caption,"yyyymmdd"),
'                    RS.MoveNext
'                Wend
'                Set RS = Nothing
'            Else
'                MyReport.PtId = txtPtId.Text
'                MyReport.PtNm = lblPtNm.Caption
'                MyReport.PtSex = lblSex.Caption
'                MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
'                MyReport.DeptNm = lblDeptNm.Caption
'                MyReport.VfyDt = lblVerifyDt.Caption
'                MyReport.VfyNM = lblVerifierNm.Caption
'                MyReport.ICD = lblDisease.Caption
'                Call MyReport.MicSensiReport(pWorkArea, pAccDt, pAccSeq, picESign)
'            End If
'            Set RS = Nothing
'            Set MyReport = Nothing
'    End Select
    '--------------------------------------------------------------------------------------------------
End Sub

Private Sub objTB_Click()
    Set objTB = Nothing
End Sub

Private Sub optOrdDiv_Click(Index As Integer)
    optOrdDiv(0).ForeColor = &H404040
    optOrdDiv(2).ForeColor = &H404040
    optOrdDiv(3).ForeColor = &H404040
    optOrdDiv(Index).ForeColor = DCM_LightRed
    ChkDivAll.Value = 0
    If optOrdDiv(2).Value = True Then optQueryKey(0).Value = True
    
End Sub

Private Sub optQueryKey_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then dtpFromDate.SetFocus
End Sub

Private Sub rtfResult_DblClick()
    
    Dim sLabNo()
    Dim strTag      As String
    Dim strLabNo    As String
    Dim aryLabNo    As Variant
    
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    
    strTag = rtfResult.Tag
    strLabNo = medGetP(strTag, 1, COL_DIV)
    aryLabNo = Split(strLabNo, "-")
    
    If aryLabNo(3) = BBS_ORDDIV Then Exit Sub
    
    frmAPS905.rtfResultText.Visible = True
    frmAPS905.OrdDiv = aryLabNo(3)
    
    If aryLabNo(3) = LIS_ORDDIV Then
        frmAPS905.Caption = medGetP(strTag, 2, COL_DIV)
        frmAPS905.rtfResultText.TextRTF = rtfResult.TextRTF
    End If
    
    Screen.MousePointer = vbDefault
    
    frmAPS905.WindowState = 0
    frmAPS905.Show vbModal
    DoEvents
    
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim pTestCd     As String
    Dim pSpcCd      As String
    Dim pErFg       As String
    Dim pTestNm     As String
    
    Dim pVfyDate    As String
    Dim pRcvDate    As String
    Dim pOrdDate    As String
    Dim pOrdDiv     As String
    
    Dim iNo         As Integer
    
    Dim objResult   As clsLISResultReview
    Dim strTATS     As String
   
    On Error GoTo Err_Trap:
   
    With tblOrdSheet
        .Row = Row
        .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value
        .Col = enREVIEW1.tcACCDT:    pAccDt = .Value
        .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value
        .Col = enREVIEW1.tcTESTCD:   pTestCd = .Value
        .Col = enREVIEW1.tcSPCCD:    pSpcCd = .Value
        .Col = enREVIEW1.tcSTATFG:   pErFg = .Value
        .Col = enREVIEW1.tcTESTNM:   pTestNm = .Value
        
        .Col = enREVIEW1.tcRCVDT:    pRcvDate = Format(.Value, "YYYY-MM-DD HH:MM")
        .Col = enREVIEW1.tcORDDATE:  pOrdDate = .Value
        .Col = enREVIEW1.tcVFYDATE:  pVfyDate = .Value
        .Col = enREVIEW1.tcORDDIV:   pOrdDiv = .Value
        If pOrdDiv <> LIS_ORDDIV Then Exit Sub
        
        '검사소요시간 읽어오기...
        Set objResult = New clsLISResultReview
        
        strTATS = objResult.GetTAT(pTestCd, pSpcCd, pErFg)
        If pAccSeq = "" Then
            iNo = 1
        Else
            iNo = objResult.GetBuildNoForTAT(pWorkArea, pAccDt, pAccSeq)
        End If
        .Col = enREVIEW1.tcTAT
        .CellType = CellTypeEdit
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .Text = medGetP(strTATS, iNo, ":")
    
    End With
    DoEvents
    
    With frmOnLineHelp
        .WorkArea = pWorkArea
        .AccDt = pAccDt
        .AccSeq = pAccSeq
        .TestCd = pTestCd
        
        .SpcCd = pSpcCd
        .TestNm = pTestNm
        .RcvDate = IIf(Mid(pRcvDate, 1, 1) = "2", pRcvDate, "")
        .VfyDate = IIf(Mid(pVfyDate, 1, 1) = "2", pVfyDate, "")
        .OrdDate = "20" & pOrdDate
        .TAT = medGetP(strTATS, iNo, ":")
        .Show , MainFrm
        
    End With
    Set objResult = Nothing
    Exit Sub
    
Err_Trap:
    Set objResult = Nothing
End Sub

'% 처방 선택(Click)하면 해당 결과 디스플레이...
Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)

    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim strOrdDiv   As String
    Dim strWardId   As String
    Dim strHosilId  As String
    Dim strOrdDt    As String
    Dim strOrdNo    As String
    Dim strSpcYY    As String
    Dim strSpcNo    As String
    Dim strTmp      As String
    Dim objResult   As clsLISResultReview
    Dim strRstSeq   As String
    
    With tblOrdSheet
      
        If Row = 0 Then Exit Sub
        If Row > .DataRowCnt Then Exit Sub
        
        '소요시간
        If Col = enREVIEW1.tcTAT Then
            Call tblOrdSheet_ButtonClicked(Col, Row, 1)
            Exit Sub
        End If
        If OldRow = Row Then Exit Sub
        
        Set objResult = New clsLISResultReview
        
        .Row = Row
        .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value: svWorkArea = .Value
        .Col = enREVIEW1.tcACCDT:    pAccDt = .Value:    svAccDt = .Value
        .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value:   svAccSeq = .Value
        .Col = enREVIEW1.tcORDDIV:   strOrdDiv = .Value
        .Col = enREVIEW1.tcWARDID:   strWardId = .Value
        .Col = enREVIEW1.tcHOSILID:  strHosilId = .Value
        .Col = enREVIEW1.tcTESTCD:   SvTestCd = .Value
        .Col = enREVIEW1.tcORDDATE:  strOrdDt = Format(.Value, CS_DateDbFormat)
        .Col = enREVIEW1.tcORDNUM:   strOrdNo = .Value
        .Col = enREVIEW1.tcSPCYY:   strSpcYY = .Value                    '검체년도
        .Col = enREVIEW1.tcSPCNO:    strSpcNo = .Value
        .Col = 35:  strRstSeq = .Value
        
        lblDeptNm.Caption = objResult.GetOrderDept(txtPtId.Text, strOrdDt, strOrdNo)
       
        Set objResult = Nothing
        
        '병동 (처방난 시점)
        If strWardId <> "" Then
            '--- 외부정도관리수정본(20161004)
            If Mid(txtPtId.Text, 1, 1) = "L" Then
                LisLabel4(4).Caption = "검체번호"
                lblLocation.Caption = strWardId
            Else
                LisLabel4(4).Caption = "병     실"
                lblLocation.Caption = strWardId & " - " & strHosilId
            End If
        Else
            lblLocation.Caption = ""
        End If
        
        If strOrdDiv = LIS_ORDDIV And (pWorkArea = "" Or pAccDt = "" Or pAccSeq = "") Then
            .Col = enREVIEW1.tcSTSCD
            If (.Value <> enStsCd.StsCd_LIS_Order) Then       '처방
                MsgBox "접수번호가 없습니다. (전산실 혹은 임상병리과로 연락바람 ☎" & ObjSysInfo.HelpLine & ")", vbExclamation, "오류발생"
            End If
            Exit Sub
        End If
      
        Call ResultClear
      
        If OldRow > 0 Then
            .Row = OldRow
            .Col = -1: .ForeColor = DCM_Black   '검정색
            
            .Col = enREVIEW1.tcSTSCD    '상태(처방,채혈,접수,검사중)
            If OldOrdDiv = LIS_ORDDIV And .Value = enStsCd.StsCd_LIS_Order Or .Value = enStsCd.StsCd_LIS_Collection Or _
               .Value = enStsCd.StsCd_LIS_Accession Or .Value = enStsCd.StsCd_LIS_InProcess Then
                .Col = enREVIEW1.tcSTSNM: .ForeColor = DCM_Gray            '회색
            End If
        End If
         
        .Row = Row
        .Col = -1: .ForeColor = DCM_Blue        '파랑색
        OldRow = Row
        OldOrdDiv = strOrdDiv
      
        MouseRunning  '13
      
        tblResult.ReDraw = False
        
        .Col = enREVIEW1.tcSPCNAME: lblSpecimenNm.Caption = .Value      '검체
        .Col = enREVIEW1.tcORDDATE: lblOrdDt.Caption = Format(.Value, "YYYY-MM-DD HH:MM")  '처방일
        .Col = enREVIEW1.tcORDDOCT: lblDoctNm.Caption = .Value          '처방의
        .Col = enREVIEW1.tcVFYNM:   lblVerifierNm.Caption = .Value      '보고자
        .Col = enREVIEW1.tcVFYDATE: lblVerifyDt.Caption = .Value        '보고일시
        .Col = enREVIEW1.tcRCVDT:   lblReceiveDt.Caption = Format(.Value, "YYYY-MM-DD HH:MM")
        

        .Col = enREVIEW1.tcSTSCD
        If .Value >= enStsCd.StsCd_LIS_FinRst Then
' 2009.04.13 양성현 상태가 결과이상의 상태인 경우로 전문의 처리부문 수정
                '** 추가 확인자 Display By M.G.Choi 2009.09.02
                lblLisDoctNm.Caption = GetLisDoctNm(SvTestCd)
        Else
                lblLisDoctNm.Caption = ""
        End If
        lblReceiverNm.Caption = ""
        
        Select Case strOrdDiv
            Case BBS_ORDDIV:
                rtfResult.Text = ""
                rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
                If strOrdDiv = BBS_ORDDIV Then
                    If P_IncludeBBSSystem Then
                        Screen.MousePointer = vbArrowHourglass
                        rtfResult.Visible = True
                        rtfResult.ZOrder 0
                        DoEvents
                        cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
                        DoEvents
                        Call DisplayBBSResult(pWorkArea, pAccDt, Val(pAccSeq), Row, strRstSeq)
                        Screen.MousePointer = vbDefault
                    End If
                End If
            Case LIS_ORDDIV:
                Screen.MousePointer = vbArrowHourglass
                rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
                rtfResult.Visible = False
                DoEvents
                tblResult.ReDraw = False
                cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
                DoEvents
                Call DisplayLISResult(pWorkArea, pAccDt, Val(pAccSeq))
                tblResult.ReDraw = True
                Screen.MousePointer = vbDefault
            Case POC_ORDDIV:
                .Col = enREVIEW1.tcORDDATE:     pAccDt = Format(.Value, CS_DateDbFormat)
                Call DisplayPOCResult(txtPtId.Text, pAccDt)
            Case CMT_ORDDIV:
                Call DisplayLABComment(Row)
        End Select
        
        tblResult.TopRow = 1
        ResultFg = True
      
        tblResult.ReDraw = True
        tblResult.Refresh
   
        chkPtList.Value = 0
        chkSize.Value = 0
        tblResult.Row = -1
        tblResult.Col = -1
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.BlockMode = False
        tblResult.ColWidth(3) = 30
        
        chkSize.Value = 1
        
        Call MouseDefault
   
    End With
End Sub

'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayBBSResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As Long, ByVal iRow As Long, ByVal pRstSeq As String)
    Dim ObjABO         As clsABO
    Dim objTransReason As clsQueryOrder
    Dim objRmk         As clsCrossMatching
    Dim objResult      As clsLISResultReview
    Dim RS             As Recordset
    Dim objTmp()       As String
    Dim strTransResult As String
    Dim strUnitQty     As String
    Dim strReqDtTm     As String
    Dim strReason      As String
    Dim strOrdDt       As String
    Dim strOrdNo       As String
    Dim strRmk         As String
    Dim strTmp         As String
    Dim strTmpBlood    As String
    Dim strJudge       As String
    Dim TF             As Boolean
    
    Dim lngAssignCnt   As Long
    Dim lngDeliveryCnt As Long
    Dim lngReturnCnt As Long
    Dim lngExpCnt As Long
    Dim strDelivelyDt As String
    Dim strDelivelyTm As String
    Dim ii             As Integer
    
    Set ObjABO = New clsABO
    Set objTransReason = New clsQueryOrder
    Set objRmk = New clsCrossMatching
    Set objResult = New clsLISResultReview
    
    With tblOrdSheet
        .Row = iRow
        .Col = enREVIEW1.tcUNITQTY: strUnitQty = .Value
        .Col = enREVIEW1.tcREQDATE: strReqDtTm = Format(.Value, CS_DateMask)
        .Col = enREVIEW1.tcREQTIME: strReqDtTm = strReqDtTm & " " & Format(Mid(.Value, 1, 4), CS_TimeShortMask)
        .Col = enREVIEW1.tcORDDATE: strOrdDt = Format(.Value, CS_DateDbFormat)
        .Col = enREVIEW1.tcORDNUM:  strOrdNo = .Value
    End With
    
    strTmp = objResult.GetOrderColid(txtPtId.Text)
    If strTmp <> "" Then
        lblCollectorNm.Caption = medGetP(strTmp, 1, COL_DIV)
        lblCollectDt.Caption = medGetP(strTmp, 2, COL_DIV)
        lblReceiverNm.Caption = medGetP(strTmp, 3, COL_DIV)
    End If
    
    strTmp = ""
    strReason = objTransReason.GetTransReason(txtPtId.Text, strOrdDt, strOrdNo)
    
    strRmk = ""
    strRmk = objRmk.GetptidRmk(txtPtId.Text)
    
    If strRmk <> "" Then
        objTmp() = Split(strRmk, vbCrLf)
    End If
    
'    strTmp = objSQL.GetDeliveryCnt(pWorkArea, pAccDt, CStr(pAccSeq))
    ' -- 2012.07.10 출고일자 추가
    ' -- by 온승호
    strTmp = objSQL.GetDeliveryCnt_New(pWorkArea, pAccDt, CStr(pAccSeq), pRstSeq)
    
    If strTmp <> "" Then
        lngAssignCnt = Val(medGetP(strTmp, 1, COL_DIV)) - Val(medGetP(strTmp, 2, COL_DIV))
        lngDeliveryCnt = Val(medGetP(strTmp, 3, COL_DIV))
        lngReturnCnt = Val(medGetP(strTmp, 4, COL_DIV))
        lngExpCnt = Val(medGetP(strTmp, 5, COL_DIV))
        strDelivelyDt = Trim(medGetP(strTmp, 6, COL_DIV))
        strDelivelyTm = Trim(medGetP(strTmp, 7, COL_DIV))
    End If

    ObjABO.PTid = txtPtId.Text  '혈액형을 구하자.
    ObjABO.GetABO
    
    Set RS = objTransReason.DonorInformation(txtPtId.Text)
    
    With rtfResult
        .Visible = False
        .Text = vbCrLf & Space(13) & "◈ 수혈 진행상황 ◈" & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 혈 액 형  :  " & ObjABO.ABO & ObjABO.Rh & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 예 정 일  :  " & strReqDtTm & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 수혈사유  :  " & strReason & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 수    량  :  " & strUnitQty & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 결과등록  :  " & lngAssignCnt & vbCrLf & vbCrLf
        .Text = .Text & Space(3) & "▶ 출고수량  :  " & lngDeliveryCnt & vbCrLf & vbCrLf
        ' -- 2012.07.10 출고일자 추가
        ' -- by 온승호
        .Text = .Text & Space(3) & "▶ 출고일자  :  " & Format(strDelivelyDt, "####-##-##") & " " & Format(strDelivelyTm, "##:##:##") & vbCrLf & vbCrLf
        
        If lngReturnCnt <> 0 Then
            .Text = .Text & Space(3) & "▶ 반환수량  :  " & lngReturnCnt & vbCrLf & vbCrLf
        End If
        If lngExpCnt <> 0 Then
            .Text = .Text & Space(3) & "▶ 폐기수량  :  " & lngExpCnt & vbCrLf & vbCrLf
        End If
        If strRmk <> "" Then
            .Text = .Text & Space(3) & "▶ 환자기록사항 " & vbCrLf & vbCrLf
            For ii = LBound(objTmp) To UBound(objTmp)
                .Text = .Text & Space(6) & objTmp(ii) & vbCrLf
            Next
        End If
        If Not RS.EOF Then
            Do Until RS.EOF
                Select Case RS.Fields("okdiv3").Value & ""
                    Case "1":  strJudge = "적격"
                    Case "0":  strJudge = "부적격"
                    Case Else: strJudge = "미등록"
                End Select
                
                strTmpBlood = RS.Fields("donornm").Value & "" & "(" & RS.Fields("tmpid").Value & "" & "," & strJudge & ")" & vbCrLf
                If TF = False Then
                    .Text = .Text & Space(3) & "▶ 헌 혈 자 : " & strTmpBlood & vbCrLf '& vbCrLf
                Else
                    .Text = .Text & Space(3) & Space(13) & strTmpBlood & vbCrLf
                End If
                TF = True
                RS.MoveNext
            Loop
        End If
        
        .SelStart = 15: .SelLength = Len(.Text)
        .SelProtected = False
        .SelFontName = "굴림체"
        .SelFontSize = 13
        .SelBold = True
        
        .SelStart = 30: .SelLength = Len(.Text)
        .SelProtected = False
        .SelFontName = "돋움체"
        .SelFontSize = 10
        .SelBold = True
        
        Call HighlightText(rtfResult, "◈ 수혈 진행상황 ◈", True, , &H4A4189)
        Call HighlightText(rtfResult, "▶ 혈 액 형 :", False, , &H553755)
        Call HighlightText(rtfResult, ObjABO.ABO & ObjABO.Rh, False, , &H7477EF, 15)  '약간 붉은색
        Call HighlightText(rtfResult, "▶ 예 정 일 :", False, , &H553755)
        Call HighlightText(rtfResult, strReqDtTm, False, , &HE48372)
        Call HighlightText(rtfResult, "▶ 수혈사유 :", False, , &H553755)
        Call HighlightText(rtfResult, "▶ 수    량 :", False, , &H553755)
        Call HighlightText(rtfResult, "▶ 결과등록 :", False, , &H553755)
        Call HighlightText(rtfResult, "▶ 출고수량 :", False, , &H553755)
        If lngReturnCnt <> 0 Then
            Call HighlightText(rtfResult, "▶ 반환수량 :", False, , &H553755)
        End If
        If lngExpCnt <> 0 Then
            Call HighlightText(rtfResult, "▶ 폐기수량 :", False, , &H553755)
        End If
        Call HighlightText(rtfResult, "▶ 환자기록사항 ", False, , &H553755)
        
        If TF = True Then
            Call HighlightText(rtfResult, "▶ 헌 혈 자 :", False, , &H553755)
        End If
        .Visible = True
    
    End With
    
    Set RS = Nothing
    Set objRmk = Nothing
    Set ObjABO = Nothing
    Set objResult = Nothing
    Set objTransReason = Nothing
    
End Sub


Private Sub DisplayPOCResult(ByVal pPtID As String, ByVal pVfyDt As String)
    Dim objResult       As New clsLISResultReview
    Dim ResultBuffer    As String
    Dim I               As Integer
    Dim J               As Integer
    
    Set objResult = New clsLISResultReview
    
    With objResult
        ResultBuffer = .POCResultQuery(pPtID, pVfyDt)
        For I = 1 To .RstRow
            tblResult.Row = I   '+ .OffSet
            For J = 1 To 8
                tblResult.Col = J
                tblResult.ForeColor = .Get_ForeColor(J, I)
            Next
        Next
    End With
    
    '결과내역 Display
    tblResult.Row = 1
    tblResult.Row2 = tblResult.MaxRows
    tblResult.Col = 2
    tblResult.Col2 = tblResult.MaxCols
    tblResult.BlockMode = True
    tblResult.AllowCellOverflow = True
    tblResult.Clip = ResultBuffer
    tblResult.BlockMode = False
    
End Sub


'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayLISResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
    Dim objResult       As clsLISResultReview
    Dim ResultBuffer    As String
    Dim RstTxtBuffer    As String
    Dim SamTxtBuffer    As String
    Dim strTestDiv      As String
    Dim strTestcd       As String
    Dim strOCSTmp       As String
    
    Dim blnImage        As Boolean
    
    Dim I               As Integer
    Dim J               As Integer
    
    
    Set objResult = New clsLISResultReview
    With objResult
        Call .ResultQuery(pWorkArea, pAccDt, pAccSeq)
        lblWorkArea.Caption = .GetWorkAreaNm(pWorkArea)     'WorkAra Name
        lblCollectorNm.Caption = GetDoctNm(.ColId)         '채혈자(간호사)
        If Trim(lblCollectorNm.Caption) = "" Then
            lblCollectorNm.Caption = GetEmpNm(.ColId)       '채혈자(병리사)
        End If
        lblReceiverNm.Caption = GetEmpNm(.RcvId)            '접수자
        lblCollectDt.Caption = .ColDtTm                     '채혈일시
        lblDeptNm.Caption = .DeptNm
        lblBedinDt.Caption = .BedinDt
        
        '** 추가루틴 By M.G.Choi 2005.08.11 ==========================================
        ' * 예수병원 현재환자위치 조회 함
        strOCSTmp = INPt_PArea(Trim(txtPtId.Text), .BedinDt)
        If strOCSTmp <> "" Then
            lblLocation.Caption = lblLocation.Caption & "(" & strOCSTmp & ")"
        End If
        '=============================================================================
            
        '상태가 처방/채혈/접수/검사중이면 Exit
        tblOrdSheet.Col = enREVIEW1.tcSTSCD
        If tblOrdSheet.Value = enStsCd.StsCd_LIS_Order Or tblOrdSheet.Value = enStsCd.StsCd_LIS_Collection Or _
            tblOrdSheet.Value = enStsCd.StsCd_LIS_Accession Or _
            tblOrdSheet.Value = enStsCd.StsCd_LIS_InProcess Then GoTo NoData
        
        If .ResultCnt = 0 Then GoTo NoData
        
        tblOrdSheet.Row = tblOrdSheet.ActiveRow
        tblOrdSheet.Col = enREVIEW1.tcTESTCD: strTestcd = tblOrdSheet.Value
        tblOrdSheet.Col = 34
    
        
        If tblOrdSheet.Value = 1 Then
            For I = 1 To .RstRow
                frmAPS905.tblResult.Row = I   '+ .OffSet
                For J = 1 To 8
                    frmAPS905.tblResult.Col = J
                    frmAPS905.tblResult.ForeColor = .Get_ForeColor(J, I)
                Next
            Next
            
            '결과내역 Display
            frmAPS905.tblResult.Row = 1
            frmAPS905.tblResult.Row2 = frmAPS905.tblResult.MaxRows
            frmAPS905.tblResult.Col = 2
            frmAPS905.tblResult.Col2 = frmAPS905.tblResult.MaxCols
            frmAPS905.tblResult.BlockMode = True
            frmAPS905.tblResult.AllowCellOverflow = True
            frmAPS905.tblResult.Clip = .ResultClipText '& .SenClipText 'ResultBuffer
            frmAPS905.tblResult.BlockMode = False
            
            '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
            If .SortFg Then
                For I = 1 To .SensiCount
                    frmAPS905.tblResult.SortBy = SortByRow
                    frmAPS905.tblResult.SortKey(1) = 2  '항생제명
                    frmAPS905.tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                    frmAPS905.tblResult.Col = -1
                    frmAPS905.tblResult.Row = .AntiSortStartRow(I)   '+ .OffSet
                    frmAPS905.tblResult.Row2 = .AntiSortEndRow(I)    '+ .OffSet
                    frmAPS905.tblResult.Action = ActionSort
                    frmAPS905.tblResult.Row = .SortStartRow - 1 '+ .OffSet
                    frmAPS905.tblResult.Col = 2
                    frmAPS905.tblResult.FontUnderline = True
                Next
            Else
                frmAPS905.tblResult.Col = 6
                frmAPS905.tblResult.Row = -1
                frmAPS905.tblResult.ForeColor = DCM_LightRed
                frmAPS905.tblResult.FontBold = True
            End If
            If Val(.TestDiv) = TST_MicTest Then
                '미생물 결과 : 균명컬럼 Align Left
                frmAPS905.tblResult.Row = -1
                frmAPS905.tblResult.Col = -1
                frmAPS905.tblResult.BlockMode = True
                frmAPS905.tblResult.AllowCellOverflow = True
                frmAPS905.tblResult.TypeHAlign = TypeHAlignLeft
                frmAPS905.tblResult.BlockMode = False
            
                frmAPS905.tblResult.Col = 3: tblResult.Col2 = 7
                frmAPS905.tblResult.Row = -1
                frmAPS905.tblResult.BlockMode = True
                frmAPS905.tblResult.FontBold = False
                frmAPS905.tblResult.BlockMode = False
            Else
                '일반결과 : 결과컬럼 Align Center
                frmAPS905.tblResult.Row = 1: frmAPS905.tblResult.Row2 = frmAPS905.tblResult.MaxRows
                frmAPS905.tblResult.Col = 3: frmAPS905.tblResult.Col2 = 7
                frmAPS905.tblResult.BlockMode = True
                frmAPS905.tblResult.TypeHAlign = TypeHAlignCenter
                frmAPS905.tblResult.BlockMode = False
            End If
            '텍스트결과 Display
            If .TextFg Then
                frmAPS905.txtRstCmt.TextRTF = .RstTextBuffer      'RstTxtBuffer
                Call HighlightText(frmAPS905.txtRstCmt, "<< 검사 소견 >>", True, , &H4A4189)
                Call HighlightText(frmAPS905.txtRstCmt, "<< Supplemental Report >>", False, , &H4A4189)
            End If
            '검체리마크 & 풋노트 Display
            If .CommentFg Then
                frmAPS905.txtSamCmt.Text = .SamTextBuffer
                Call HighlightText(frmAPS905.txtSamCmt, "<< Remark >>", True)
                Call HighlightText(frmAPS905.txtSamCmt, "<< Foot Note >>", False)
            End If
            '특수검사 결과 Display
            frmAPS905.OrdDiv = VBA.Strings.Right(rtfResult.Tag, 1)
            If .SpecialFg Then
                frmAPS905.Special = True
                Call frmAPS905.DisplayForm(strTestcd, pWorkArea, pAccDt, pAccSeq)
                rtfResult.TextRTF = .SpeTextBuffer
                rtfResult.Tag = rtfResult.Tag & COL_DIV & .SpeRstTitle
                tblOrdSheet.Row = OldRow
                tblOrdSheet.Col = 32: strTestDiv = tblOrdSheet.Value
                If strTestDiv = CStr(enTestDiv.TST_SpeTest) Then Call rtfResult_DblClick
            Else
                frmAPS905.Special = False
                Call frmAPS905.DisplayForm(strTestcd, pWorkArea, pAccDt, pAccSeq)
                frmAPS905.WindowState = 0
                frmAPS905.Show vbModal
                DoEvents
            End If
            blnImage = True
        Else
            blnImage = False
        End If
        ' 일반검사 - High / Low 컬럼 ForeColor 설정
        For I = 1 To .RstRow
            tblResult.Row = I   '+ .OffSet
            For J = 1 To 8
                tblResult.Col = J
                tblResult.ForeColor = .Get_ForeColor(J, I)
'                Debug.Print I & J & " : " & .Get_ForeColor(J, I)
            Next
        Next
        
        '결과내역 Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.Col2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText    '& .SenClipText 'ResultBuffer
        tblResult.BlockMode = False
        
        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
        If .SortFg Then
            For I = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  '항생제명
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(I)   '+ .OffSet
                tblResult.Row2 = .AntiSortEndRow(I)    '+ .OffSet
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 '+ .OffSet
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If Val(.TestDiv) = TST_MicTest Then
            '미생물 결과 : 균명컬럼 Align Left
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            For I = 1 To 5
                If .MicFg(I) Then
                    tblResult.ColWidth(I + 2) = 9
                Else
                    tblResult.ColWidth(I + 2) = 4
                End If
            Next
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            '일반결과 : 결과컬럼 Align Center
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
        '텍스트결과 Display
        If .TextFg Then
            txtRstCmt.TextRTF = .RstTextBuffer      'RstTxtBuffer
            chkRstCmt.Value = 1
            chkRstCmt.Enabled = True
            Call HighlightText(txtRstCmt, "<< 검사 소견 >>", True, , &H4A4189)
            Call HighlightText(txtRstCmt, "<< Supplemental Report >>", False, , &H4A4189)
        Else
            chkRstCmt.Value = 0
            chkRstCmt.Enabled = False
        End If
        '특수검사 결과 Display
        If .SpecialFg And blnImage = False Then
            rtfResult.TextRTF = .SpeTextBuffer
            rtfResult.Tag = rtfResult.Tag & COL_DIV & .SpeRstTitle
            tblOrdSheet.Row = OldRow
            tblOrdSheet.Col = 32: strTestDiv = tblOrdSheet.Value
            If strTestDiv = CStr(enTestDiv.TST_SpeTest) Then Call rtfResult_DblClick
        End If
        '검체리마크 & 풋노트 Display
        If .CommentFg Then
            txtSamCmt.Text = .SamTextBuffer
            chkSamCmt.Value = 1
            chkSamCmt.Enabled = True
            Call HighlightText(txtSamCmt, "<< Remark >>", True)
            Call HighlightText(txtSamCmt, "<< Foot Note >>", False)
        Else
            chkSamCmt.Value = 0
            chkSamCmt.Enabled = False
        End If
        
        '** 추가 수정 History By M.G.Choi 2006.09.06
'        Dim strModifyHistory    As String
'
'        strModifyHistory = GetModifyHistory(pWorkArea, pAccDt, pAccSeq)
'        If strModifyHistory <> "" Then
'            txtSamCmt.Text = txtSamCmt.Text & "<< Supplemental Report >>" & vbNewLine & strModifyHistory
'            chkSamCmt.Value = 1
'            chkSamCmt.Enabled = True
''            Call HighlightText(txtSamCmt, "<< Supplemental Report >>", False)
'        Else
'            chkSamCmt.Value = 0
'            chkSamCmt.Enabled = False
'        End If
        
        If .TBFg = True Then Call DisplayTBReport(pWorkArea, pAccDt, pAccSeq)
    End With
NoData:
    Set objResult = Nothing
    
'    txtPtId.SetFocus

End Sub

Private Function GetModifyHistory(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim strTmp      As String
    
    On Error Resume Next
    
    strSQL = " select rsttxt from " & T_LAB304 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq) & _
             "  order by seq "
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.BOF = False Then
        Do Until RS.EOF = True
            If strTmp = "" Then
                strTmp = RS.Fields("rsttxt").Value & ""
            Else
                strTmp = strTmp & vbNewLine & "---------------------" & vbNewLine & RS.Fields("rsttxt").Value & ""
            End If
            
            RS.MoveNext
        Loop
    End If
    
    GetModifyHistory = strTmp
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub DisplayTBReport(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
    Set objTB = New frmTBReport
    With objTB
        .Top = Me.Top
        .Left = Me.Left + 7000
        .tblRst.Visible = True
        .MousePointer = vbDefault
        Call objTB.GetTBResult(pWorkArea, pAccDt, pAccSeq)
        .WindowState = 0
        .Show vbModal
        DoEvents
    End With
    Set objTB = Nothing
End Sub

'% 폼 로드
Private Sub Form_Load()
    txtSearchKey.Text = ""
    gPatientId = ""
    chkPtList.Value = 0
    picPtList.Visible = False:  OrderFg = False
    ResultFg = False:           ClearFg = True
    PtFg = False:               optSort(1).Value = True
    
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame2.Visible = False
    
    OldRow = 0
    medInitLvwHead lvwPtList, "환자ID,환자성명,주민등록번호,생년월일,성별/나이,주소,전화번호", _
                       "300,300,1000,500,400,1800,800"
                       
    If P_ReviewStartDate <> "" Then
        dtpFromDate.Value = Format(P_ReviewStartDate, CS_DateLongMask)
    Else
        dtpFromDate.Value = DateAdd("yyyy", -1, GetSystemDate)
    End If
    
    optQueryKey(1).Value = True     '처방일
    
    dtpToDate.Value = GetSystemDate
    Call ClearRtn
    ChkDivAll.Value = 0
    optOrdDiv(0).Value = True
    
    If gUsingInWardMenu Then
'        ChkDivAll.Value = 1
        Call ChkDivAll_Click
    Else
'        ChkDivAll.Value = 0
        Call ChkDivAll_Click
' 2009.08.20 양성현 강성수선생 요청에 의해 막음.
'        Select Case ObjSysInfo.ProjectId
'            Case "LIS": optOrdDiv(0).Value = True
'            Case "BBS": optOrdDiv(2).Value = True
'        End Select
    End If

    Set objPatient = New clsPatient
    Set objSQL = New clsLISSqlReview
'    Me.Show
End Sub


'% 정렬 기준 선택
Private Sub optSort_Click(Index As Integer)
    If Not picPtList.Visible Then Exit Sub
    If txtSearchKey.Text <> "" Then
        Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    txtSearchKey.SetFocus
End Sub

'% 처방테이블 Set Focus
Private Sub tblOrdSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Trap
    If OrderFg Then tblOrdSheet.SetFocus
Err_Trap:
End Sub
Private Sub tblOrdSheet_LostFocus()
'    txtPtId.SetFocus
End Sub

'처방내역 테이블에 ToolTip 보여주기...
Private Sub tblOrdSheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim RS          As Recordset
    Dim lngseq      As String
    Dim tmpToolTip  As String
    Dim tmpPanelFg  As String
    Dim strSQL      As String
    
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strReqdt    As String
    Dim strMdRsn    As String
    Dim strTestcd   As String
    
    If Not OrderFg Then Exit Sub
   
    If Row <= 0 Then Exit Sub
    tmpToolTip = vbCrLf
   
    With tblOrdSheet
        .Row = Row
        
        .Col = 3: If Trim(.Value) = "" Then Exit Sub
        
        .Col = enREVIEW1.tcREQDATE:   '.Value = Trim("" & tmpRs.GetValue("ReqDt"))         '수혈예정일
                    strReqdt = Format(.Value, "####-##-##")
        .Col = enREVIEW1.tcREQTIME:  ' .Value = Trim("" & tmpRs.GetValue("ReqTm"))         '수혈예정시간
                    strReqdt = strReqdt & "  " & Format(.Value, "0#:##:##")
        .Col = 9:   tmpToolTip = tmpToolTip & "  처방일시 : " & .Value & vbCrLf  '처방일시
        .Col = 13:  tmpToolTip = tmpToolTip & "  처방번호 : " & .Value & vbCrLf  '처방번호
        .Col = 3:   tmpToolTip = tmpToolTip & "  검 사 명 : " & .Value & vbCrLf  '검사명
        .Col = 20:   tmpToolTip = tmpToolTip & "  검사코드 : " & .Value & vbCrLf  '검사코드
        .Col = 4:   tmpToolTip = tmpToolTip & "  검    체 : " & .Value & vbCrLf  '검체
        .Col = 11:  tmpToolTip = tmpToolTip & "  처 방 의 : " & .Value & vbCrLf  '처방의
        .Col = 14:  strWorkArea = .Value
        .Col = 15:  strAccDt = .Value
        .Col = 16:  strAccSeq = .Value
        .Col = 31:  tmpPanelFg = .Value
        .Col = 33:  strMdRsn = .Value
        
        
        '검체번호 추가 By M.G.Choi 2007.02.06
        '----------------------------------------------------------------------------------
        Set RS = New Recordset
        
        strSQL = " select spcyy, spcno from " & T_LAB201 & _
                 "  where workarea = " & DBS(strWorkArea) & _
                 "    and accdt = " & DBS(strAccDt) & _
                 "    and accseq = " & DBN(strAccSeq)
        
        RS.Open strSQL, DBConn
        
        If RS.EOF = False Then
            tmpToolTip = tmpToolTip & "  검체번호 : " & RS.Fields("spcyy").Value & "" & "-" & RS.Fields("spcno").Value & "" & vbCrLf  '검체번호
        End If
        
        RS.Close
        Set RS = Nothing
        '----------------------------------------------------------------------------------
        
        If tmpPanelFg = PN_Group Then
            lngseq = 0
            strSQL = objSQL.SqlMultiTest(strWorkArea, strAccDt, Val(strAccSeq))
            Set RS = New Recordset
            RS.Open strSQL, DBConn
            
            If Not RS.EOF Then
                tmpToolTip = tmpToolTip & "  접수번호 : " & vbCrLf
                While Not RS.EOF
                    lngseq = lngseq + 1
                    tmpToolTip = tmpToolTip & "      복수검체 " & CStr(lngseq) & " : " & RS.Fields("WorkArea").Value & "-"
                    tmpToolTip = tmpToolTip & Mid("" & RS.Fields("AccDt").Value, 3) & "-"
                    tmpToolTip = tmpToolTip & RS.Fields("AccSeq").Value & vbCrLf
                    RS.MoveNext
                Wend
                Set RS = Nothing
            Else
                tmpToolTip = tmpToolTip & "  접수번호 : " & strWorkArea & "-"
                tmpToolTip = tmpToolTip & Mid(strAccDt, 3) & "-"
                tmpToolTip = tmpToolTip & strAccSeq & vbCrLf
            End If
        Else
            tmpToolTip = tmpToolTip & "  접수번호 : " & strWorkArea & "-"
            tmpToolTip = tmpToolTip & Mid(strAccDt, 3) & "-"
            tmpToolTip = tmpToolTip & strAccSeq & vbCrLf
        End If
        If UBound(aryMesg) >= Row Then
            If aryMesg(Row) <> "" Then tmpToolTip = tmpToolTip & vbCrLf & "  " & aryMesg(Row) & vbCrLf
        End If
        
'08.09.25 양성현 수납일시 추가

        tmpToolTip = tmpToolTip & "  수 납  일 시 :" & strReqdt & vbCrLf


        tmpToolTip = tmpToolTip & "  희망채취일시 :" & strReqdt & vbCrLf
        
        If Trim(strMdRsn) <> "" Then tmpToolTip = tmpToolTip & "  결과수정사유 :" & strMdRsn & vbCrLf
        
        strSQL = objSQL.SqlGetWsUnit(strWorkArea, strAccDt, Val(strAccSeq))
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open strSQL, DBConn
        
        If Not RS.EOF Then
            tmpToolTip = tmpToolTip & "  Worksheet Unit : " & vbCrLf
            While Not RS.EOF
                tmpToolTip = tmpToolTip & _
                             "     " & RS.Fields("WSCD").Value & "-" & _
                             RS.Fields("WSUNIT").Value & vbCrLf
                RS.MoveNext
            Wend
            Set RS = Nothing
        End If
        
        '검사항목별 Comment 조회
       .Col = enREVIEW1.tcTESTCD: strTestcd = .Value
        
        tmpToolTip = tmpToolTip & vbCrLf & TestItemToolTipString(strTestcd)
        
        '검사항목별 장비명 조회
        Set RS = New Recordset
        
        strSQL = " select eqpcd from " & T_LAB302 & _
                 "  where workarea = " & DBS(strWorkArea) & _
                 "    and accdt = " & DBS(strAccDt) & _
                 "    and accseq = " & DBN(strAccSeq) & _
                 "    and testcd = " & DBS(strTestcd)
        
        RS.Open strSQL, DBConn
        
        If RS.EOF = False Then
            tmpToolTip = tmpToolTip & vbCrLf & "  장비코드 : " & RS.Fields("eqpcd").Value & ""
        End If
        
        RS.Close
        Set RS = Nothing
        
        '결과수정일시/수정자 조회
        Set RS = New Recordset
        
        strSQL = " select mfyid,mfydt,mfytm from " & T_LAB308 & _
                 "  where workarea = " & DBS(strWorkArea) & _
                 "    and accdt = " & DBS(strAccDt) & _
                 "    and accseq = " & DBN(strAccSeq)
        
        RS.Open strSQL, DBConn
        
        If RS.EOF = False Then
            tmpToolTip = tmpToolTip & vbCrLf & "  수정자 : " & RS.Fields("mfyid").Value & ""
            tmpToolTip = tmpToolTip & vbCrLf & "  수정일자 : " & Format(RS.Fields("mfydt").Value & "", "####-##-##") & " " & Format(RS.Fields("mfytm").Value & "", "0#:##:##")
        End If
        
        RS.Close
        Set RS = Nothing
        
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 7000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
   Set RS = Nothing
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    tblResult.Row = Row
    tblResult.Col = Col
    If tblResult.Value = "☞ RESULT" Then
        If Trim(rtfResult.Text) <> "" Then
            Call rtfResult_DblClick
        End If
    End If
End Sub

'% 결과테이블 Set Focus
Private Sub tblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Trap
    If ResultFg Then tblResult.SetFocus
Err_Trap:
End Sub

Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
'    If Not optOrdDiv(3).Value Then Exit Sub
    If tblResult.DataRowCnt = 0 Then Exit Sub
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_PRT, "결과지 출력"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuPrint = frmControls.mnuSub
'    mnuPrint.Caption = "결과지 출력"
'
'    Me.PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuPrint = Nothing
'    Unload frmControls
'    Set frmControls = Nothing
    
End Sub

'결과내역 테이블에 ToolTip 보여주기...
Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim RS       As Recordset
    Dim tmpToolTip  As String
    Dim strSQL      As String
    Dim strModRst   As String
    Dim strRstVal   As String
    
    If Not ResultFg Then Exit Sub
    
    tmpToolTip = vbCrLf
   
    With tblResult
        .Row = Row
        .Col = 2:
                 If .Value = "" Then
                    ShowTip = False
                    GoTo Skip
                 End If
        .Col = 8:  tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf   '처방명(Long)
        .Col = 9:
                If .Value <> "" Then
                    tmpToolTip = tmpToolTip & vbCrLf & "  최근결과 : " & .Value & vbCrLf   '최근결과
                    .Col = 10
                    tmpToolTip = tmpToolTip & "  보고일시 : " & .Value & vbCrLf  '최근결과일
                End If
        .Col = 13:
                '## 5.0.1: 이상대(2004-12-31)
                '   - 검사항목별 결과코드에 등록된 코드는 결과명 표시
                strSQL = objSQL.SqlGetOldResult(svWorkArea, svAccDt, svAccSeq, .Value)
                Set RS = New Recordset
                RS.Open strSQL, DBConn
                
                If Not RS.EOF Then
                   tmpToolTip = tmpToolTip & vbCrLf & "  [ 수정전 결과 ]  " '& vbCrLf
                   
                   While (Not RS.EOF)
                      strRstVal = RS.Fields("rstval").Value & ""
                      strRstVal = IIf(strRstVal = "", RS.Fields("rstcd").Value & "", strRstVal)
                      
                      strModRst = strRstVal & Space(3)
                      strModRst = strModRst & Format("" & RS.Fields("vfydt").Value, "####-##-##") & Space(1) & Format(Mid("" & RS.Fields("vfytm").Value, 1, 4), "0#:##") & Space(2)
                      strModRst = strModRst & "by " & RS.Fields("EmpNm").Value & vbCrLf
                      tmpToolTip = tmpToolTip & strModRst & Space(19)
                      RS.MoveNext
                   Wend
                End If
Skip:
        Set RS = Nothing
        MultiLine = 1
        If Trim(Replace(tmpToolTip, vbCrLf, "", 1, -1, vbBinaryCompare)) = "" Then
            ShowTip = False
            Exit Sub
        End If
        TipText = tmpToolTip
        TipWidth = 5500
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not ClearFg Then
        Call ClearRtn
    End If
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
    
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub txtPtId_LostFocus()
    Dim objDisease  As S2LIS_ReportLib.clsDisease
    Dim strWardId   As String
    Dim strTmp1     As String
    Dim strOCSTmp   As String
    
    If Screen.ActiveForm.Name <> Me.Name Then
        Exit Sub
    End If
    If Not gUsingInWardMenu Then
        If Screen.ActiveControl Is Nothing Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
        If Screen.ActiveControl.Name = chkPtList.Name Then Exit Sub
        If Screen.ActiveControl.Name = chkVerified.Name Then Exit Sub
        If Screen.ActiveControl.Name = txtSearchKey.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    End If
    DoEvents
    
    If MsgFg Then Exit Sub
      
On Error GoTo Errors
    If Trim(txtPtId.Text) = "" Then Exit Sub
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    Call ICSPatientMark(txtPtId.Text, enICSNum.ResultReview)
    '성바오로병원 외부수탁환자
    
    With objPatient
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SEXNM
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AGEDIV
            If lblDeptNm.Caption = "" Then lblDeptNm.Caption = .DeptNm
            strWardId = .WardId
            If strWardId <> "" Then
                If .RoomId <> "" Then strWardId = strWardId & "-" & .RoomId
            End If
            
            lblLocation.Caption = strWardId
            lblBedinDt.Caption = Format(.BedinDt, CS_DateMask)
            lblBedoutDt.Caption = Format(.BEDOUTDT, CS_DateMask)
            
            '** 추가루틴 By M.G.Choi 2005.08.11 ==========================================
            ' * 예수병원 현재환자위치 조회 함
            strOCSTmp = INPt_PArea(Trim(txtPtId.Text), .BedinDt)
            If strOCSTmp <> "" Then
                lblLocation.Caption = strWardId & "(" & strOCSTmp & ")"
            End If
            '=============================================================================
            
            If lblBedoutDt.Caption <> "" Then
'            '최근의 처방과를 가지고 온다.
'            '처방헤더의 deptcd,wardid,hosilid를 가지고온다.
                strTmp1 = objSQL.GetDeptInfo(txtPtId.Text)
                If strTmp1 <> "" Then
                    lblLocation.Caption = ""
                    lblDeptNm.Caption = medGetP(strTmp1, 1, COL_DIV)
                    lblDoctNm.Caption = medGetP(strTmp1, 2, COL_DIV)
                End If
                txtPtId.SetFocus
            End If
            
            Set objDisease = New S2LIS_ReportLib.clsDisease
            objDisease.PTid = txtPtId.Text
            lblDisease.Caption = objDisease.Disease
            Set objDisease = Nothing
            gPatientId = ""
            gPatientId = txtPtId.Text
            PtFg = True
            Call cmdCaution_Click
        Else
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            txtPtId = ""
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With

On Error GoTo Errors
    If ActiveControl.Name <> cmdRefresh.Name Then dtpFromDate.SetFocus
    If ClearFg Then Call cmdRefresh_Click
    ClearFg = False
    Exit Sub
Errors:
    Resume Next
End Sub

Private Function GetLastResultDate(ByVal sPTID As String) As String

    Dim RS      As Recordset
    Dim strSQL  As String
    
    strSQL = " select max(examdt) lastdt from " & T_LAB102 & _
             " where " & DBW("ptid = ", sPTID)
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        GetLastResultDate = Format(GetSystemDate, CS_DateLongFormat)
    Else
        If RS.Fields("lastdt").Value & "" = "" Then
            GetLastResultDate = Format(GetSystemDate, CS_DateLongFormat)
        Else
            GetLastResultDate = Format(RS.Fields("lastdt").Value, CS_DateMask)
        End If
    End If
    Set RS = Nothing
End Function

'% 텍스트결과 박스 더블클릭 - Larger Box Popup
Private Sub txtRstCmt_DblClick()
    With frmAPS905
        .rtfResultText.Visible = True
        .rtfResultText.TextRTF = txtRstCmt.Text
        Call HighlightText(.rtfResultText, "<< 검사 소견 >>", True, , &H4A4189)
        Call HighlightText(.rtfResultText, "<< Supplemental Report >>", False, , &H4A4189)
        .Show vbModal
    End With
    DoEvents
End Sub

'% 풋노트 박스 더블클릭 - Larger Box Popup
Private Sub txtSamCmt_DblClick()
    With frmAPS905
        .rtfResultText.Visible = False
        .rtfResultText.Text = txtRstCmt.Text & vbCrLf & vbCrLf
        .rtfResultText.Text = txtSamCmt.Text
        Call HighlightText(.rtfResultText, "<< 검사 소견 >>", True, , &H4A4189)
        Call HighlightText(.rtfResultText, "<< Supplemental Report >>", False, , &H4A4189)
        .rtfResultText.Visible = True
        .Show vbModal
    End With
    DoEvents

End Sub

'% 환자 검색 (ID 또는 성명으로...)
Private Sub txtSearchKey_GotFocus()
    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
    Dim objPtInfo   As clsPatient '  clsHosComSQLStmt
    Dim RS          As Recordset
    Dim itmx        As ListItem
    Dim strPtId     As String
    Dim lngSearch   As Long
    Dim ColCnt      As Long
    Dim RowCnt      As Long
    
    Set RS = New Recordset
    Set objPtInfo = New clsPatient 'clsHosComSQLStmt
    
    If KeyAscii = vbKeyReturn Then
        lngSearch = IIf(optSort(0).Value, 1, 2)  'True:환자ID, False:환자명
        If lngSearch = 1 And Not IsNumeric(txtSearchKey.Text) Then
            Set RS = Nothing
            Set objPtInfo = Nothing
            Exit Sub
        End If
        
        If chkVerified.Value = 0 Then
            If lngSearch = 2 And Len(txtSearchKey.Text) < 2 Then
                MsgBox "2문자 이상 입력하신후 검색하십시오.", vbInformation, "환자검색"
                txtSearchKey.SetFocus
                Exit Sub
            End If
            
            If optSort(0).Value = True Then
'                ColCnt = RS.OpenCursor(, objPtInfo.SqlPtntSearch("99", txtSearchKey.Text))
'                Rs.Open objPtInfo.SqlPtntSearch("99", txtSearchKey.Text), dbconn
                RS.Open objPtInfo.GetSQLPtNt("99", txtSearchKey.Text), DBConn
            Else
'                ColCnt = RS.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text))
'                Rs.Open objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text), dbconn
                RS.Open objPtInfo.GetSQLPtNt(lngSearch, txtSearchKey.Text), DBConn
            End If
        Else
            If optSort(0).Value = True Then
'                ColCnt = RS.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text, _
'                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)))
                RS.Open objPtInfo.GetSQLPtNt(lngSearch, txtSearchKey.Text, _
                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)), DBConn
'                Rs.Open objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text, _
'                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)), dbconn
            Else
'                ColCnt = RS.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text, _
'                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)))
                RS.Open objPtInfo.GetSQLPtNt(lngSearch, txtSearchKey.Text, _
                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)), DBConn
'                Rs.Open objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey.Text, _
'                              mvarDeptCd, Format(GetSystemDate, CS_DateDbFormat)), dbconn
            End If
        End If
        
        Me.MousePointer = vbHourglass
        
        lvwPtList.ListItems.Clear
'        If ColCnt > 0 Then
        If RS.EOF Then
            RowCnt = 0
            With lvwPtList
'                Do While (RS.FetchCursor(ColCnt))
                Do Until RS.EOF
                    RowCnt = RowCnt + 1
                    
                    Set itmx = .ListItems.Add(, , "" & RS.Fields("ptid").Value)
                    itmx.SubItems(1) = "" & RS.Fields("ptnm").Value
                    itmx.SubItems(2) = "" & RS.Fields("SSN").Value
                    itmx.SubItems(3) = "" & RS.Fields("DOB").Value
                    If Not IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(3) = Mid(itmx.SubItems(3), 1, 4) & "-01-01"
                    End If
                    If IsNumeric(Mid("" & RS.Fields("ssn").Value, 8, 1)) Then
                        itmx.SubItems(4) = IIf((Mid("" & RS.Fields("ssn").Value, 8, 1) Mod 2) = 1, "남", "여").Value
                    Else
                        itmx.SubItems(4) = "모름"
                    End If
                    If IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(4) = itmx.SubItems(4) & " / " & DateDiff("yyyy", itmx.SubItems(3), Now)
                    Else
                        itmx.SubItems(4) = itmx.SubItems(4) & " / ? "
                    End If
                    itmx.SubItems(5) = "" & RS.Fields("address").Value
                    itmx.SubItems(6) = "" & RS.Fields("telno").Value
                    
                    If RowCnt > 1000 Then Exit Do
                    RS.MoveNext
                Loop
            End With
        Else
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, Me.Caption
        End If
        
        Me.MousePointer = vbDefault
        
'        RS.CloseCursor
    End If
    Set RS = Nothing
    Set objPtInfo = Nothing
    
End Sub

'% Clear 루틴
Public Sub ClearRtn()
' Private Sub ClearRtn()
    chkSize.Caption = "참고치보기"
    
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    lblBedinDt.Caption = ""
    lblBedoutDt.Caption = ""
    lblDisease.Caption = ""
    rtfResult.Visible = False
    Call FieldClear
    Call TableClear
    ClearFg = True
    OrderFg = False
    MsgFg = False
    QueryFg = False
    OldRow = 0
End Sub

Private Sub FieldClear()

    lblDoctNm.Caption = ""
    lblCollectorNm.Caption = ""
    lblReceiverNm.Caption = ""
    lblVerifierNm.Caption = ""
    lblOrdDt.Caption = ""
    lblCollectDt.Caption = ""
    lblReceiveDt.Caption = ""
    lblVerifyDt.Caption = ""
    lblLisDoctNm.Caption = ""
    txtSamCmt.Text = ""
    txtRstCmt.Text = ""
    lblWorkArea.Caption = ""
    lblSpecimenNm.Caption = ""

End Sub

Private Sub TableClear()
    tblOrdSheet.MaxRows = 0
    tblOrdSheet.MaxRows = 25
    
    tblResult.MaxRows = 0
    tblResult.MaxRows = 127
End Sub

'% 결과 Part Clear
Private Sub ResultClear()
    txtRstCmt.Text = ""
    txtSamCmt.Text = ""
      
    lblWorkArea.Caption = ""
    lblSpecimenNm.Caption = ""
   
    rtfResult.Tag = ""
    rtfResult.Text = ""
    ResultFg = False
   
    With tblResult
        '결과테이블 Clear
        .Row = -1:  .Col = -1
        .BlockMode = True
        .FontBold = False
        .Action = ActionClearText
        .BlockMode = False
        '검사명/결과 컬럼 Bold
        .Row = -1: .Col = 2: .Col2 = 3
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        'High/Low field font 지정
        .Row = -1: .Col = 5: .Col2 = 5
        .BlockMode = True
        .FontName = "돋움"
        .BlockMode = False
        .RowsFrozen = 0
    End With
   

End Sub

Public Sub Call_ToDate_LostFocus()

    If Not gUsingInWardMenu Then
        On Error Resume Next
        If ActiveControl.Name = cmdExit.Name Then Exit Sub
        If ActiveControl.Name = cmdClear.Name Then Exit Sub
        If ActiveControl.Name = chkPtList.Name Then Exit Sub
        
    End If
    Call cmdRefresh_Click
End Sub


Public Sub Call_PtId_KeyPress()
    Dim objDisease  As S2LIS_ReportLib.clsDisease
    Dim strLastDT   As String
    
On Error GoTo Errors
    
    strLastDT = GetLastResultDate(txtPtId.Text)
    dtpFromDate.Value = DateAdd("d", -7, strLastDT)     '외래

    If txtPtId.Text = "" Then
        If Screen.ActiveForm.Name = Me.Name Then txtPtId.SetFocus
        Exit Sub
    End If

      With objPatient
         If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SEXNM
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AGEDIV
            lblDeptNm.Caption = .DeptNm
            txtPtId.SetFocus
            PtFg = True
            
            Set objDisease = New S2LIS_ReportLib.clsDisease
            objDisease.PTid = txtPtId.Text
            lblDisease.Caption = objDisease.Disease
            Set objDisease = Nothing

            gPatientId = txtPtId.Text
            ClearFg = False
         Else
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            Me.Enabled = True
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
         End If
      End With
      If ClearFg Then Call dtpToDate.SetFocus
      Exit Sub
Errors:
    Resume Next

End Sub

Private Sub DisplayLABComment(ByVal iRow As Long)
    Dim sBedinDt As String
    
    tblOrdSheet.Row = iRow
    tblOrdSheet.Col = enREVIEW1.tcORDDATE
    sBedinDt = tblOrdSheet.Value
    
    With frmLabReport
        .ZOrder 0
        DoEvents
        .PTid = txtPtId.Text
        .BedinDt = Format(sBedinDt, CS_DateDbFormat)
        Call .StartQuery
        .Show 1
    End With
    
End Sub

'** 예수병원 추가 By M.G. Choi ========================================
Private Function INPt_PArea(ByVal pPtID As String, ByVal pINDate As String) As String
    Dim objOCSSql   As New clsLISTransfer
    
    INPt_PArea = objOCSSql.OCS_INPT_PArea(pPtID, pINDate)
    
    Set objOCSSql = Nothing
    
End Function
'======================================================================

Private Function GetLisDoctNm(ByVal pTestCd As String) As String
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    
    On Error Resume Next
    
    strSQL = " select b.empnm " & _
             "   from " & T_LAB031 & " a, " & T_COM006 & " b " & _
             "  where a.cdindex = " & DBS(LC2_DoctTest) & _
             "    and a.cdval2 = " & DBS(pTestCd) & _
             "    and a.cdval1 = b.empid "
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        GetLisDoctNm = RS.Fields("empnm").Value & ""
    End If
    
    RS.Close
    Set RS = Nothing

End Function

