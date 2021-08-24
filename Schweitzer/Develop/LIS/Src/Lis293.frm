VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm293SpecialTest 
   BackColor       =   &H00DBE6E6&
   Caption         =   "임상병리 특수검사 결과 입력"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14580
   WindowState     =   2  '최대화
   Begin VB.Frame frmSMS 
      BackColor       =   &H00F8E4D8&
      Caption         =   "SMS전송"
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
      Height          =   5415
      Left            =   6420
      TabIndex        =   137
      Top             =   1890
      Width           =   4515
      Begin VB.TextBox txtExDtNo 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   152
         Tag             =   "opt"
         Top             =   2190
         Width           =   2325
      End
      Begin VB.TextBox txtExDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   151
         Tag             =   "opt"
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox txtExDtId 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   150
         Tag             =   "opt"
         Top             =   1800
         Width           =   1305
      End
      Begin VB.CommandButton cmdTrans 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전송"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   149
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   3030
         Style           =   1  '그래픽
         TabIndex        =   148
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.TextBox txtTransId 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   147
         Tag             =   "opt"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtTransNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   146
         Tag             =   "opt"
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtTransNo 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   145
         Tag             =   "opt"
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtDtId 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   144
         Tag             =   "opt"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   143
         Tag             =   "opt"
         Top             =   1020
         Width           =   1005
      End
      Begin VB.TextBox txtDetpCd 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   142
         Tag             =   "opt"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDeptNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   141
         Tag             =   "opt"
         Top             =   2580
         Width           =   3195
      End
      Begin VB.TextBox txtDtNo 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   140
         Tag             =   "opt"
         Top             =   1410
         Width           =   2325
      End
      Begin VB.TextBox txtTransDt 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   139
         Tag             =   "opt"
         Top             =   4170
         Width           =   3195
      End
      Begin VB.TextBox txtTestCd 
         Appearance      =   0  '평면
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
         Height          =   360
         Left            =   5100
         MaxLength       =   15
         TabIndex        =   138
         Tag             =   "opt"
         Top             =   1350
         Width           =   1305
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   9
         Left            =   180
         TabIndex        =   153
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "전송자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1905
         Index           =   11
         Left            =   180
         TabIndex        =   154
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   3360
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
         Caption         =   "수신자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   20
         Left            =   180
         TabIndex        =   155
         Top             =   2970
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "메시지"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   21
         Left            =   180
         TabIndex        =   156
         Top             =   4200
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "전송일시"
         Appearance      =   0
      End
      Begin RichTextLib.RichTextBox rtfMessage 
         Height          =   1170
         Left            =   1140
         TabIndex        =   157
         Top             =   2970
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis293.frx":0000
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   22
         Left            =   180
         TabIndex        =   158
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   765
         Index           =   23
         Left            =   1110
         TabIndex        =   159
         Top             =   1020
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   765
         Index           =   24
         Left            =   1110
         TabIndex        =   160
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
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
         Caption         =   "주치의"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSMS 
      BackColor       =   &H008080FF&
      Caption         =   "SMS"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   7830
      Style           =   1  '그래픽
      TabIndex        =   106
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdCaution 
      BackColor       =   &H008080FF&
      Caption         =   "Caution"
      Height          =   345
      Left            =   3960
      MaskColor       =   &H8000000F&
      Style           =   1  '그래픽
      TabIndex        =   134
      Top             =   1620
      Width           =   1005
   End
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
      Height          =   7035
      Left            =   3960
      TabIndex        =   107
      Top             =   2010
      Width           =   7005
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
         TabIndex        =   135
         Text            =   "Caution 수정은 감염관리실에 요청하여 주십시요."
         Top             =   6090
         Width           =   6795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "종 료"
         Height          =   495
         Left            =   5250
         TabIndex        =   129
         Top             =   6480
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         Caption         =   "특이소견"
         Enabled         =   0   'False
         Height          =   2685
         Left            =   90
         TabIndex        =   121
         Top             =   3390
         Width           =   6795
         Begin VB.CheckBox Check1 
            Caption         =   "Fungus"
            Height          =   195
            Index           =   14
            Left            =   5610
            TabIndex        =   127
            Top             =   270
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "C.difficile"
            Height          =   195
            Index           =   13
            Left            =   4200
            TabIndex        =   126
            Top             =   270
            Width           =   1155
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tb"
            Height          =   195
            Index           =   12
            Left            =   3360
            TabIndex        =   125
            Top             =   270
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "AFB"
            Height          =   195
            Index           =   11
            Left            =   2370
            TabIndex        =   124
            Top             =   270
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRE"
            Height          =   195
            Index           =   10
            Left            =   1290
            TabIndex        =   123
            Top             =   270
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRSA"
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   122
            Top             =   270
            Width           =   885
         End
         Begin RichTextLib.RichTextBox RichText 
            Height          =   1935
            Left            =   150
            TabIndex        =   128
            Top             =   570
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3413
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"Lis293.frx":009D
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Drug Allergy"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   90
         TabIndex        =   117
         Top             =   2250
         Width           =   6795
         Begin VB.TextBox txtDrug 
            Height          =   315
            Left            =   180
            TabIndex        =   120
            Text            =   "Text1"
            Top             =   570
            Width           =   6465
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RadioContrast"
            Height          =   195
            Index           =   8
            Left            =   2640
            TabIndex        =   119
            Top             =   270
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Penicillin"
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   118
            Top             =   270
            Width           =   2265
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Viral Marker"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   90
         TabIndex        =   108
         Top             =   870
         Width           =   6795
         Begin VB.TextBox txtVival 
            Height          =   315
            Left            =   210
            TabIndex        =   116
            Text            =   "Text1"
            Top             =   870
            Width           =   6465
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HAV lgM"
            Height          =   195
            Index           =   6
            Left            =   2610
            TabIndex        =   115
            Top             =   570
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HBc lgM"
            Height          =   195
            Index           =   5
            Left            =   210
            TabIndex        =   114
            Top             =   570
            Width           =   1845
         End
         Begin VB.CheckBox Check1 
            Caption         =   "기 타"
            Height          =   195
            Index           =   4
            Left            =   5670
            TabIndex        =   113
            Top             =   330
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HCV"
            Height          =   195
            Index           =   3
            Left            =   4080
            TabIndex        =   112
            Top             =   330
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HBsAg"
            Height          =   195
            Index           =   2
            Left            =   2610
            TabIndex        =   111
            Top             =   330
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VDRL"
            Height          =   195
            Index           =   1
            Left            =   1290
            TabIndex        =   110
            Top             =   330
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HIV"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   109
            Top             =   330
            Width           =   975
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   130
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
         TabIndex        =   131
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
         TabIndex        =   132
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
         TabIndex        =   133
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
   Begin RichTextLib.RichTextBox txtPNote 
      Height          =   705
      Left            =   2550
      TabIndex        =   58
      Top             =   7200
      Visible         =   0   'False
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   1244
      _Version        =   393217
      BackColor       =   15658734
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Lis293.frx":012C
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
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   8
      Left            =   135
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   6390
      Width           =   2505
      _ExtentX        =   4419
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
      Caption         =   "처방 Remark"
      Appearance      =   0
   End
   Begin VB.TextBox txtOCSMesg 
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
      Height          =   1590
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   101
      ToolTipText     =   "검사 리마크를 입력하세요."
      Top             =   6765
      Width           =   2520
   End
   Begin VB.CommandButton cmdTemplet2 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&T2 Report"
      Height          =   510
      Left            =   2895
      Style           =   1  '그래픽
      TabIndex        =   100
      Tag             =   "135"
      Top             =   8535
      Visible         =   0   'False
      Width           =   990
   End
   Begin DRcontrol1.DrFrame fraResult 
      Height          =   3735
      Left            =   6240
      TabIndex        =   42
      Top             =   4575
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6588
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   15518662
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwResult 
         Height          =   3600
         Left            =   60
         TabIndex        =   43
         Top             =   60
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   6350
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16775406
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Template Text"
            Object.Width           =   12435
         EndProperty
      End
   End
   Begin VB.CommandButton cmdTemplet 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&T1 Report"
      Height          =   510
      Left            =   1920
      Style           =   1  '그래픽
      TabIndex        =   81
      Tag             =   "135"
      Top             =   8535
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Bone Marrow Report"
      Height          =   510
      Left            =   3915
      Style           =   1  '그래픽
      TabIndex        =   74
      Tag             =   "135"
      Top             =   8535
      Width           =   1980
   End
   Begin VB.PictureBox picESign 
      Height          =   500
      Left            =   2070
      ScaleHeight     =   435
      ScaleWidth      =   1140
      TabIndex        =   55
      Top             =   8610
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00EDE2ED&
      Caption         =   "결과지 출력 (&P)"
      Height          =   510
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   54
      Tag             =   "135"
      Top             =   8535
      Width           =   1845
   End
   Begin VB.ListBox lstTstNo 
      BackColor       =   &H00FBEDEA&
      Height          =   2220
      Left            =   135
      TabIndex        =   24
      Top             =   1470
      Width           =   2520
   End
   Begin VB.ListBox lstLabNo 
      BackColor       =   &H00F4FAFF&
      Height          =   2220
      Left            =   135
      TabIndex        =   23
      Top             =   4110
      Width           =   2520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1170
      Left            =   75
      TabIndex        =   8
      Top             =   -45
      Width           =   2625
      Begin VB.OptionButton optInput 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과 확인 대상 리스트"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   10
         Top             =   905
         Width           =   2115
      End
      Begin VB.OptionButton optInput 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수 번호별 결과 입력"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   9
         Top             =   530
         Width           =   2145
      End
      Begin VB.Label lblDoctID 
         Height          =   255
         Left            =   390
         TabIndex        =   136
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdVerify 
      BackColor       =   &H00F4F0F2&
      Caption         =   "확 인 (&V)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrintWS 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Worksheet 출력"
      Height          =   510
      Left            =   5910
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "135"
      Top             =   8535
      Width           =   1905
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저 장 (&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   13035
      Top             =   615
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtFNote 
      Height          =   705
      Left            =   2670
      TabIndex        =   59
      Top             =   8370
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1244
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Lis293.frx":03AF
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
   Begin DRcontrol1.DrFrame fraLastRst 
      Height          =   3390
      Left            =   2925
      TabIndex        =   82
      Top             =   1980
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   5980
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
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
      Begin MSComctlLib.ListView lvwLResult 
         Height          =   2955
         Left            =   30
         TabIndex        =   84
         Top             =   390
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5212
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "접수번호"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "보고일자"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "보고시간"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "진료과/병동"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "처방의"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Sex/Age"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "workarea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "accdt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "accseq"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "testcd"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "mfyseq"
            Object.Width           =   0
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Left            =   30
         TabIndex        =   83
         Top             =   60
         Width           =   3120
         _ExtentX        =   5503
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
         Alignment       =   1
         Caption         =   "특수검사 최근결과 리스트"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin VB.Frame fraTest 
      BackColor       =   &H00E8EEEE&
      Height          =   3420
      Left            =   4410
      TabIndex        =   85
      Top             =   1980
      Visible         =   0   'False
      Width           =   9105
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Cl&ose"
         Height          =   540
         Left            =   7545
         Style           =   1  '그래픽
         TabIndex        =   90
         Top             =   165
         Width           =   1395
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00F4F0F2&
         Caption         =   "&Apply"
         Height          =   510
         Left            =   6030
         Style           =   1  '그래픽
         TabIndex        =   89
         Top             =   180
         Width           =   1470
      End
      Begin FPSpread.vaSpread tblData 
         Height          =   2520
         Left            =   30
         TabIndex        =   86
         Top             =   855
         Width           =   9030
         _Version        =   196608
         _ExtentX        =   15928
         _ExtentY        =   4445
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         GrayAreaBackColor=   15265518
         MaxCols         =   9
         MaxRows         =   6
         ShadowColor     =   15265518
         ShadowDark      =   15265518
         SpreadDesigner  =   "Lis293.frx":0632
      End
      Begin VB.Label lblTitle1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   315
         TabIndex        =   88
         Top             =   300
         Width           =   3705
      End
      Begin VB.Shape shpSubMenu 
         BackColor       =   &H80000001&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   135
         Top             =   195
         Width           =   4095
      End
      Begin VB.Label lblTest 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   60
         TabIndex        =   87
         Top             =   270
         Width           =   4095
      End
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00DBE6E6&
      Height          =   8370
      Left            =   2730
      ScaleHeight     =   8310
      ScaleWidth      =   11670
      TabIndex        =   11
      Top             =   45
      Width           =   11730
      Begin MedControls1.LisLabel lblMode 
         Height          =   360
         Left            =   8505
         TabIndex        =   40
         Top             =   75
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   635
         BackColor       =   14737632
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "결과 등록"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdFont 
         BackColor       =   &H00C4DBDD&
         Caption         =   "&Font"
         Height          =   360
         Left            =   9900
         Style           =   1  '그래픽
         TabIndex        =   57
         Top             =   75
         Width           =   465
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00E7BAB4&
         Caption         =   "색상표"
         Height          =   360
         Left            =   10365
         Style           =   1  '그래픽
         TabIndex        =   53
         Top             =   75
         Width           =   660
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00DAC7DA&
         Caption         =   "&Edit"
         Height          =   360
         Left            =   11025
         Style           =   1  '그래픽
         TabIndex        =   47
         Top             =   75
         Width           =   465
      End
      Begin VB.TextBox txtWorkArea 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
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
         Height          =   210
         Left            =   1395
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "BM"
         Top             =   150
         Width           =   300
      End
      Begin VB.TextBox txtAccDt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
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
         Height          =   225
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "010515"
         Top             =   150
         Width           =   630
      End
      Begin VB.TextBox txtAccSeq 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
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
         Height          =   210
         Left            =   2775
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "10012"
         Top             =   150
         Width           =   525
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Left            =   135
         TabIndex        =   12
         Top             =   75
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483634
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
      Begin MedControls1.LisLabel lblLastDate 
         Height          =   360
         Left            =   8505
         TabIndex        =   30
         Top             =   75
         Visible         =   0   'False
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   635
         BackColor       =   16702665
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최근 결과 등록일자"
         Appearance      =   0
      End
      Begin TabDlg.SSTab sstRst 
         Height          =   8160
         Left            =   3465
         TabIndex        =   31
         Top             =   45
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   14393
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   706
         BackColor       =   14411494
         TabCaption(0)   =   "Text 결과 "
         TabPicture(0)   =   "Lis293.frx":0E18
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cboTemplate"
         Tab(0).Control(1)=   "cboAppend"
         Tab(0).Control(2)=   "txtRst"
         Tab(0).Control(3)=   "txtCRst"
         Tab(0).Control(4)=   "slider"
         Tab(0).Control(5)=   "lblMesg"
         Tab(0).Control(6)=   "Label1"
         Tab(0).Control(7)=   "Label11"
         Tab(0).Control(8)=   "Shape4"
         Tab(0).Control(9)=   "Shape5"
         Tab(0).Control(10)=   "Shape6"
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "최근결과 (01/06/30 13:00)"
         TabPicture(1)   =   "Lis293.frx":0E34
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "txtLastRst"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "관련검사결과"
         TabPicture(2)   =   "Lis293.frx":0E50
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Line1"
         Tab(2).Control(1)=   "tblResult"
         Tab(2).Control(2)=   "lblVfyDtTm"
         Tab(2).Control(3)=   "lblColDtTm"
         Tab(2).Control(4)=   "LisLabel3"
         Tab(2).Control(5)=   "LisLabel2"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "이미지 조회"
         TabPicture(3)   =   "Lis293.frx":0E6C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraImage"
         Tab(3).ControlCount=   1
         Begin VB.ComboBox cboTemplate 
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -73500
            Style           =   2  '드롭다운 목록
            TabIndex        =   33
            Top             =   507
            Width           =   2805
         End
         Begin VB.ComboBox cboAppend 
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -69315
            Style           =   2  '드롭다운 목록
            TabIndex        =   32
            Top             =   507
            Width           =   2250
         End
         Begin RichTextLib.RichTextBox txtRst 
            Height          =   7020
            Left            =   -74880
            TabIndex        =   34
            Top             =   1080
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   12383
            _Version        =   393217
            BackColor       =   15924219
            ScrollBars      =   3
            RightMargin     =   9000
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Lis293.frx":0E88
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
         Begin RichTextLib.RichTextBox txtLastRst 
            Height          =   7605
            Left            =   105
            TabIndex        =   35
            Top             =   507
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   13414
            _Version        =   393217
            BackColor       =   15658734
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"Lis293.frx":10FE
         End
         Begin RichTextLib.RichTextBox txtCRst 
            Height          =   5790
            Left            =   -74850
            TabIndex        =   36
            Top             =   1095
            Visible         =   0   'False
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   10213
            _Version        =   393217
            BackColor       =   15857140
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"Lis293.frx":131E
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
         Begin MSComctlLib.Slider slider 
            Height          =   7125
            Left            =   -74970
            TabIndex        =   37
            Top             =   990
            Visible         =   0   'False
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   12568
            _Version        =   393216
            Orientation     =   1
            LargeChange     =   50
            SmallChange     =   5
            TickStyle       =   3
         End
         Begin MedControls1.LisLabel LisLabel2 
            Height          =   330
            Left            =   -74850
            TabIndex        =   48
            Top             =   675
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16702665
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            AutoSize        =   -1  'True
            Caption         =   "채혈일시 : "
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel3 
            Height          =   330
            Left            =   -74850
            TabIndex        =   49
            Top             =   1020
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16702665
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            AutoSize        =   -1  'True
            Caption         =   "보고일시 : "
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblColDtTm 
            Height          =   330
            Left            =   -73785
            TabIndex        =   50
            Top             =   675
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   582
            BackColor       =   16510442
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblVfyDtTm 
            Height          =   330
            Left            =   -73785
            TabIndex        =   51
            Top             =   1020
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   582
            BackColor       =   16510442
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   0
         End
         Begin FPSpread.vaSpread tblResult 
            Height          =   6525
            Left            =   -74850
            TabIndex        =   52
            Top             =   1500
            Width           =   7740
            _Version        =   196608
            _ExtentX        =   13653
            _ExtentY        =   11509
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
            GridColor       =   14013909
            GridShowVert    =   0   'False
            MaxCols         =   8
            OperationMode   =   1
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            ShadowText      =   0
            SpreadDesigner  =   "Lis293.frx":1546
            TextTip         =   4
         End
         Begin VB.Frame fraImage 
            BackColor       =   &H00DBE6E6&
            Height          =   7710
            Left            =   -74940
            TabIndex        =   75
            Top             =   435
            Width           =   7950
            Begin VB.Frame frmImage 
               Height          =   7185
               Left            =   420
               TabIndex        =   80
               Top             =   195
               Visible         =   0   'False
               Width           =   7095
               Begin VB.Image imgImage 
                  BorderStyle     =   1  '단일 고정
                  Height          =   7005
                  Left            =   45
                  Picture         =   "Lis293.frx":5244
                  Stretch         =   -1  'True
                  Top             =   135
                  Width           =   7005
               End
            End
            Begin MSComctlLib.ListView lvwHxList 
               Height          =   3330
               Left            =   90
               TabIndex        =   78
               Top             =   480
               Width           =   3690
               _ExtentX        =   6509
               _ExtentY        =   5874
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   14737632
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin MSComctlLib.ListView lvwList 
               Height          =   3255
               Left            =   75
               TabIndex        =   76
               Top             =   4260
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   5741
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   14737632
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin VB.Image imgList 
               BorderStyle     =   1  '단일 고정
               Height          =   3600
               Left            =   3990
               Picture         =   "Lis293.frx":1AD9D
               Stretch         =   -1  'True
               Top             =   3975
               Width           =   3795
            End
            Begin VB.Image imgHx 
               BorderStyle     =   1  '단일 고정
               Height          =   3600
               Left            =   3990
               Picture         =   "Lis293.frx":3006B
               Stretch         =   -1  'True
               Top             =   225
               Width           =   3795
            End
            Begin VB.Shape Shape11 
               BackColor       =   &H00E4F3F8&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00808080&
               Height          =   3720
               Left            =   3855
               Top             =   3870
               Width           =   4035
            End
            Begin VB.Shape Shape10 
               BackColor       =   &H00E4F3F8&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00808080&
               Height          =   3720
               Left            =   3855
               Top             =   135
               Width           =   4035
            End
            Begin VB.Label Label14 
               Alignment       =   2  '가운데 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "◈ 환자 Histroy 이미지 리스트"
               Height          =   180
               Left            =   90
               TabIndex        =   79
               Tag             =   "104"
               Top             =   225
               Width           =   2460
            End
            Begin VB.Label Label5 
               Alignment       =   2  '가운데 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "◈ 이미지 리스트 정보"
               Height          =   180
               Left            =   105
               TabIndex        =   77
               Tag             =   "104"
               Top             =   4005
               Width           =   1815
            End
            Begin VB.Shape Shape8 
               BackColor       =   &H00E4F3F8&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00808080&
               Height          =   3735
               Left            =   30
               Top             =   3870
               Width           =   3795
            End
            Begin VB.Shape Shape9 
               BackColor       =   &H00E4F3F8&
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00808080&
               Height          =   3720
               Left            =   30
               Top             =   135
               Width           =   3795
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   -74820
            X2              =   -67125
            Y1              =   1425
            Y2              =   1425
         End
         Begin VB.Label lblMesg 
            BackColor       =   &H00F2FBFB&
            BackStyle       =   0  '투명
            Caption         =   "  ☞ [F2]버튼을 누르시면 커서가 다음 결과입력 위치로 이동합니다."
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   -74835
            TabIndex        =   41
            Top             =   885
            Width           =   7755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Template Code"
            ForeColor       =   &H00313D46&
            Height          =   180
            Left            =   -74850
            TabIndex        =   39
            Top             =   585
            Width           =   1440
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Append Code "
            ForeColor       =   &H00313D46&
            Height          =   180
            Left            =   -70530
            TabIndex        =   38
            Top             =   585
            Width           =   1350
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00808080&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00E0E0E0&
            Height          =   225
            Left            =   -74850
            Shape           =   4  '둥근 사각형
            Top             =   855
            Width           =   7785
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00FFF9F7&
            BackStyle       =   1  '투명하지 않음
            Height          =   300
            Left            =   -74880
            Shape           =   4  '둥근 사각형
            Top             =   510
            Width           =   1365
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFF9F7&
            BackStyle       =   1  '투명하지 않음
            Height          =   300
            Left            =   -70680
            Shape           =   4  '둥근 사각형
            Top             =   510
            Width           =   1350
         End
      End
      Begin VB.Frame fraRst 
         BackColor       =   &H00DBE6E6&
         Height          =   7710
         Left            =   150
         TabIndex        =   13
         Top             =   525
         Width           =   3255
         Begin VB.CommandButton cmdOrderView 
            BackColor       =   &H00F4F0F2&
            Caption         =   "처방별조회"
            Height          =   360
            Left            =   2070
            Style           =   1  '그래픽
            TabIndex        =   105
            Top             =   1005
            Width           =   1100
         End
         Begin MedControls1.LisLabel lblOrdD 
            Height          =   315
            Index           =   9
            Left            =   75
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   3015
            Width           =   930
            _ExtentX        =   1640
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
            Caption         =   "처방의"
            Appearance      =   0
         End
         Begin VB.CheckBox chkApply 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Add to Result"
            ForeColor       =   &H00404080&
            Height          =   195
            Left            =   1680
            TabIndex        =   73
            Top             =   6015
            Width           =   1440
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00CEE1E1&
            Caption         =   "X"
            Height          =   330
            Index           =   0
            Left            =   2835
            Style           =   1  '그래픽
            TabIndex        =   72
            Top             =   6240
            Width           =   270
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00CEE1E1&
            Caption         =   "X"
            Height          =   330
            Index           =   1
            Left            =   2835
            Style           =   1  '그래픽
            TabIndex        =   71
            Top             =   6600
            Width           =   270
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00CEE1E1&
            Caption         =   "X"
            Height          =   330
            Index           =   2
            Left            =   2835
            Style           =   1  '그래픽
            TabIndex        =   70
            Top             =   6945
            Width           =   270
         End
         Begin VB.CommandButton cmdRstCd 
            BackColor       =   &H00A2BEBF&
            Caption         =   "..."
            Height          =   330
            Index           =   2
            Left            =   2505
            Style           =   1  '그래픽
            TabIndex        =   69
            Top             =   6945
            Width           =   315
         End
         Begin VB.CommandButton cmdRstCd 
            BackColor       =   &H00A2BEBF&
            Caption         =   "..."
            Height          =   330
            Index           =   1
            Left            =   2505
            Style           =   1  '그래픽
            TabIndex        =   68
            Top             =   6600
            Width           =   315
         End
         Begin VB.CommandButton cmdRstCd 
            BackColor       =   &H00A2BEBF&
            Caption         =   "..."
            Height          =   330
            Index           =   0
            Left            =   2490
            Style           =   1  '그래픽
            TabIndex        =   67
            Top             =   6240
            Width           =   315
         End
         Begin MedControls1.LisLabel lblPtNm 
            Height          =   330
            Left            =   1305
            TabIndex        =   27
            Top             =   1380
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            BackColor       =   13752531
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
            Alignment       =   1
            Appearance      =   0
         End
         Begin VB.ComboBox cboRemark 
            BackColor       =   &H00F1F5F4&
            Height          =   300
            Left            =   75
            Style           =   2  '드롭다운 목록
            TabIndex        =   21
            Top             =   4755
            Width           =   3075
         End
         Begin MedControls1.LisLabel lblPtId 
            Height          =   330
            Left            =   75
            TabIndex        =   28
            Top             =   1380
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            BackColor       =   13752531
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
            Alignment       =   1
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPtSA 
            Height          =   330
            Left            =   2385
            TabIndex        =   29
            Top             =   1380
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
            BackColor       =   13752531
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
            Alignment       =   1
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   330
            Index           =   0
            Left            =   75
            TabIndex        =   61
            Top             =   6255
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   582
            BackColor       =   11259341
            ForeColor       =   16448
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "1"
         End
         Begin MedControls1.LisLabel lblRstCd 
            Height          =   330
            Index           =   0
            Left            =   405
            TabIndex        =   62
            Top             =   6255
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   582
            BackColor       =   13558241
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "AML"
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel LisLabel6 
            Height          =   330
            Left            =   75
            TabIndex        =   63
            Top             =   6615
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   582
            BackColor       =   11259341
            ForeColor       =   16448
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "2"
         End
         Begin MedControls1.LisLabel lblRstCd 
            Height          =   330
            Index           =   1
            Left            =   405
            TabIndex        =   64
            Top             =   6615
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   582
            BackColor       =   13558241
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "AML"
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel LisLabel8 
            Height          =   330
            Left            =   75
            TabIndex        =   65
            Top             =   6960
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   582
            BackColor       =   11259341
            ForeColor       =   16448
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "3"
         End
         Begin MedControls1.LisLabel lblRstCd 
            Height          =   330
            Index           =   2
            Left            =   405
            TabIndex        =   66
            Top             =   6960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            BackColor       =   13558241
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "AML"
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   10
            Left            =   75
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   1845
            Width           =   930
            _ExtentX        =   1640
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
            Index           =   1
            Left            =   75
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   2235
            Width           =   930
            _ExtentX        =   1640
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
            Caption         =   "진료과"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   2625
            Width           =   930
            _ExtentX        =   1640
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
            Caption         =   "병  동"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   1035
            Width           =   975
            _ExtentX        =   1720
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
            Caption         =   "환자 정보"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   3405
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "임상진단"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   6
            Left            =   75
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   4425
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "검체 Remark"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   7
            Left            =   75
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   5910
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "결과코드"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   3
            Left            =   1770
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   3330
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
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
            Caption         =   "연락처"
            Appearance      =   0
         End
         Begin VB.Label lblOrdDoct 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1035
            TabIndex        =   104
            Top             =   3015
            Width           =   2085
         End
         Begin VB.Label lblTelNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2730
            TabIndex        =   91
            Top             =   3330
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label lblTestNm 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   570
            Left            =   225
            TabIndex        =   56
            Top             =   315
            Width           =   2805
            WordWrap        =   -1  'True
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00FBEDEA&
            BackStyle       =   1  '투명하지 않음
            Height          =   735
            Left            =   75
            Top             =   210
            Width           =   3105
         End
         Begin VB.Label lblRemark 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Height          =   810
            Left            =   75
            TabIndex        =   46
            Top             =   5055
            Width           =   3075
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMajDoct 
            Caption         =   "주치의"
            Height          =   195
            Left            =   1845
            TabIndex        =   45
            Top             =   3270
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblWardId 
            Caption         =   "WardId"
            Height          =   195
            Left            =   1185
            TabIndex        =   44
            Top             =   3270
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblAttr 
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            Caption         =   "1234567"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   75
            TabIndex        =   22
            Top             =   3750
            Width           =   3060
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblWard 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1035
            TabIndex        =   20
            Top             =   2625
            Width           =   2085
         End
         Begin VB.Label lblDeptCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1035
            TabIndex        =   19
            Top             =   2235
            Width           =   2100
         End
         Begin VB.Label lblSpecimen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D1D8D3&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1035
            TabIndex        =   18
            Top             =   1845
            Width           =   2100
         End
      End
      Begin VB.ListBox lstTest 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF9F7&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   1260
         TabIndex        =   14
         Top             =   450
         Visible         =   0   'False
         Width           =   5370
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   17
         Top             =   135
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2595
         TabIndex        =   16
         Top             =   135
         Width           =   195
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00F1F5F4&
         BackStyle       =   1  '투명하지 않음
         Height          =   360
         Left            =   1290
         Shape           =   4  '둥근 사각형
         Top             =   75
         Width           =   2055
      End
      Begin VB.Label Label40 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수 번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   15
         Tag             =   "25601"
         Top             =   165
         Width           =   840
      End
   End
   Begin VB.Label lblFNote 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ Foot Note"
      Height          =   180
      Left            =   2700
      TabIndex        =   60
      Tag             =   "25607"
      Top             =   7425
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label13 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 대상 리스트"
      Height          =   180
      Left            =   195
      TabIndex        =   26
      Tag             =   "104"
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 검사 항목"
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Tag             =   "104"
      Top             =   1275
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E4F3F8&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   2640
      Left            =   75
      Top             =   1140
      Width           =   2625
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E4F3F8&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   4665
      Left            =   75
      Top             =   3795
      Width           =   2625
   End
End
Attribute VB_Name = "frm293SpecialTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cNoLastRst As String = "(최근 결과 없음)"
Const cSep As Integer = 30

Const cGenMode As Long = &HC0C0C0       ' 현재 상태를 색으로 표시
Const cInsMode As Long = &HFF0000       ' 일반 모드
Const cMfyMode As Long = &HFF&          ' 수정 모드

Private fTestCd As String               ' 현재 에디트 중인 검사 항목
Private fFNSeq As Integer               ' Foot Note Seq. Number
Private fStatus As String               ' 현재 Status
Private fMfySeq As Integer              ' 수정 횟수 (결과 Reading시 필요)
Private fRstType  As String             ' 결과유형
Private blnTmpDisplay  As Boolean

Private fLWorkArea As String, fLAccDt As String, fLAccSeq As String, fLM As Integer

Private WithEvents objDesc      As frmTmpBMReport
Attribute objDesc.VB_VarHelpID = -1
Private WithEvents objTemplet   As frmTmpResult
Attribute objTemplet.VB_VarHelpID = -1
Private WithEvents objTemplet2  As frmTmpEIReport
Attribute objTemplet2.VB_VarHelpID = -1

'Private WithEvents mnuPopup As menu
'Private WithEvents mnuSave  As menu

Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_SAVE& = 1

Private objETest        As New clsLISSpecialTest
Private dicRstFields    As New clsDictionary

Private blnInitFg   As Boolean
Private blnMsgFg    As Boolean

Private strRtfHead      As String
Private strRtfEnd       As String
Private strImagePath    As String

Private mWorkarea   As String   'Workarea
Private mAccDt      As String   'AccDt
Private mAccSeq     As String   'AccSeq
Private mTestCd     As String   '검사코드

Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset
Dim strRcvDt            As String

Private Sub cboAppend_Click()
    
    Dim sApdCd As String

    If cboAppend.ListIndex < 0 Then Exit Sub
    
    sApdCd = medGetP(cboAppend.List(cboAppend.ListIndex), 1, vbTab)
       
    'txtRst.SelText = objETest.GetAppendText(fTestCd, sApdCd)
    Call objETest.LoadRstTemplate(sApdCd, lvwResult)
    
    cboAppend.ListIndex = -1
    'txtRst.SetFocus
    fraResult.Visible = True
    fraResult.ZOrder 0
    lvwResult.SetFocus

End Sub

Private Sub cboRemark_Click()
    
    Dim sRMCd As String, sRMNm As String
 
    If cboRemark.ListIndex < 0 Then Exit Sub

    sRMCd = cboRemark.List(cboRemark.ListIndex)
    If sRMCd = LIS_Nothing Then lblRemark.Caption = "": Exit Sub
    lblRemark.Caption = objETest.GetRemark(sRMCd)

End Sub

Private Sub cboTemplate_Click()
    
    Dim sTemp As String, sRType As String, sTCode As String
    Dim aryFld As Variant
    Dim i As Long, iPos As Long

    If cboTemplate.ListIndex < 0 Then Exit Sub

    sTemp = cboTemplate.List(cboTemplate.ListIndex)
    sRType = medGetP(sTemp, 4, vbTab)
    sTCode = medGetP(sTemp, 1, vbTab)

    With txtRst
        .TextRTF = objETest.GetTemplateRst(sRType, sTCode)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelProtected = False
        '.SelColor = &H404040
        .SelProtected = True
'        If dicRstFields.Exists(sRType & COL_DIV & sTCode) Then
'            Call dicRstFields.KeyChange(sRType & COL_DIV & sTCode)
'            aryFld = Split(dicRstFields.Fields("rstfields"), vbTab)
'            For i = LBound(aryFld) To UBound(aryFld)
'                iPos = .Find(aryFld(i), 0, , rtfWholeWord)
'                While (iPos < Len(.Text)) AND (iPos >= 0)
'                    .SelStart = iPos
'                    .SelLength = Len(aryFld(i))
'                    .SelProtected = False
'                    .SelColor = vbBlue
'                    .SelProtected = True
'                    iPos = .Find(aryFld(i), iPos + Len(aryFld(i)), , rtfWholeWord)
'                Wend
'            Next
'        End If
        .SelStart = 0
        .SelLength = 0
    End With
    
    blnTmpDisplay = True
    fraResult.Visible = False
    txtRst.SetFocus
    
End Sub

Private Sub cmdApply_Click()
    Dim strHeader   As String
    Dim strNewLine  As String
    Dim strSpHeader As String
    Dim strTmp      As String
    Dim strTitle    As String
    Dim ii          As Long
    Dim jj          As Long
    
    strHeader = "\rtf1\ansi\ansicpg949\deff0{\fonttbl{\f0\fnil\fcharset129 \'b1\'bc\'b8\'b2;}}" & _
                "\viewkind4\uc1\pard\b\lang1042\f0\fs18 "
    strNewLine = "\par"
    
    strSpHeader = "\rtf1\ansi\ansicpg949\deff0{\fonttbl{\f0\fnil\fcharset129 \'b5\'b8\'bf\'f2\'c3\'bc;}}" & _
                  "\viewkind4\uc1\pard\lang1042\f0\fs18 "
    
    With tblData
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 1 To .MaxCols
                .Col = jj:
                If (jj Mod 3) = 1 Then
                    If jj = 1 Then
                        .Value = Format(.Value, "!" & String(16, "@"))
                    Else
                        .Value = Format(.Value, String(16, "@"))
                    End If
                ElseIf (jj Mod 3) = 2 Then
                    .Value = Format(Trim(.Value), String(9, "@"))
                Else
                    .Value = Format(Trim(.Value), String(5, "@"))
                End If
            Next jj
        Next ii
    End With
    
    If Trim(lblTitle1.Caption) <> "" And tblData.DataRowCnt > 0 Then
        tblData.Row = 1: tblData.Row2 = tblData.MaxRows
        tblData.Col = 1: tblData.COL2 = tblData.MaxCols
        tblData.BlockMode = True
        strTmp = tblData.Clip
        tblData.BlockMode = False
        strTmp = Replace(strTmp, vbTab, Space(1))
        strTmp = Replace(strTmp, vbNewLine, Space(1) & strNewLine & Space(1))
    End If
    
    If lblTitle1.Caption <> "" Then
        strTitle = strHeader & lblTitle1.Caption & strNewLine & strNewLine
    Else
        strTitle = ""
    End If
                  
    txtRst.TextRTF = "{" & strHeader & Mid(strRtfHead, 2, Len(strRtfHead) - 4) & strNewLine & _
                    strTitle & strSpHeader & strTmp & strNewLine & Mid(strRtfEnd, 2, Len(strRtfEnd))
       
    fraTest.Visible = False
End Sub

Private Sub cmdBM_Click()
    Dim strTmp  As VbMsgBoxResult
    
    If Trim(txtRst.Text) <> "" Then
        strTmp = MsgBox("기존 데이타가 없어 질 수 있습니다!", vbCritical + vbOKCancel, Me.Caption)
    
        If strTmp = vbCancel Then Exit Sub
    
    End If
    Set objDesc = frmTmpBMReport
    objDesc.Left = 2800
    objDesc.Top = 1000
    
    objDesc.Show 1
End Sub

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
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
    
    SQL = "SELECT HIVYN,                                                         "
    SQL = SQL + "       VDRLYN,                                                        "
    SQL = SQL + "       HBSAGYN,                                                       "
    SQL = SQL + "       HCVYN,                                                         "
    SQL = SQL + "       VMOTHRYN,                                                      "
    SQL = SQL + "       HBCYN,                                                         "
    SQL = SQL + "       HAVYN,                                                         "
    SQL = SQL + "       PENICILN,                                                      "
    SQL = SQL + "       RADCONT,                                                       "
    SQL = SQL + "       MRSAYN,                                                        "
    SQL = SQL + "       VREYN,                                                         "
    SQL = SQL + "       AFBYN,                                                         "
    SQL = SQL + "       TBYN,                                                          "
    SQL = SQL + "       CDIFFIYN,                                                      "
    SQL = SQL + "       FUNGUSYN,                                                      "
    SQL = SQL + "       VMREMARK,                                                      "
    SQL = SQL + "       OTHERRMK,                                                      "
    SQL = SQL + "       DRUGALGY,                                                      "
    SQL = SQL + "       PATNO,                                                         "
    SQL = SQL + "       SEQ,                                                           "
    SQL = SQL + "       TO_CHAR(EDITDATE,'YYYYMMDD') AS EDITDATE,                      "
    SQL = SQL + "       EDITID,                                                        "
    SQL = SQL + "       FN_USERNAME_SELECT(EDITID) AS EDITNM                           "
    SQL = SQL + "  FROM MDCAUTNT                                                       "
    SQL = SQL + " WHERE PATNO = '" & Trim(lblPtId.Caption) & "'                                             "
    SQL = SQL + "   AND SEQ = (SELECT MAX(SEQ) FROM MDCAUTNT WHERE PATNO = '" & Trim(lblPtId.Caption) & "') "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open SQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            For iCnt = 0 To 14
                If .Fields(iCnt).Value = "Y" Then
                    Check1(iCnt).Value = 1
                Else
                    Check1(iCnt).Value = 0
                End If
            Next
            lblWDt.Caption = Format(.Fields("EDITDATE").Value & "", "####-##-##")
            lblWNm.Caption = .Fields("EDITNM").Value & ""
            txtVival.Text = .Fields("VMREMARK").Value & ""
            txtDrug.Text = .Fields("DRUGALGY").Value & ""
            RichText.Text = .Fields("OTHERRMK").Value & ""
            Frame2.Visible = True
        Else
            Frame2.Visible = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing
End Sub

Private Sub cmdClear_Click()

    On Error GoTo Err_Trap
    
    ClearForm

'    lstTest.Clear
    lblTestNm.Caption = "": fTestCd = ""
    lstLabNo.Clear
    lstLabNo.ListIndex = -1
    txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
    blnTmpDisplay = False
    If txtWorkArea.Enabled Then txtWorkArea.SetFocus

Err_Trap:
    Resume Next
End Sub

Private Sub cmdClose_Click()
    Dim strHeader   As String
    strHeader = "\rtf1\ansi\ansicpg949\deff0{\fonttbl{\f0\fnil\fcharset129 \'b1\'bc\'b8\'b2;}}" & _
                "\viewkind4\uc1\pard\lang1042\f0\fs18 "
                  
    txtRst.TextRTF = "{" & strHeader & Mid(strRtfHead, 2, Len(strRtfHead) - 4) & Mid(strRtfEnd, 2, Len(strRtfEnd))
    fraTest.Visible = False
End Sub

Private Sub cmdColor_Click()
    
    DlgSave.ShowColor
    txtRst.SelProtected = False
    txtRst.SelColor = DlgSave.Color
    txtRst.SelProtected = True
End Sub

Private Sub cmdEdit_Click()
    txtRst.SelStart = 0
    txtRst.SelLength = Len(txtRst.Text)
    txtRst.SelProtected = False
    txtRst.SelLength = 0
    'txtRst.SelColor = DCM_LightBlue
    txtRst.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm293SpecialTest = Nothing
End Sub

Private Sub cmdFont_Click()
    If txtRst.SelProtected = True Then Exit Sub
    DlgSave.Flags = cdlCFBoth
    DlgSave.ShowFont
    txtRst.SelBold = DlgSave.FontBold
    txtRst.SelFontName = DlgSave.FontName
    txtRst.SelFontSize = DlgSave.FontSize
    txtRst.SelItalic = DlgSave.FontItalic
    txtRst.SelStrikeThru = DlgSave.FontStrikethru
    txtRst.SelUnderline = DlgSave.FontUnderline
    txtRst.SelProtected = True
End Sub

Private Sub cmdOrderView_Click()
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

Private Sub cmdReport_Click()

    Dim strVfyDt As String
    Dim sAccDt As String
    
    sAccDt = txtAccDt.Text
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If
    strVfyDt = objETest.GetVfyDate(txtWorkArea.Text, sAccDt, txtAccSeq.Text)
    Call PrintReport(strVfyDt)
End Sub

Private Sub PrintReport(ByVal pVfyDt As String)
    Dim objReport As New clsBatchReport
    Dim strLastDt As String, strLastTm As String

    With objReport
        .PtId = lblPtId.Caption
        .ptnm = lblPtNm.Caption
        .PtSex = Trim(medGetP(lblPtSA.Caption, 1, "/"))
        .PtAge = Trim(medGetP(lblPtSA.Caption, 2, "/"))
        .VfyDt = pVfyDt
        .VfyNM = ObjSysInfo.EmpNm
        .ICD = lblAttr.Caption
                
        .Rouding = False       '회진레포트 여부
        .Reprint = True        '재발행 여부
        .BatchReprint = False
        .Special = True
                
        .Dept = lblDeptCd.Caption
        .DeptNm = GetDeptNm(lblDeptCd.Caption)
        .Ward = lblWard.Caption
        
        Call .SpecialReport(mWorkarea, mAccDt, mAccSeq, mTestCd, lblPtId.Caption, pVfyDt, pVfyDt, _
                            enTestDiv.TST_SpeTest, "", picESign, Nothing, strLastDt, strLastTm)
    End With
    Set objReport = Nothing
End Sub

Private Sub cmdReset_Click(Index As Integer)
    lblRstCd(Index).Caption = ""
End Sub

Private Sub cmdRstCd_Click(Index As Integer)

    Dim objSQL As New clsLISSqlETest
    Dim objHelp As New clsPopUpList
    
    With objHelp
        .Connection = DBConn
        .FormCaption = "결과코드"
        .ColumnHeaderText = "코드;코드명"
        .LoadPopUp objSQL.SqlGetSpeRstCode(fTestCd, fRstType) ', 6000, 6000
        
        lblRstCd(Index).Caption = medGetP(.SelectedString, 1, ";")
        
        If chkApply.Value = 1 Then  '결과텍스트에 반영
            txtRst.SelProtected = False
            txtRst.SelColor = DCM_Black
            txtRst.SelText = medGetP(.SelectedString, 2, ";")
        End If
        
        txtRst.SetFocus
        
    End With
    Set objHelp = Nothing
    Set objSQL = Nothing
    
End Sub

Private Sub cmdSave_Click()
    
    Dim dsDT As Recordset, sSysDate As String, sDate As String, sTime As String
    Dim sChkVal As String, sChkTxt As String
    Dim blnSave As Boolean

    If txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then
        MsgBox "Accession Number가 정확하지 않습니다. 확인후 처리 하세요"
        Exit Sub
    End If
    
    If fStatus >= enStsCd.StsCd_LIS_FinRst Then
        MsgBox "이미 확인된 결과입니다. 일반 저장은 할 수 없습니다."
        Exit Sub
    End If

    ' 시스템 일자/시간 설정
    sSysDate = Format(GetSystemDate, "yyyymmdd hhmmss")
    sDate = Mid$(sSysDate, 1, 8)
    sTime = Mid$(sSysDate, 10, 6)

    If txtRst.Text = "" Then
        sChkTxt = "0"
    Else
        sChkTxt = ERT_TxtRst
    End If

On Error GoTo DBExecError

    DBConn.BeginTrans

    blnSave = SaveStatus(enStsCd.StsCd_LIS_MidRst, fMfySeq, sChkVal, sChkTxt, sDate, sTime)
    If Not blnSave Then GoTo DBExecError
    blnSave = SaveTxtRst(fMfySeq)             ' 안그러면 데이타 저장후 재 저장시 언매치 가능성
    If Not blnSave Then GoTo DBExecError
    blnSave = SaveFootnote(enStsCd.StsCd_LIS_MidRst)
    If Not blnSave Then GoTo DBExecError
    
    DBConn.CommitTrans

    'Call cmdClear_Click
    Call ClearForm

    If lstLabNo.ListCount > lstLabNo.ListIndex + 1 Then
        lstLabNo.ListIndex = lstLabNo.ListIndex + 1
        Call lstLabNo_KeyDown(vbKeyReturn, 0)
    End If
    
    Exit Sub

DBExecError:

    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Sub

Private Sub cmdSMS_Click()
    Dim SSQL As String
    
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
    
    frmSMS.Visible = True
    txtTransId.Text = Trim(ObjSysInfo.EmpId)
    txtTransNm.Text = GetEmpNm(Trim(ObjSysInfo.EmpId))
    txtTransNo.Text = txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccSeq.Text
    txtDtNo.Text = ""
    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:DD")
    txtDeptNm.Text = lblDeptCd.Caption
    
    rtfMessage.Text = "환자명 : " & lblPtNm.Caption & "(" & lblPtId.Caption & ")" & vbCRLF
    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value 즉시처치요함" & vbCr ' & rtfComment.Text
    If txtDtId.Text <> "" Then
        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
'
'        Set AdoCn_ORACLE = Nothing
    End If
    
    If txtExDtId.Text <> "" Then
        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtExDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtExDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtExDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
        
        Set AdoCn_ORACLE = Nothing
    End If

'
'    Dim SSQL As String
'    frmSMS.Visible = True
'
'    Set AdoCn_ORACLE = New ADODB.Connection
'
'    With AdoCn_ORACLE
'        .ConnectionTimeout = 25
''        .Provider = "OraOLEDB.Oracle.1"
'        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
'        .Properties("Data Source").Value = "PMC"
''        .Properties("Initial Catalog").Value = DatabaseName
'        .Properties("Persist Security Info") = True
'
'        .Properties("User ID").Value = "oral1"
'        .Properties("Password").Value = "oral1"
'
''        Screen.MousePointer = vbHourglass
'        .Open
'    End With
'
'    frmSMS.Visible = True
'    txtTransId.Text = Trim(ObjSysInfo.EmpId)
'    txtTransNm.Text = GetEmpNm(Trim(ObjSysInfo.EmpId))
'    txtTransNo.Text = txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccSeq.Text
'    txtDtNo.Text = ""
'    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:DD")
'    txtDtNm.Text = Trim(lblOrdDoct.Caption)
'
'    rtfMessage.Text = "환자명 : " & Trim(lblPtNm.Caption) & "(" & Trim(lblPtId.Caption) & ")" & vbCRLF & "Critical value 즉시처치요함" & vbCr ' & rtfComment.Text
'
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
''        SSQL = ""
''        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
''        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
'    Set AdoRs_ORACLE = New ADODB.Recordset
'
'    AdoRs_ORACLE.CursorLocation = adUseClient
'    AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
'
'    If AdoRs_ORACLE.RecordCount > 0 Then
'        txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
'    End If
'
'    Set AdoCn_ORACLE = Nothing
End Sub

Private Sub cmdTemplet_Click()
    Dim strTmp  As VbMsgBoxResult
    
    If Trim(txtRst.Text) <> "" Then
        strTmp = MsgBox("기존 데이타가 없어 질 수 있습니다!", vbCritical + vbOKCancel, Me.Caption)
    
        If strTmp = vbCancel Then Exit Sub
    
    End If
    Set objTemplet = frmTmpResult
    Set objTemplet.objcbo = cboTemplate
    Call objTemplet.LoadData(fTestCd)
    objTemplet.Left = 3000
    objTemplet.Top = 400
    
    objTemplet.Show vbModal
End Sub

Private Sub cmdTemplet2_Click()
    Dim strTmp  As VbMsgBoxResult
    
    If Trim(txtRst.Text) <> "" Then
        strTmp = MsgBox("기존 데이타가 없어 질 수 있습니다!", vbCritical + vbOKCancel, Me.Caption)
    
        If strTmp = vbCancel Then Exit Sub
    
    End If
    Set objTemplet2 = frmTmpEIReport
    'Set objTemplet2.objcbo = cboTemplate
    'Call objTemplet2.LoadData(fTestCd)
    objTemplet2.Left = 3000
    objTemplet2.Top = 400
    
    objTemplet2.Show vbModal
End Sub

Private Sub cmdTrans_Click()
    Dim ServerName   As String
    Dim DatabaseName As String
    Dim UserName     As String
    Dim Password     As String
    Dim strTransCd   As String
    Dim strDoctCd    As String
    Dim strTransDt   As String
    Dim strTransStatus As String
    Dim strTansEtc   As String
    Dim strMessage   As String
    Dim strTransNo   As String
    Dim strDoctNo    As String
    Dim strSQL       As String
    Dim strDeptNm    As String
    Dim strTranNm    As String
    Dim strSMSIP     As String
    Dim strBackNo    As String
    Dim strTmpTestCd As String
    Dim strMaDtId  As String
    Dim strMaTransNo As String
    
    Set AdoCn_ORACLE = New ADODB.Connection
    
    On Error Resume Next    '2013-09-11 PSK
    
    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
        .Properties("Persist Security Info") = True
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        .Open
    End With
           
    Set AdoRs_ORACLE = New ADODB.Recordset
        
    strSQL = ""
    strSQL = "SELECT * FROM S2lab032  "
    strSQL = strSQL + " WHERE cdindex = 'C232'"
    strSQL = strSQL + "   AND cdval1 = 'SVR1'  "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open strSQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            strSMSIP = AdoRs_ORACLE.Fields("FIELD4") & ""
        Else
            strSMSIP = "172.16.200.37"
        End If
        .Close
    End With
    
    Set AdoCn_SQL = New ADODB.Connection

    ServerName = strSMSIP
    DatabaseName = "medicalCRM_jesus"
    UserName = "jesus"
    Password = "jesus"
   
    With AdoCn_SQL
        .ConnectionTimeout = 10
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("User ID").Value = UserName
        .Properties("Password").Value = Password
        Screen.MousePointer = vbHourglass
        .Open
    End With
    Screen.MousePointer = vbDefault
    
'    If txtDtNo.Text = "" Then
'        MsgBox "수신번호를 입력하세요.", vbCritical + vbOKOnly, "수신번호등록 Message"
'        txtDtNo.SetFocus
'        Exit Sub
'    End If
    
    strTransCd = ObjSysInfo.EmpId
    strTransNo = txtTransNo.Text
    strDoctCd = txtDtId.Text
    strMaDtId = txtExDtId.Text
    strMaTransNo = txtExDtNo.Text
    strTransDt = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strDoctNo = txtDtNo.Text
    strTransStatus = "1"
    strTansEtc = "LIS"
    strDeptNm = txtDeptNm.Text
    strTranNm = txtTransNm.Text
    strMessage = rtfMessage.Text & vbCRLF & "- " & strTranNm
    strBackNo = "063-230-8753"
    strTmpTestCd = txtTestCd.Text
    
    If Len(strMessage) > 80 Then
        MsgBox "메시지의 크기를 줄여주세요.", vbCritical + vbOKOnly, "메시지내용수정 Message"
        rtfMessage.SetFocus
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
    strSQL = strSQL & " values('" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & strBackNo & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransStatus & "' ,"
    strSQL = strSQL & "        '" & strTansEtc & "')"
    
    AdoCn_SQL.Execute strSQL
    
    ' 검사코드 추가
    ' 2019-05-03 SMS 조회 검사코드로 조회 용
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
    strSQL = strSQL & " values('" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtId.Text) & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtNm.Text) & "' ,"
    strSQL = strSQL & "        '" & strDeptNm & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '정상' ,"
    strSQL = strSQL & "        '" & strTransNo & "',"
    strSQL = strSQL & "        '" & strRcvDt & "',"
    strSQL = strSQL & "        '" & strTmpTestCd & "')"
    
    AdoCn_ORACLE.Execute strSQL
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
    strSQL = strSQL & "        '7' ,"
    strSQL = strSQL & "        SYSDATE ,"
    strSQL = strSQL & "        '[CVR(이상결과보고)]' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
    
    AdoCn_ORACLE.Execute strSQL
    
    If Trim(txtDtId.Text) <> Trim(txtExDtId.Text) Then
        strSQL = ""
        strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
        strSQL = strSQL & " values('" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & strBackNo & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransStatus & "' ,"
        strSQL = strSQL & "        '" & strTansEtc & "')"
        
        AdoCn_SQL.Execute strSQL
        
        ' 검사코드 추가
        ' 2019-05-03 SMS 조회 검사코드로 조회 용
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
        strSQL = strSQL & " values('" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtId.Text) & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtNm.Text) & "' ,"
        strSQL = strSQL & "        '" & strDeptNm & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '정상' ,"
        strSQL = strSQL & "        '" & strTransNo & "',"
        strSQL = strSQL & "        '" & strRcvDt & "',"
        strSQL = strSQL & "        '" & strTmpTestCd & "')"
        
        AdoCn_ORACLE.Execute strSQL
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
        strSQL = strSQL & " (select '" & strMaDtId & "' ,"
        strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
        strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
        strSQL = strSQL & "        '7' ,"
        strSQL = strSQL & "        SYSDATE ,"
        strSQL = strSQL & "        '[CVR(이상결과보고)]' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strMaDtId & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
        
        AdoCn_ORACLE.Execute strSQL
    End If
    
    strRcvDt = ""
    
    frmSMS.Visible = False
    Set AdoCn_SQL = Nothing
    Set AdoCn_ORACLE = Nothing
    
End Sub

Private Sub cmdVerify_Click()
    
    Dim sSysDate As String, sDate As String, sTime As String
    Dim sStatus As String, iMfyCnt As Integer
    Dim blnSave As Boolean
    Dim tmpDept As String
    Dim tmpBussDiv As String
    Dim lngResp As VbMsgBoxResult
    
    '-- 전주예수병원 추가 변수=============
    Dim clsTransfer As New clsLISTransfer
    '======================================
    
'   lstLabNo.Enabled = False

    If txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then
        MsgBox "Accession Number가 정확하지 않습니다. 확인후 처리 하세요", vbInformation, "결과입력"
        Exit Sub
    End If

    If txtRst.Text = "" Then
        MsgBox "Text 결과를 입력하지 않았습니다. 확인 후 처리하세요", vbInformation, "결과입력"
        Exit Sub
    End If

    Dim objESign        As clsLISElectronSign
    Dim strTmp          As VbMsgBoxResult
    Dim strWorkArea     As String
    Dim strAccDt        As String
    Dim strAccSeq       As String
    Dim strKey          As String
    Dim strData         As String
    Dim strRstEntryType As String
    Dim i               As Long
    Dim j               As Long
    
    '-------------------------------------------------------------------------------------------
    '전자서명 Validation Check
    If P_ElectronicSignature Then
        Set objESign = New clsLISElectronSign
        If objESign.LoadElectronSign(ObjMyUser.EmpId, InstallDir & "LIS\Bin") = False Then
            '전자서명 인증 에러
            medBeep 20
            MsgBox objESign.ErrMsg, vbCritical, "전자서명 확인"
            Exit Sub
        Else
            '전자서명 인증
            objESign.ShowESignForm
            If objESign.ElectronSingOk = True Then
            Else
                MsgBox "전자서명 인증을 하지 않으셨습니다.", vbInformation, "전자서명 인증"
                Exit Sub
            End If
        End If
    End If
    '-------------------------------------------------------------------------------------------
    
    ' 시스템 일자/시간 설정
    sSysDate = Format(GetSystemDate, "yyyymmdd hhmmss")
    sDate = Mid$(sSysDate, 1, 8)
    sTime = Mid$(sSysDate, 10, 6)

    ' 적용 Status 설정
    If fStatus < enStsCd.StsCd_LIS_FinRst Then sStatus = enStsCd.StsCd_LIS_FinRst: iMfyCnt = fMfySeq
    If fStatus >= enStsCd.StsCd_LIS_FinRst Then sStatus = enStsCd.StsCd_LIS_Modify: iMfyCnt = fMfySeq + 1

On Error GoTo DBExecError

    DBConn.BeginTrans

    blnSave = SaveStatus(sStatus, iMfyCnt, ERT_ValRst, ERT_TxtRst, sDate, sTime)
    If Not blnSave Then GoTo DBExecError
    blnSave = SaveTxtRst(iMfyCnt)                     ' 결과 내역 등록
    If Not blnSave Then GoTo DBExecError
    
    If Trim(lblWardId.Caption) = "" Then
        tmpDept = lblDeptCd.Caption
        tmpBussDiv = enBussDiv.BussDiv_OutPatient
    Else
        tmpDept = lblWardId.Caption
        tmpBussDiv = enBussDiv.BussDiv_InPatient
    End If
    blnSave = objETest.SubmitVerifyList(tmpDept, sDate, sTime, lblPtId.Caption, sStatus, ObjMyUser.EmpId, lblMajDoct.Caption, tmpBussDiv)
    If Not blnSave Then GoTo DBExecError
    blnSave = SaveAccStatus(sStatus, sDate, sTime)        ' 최종적으로 접수내역에 Status 반영
    Call SaveFootnote(sStatus)
    
    '** 전주예수병원 추가 루틴 =============================================================
    Dim strMfyFg As String
    
    strWorkArea = Trim(txtWorkArea.Text)
    strAccDt = Trim(txtAccDt.Text)
    strAccSeq = Trim(txtAccSeq.Text)
    
    If iMfyCnt = 0 Then
        strMfyFg = "0"
    Else
        strMfyFg = "1"
    End If
    
    '** 일단 OCS 결과전송 시 일반 텍스트 형식으로 넘기기로 함. (2004.12.06)
    If clsTransfer.SpecialTest_Main(strWorkArea, strAccDt, strAccSeq, fTestCd, _
        sDate, sTime, ObjMyUser.EmpId, txtRst.Text, "", strMfyFg) = False Then
                                    
        MsgBox "결과전송 시 오류가 발생하였습니다.", vbCritical
        GoTo DBExecError
        
    End If
    
    '=======================================================================================
    
    DBConn.CommitTrans
    
    '-- 예수병원 일단 보류 함 =============================================================
'    lngResp = MsgBox("결과지를 출력하시겠습니까?", vbQuestion + vbYesNo, "결과지 출력")
'    If lngResp = vbYes Then
'        Call PrintReport(Format(Now, CS_DateDbFormat))
'    End If
    '======================================================================================
    
    'Call cmdClear_Click
    Call ClearForm
    
On Error GoTo Err_Trap
    
    If optInput(0).Value Then
        If txtWorkArea.Enabled Then txtWorkArea.SetFocus
    End If
    'If optInput(1).Value AND lstLabNo.ListIndex > -1 Then lstLabNo.RemoveItem lstLabNo.ListIndex
    'If optInput(1).Value AND lstTstNo.ListIndex > -1 Then Call lstTstNo_KeyDown(13, 0)
    If lstLabNo.ListCount > lstLabNo.ListIndex + 1 Then
        lstLabNo.ListIndex = lstLabNo.ListIndex + 1
        Call lstLabNo_KeyDown(vbKeyReturn, 0)
        
        lstLabNo.RemoveItem lstLabNo.ListIndex - 1
    Else
        lstLabNo.RemoveItem lstLabNo.ListIndex
        If txtWorkArea.Enabled Then txtWorkArea.SetFocus
    End If

'   lstLabNo.Enabled = True
    
    Set clsTransfer = Nothing
    
    Exit Sub

DBExecError:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
    Set clsTransfer = Nothing
    Exit Sub
Err_Trap:
    Resume Next
End Sub

Private Function SaveStatus(ByVal pStatus As String, ByVal pMfyCnt As Integer, ByVal pValRst As String, _
                            ByVal pTxtRst As String, ByVal pDate As String, pTime As String) As Boolean
    
    Dim sqlUpRst As String, sqlUpOrd As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim blnSave As Boolean

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

'*1    '기타검사 결과 내역 Update (주의 : 같은 번호에 다수개의 검사가 존재 할 수 있슴)
    If pStatus = enStsCd.StsCd_LIS_Modify And pMfyCnt > 0 Then
         blnSave = objETest.SetHistory(sWorkArea, sAccDt, sAccSeq, fTestCd) '(pMfyCnt)       ' History 작업
         If Not blnSave Then GoTo Err_Trap
    End If

    ' 등록과 수정 모두에 공통 적용
    blnSave = objETest.SaveStatus(sWorkArea, sAccDt, sAccSeq, fTestCd, pStatus, _
                                  pMfyCnt, pValRst, pTxtRst, pDate, pTime, ObjMyUser.EmpId)
    If Not blnSave Then GoTo Err_Trap
    
    SaveStatus = True
    Exit Function

Err_Trap:
    SaveStatus = False
    
End Function

Private Function SaveTxtRst(ByVal pMfySeq As Integer) As Boolean
    
    Dim sInsHead As String, sInsData As String
    Dim sqlDelete As String, sqlInsert As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim blnSave As Boolean

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

    '2001-12-27 수정
    'SaveTxtRst = objETest.SaveValResult(sWorkArea, sAccDt, sAccSeq, fTestCd, pMfySeq, txtRst.TextRTF)
    SaveTxtRst = objETest.SaveSpeResult(sWorkArea, sAccDt, sAccSeq, fTestCd, pMfySeq, txtRst.TextRTF, _
                                        lblRstCd(0).Caption, lblRstCd(1).Caption, lblRstCd(2).Caption)

End Function

Private Function SaveAccStatus(ByVal pStatus As String, ByVal pDate As String, ByVal pTime As String) As Boolean
    
    Dim sqlAcc As String
    Dim iTotalCount As Integer, iInputCount As Integer, sStatus As String
    Dim upStatus As String, upInputCount As Integer
    Dim sqlUpdAcc As String

    ' 중간결과를 사용하지 않는 검사인 경우에 성립
    ' 만약 감수성 검사 등의 경우라면 중간결과시 카운터에는 반영않지만 Status는 반영함
    If pStatus < enStsCd.StsCd_LIS_FinRst Then Exit Function

    Dim sWorkArea As String, sAccDt As String, sAccSeq As String

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

    ' 검사종류 확인
    SaveAccStatus = objETest.SaveAccStatus(sWorkArea, sAccDt, sAccSeq, pStatus, pDate, pTime, ObjMyUser.EmpId)
    
End Function

Private Function SaveFootnote(ByVal pStatus As String) As Boolean
    
    Dim sRemarkCd As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If
    sRemarkCd = cboRemark.List(cboRemark.ListIndex)

    SaveFootnote = objETest.SaveFootnote(sWorkArea, sAccDt, sAccSeq, ObjSysInfo.EmpId, _
                                         txtFNote.Text, sRemarkCd, pStatus, fFNSeq)

End Function

Private Sub Command1_Click()
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtVival.Text = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame2.Visible = False
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    If blnInitFg Then Exit Sub
    medInitLvwHead lvwResult, "결과 Temlate,코드", "7050,0"
    blnInitFg = True
End Sub

Private Sub Form_Load()

    txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
    
    If P_ImageSystem = False Then sstRst.TabVisible(3) = False
'    sstRst.TabVisible(3) = False
    Call objETest.LoadRemark(cboRemark)
    
'    slider.Max = txtCRst.Top + txtCRst.Height + cSep + txtRst.Height
'    slider.Min = txtCRst.Top

    ' 결과 확인/수정은 과장급만 사용 가능
    'If objMyUser.Degree = "0" Then
    If ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper Then
         cmdVerify.Enabled = True
         lstTstNo.Enabled = True: lstLabNo.Enabled = True
         optInput(0).Value = True
         Call objETest.LoadTstNo(lstTstNo)
    Else
         
         cmdVerify.Enabled = False
         lstTstNo.Enabled = False: lstLabNo.Enabled = False
         optInput(0).Value = True: optInput(1).Enabled = False
    End If
    
    blnInitFg = False
    ClearForm
    Call objETest.LoadRstFields(dicRstFields)
    
    cmdTemplet.Visible = True
    cmdTemplet2.Visible = True
    
    frmSMS.Visible = False
    
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtVival.Text = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame2.Visible = False
End Sub

Private Sub ClearForm()
    fFNSeq = 0

    lblLastDate.Caption = ""
    
    fMfySeq = 0
    fLWorkArea = "": fLAccDt = "": fLAccSeq = 0: fLM = 0
    lblMode.Caption = ""
    cmdSave.Enabled = True
    cmdVerify.Caption = "확 인 (&V)"
    cmdBM.Enabled = False

    lblMode.ForeColor = cGenMode
    lblMode.Visible = False

    lblPtId.Caption = "":    lblPtNm.Caption = "": lblPtSA.Caption = ""
    lblDeptCd.Caption = "":  lblWard.Caption = "": lblSpecimen.Caption = ""
    lblWardId.Caption = "":  lblAttr.Caption = ""
    lblOrdDoct.Caption = ""
    lblDoctID.Caption = ""
    txtPNote.Text = "":      txtFNote.Text = ""
    cboRemark.ListIndex = 0: lblRemark.Caption = ""
    
    lblRstCd(0).Caption = ""
    lblRstCd(1).Caption = ""
    lblRstCd(2).Caption = ""
    chkApply.Value = 1
    
    sstRst.Tab = 0
    cboTemplate.Clear
    slider.Value = slider.Max / 8
    txtCRst.Text = "": txtRst.Text = "": txtLastRst.Text = ""
    fraRst.Enabled = False

    fraResult.Visible = False
    
    lvwHxList.ListItems.Clear
    lvwList.ListItems.Clear
    imgHx.Visible = False
    imgList.Visible = False
    
    txtOCSMesg.Text = ""
    
    mWorkarea = ""
    mAccDt = ""
    mAccSeq = ""
    mTestCd = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
'    Set mnuPopup = Nothing
'    Set mnuSave = Nothing
    If P_SLIDE_SERVER_PATH = "" Then ClearImage
    Set frm293SpecialTest = Nothing
End Sub

Private Sub imgHx_DblClick()
    If imgHx.Visible = True Then
        strImagePath = lvwHxList.SelectedItem.SubItems(6)
        imgImage.Picture = imgHx.Picture
        frmImage.Visible = True
    Else
        strImagePath = ""
    End If
End Sub

Private Sub imgImage_DblClick()
    frmImage.Visible = False
End Sub

Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If strImagePath = "" Then Exit Sub
        Set objPop = Nothing
        Set objPop = New clsPopupMenu
        
        With objPop
            .AddMenu MENU_SAVE, "IMAGE SAVE"
            
            .PopupMenus Me.hwnd
        End With
        
        Set objPop = Nothing
        
'        Set mnuPopup = frmControls.mnuPopup
'        Set mnuSave = frmControls.mnuSub
'        frmControls.mnuSub1.Visible = False
'        frmControls.mnuSub2.Visible = False
'        mnuSave.Caption = "Image Save"
'        Me.PopupMenu mnuPopup
'
'        Set mnuPopup = Nothing
'        Set mnuSave = Nothing
'        Unload frmControls
'        Set frmControls = Nothing
    End If
End Sub

'Private Sub mnuSave_Click()
'    Dim strImgDir As String
'
'    If Trim(strImagePath) = "" Then Exit Sub
'
'    strImgDir = strImagePath
'
'    DlgSave.InitDir = "C:\"
'    DlgSave.Filter = "JPEG"
'    DlgSave.FileName = Mid(strImgDir, InStrRev(strImgDir, "\", , vbTextCompare) + 1, Len(strImgDir))
'    DlgSave.ShowSave
'
'    FileCopy strImgDir, DlgSave.FileName
'End Sub

Private Sub imgList_Click()
    If imgList.Visible = True Then
        strImagePath = lvwList.SelectedItem.SubItems(6)
        imgImage.Picture = imgList.Picture
        frmImage.Visible = True
    Else
        strImagePath = ""
    End If
End Sub

Private Sub lstLabNo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim sTmp As String, sTmp2 As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim iTestCount As Integer
    Dim sTestcd As String, sTestNm As String, sRstFg As String, sRstType As String, iMfySeq As Long

    If lstTstNo.ListIndex < 0 Or lstLabNo.ListIndex < 0 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        
        sTmp = medGetP(lstLabNo.List(lstLabNo.ListIndex), 1, " ") 'medGetP(lstLabNo.List(lstLabNo.ListIndex), 1, vbTab)
        txtWorkArea = medGetP(sTmp, 1, "-"): sWorkArea = txtWorkArea
        txtAccDt = medGetP(sTmp, 2, "-"): sAccDt = IIf(Mid$(txtAccDt, 1, 1) = "9", "19", "20") & txtAccDt
        txtAccSeq = medGetP(sTmp, 3, "-"): sAccSeq = medGetP(txtAccSeq, 1, " ")
        
        '감염관리
        Call ICSLabNoMark(sWorkArea, sAccDt, sAccSeq, enICSNum.LIS_ALL)
    
        sTmp2 = lstTstNo.List(lstTstNo.ListIndex)
        sTestcd = medGetP(sTmp2, 1, vbTab)
        sTestNm = Trim(medGetP(sTmp2, 2, vbTab))
        sRstType = medGetP(sTmp2, 3, vbTab)
    
        Call objETest.LoadResultByLabNo(sTestcd, sWorkArea, sAccDt, sAccSeq, sRstFg, iMfySeq)
        
        lstTest.Clear
        lstTest.AddItem sTestcd & vbTab & sRstFg & vbTab & sTestNm & _
                        vbTab & vbTab & sRstType & vbTab & iMfySeq
        lstTest.ListIndex = 0
        Call lstTest_KeyPress(vbKeyReturn)
                  
    End If

End Sub

Private Sub lstLabNo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If Button = vbLeftButton Then
        Call lstLabNo_KeyDown(13, 0)
    End If

End Sub

Private Sub lvwLResult_Click()
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strTestCd   As String
    Dim strMfySeq   As String
    Dim strVfyDt    As String
    Dim strVfyTm    As String
    
    strWorkArea = lvwLResult.SelectedItem.SubItems(6)
    strAccDt = lvwLResult.SelectedItem.SubItems(7)
    strAccSeq = lvwLResult.SelectedItem.SubItems(8)
    strVfyDt = lvwLResult.SelectedItem.SubItems(1)
    strVfyTm = lvwLResult.SelectedItem.SubItems(2)
    strTestCd = lvwLResult.SelectedItem.SubItems(9)
    strMfySeq = lvwLResult.SelectedItem.SubItems(10)
    
    
    sstRst.Tab = 1: sstRst.Caption = " 최근결과 " & "(" & _
                                        Format(Mid(strVfyDt, 3, 6), "00/00/00") & " " & _
                                        Format(Mid(strVfyTm, 1, 4), "00:00") & ")"
    
    txtLastRst.TextRTF = objETest.GetResultText(strWorkArea, strAccDt, strAccSeq, strTestCd, strMfySeq)
End Sub

Private Sub lstTest_KeyPress(KeyAscii As Integer)
    
    Dim sTmp As String, sRstType As String
    
    On Error GoTo Err_Trap
    
    Select Case True
        Case KeyAscii = 27
            Call ClearForm
            If optInput(0).Value = True Then
                If txtWorkArea.Enabled Then txtWorkArea.SetFocus
            End If
            lstTest.Visible = False
        
        Case KeyAscii = vbKeyReturn And lstTest.ListIndex >= 0
            sTmp = lstTest.List(lstTest.ListIndex)
            fraRst.Enabled = True
            fTestCd = "" & medGetP(sTmp, 1, vbTab)
            fStatus = "" & medGetP(sTmp, 2, vbTab)
            lblTestNm.Caption = "" & medGetP(sTmp, 3, vbTab)
            sRstType = "" & medGetP(sTmp, 5, vbTab)
            fMfySeq = Val("" & medGetP(sTmp, 6, vbTab))
            fRstType = sRstType
            Call SetEditMode
            Call DisplayForm(fTestCd, sRstType)
            
            '임상진단....
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.PtId = lblPtId.Caption
            lblAttr.Caption = objDisease.Disease
            Set objDisease = Nothing
                
            lstTest.Visible = False
            
            txtRst.SetFocus
    End Select

Err_Trap:
    Resume Next
    
End Sub

Private Sub lstTest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        Call lstTest_KeyPress(13)
    ElseIf Button = vbRightButton Then
        Call lstTest_KeyPress(27)
    End If
    
End Sub

Private Sub SetEditMode()

    If fStatus < enStsCd.StsCd_LIS_FinRst Then
        lblMode.Caption = "결과 등록"
        lblMode.ForeColor = cInsMode
        cmdSave.Enabled = True
        cmdVerify.Caption = "확 인 (&V)"
    Else
        lblMode.Caption = "결과 수정"
        lblMode.ForeColor = cMfyMode
        cmdSave.Enabled = False
        cmdVerify.Caption = "수 정 (&V)"
    End If
    
    lblMode.Visible = True
    cmdBM.Enabled = True
End Sub

Private Sub DisplayForm(ByVal pTestCd As String, ByVal pRstType As String)
    
    Dim i As Integer, j As Integer
    Dim sqlBasic As String, dsBasic As New Recordset, iBasicCol As Integer
    Dim sSTS As String, sRstChk As String
    Dim sRemarkCd As String, sRemarkIdx As Integer
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim objSQL As New clsLISSqlETest
    
    Dim objPatient As New clsPatient     '환자 클래스
    
    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

    sqlBasic = objSQL.SqlGetDataByLabNo(pTestCd, sWorkArea, sAccDt, sAccSeq)
               
    dsBasic.Open sqlBasic, DBConn
    

    If dsBasic.EOF = False Then
        mWorkarea = sWorkArea
        mAccDt = sAccDt
        mAccSeq = sAccSeq
        mTestCd = pTestCd
        
        lblPtId.Caption = "" & dsBasic.Fields("ptid").Value
        lblPtNm.Caption = "" & dsBasic.Fields("ptnm").Value
        If "" & dsBasic.Fields("ageday").Value < 365 Then
            lblPtSA.Caption = "" & dsBasic.Fields("sex").Value & "/" & Str$("" & dsBasic.Fields("ageday").Value) & " D"
        Else
'            lblPtSA.Caption = "" & dsBasic.Fields("sex").Value & "/" & Str$((Val("" & dsBasic.Fields("ageday").Value) \ 365) + 1)
            With objPatient
                If Trim(lblPtId.Caption) <> "" And .GETPatient(lblPtId.Caption) Then
    '                lblPtNm.Caption = .ptnm
    '                lblSex.Caption = .SEXNM
    '                lblAge.Caption = .Age
                    lblPtSA.Caption = "" & dsBasic.Fields("sex").Value & "/" & .Age
                End If
            End With
        End If
        lblDeptCd.Caption = "" & dsBasic.Fields("deptcd").Value
        
        '-- 병동정보 가져오기
        Call objPatient.GETPatient(lblPtId.Caption)
        lblWard.Caption = objPatient.WardId & "-" & objPatient.RoomId
        '-- 원본 -------------------------------
'        lblWard.Caption = "" & dsBasic.Fields("wardid").Value & "-" & dsBasic.Fields("roomid").Value & "-" & dsBasic.Fields("bedid").Value
        '---------------------------------------
        
        lblOrdDoct.Caption = GetLisDoctNm("" & dsBasic.Fields("orddoct").Value)
        lblDoctID.Caption = "" & dsBasic.Fields("orddoct").Value
        lblWardId.Caption = "" & dsBasic.Fields("wardid").Value
        lblSpecimen.Caption = "" & dsBasic.Fields("spcnm").Value
    
        sSTS = "" & dsBasic.Fields("stscd").Value
        sRstChk = "" & dsBasic.Fields("valfg").Value
        sRemarkCd = "" & dsBasic.Fields("rmkcd").Value
        
        fFNSeq = Val("" & dsBasic.Fields("footnotefg").Value)
    
        txtDtId.Text = "" & dsBasic.Fields("orddoct").Value
        txtExDtId.Text = "" & dsBasic.Fields("majdoct").Value
        rtfMessage.Text = ""
        strRcvDt = "" & dsBasic.Fields("rcvdt").Value
        txtTestCd.Text = "" & dsBasic.Fields("testcd").Value
    Else
        MsgBox "데이타에 이상이 있습니다. 전산실 혹은 임상병리과로 연락 바랍니다(☎" & ObjSysInfo.HelpLine & ").value", _
                vbCritical, "데이터오류"
        ClearForm
        Set dsBasic = Nothing
        Exit Sub
    End If
    
    Set dsBasic = Nothing

    ' Foot Note 내역
    Call DispFootNote(sWorkArea, sAccDt, sAccSeq, sSTS)

    ' 검체 Remark Display
    cboRemark.ListIndex = medComboFind(cboRemark, sRemarkCd)

    ' Text 결과 Template 세팅
    Call objETest.SetTemplate(pRstType, cboTemplate)
    'Call objETest.SetAppend(pTestCd, cboAppend)
    Call objETest.LoadAppendTemp(pRstType, cboAppend)
    
    ' 과거 결과 유뮤 및 최근 결과 Lab-No Setting
    Call GetLastAccNo(lblPtId.Caption)

    Call DispTxtResult
    
    '** 예수병원 추가루틴 By M.G.Choi 2004.12.14
    ' * 처방Remark Display ============================================
    txtOCSMesg.Text = OCSRemark_Info(sWorkArea, sAccDt, sAccSeq, pTestCd)
    '==================================================================
    
    '이미지 리스트 Loading
    If P_ImageSystem = True Then Call LoadImage(Trim(lblPtId.Caption), sWorkArea, sAccDt, sAccSeq)
    
    Set objSQL = Nothing
    Set objPatient = Nothing
    
End Sub

'** 예수병원 추가루틴 By M.G.Choi 2004.12.14
' * 처방Remark Display
Private Function OCSRemark_Info(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                            ByVal pTestCd) As String
    Dim RS      As New ADODB.Recordset
    Dim strSQL  As String
    
    strSQL = " select * from " & T_LAB102 & _
             "  where workarea = " & DBS(pWorkArea) & _
             "    and accdt    = " & DBS(pAccDt) & _
             "    and accseq   = " & DBN(pAccSeq) & _
             "    and ordcd    = " & DBS(pTestCd)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        OCSRemark_Info = RS.Fields("mesg").Value & ""
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function
'=============================================================================================

Private Sub LoadImage(ByVal pPtId As String, ByVal pWorkArea As String, _
                      ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim cn As ADODB.Connection, RS As ADODB.Recordset, SQL As String
    Dim Cnt As Long
    Dim ii As Long
    Dim jj As Long
    Dim iTmx        As ListItem
    Dim itmx1       As ListItem
    
    
    Set cn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    cn.CursorLocation = adUseServer
    cn.Open "Driver={Microsoft ODBC for Oracle};" & _
    "Server=" & GetSetting("Schweitzer2000 LIS", "Server", "DB", "") & ";" & _
    "Uid=" & GetSetting("Schweitzer2000 LIS", "Server", "UID", "") & ";" & _
    "Pwd=" & GetSetting("Schweitzer2000 LIS", "Server", "PWD", "") & ";"
    
    Dim strFrDt As String
    Dim strToDt As String
    
    strFrDt = Format(DateAdd("YYYY", -1, GetSystemDate), "YYYYMMDD")
    strToDt = Format(GetSystemDate, "YYYYMMDD")
    
'    For ii = 1 To Cnt
        SQL = " SELECT * FROM " & T_LAB310 & _
              "  WHERE " & DBW("ptid", pPtId, 2) & _
              "    AND " & DBW("rcvdt>=", strFrDt) & _
              "    AND " & DBW("rcvdt<=", strToDt) & _
              "  ORDER BY workarea,accdt,accseq,seq "
'              & _
              "    AND " & DBW("seq", "2", 2)
        RS.Open SQL, cn, adOpenStatic, adLockReadOnly
    
        ' Save using GetChunk AND known size.
        ' FieldSize (ActualSize) > Threshold arg (16384)
        If P_SLIDE_SERVER_PATH = "" Then ClearImage
        
        medInitLvwHead lvwHxList, _
             "No,Slide No,Status,File Size,File Date,Description,Directory", _
             "-50,950,300,650,1300,3000,3000"
        
        medInitLvwHead lvwList, _
             "No,Slide No,Status,File Size,File Date,Description,Directory", _
             "-50,950,300,650,1300,3000,3000"
        
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            ii = 0: jj = 0
            lvwList.ListItems.Clear
            lvwHxList.ListItems.Clear
            Do Until RS.EOF
            
                If P_SLIDE_SERVER_PATH = "" Then BlobToFile RS!imgfile, RS!imgdir
                
                If RS!WorkArea = pWorkArea And RS!AccDt = pAccDt And RS!AccSeq = pAccSeq Then
                    ii = ii + 1
                    With lvwList
                        Set iTmx = .ListItems.Add(, , ii)
                                
                        iTmx.SubItems(1) = RS!WorkArea & "-" & Mid(RS!AccDt, 3) & "-" & RS!AccSeq & "-" & Format(RS!SEQ, "00")
                        iTmx.SubItems(2) = "이미지" & COL_DIV & RS!PrtFg
                        iTmx.SubItems(3) = VBA.FileLen(RS!imgdir) & " byte"
                        iTmx.SubItems(4) = VBA.FileDateTime(RS!imgdir)
                        iTmx.SubItems(5) = RS!Rmk
                        iTmx.SubItems(6) = RS!imgdir
                    End With
                    
                Else
                    jj = jj + 1
                    With lvwHxList
                        Set itmx1 = .ListItems.Add(, , jj)
                                
                        itmx1.SubItems(1) = RS!WorkArea & "-" & Mid(RS!AccDt, 3) & "-" & RS!AccSeq & "-" & Format(RS!SEQ, "00")
                        itmx1.SubItems(2) = "이미지" & COL_DIV & RS!PrtFg
                        itmx1.SubItems(3) = VBA.FileLen(RS!imgdir) & " byte"
                        itmx1.SubItems(4) = VBA.FileDateTime(RS!imgdir)
                        itmx1.SubItems(5) = RS!Rmk
                        itmx1.SubItems(6) = RS!imgdir
                    End With
                    
                End If
                RS.MoveNext
            Loop
        End If
'
    RS.Close
    cn.Close
    
    If lvwList.ListItems.Count > 0 Then
        lvwList.ListItems(1).Selected = True
        imgList.Picture = LoadPicture(lvwList.SelectedItem.SubItems(6))
        imgList.Visible = True
    End If
    If lvwHxList.ListItems.Count > 0 Then
        lvwHxList.ListItems(1).Selected = True
        imgHx.Picture = LoadPicture(lvwHxList.SelectedItem.SubItems(6))
        imgHx.Visible = True
    End If
    
    Set iTmx = Nothing
    Set itmx1 = Nothing
    Set RS = Nothing
    Set cn = Nothing
End Sub

Private Sub ClearImage()
    Dim ii As Long
    
    If Dir(P_SLIDE_DB_PATH, vbDirectory) = "" Then
        MkDir P_SLIDE_DB_PATH
    Else
        If lvwList.ListItems.Count > 0 Then
            For ii = 1 To lvwList.ListItems.Count
                If Dir(Trim(lvwList.ListItems(ii).SubItems(6))) <> "" Then Kill Trim(lvwList.ListItems(ii).SubItems(6))
            Next
        End If
        
        If lvwHxList.ListItems.Count > 0 Then
            For ii = 1 To lvwHxList.ListItems.Count
                If Dir(Trim(lvwHxList.ListItems(ii).SubItems(6))) <> "" Then Kill Trim(lvwHxList.ListItems(ii).SubItems(6))
            Next
        End If
        
    End If
    
    imgHx.Visible = False
    imgList.Visible = False
End Sub

Private Sub DispFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pSTS As String)
    
    Dim strFNote As String

    If fFNSeq > 0 Then

        strFNote = objETest.ReadFootNote(pWorkArea, pAccDt, pAccSeq)
        
        txtPNote.Text = "": txtFNote.Text = ""
            
        If pSTS < enStsCd.StsCd_LIS_FinRst Then
            txtFNote.Text = strFNote
            fFNSeq = 0                          ' 만약 데이타 이상이 있어도 자동 교정( 세이브후 )
        Else                                    ' 이미 확인된 결과이면 추가만 가능
            txtPNote.Text = strFNote
        End If
    
    End If

End Sub

Private Sub GetLastAccNo(ByVal pPtId As String)
    
    Dim sqlLast As String, dsLast As Recordset
    Dim sTWorkArea As String, sTAccDt As String, sTAccSeq As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim objSQL As New clsLISSqlETest
    Dim RS As Recordset
    Dim iTmx As ListItem

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

    Set dsLast = New Recordset
    dsLast.Open objSQL.SqlGetLastLabNo(pPtId, fTestCd), DBConn
    
    If dsLast.EOF Then
        lblLastDate.Caption = cNoLastRst
        sstRst.Tab = 1: sstRst.Caption = cNoLastRst
        sstRst.Tab = 0
        Set dsLast = Nothing
        Set objSQL = Nothing
        Exit Sub
    End If
    
    '----------------
    '과거결과 리스트
    '----------------
    dsLast.MoveFirst
    lvwLResult.ListItems.Clear
    Do Until dsLast.EOF
        If "" & dsLast.Fields("workarea").Value = sWorkArea And "" & dsLast.Fields("accdt").Value = sAccDt And "" & dsLast.Fields("accseq").Value = sAccSeq Then
        Else
            Set RS = Nothing
            Set RS = New Recordset
            RS.Open objSQL.SqlGetLastLabNoData("" & dsLast.Fields("workarea").Value _
                    , "" & dsLast.Fields("accdt").Value, "" & dsLast.Fields("accseq").Value), DBConn
                    
            With lvwLResult
                Set iTmx = .ListItems.Add()
                iTmx.Text = "" & dsLast.Fields("workarea").Value & "-" & _
                            "" & dsLast.Fields("accdt").Value & "-" & _
                            "" & dsLast.Fields("accseq").Value
                iTmx.SubItems(1) = "" & dsLast.Fields("vfydt").Value
                iTmx.SubItems(2) = "" & dsLast.Fields("vfytm").Value
                
                iTmx.SubItems(3) = "" & RS.Fields("deptcd").Value & "/" & "" & RS.Fields("wardid").Value
                iTmx.SubItems(4) = GetEmpNm("" & RS.Fields("orddoct").Value)
                If iTmx.SubItems(4) = "" Then
                    iTmx.SubItems(4) = "" & RS.Fields("orddoct").Value
                End If
                
                If "" & RS.Fields("ageday").Value < 365 Then
                    iTmx.SubItems(5) = "" & RS.Fields("sex").Value & "/" & Str$("" & RS.Fields("ageday").Value) & " D"
                Else
                    iTmx.SubItems(5) = "" & RS.Fields("sex").Value & "/" & Str$((Val("" & RS.Fields("ageday").Value) \ 365) + 1)
                End If
                iTmx.SubItems(6) = "" & dsLast.Fields("workarea").Value
                iTmx.SubItems(7) = "" & dsLast.Fields("accdt").Value
                iTmx.SubItems(8) = "" & dsLast.Fields("accseq").Value
                iTmx.SubItems(9) = "" & dsLast.Fields("testcd").Value
                iTmx.SubItems(10) = "" & dsLast.Fields("mfyseq").Value
                Set iTmx = Nothing
            End With
            Set RS = Nothing
            
        End If
        dsLast.MoveNext
    Loop
    
    dsLast.MoveFirst
    fLWorkArea = "" & dsLast.Fields("workarea").Value
    fLAccDt = "" & dsLast.Fields("accdt").Value
    fLAccSeq = Val("" & dsLast.Fields("accseq").Value)
    fLM = Val("" & dsLast.Fields("mfyseq").Value)
    lblLastDate.Caption = "최근 결과 확인일 : "
    sstRst.Tab = 1: fraLastRst.Visible = False
                    sstRst.Caption = " 최근결과 " & "(" & _
                                        Format(Mid("" & dsLast.Fields("vfydt").Value, 3, 6), "00/00/00") & " " & _
                                        Format(Mid("" & dsLast.Fields("vfytm").Value, 1, 4), "00:00") & ")"
    sstRst.Tab = 0
    
    If sWorkArea = fLWorkArea And sAccDt = fLAccDt And sAccSeq = fLAccSeq Then
        
        dsLast.MoveNext
        If dsLast.EOF Then
            lblLastDate.Caption = cNoLastRst
            sstRst.Tab = 1: sstRst.Caption = cNoLastRst
            sstRst.Tab = 0
            fLM = 9999              ' 반드시 필요 (최근 수치 결과 처리, 급히 수정 하느라...)
        Else
            fLWorkArea = "" & dsLast.Fields("workarea").Value
            fLAccDt = "" & dsLast.Fields("accdt").Value
            fLAccSeq = Val("" & dsLast.Fields("accseq").Value)
            fLM = Val("" & dsLast.Fields("mfyseq").Value)
            lblLastDate.Caption = "최근 결과 확인일 : "
            sstRst.Tab = 1: sstRst.Caption = " 최근 결과 " & "(" & _
                                                Format("" & dsLast.Fields("vfydt").Value, "0000/00/00") & " " & _
                                                Format("" & dsLast.Fields("vfytm").Value, "00:00:00") & ")"
            sstRst.Tab = 0
        End If
        
    End If
    
    Set dsLast = Nothing
    Set objSQL = Nothing

End Sub


Private Sub DispTxtResult()
    
    Dim i As Integer
    Dim sqlRst As String
    Dim sqlLRst As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim objProBar As clsProgress

    Dim pRstCd1 As String, pRstCd2 As String, pRstCd3 As String
    
    Set objProBar = New clsProgress
    With objProBar
        .Container = Me
        .Left = picResult.Left + sstRst.Left + lblMesg.Left
        .Top = picResult.Top + sstRst.Top + lblMesg.Top
        .Width = lblMesg.Width
        .Height = 260
        .Message = "특수검사 내역을 검색중입니다..."
        .Max = 90
        .Value = 10
'        .SetMyForm Me
'        .Choice = True
'        .XPos = picResult.Left + sstRst.Left + lblMesg.Left
'        .YPos = picResult.Top + sstRst.Top + lblMesg.Top
'        .XWidth = lblMesg.Width
''        .ForeColor = &H864B24
'        .ForeColor = DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = 260
'        .MSG = "특수검사 내역을 검색중입니다..."
'        .Max = 90
'        .Min = 0
'        .Value = 10
        DoEvents
    End With

    MouseRunning
    
    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If

    If lblLastDate.Caption <> cNoLastRst Then
        txtLastRst.TextRTF = objETest.GetResultText(fLWorkArea, fLAccDt, fLAccSeq, fTestCd, fLM)
    End If
    
    '관련검사결과
    objProBar.Message = "관련검사내역을 검색중입니다..."
    DoEvents
    For i = 1 To 40
        objProBar.Value = i
        DoEvents
    Next
    
    Dim MyResult As New clsLISSpecialTest
    If MyResult.DisplayRelTest(sWorkArea, sAccDt, sAccSeq, tblResult) Then
        tblResult.Row = 1
        tblResult.Col = 7: lblColDtTm.Caption = tblResult.Value
        tblResult.Col = 8: lblVfyDtTm.Caption = tblResult.Value
    Else
        lblColDtTm.Caption = ""
        lblVfyDtTm.Caption = ""
    End If
    
    ' 현재 결과
    objProBar.Message = "현재 결과내역을 검색중입니다..."
    DoEvents
    For i = 41 To 70
        objProBar.Value = i
        DoEvents
    Next
    
    txtCRst.Text = "": txtRst.Text = ""
    
'    If fStatus < enStsCd.StsCd_LIS_FinRst Then              ' 결과 등록 Mode
        txtRst.TextRTF = objETest.GetSpeResultText(sWorkArea, sAccDt, sAccSeq, fTestCd, pRstCd1, pRstCd2, pRstCd3, fMfySeq)
        lblRstCd(0).Caption = pRstCd1
        lblRstCd(1).Caption = pRstCd2
        lblRstCd(2).Caption = pRstCd3
'        txtRst.SelStart = 0
'        txtRst.SelLength = Len(txtRst.Text)
'        txtRst.SelProtected = True
'        slider.Value = slider.Min
'    Else                                                    ' 결과 수정 Mode
'        i = 0
'        txtCRst.TextRTF = objETest.GetResultText(sWorkArea, sAccDt, sAccSeq, fTestCd)   ', fMfySeq)
'        txtCRst.Text = vbCrLf & "<< Text 결과 (수정) >>" & vbCrLf & vbCrLf & Trim(txtCRst.Text)
'        txtCRst.SelStart = 0
'        txtCRst.SelLength = Len(txtCRst.Text)
'        txtCRst.SelProtected = False
'        Call HighlightText(txtCRst, "<< Text 결과 (수정) >>", False, , vbRed)
'        txtCRst.SelStart = 0
'        txtCRst.SelLength = Len(txtCRst.Text)
'        txtCRst.SelProtected = True
'        txtCRst.SelLength = 0
'        txtCRst.SelStart = 0
'        txtRst.Text = ""
'        slider.Value = slider.Max / 8 * 7
'    End If
    
    For i = 71 To 90
        objProBar.Value = i
        DoEvents
    Next

    MouseDefault

End Sub


Private Sub lstTest_LostFocus()
    lstTest.Visible = False
End Sub

Private Sub lstTstNo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim sTmp As String, sTstCd As String, sTstNm As String, sAccDt As String
    Dim sqlLN As String, dsLN As New Recordset, iLNCol As Integer
    
    If lstTstNo.ListIndex < 0 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        
        Call cmdClear_Click
        
        sTmp = lstTstNo.List(lstTstNo.ListIndex)
        sTstCd = medGetP(sTmp, 1, vbTab)
        sTstNm = Trim(medGetP(sTmp, 2, vbTab))
        
        lblTestNm.Caption = sTstNm
        
        If optInput(0).Value Then
            Call objETest.GetAccList(sTstCd, lstLabNo)
        Else
            Call objETest.GetLabNoList(sTstCd, lstLabNo)
        End If
        
    End If
    
End Sub

Private Sub lstTstNo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If Button = vbLeftButton Then
        Call lstTstNo_KeyDown(vbKeyReturn, 0)
   End If
   
End Sub

Private Sub lvwHxList_DblClick()
    If lvwHxList.ListItems.Count > 0 Then
        imgHx.Picture = LoadPicture(lvwHxList.SelectedItem.SubItems(6))
    End If
End Sub

Private Sub lvwList_DblClick()
    If lvwList.ListItems.Count > 0 Then
        imgList.Picture = LoadPicture(lvwList.SelectedItem.SubItems(6))
    End If
End Sub

Private Sub lvwResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRst.SetFocus
        fraResult.Visible = False
    End If
End Sub

Private Sub lvwResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iTmx As ListItem
    
    Set iTmx = lvwResult.HitTest(X, Y)
    If iTmx Is Nothing Then Exit Sub
    
    txtRst.SelProtected = False
    txtRst.SelColor = DCM_Black
    txtRst.SelText = iTmx.Text
    
'    If lblRstCd(0).Caption = "" Then
'        lblRstCd(0).Caption = itmx.SubItems(1)
'    ElseIf lblRstCd(1).Caption = "" Then
'        lblRstCd(1).Caption = itmx.SubItems(1)
'    ElseIf lblRstCd(2).Caption = "" Then
'        lblRstCd(2).Caption = itmx.SubItems(1)
'    End If
  
    txtRst.SetFocus
    fraResult.Visible = False
    DoEvents
    
    'lvwResult.ZOrder 1
    'DoEvents
    txtRst.SetFocus

End Sub

Private Sub lvwResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraResult.Visible = False
    txtRst.SetFocus
End Sub

Private Sub objDesc_DescClick(ByVal SelDesc As String)
    
    If SelDesc <> "" Then txtRst.TextRTF = SelDesc
  
    Set objDesc = Nothing
    
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_SAVE
            Dim strImgDir As String
            
            If Trim(strImagePath) = "" Then Exit Sub
            
            strImgDir = strImagePath
            
            DlgSave.InitDir = "C:\"
            DlgSave.Filter = "JPEG"
            DlgSave.FileName = Mid(strImgDir, InStrRev(strImgDir, "\", , vbTextCompare) + 1, Len(strImgDir))
            DlgSave.ShowSave
        
            FileCopy strImgDir, DlgSave.FileName
    End Select
End Sub

Private Sub objTemplet_DescClick(ByVal SelDesc As String)
    If SelDesc <> "" Then
        txtRst.TextRTF = SelDesc
        txtRst.SelStart = 0
        txtRst.SelLength = Len(txtRst.Text)
        txtRst.SelProtected = True
        txtRst.SelStart = 0
        txtRst.SelLength = 0
    End If
    Set objTemplet = Nothing
End Sub

Private Sub objTemplet2_DescClick(ByVal SelDesc As String)
    If SelDesc <> "" Then
        txtRst.TextRTF = SelDesc
        txtRst.SelStart = 0
        txtRst.SelLength = Len(txtRst.Text)
        txtRst.SelProtected = True
        txtRst.SelStart = 0
        txtRst.SelLength = 0
    End If
    Set objTemplet2 = Nothing
End Sub

Private Sub optInput_Click(Index As Integer)
   
    On Error GoTo Err_Trap
   
    lstLabNo.Clear
    If Index = 0 Then
        ClearForm
        txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
        txtWorkArea.Enabled = True: txtAccDt.Enabled = True: txtAccSeq.Enabled = True
'        lstTstNo.Enabled = False: lstLabNo.Enabled = False
        If txtWorkArea.Enabled Then txtWorkArea.SetFocus
    Else
        ClearForm
        txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
        txtWorkArea.Enabled = False: txtAccDt.Enabled = False: txtAccSeq.Enabled = False
        lstTstNo.Enabled = True: lstLabNo.Enabled = True
    End If
Err_Trap:
    Resume Next
End Sub


Private Sub slider_Change()

'    txtCRst.Visible = False: txtRst.Visible = False
'
'    If slider.Value < slider.Min + 200 Then slider.Value = slider.Min + 200
'    If slider.Value > slider.Max - 200 Then slider.Value = slider.Max - 200
'
'    txtCRst.Height = slider.Value - slider.Min
'    txtRst.Top = txtCRst.Height + slider.Min
'    txtRst.Height = slider.Max - txtRst.Top
'
'    txtCRst.Visible = True: txtRst.Visible = True

End Sub

Private Sub sstRst_Click(PreviousTab As Integer)
    If lvwLResult.ListItems.Count = 0 Then Exit Sub
    Select Case sstRst.Tab
        Case "1": fraLastRst.Visible = True
        Case Else
            fraLastRst.Visible = False
    End Select
End Sub

Private Sub txtRst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cursorPos As Long
    Dim lineCount As Long
    Dim ChrsUpToLast As Long
    Dim lastLineLen As Long
    Dim sTemp As String
    Dim sRType As String
    Dim sTCode As String
    Dim iSPos As Long, iEPos As Long
    Dim strKey As String
    
    On Local Error Resume Next
    
    If KeyCode = vbKeyF2 Then
        'If cboTemplate.ListIndex < 0 Then Exit Sub
        
        sTemp = cboTemplate.List(cboTemplate.ListIndex)
        sRType = medGetP(sTemp, 4, vbTab)
        sTCode = medGetP(sTemp, 1, vbTab)
        
        If blnTmpDisplay Then
            With txtRst
                iSPos = .Find("<#" & sTCode & "_VALUE", .SelStart)
                If iSPos < 0 Then
                    iSPos = .Find("<#" & sTCode & "_VALUE", 0)
                    If iSPos < 0 Then Exit Sub
                End If
    
                iEPos = .Find(">", iSPos)
                .SelStart = iSPos
                .SelLength = iEPos - iSPos + 1
                .SelProtected = False
                strKey = .SelText
            End With
            
            Call objETest.LoadRstTemplate(strKey, lvwResult)
            fraResult.Visible = True
            fraResult.ZOrder 0
            lvwResult.SetFocus
        End If
    ElseIf KeyCode = vbKeyF3 Then
        strRtfHead = ""
        strRtfEnd = ""
        
        'get the character position of the cursor
        cursorPos = SendMessage(txtRst.hwnd, _
                                EM_GETSEL, 0, ByVal 0&) \ &H10000
    
       'SELECT the text FROM position 0 to the cursor
        txtRst.SetFocus
        Call SendMessage(txtRst.hwnd, EM_SETSEL, 0, ByVal cursorPos)
        
        strRtfHead = txtRst.SelRTF
        
       
       'get the cursor position in the textbox
        cursorPos = SendMessage(txtRst.hwnd, _
                                EM_GETSEL, 0, ByVal 0&) \ &H10000
    
       'get the number of lines in the textbox
        lineCount = SendMessage(txtRst.hwnd, _
                                EM_GETLINECOUNT, 0, ByVal 0&)
       
       'the number of characters in the textbox,
       'up to but not including the the last line
       '(0-based)
        ChrsUpToLast = SendMessage(txtRst.hwnd, _
                                   EM_LINEINDEX, _
                                   lineCount - 1, ByVal 0&)
    
       'the number of characters in the last line
        lastLineLen = SendMessage(txtRst.hwnd, _
                                  EM_LINELENGTH, _
                                  lineCount, ByVal 0&)
    
       'SELECT the text FROM the cursor
       'position to the last line
        txtRst.SetFocus
        Call SendMessage(txtRst.hwnd, _
                         EM_SETSEL, _
                         cursorPos, _
                         ByVal ChrsUpToLast + lastLineLen)
       
        strRtfEnd = txtRst.SelRTF
        
        LoadData
        fraTest.Visible = True
    End If
End Sub

Private Sub LoadData()
    Dim RS          As New Recordset
    Dim objSQL      As New clsLISSqlStatement
    Dim strSQL      As String
    Dim strTitle    As String
    Dim ii          As Long
    Dim jj          As Long
  
    medClearTable tblData
    lblTest.Caption = ""
    strSQL = objSQL.SqlLAB031CodeList(LC2_TempletTest, "cdval2,field1,field2,field3,field4,text1", fTestCd, , "ORDER BY cdval2")
    RS.Open strSQL, DBConn
    lblTitle1.Caption = ""
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        With tblData
            .MaxRows = medGetP(RS.Fields("field4").Value & "", 1, "*")
            .MaxCols = medGetP(RS.Fields("field4").Value & "", 2, "*") * 3
            If lblTitle1.Caption = "" Then lblTitle1.Caption = RS.Fields("text1").Value & ""
            ii = 1
            Do Until RS.EOF
                .Row = ii
                .RowHeight(ii) = 16.8
                For jj = 1 To .MaxCols
                    .Col = jj
                    If (jj Mod 3) = 1 Then
                        .Value = RS.Fields("field1").Value & ""
                        .ColWidth(jj) = 15.63
                    ElseIf (jj Mod 3) = 2 Then
                        .Value = RS.Fields("field2").Value & ""
                        .ColWidth(jj) = 8
                    Else
                        .Value = RS.Fields("field3").Value & ""
                        .ColWidth(jj) = 5
                        RS.MoveNext
                    End If
                Next jj
                ii = ii + 1
            Loop
        End With
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub txtWorkArea_Change()
    On Error GoTo Err_Trap
    If Not txtAccDt.Enabled Then Exit Sub
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtWorkArea_GotFocus()
    txtWorkArea.SelStart = 0
    txtWorkArea.SelLength = Len(txtWorkArea)
End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Trap
    
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    
    If Not txtAccDt.Enabled Then Exit Sub
    If KeyAscii = vbKeyReturn And Len(txtWorkArea) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtAccDt_Change()
    On Error GoTo Err_Trap
    If Not txtAccSeq.Enabled Then Exit Sub
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then txtAccSeq.SetFocus
Err_Trap:
    Resume Next
End Sub

Private Sub txtAccDt_GotFocus()
    txtAccDt.SelStart = 0
    txtAccDt.SelLength = Len(txtAccDt)
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Err_Trap
    
    If KeyAscii = vbKeyReturn And Len(txtAccDt) >= (txtAccDt.MaxLength - 4) Then
        If txtAccSeq.Enabled Then txtAccSeq.SetFocus
    End If
    
    ' 숫자와 백스페이스만 허용
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If

Err_Trap:
    Resume Next
End Sub

Private Sub txtAccSeq_GotFocus()
    txtAccSeq.SelStart = 0
    txtAccSeq.SelLength = Len(txtAccSeq)
End Sub

Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)

    Dim i As Integer
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim sqlTest As String, dsTest As Recordset
    Dim iTestCount As Integer
    Dim sTestcd As String, sTestNm As String, sRstFg As String, sRstType As String, iMfySeq As Integer
    Dim objSQL As New clsLISSqlETest

    On Error GoTo Err_Trap
    
    If KeyAscii <> 13 Or txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then Exit Sub
    
    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    If Mid$(sAccDt, 1, 1) = "9" Then
       sAccDt = "19" & sAccDt
    Else
       sAccDt = "20" & sAccDt
    End If
    
    
    '병동/진료과 연락처(환자ID,CONTROL)
    Call GetPtTelInfo(sWorkArea, sAccDt, sAccSeq, lblTelNo)

    Set dsTest = New Recordset
    dsTest.Open objSQL.SqlGetResultData(sWorkArea, sAccDt, sAccSeq), DBConn

    iTestCount = dsTest.RecordCount
    
    Select Case iTestCount
        
        Case Is < 1
            MsgBox "임상병리 특수검사 접수 되지 않은 Lab-No 입니다. 확인 후 처리하세요"
            Call ClearForm
            Exit Sub
        
        Case Is = 1
            lstTest.Clear
            sTestcd = "" & dsTest.Fields("testcd").Value
            sTestNm = "" & dsTest.Fields("testnm").Value
            sRstFg = "" & dsTest.Fields("stscd").Value
            sRstType = "" & dsTest.Fields("rsttype").Value
            iMfySeq = Val("" & dsTest.Fields("mfyseq").Value)
            lstTest.AddItem sTestcd & vbTab & sRstFg & vbTab & sTestNm & _
                            vbTab & vbTab & sRstType & vbTab & iMfySeq
            lstTest.ListIndex = 0
            Call lstTest_KeyPress(vbKeyReturn)
                
        Case Is > 1
    
            lstTest.Clear
            dsTest.MoveFirst
            Do Until dsTest.EOF
                sTestcd = "" & dsTest.Fields("testcd").Value
                sTestNm = "" & dsTest.Fields("testnm").Value
                sRstFg = "" & dsTest.Fields("stscd").Value
                sRstType = "" & dsTest.Fields("rsttype").Value
                iMfySeq = Val("" & dsTest.Fields("mfyseq").Value)
                lstTest.AddItem sTestcd & vbTab & sRstFg & vbTab & sTestNm & _
                                vbTab & vbTab & sRstType & vbTab & iMfySeq
                lstTest.ListIndex = 0
                dsTest.MoveNext
            Loop

            lstTest.Visible = True
            lstTest.ZOrder 0
            lstTest.SetFocus
            
    End Select

    Set dsTest = Nothing
    Set objSQL = Nothing

    ' 만약에 숫자가 아니면 문자를 없애버려도 좋음(백스페이스 허용)
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If
Err_Trap:
    Resume Next
End Sub

Public Sub Call_txtAccSeq_KeyPress()
    Call txtAccSeq_KeyPress(vbKeyReturn)
End Sub

Private Function GetLisDoctNm(ByVal pDoctID As String) As String
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    
    On Error Resume Next
    
    strSQL = " select username " & _
             "   from " & T_HIS005 & _
             "  where userid = " & DBS(pDoctID)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        GetLisDoctNm = RS.Fields("username").Value & ""
    End If
    
    RS.Close
    Set RS = Nothing

End Function



