VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm255MStain 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Stain 결과등록"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14640
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   14640
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
      Left            =   6180
      TabIndex        =   77
      Top             =   1740
      Width           =   4515
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
         TabIndex        =   92
         Tag             =   "opt"
         Top             =   1350
         Width           =   1305
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
         TabIndex        =   91
         Tag             =   "opt"
         Top             =   4170
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
         TabIndex        =   90
         Tag             =   "opt"
         Top             =   1410
         Width           =   2325
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
         TabIndex        =   89
         Tag             =   "opt"
         Top             =   2580
         Width           =   3195
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
         TabIndex        =   88
         Tag             =   "opt"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   87
         Tag             =   "opt"
         Top             =   1020
         Width           =   1005
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
         TabIndex        =   86
         Tag             =   "opt"
         Top             =   1020
         Width           =   1305
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
         TabIndex        =   85
         Tag             =   "opt"
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtTransNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   84
         Tag             =   "opt"
         Top             =   300
         Width           =   1875
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
         TabIndex        =   83
         Tag             =   "opt"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   3030
         Style           =   1  '그래픽
         TabIndex        =   82
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.CommandButton cmdTrans 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전송"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   81
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
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
         TabIndex        =   80
         Tag             =   "opt"
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox txtExDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   79
         Tag             =   "opt"
         Top             =   1800
         Width           =   1005
      End
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
         TabIndex        =   78
         Tag             =   "opt"
         Top             =   2190
         Width           =   2325
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   18
         Left            =   180
         TabIndex        =   93
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
         Index           =   19
         Left            =   180
         TabIndex        =   94
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
         TabIndex        =   95
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
         TabIndex        =   96
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
         TabIndex        =   97
         Top             =   2970
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         ScrollBars      =   2
         TextRTF         =   $"Lis255.frx":0000
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
         TabIndex        =   98
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
         TabIndex        =   99
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
         TabIndex        =   100
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
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2700
      TabIndex        =   76
      Text            =   "Text3"
      Top             =   45
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1620
      TabIndex        =   75
      Text            =   "Text2"
      Top             =   45
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   855
      TabIndex        =   74
      Text            =   "Text1"
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdSMS 
      BackColor       =   &H008080FF&
      Caption         =   "SMS"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   71
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraLabList 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Index           =   0
      Left            =   15
      TabIndex        =   45
      Top             =   1100
      Width           =   3500
      Begin VB.ComboBox cboMonth 
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
         ItemData        =   "Lis255.frx":009D
         Left            =   1050
         List            =   "Lis255.frx":009F
         Style           =   2  '드롭다운 목록
         TabIndex        =   70
         Top             =   1200
         Width           =   705
      End
      Begin VB.CheckBox chkViewAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체 항목 보기"
         Height          =   240
         Left            =   1680
         TabIndex        =   65
         Top             =   2280
         Width           =   1545
      End
      Begin VB.ListBox lstAccList 
         Appearance      =   0  '평면
         BackColor       =   &H00FEF7ED&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4905
         Left            =   120
         TabIndex        =   60
         Top             =   2520
         Width           =   3210
      End
      Begin VB.CommandButton cmdWSList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         Style           =   1  '그래픽
         TabIndex        =   59
         Top             =   780
         Width           =   345
      End
      Begin VB.TextBox txtWSUnit 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   1050
         TabIndex        =   58
         Text            =   "19990005"
         Top             =   780
         Width           =   1905
      End
      Begin VB.ComboBox cboWSCode 
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
         ItemData        =   "Lis255.frx":00A1
         Left            =   1050
         List            =   "Lis255.frx":00A3
         Style           =   2  '드롭다운 목록
         TabIndex        =   57
         Top             =   345
         Width           =   2025
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00F4F0F2&
         Caption         =   "이전(&P)"
         Height          =   510
         Left            =   195
         Style           =   1  '그래픽
         TabIndex        =   56
         Top             =   7500
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00F4F0F2&
         Caption         =   "다음(&N)"
         Height          =   510
         Left            =   1185
         Style           =   1  '그래픽
         TabIndex        =   55
         Top             =   7500
         Width           =   945
      End
      Begin VB.CheckBox chkAutoNext 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Auto Next"
         Height          =   225
         Left            =   2280
         TabIndex        =   54
         Top             =   7560
         Value           =   1  '확인
         Width           =   1125
      End
      Begin VB.CheckBox chkFix 
         BackColor       =   &H00DBE6E6&
         Caption         =   "고정"
         Height          =   210
         Left            =   2280
         TabIndex        =   53
         Top             =   7800
         Value           =   1  '확인
         Width           =   660
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   345
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
         Caption         =   "검 체 군 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   780
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
         Caption         =   "WS Unit"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   1650
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
         Caption         =   "작성일/시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   1995
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
         Caption         =   "마감일/시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   17
         Left            =   120
         TabIndex        =   69
         Top             =   1200
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
         Caption         =   "조회기간"
         Appearance      =   0
      End
      Begin VB.Label lblBltDate 
         BackStyle       =   0  '투명
         Caption         =   "Feb 03 1999 10:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   67
         Top             =   1620
         Width           =   2205
      End
      Begin VB.Label lblRcvDate 
         BackStyle       =   0  '투명
         Caption         =   "Feb 03 1999 10:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1095
         TabIndex        =   66
         Top             =   1965
         Width           =   2205
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   3255
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.OptionButton optGetList 
      BackColor       =   &H00FFFCF7&
      Caption         =   "결과입력대상"
      Height          =   360
      Index           =   0
      Left            =   165
      Style           =   1  '그래픽
      TabIndex        =   44
      Top             =   675
      Width           =   1590
   End
   Begin VB.OptionButton optGetList 
      BackColor       =   &H00EDE2ED&
      Caption         =   "최종결과대상"
      Height          =   360
      Index           =   1
      Left            =   1860
      Style           =   1  '그래픽
      TabIndex        =   43
      Top             =   675
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1770
      Left            =   3600
      TabIndex        =   12
      Top             =   -75
      Width           =   10995
      Begin VB.TextBox txtBarNo 
         BorderStyle     =   0  '없음
         Height          =   285
         Left            =   1170
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   225
         Width           =   1635
      End
      Begin VB.CheckBox chkBar 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3285
         TabIndex        =   72
         Top             =   225
         Width           =   240
      End
      Begin VB.CommandButton cmdOrderView 
         BackColor       =   &H00F4F0F2&
         Caption         =   "처방별조회(&C)"
         Height          =   390
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   68
         Top             =   155
         Visible         =   0   'False
         Width           =   1300
      End
      Begin VB.ComboBox cboRelTest 
         BackColor       =   &H00FFF9F7&
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
         Left            =   1470
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   1380
         Width           =   7995
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
         Height          =   225
         Left            =   2580
         MaxLength       =   5
         TabIndex        =   22
         Text            =   "10011"
         Top             =   255
         Width           =   555
      End
      Begin VB.TextBox txtAccDt 
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
         Height          =   240
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "9906"
         Top             =   255
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
         Height          =   240
         Left            =   1185
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "41"
         Top             =   255
         Width           =   330
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   635
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   8400
         TabIndex        =   14
         Top             =   180
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
         Caption         =   "연락처 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   5280
         TabIndex        =   15
         Top             =   180
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   5280
         TabIndex        =   16
         Top             =   615
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
         Caption         =   "병      동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   990
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
         Caption         =   "검     체"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   18
         Top             =   615
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
         Caption         =   "환자정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   9480
         TabIndex        =   19
         Top             =   975
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "F2  : 결과조회"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   2610
         TabIndex        =   25
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
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
      Begin MedControls1.LisLabel lblPtId 
         Height          =   330
         Left            =   1110
         TabIndex        =   26
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
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
         Left            =   4290
         TabIndex        =   27
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
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
      Begin MedControls1.LisLabel lblSpecimen 
         Height          =   330
         Left            =   1080
         TabIndex        =   29
         Top             =   990
         Width           =   4155
         _ExtentX        =   7329
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
      Begin MedControls1.LisLabel lblDept 
         Height          =   330
         Left            =   6250
         TabIndex        =   30
         Top             =   180
         Width           =   1980
         _ExtentX        =   3493
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
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTelno 
         Height          =   345
         Left            =   9405
         TabIndex        =   31
         Top             =   180
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   9495
         TabIndex        =   32
         Top             =   1320
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "Esc : 결과숨김"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
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
         Caption         =   "관련검사 결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   15
         Left            =   5280
         TabIndex        =   39
         Top             =   990
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   330
         Left            =   6250
         TabIndex        =   40
         Top             =   990
         Width           =   3060
         _ExtentX        =   5398
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   16
         Left            =   8400
         TabIndex        =   41
         Top             =   600
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   345
         Left            =   9405
         TabIndex        =   42
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Label lblWardId 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   180
         Left            =   420
         TabIndex        =   34
         Top             =   1380
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblMajDoct 
         Caption         =   "주치의"
         Height          =   195
         Left            =   435
         TabIndex        =   33
         Top             =   1545
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblWard 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         Caption         =   "5NCU-01-12"
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
         Height          =   330
         Left            =   6250
         TabIndex        =   28
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label4 
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
         Left            =   1560
         TabIndex        =   23
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label5 
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
         Left            =   2355
         TabIndex        =   24
         Top             =   255
         Width           =   195
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F1F5F4&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         Height          =   360
         Left            =   1110
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00F1F5F4&
      Height          =   2220
      ItemData        =   "Lis255.frx":00A5
      Left            =   960
      List            =   "Lis255.frx":00A7
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CommandButton cmdCommentTemplete 
      BackColor       =   &H00DEDBDD&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5145
      Picture         =   "Lis255.frx":00A9
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   6165
      Width           =   300
   End
   Begin VB.ComboBox cboRemark 
      BackColor       =   &H00F1F5F4&
      Height          =   300
      ItemData        =   "Lis255.frx":05DB
      Left            =   5475
      List            =   "Lis255.frx":05DD
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   8115
      Width           =   4335
   End
   Begin FPSpread.vaSpread ssRst 
      Height          =   4395
      Left            =   3615
      TabIndex        =   1
      Top             =   1725
      Width           =   10830
      _Version        =   196608
      _ExtentX        =   19103
      _ExtentY        =   7752
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      EditEnterAction =   2
      EditModePermanent=   -1  'True
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
      GrayAreaBackColor=   15857140
      MaxCols         =   9
      MaxRows         =   10
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "Lis255.frx":05DF
      UserResize      =   0
      TextTip         =   1
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdVerify 
      BackColor       =   &H00F4F0F2&
      Caption         =   "결과 확인(&S)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8535
      Width           =   1320
   End
   Begin RichTextLib.RichTextBox txtFNote 
      Height          =   1905
      Left            =   5460
      TabIndex        =   2
      Top             =   6165
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   3360
      _Version        =   393217
      BackColor       =   15857140
      BorderStyle     =   0
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Lis255.frx":0C1B
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
   Begin MedControls1.LisLabel lblRemark 
      Height          =   300
      Left            =   9810
      TabIndex        =   7
      Top             =   8100
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   529
      BackColor       =   14737632
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
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   5
      Left            =   3615
      TabIndex        =   35
      Top             =   6150
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
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
      Caption         =   "◈ Foot Note"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   13
      Left            =   3615
      TabIndex        =   36
      Top             =   8115
      Width           =   1515
      _ExtentX        =   2672
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
      Caption         =   "◈ 검체 Remark"
      Appearance      =   0
   End
   Begin VB.Frame fraLabList 
      BackColor       =   &H00DBE6E6&
      Height          =   4905
      Index           =   1
      Left            =   15
      TabIndex        =   46
      Top             =   1100
      Width           =   3500
      Begin VB.ListBox lstFinList 
         Appearance      =   0  '평면
         BackColor       =   &H00F4FDF5&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3810
         Left            =   150
         TabIndex        =   49
         Top             =   960
         Width           =   3135
      End
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
         Height          =   300
         Left            =   2130
         MaskColor       =   &H00C0FFFF&
         Style           =   1  '그래픽
         TabIndex        =   47
         Tag             =   "128"
         Top             =   270
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpMidFVfyDt 
         Height          =   285
         Left            =   240
         TabIndex        =   48
         Top             =   585
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
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
         Format          =   85655553
         CurrentDate     =   37083
      End
      Begin MSComCtl2.DTPicker dtpMidVfyDt 
         Height          =   285
         Left            =   1920
         TabIndex        =   50
         Top             =   585
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
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
         Format          =   85655553
         CurrentDate     =   37083
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00B88FA5&
         BorderWidth     =   2
         Height          =   4650
         Left            =   105
         Top             =   180
         Width           =   3255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "▷ 중간보고일"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   345
         Width           =   1140
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   225
         Left            =   1605
         TabIndex        =   51
         Tag             =   "40110"
         Top             =   630
         Width           =   195
      End
   End
   Begin VB.ListBox lstBtRCd 
      Appearance      =   0  '평면
      BackColor       =   &H00F5FDEE&
      Columns         =   2
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   8940
      TabIndex        =   9
      Top             =   6135
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.ListBox lstRstCd 
      Appearance      =   0  '평면
      BackColor       =   &H00FBFCEB&
      Columns         =   2
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   3600
      TabIndex        =   6
      Top             =   6135
      Width           =   5325
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "WorkSheet Unit 별 결과 등록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   3165
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFCF7&
      FillStyle       =   0  '단색
      Height          =   420
      Left            =   75
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "frm255MStain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public WithEvents clsTemplete As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1

Dim fWorkSheet() As tpMicWorkSheet
Dim fFNSeq As Integer

Dim fSkColor As Long                    ' 입력 가능 셀 배경색
Dim fOkColor As Long                    ' 입력 불가 셀 배경색
Dim blnPtFg As Boolean

Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture

Private mvarCurRow  As Long

Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset
Dim strRcvDt            As String

Private Sub cboRemark_Click()
    
    Dim iIndex As Integer, sRMCd As String, sRMNm As String

    iIndex = cboRemark.ListIndex
    If iIndex < 0 Then Exit Sub

    sRMCd = Trim(Mid(cboRemark.List(iIndex), 1, 6))
    If sRMCd = LIS_Nothing Then lblRemark.Caption = "": Exit Sub

    lblRemark.Caption = objMicRst.GetRemark(sRMCd)
    
End Sub

Private Sub chkBar_Click()
    If chkBar.Value = 0 Then
        LisLabel4(6).Caption = "접수 번호"
        txtWorkArea.Visible = True
        txtAccDt.Visible = True
        txtAccSeq.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        txtBarNo.Visible = False
    Else
        LisLabel4(6).Caption = "검체 번호"
        txtWorkArea.Visible = False
        txtAccDt.Visible = False
        txtAccSeq.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        txtBarNo.Visible = True
    End If
End Sub

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
End Sub

Private Sub cmdOrderView_Click()
' 2008.12.17. 양성현 작업중입니다.
' 2009.01.09 양성현 환자ID 파라메터 추가
' 2009.04.13 양성현 추가
    Dim i As Integer
    Dim pFrmName As String
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

End Sub

Private Sub cboWSCode_Click()
    
    Dim i As Integer

    If cboWSCode.ListIndex < 0 Then Exit Sub

    txtWSUnit = ""
    Call ScreenClear
    If txtWorkArea.Enabled Then txtWorkArea.SetFocus

End Sub


Private Sub clsTemplete_CopyTemplete()
   '
    txtFNote.Text = clsTemplete.rtfText.Text
    txtFNote.SetFocus
    Set clsTemplete = Nothing

End Sub

Private Sub cmdClear_Click()
    
    txtWSUnit = ""
    Call ScreenClear
    If chkFix.Value = 0 Then
        cboWSCode.ListIndex = -1
        cboWSCode.SetFocus
    Else
        If chkBar.Value = 0 Then
            If txtWorkArea.Enabled Then txtWorkArea.SetFocus
        Else
            If txtBarNo.Enabled Then txtBarNo.SetFocus
        End If
    End If
End Sub

Private Sub cmdCommentTemplete_Click()

   If ssRst.MaxRows < 1 Then Exit Sub
   Call CallTemplete(3, 0)

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm255MStain = Nothing
End Sub

Private Sub cmdNext_Click()

    If lstAccList.ListIndex < lstAccList.ListCount - 1 Then
        lstAccList.ListIndex = lstAccList.ListIndex + 1
        Call lstAccList_KeyDown(vbKeyReturn, 0)
    End If

    DoEvents

End Sub

Private Sub cmdPrev_Click()

    If lstAccList.ListIndex > 0 Then
        lstAccList.ListIndex = lstAccList.ListIndex - 1
        Call lstAccList_KeyDown(vbKeyReturn, 0)
    End If

    DoEvents

End Sub

Private Sub cmdRefresh_Click()
    Call dtpMidVfyDt_Change
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
    txtDeptNm.Text = lblDept.Caption
    
    rtfMessage.Text = "환자명 : " & Trim(lblPtNm.Caption) & "(" & Trim(lblPtId.Caption) & ")"
    rtfMessage.Text = rtfMessage.Text & vbCRLF & txtFNote.Text
    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value 즉시처치요함"

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
    
'    Dim SSQL As String
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
'    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
'
'    txtDtNm.Text = lblDoctNm.Caption
'    txtDeptNm.Text = lblDept.Caption
'    rtfMessage.Text = "환자명 : " & Trim(lblPtNm.Caption) & "(" & Trim(lblPtId.Caption) & ")"
'    rtfMessage.Text = rtfMessage.Text & vbCRLF & txtFNote.Text
'    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value 즉시처치요함"
'
'    If txtDtNm.Text <> "" Then
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
''        SSQL = ""
''        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
''        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
'        Set AdoRs_ORACLE = New ADODB.Recordset
'
'        AdoRs_ORACLE.CursorLocation = adUseClient
'        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
'
'        If AdoRs_ORACLE.RecordCount > 0 Then
'            txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
'            txtDtId.Text = AdoRs_ORACLE.Fields("EMPNO") & ""
'        End If
'
'        Set AdoCn_ORACLE = Nothing
'    End If
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
    
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    Dim sRemarkCd As String, sWsCd As String
    Dim sSysDate As String, sDate As String, sTime As String
    Dim blnSave As Boolean, tmpDept As String, tmpBussDiv As String
    Dim strRstVal   As String
    Dim bRstFlag    As Boolean
    Dim i       As Integer
    
    '** 결과값이 없을 경우 결과 등록되지 않도록 수정
    With ssRst
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 5: strRstVal = .Value
            
            If Trim(strRstVal) <> "" Then
                bRstFlag = True
                Exit For
            Else
                bRstFlag = False
            End If
        Next
    End With
    
    If bRstFlag = False Then
        Exit Sub
    End If
    
    pWorkArea = Trim(txtWorkArea.Text): pAccDt = Trim(txtAccDt.Text): pAccSeq = Trim(txtAccSeq.Text)
    pAccDt = IIf(Mid(pAccDt, 1, 1) = "9", "19" & pAccDt, "20" & pAccDt)

    If pWorkArea = "" Or pAccDt = "" Or pAccSeq = "" Then
        MsgBox "접수번호가 정확하지 않습니다. 확인 후 처리 하세요", vbExclamation, "Stain결과등록"
        Exit Sub
    End If

    sSysDate = Format(GetSystemDate, "yyyymmdd hhmmss")
    sDate = Mid$(sSysDate, 1, 8)
    sTime = Mid$(sSysDate, 10, 6)
    
    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    sRemarkCd = Trim(Mid(cboRemark.List(cboRemark.ListIndex), 1, 6))
    If sRemarkCd = LIS_Nothing Then sRemarkCd = ""
    
    blnSave = objMicRst.SaveStainResult(pWorkArea, pAccDt, pAccSeq, ssRst, sWsCd, ObjSysInfo.EmpId, txtFNote.Text, sRemarkCd)
    If Not blnSave Then GoTo DBExecError
    
    '결과보고 대기내역 생성

'    tmpBussDiv = objMicRst.Get_Bussdiv(pWorkArea, pAccDt, pAccSeq)
'    If Trim(lblWardId.Caption) = "" Then
'        tmpDept = lblDept.Caption
'    Else
'        tmpDept = lblWardId.Caption
'    End If
'
'    If tmpBussDiv = "" Then
'        If Trim(lblWardId.Caption) = "" Then
'            tmpDept = lblDept.Caption
'            tmpBussDiv = enBussDiv.BussDiv_OutPatient
'        Else
'            tmpDept = lblWardId.Caption
'            tmpBussDiv = enBussDiv.BussDiv_InPatient
'        End If
'    End If
'
'    DBConn.BeginTrans
'    blnSave = objMicRst.SubmitVerifyList(tmpDept, sDate, sTime, lblPtId.Caption, enStsCd.StsCd_LIS_FinRst, ObjMyUser.EmpId, lblMajDoct.Caption, tmpBussDiv)
'    DBConn.CommitTrans
'    If Not blnSave Then GoTo DBExecError1

'        '감염관리
    Call ICSStainResultCheck(pWorkArea, pAccDt, pAccSeq, lblPtId.Caption, lblPtNm.Caption, _
                                    lblDept.Caption, medGetP(lblWard.Caption, 1, "-"), ssRst)

    ' *** 처리후 다음 데이타 로드
    Call LoadNewData
    
    If chkBar.Value = 0 Then
        txtAccSeq.SetFocus
    End If
    
    Exit Sub
    
DBExecError:
    MsgBox "결과저장 중 오류가 발생했습니다.", vbCritical, "오류"
DBExecError1:
    MsgBox "결과보고대기내역저장중 오류가 발생했습니다.", vbCritical, "오류"
End Sub
 
Private Sub LoadNewData()
    
    Dim iPLIdx As Integer, iPLDat As String

    iPLIdx = lstAccList.ListIndex
    iPLDat = lstAccList.List(iPLIdx)

    Call txtWSUnit_KeyPress(vbKeyReturn)

    If chkAutoNext.Value = 1 Then

         If lstAccList.List(iPLIdx) = iPLDat Then iPLIdx = iPLIdx + 1

         If iPLIdx < lstAccList.ListCount Then
            lstAccList.ListIndex = iPLIdx
         Else
            lstAccList.ListIndex = lstAccList.ListCount - 1
         End If
    End If

    Call lstAccList_KeyDown(vbKeyReturn, 0)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    Select Case KeyCode

        Case vbKeyEscape
            lstBtRCd.Visible = False
            lstRstCd.Visible = False
        Case vbKeyF2
            If Me.ActiveControl.Name = ssRst.Name Then
                Call ssRst_EditMode(ssRst.ActiveCol, ssRst.ActiveRow, 1, True)
            End If

    End Select

'    Me.ActiveControl.SetFocus
'
End Sub

Private Sub Form_Load()

    Me.Show

    KeyPreview = True

    ssRst.Col = enSTAIN.tcTESTNM: ssRst.Row = 1: fSkColor = ssRst.BackColor
    ssRst.Col = enSTAIN.tcRSTCD:  ssRst.Row = 1: fOkColor = ssRst.BackColor

    objMicRst.LoadWorkSheetCode MWS_ForStain, cboWSCode, fWorkSheet
    cboWSCode.ListIndex = -1: txtWSUnit.Text = ""
    objMicRst.LoadRemark cboRemark
    ScreenClear

    chkAutoNext.Value = 1
    chkFix.Value = 1
    txtWorkArea.Enabled = False: txtAccDt.Enabled = False: txtAccSeq.Enabled = False

'    fraWSUnit.Enabled = True
    If ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper Then
         optGetList(1).Enabled = True
'         cmdFinEnter.Enabled = True
    Else
         optGetList(1).Enabled = False
'         cmdFinEnter.Enabled = False
    End If

    

    optGetList(0).Value = True

    cboWSCode.SetFocus
    
    cboMonth.Clear
    cboMonth.AddItem "1"
    cboMonth.AddItem "2"
    cboMonth.AddItem "3"
    cboMonth.AddItem "4"
    cboMonth.AddItem "5"
    cboMonth.AddItem "6"
    cboMonth.ListIndex = 0
    
    frmSMS.Visible = False
    txtBarNo.Visible = False
    txtBarNo.Text = ""
End Sub


Private Sub ScreenClear()

    'txtWSUnit = ""
    fFNSeq = 0
    lstWSUnit.Clear
    lblBltDate.Caption = "": lblRcvDate.Caption = ""
    txtBarNo.Text = ""
    lstAccList.Clear

    Call ClearResult

    lstBtRCd.Visible = False    ': lstBtRCd.ZOrder 0
    lstRstCd.Visible = False    ': lstRstCd.ZOrder 0

    cmdOrderView.Visible = False

End Sub

Private Sub ClearResult()

    If cboWSCode.ListIndex >= 0 Then
        txtWorkArea.Enabled = True: txtAccDt.Enabled = True: txtAccSeq.Enabled = True
    Else
        txtWorkArea.Enabled = False: txtAccDt.Enabled = False: txtAccSeq.Enabled = False
    End If
    txtWorkArea.Text = "": txtAccDt.Text = "":   txtAccSeq.Text = ""
    lblPtId.Caption = "":  lblPtNm.Caption = "": lblPtSA.Caption = ""
    lblDept.Caption = "":  lblWard.Caption = "": lblWardId.Caption = ""
    lblSpecimen.Caption = ""

    ssRst.MaxRows = 0

    txtFNote.Text = ""
    cboRemark.ListIndex = 0: lblRemark.Caption = ""
    
    cboRelTest.Clear
    
End Sub

Private Sub cmdWSList_Click()
    
    Dim sWsCd As String
    Dim sMonth As String

    If cboWSCode.ListIndex < 0 Then Exit Sub

    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    sMonth = cboMonth.Text
    
    'objMicRst.LoadMicWorkList sWsCd, lstWSUnit
    objMicRst.LoadMicWorkList_New sWsCd, sMonth, lstWSUnit
    
    If lstWSUnit.ListCount <= 0 Then Exit Sub
    
    lstWSUnit.ListIndex = 0
    lstWSUnit.Visible = True
    lstWSUnit.ZOrder

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub


Private Sub LoadRstData()
    
    Dim i As Integer, iWSIndex As Integer
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    Dim pSpcYY    As String, pSpcNO As String

    If optGetList(0) Then
        iWSIndex = cboWSCode.ListIndex
    Else
        iWSIndex = 0
    End If
    
    pWorkArea = Trim(txtWorkArea.Text): pAccDt = Trim(txtAccDt.Text): pAccSeq = Trim(txtAccSeq.Text)
    pAccDt = IIf(Mid(pAccDt, 1, 1) = "9", "19" & pAccDt, "20" & pAccDt)

    If optGetList(0) Then
        iWSIndex = cboWSCode.ListIndex
    Else
        iWSIndex = 0
    End If
    
    Call ClearTable

    If chkBar.Value = 0 Then
        Call DispPtInfo(pWorkArea, pAccDt, pAccSeq)
    Else
        Call DispPtInfo_New(pSpcYY, pSpcNO)
    End If
    '병동/진료과 연락처(환자ID,CONTROL)
    
    If chkBar.Value = 0 Then
        Call GetPtTelInfo(pWorkArea, pAccDt, pAccSeq, lblTelno)
    Else
        pWorkArea = Text1.Text
        pAccDt = Text1.Text
        pAccSeq = Text1.Text
        Call GetPtTelInfo(pWorkArea, pAccDt, pAccSeq, lblTelno)
    End If
    If blnPtFg Then
        If chkBar.Value = 0 Then
            Call objMicRst.DispStainTable(ssRst, pWorkArea, pAccDt, pAccSeq, fWorkSheet(iWSIndex).WsRstType)
        Else
            pWorkArea = Text1.Text
            pAccDt = Text2.Text
            pAccSeq = Text3.Text
            Call objMicRst.DispStainTable(ssRst, pWorkArea, pAccDt, pAccSeq, fWorkSheet(iWSIndex).WsRstType)
        End If
        For i = 1 To ssRst.MaxRows
            ssRst.Col = enSTAIN.tcRSTCD: ssRst.Row = i
            If ssRst.CellType = CellTypeEdit Then
                ssRst.Action = ActionActiveCell
                ssRst.SetFocus
                Exit For
            End If
        Next i
    End If
    
'2009.10.06 추가
    cmdOrderView.Visible = True
    
    '** 관련검사 추가 By M.G.Choi 2008.02.19
    Dim MyResult As New clsLISResultReview
    
    Call MyResult.GetMicRelTest(cboRelTest, pWorkArea & "-" & pAccDt & "-" & pAccSeq)
    cboRelTest.ListIndex = 0
    '------------------------추가 사항----------------------------
        
    'Call ssRst_LeaveCell(1, 1, ssRst.Col, ssRst.Row, False)

End Sub

Private Sub lstAccList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then Call lstAccList_KeyDown(13, 0)

End Sub

Private Sub ClearTable()
    ssRst.Col = 1: ssRst.COL2 = ssRst.MaxCols
    ssRst.Row = 1: ssRst.Row2 = ssRst.MaxRows
    ssRst.BlockMode = True
    ssRst.Action = ActionClearText
    ssRst.BlockMode = False
End Sub

Private Sub DispPtInfo(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    
    Dim sRemarkCd As String, sRemarkIdx As Integer
    Dim objPtDic As clsDictionary
    Dim iWSIndex As Long

    blnPtFg = False
    
    If optGetList(0) Then
        iWSIndex = cboWSCode.ListIndex
    Else
        iWSIndex = 0
    End If


    Set objPtDic = objMicRst.DispPtInfoByLabno(pWorkArea, pAccDt, pAccSeq, fWorkSheet(iWSIndex).WsRstType)
'    Set objPtDic = objMicRst.DispPtInfoByLabno(pWorkArea, pAccDt, pAccSeq) ' , fWorkSheet(iWSIndex).WsRstType)

    If objPtDic Is Nothing Then
       MsgBox "데이타가 없습니다. 접수번호를 확인하십시오.", vbInformation, "메세지"
       Exit Sub
    ElseIf objPtDic.Fields("StsCd") = enStsCd.StsCd_LIS_Collection Then
       MsgBox "아직 접수되지 않은 검체입니다.", vbInformation, "메세지"
       Call txtAccSeq_GotFocus
       Exit Sub
    End If

    lblPtId.Caption = objPtDic.Fields("ptid")
    lblPtNm.Caption = objPtDic.Fields("ptnm")
    lblPtSA.Caption = objPtDic.Fields("sexage")
'    lblDept.Caption = objPtDic.Fields("deptcd")
    lblDept.Caption = objPtDic.Fields("deptnm")

    lblWard.Caption = objPtDic.Fields("location")
    lblWardId.Caption = objPtDic.Fields("wardid")
    lblSpecimen.Caption = objPtDic.Fields("spcnm")
    lblMajDoct.Caption = objPtDic.Fields("orddoct")

    lblDoctNm.Caption = objPtDic.Fields("orddrnm")
    lblTelno.Caption = objPtDic.Fields("phone")
    lblDisease.Caption = objPtDic.Fields("mesg")

    fFNSeq = Val(objPtDic.Fields("footnotefg"))
    sRemarkCd = objPtDic.Fields("rmkcd")
    sRemarkIdx = -1
    
    txtDtId.Text = objPtDic.Fields("orddoct")
    txtExDtId.Text = objPtDic.Fields("majdoct")
    strRcvDt = objPtDic.Fields("rcvdt")
    txtTestCd = objPtDic.Fields("testcd")
    
    rtfMessage.Text = ""
    
    blnPtFg = True

    ' footnote Display
    txtFNote.Text = ""
    If fFNSeq > 0 Then txtFNote.Text = objMicRst.DispFootNote(pWorkArea, pAccDt, pAccSeq)

    ' 검체 Remark Display
    sRemarkIdx = medComboFind(cboRemark, sRemarkCd)
    If sRemarkIdx < 0 Then
        cboRemark.ListIndex = 0
    Else
        cboRemark.ListIndex = sRemarkIdx
    End If

    Call ICSPatientMark(lblPtId.Caption, enICSNum.LIS_ALL)
End Sub

Private Sub DispPtInfo_New(ByVal pSpcYY As String, ByVal pSpcNO As String)
    
    Dim sRemarkCd As String, sRemarkIdx As Integer
    Dim objPtDic As clsDictionary
    Dim iWSIndex As Long
    Dim pWorkArea, pAccDt, pAccSeq As String
    
    blnPtFg = False
    
    If optGetList(0) Then
        iWSIndex = cboWSCode.ListIndex
    Else
        iWSIndex = 0
    End If

    pSpcYY = Mid(txtBarNo.Text, 1, 2)
    pSpcNO = Val(Mid(txtBarNo.Text, 3))
    Set objPtDic = objMicRst.DispPtInfoByBarno(pSpcYY, pSpcNO, fWorkSheet(iWSIndex).WsRstType)

    If objPtDic Is Nothing Then
       MsgBox "데이타가 없습니다. 접수번호를 확인하십시오.", vbInformation, "메세지"
       Exit Sub
    ElseIf objPtDic.Fields("StsCd") = enStsCd.StsCd_LIS_Collection Then
       MsgBox "아직 접수되지 않은 검체입니다.", vbInformation, "메세지"
       Call txtAccSeq_GotFocus
       Exit Sub
    End If

    lblPtId.Caption = objPtDic.Fields("ptid")
    lblPtNm.Caption = objPtDic.Fields("ptnm")
    lblPtSA.Caption = objPtDic.Fields("sexage")
'    lblDept.Caption = objPtDic.Fields("deptcd")
    lblDept.Caption = objPtDic.Fields("deptnm")

    lblWard.Caption = objPtDic.Fields("location")
    lblWardId.Caption = objPtDic.Fields("wardid")
    lblSpecimen.Caption = objPtDic.Fields("spcnm")
    lblMajDoct.Caption = objPtDic.Fields("majdoct")

    lblDoctNm.Caption = objPtDic.Fields("orddrnm")
    lblTelno.Caption = objPtDic.Fields("phone")
    lblDisease.Caption = objPtDic.Fields("mesg")

    pWorkArea = objPtDic.Fields("workarea")
    pAccDt = objPtDic.Fields("accdt")
    pAccSeq = objPtDic.Fields("accseq")

    Text1.Text = pWorkArea
    Text2.Text = pAccDt
    Text3.Text = pAccSeq

    txtWorkArea.Text = pWorkArea
    txtAccDt.Text = Mid(pAccDt, 3)
    txtAccSeq.Text = pAccSeq

    fFNSeq = Val(objPtDic.Fields("footnotefg"))
    sRemarkCd = objPtDic.Fields("rmkcd")
    sRemarkIdx = -1

    blnPtFg = True

    ' footnote Display
    txtFNote.Text = ""
    If fFNSeq > 0 Then txtFNote.Text = objMicRst.DispFootNote(pWorkArea, pAccDt, pAccSeq)

    ' 검체 Remark Display
    sRemarkIdx = medComboFind(cboRemark, sRemarkCd)
    If sRemarkIdx < 0 Then
        cboRemark.ListIndex = 0
    Else
        cboRemark.ListIndex = sRemarkIdx
    End If

    Call ICSPatientMark(lblPtId.Caption, enICSNum.LIS_ALL)
End Sub

Private Sub lstBtRCd_DblClick()
    Dim strTestCd   As String   '검사코드
    Dim strRstCd    As String   '결과코드
    Dim strRstNm    As String   '결과명
    
    With ssRst
        strRstCd = Trim$(medGetP(lstBtRCd.List(lstBtRCd.ListIndex), 1, Chr$(9)))
        
        .Row = mvarCurRow: .Col = enSTAIN.tcRSTCD
        .Text = strRstCd
        
        .Col = enSTAIN.tcTESTCD: strTestCd = Trim(.Text)
        Call objMicRst.ResultCheck(strTestCd, strRstCd, strRstNm)
        .Col = enSTAIN.tcRSTNM:  .ForeColor = &HDF6A3E
        .Text = strRstNm
        
        .Col = enSTAIN.tcRSTCD:  .Row = mvarCurRow + 1
        
        If .DataRowCnt < .Row Then
            lstRstCd.Visible = False
            lstBtRCd.Visible = False
            txtFNote.SetFocus
        Else
            .SetFocus
            .Action = ActionActiveCell
        End If
    End With
End Sub

Private Sub lstBtRCd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      lstBtRCd.Visible = False
      lstRstCd.Visible = False
   End If
End Sub

Private Sub lstRstCd_DblClick()
    Dim strTestCd   As String   '검사코드
    Dim strRstCd    As String   '결과코드
    Dim strRstNm    As String   '결과명
    
    With ssRst
        strRstCd = Trim$(medGetP(lstRstCd.List(lstRstCd.ListIndex), 1, Chr$(9)))
        
        .Row = mvarCurRow: .Col = enSTAIN.tcRSTCD
        .Text = strRstCd
        
        .Col = enSTAIN.tcTESTCD: strTestCd = Trim(.Text)
        Call objMicRst.ResultCheck(strTestCd, strRstCd, strRstNm)
        .Col = enSTAIN.tcRSTNM:  .ForeColor = &HDF6A3E
        .Text = strRstNm

        .Col = enSTAIN.tcRSTCD:  .Row = mvarCurRow + 1
        
        If .DataRowCnt < .Row Then
            lstRstCd.Visible = False
            lstBtRCd.Visible = False
            txtFNote.SetFocus
        Else
            .SetFocus
            .Action = ActionActiveCell
        End If
    End With
End Sub

Private Sub lstRstCd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      lstBtRCd.Visible = False
      lstRstCd.Visible = False
   End If
End Sub

Private Sub lstWSUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iListIndex As Integer, iWSIndex As Integer

    'Call ScreenClear

    If optGetList(0) Then
        iWSIndex = cboWSCode.ListIndex
    Else
        iWSIndex = 0
    End If

    iListIndex = lstWSUnit.ListIndex

    Call ClearResult

    If Button = vbLeftButton And iListIndex >= 0 Then
        txtWSUnit.Text = medGetP(lstWSUnit.List(iListIndex), 1, " ")
        Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text)
    End If

    lstWSUnit.Clear
    lstWSUnit.Visible = False

End Sub

Private Sub DisplayData(ByVal pWsCd As String, ByVal pWsUnit As String)

    Dim strBuildDtTm As String, strRcvDtTm As String
    
    Call objMicRst.DispWorksheetInfo(pWsCd, pWsUnit, strBuildDtTm, strRcvDtTm)
    lblBltDate.Caption = strBuildDtTm
    lblRcvDate.Caption = strRcvDtTm

    Call objMicRst.DispWorksheetList(pWsCd, pWsUnit, lstAccList)

End Sub


Private Sub optGetList_Click(Index As Integer)
    fraLabList(0).Visible = IIf(Index = 0, True, False)
    optGetList(0).ForeColor = IIf(Index = 0, vbBlue, vbBlack)
    fraLabList(1).Visible = IIf(Index = 0, False, True)
    optGetList(1).ForeColor = IIf(Index = 0, vbBlack, vbBlue)
    'lstWSUnit.Visible = IIf(Index = 0, True, False)
    DoEvents
    If Index = 1 Then
        lstAccList.Clear
        dtpMidFVfyDt.Value = DateAdd("d", -7, GetSystemDate)
        dtpMidVfyDt.Value = GetSystemDate
        Call dtpMidVfyDt_Change
    Else
        lstFinList.Clear
    End If
End Sub
Private Sub lstAccList_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sTmp As String

    If KeyCode = vbKeyReturn Then

        If lstAccList.ListIndex < 0 Then Exit Sub

        sTmp = medGetP(lstAccList.List(lstAccList.ListIndex), 1, vbTab)
' 2008.12.18 양성현
'        txtWorkArea.Enabled = False: txtAccDt.Enabled = False: txtAccSeq.Enabled = False
        txtWorkArea.Text = medGetP(sTmp, 1, "-"): txtAccDt.Text = medGetP(sTmp, 2, "-"): txtAccSeq.Text = medGetP(sTmp, 3, "-")

' 2008.12.18 양성현
'        fraWSUnit.Enabled = False
'        Call LoadRstData
'        fraWSUnit.Enabled = True
    
        DoEvents
    
        Call txtAccSeq_KeyPress(vbKeyReturn)
    
    End If

End Sub

Private Sub lstFinList_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sTmp As String

    If KeyCode = vbKeyReturn Then

        If lstFinList.ListIndex < 0 Then Exit Sub

        sTmp = medGetP(lstFinList.List(lstFinList.ListIndex), 1, vbTab)
        'txtWorkArea.Enabled = False: txtAccDt.Enabled = False: txtAccSeq.Enabled = False
        txtWorkArea.Text = medGetP(sTmp, 1, "-"): txtAccDt.Text = medGetP(sTmp, 2, "-"): txtAccSeq.Text = medGetP(sTmp, 3, "-")
        'fraWSUnit.Enabled = False
        'Call LoadRstData
        'fraWSUnit.Enabled = True
        DoEvents
        
        Call txtAccSeq_KeyPress(vbKeyReturn)
        
    End If

End Sub

Private Sub lstFinList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then Call lstFinList_KeyDown(13, 0)

End Sub

Private Sub dtpMidVfyDt_Change()
    '** 원본 ----------------------------------------------------------------------------
'    Call objMicCul.GetFinRstList(Format(dtpMidVfyDt.Value, CS_DateDbFormat), lstFinList)
    '------------------------------------------------------------------------------------
    
    '-- 변경 By M.G.Choi 2006.04.05
    Call objMicCul.GetFinRstList_Gra(Format(dtpMidFVfyDt.Value, CS_DateDbFormat), Format(dtpMidVfyDt.Value, CS_DateDbFormat), lstFinList)
    
    If lstFinList.ListCount = 0 Then
        MsgBox "해당일에 중간보고된 결과가 없거나 모두 최종확인되었습니다.", vbInformation, "Stain결과 최종확인"
    End If
End Sub

Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
   If AdvanceNext Then
      'Call ssRst_LeaveCell(6, ssRst.MaxRows, ssRst.Col, ssRst.Row, False)
      Call ssRst_LeaveCell(enSTAIN.tcRSTCD, ssRst.MaxRows, -1, -1, False)
      lstRstCd.Visible = False
      lstBtRCd.Visible = False
      txtFNote.SetFocus
      DoEvents
   End If
End Sub

Private Sub ssRst_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp As Variant
    Dim strTest As String
    Dim strResult As String
    
    With ssRst
        .GetText 1, Row, varTmp: strTest = Trim(varTmp)
        .GetText 6, Row, varTmp: strResult = Trim(varTmp)
    End With
    rtfMessage.Text = rtfMessage.Text & strTest & " : " & strResult & vbCRLF
End Sub

Private Sub ssRst_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal mode As Integer, ByVal ChangeMade As Boolean)
    
    Dim sTestcd As String

    If Col = enSTAIN.tcRSTCD And Row > 0 Then

        ' 새로운 리스트 Load
        ssRst.Col = enSTAIN.tcTESTCD: ssRst.Row = Row: sTestcd = ssRst.Text

        Call objMicRst.LoadStainRstCd(sTestcd, lstBtRCd, lstRstCd)

         If mode = 1 Then
            lstRstCd.Visible = True: lstBtRCd.ZOrder 0
            lstBtRCd.Visible = True: lstRstCd.ZOrder 0
            mvarCurRow = Row
         End If

    End If

    If Col = enSTAIN.tcEXCPT And Row > 0 Then
        lstRstCd.Visible = False
        lstBtRCd.Visible = False
    End If
End Sub

Private Sub ssRst_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyEscape Then

'      ssRst.col = 5: ssRst.Row = ssRst.MaxRows
'      ssRst.Action = ActionActiveCell
'
'      Call ssRst_Advance(True)

      lstRstCd.Visible = False
      lstBtRCd.Visible = False
      DoEvents

   End If

End Sub

Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim i As Integer, sRstCd As String, sRstNm As String, sChk As String, sTestcd As String
    Dim sTmp As String, sExistCd As String, sExistNm As String
    Dim sqlRst As String

    If Col = enSTAIN.tcRSTCD And Row > 0 Then
        ' 현재 리스트에 존재 하는지 check
        ssRst.Col = enSTAIN.tcTESTCD: ssRst.Row = Row: sTestcd = Trim(ssRst.Text)
        ssRst.Col = Col: ssRst.Row = Row: sRstCd = UCase$(Trim(ssRst.Text))

        If Not objMicRst.ResultCheck(sTestcd, sRstCd, sRstNm) Then
            ssRst.Col = Col: ssRst.Row = Row: ssRst.Text = ""
            ssRst.Col = Col + 1: ssRst.Row = Row: ssRst.Text = ""
        Else
            ssRst.Col = Col: ssRst.Row = Row: ssRst.Text = sRstCd
            ssRst.Col = Col + 1: ssRst.Row = Row: ssRst.ForeColor = &HDF6A3E: ssRst.Text = sRstNm
        End If
    End If

    ssRst.Col = NewCol: ssRst.Row = NewRow
    If ssRst.CellType = CellTypeEdit Or ssRst.CellType = CellTypeCheckBox Then
        ssRst.Col = NewCol
    Else
        ssRst.Col = enSTAIN.tcRSTCD
        If ssRst.CellType <> CellTypeEdit Then ssRst.Col = enSTAIN.tcEXCPT
    End If

    ssRst.Action = ActionActiveCell
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim RS          As Recordset
    Dim tmpToolTip  As String
    Dim SSQL        As String
    Dim WorkArea       As String
    Dim AccDt      As String
    Dim AccSeq      As Long
    Dim sTestcd     As String
    
    If Row = 0 Then Exit Sub
    If Row > ssRst.DataRowCnt Then Exit Sub
    
    With ssRst
        WorkArea = Trim(txtWorkArea)
        AccDt = Mid(Now, 1, 2) & Trim(txtAccDt)
        AccSeq = txtAccSeq.TabIndex
        
        .Row = Row: .Col = 2: sTestcd = .Value
        
        SSQL = " SELECT vfydt,lastvfydt,lastvfytm,lastvfyid " & _
               "  FROM " & T_LAB404 & _
               " WHERE " & DBW("workarea=", WorkArea) & _
               "   AND " & DBW("accdt=", AccDt) & _
               "   AND " & DBW("accseq=", AccSeq) & _
               "   AND testcd = " & DBS(sTestcd)
        
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            
            If Not IsNull(RS.Fields("lastvfydt").Value & "") Then
                tmpToolTip = vbCRLF & " 최근 결과일시 : " & Format(RS.Fields("lastvfydt").Value & "", "0###-##-##") & " " & _
                                                     Format(Mid(RS.Fields("lastvfytm").Value & "", 1, 4), "0#:##") & vbCRLF & _
                                        " 결과 보 고 자 : " & GetEmpNm(RS.Fields("lastvfyid").Value & "") & vbCRLF
        
            End If
            
'            Do Until Rs.EOF
'                If Not IsNull(Rs.Fields("lastvfydt").Value & "") Then
'                    tmpToolTip = vbCRLF & " 최근 결과일시 : " & Format(Rs.Fields("lastvfydt").Value & "", "0###-##-##") & " " & _
'                                                         Format(Mid(Rs.Fields("lastvfytm").Value & "", 1, 4), "0#:##") & vbCRLF & _
'                                            " 결과 보 고 자 : " & GetEmpNm(Rs.Fields("lastvfyid").Value & "") & vbCRLF
'
'                End If
'                Rs.MoveNext
'            Loop
            
            MultiLine = 1
            TipText = tmpToolTip
            TipWidth = 5500
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
            
        End If
    End With
    Set RS = Nothing

End Sub

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Or txtBarNo = "" Then Exit Sub

    Call LoadRstData
End Sub

Private Sub txtWorkArea_Change()
    If Not txtAccDt.Enabled Then Exit Sub
    If chkBar.Value = 0 Then
        If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtWorkArea_GotFocus()
    txtWorkArea.SelStart = 0
    txtWorkArea.SelLength = Len(txtWorkArea)
End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr$(KeyAscii)))

    If KeyAscii = vbKeyReturn And Len(txtWorkArea) = txtWorkArea.MaxLength Then txtAccDt.SetFocus

End Sub

Private Sub txtAccDt_Change()
    If Not txtAccSeq.Enabled Then Exit Sub
    If chkBar.Value = 0 Then
        If Len(txtAccDt.Text) = txtAccDt.MaxLength Then txtAccSeq.SetFocus
    End If
End Sub

Private Sub txtAccDt_GotFocus()
    txtAccDt.SelStart = 0
    txtAccDt.SelLength = Len(txtAccDt)
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Len(txtAccDt) >= 2 Then txtAccSeq.SetFocus

    ' 숫자와 백스페이스만 허용
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub txtAccSeq_GotFocus()
    txtAccSeq.SelStart = 0
    txtAccSeq.SelLength = Len(txtAccSeq)
End Sub

Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Or txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then Exit Sub

    Call LoadRstData

End Sub

Private Sub txtWSUnit_KeyPress(KeyAscii As Integer)
    
    Dim iWSIndex As Integer

    If KeyAscii = vbKeyReturn Then

        Call ClearResult

        If optGetList(0) Then
            iWSIndex = cboWSCode.ListIndex
        Else
            iWSIndex = 0
        End If


        If ExistWS(fWorkSheet(iWSIndex).WsCode, txtWSUnit) Then
            Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text)
        Else
            Call ScreenClear
        End If

    End If

End Sub



Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    Dim gintTemplete As Integer
   
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    With clsTemplete
        .Show
        If pintMode = 0 Then
            .lblName.Caption = "Edit " & strTitle
        Else
            .lblName.Caption = "Modify " & strTitle
        End If
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
           Case 1:
              '.lblCode.Caption = objPtInfo.RmkCd
              '.rtfText = rtfRemark.Text
           Case 2:
              '.rtfText = rtfText.Text
           Case 3:
              .rtfText.Text = txtFNote.Text
        End Select
    End With
    gintTemplete = pintPrg
End Sub


