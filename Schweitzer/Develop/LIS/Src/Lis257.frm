VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm257MCultureModify 
   BackColor       =   &H00DBE6E6&
   Caption         =   "감수성 결과수정"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   15345
   Tag             =   "25700"
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
      Left            =   6210
      TabIndex        =   91
      Top             =   1830
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
         Tag             =   "opt"
         Top             =   2190
         Width           =   2325
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   18
         Left            =   180
         TabIndex        =   107
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
         TabIndex        =   108
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
         TabIndex        =   109
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
         TabIndex        =   110
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
         TabIndex        =   111
         Top             =   2970
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis257.frx":0000
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
         TabIndex        =   112
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
         TabIndex        =   113
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
         TabIndex        =   114
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
      Left            =   7770
      Style           =   1  '그래픽
      TabIndex        =   88
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdOrderView 
      BackColor       =   &H00F4F0F2&
      Caption         =   "처방별조회(&C)"
      Height          =   510
      Left            =   9140
      Style           =   1  '그래픽
      TabIndex        =   85
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1845
      Left            =   75
      TabIndex        =   47
      Top             =   -45
      Width           =   8670
      Begin VB.TextBox txtBarcode 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1305
         TabIndex        =   87
         Top             =   180
         Width           =   2535
      End
      Begin VB.CommandButton cmdLisLabel11 
         BackColor       =   &H000080FF&
         Caption         =   "접수 번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaskColor       =   &H80000005&
         Style           =   1  '그래픽
         TabIndex        =   86
         Top             =   180
         Width           =   1150
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
         Left            =   2985
         MaxLength       =   5
         TabIndex        =   51
         Text            =   "10013"
         Top             =   255
         Width           =   630
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
         Left            =   2190
         MaxLength       =   4
         TabIndex        =   50
         Text            =   "9906"
         Top             =   255
         Width           =   540
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
         Height          =   225
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "41"
         Top             =   255
         Width           =   510
      End
      Begin MedControls1.LisLabel lblTestType 
         Height          =   360
         Left            =   3900
         TabIndex        =   52
         Top             =   165
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   635
         BackColor       =   13752531
         ForeColor       =   192
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   2205
         TabIndex        =   53
         Top             =   630
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         BackColor       =   13752531
         ForeColor       =   16711680
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpecimen 
         Height          =   315
         Left            =   6405
         TabIndex        =   54
         Top             =   1005
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
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
      End
      Begin MedControls1.LisLabel lblDept 
         Height          =   315
         Left            =   6405
         TabIndex        =   55
         Top             =   630
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         BackColor       =   13752531
         ForeColor       =   16711680
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
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   975
         TabIndex        =   56
         Top             =   630
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         BackColor       =   13752531
         ForeColor       =   16711680
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtSA 
         Height          =   315
         Left            =   4545
         TabIndex        =   57
         Top             =   630
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   13752531
         ForeColor       =   16711680
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWard 
         Height          =   315
         Left            =   975
         TabIndex        =   58
         Top             =   1005
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
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
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   315
         Left            =   975
         TabIndex        =   59
         Top             =   1365
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
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
      End
      Begin MedControls1.LisLabel lblTelno 
         Height          =   315
         Left            =   3690
         TabIndex        =   71
         Top             =   1005
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
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
         Index           =   1
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   615
         Width           =   810
         _ExtentX        =   1429
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
         Index           =   2
         Left            =   120
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   990
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "병   동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1365
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "진단명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2850
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1005
         Width           =   810
         _ExtentX        =   1429
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   5535
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   630
         Width           =   810
         _ExtentX        =   1429
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
         Index           =   7
         Left            =   5550
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1005
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "검 체"
         Appearance      =   0
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
         Left            =   3855
         TabIndex        =   48
         Top             =   180
         Visible         =   0   'False
         Width           =   4110
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   300
         Left            =   6405
         TabIndex        =   89
         Top             =   1380
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   5550
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1380
         Width           =   810
         _ExtentX        =   1429
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
      Begin VB.Label lblStainRst 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "GramStain Result"
         Height          =   330
         Left            =   6885
         TabIndex        =   63
         Tag             =   "25607"
         Top             =   180
         Width           =   1635
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
         Left            =   2760
         TabIndex        =   62
         Top             =   270
         Width           =   195
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
         Left            =   1950
         TabIndex        =   61
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblSenType 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Label6"
         Height          =   285
         Left            =   5760
         TabIndex        =   60
         Top             =   210
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F1F5F4&
         BackStyle       =   1  '투명하지 않음
         Height          =   360
         Left            =   1305
         Top             =   180
         Width           =   2535
      End
   End
   Begin VB.Frame fraNogr 
      BackColor       =   &H00DBE6E6&
      Height          =   1335
      Left            =   75
      TabIndex        =   6
      Top             =   1725
      Width           =   4620
      Begin VB.ComboBox cboNGRst 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         ItemData        =   "Lis257.frx":009D
         Left            =   1920
         List            =   "Lis257.frx":00B6
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   765
         Width           =   2325
      End
      Begin MedControls1.LisLabel lblNogrowth 
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   315
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "3 day(s) Nogrowth"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   315
         Width           =   1755
         _ExtentX        =   3096
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
         Caption         =   "Nogrowth History"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   750
         Width           =   1755
         _ExtentX        =   3096
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
         Caption         =   "Nogrowth Result"
         Appearance      =   0
      End
   End
   Begin DRcontrol1.DrFrame fraOldSensi 
      Height          =   5940
      Left            =   1605
      TabIndex        =   64
      Top             =   3120
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   10478
      Title           =   "결과보고일 : "
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
      Begin VB.ListBox lstlastAcc 
         Appearance      =   0  '평면
         BackColor       =   &H00E4F3F8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5250
         Left            =   135
         TabIndex        =   67
         Top             =   450
         Width           =   2145
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F2FBFB&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10035
         Style           =   1  '그래픽
         TabIndex        =   66
         Tag             =   "0"
         Top             =   90
         Width           =   270
      End
      Begin VB.CommandButton cmdUpDown 
         BackColor       =   &H00F2FBFB&
         Caption         =   "▲"
         Height          =   255
         Left            =   9750
         Style           =   1  '그래픽
         TabIndex        =   65
         Tag             =   "0"
         Top             =   90
         Width           =   270
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   4260
         Left            =   2295
         TabIndex        =   68
         Top             =   450
         Width           =   8055
         _Version        =   196608
         _ExtentX        =   14208
         _ExtentY        =   7514
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
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
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis257.frx":0130
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblVfyDt 
         Height          =   300
         Left            =   2295
         TabIndex        =   69
         Top             =   60
         Width           =   8040
         _ExtentX        =   14182
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
         Caption         =   "결과보고일 : "
         Appearance      =   0
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Left            =   135
         TabIndex        =   70
         Top             =   60
         Width           =   2145
         _ExtentX        =   3784
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
      Begin RichTextLib.RichTextBox txtSamCmt 
         Height          =   960
         Left            =   2310
         TabIndex        =   84
         Top             =   4740
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   1693
         _Version        =   393217
         BackColor       =   15728382
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis257.frx":08E1
         MouseIcon       =   "Lis257.frx":0986
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
   Begin VB.CommandButton cmdGetOldResult 
      BackColor       =   &H00EDE2ED&
      Caption         =   "최근 감수성결과"
      Height          =   510
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   44
      Top             =   8535
      Width           =   1515
   End
   Begin VB.TextBox txtVfyDt 
      Height          =   300
      Left            =   4410
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   8730
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox picESign 
      Height          =   500
      Left            =   3060
      ScaleHeight     =   435
      ScaleWidth      =   1140
      TabIndex        =   41
      Top             =   8595
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   40
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame fraSusc 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   75
      TabIndex        =   16
      Tag             =   "25615"
      Top             =   2955
      Width           =   14400
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         BackColor       =   &H00F5F2FF&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   180
         Left            =   11145
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "☞ Red : 감수성 Italic : 농도"
         Top             =   765
         Width           =   2490
      End
      Begin VB.CommandButton cmdDefAnti 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   13785
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   1905
         Width           =   360
      End
      Begin VB.CommandButton cmdDefAnti 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   13785
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   1095
         Width           =   360
      End
      Begin VB.CommandButton cmdDefAnti 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   13785
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   2715
         Width           =   360
      End
      Begin VB.CommandButton cmdDefAnti 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   13785
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   3540
         Width           =   360
      End
      Begin VB.CommandButton cmdDefAnti 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   13785
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   4350
         Width           =   360
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   825
         Left            =   105
         TabIndex        =   22
         Top             =   1095
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   1455
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "1"
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   810
         Left            =   105
         TabIndex        =   23
         Top             =   1920
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   1429
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "2"
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   810
         Left            =   105
         TabIndex        =   24
         Top             =   2730
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   1429
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "3"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   810
         Index           =   0
         Left            =   105
         TabIndex        =   25
         Top             =   3540
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   1429
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "4"
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   825
         Left            =   105
         TabIndex        =   26
         Top             =   4350
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   1455
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "5"
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   435
         Left            =   450
         TabIndex        =   27
         Top             =   630
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   767
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "미 생 물"
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   435
         Left            =   3330
         TabIndex        =   28
         Top             =   630
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   767
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "정 도"
      End
      Begin MedControls1.LisLabel LisLabel10 
         Height          =   435
         Left            =   4620
         TabIndex        =   29
         Top             =   630
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   767
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "적 용 항 생 제"
      End
      Begin FPSpread.vaSpread ssSusc 
         Height          =   4305
         Left            =   450
         TabIndex        =   30
         Tag             =   "25616"
         Top             =   1095
         Width           =   13305
         _Version        =   196608
         _ExtentX        =   23469
         _ExtentY        =   7594
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         MaxCols         =   32
         MaxRows         =   15
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   1
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis257.frx":0AE8
         UserResize      =   0
         VisibleCols     =   2
         VisibleRows     =   2
         TextTip         =   2
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   435
         Left            =   2955
         TabIndex        =   31
         Top             =   630
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BackColor       =   16118527
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "구분"
      End
      Begin MedControls1.LisLabel LisLabel15 
         Height          =   375
         Left            =   450
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   210
         Width           =   2010
         _ExtentX        =   3545
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
         Caption         =   "◈ 항균제 감수성 결과"
         Appearance      =   0
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "WARNING"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5655
         TabIndex        =   46
         ToolTipText     =   "새로운 미생물이 발견되었습니다."
         Top             =   330
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "( 결과코드 : R,I,S,P,N  없음 : - )"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2505
         TabIndex        =   33
         Top             =   315
         Width           =   3060
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "☞ Esc : 균코드/정도코드 숨김"
         Height          =   180
         Index           =   0
         Left            =   11130
         TabIndex        =   32
         Top             =   375
         Width           =   2520
      End
      Begin VB.Shape shpWarning 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   5580
         Shape           =   4  '둥근 사각형
         Top             =   225
         Visible         =   0   'False
         Width           =   1380
      End
   End
   Begin VB.PictureBox picMIC 
      Appearance      =   0  '평면
      BackColor       =   &H00B88FA5&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2025
      ScaleHeight     =   3855
      ScaleWidth      =   2820
      TabIndex        =   36
      Top             =   4125
      Visible         =   0   'False
      Width           =   2820
      Begin VB.ListBox lstMIC 
         BackColor       =   &H00F7FDFD&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   3765
         Index           =   1
         ItemData        =   "Lis257.frx":6571
         Left            =   945
         List            =   "Lis257.frx":65AE
         TabIndex        =   39
         Top             =   45
         Width           =   915
      End
      Begin VB.ListBox lstMIC 
         BackColor       =   &H00FFFCF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C76456&
         Height          =   3765
         Index           =   0
         ItemData        =   "Lis257.frx":6605
         Left            =   30
         List            =   "Lis257.frx":6642
         TabIndex        =   38
         Top             =   45
         Width           =   915
      End
      Begin VB.ListBox lstMIC 
         BackColor       =   &H00F8F8FE&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   3765
         Index           =   2
         ItemData        =   "Lis257.frx":66BF
         Left            =   1860
         List            =   "Lis257.frx":66FC
         TabIndex        =   37
         Top             =   45
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00F4F0F2&
      Caption         =   "결과 수정 (&S)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "25611"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "25612"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.ListBox lstMicNm 
      Appearance      =   0  '평면
      BackColor       =   &H00F8FADE&
      Height          =   3450
      Left            =   11310
      TabIndex        =   11
      Top             =   5430
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.ListBox lstQty 
      Appearance      =   0  '평면
      BackColor       =   &H00EEFFEE&
      Height          =   3450
      Left            =   10785
      TabIndex        =   9
      Top             =   5430
      Visible         =   0   'False
      Width           =   3675
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
      Left            =   6255
      Picture         =   "Lis257.frx":6779
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   1830
      Width           =   315
   End
   Begin VB.ComboBox cboRemark 
      BackColor       =   &H00F1F5F4&
      Height          =   300
      Left            =   6615
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   2640
      Width           =   3255
   End
   Begin VB.ListBox lstGramStain 
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      ItemData        =   "Lis257.frx":6CAB
      Left            =   8775
      List            =   "Lis257.frx":6CB2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   390
      Width           =   5670
   End
   Begin VB.ListBox lstMicCd 
      Appearance      =   0  '평면
      BackColor       =   &H00F8FADE&
      Height          =   3450
      Left            =   10050
      TabIndex        =   10
      Top             =   5415
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstMGroup 
      BackColor       =   &H00FFF2EC&
      Height          =   3120
      Left            =   8745
      TabIndex        =   8
      Top             =   5805
      Visible         =   0   'False
      Width           =   1290
   End
   Begin RichTextLib.RichTextBox txtPNote 
      Height          =   765
      Left            =   6615
      TabIndex        =   5
      Top             =   1845
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1349
      _Version        =   393217
      BackColor       =   15658734
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Lis257.frx":6CBB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtFNote 
      Height          =   765
      Left            =   10380
      TabIndex        =   1
      Top             =   1845
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1349
      _Version        =   393217
      BackColor       =   15857140
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Lis257.frx":6D60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MedControls1.LisLabel lblRemark 
      Height          =   285
      Left            =   9870
      TabIndex        =   13
      Top             =   2640
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   503
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
      Alignment       =   1
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel14 
      Height          =   315
      Left            =   8760
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   45
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "◈ GramStain Result"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel12 
      Height          =   315
      Left            =   4770
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1830
      Width           =   1470
      _ExtentX        =   2593
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
      Caption         =   "◈ Foot Note"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel13 
      Height          =   285
      Left            =   4770
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   "◈ 검체 Remark"
      Appearance      =   0
   End
   Begin VB.Label lblSpcCd 
      AutoSize        =   -1  'True
      Caption         =   "검체"
      Height          =   195
      Left            =   13230
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblWardId 
      AutoSize        =   -1  'True
      Caption         =   "WardId"
      Height          =   195
      Left            =   13725
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblMajDoct 
      AutoSize        =   -1  'True
      Caption         =   "주치의"
      Height          =   195
      Left            =   13725
      TabIndex        =   34
      Top             =   195
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frm257MCultureModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsTemplete As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents frmMic As frmMicOption
Attribute frmMic.VB_VarHelpID = -1

Private fTestCd As String                   ' 결과 작업 중인 검사항목
Private fMfySeq As Long                     ' 결과 작업 중인 항목의 수정 횟수
Private fFNSeq As Integer                   ' 결과 작업 중인 FootNote Seq. Number
Private fMicFg As String                    ' MIC 감수성 여부
Private fCurMic As Integer

Private fSkColor As Long                    ' 입력 가능 셀 배경색
Private fOkColor As Long                    ' 입력 불가 셀 배경색

Private aryNGCD() As Variant                   ' Nogrowth 결과 코드
'Private fSSRow As Integer                   ' 현재 에디팅 중인 라인
'Private fSSCol As Integer                   ' 현재 에디팅 중인 컬럼
Private fPrevCode As String                 ' 균명이 바뀌었는지 체크

Private blnPtFg As Boolean
Private blnMsgFg As Boolean
Private blnSendKeys As Boolean
Private bEsign   As Boolean

Private fVfyDt As String                    '보고일자

Private objMicSql As New clsLISSqlMicRst
Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture
Private objMicLib As New clsLISMicroLib

Private blnForce As Boolean
Private cmdLisLabel11fg As Boolean

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset
Dim strRcvDt            As String

'Private PreRstSens As String    '스프레드에 입력하기 전 결과(Sensi결과)
'Private PreRstMedi As String    '스프레드에 입력하기 전 결과(Medi결과)

Public Sub ApplyDefAnti(ByVal pRow As Integer, ByVal pCnt As Integer, ByVal pBuf As String)

    Dim sTmp As String
    Dim i As Integer

    ssSusc.Col = 7: ssSusc.Row = pRow: ssSusc.Text = pCnt

    For i = 8 To (7 + pCnt)

        sTmp = medShift(pBuf, ";")

        ssSusc.Col = i: ssSusc.Row = pRow - 1
        ssSusc.Text = medGetP(sTmp, 1, ":")
        
        ssSusc.Col = i: ssSusc.Row = pRow
        ssSusc.CellType = CellTypeEdit
        ssSusc.TypeEditCharCase = TypeEditCharCaseSetUpper
        ssSusc.TypeHAlign = TypeHAlignCenter
        ssSusc.BackColor = fOkColor
        ssSusc.Text = medGetP(sTmp, 2, ":")
        
    Next i

    For i = pCnt + 8 To ssSusc.MaxCols
        ssSusc.Col = i: ssSusc.Row = pRow - 1
        ssSusc.Text = ""

        ssSusc.Col = i: ssSusc.Row = pRow
        ssSusc.CellType = CellTypeStaticText
        ssSusc.BackColor = fSkColor
        ssSusc.Text = ""
    Next i

End Sub

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
End Sub

Private Sub cmdLisLabel11_Click()
    Call ClearForm
If cmdLisLabel11fg Then
    cmdLisLabel11.Caption = "접수 번호"
'    Shape1.Visible = True
    txtWorkArea.Visible = True
    txtAccDt.Visible = True: txtAccDt.Enabled = True
    txtAccSeq.Visible = True
    
    txtWorkArea.Text = "04"
    txtAccDt.Text = ""
    txtAccSeq.Text = ""
    Label4.Visible = True
    Label5.Visible = True
    cmdLisLabel11fg = False
    txtBarcode.Visible = False
    txtWorkArea.SetFocus
Else
    cmdLisLabel11.Caption = "검체 번호"
'    Shape1.Visible = False
    txtWorkArea.Visible = False
    txtAccDt.Visible = False: txtAccDt.Enabled = False
    txtAccSeq.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    cmdLisLabel11fg = True
    txtBarcode.Visible = True
    txtBarcode.Text = ""
    txtBarcode.SetFocus

End If
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

Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
    Dim strLabNo As String
    Dim strSpcYY As String
    Dim strSpcNo As String
' txtWorkArea.Text , txtAccDt.Text, txtAccSeq.Text
    If KeyAscii = vbKeyReturn Then
        If txtBarcode.Text = "" Then Exit Sub
         Dim blnAccFg As Boolean
'            Call ClearForm
            strSpcYY = Mid(txtBarcode.Text, 1, 2)
            strSpcNo = Mid(txtBarcode.Text, 3)
            strLabNo = objMicLib.GetWADNoByBarCode(strSpcYY, strSpcNo, txtWorkArea.Text, txtAccDt.Text, txtAccSeq.Text)
            txtWorkArea.Text = medGetP(strLabNo, 1, "-")
            txtAccDt.Text = Mid(medGetP(strLabNo, 2, "-"), 3)
            txtAccSeq.Text = medGetP(strLabNo, 3, "-")
            If strLabNo <> "" Then
                Call txtAccSeq_KeyPress(999)
            Else
                MsgBox " 미생물 검체 정보가 아닙니다. "
                txtBarcode.Text = ""
            End If
        cmdLisLabel11fg = False
'        Call cmdLisLabel11_Click
    End If
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
Private Sub cboNGRst_Click()
    
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer

    If txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then
        fraSusc.Enabled = False
    Else
        If aryNGCD(cboNGRst.ListIndex) = "" Then
            fraSusc.Enabled = True
        Else
            ssSusc.Col = 2: ssSusc.Row = 2
            If ssSusc.Text <> "" Then
                sMsg = "현재 저장되어 있는 감수성 결과를 모두 무시해도 좋습니까?"
                sStyle = vbYesNo + vbCritical + vbDefaultButton2
                sRes = MsgBox(sMsg, sStyle, "Nogrowth 결과 입력 확인")
                If sRes = vbYes Then
                    Call ClearSuscTable
                    fraSusc.Enabled = False: txtFNote.SetFocus
                Else
                    cboNGRst.ListIndex = 0
                End If
            Else
                Call ClearSuscTable
                fraSusc.Enabled = False: txtFNote.SetFocus
            End If
        End If
    End If

End Sub

Private Sub cboNGRst_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub cboNGRst_LostFocus()
    
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdModify.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If cboNGRst.ListIndex = -1 Then Exit Sub
    If txtWorkArea = "" Or txtAccDt = "" Or txtAccSeq = "" Then
        fraSusc.Enabled = False: cboNGRst.ListIndex = 0: If Not cmdLisLabel11fg Then txtWorkArea.SetFocus
    Else
        If aryNGCD(cboNGRst.ListIndex) = "" Then
            fraSusc.Enabled = True
            objMicLib.CRow = ssSusc.DataRowCnt + 2
            ssSusc.Col = 1: ssSusc.Row = objMicLib.CRow: ssSusc.Action = ActionActiveCell
            DoEvents: ssSusc.SetFocus
        Else
            fraSusc.Enabled = False: txtFNote.SetFocus
        End If
    End If

End Sub

Private Sub cboRemark_Click()
    
    Dim iIndex As Integer, sRMCd As String, sRMNm As String
    iIndex = cboRemark.ListIndex
    If iIndex < 0 Then Exit Sub
    sRMCd = Trim(Mid(cboRemark.List(iIndex), 1, 6))
    If sRMCd = LIS_Nothing Then lblRemark.Caption = "": Exit Sub
    lblRemark.Caption = objMicRst.GetRemark(sRMCd)
End Sub

Private Sub clsTemplete_CopyTemplete()
    txtFNote.Text = clsTemplete.rtfText.Text
    txtFNote.SetFocus
    Set clsTemplete = Nothing
End Sub

Private Sub cmdClear_Click()

    Call ClearForm

    'txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
    txtWorkArea = MIC_WorkArea: txtAccDt = "": txtAccSeq = ""
    txtWorkArea.Locked = False: txtAccDt.Locked = False: txtAccSeq.Locked = False
    cmdLisLabel11fg = False
    txtBarcode.Visible = False
    txtBarcode.Text = ""
    
    cmdLisLabel11fg = False

    cmdLisLabel11.Caption = "접수 번호"
    txtWorkArea.Visible = True
    txtAccDt.Visible = True: txtAccDt.Enabled = True
    txtAccSeq.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    cmdLisLabel11fg = False
    txtBarcode.Visible = False
    
    txtAccDt.SetFocus
    cmdOrderView.Visible = False

End Sub

Private Sub cmdClose_Click()
    fraOldSensi.Visible = False
End Sub

Private Sub cmdCommentTemplete_Click()

   If ssSusc.MaxRows < 1 Then Exit Sub
   Call CallTemplete(3, 0)

End Sub

Private Sub cmdDefAnti_Click(Index As Integer)
    
    Dim sMicNm As String, sAntiCnt As Integer, sBuf As String
    Dim tmpMicCd As String, tmpRst As String, tmpMic As String

    ssSusc.Col = 1: ssSusc.Row = (Index + 1) * 3 - 1: sMicNm = ssSusc.Text

    If sMicNm = "" Then Exit Sub

    ssSusc.Col = 7: ssSusc.Row = (Index + 1) * 3 - 1: sAntiCnt = Val(ssSusc.Text)

    Dim i As Integer
    sBuf = ""
    For i = 8 To 7 + sAntiCnt
        ssSusc.Col = i: ssSusc.Row = ((Index + 1) * 3) - 2: tmpMic = ssSusc.Text
        ssSusc.Col = i: ssSusc.Row = ((Index + 1) * 3) - 1: tmpRst = ssSusc.Text
        'ssSusc.Col = i: ssSusc.Row = ((Index + 1) * 3): tmpMicCd = ssSusc.Text
        'sBuf = sBuf & tmpMicCd & ":" & tmpRst & ":" & tmpMic & ";"
        sBuf = sBuf & tmpMic & ":" & tmpRst & ";"
    Next i

    Call frm260MDefAnti.SetCurAnti(Me, Index, sMicNm, sAntiCnt, sBuf)
    frm260MDefAnti.Show 1

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frm257MCultureModify = Nothing
End Sub

Private Sub cmdGetOldResult_Click()
    Dim strLabNo    As String
    Dim aryTmp()    As String
    Dim ii          As Integer
    
    
    strLabNo = objMicLib.GetAccNoOfLatestRst(, txtWorkArea.Text & Mid(Format(GetSystemDate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text)
    

    lstlastAcc.Clear

    
    If Trim(strLabNo) = "" Then
        MsgBox "해당 환자의 " & lblSpecimen.Caption & " 검체에 대한 최근 감수성검사 내역이 없습니다.", vbInformation
        fraOldSensi.Visible = False
        Exit Sub
    Else
        aryTmp() = Split(strLabNo, COL_DIV)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            lstlastAcc.AddItem aryTmp(ii)
        Next
    End If
    
    lstlastAcc.ListIndex = 0
    Call lstlastAcc_KeyDown(13, 0)
    
'    Dim strLabNo As String
'    Dim sWorkArea As String
'    Dim sAccDt As String
'    Dim sAccSeq As String
'    Dim i As Long, j As Long, k As Long
'
'    tblResult.Row = -1
'    tblResult.Col = -1
'    tblResult.BlockMode = True
'    tblResult.Action = ActionClearText
'    tblResult.BlockMode = False
'    lblVfyDt.Caption = ""
'
''    objMicLib.AccNo = txtWorkArea.Text & Mid(Format(GetSystemdate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text
''    strLabNo = objMicCul.GetLatestSensiAccNo(lblPtid.Caption, _
''                        Format(GetSystemdate, CS_DateDbFormat), lblSpcCd.Caption)
'    strLabNo = objMicLib.GetAccNoOfLatestRst(, txtWorkArea.Text & Mid(Format(GetSystemdate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text)
'
'    If Trim(strLabNo) = "" Then
'        MsgBox "해당 환자의 " & lblSpecimen.Caption & " 검체에 대한 최근 감수성검사 내역이 없습니다.", vbInformation
'        fraOldSensi.Visible = False
'        Exit Sub
'    End If
'
'    sWorkArea = medGetP(strLabNo, 1, "-")
'    sAccDt = medGetP(strLabNo, 2, "-")
'    sAccSeq = medGetP(strLabNo, 3, "-")
'
'    lblVfyDt.Caption = "최근 감수성결과 보고일시 : " & Format(medGetP(strLabNo, 4, "-"), CS_DateLongMask)
'    lblVfyDt.Caption = lblVfyDt.Caption & "  " & Format(medGetP(strLabNo, 5, "-"), CS_TimeLongMask)
'    lblVfyDt.Caption = lblVfyDt.Caption & String(2, " ") & "접수번호 : " & sWorkArea & "-" & Mid(sAccDt, 3) & "-" & sAccSeq
'
'    Dim MyResult As New clsLISResultReview
'
'    With MyResult
'
'        MouseRunning
'
'        Call .MicrobeSensiRst(sWorkArea, sAccDt, sAccSeq)
'
'        If .ResultCnt = 0 Then
'            MouseDefault
'            Exit Sub
'        End If
'
'        For i = 1 To .RstRow
'            tblResult.Row = i + .OffSet
'            For j = 1 To 8
'                tblResult.Col = j
'                If .Get_ForeColor(j, i) <> 0 Then tblResult.ForeColor = .Get_ForeColor(j, i)
'            Next
'        Next
'
'        '결과내역 Display
'        tblResult.Row = 1
'        tblResult.Row2 = tblResult.MaxRows
'        tblResult.Col = 2
'        tblResult.COL2 = tblResult.MaxCols
'        tblResult.BlockMode = True
'        tblResult.AllowCellOverflow = True
'        tblResult.Clip = .ResultClipText    '& .SenClipText             'ResultBuffer
'        tblResult.BlockMode = False
'
'        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
'        'If .SortFg Then
'        If .SortFg Then
''            tblResult.SortBy = SortByRow
''            tblResult.SortKey(1) = 2  '항생제명
''            tblResult.SortKeyOrder(1) = SortKeyOrderAscending
''            tblResult.Col = -1
''            tblResult.Row = .SortStartRow   '+ .OffSet
''            tblResult.Row2 = .SortEndRow  '+ .OffSet
''            tblResult.Action = ActionSort
''            tblResult.Row = .SortStartRow - 1   '+ .OffSet
''            tblResult.Col = 2
''            tblResult.FontUnderline = True
'        Else
'            tblResult.Col = 6
'            tblResult.Row = -1
'            tblResult.ForeColor = DCM_LightRed
'            tblResult.FontBold = True
'        End If
'
'        '미생물 결과 : 균명컬럼 Align Left
'        tblResult.Row = -1
'        tblResult.Col = -1
'        tblResult.BlockMode = True
'        tblResult.AllowCellOverflow = True
'        tblResult.TypeHAlign = TypeHAlignLeft
'        tblResult.BlockMode = False
'        tblResult.ColWidth(2) = 10
'        'tblResult.ColWidth(3) = 60
'        For i = 1 To 5
'            If .MicFg(i) Then
'                tblResult.ColWidth(i + 2) = 9
'            Else
'                tblResult.ColWidth(i + 2) = 4
'            End If
'        Next
'        tblResult.ColWidth(8) = 20
'        tblResult.Col = 3: tblResult.COL2 = 7
'        tblResult.Row = -1
'        tblResult.BlockMode = True
'        tblResult.FontBold = False
'        tblResult.BlockMode = False
'
'    End With
'
'    Dim strAntiCd As String
'    Dim strAntiRst As String
'    Dim lngMicCnt As Long
'    Dim blnFind As Boolean
'
'    With ssSusc
'        For i = 1 To ssSusc.DataRowCnt Step 3
'            .Row = i + 1
'            .Col = 1
'            If .Value <> "" Then
'                .Col = 7
'                lngMicCnt = Val(.Value)
'                For k = 8 To lngMicCnt + 8
'                    .Row = i
'                    .Col = k: strAntiCd = .Value
'                    .Row = i + 1
'                    .Col = k: strAntiRst = .Value
'
'                    tblResult.Row = MyResult.SortStartRow - 1
'                    tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
'                    tblResult.Value = "현재" & CStr(i \ 3 + 1)
'                    tblResult.ForeColor = DCM_LightRed
'
'                    blnFind = False
'                    For j = MyResult.SortStartRow To tblResult.DataRowCnt
'                        tblResult.Row = j
'                        tblResult.Col = 2
'                        If tblResult.Value = strAntiCd Then
'                            tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
'                            tblResult.Value = strAntiRst
'                            tblResult.ColWidth(MyResult.MicrobeCount + (i \ 3) + 3) = 4
'                            If strAntiRst Like "S*" Then
'                                tblResult.ForeColor = DCM_Red
'                            Else
'                                tblResult.ForeColor = DCM_Gray
'                            End If
'                            blnFind = True
'                            Exit For
'                        End If
'                    Next
'                    If Not blnFind Then
'                        tblResult.Row = tblResult.DataRowCnt + 1
'                        tblResult.Col = 2
'                        tblResult.Value = strAntiCd
'                        tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
'                        tblResult.Value = strAntiRst
'                        If strAntiRst Like "S*" Then
'                            tblResult.ForeColor = DCM_Red
'                        Else
'                            tblResult.ForeColor = DCM_Gray
'                        End If
'                    End If
'                Next
'            End If
'        Next
'    End With
'
'    MouseDefault
'
'    fraOldSensi.Visible = True
'    fraOldSensi.ZOrder 0
'    Set MyResult = Nothing
End Sub

Private Sub cmdModify_Click()
    
    Dim sSysDate As String, sDate As String, sTime As String
    Dim sTmp1 As String, sTmp2 As String, sTmp3 As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim sRmkCd As String
    Dim blnSave As Boolean
    Dim tmpDeptCd As String, tmpBussDiv As String
    Dim strSpcNm  As String
    Dim iCnt As Integer
    Dim chkNull As Boolean

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)

    If sWorkArea = "" Or sAccDt = "" Or sAccSeq = "" Then
        MsgBox "Accession Number가 정확하지 않습니다. 확인후 처리 하세요"
        Exit Sub
    End If

    sDate = Format(GetSystemDate, CS_DateDbFormat)
    sTime = Format(GetSystemDate, CS_TimeDbFormat)

    If cboRemark.ListIndex > 0 Then sRmkCd = Trim(Mid(cboRemark.Text, 1, 6))

'-------------------------------------------------------------------------------------------
    '전자서명 Validation Check
    Dim objESign        As clsLISElectronSign
    
    If P_MicElectronicSign And aryNGCD(cboNGRst.ListIndex) = "" Then
        Set objESign = New clsLISElectronSign
        If objESign.LoadElectronSign(ObjMyUser.EmpId, InstallDir & "LIS") = False Then
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
    Set objESign = Nothing
'-------------------------------------------------------------------------------------------

On Error GoTo DBExecError

    DBConn.BeginTrans   '시작

    If aryNGCD(cboNGRst.ListIndex) = "" Then       ' 감수성 결과 중간 등록
        chkNull = True
        For iCnt = 1 To ssSusc.MaxRows
            ssSusc.Col = 2: ssSusc.Row = iCnt: sTmp1 = Trim(ssSusc.Text)
            ssSusc.Col = 5: ssSusc.Row = iCnt: sTmp2 = Trim(ssSusc.Text)
            If Len(sTmp1) > 0 Or Len(sTmp2) > 0 Then
                chkNull = chkNull And False
            End If
        Next
        'If sTmp1 = "" Or sTmp2 = "" Then
        If chkNull = True Then
            MsgBox "감수성 결과 입력이 잘못되었습니다. 확인 후 처리 하세요", vbExclamation, "감수성 결과"
            DBConn.RollbackTrans
            Exit Sub
        Else
            If objMicCul.SaveGRResult(sWorkArea, sAccDt, sAccSeq, fTestCd, "", enStsCd.StsCd_LIS_Modify, _
                                        ObjSysInfo.EmpId, txtFNote.Text, sRmkCd, fMfySeq + 1) = False Then GoTo DBExecError
                                        
            '** 원본 ============================================================================================
            'If objMicCul.SaveSenResult(ssSusc, sWorkArea, sAccDt, sAccSeq, fTestCd, fMfySeq + 1) = False Then GoTo DBExecError
            '====================================================================================================
            
            '** 전주예수병원 추가루틴 ===========================================================================
            'If objMicCul.SaveSenResult_New3(ssSusc, sWorkArea, sAccDt, sAccSeq, fTestCd, ObjSysInfo.EmpId, txtFNote.Text, sRmkCd, enStsCd.StsCd_LIS_Modify, fMfySeq + 1) = False Then GoTo DBExecError
            '====================================================================================================
            
            '** 전주예수병원 추가루틴 ===========================================================================
            ' * 검체명 추가로 인한 변경
            strSpcNm = Trim(lblSpecimen.Caption)
            
            If objMicCul.SaveSenResult_New4(ssSusc, sWorkArea, sAccDt, sAccSeq, fTestCd, ObjSysInfo.EmpId, txtFNote.Text, sRmkCd, enStsCd.StsCd_LIS_Modify, fMfySeq + 1, strSpcNm) = False Then GoTo DBExecError
            '====================================================================================================
        End If
    Else                                        ' Nogrowth 결과 중간 등록
        '-----------------------------------------------------------------------
        '2001-12-03 수정 : Nogrowth 결과입력 시 MIC Status도 반영
        '-----------------------------------------------------------------------
        Dim i As Long
        fTestCd = ""
        For i = 0 To lstTest.ListCount - 1
            If i = 0 Then
                fTestCd = "'" & medGetP(lstTest.List(i), 1, vbTab) & "'"
            Else
                fTestCd = fTestCd & ",'" & medGetP(lstTest.List(i), 1, vbTab) & "'"
            End If
        Next
        '-----------------------------------------------------------------------
        If objMicCul.SaveNGResult(sWorkArea, sAccDt, sAccSeq, fTestCd, _
                                    aryNGCD(cboNGRst.ListIndex), enStsCd.StsCd_LIS_Modify, _
                                    ObjSysInfo.EmpId, txtFNote.Text, sRmkCd, fMfySeq + 1) = False Then GoTo DBExecError
    End If

    tmpBussDiv = objMicRst.Get_Bussdiv(sWorkArea, sAccDt, sAccSeq)
    If tmpBussDiv = "" Then
        If Trim(lblWardId.Caption) = "" Then
            tmpDeptCd = lblDept.Caption
            tmpBussDiv = enBussDiv.BussDiv_OutPatient
        Else
            tmpDeptCd = lblWardId.Caption
            tmpBussDiv = enBussDiv.BussDiv_InPatient
        End If
    Else
        If tmpBussDiv = "1" Then
            tmpDeptCd = lblDept.Caption
        ElseIf tmpBussDiv = "2" Then
            tmpDeptCd = lblWardId.Caption
        
        End If
        
    End If
        
    
    blnSave = objMicRst.SubmitVerifyList(tmpDeptCd, sDate, sTime, lblPtId.Caption, enStsCd.StsCd_LIS_Modify, ObjMyUser.EmpId, lblMajDoct.Caption, tmpBussDiv)
    
    If Not blnSave Then GoTo DBExecError
    
    DBConn.CommitTrans  '끝
    
    '감염관리
    Call ICSSensiResultCheck(ssSusc, lblPtId.Caption, lblPtNm.Caption, sWorkArea, sAccDt, sAccSeq, _
                             fTestCd, lblWard.Caption, lblDept.Caption, True)
   
    Call cmdClear_Click

    Exit Sub

DBExecError:
    'DbConn.RollbackTrans
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Sub

Private Sub cmdPrint_Click()
    If Trim(txtWorkArea.Text) <> "" And Trim(txtAccDt.Text) <> "" And Trim(txtAccSeq.Text) <> "" Then
'        With ssSusc
'            If .DataRowCnt < 1 Then
'                MsgBox "출력할 내용이 없습니다!", vbCritical, "확인"
'                Exit Sub
'            End If
'        End With
        
        Me.MousePointer = vbHourglass
        
        Call FinalResult_Print
        
        Me.MousePointer = vbDefault
    Else
        MsgBox "접수번호를 확인하세요!", vbCritical, "확인"
    End If
End Sub

Private Sub FinalResult_Print()
    Dim strRstEntryType As String
    Dim strPtID         As String
    Dim strTestDiv      As String
    Dim strVfyDt        As String
    Dim strTable        As String
    Dim strSQL          As String
    Dim strImgPath      As String
    Dim strStartDate    As String
    Dim strRptFg        As String
    Dim strRptDt        As String
    Dim strRptTm        As String
    Dim strRptID        As String
    Dim strAccDt        As String
    Dim bFeedBackFg     As Boolean
    Dim bTemp           As Boolean
    Dim i               As Long
    Dim j               As Long
    Dim objProgress     As clsProgress
    Dim objProgress1    As clsProgress
    Dim objReport       As clsBatchReport
    Dim objSQL          As New clsLISSqlReport
    Dim objSqlTmp       As New clsLISSqlReview
    Dim strLastDt       As String
    Dim strLastTm       As String
    Dim strPrtDt        As String
    Dim strPrtTm        As String
    Dim rsTmp           As Recordset
    Dim RS              As Recordset
    Dim lngErrCount     As Long
    Dim lngFileNo       As Long

    Dim strFDate As String
    Dim strTDate As String
    

    strPtID = lblPtId.Caption
    strVfyDt = txtVfyDt.Text
    
    '새로 추가
    
    strFDate = txtVfyDt.Text
    strTDate = txtVfyDt.Text
    
    strTestDiv = "2"
    picESign.Picture = LoadPicture(strImgPath)
    strStartDate = Format(GetSystemDate, "yyyymmdd")

    Set objProgress = New clsProgress
    Set objReport = New clsBatchReport

    objReport.PtId = lblPtId.Caption
    objReport.ptnm = lblPtNm.Caption
    objReport.Ward = medGetP(lblWardId.Caption, 1, "-")
    objReport.PtSex = medGetP(lblPtSA.Caption, 1, "/")
    objReport.PtAge = medGetP(lblPtSA.Caption, 2, "/")
    objReport.VfyDt = Format(GetSystemDate, "yyyy/mm/dd")

    objReport.Doct = GetEmpNm(lblMajDoct.Caption)

'    If ObjLISComCode.WardId.Exists(medGetP(lblWard.Caption, 1, "-")) = True Then
'        ObjLISComCode.WardId.KeyChange (medGetP(lblWard.Caption, 1, "-"))
'        objReport.Ward = ObjLISComCode.WardId.Fields("wardnm")
'
'        If objReport.Ward <> "" Then
'            objReport.Ward = objReport.Ward & " " & Mid(lblWard.Caption, Len(medGetP(lblWard.Caption, 1, "-")) + 2)
'        Else
'            objReport.Ward = Mid(lblWard.Caption, Len(medGetP(lblWard.Caption, 1, "-")) + 2)
'        End If
'    End If

    '임상진단....
    Dim objDisease  As New clsDisease

    objDisease.PtId = lblPtId.Caption

    objReport.ICD = objDisease.Disease

    Set objDisease = Nothing
    
    '-- 성바오로병원 결과지 출력
    strAccDt = Trim(txtAccDt.Text)
    strAccDt = IIf(Mid(strAccDt, 1, 1) = "9", "19" & strAccDt, "20" & strAccDt)
    
    '-- 복수검체일 경우는 최종 접수번호로 출력한다. By M.G.Choi
    ' - 복수검체 판단여부 : 처방Body(ORD102)테이블에 접수번호 없을 경우 => 복수검체
    strSQL = " SELECT workarea, accdt, accseq,ptid,orddt,ordno,ordseq FROM " & T_LAB102 & _
             "  WHERE workarea = '" & Trim(txtWorkArea.Text) & "'" & _
             "    AND accdt = '" & strAccDt & "'" & _
             "    AND accseq = '" & Trim(txtAccSeq) & "'"
             
    Set rsTmp = New Recordset
    rsTmp.Open strSQL, DBConn
    
    If rsTmp.EOF = False Then
        bTemp = False
    Else
        bTemp = True
    End If
    
    If bTemp = True Then
        MsgBox "선택한 항목은 복수검체 입니다!" & Chr(13) & _
               "최종 접수번호로만 출력이 가능 합니다.", vbCritical, "확인"
        txtAccSeq.Text = ""
        txtAccSeq.SetFocus
        Set rsTmp = Nothing
        Set objSQL = Nothing
        Set objSqlTmp = Nothing
        Exit Sub
    End If
    
    rsTmp.MoveFirst
    Do Until rsTmp.EOF
    
        strSQL = " SELECT min(vfydt) as FROMdt,max(vfydt) as todt FROM " & T_LAB404 & _
                 " WHERE " & DBW("ptid=", rsTmp.Fields("ptid").Value & "") & _
                 " AND " & DBW("orddt=", rsTmp.Fields("orddt").Value & "") & _
                 " AND " & DBW("ordno=", rsTmp.Fields("ordno").Value & "") & _
                 " AND " & DBW("ordseq=", rsTmp.Fields("ordseq").Value & "")
               
               
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open strSQL, DBConn
        
        
        If Not RS.EOF Then
            If strFDate > RS.Fields("FROMdt").Value & "" Or strFDate = "" Then strFDate = RS.Fields("FROMdt").Value & ""
            If strTDate < RS.Fields("todt").Value & "" Or strTDate = "" Then strTDate = RS.Fields("todt").Value & ""
        End If
        Set RS = Nothing
        rsTmp.MoveNext
    Loop
           
    strSQL = objSqlTmp.SqlGet_LAB404(Trim(txtWorkArea.Text), strAccDt, _
                                     Trim(txtAccSeq.Text), fTestCd)
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    bFeedBackFg = False
    If RS.EOF = False Then
        strRptFg = RS.Fields("rptfg").Value & ""
        strRptDt = RS.Fields("rptdt").Value & ""
        strRptTm = RS.Fields("rpttm").Value & ""
        strRptID = RS.Fields("rptid").Value & ""
        
        bFeedBackFg = True
    End If
    
    Set RS = Nothing

    objReport.Reprint = True
    Call objReport.ReportForOnePatient(strPtID, strFDate, strTDate, _
                                       strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
'    Call objReport.ReportForOnePatient(strPtid, strVfyDt, strVfyDt, _
'                                       strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
    '-- FeedBack
    If bFeedBackFg = True Then
        Call PrintReUpdate(Trim(txtWorkArea.Text), Trim(txtAccDt.Text), Trim(txtAccSeq.Text), fTestCd, _
                           strRptFg, strRptDt, strRptTm, strPtID)
    End If
    
    Set objSQL = Nothing
    Set objSqlTmp = Nothing
End Sub

Private Function PrintReUpdate(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal TestCd As String, _
                               ByVal pRptFg As String, ByVal pRptDt As String, ByVal pRptTm As String, ByVal pRptID As String) As Boolean
    Dim strTblNm As String
    Dim SSQL     As String
    
'    SELECT Case testdiv
'        Case "0": strTblNm = T_LAB302
'        Case "1": strTblNm = T_LAB351
'        Case "2": strTblNm = T_LAB404
'    End SELECT
    
    On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    
    SSQL = " UPDATE " & T_LAB404 & _
                " set " & _
                        DBW("rptfg", pRptFg, 3) & _
                        DBW("rptdt", pRptDt, 3) & _
                        DBW("rpttm", pRptTm, 3) & _
                        DBW("rptid", pRptID, 2) & _
                " WHERE " & _
                        DBW("WORKAREA=", WorkArea) & " AND " & _
                        DBW("ACCDT=", AccDt) & " AND " & _
                        DBW("ACCSEQ=", AccSeq) & " AND " & _
                        DBW("TESTCD=", TestCd)
                
    DBConn.Execute SSQL
    DBConn.CommitTrans
    PrintReUpdate = True
    Exit Function
SAVE_ERROR:
    DBConn.RollbackTrans
    PrintReUpdate = False
End Function

Private Sub cmdUpDown_Click()
    
    If cmdUpDown.Tag = "0" Then
        cmdUpDown.Tag = "1"
        cmdUpDown.Caption = "▼"
        fraOldSensi.Height = 465
        blnForce = True
    Else
        cmdUpDown.Tag = "0"
        cmdUpDown.Caption = "▲"
        fraOldSensi.Height = 5925
        blnForce = False
    End If

End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()

    ssSusc.Col = 1: ssSusc.Row = 1: fSkColor = ssSusc.BackColor
    ssSusc.Col = 1: ssSusc.Row = 2: fOkColor = ssSusc.BackColor

    With objMicRst
        Call .LoadNGRstCd("GC", cboNGRst, aryNGCD)  'LoadNGCode
        Call .LoadRemark(cboRemark)
    End With
    With objMicCul
        Call .LoadMicrobe(lstMicCd, lstMicNm, lstMGroup)
        Call .LoadQuantity(lstQty)
    End With
    
    Me.Show
    DoEvents
    
    txtWorkArea = MIC_WorkArea: txtAccDt = "": txtAccSeq = ""
    'txtWorkArea = "": txtAccDt = "": txtAccSeq = ""
    'txtWorkArea.SetFocus
    txtAccDt.SetFocus
    ClearForm
    
    frmSMS.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objMicSql = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
    Set objMicLib = Nothing
End Sub

Private Sub fraOldSensi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnForce Then Exit Sub
    If cmdUpDown.Tag = "1" Then
        cmdUpDown.Tag = "0"
        cmdUpDown.Caption = "▲"
        fraOldSensi.Height = 5925
    End If
End Sub

Private Sub lblVfyDt_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    fraOldSensi.Move X, Y
End Sub

Private Sub lblVfyDt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If blnForce Then Exit Sub
    If cmdUpDown.Tag = "1" Then
        cmdUpDown.Tag = "0"
        cmdUpDown.Caption = "▲"
        fraOldSensi.Height = 5925
    End If

End Sub

Private Sub lstMIC_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn                      'Enter Key 또는 Space
            Call lstMIC_MouseDown(Index, 1, 0, 0, 0)
        Case vbKeyEscape    '  27                      'ESC
            picMIC.Visible = False
            DoEvents: ssSusc.SetFocus
        Case vbKeyRight
            lstMIC((Index + 1) Mod 3).SetFocus
        Case vbKeyLeft
            lstMIC((Index + 2) Mod 3).SetFocus
    End Select

End Sub

Private Sub lstMIC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    ssSusc.Row = objMicLib.CRow
    ssSusc.Col = objMicLib.CCol
    ssSusc.Text = lstMIC(Index).Text
    If blnSendKeys Then SendKeys "{Enter}"
End Sub

Private Sub lstTest_KeyPress(KeyAscii As Integer)
    Dim iX As Single, iY As Single
    If KeyAscii = vbKeyReturn Then
        Call lstTest_MouseDown(1, 0, iX, iY)
    End If
End Sub

Private Sub lstTest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 1 Then Exit Sub
    
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String
    
    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)

    fTestCd = medGetP(lstTest.Text, 1, vbTab)
    lblTestType.Caption = medGetP(lstTest.Text, 2, vbTab)
    lstTest.Visible = False
    
    If DispPtInfo(sWorkArea, sAccDt, sAccSeq, fTestCd) Then
        Call objMicCul.DispStainResult(sWorkArea, sAccDt, sAccSeq, lstGramStain)
        Call DispGrowthRst(sWorkArea, sAccDt, sAccSeq, fTestCd)
        fraSusc.Enabled = True
        cboNGRst.SetFocus
        Exit Sub
    End If

End Sub

Private Sub ClearForm()

    objMicLib.CRow = -1:   objMicLib.CCol = -1: fPrevCode = "": fFNSeq = 0: fCurMic = -1
    fTestCd = "": lblTestType.Caption = "": lblSenType = "": fMicFg = ""

    lblPtId.Caption = "": lblPtNm.Caption = "": lblPtSA.Caption = ""
    lblDept.Caption = "": lblWard.Caption = "": lblSpecimen.Caption = "": lblDisease.Caption = ""
    txtPNote.Text = "": txtFNote.Text = ""
    lblStainRst = ""
    lblSpcCd.Caption = ""
    lblDoctNm.Caption = ""
    lstGramStain.Clear
    cboNGRst.ListIndex = 0: lblNogrowth.Caption = ""
    cboRemark.ListIndex = 0: lblRemark.Caption = ""

    picMIC.Visible = False
    lstTest.Clear: lstTest.Visible = False
    lstMicCd.Visible = False: lstMicNm.Visible = False
    fraSusc.Enabled = False
    Call ClearSuscTable
    
    With objMicLib
        .WorkArea = ""
        .AccDt = ""
        .AccSeq = ""
        .PtId = ""
        .TestCd = ""
        .SpcCd = ""
        .PreRstMedi = ""
        .PreRstSens = ""
    End With
    
    txtBarcode.Visible = False
    txtBarcode.Text = ""
    
    cmdOrderView.Visible = False

End Sub

Private Sub ClearSuscTable()

    Dim i As Integer, j As Integer

    With ssSusc
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    
        For i = 2 To .MaxRows Step 3
            For j = 8 To .MaxCols
                'SRI결과입력부분
                .Col = j: .Row = i: .Text = ""
                .CellType = CellTypeEdit
                .TypeEditCharCase = TypeEditCharCaseSetUpper
                .TypeMaxEditLen = 1
                .TypeHAlign = TypeHAlignCenter
                .BackColor = fOkColor
                'MIC결과입력부분
                .Col = j: .Row = i + 1: .Text = ""
                .CellType = CellTypeEdit
                .TypeMaxEditLen = 999
                .TypeHAlign = TypeHAlignCenter
                .BackColor = fOkColor
            Next j
        Next i
    End With

End Sub


Private Sub frmMic_MicSELECT(ByVal strMicFg As String)

    Dim sMicCd As String
    Dim sMicFg As String
    
    If objMicLib.CRow = -1 Then Exit Sub
    
    ssSusc.Row = objMicLib.CRow
    ssSusc.Col = 2: sMicCd = ssSusc.Value
    ssSusc.Col = 4: sMicFg = ssSusc.Value
    
    If fPrevCode <> sMicCd Or sMicFg <> strMicFg Then
        ssSusc.Text = strMicFg
        fMicFg = strMicFg
        Call ShowAntiList
    End If
    
    ssSusc.Row = objMicLib.CRow
    ssSusc.Col = 5
    ssSusc.Action = ActionActiveCell

'    Dim sWorkArea As String, sAccDt As String, sAccSeq As String, sRstTp As String
'
'    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtWorkArea): sAccSeq = Trim(txtAccSeq)
'    sAccDt = IIf(Mid$(sAccDt, 1, 1) = "9", "19", "20") & sAccDt
'
'    If strMicFg = MRT_MicSen Then fMicFg = ""
'    If DispPtInfo(sWorkArea, sAccDt, sAccSeq, strMicFg) Then
'        lblStainRst.Caption = objMicCul.DispStainResult(sWorkArea, sAccDt, sAccSeq, lstGramStain)
'        Call DispGrowthRst(sWorkArea, sAccDt, sAccSeq, strMicFg)
'        fraSusc.Enabled = True
'        cboNGRst.SetFocus
'        Exit Sub
'    End If



End Sub

Private Sub lstMicCd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn                      'Enter Key 또는 Space
            Call lstMicNM_MouseDown(1, 0, 0, 0)
        Case vbKeyEscape    '  27                      'ESC
            lstMicCd.Visible = False: lstMicNm.Visible = False
            DoEvents: ssSusc.SetFocus
    End Select
End Sub

Private Sub lstMicCd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And lstMicCd.ListIndex >= 0 Then
        Call SelMicList(lstMicCd.ListIndex)
    End If
    DoEvents: ssSusc.SetFocus
    If blnSendKeys Then SendKeys "{Enter}"
End Sub

Private Sub lstMicCd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      lstMicCd.Visible = False
      lstMicNm.Visible = False
   End If
End Sub

Private Sub lstMicCd_Scroll()
    lstMicNm.TopIndex = lstMicCd.TopIndex
End Sub

Private Sub lstMicNm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      lstMicCd.Visible = False
      lstMicNm.Visible = False
   End If
End Sub

Private Sub lstMicNm_Scroll()
    lstMicCd.TopIndex = lstMicNm.TopIndex
End Sub

Private Sub lstMicNM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And lstMicNm.ListIndex >= 0 Then
        Call SelMicList(lstMicNm.ListIndex)
    End If
    DoEvents: ssSusc.SetFocus
    If blnSendKeys Then SendKeys "{Enter}"
End Sub

Private Sub SelMicList(ByVal pIdx As Integer)
    

    Dim sMicCd As String, sMicNm As String

    sMicCd = lstMicCd.List(pIdx)
    sMicNm = lstMicNm.List(pIdx)

    If fPrevCode = sMicCd Then Exit Sub
    If objMicLib.CRow = -1 Then Exit Sub
    
    If sMicCd = LIS_Nothing Then
        Call ClearSuscRow(objMicLib.CRow, True)
        Exit Sub
    End If
    
    With ssSusc
        .Row = objMicLib.CRow
        .Col = 1: .Text = sMicNm    '균이름
        .Col = 2: .Text = sMicCd    '균코드
        .Col = 3: .Text = lstMGroup.List(pIdx) '균종
    End With
    
    Call objMicLib.GetWarningForMnm(ssSusc)
End Sub

'Private Sub GetWarningForMnm()
'    Dim RS As recordset
'    Dim blnWarning As Boolean
'    Dim strMnmCd As String
'
'    If fSSRow = -1 Then Exit Sub
'
'    If ChkInPatient = False Then
'        blnWarning = False
'    Else
'        Set RS = GetLastestRst
'
'        If RS Is Nothing Then
'            blnWarning = False
'        Else
'            Do Until RS.EOF
'                strMnmCd = strMnmCd & RS.Fields("mnmcd").Value & "" & COL_DIV
'
'                RS.MoveNext
'            Loop
'            blnWarning = True
'        End If
'    End If
'
'    ssSusc.Col = 1: ssSusc.Row = fSSRow
'    shpWarning.Visible = False
'    lblWarning.Visible = False
'    ssSusc.FontItalic = False
'    ssSusc.FontBold = False
'
'    If blnWarning Then
'        Dim varMnmCd As Variant
'        Dim varMnmNm As Variant
'        Dim i As Long
'        Call ssSusc.GetText(2, fSSRow, varMnmCd)
'        Call ssSusc.GetText(1, fSSRow, varMnmNm)
'
'        If InStr(strMnmCd, varMnmCd) = 0 Then
'            shpWarning.Visible = True
'            lblWarning.Visible = True
'            ssSusc.FontItalic = True
'            ssSusc.FontBold = True
'        End If
'    End If
'
'    Set RS = Nothing
'End Sub

'Private Function ChkInPatient() As Boolean
'    Dim RS As recordset
'    Dim strSQL As String
'
'    strSQL = " SELECT c." & F_INPTID & " as ptid, c." & F_BEDINDT & " as bedindt FROM " & T_HIS002 & " c " & _
'             " WHERE c." & F_INPTID & "=(SELECT d.ptid FROM " & T_LAB201 & " d WHERE " & DBW("d.workarea =", txtWorkArea.Text) & " AND " & DBW("d.accdt =", Mid(Format(GetSystemdate, "yyyyMMdd"), 1, 2) & txtAccDt.Text) & " AND " & DBW("d.accseq=", txtAccSeq.Text) & ")" & _
'             " AND c." & F_BEDINDT & " = (SELECT max(b." & F_BEDINDT & ") FROM " & T_HIS002 & " b WHERE b." & F_INPTID & "=c." & F_INPTID & ") " & _
'             " AND (c." & F_BEDOUTDT & "='' or c." & F_BEDOUTDT & " is null)"
'
'    Set RS = OpenRecordSet(strSQL)
'
'    If RS.EOF Then
'        ChkInPatient = False
'    Else
'        ChkInPatient = True
'    End If
'
'    Set RS = Nothing
'End Function

'Private Function GetLastestRst() As recordset
'    Dim strSQL As String
'    Dim strAccNo As String
'
'    objMicLib.AccNo = txtWorkArea.Text & Mid(Format(GetSystemdate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text
'    strAccNo = objMicCul.GetLatestSensiAccNo(lblPtId.Caption, _
'                        Format(GetSystemdate, CS_DateDbFormat), lblSpcCd.Caption)
'
'    If strAccNo <> "" Then
'        strSQL = " SELECT * FROM " & T_LAB405 & _
'                  " WHERE " & DBW("workarea=", medGetP(strAccNo, 1, "-")) & _
'                  " AND " & DBW("accdt=", medGetP(strAccNo, 2, "-")) & _
'                  " AND " & DBW("accseq=", medGetP(strAccNo, 3, "-"))
'
'        Set GetLastestRst = OpenRecordSet(strSQL)
'    End If
'End Function

Private Sub ShowAntiList()

    Dim sAntiList As String, iCnt As Integer
    Dim sMicGrp As String
    Dim sMicCd As String
    Dim sMicFg As String
    Dim sSenType As String
    
    If objMicLib.CRow = -1 Then Exit Sub
    
    ssSusc.Row = objMicLib.CRow
    ssSusc.Col = 2
    sMicCd = ssSusc.Text
    ssSusc.Col = 3
    sMicFg = ssSusc.Text
    
    If fPrevCode = sMicCd And fMicFg = sMicFg Then Exit Sub

'2009.10.06 양성현 고객(김정애)의 요청으로 ....
'    Call ClearSuscRow(objMicLib.CRow, False)

    With ssSusc
    
        .Row = objMicLib.CRow

        .Col = 3: sMicGrp = .Text   '균종
        .Col = 4: sSenType = .Text  'MIC여부
        
        iCnt = objMicCul.LoadAntibiotic(sSenType, sMicGrp, sAntiList)

        .Col = 7: .Value = iCnt     '항생제갯수
        
        .Col = 8: .COL2 = iCnt + 7
        .Row = objMicLib.CRow - 1: .Row2 = objMicLib.CRow - 1
        .BlockMode = True
        .Clip = sAntiList
        .BlockMode = False
        
        .Col = iCnt + 8: .COL2 = .MaxCols
        .Row = objMicLib.CRow: .Row2 = objMicLib.CRow
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BackColor = fSkColor
        .BlockMode = False

        If sSenType = MRT_MicSen Then
            .Col = iCnt + 8
        Else
            .Col = 8
        End If
        .COL2 = .MaxCols
        .Row = objMicLib.CRow + 1: .Row2 = objMicLib.CRow + 1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .BackColor = fSkColor
        .BlockMode = False

    End With

    lstMicCd.ListIndex = -1: lstMicCd.Visible = False
    lstMicNm.ListIndex = -1: lstMicNm.Visible = False

End Sub


Private Sub ClearSuscRow(ByVal pRow As Integer, ByVal blnAll As Boolean)
    
    Dim i As Integer
    
    With ssSusc
        If blnAll Then
            .Col = 1:
        Else
            .Col = 5
        End If
        .COL2 = .MaxCols
        .Row = pRow - 1: .Row2 = pRow + 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        
        .Col = 8: .COL2 = .MaxCols
        .Row = pRow: .Row2 = pRow
        .BlockMode = True
        .CellType = CellTypeEdit
        .TypeEditCharCase = TypeEditCharCaseSetUpper
        .TypeHAlign = TypeHAlignCenter
        .ForeColor = &HC6614F
        .BackColor = fOkColor
        .BlockMode = False
    
        .Col = 8: .COL2 = .MaxCols
        .Row = pRow + 1: .Row2 = pRow + 1
        .BlockMode = True
        .CellType = CellTypeEdit
        .TypeEditCharCase = TypeEditCharCaseSetUpper
        .TypeHAlign = TypeHAlignCenter
        .ForeColor = &H6A6FA6
        .BackColor = fOkColor
        .BlockMode = False
    
    End With
    
End Sub

Private Sub lstQty_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyReturn                      'Enter Key 또는 Space
            Call lstQty_MouseDown(1, 0, 0, 0)
        Case vbKeyEscape                      'ESC
            lstQty.Visible = False
            DoEvents: ssSusc.SetFocus
    End Select

End Sub

Private Sub lstQty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'배양된 균의 정도코드 선택

    Dim i As Integer
    Dim sTmp As String, sCode As String, sName As String

    If Button = 1 And lstQty.ListIndex >= 0 Then
        If objMicLib.CRow = -1 Then Exit Sub
        
        sTmp = lstQty.List(lstQty.ListIndex)
        sCode = medGetP(sTmp, 1, vbTab)
        sName = medGetP(sTmp, 2, vbTab)

        ssSusc.Row = objMicLib.CRow: ssSusc.Col = 5: ssSusc.Value = sName
        ssSusc.Row = objMicLib.CRow: ssSusc.Col = 6: ssSusc.Value = sCode

        lstQty.ListIndex = -1: lstQty.Visible = False

        Call objMicLib.GetWarningForQty(ssSusc)
    End If

    DoEvents: ssSusc.SetFocus
    If blnSendKeys Then SendKeys "{Enter}"

End Sub

Private Sub lstQty_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      lstQty.Visible = False
   End If
End Sub

'Private Sub GetWarningForQty()
''접수번호를 구해서 구한걸로 405를 뒤져서 같은 균이 있는지 찾어..
''그 균의 정도를 찾어..
'    Dim RS As recordset
'    Dim blnWarning As Boolean
'    Dim strMnmCd As String
'
'    If fSSRow = -1 Then Exit Sub
'
'    If ChkInPatient = False Then
'        blnWarning = False
'    Else
'        Set RS = GetLastestRst
'
'        If RS Is Nothing Then
'            blnWarning = False
'        Else
'            Do Until RS.EOF
'                strMnmCd = strMnmCd & RS.Fields("mnmcd").Value & "" & COL_DIV & RS.Fields("mqtcd").Value & "" & LINE_DIV
'
'                RS.MoveNext
'            Loop
'            blnWarning = True
'        End If
'    End If
'
'    ssSusc.Col = 5: ssSusc.Row = fSSRow
'    shpWarning.Visible = False
'    lblWarning.Visible = False
'    ssSusc.FontBold = False
'    ssSusc.FontItalic = False
'
'    If blnWarning Then
'        Dim varMnmCd As Variant
'        Dim varMqyCd As Variant
'        Dim i As Long
'        Dim aryRec() As String
'        Dim aryFld() As String
'
'        Call ssSusc.GetText(2, fSSRow, varMnmCd)
'        Call ssSusc.GetText(6, fSSRow, varMqyCd)
'
'        If InStr(strMnmCd, varMnmCd) > 0 Then
'            aryRec = Split(strMnmCd, LINE_DIV)
'
'            For i = LBound(aryRec) To UBound(aryRec) - 1
'                aryFld = Split(aryRec(i), COL_DIV)
'                If aryFld(0) = varMnmCd Then
'                    If aryFld(1) <> varMqyCd Then
'                        shpWarning.Visible = True
'                        lblWarning.Visible = True
'                        ssSusc.FontBold = True
'                        ssSusc.FontItalic = True
'                        Exit For 'i
'                    End If
'                End If
'            Next i
'        End If
'    End If
'
'    Set RS = Nothing
'End Sub

Private Sub ssSusc_Click(ByVal Col As Long, ByVal Row As Long)
    If Col > 7 Then
        objMicLib.PreRstSens = "": objMicLib.PreRstMedi = ""
        
        ssSusc.Row = Col: ssSusc.Row = Row
        
        If (Row Mod 3) = 2 Then
            objMicLib.PreRstSens = ssSusc.Value
        ElseIf (Row Mod 3) = 0 Then
            objMicLib.PreRstMedi = ssSusc.Value
        End If
    End If
End Sub

Private Sub ssSusc_GotFocus()
    If ssSusc.Col = -1 Or ssSusc.Row = -1 Then Exit Sub
    Call ssSusc_LeaveCell(0, 0, ssSusc.Col, ssSusc.Row, False)
End Sub

Private Sub ssSusc_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iListIdx As Integer

    If ssSusc.Col = 1 And lstMicCd.Visible = True Then

        iListIdx = lstMicCd.ListIndex

        Select Case KeyCode
            Case vbKeyDown ', vbKeyPageDown
                If lstMicCd.ListCount - 1 > iListIdx Then
                    lstMicCd.ListIndex = iListIdx + 1
                    lstMicNm.ListIndex = iListIdx + 1
                End If
                KeyCode = 0
            Case vbKeyPageDown
                If lstMicCd.ListCount - 18 > iListIdx Then
                    lstMicCd.ListIndex = iListIdx + 18
                    lstMicNm.ListIndex = iListIdx + 18
                Else
                    lstMicCd.ListIndex = lstMicCd.ListCount - 1
                    lstMicNm.ListIndex = lstMicCd.ListCount - 1
                End If
                KeyCode = 0
            Case vbKeyUp
                If iListIdx > 0 Then
                    lstMicCd.ListIndex = iListIdx - 1
                    lstMicNm.ListIndex = iListIdx - 1
                End If
                KeyCode = 0
            Case vbKeyPageUp
                If iListIdx - 18 > 0 Then
                    lstMicCd.ListIndex = iListIdx - 18
                    lstMicNm.ListIndex = iListIdx - 18
                Else
                    lstMicCd.ListIndex = 0
                    lstMicNm.ListIndex = 0
                End If
                KeyCode = 0
        End Select
        DoEvents

    End If

    If ssSusc.Col = 5 And lstQty.Visible = True Then

        iListIdx = lstQty.ListIndex

        Select Case KeyCode
            Case vbKeyDown, vbKeyPageDown
                If lstQty.ListCount - 1 > iListIdx Then lstQty.ListIndex = iListIdx + 1
                KeyCode = 0
            Case vbKeyUp, vbKeyPageUp
                If iListIdx > 0 Then lstQty.ListIndex = iListIdx - 1
                KeyCode = 0
        End Select
        DoEvents
    End If
    
    If ssSusc.Col > 7 Then
        If (ssSusc.Row Mod 3) = 0 And fCurMic >= 0 Then
            iListIdx = lstMIC(fCurMic).ListIndex
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown
                    If lstMIC(fCurMic).ListCount - 1 > iListIdx Then lstMIC(fCurMic).ListIndex = iListIdx + 1
                    KeyCode = 0
                Case vbKeyUp, vbKeyPageUp
                    If iListIdx > 0 Then lstMIC(fCurMic).ListIndex = iListIdx - 1
                    KeyCode = 0
                Case vbKeyLeft
                    lstMIC(fCurMic).ListIndex = -1
                    fCurMic = (fCurMic + 2) Mod 3
                    lstMIC(fCurMic).ListIndex = iListIdx
                Case vbKeyRight
                    lstMIC(fCurMic).ListIndex = -1
                    fCurMic = (fCurMic + 1) Mod 3
                    lstMIC(fCurMic).ListIndex = iListIdx
            End Select
            DoEvents
        Else
            Select Case KeyCode
                Case vbKeyUp
                    If ssSusc.Row > 3 Then
                        ssSusc.Row = ssSusc.Row - 3
                        ssSusc.Action = ActionActiveCell
                    End If
            End Select
        End If
    End If

End Sub

Private Sub ssSusc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        lstMicCd.Visible = False
        lstMicNm.Visible = False
        lstQty.Visible = False
        picMIC.Visible = False
    End If

    If KeyAscii = vbKeyReturn Then
        If objMicLib.CRow = -1 Then Exit Sub
        
        blnSendKeys = False
        
        Select Case ssSusc.Col
        
            Case 1:
                If lstMicCd.ListIndex >= 0 And lstMicCd.Visible Then
                    Call lstMicCd_MouseDown(1, 0, 0, 0)
                Else                    ' 여기가 이상하다 없는 인덱스에는 동작 안하다니..
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 1: ssSusc.Text = ""
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 2: ssSusc.Text = ""
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 1: ssSusc.Action = ActionActiveCell
                    KeyAscii = 0
                    DoEvents: ssSusc.SetFocus
                End If
                lstMicCd.Visible = False: lstMicNm.Visible = False
                DoEvents
            
            Case 5:
                If lstQty.ListIndex >= 0 And lstQty.Visible Then
                    Call lstQty_MouseDown(1, 0, 0, 0)
                Else                    ' 여기가 이상하다 없는 인덱스에는 동작 안하다니..
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 5: ssSusc.Text = ""
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 6: ssSusc.Text = ""
                    ssSusc.Row = objMicLib.CRow: ssSusc.Col = 5: ssSusc.Action = ActionActiveCell
                    KeyAscii = 0
                    DoEvents: ssSusc.SetFocus
                End If
                lstQty.Visible = False
                DoEvents
        
            Case Is > 7:
                If (objMicLib.CRow Mod 3) = 0 Then
                    If lstMIC(fCurMic).ListIndex >= 0 And lstMIC(fCurMic).Visible Then
                        Call lstMIC_MouseDown(fCurMic, 1, 0, 0, 0)
                    Else                    ' 여기가 이상하다 없는 인덱스에는 동작 안하다니..
                        ssSusc.Row = objMicLib.CRow: ssSusc.Col = objMicLib.CCol: ssSusc.Text = ""
                        ssSusc.Row = objMicLib.CRow: ssSusc.Col = objMicLib.CCol: ssSusc.Action = ActionActiveCell
                        KeyAscii = 0
                        DoEvents: ssSusc.SetFocus
                    End If
                    picMIC.Visible = False
                    DoEvents
                End If
        End Select
        
        blnSendKeys = True
    Else 'If (KeyAscii = Asc("R")) Or (KeyAscii = Asc("I")) Or (KeyAscii = Asc("S")) Or (KeyAscii = Asc("P")) Or (KeyAscii = Asc("N")) Or (KeyAscii = Asc("-")) Then
        If (ssSusc.Col > 7) Then
            objMicLib.PreRstSens = "": objMicLib.PreRstMedi = ""

            ssSusc.Row = objMicLib.CRow: ssSusc.Row = objMicLib.CCol

            If (objMicLib.CRow Mod 3) = 2 Then
                objMicLib.PreRstSens = ssSusc.Value
            ElseIf (objMicLib.CRow Mod 3) = 0 Then
                objMicLib.PreRstMedi = ssSusc.Value
            End If
        End If
    End If

End Sub

Private Sub ssSusc_EditChange(ByVal Col As Long, ByVal Row As Long)

    Dim sCurIdx As Integer
    Dim sIdxCd As Integer, sIdxNm As Integer, sIdxQty As Integer
    Dim sMicCd As String, sMicGrp As String

    Select Case Col

        Case 1
            With ssSusc
                .Col = Col: .Row = Row
                sIdxCd = medListFind(lstMicCd, .Value)
                sIdxNm = medListFind(lstMicNm, .Value)

                If sIdxCd >= 0 Then                         ' 코드에서 같은 문자
                    sCurIdx = sIdxCd
                ElseIf sIdxCd = -1 And sIdxNm >= 0 Then     ' 이름에서 같은 문자
                    sCurIdx = sIdxNm
                Else                                        ' 같은 문자 없슴
                    sCurIdx = lstMicCd.ListIndex
                End If

                medLockWindowUpdate lstMicCd.hwnd
                lstMicCd.ListIndex = sCurIdx
                medLockWindowUpdate 0&
                
                medLockWindowUpdate lstMicNm.hwnd
                lstMicNm.ListIndex = sCurIdx
                medLockWindowUpdate 0&
            End With

        Case 4  'MIC여부
            Call ShowAntiList
            SendKeys "{ENTER}"
        
        Case 5  '정도코드
            With ssSusc
                .Col = Col: .Row = Row
                sIdxQty = medListFind(lstQty, .Value)

                If sIdxQty >= 0 Then                         ' 코드에서 같은 문자
                    sCurIdx = sIdxQty
                Else                                         ' 같은 문자 없슴
                    sCurIdx = lstQty.ListIndex
                End If

                lstQty.ListIndex = sCurIdx

            End With

         Case Is > 7
            
            With ssSusc
                
                If (Row Mod 3) = 2 Then
                    .Col = 1: .Row = Row
                    If .Value = "" Then
                        .Col = Col: .Row = Row
                        .Value = ""
                        Exit Sub
                    End If
                    .Col = Col: .Row = Row
                    If .Value Like "S*" Then
                        .ForeColor = DCM_LightRed
                    Else
                        .ForeColor = DCM_LightBlue
                    End If
                    If (Trim(.Value) <> "") And (InStr(1, MRT_SenRstCd, Trim(.Value))) = 0 Then
                        .SelStart = 0: .SelLength = Len(.Value)
                        medBeep (10)
                        Exit Sub
                    End If
                    If Len(.Value) = 1 Then SendKeys "{ENTER}"
                
                ElseIf (Row Mod 3) = 0 Then
                    .Row = Row: .Col = Col
                    picMIC.Visible = True
                    picMIC.ZOrder 0
                    lstMIC(0).ListIndex = 0
                    lstMIC(1).ListIndex = -1
                    lstMIC(2).ListIndex = -1
                    fCurMic = 0
                End If
                
            End With
    End Select

End Sub

Private Sub ssSusc_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim sIdxNm As Integer, sIdxQty As Integer

    If blnMsgFg Then Exit Sub
    
    If (NewRow Mod 3) = 1 Then
        NewRow = NewRow + 1
        ssSusc.Col = NewCol: ssSusc.Row = NewRow
        ssSusc.Action = ActionActiveCell
        DoEvents
    End If

    If (NewRow Mod 3) = 0 And NewCol < 8 And NewRow < ssSusc.MaxRows Then
        NewRow = NewRow + 2
        ssSusc.Col = NewCol: ssSusc.Row = NewRow
        ssSusc.Action = ActionActiveCell
        DoEvents
    End If

    If NewCol < 8 Or (NewRow Mod 3) > 0 Then picMIC.Visible = False
    
    Select Case Col
        Case 1
            lstMicCd.Visible = False: lstMicNm.Visible = False
        Case 5
            lstQty.Visible = False
        Case Is > 7
            If (Row Mod 3) = 2 Then
                ssSusc.Col = Col: ssSusc.Row = Row
                If Trim(ssSusc.Value) <> "" And (InStr(1, MRT_SenRstCd, Trim(ssSusc.Value))) = 0 Then
                    Cancel = True
                    medBeep (10)
                    Exit Sub
                Else '과거 항생제 결과 비교
                    If objMicLib.PreRstSens <> ssSusc.Value Then
                        Call objMicLib.GetWarningForSens(ssSusc)
                    End If
                End If
                
                ssSusc.Row = Row - 1
                If (ssSusc.ForeColor = vbRed) Or (ssSusc.FontItalic = True) Then
                    shpWarning.Visible = True
                    lblWarning.Visible = True
                    ssSusc.FontBold = True
                Else
                    shpWarning.Visible = False
                    lblWarning.Visible = False
                    ssSusc.FontBold = False
                End If
            ElseIf (Row Mod 3) = 0 Then
                Call objMicLib.GetWarningForMedi(ssSusc)
                
                ssSusc.Col = Col
                ssSusc.Row = Row - 2
                If (ssSusc.ForeColor = vbRed) Or (ssSusc.FontItalic = True) Then
                    shpWarning.Visible = True
                    lblWarning.Visible = True
                    ssSusc.FontBold = True
                Else
                    shpWarning.Visible = False
                    lblWarning.Visible = False
                    ssSusc.FontBold = False
                End If
            End If
    End Select

    ' 현재 에디팅 중인 라인 설정
    objMicLib.CRow = NewRow
    objMicLib.CCol = NewCol
    blnSendKeys = True
    ssSusc.ArrowsExitEditMode = True

    Select Case NewCol

        Case 1  '균명

            ssSusc.Col = 2: ssSusc.Row = NewRow
            fPrevCode = ssSusc.Text
            ssSusc.Col = NewCol: ssSusc.Row = NewRow
            sIdxNm = medListFind(lstMicNm, ssSusc.Value)
            If sIdxNm = -1 Then sIdxNm = 0

            lstMicCd.ListIndex = sIdxNm: lstMicNm.ListIndex = sIdxNm
            lstMicCd.Visible = True: lstMicNm.Visible = True
            lstMicCd.ZOrder: lstMicNm.ZOrder
            
        Case 4  'MIC 여부
            ssSusc.Col = 1: ssSusc.Row = NewRow
            If Trim(ssSusc.Text) <> "" Then
                ssSusc.Col = NewCol: fMicFg = ssSusc.Text
                If P_MICSelectedByUser Then     'MIC 선택여부
                    Dim lngY As Long
                    Dim lngX As Long
                    Dim lngW As Long, lngH As Long
                    
                    objMicLib.CRow = NewRow
                    Me.ScaleMode = 1
                    Call ssSusc.GetCellPos(NewCol, NewRow, lngX, lngY, lngW, lngH)
                    
                    blnMsgFg = True
                    
                    Set frmMic = New frmMicOption
                    frmMic.Top = ssSusc.Top + fraSusc.Top + lngY '+ 1400 'lblSenType.Top + 1800
                    frmMic.Left = ssSusc.Left + fraSusc.Left + lngX '+ 500  '7600
                    frmMic.Show 1
                    
                    ssSusc.SetFocus
                    blnMsgFg = False
                    Exit Sub
                Else
                    ssSusc.Row = NewRow: ssSusc.Col = NewCol
                    ssSusc.Text = MRT_GenSen
                End If
            End If

        Case 5  '정도코드

            ssSusc.Col = 1: ssSusc.Row = NewRow
            If Trim(ssSusc.Text) <> "" Then

               ssSusc.Col = NewCol + 1: ssSusc.Row = NewRow
               sIdxQty = medListFind(lstQty, ssSusc.Value)
               If sIdxQty = -1 Then sIdxQty = 0
               lstQty.ListIndex = sIdxQty
               lstQty.Visible = True
               lstQty.ZOrder

            End If
'감수성 결과 로직과 동일하게 할려고 리마크 처리하고 아래 로직 추가함
'Modify By Legends 2003/07/28
'        Case Is > 7
'            If (NewRow Mod 3) = 0 Then
'                ssSusc.ArrowsExitEditMode = False
'
'                picMIC.Visible = True
'                picMIC.ZOrder 0
'                lstMIC(0).ListIndex = 0
'                lstMIC(1).ListIndex = -1
'                lstMIC(2).ListIndex = -1
'                fCurMic = 0
'
'                ssSusc.Col = NewCol: ssSusc.Row = NewRow
'                sIdxNm = medListFind(lstMIC(0), ssSusc.Value)
'                lstMIC(0).ListIndex = sIdxNm
'                If sIdxNm = -1 Then
'                    sIdxNm = medListFind(lstMIC(1), ssSusc.Value)
'                    lstMIC(1).ListIndex = sIdxNm
'                    If sIdxNm = -1 Then
'                        sIdxNm = medListFind(lstMIC(2), ssSusc.Value)
'                        lstMIC(2).ListIndex = sIdxNm
'                        If sIdxNm = -1 Then
'                            lstMIC(0).ListIndex = 0
'                        Else
'                            fCurMic = 2
'                        End If
'                    Else
'                        fCurMic = 1
'                    End If
'                Else
'                    fCurMic = 0
'                End If
'
'            Else
'                ssSusc.ArrowsExitEditMode = True
'            End If
            If Col = 4 Then
                ssSusc.Col = 4: ssSusc.Row = Row
                If fMicFg <> ssSusc.Value Then
                    Call ShowAntiList
                    'SendKeys "{ENTER}"
                End If
            End If
            
        Case Is > 7
            If (NewRow Mod 3) = 0 Then
                ssSusc.ArrowsExitEditMode = False
                picMIC.Visible = True
                picMIC.ZOrder 0
                lstMIC(0).ListIndex = 0
                lstMIC(1).ListIndex = -1
                lstMIC(2).ListIndex = -1
                fCurMic = 0
            Else
                ssSusc.ArrowsExitEditMode = True
            End If
    End Select

    ssSusc.Col = NewCol: ssSusc.Row = NewRow

End Sub

'Private Sub GetWarningForSens()
'    Dim RS As recordset
'    Dim blnWarning As Boolean
'    Dim strMnmCd As String
'    Dim strRst As String
'    Dim i As Long
'    Dim j As Long
'
'    Dim varMnmCd As Variant
'    Dim varAnti As Variant '항생제 타이틀
'    Dim varSens As Variant
'    Dim varMedi As Variant
'    Dim aryRec() As String
'    Dim aryFld() As String
'
'    Call ssSusc.GetText(fSSCol, fSSRow, varSens)
'
'    If varSens = "" Then Exit Sub
'
'    If fSSRow = -1 Or fSSCol = -1 Then Exit Sub
'
'    If ChkInPatient = False Then
'        blnWarning = False
'    Else
'        Set RS = GetLastestRst
'
'        If RS Is Nothing Then
'            blnWarning = False
'        Else
'            Do Until RS.EOF
'                strRst = ""
'
'                For i = 1 To RS.Fields("scnt").Value & ""
'                    strRst = strRst & RS.Fields("srst" & i).Value & "" & COL_DIV
'                Next
'
'                strMnmCd = strMnmCd & RS.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
'
'                RS.MoveNext
'            Loop
'            blnWarning = True
'        End If
'    End If
'
'    ssSusc.Col = fSSCol: ssSusc.Row = fSSRow - 1
'    ssSusc.ForeColor = vbBlack
'    shpWarning.Visible = False
'    lblWarning.Visible = False
'
'    If blnWarning Then
'        Call ssSusc.GetText(2, fSSRow, varMnmCd)
'        Call ssSusc.GetText(fSSCol, fSSRow - 1, varAnti)
'        Call ssSusc.GetText(fSSCol, fSSRow, varSens)
'        Call ssSusc.GetText(fSSCol, fSSRow + 1, varMedi)
'
'        If InStr(strMnmCd, varMnmCd) > 0 Then
'            aryRec = Split(strMnmCd, LINE_DIV)
'            For i = LBound(aryRec) To UBound(aryRec) - 1
'                aryFld = Split(aryRec(i), COL_DIV)
'                If aryFld(0) = varMnmCd Then
'                    For j = 1 To UBound(aryFld) - 1
'                        If medGetP(aryFld(j), 1, ";") = varAnti Then
'                            If medGetP(aryFld(j), 2, ";") = varSens Then
'                                fSSRow = fSSRow + 1
'                                Call GetWarningForMedi
'                                Exit For
'                            Else
'                                ssSusc.ForeColor = vbRed
'
'                                If PreRstSens <> ssSusc.Value Then
'                                    If ssSusc.FontItalic Then ssSusc.FontItalic = False
'                                End If
'                                Exit For
'                            End If
'                        End If
'                    Next j
'                End If
'            Next i
'        End If
'    End If
'
'    Set RS = Nothing
'End Sub

'Private Sub GetWarningForMedi()
'    Dim RS As recordset
'    Dim blnWarning As Boolean
'    Dim strMnmCd As String
'    Dim strRst As String
'    Dim i As Long
'    Dim j As Long
'
'    Dim varMnmCd As Variant
'    Dim varAnti As Variant '항생제 타이틀
'    Dim varSens As Variant
'    Dim varMedi As Variant
'    Dim aryRec() As String
'    Dim aryFld() As String
'
'    If fSSRow = -1 Or fSSCol = -1 Then Exit Sub
'
'    If ChkInPatient = False Then
'        blnWarning = False
'    Else
'        Set RS = GetLastestRst
'
'        If RS Is Nothing Then
'            blnWarning = False
'        Else
'            Do Until RS.EOF
'                strRst = ""
'
'                For i = 1 To RS.Fields("scnt").Value & ""
'                    strRst = strRst & RS.Fields("srst" & i).Value & "" & COL_DIV
'                Next
'
'                strMnmCd = strMnmCd & RS.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
'
'                RS.MoveNext
'            Loop
'            blnWarning = True
'        End If
'    End If
'
'    ssSusc.Col = fSSCol: ssSusc.Row = fSSRow - 2
'    shpWarning.Visible = False
'    lblWarning.Visible = False
'    ssSusc.FontItalic = False
'
'    If blnWarning Then
'        Call ssSusc.GetText(2, fSSRow - 1, varMnmCd)
'        Call ssSusc.GetText(fSSCol, fSSRow - 2, varAnti)
'        Call ssSusc.GetText(fSSCol, fSSRow - 1, varSens)
'        Call ssSusc.GetText(fSSCol, fSSRow, varMedi)
'
'        If InStr(strMnmCd, varMnmCd) > 0 Then
'            aryRec = Split(strMnmCd, LINE_DIV)
'
'            For i = LBound(aryRec) To UBound(aryRec) - 1
'                aryFld = Split(aryRec(i), COL_DIV)
'                If aryFld(0) = varMnmCd Then
'                    For j = 1 To UBound(aryFld) - 1
'                        If medGetP(aryFld(j), 1, ";") = varAnti Then
'                            If medGetP(aryFld(j), 2, ";") = varSens Then
'                                If medGetP(aryFld(j), 3, ";") <> varMedi Then
'                                    ssSusc.FontItalic = True
'                                    Exit For
'                                End If
'                            End If
'                        End If
'                    Next j
'                End If
'            Next i
'        End If
'    End If
'
'    Set RS = Nothing
'End Sub

'Private Sub GetWarningForToolTip(ByVal pCol As Long, ByVal pRow As Long, ByRef pTextTip As String)
'    Dim RS As recordset
'    Dim blnWarning As Boolean
'    Dim strMnmCd As String
'    Dim strRst As String
'    Dim i As Long
'    Dim j As Long
'
'    Dim varMnmCd As Variant
'    Dim varAnti As Variant '항생제 타이틀
'    Dim aryRec() As String
'    Dim aryFld() As String
'
'    pTextTip = ""
'
'    If ChkInPatient = False Then
'        blnWarning = False
'    Else
'        Set RS = GetLastestRst
'
'        If RS Is Nothing Then
'            blnWarning = False
'        Else
'            Do Until RS.EOF
'                strRst = ""
'
'                For i = 1 To RS.Fields("scnt").Value & ""
'                    strRst = strRst & RS.Fields("srst" & i).Value & "" & COL_DIV
'                Next
'
'                strMnmCd = strMnmCd & RS.Fields("mnmcd").Value & "" & COL_DIV & strRst & "" & LINE_DIV
'
'                RS.MoveNext
'            Loop
'            blnWarning = True
'        End If
'    End If
'
'    If blnWarning Then
'        Call ssSusc.GetText(2, pRow + 1, varMnmCd)
'        Call ssSusc.GetText(pCol, pRow, varAnti)
'
'        aryRec = Split(strMnmCd, LINE_DIV)
'
'        For i = LBound(aryRec) To UBound(aryRec) - 1
'            aryFld = Split(aryRec(i), COL_DIV)
'            If aryFld(0) = varMnmCd Then
'                For j = 1 To UBound(aryFld) - 1
'                    If medGetP(aryFld(j), 1, ";") = varAnti Then
'                        pTextTip = vbNewLine & "  - 이전결과 - " & vbNewLine & _
'                                               "    항 생 제 : " & medGetP(aryFld(j), 1, ";") & vbNewLine & _
'                                               "    감 수 성 : " & medGetP(aryFld(j), 2, ";") & vbNewLine & _
'                                               "    농    도 : " & medGetP(aryFld(j), 3, ";") & vbNewLine
'                        Exit For 'j
'                    End If
'                Next j
'                Exit For 'i
'            End If
'        Next i
'    End If
'
'    Set RS = Nothing
'End Sub

Private Sub ssSusc_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Col = 1 Then
        With ssSusc
            .Row = Row: .Col = Col
            If .Value = "" Then Exit Sub
            MultiLine = 1
            TipText = vbCRLF & "  " & .Text & vbCRLF
            TipWidth = 5500
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End With
    End If
    
    If Col > 7 Then
        Dim TextTip As String
        
        With ssSusc
            If (Row Mod 3) = 1 Then
                .Row = Row: .Col = 1: If .FontItalic Then Exit Sub
                                .Col = Col: If .FontBold = False Then Exit Sub
                .Row = Row: .Col = Col: If .Value = "" Then Exit Sub
                Call objMicLib.GetWarningForToolTip(ssSusc, Col, Row, TextTip)
            ElseIf (Row Mod 3) = 2 Then
                .Row = Row - 1: .Col = 1: If .FontItalic Then Exit Sub
                            .Col = Col: If .FontBold = False Then Exit Sub
                Call objMicLib.GetWarningForToolTip(ssSusc, Col, Row - 1, TextTip)
            ElseIf (Row Mod 3) = 0 Then
                .Row = Row - 2: .Col = 1: If .FontItalic Then Exit Sub
                                .Col = Col: If .FontBold = False Then Exit Sub
                Call objMicLib.GetWarningForToolTip(ssSusc, Col, Row - 2, TextTip)
            End If
            
            If TextTip = "" Then Exit Sub
            
            MultiLine = 1
            TipText = TextTip
            TipWidth = 2000
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End With
    End If
End Sub

Private Sub txtAccSeq_Change()
    fraOldSensi.Visible = False
End Sub

Private Sub txtWorkArea_Change()
    If Not txtAccDt.Enabled Then Exit Sub
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
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
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then If Not cmdLisLabel11fg Then txtAccSeq.SetFocus
End Sub

Private Sub txtAccDt_GotFocus()
    txtAccDt.SelStart = 0
    txtAccDt.SelLength = Len(txtAccDt)
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then txtAccSeq.SetFocus

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
    
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String, sMicFg As String
    Dim sqlInfo As String
    Dim lngTestCnt As Long

    sWorkArea = Trim(txtWorkArea): sAccDt = Trim(txtAccDt): sAccSeq = Trim(txtAccSeq)
    sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)
    
    Call ICSLabNoMark(sWorkArea, sAccDt, sAccSeq, enICSNum.LIS_ALL)
    
    '병동/진료과 연락처(환자ID,CONTROL)
    Call GetPtTelInfo(sWorkArea, sAccDt, sAccSeq, lblTelno)
    
    If KeyAscii = vbKeyReturn Then
        Call ClearForm
    End If
    
    If KeyAscii = vbKeyReturn Or KeyAscii = 999 Then
        lngTestCnt = objMicCul.GetTestByLabNo(sWorkArea, sAccDt, sAccSeq, lstTest)
        
        If lngTestCnt >= 1 Then
            lstTest.ListIndex = 0
            fTestCd = medGetP(lstTest.Text, 1, vbTab)
            lblTestType.Caption = medGetP(lstTest.Text, 2, vbTab)
            lstTest.Visible = False
        End If

        If DispPtInfo(sWorkArea, sAccDt, sAccSeq, fTestCd) Then
            With objMicLib
                .WorkArea = sWorkArea
                .AccDt = sAccDt
                .AccSeq = sAccSeq
                .TestCd = fTestCd
                .SpcCd = lblSpcCd.Caption
                .PtId = lblPtId.Caption
            End With
        
            Call objMicCul.DispStainResult(sWorkArea, sAccDt, sAccSeq, lstGramStain)
            Call DispGrowthRst(sWorkArea, sAccDt, sAccSeq, sMicFg)
            
            Dim strLabNo As String
'            objMicLib.AccNo = txtWorkArea.Text & Mid(Format(GetSystemdate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text
'            strLabNo = objMicCul.GetLatestSensiAccNo(lblPtId.Caption, _
'                                Format(GetSystemdate, CS_DateDbFormat), lblSpcCd.Caption)
            strLabNo = objMicLib.GetAccNoOfLatestRst(, txtWorkArea.Text & Mid(Format(GetSystemDate, "yyyyMMdd"), 1, 2) & txtAccDt.Text & txtAccSeq.Text)
            If Trim(strLabNo) = "" Then
                cmdGetOldResult.Visible = False
            Else
                cmdGetOldResult.Visible = True
            End If
            
            fraSusc.Enabled = True
            cboNGRst.SetFocus
            Exit Sub
        End If
    End If

'    cmdOrderView.Visible = True

    ' 만약에 숫자가 아니면 문자를 없애버려도 좋음(백스페이스 허용)
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Function DispPtInfo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccSeq As String, ByVal pTestCd As String) As Boolean
    
    Dim i As Integer, j As Integer
    Dim sRemarkCd As String, sRemarkIdx As Integer
    Dim objPtDic As clsDictionary

    blnPtFg = False

    Set objPtDic = objMicRst.DispSenDataByTestCd(pWorkArea, pAccDt, pAccSeq, pTestCd)

    If objPtDic Is Nothing Then
        MsgBox "감수성 검사로 접수 되지 않은 Lab-No 입니다. 확인 후 처리하십시오", vbInformation, "감수성결과등록"
        ClearForm
        txtAccSeq.SelStart = 0: txtAccSeq.SelLength = Len(txtAccSeq)
        DispPtInfo = False: Exit Function
    End If

    If objPtDic.Fields("stscd") < enStsCd.StsCd_LIS_FinRst Then
        MsgBox "아직 최종 확인되지 않은 결과입니다. 등록 화면을 이용하십시오", vbInformation, "감수성결과등록"
        ClearForm
        txtAccSeq.SelStart = 0: txtAccSeq.SelLength = Len(txtAccSeq)
        DispPtInfo = False: Exit Function
    End If

    fTestCd = objPtDic.Fields("testcd")
'    lblSenType = objPtDic.Fields("rsttype")
    fMfySeq = Val(objPtDic.Fields("mfyseq"))
'
'    If lblSenType = MRT_GenSen Then lblTestType.Caption = MNM_GSen
'    If lblSenType = MRT_MicSen Then lblTestType.Caption = MNM_MSen

    ' 데이타 화면에 출력
    lblPtId.Caption = objPtDic.Fields("ptid")
    lblPtNm.Caption = objPtDic.Fields("ptnm")
    lblPtSA.Caption = objPtDic.Fields("sexage")
    lblDept.Caption = objPtDic.Fields("deptcd")
    lblWard.Caption = objPtDic.Fields("location")
    lblWardId.Caption = objPtDic.Fields("wardid")
    lblSpecimen.Caption = objPtDic.Fields("spcnm")
    lblSpcCd.Caption = objPtDic.Fields("spccd")
    lblMajDoct.Caption = objPtDic.Fields("orddoct")
    lblDoctNm.Caption = objPtDic.Fields("orddrnm")
    
    txtDtId.Text = objPtDic.Fields("orddoct")
    txtExDtId.Text = objPtDic.Fields("majdoct")
    rtfMessage.Text = ""
    strRcvDt = objPtDic.Fields("rcvdt")
    txtTestCd = objPtDic.Fields("testcd")
    
    fFNSeq = Val(objPtDic.Fields("footnotefg"))
    sRemarkCd = objPtDic.Fields("rmkcd")
    sRemarkIdx = -1

    ' footnote Display
    txtPNote.Text = "": txtFNote.Text = ""
    If fFNSeq > 0 Then txtPNote.Text = objMicRst.DispFootNote(pWorkArea, pAccDt, pAccSeq)

    ' 검체 Remark Display
    sRemarkIdx = medComboFind(cboRemark, sRemarkCd)
    If sRemarkIdx < 0 Then
        cboRemark.ListIndex = 0
    Else
        cboRemark.ListIndex = sRemarkIdx
    End If

    cmdOrderView.Visible = True

    DispPtInfo = True

End Function



Private Sub DispGrowthRst(ByVal pWorkArea As String, ByVal pAccDt As String, _
                          ByVal pAccSeq As String, ByVal pMicFg As String)

    Dim strGrowthRst As String
    Dim sSenFg As String
    Dim i As Long, iACnt As Long, iCnt As Long, sMK As String
    Dim j As Long
    
    sMK = "GC"

    strGrowthRst = objMicCul.DispGrowthResult(pWorkArea, pAccDt, pAccSeq, sMK, fTestCd) 'pMicFg)
    lblNogrowth.Caption = medGetP(strGrowthRst, 1, COL_DIV)
    sSenFg = medGetP(strGrowthRst, 2, COL_DIV)

    txtVfyDt.Text = medGetP(strGrowthRst, 3, COL_DIV)
    ssSusc.ReDraw = False

    ' 감수성 결과 Display
    If sSenFg = MRT_SenRst Then

'        Call objMicCul.DispSensiResult(ssSusc, pWorkArea, pAccDt, pAccseq, fTestCd, fMfySeq)
        Call objMicLib.DispSensiResultForWarn(ssSusc, pWorkArea, pAccDt, pAccSeq, fTestCd, fMfySeq)
        With ssSusc
            For i = 1 To .MaxRows Step 3
                .Row = i + 1
                .Col = 2
                If .Value <> "" Then
                    ' 항생제 적용되지않는 나머지 셀을 사용못하게
                    .Col = 7
                    iACnt = Val(.Value)
                    
                    .Row = i + 1: .Row2 = i + 2
                    .Col = iACnt + 8: .COL2 = .MaxCols
                    .BlockMode = True
                    .CellType = 5   'CellTypeStaticText
                    .BackColor = fSkColor
                    .BlockMode = False
                    
                    .Row = i + 1: .Col = 4
                    If .Value = MRT_MicSen Then
                        .Col = iACnt + 8
                    Else
                        .Col = 8
                    End If
                    .COL2 = .MaxCols
                    .Row = i + 2: .Row2 = i + 2
                    .BlockMode = True
                    .CellType = CellTypeStaticText
                    .BackColor = fSkColor
                    .BlockMode = False
                End If
                
                .Row = i + 1
                For j = 8 To iACnt + 8
                    .Col = j
                    If .Value Like "S*" Then
                        .ForeColor = DCM_LightRed
                    Else
                        .ForeColor = DCM_LightBlue
                    End If
                Next
            Next
        End With
    End If

    ssSusc.ReDraw = True

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
                .rtfText = txtFNote.Text
        End Select
    End With
    gintTemplete = pintPrg
End Sub


Private Sub lstlastAcc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTmp  As String
    Dim objRstSql As clsLISSqlReview
    Dim tmpSQL As String
    Dim tmpRs As Recordset
    Dim FootNote As String
    Dim strWA       As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    
    Set objRstSql = New clsLISSqlReview
    If KeyCode = vbKeyReturn Then
        strTmp = (lstlastAcc.List(lstlastAcc.ListIndex))
        Call LastOldResult(strTmp)
    
        strWA = medGetP(strTmp, 1, "-")
        strAccDt = medGetP(strTmp, 2, "-")
        strAccSeq = medGetP(strTmp, 3, "-")
    
        '2007.06.28 osw 풋노트
        tmpSQL = objRstSql.SqlGetFootNote(strWA, strAccDt, strAccSeq)
        
        Set tmpRs = New Recordset
        tmpRs.Open tmpSQL, DBConn
        
        txtSamCmt.Text = ""
        If Not tmpRs.EOF Then 'GoTo NoData
       
            FootNote = "<< Foot Note >>" & vbCRLF
            FootNote = FootNote & Trim("" & tmpRs.Fields("FootNote").Value)
'            While (Not tmpRs.EOF)
'                FootNote = FootNote & Trim("" & tmpRs.Fields("FootNote").Value) & vbCRLF
'                tmpRs.MoveNext
'            Wend
            txtSamCmt.Text = FootNote
        End If
    End If
    
End Sub
Private Sub lstlastAcc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call lstlastAcc_KeyDown(13, 0)
End Sub

Private Sub LastOldResult(ByVal strTmp As String)
    
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    
    Dim i As Long, j As Long, k As Long
    
    tblResult.Row = -1
    tblResult.Col = -1
    tblResult.BlockMode = True
    tblResult.Action = ActionClearText
    tblResult.BlockMode = False
    lblVfyDt.Caption = ""
        
    
    sWorkArea = medGetP(strTmp, 1, "-")
    sAccDt = medGetP(strTmp, 2, "-")
    sAccSeq = medGetP(strTmp, 3, "-")
    
    lblVfyDt.Caption = "최근 감수성결과 보고일시 : " & Format(medGetP(strTmp, 4, "-"), CS_DateLongMask)
    lblVfyDt.Caption = lblVfyDt.Caption & String(2, " ") & Format(medGetP(strTmp, 5, "-"), CS_TimeLongMask)
    lblVfyDt.Caption = lblVfyDt.Caption & String(2, " ") & "접수번호 : " & sWorkArea & "-" & Mid(sAccDt, 3) & "-" & sAccSeq
    
    Dim MyResult As New clsLISResultReview
    
    With MyResult
      
        MouseRunning
        
        Call .MicrobeSensiRst(sWorkArea, sAccDt, sAccSeq)
      
        If .ResultCnt = 0 Then
            MouseDefault
            Exit Sub
        End If
      
        For i = 1 To .RstRow
            tblResult.Row = i + .OffSet
            For j = 1 To 8
                tblResult.Col = j
                If .Get_ForeColor(j, i) <> 0 Then tblResult.ForeColor = .Get_ForeColor(j, i)
            Next
        Next
      
        '결과내역 Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.COL2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText    '& .SenClipText             'ResultBuffer
        tblResult.BlockMode = False
      
        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
        'If .SortFg Then
        If .SortFg Then
'            tblResult.SortBy = SortByRow
'            tblResult.SortKey(1) = 2  '항생제명
'            tblResult.SortKeyOrder(1) = SortKeyOrderAscending
'            tblResult.Col = -1
'            tblResult.Row = .SortStartRow   '+ .OffSet
'            tblResult.Row2 = .SortEndRow  '+ .OffSet
'            tblResult.Action = ActionSort
'            tblResult.Row = .SortStartRow - 1   '+ .OffSet
'            tblResult.Col = 2
'            tblResult.FontUnderline = True
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        
        '미생물 결과 : 균명컬럼 Align Left
        tblResult.Row = -1
        tblResult.Col = -1
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.TypeHAlign = TypeHAlignLeft
        tblResult.BlockMode = False
        tblResult.ColWidth(2) = 10
        'tblResult.ColWidth(3) = 60
        For i = 1 To 5
            If .MicFg(i) Then
                tblResult.ColWidth(i + 2) = 9
            Else
                tblResult.ColWidth(i + 2) = 4
            End If
        Next
        tblResult.ColWidth(8) = 20
        tblResult.Col = 3: tblResult.COL2 = 7
        tblResult.Row = -1
        tblResult.BlockMode = True
        tblResult.FontBold = False
        tblResult.BlockMode = False
    
    End With
    
    Dim strAntiCd As String
    Dim strAntiRst As String
    Dim lngMicCnt As Long
    Dim blnFind As Boolean
    
    With ssSusc
        For i = 1 To ssSusc.DataRowCnt Step 3
            .Row = i + 1
            .Col = 1
            If .Value <> "" Then
                .Col = 7
                lngMicCnt = Val(.Value)
                For k = 8 To lngMicCnt + 8
                    .Row = i
                    .Col = k: strAntiCd = .Value
                    .Row = i + 1
                    .Col = k: strAntiRst = .Value
                    
                    tblResult.Row = MyResult.SortStartRow - 1
                    tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
                    tblResult.Value = "현재" & CStr(i \ 3 + 1)
                    tblResult.ForeColor = DCM_LightRed
                    
                    blnFind = False
                    For j = MyResult.SortStartRow To tblResult.DataRowCnt
                        tblResult.Row = j
                        tblResult.Col = 2
                        If tblResult.Value = strAntiCd Then
                            tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
                            tblResult.Value = strAntiRst
                            tblResult.ColWidth(MyResult.MicrobeCount + (i \ 3) + 3) = 4
                            If strAntiRst Like "S*" Then
                                tblResult.ForeColor = DCM_Red
                            Else
                                tblResult.ForeColor = DCM_Gray
                            End If
                            blnFind = True
                            Exit For
                        End If
                    Next
                    If Not blnFind Then
                        tblResult.Row = tblResult.DataRowCnt + 1
                        tblResult.Col = 2
                        tblResult.Value = strAntiCd
                        tblResult.Col = MyResult.MicrobeCount + (i \ 3) + 3
                        tblResult.Value = strAntiRst
                        If strAntiRst Like "S*" Then
                            tblResult.ForeColor = DCM_Red
                        Else
                            tblResult.ForeColor = DCM_Gray
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    MouseDefault
    
    fraOldSensi.Visible = True
    fraOldSensi.ZOrder 0
    Set MyResult = Nothing
    
End Sub
