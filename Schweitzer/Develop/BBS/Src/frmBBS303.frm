VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS303 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Blood Delivery"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15105
   Icon            =   "frmBBS303.frx":0000
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15105
   Visible         =   0   'False
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출고(&S)"
      Height          =   510
      Left            =   9135
      Style           =   1  '그래픽
      TabIndex        =   48
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   47
      Tag             =   "124"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   46
      Tag             =   "128"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdF 
      BackColor       =   &H00F4F0F2&
      Caption         =   "필터"
      Height          =   510
      Left            =   10470
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "15101"
      Top             =   8565
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame fraFilter 
      BackColor       =   &H00F4F0F2&
      Height          =   5295
      Left            =   8490
      TabIndex        =   0
      Top             =   2550
      Visible         =   0   'False
      Width           =   5760
      Begin VB.CommandButton cmdFQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   750
         Left            =   4545
         Style           =   1  '그래픽
         TabIndex        =   4
         Tag             =   "124"
         Top             =   705
         Width           =   1170
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F4F0F2&
         Caption         =   "닫기"
         Height          =   510
         Left            =   2940
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "128"
         Top             =   4500
         Width           =   1320
      End
      Begin VB.CommandButton cmdFSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "필터출고"
         Height          =   510
         Left            =   1620
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "124"
         Top             =   4500
         Width           =   1320
      End
      Begin VB.TextBox txtPtid 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1095
         TabIndex        =   5
         Top             =   345
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70320129
         CurrentDate     =   38170
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
         BackColor       =   8388608
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "FILTER 출고"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   14
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtnm1 
         Height          =   360
         Left            =   2295
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   27
         Left            =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "출고일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   2595
         TabIndex        =   10
         Top             =   345
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70320129
         CurrentDate     =   38170
      End
      Begin FPSpread.vaSpread tblFilter 
         Height          =   2580
         Left            =   30
         TabIndex        =   11
         Top             =   1470
         Width           =   5655
         _Version        =   196608
         _ExtentX        =   9975
         _ExtentY        =   4551
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         GrayAreaBackColor=   15265518
         GridColor       =   16703181
         GridShowVert    =   0   'False
         MaxCols         =   9
         MaxRows         =   7
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS303.frx":076A
         TextTip         =   2
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   16
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "출고수량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDelCnt 
         Height          =   360
         Left            =   1080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1095
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F0F2&
         Caption         =   "~"
         Height          =   255
         Left            =   2415
         TabIndex        =   14
         Top             =   420
         Width           =   360
      End
   End
   Begin DRcontrol1.DrFrame fraMode 
      Height          =   1485
      Index           =   3
      Left            =   1020
      TabIndex        =   16
      Top             =   7590
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2619
      Title           =   "Title"
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
      Begin VB.CommandButton cmdBagID 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   540
         Width           =   350
      End
      Begin VB.TextBox txtBagID 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1080
         TabIndex        =   17
         Top             =   540
         Width           =   1050
      End
      Begin MedControls1.LisLabel lblBagNm 
         Height          =   360
         Left            =   2520
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   105
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "요청자"
         Appearance      =   0
      End
   End
   Begin DRcontrol1.DrFrame fraMode 
      Height          =   1485
      Index           =   2
      Left            =   1020
      TabIndex        =   21
      Top             =   6090
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2619
      Title           =   "Title"
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
      Begin VB.TextBox txtExpRsnRmk 
         Appearance      =   0  '평면
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   25
         Text            =   "frmBBS303.frx":0C1E
         Top             =   855
         Width           =   5115
      End
      Begin VB.ComboBox cboExpRsnCd 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   1080
         Style           =   2  '드롭다운 목록
         TabIndex        =   24
         Top             =   495
         Width           =   2895
      End
      Begin VB.CommandButton cmdExpId 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   90
         Width           =   350
      End
      Begin VB.TextBox txtExpId 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1050
         TabIndex        =   22
         Top             =   105
         Width           =   1050
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   90
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "폐기사유"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   90
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   105
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "요청자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExpNm 
         Height          =   360
         Left            =   2520
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   105
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExpBillDiv 
         Height          =   360
         Left            =   5505
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   90
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   855
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "Memo"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   4125
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "환자부담여부"
         Appearance      =   0
      End
   End
   Begin DRcontrol1.DrFrame fraMode 
      Height          =   1485
      Index           =   1
      Left            =   1020
      TabIndex        =   32
      Top             =   4590
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2619
      Title           =   "Title"
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
      Begin VB.TextBox txtRetRmk 
         Appearance      =   0  '평면
         Height          =   855
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   37
         Text            =   "frmBBS303.frx":0C24
         Top             =   540
         Width           =   5115
      End
      Begin VB.CommandButton cmdRetID 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   36
         Top             =   120
         Width           =   350
      End
      Begin VB.TextBox txtRetID 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1080
         TabIndex        =   35
         Top             =   120
         Width           =   1050
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   90
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   540
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "Memo"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   90
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "요청자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRetNm 
         Height          =   360
         Left            =   2520
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   120
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
   End
   Begin DRcontrol1.DrFrame fraMode 
      Height          =   1485
      Index           =   0
      Left            =   7860
      TabIndex        =   39
      Top             =   7005
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2619
      Title           =   "Title"
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
      Begin VB.CommandButton cmdRcvId 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2445
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   43
         Top             =   300
         Width           =   350
      End
      Begin VB.TextBox txtRcvId 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   300
         Width           =   1110
      End
      Begin VB.CheckBox chkFilter 
         BackColor       =   &H8000000E&
         Caption         =   "Filter사용 여부"
         Height          =   315
         Left            =   2640
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkIrra 
         BackColor       =   &H8000000E&
         Caption         =   "Irradiation 여부"
         Height          =   315
         Left            =   300
         TabIndex        =   40
         Top             =   840
         Width           =   2055
      End
      Begin MedControls1.LisLabel lblRcvNm 
         Height          =   315
         Left            =   2805
         TabIndex        =   44
         Top             =   300
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   24
         Left            =   270
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "수령자"
         Appearance      =   0
      End
   End
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   6915
      Left            =   7845
      TabIndex        =   49
      Top             =   90
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12197
      Title           =   "처방정보"
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MedControls1.LisLabel lblDeliverySeq 
         Height          =   315
         Left            =   4800
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblWorkSeq 
         Height          =   315
         Left            =   4800
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblirrstring 
         Height          =   315
         Left            =   1215
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1665
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblirrFg 
         Height          =   315
         Left            =   135
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "Irradiation"
         Appearance      =   0
      End
      Begin VB.PictureBox picStat 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   3825
         ScaleHeight     =   675
         ScaleWidth      =   2415
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   345
         Width           =   2475
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "응급Assign"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   2205
         End
      End
      Begin VB.ListBox lstNewTest 
         Appearance      =   0  '평면
         Height          =   660
         Left            =   1080
         Style           =   1  '확인란
         TabIndex        =   62
         Top             =   5500
         Width           =   4995
      End
      Begin VB.TextBox txtRmk 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         Height          =   705
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   4605
         Width           =   6060
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   17
         Left            =   135
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   2730
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "처방코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   135
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2385
         Width           =   1035
         _ExtentX        =   1826
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   19
         Left            =   135
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3075
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "수량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   20
         Left            =   135
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3765
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "수혈사유"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   21
         Left            =   135
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "예정일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDelivery 
         Height          =   315
         Left            =   135
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "출고일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDelivery1 
         Height          =   315
         Left            =   2475
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "출고시간"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   135
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   630
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "Volume"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   135
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   13
         Left            =   135
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   975
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "XM검사일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAccNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2385
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         BackColor       =   14411494
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestCd 
         Height          =   315
         Left            =   1200
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   315
         Left            =   2445
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2730
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblUnitQty 
         Height          =   315
         Left            =   1200
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3075
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   315
         Left            =   4890
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReqDtTm 
         Height          =   315
         Left            =   1200
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3420
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BackColor       =   14411494
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBldABO 
         Height          =   315
         Left            =   1215
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVolumn 
         Height          =   315
         Left            =   1215
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   630
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblTestDt 
         Height          =   315
         Left            =   1215
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   975
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblRsnNm 
         Height          =   315
         Left            =   1200
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3765
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblIrra 
         Height          =   315
         Left            =   1260
         TabIndex        =   75
         Top             =   4980
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFilter 
         Height          =   315
         Left            =   3180
         TabIndex        =   76
         Top             =   4980
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdNo 
         Height          =   315
         Left            =   5160
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2370
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblOrdSeq 
         Height          =   315
         Left            =   5685
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2370
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblRstSeq 
         Height          =   315
         Left            =   2445
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   975
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblDeliveryDt 
         Height          =   315
         Left            =   1215
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblDeliveryTm 
         Height          =   315
         Left            =   3540
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   22
         Left            =   3825
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "처방일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   23
         Left            =   135
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   4275
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Line Line5 
         X1              =   180
         X2              =   6140
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000009&
         X1              =   180
         X2              =   6140
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Irradiation"
         Height          =   180
         Left            =   240
         TabIndex        =   90
         Top             =   5070
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Filter"
         Height          =   180
         Left            =   2580
         TabIndex        =   89
         Top             =   5070
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방코드"
         Height          =   180
         Left            =   300
         TabIndex        =   88
         Top             =   5550
         Width           =   720
      End
      Begin VB.Label lblNewTestDiv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "N"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   840
         TabIndex        =   87
         Top             =   5940
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00C0C0FF&
         BorderWidth     =   2
         FillColor       =   &H00C0E0FF&
         Height          =   975
         Left            =   180
         Shape           =   4  '둥근 사각형
         Top             =   5400
         Width           =   6075
      End
   End
   Begin DRcontrol1.DrFrame DrFrame1 
      Height          =   8415
      Left            =   75
      TabIndex        =   91
      Top             =   75
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   14843
      Title           =   "환자정보"
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtBldNo 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1380
         TabIndex        =   99
         Top             =   1020
         Width           =   2070
      End
      Begin VB.CheckBox chkBarCode 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드 입력"
         Height          =   375
         Left            =   360
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdBldNo 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   97
         ToolTipText     =   "혈액의 반환, 폐기는 출고 후 24이내에만 가능합니다."
         Top             =   1020
         Width           =   350
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         Left            =   1380
         Style           =   2  '드롭다운 목록
         TabIndex        =   96
         Top             =   1425
         Width           =   2475
      End
      Begin VB.CheckBox chkExpire 
         BackColor       =   &H00DBE6E6&
         Caption         =   "자체폐기"
         Height          =   375
         Left            =   1920
         TabIndex        =   95
         TabStop         =   0   'False
         ToolTipText     =   "입고, 반환, Assign된 혈액만 자체폐기가 가능합니다."
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Refresh"
         Height          =   510
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   94
         Tag             =   "15101"
         Top             =   7380
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   165
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "과/병동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   165
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   360
         Left            =   1365
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   2835
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   4320
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2220
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblDeptCd 
         Height          =   360
         Left            =   1365
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2610
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   2175
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblWard 
         Height          =   360
         Left            =   3480
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin FPSpread.vaSpread tblDelivery 
         Height          =   3555
         Left            =   180
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   3720
         Width           =   7305
         _Version        =   196608
         _ExtentX        =   12885
         _ExtentY        =   6271
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   9
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS303.frx":0C2A
         UserResize      =   0
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblCompoCd 
         Height          =   345
         Left            =   4380
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   609
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblCompoNm 
         Height          =   345
         Left            =   5100
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblReaction 
         Height          =   315
         Left            =   3780
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   12640511
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Reaction"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblInfection 
         Height          =   315
         Left            =   3360
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3000
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BackColor       =   12640511
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "@"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   165
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "Component"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   165
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "혈액번호"
         Appearance      =   0
      End
      Begin VB.Label lblABO 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "AB(AB)+"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   30
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   5085
         TabIndex        =   115
         Top             =   2460
         Width           =   2565
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   7410
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   210
         X2              =   7410
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이 환자에게 출고된 내역"
         Height          =   180
         Left            =   210
         TabIndex        =   114
         Top             =   3450
         Width           =   1980
      End
      Begin VB.Label lblCompoCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
         ForeColor       =   &H00004080&
         Height          =   180
         Left            =   3900
         TabIndex        =   113
         Top             =   1500
         Width           =   90
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   1095
         Left            =   5040
         TabIndex        =   116
         Top             =   2220
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmBBS303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------------
' 혈액출고, 반환, 폐기, 회수 모두 이 화면에서 처리한다.
'---------------------------------------------------------------------------------------------
' 출고시 처리
' 1. 혈액번호를 입력하면,
'    가. XM결과 내역에서 이 혈액번호로 출고대기상태인 혈액을 모두 찾는다.
'        (  (rstval = 1 or stat = 1) and 혈액입고내역.stscd = assign인것  )
'    나. 같은 혈액번호의 혈액을 cboCompo에 담는다.
' 2. component를 선택을 하면,
'    가. 환자정보와 이 환자에게 출고된 혈액리스트를 조회한다.

Public mode As Long

Private Enum EMode
    modeDELIVERY = 0
    modeRETURN = 1
    modeEXPIRE = 2
    modeBAGRETURN = 3
End Enum
Private First As Boolean
Private onPgm As Boolean

Private modeMsg(3) As String

Private Enum TblColumn
    tcDELIVERYDT = 1
    tcBldNo
    tcCOMPONENT
    tcABO
    tcVOLUMN
    tcIRRADIATION
    tcFilter
    tcRETURN
    tcEXPIRE
End Enum

Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objBldList As clsPopUpList
Attribute objBldList.VB_VarHelpID = -1

Private Sub cboCompo_Click()
    Dim BldSrc As String
    Dim BldYY  As String
    Dim BldNo  As String
    
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
'        BldNo = Format(Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2), "######")
        BldNo = Format(Mid(txtBldNo, 5, 6), "00000#")
    Else
        BldSrc = medGetP(txtBldNo.Text, 1, "-")
        BldYY = medGetP(txtBldNo.Text, 2, "-")
        BldNo = Format(medGetP(txtBldNo.Text, 3, "-"), "00000#")
    End If
    lblCompoCd.Caption = medGetP(cboCompo.Text, 1, COL_DIV)
    lblCompoNm.Caption = medGetP(cboCompo.Text, 2, COL_DIV)
        
    Call SetBloodInfo(BldSrc, BldYY, BldNo, lblCompoCd.Caption)
'
'    Dim BldSrc As String
'    Dim BldYY  As String
'    Dim BldNo  As String
'
'    If chkBarCode.value = 1 Then
'        BldSrc = Mid(txtBldNo, 1, 2)
'        BldYY = Mid(txtBldNo, 3, 2)
'        BldNo = Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
'    Else
'        BldSrc = medGetP(txtBldNo, 1, "-")
'        BldYY = medGetP(txtBldNo, 2, "-")
'        BldNo = medGetP(txtBldNo, 3, "-")
'    End If
'    lblCompoCd.Caption = medGetP(cboCompo.Text, 1, " ")
'    lblCompoNm.Caption = medGetP(cboCompo.Text, 2, " ")
'
'    Call SetBloodInfo(BldSrc, BldYY, BldNo, lblCompoCd.Caption)
End Sub

Private Sub cboExpRsnCd_Click()
    Dim div As String
    
    On Error Resume Next
    div = medGetP(cboExpRsnCd.Text, 2, vbTab)
    lblExpBillDiv.Caption = IIf(div = 1, "Yes", "No")
End Sub


Private Sub chkBarCode_Click()
'    txtBldNo.SetFocus
End Sub

Private Sub chkExpire_Click()
    Clear
    ClearAll
    cboCompo.Clear
    lblCompoCnt = "0"
    txtBldNo = ""
    txtBldNo.SetFocus
End Sub

Private Sub cmdBagID_Click()
    
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "직원조회": .ColumnHeaderText = "사번;직원명"
        txtBagID.Text = "": lblBagNm.Caption = ""
        Call .LoadPopUp(GetSQLHisEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtBagID.Text = medGetP(.SelectedString, 1, ";")
            lblBagNm.Caption = medGetP(.SelectedString, 2, ";")
            
'            Call SetHisEmpToLisEmp(txtBagID.Text, lblBagNm.Caption)
        End If
    End With
    
    Set objMyList = Nothing
End Sub


Private Sub SetHisEmpToLisEmp(ByVal vEmpId As String, ByVal vEmpNm As String)
'HIS의 직원정보를 LIS에 입력한다. (직원마스터와 사용자 마스터에 적용)
    Dim strSQL As String
    Dim RS As Recordset
    
    On Error GoTo ErrTrap
    
    'LIS 직원 마스터에 입력
    strSQL = " select * from s2com006"
    strSQL = strSQL & " where empid='" & vEmpId & "'"
    
    DBConn.BeginTrans
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then 'lis 직원 마스터에 없는 경우
        strSQL = " insert into s2com006"
        strSQL = strSQL & " (empid, empnm) values"
        strSQL = strSQL & " ('" & vEmpId & "','" & vEmpNm & "')"
        
        DBConn.Execute strSQL
    End If
    
    'LIS 사용자 마스터에 입력
    strSQL = " select * from s2com010"
    strSQL = strSQL & " where loginid=''"
    
    Set RS = Nothing
    Set RS = New Recordset
        
    RS.Open strSQL, DBConn
    
    If RS.EOF Then 'LIS 사용자 마스터에 없는 경우
        strSQL = " insert into s2com010"
        strSQL = strSQL & " (loginid, loginpass,empid,logindesc, groupid) values"
        strSQL = strSQL & " ('" & vEmpId & "','2','" & vEmpId & "','" & vEmpNm & "','G002')"
        
        DBConn.Execute strSQL
    End If
    
    Set RS = Nothing
    
    DBConn.CommitTrans
    
    Exit Sub
    
ErrTrap:
    DBConn.RollbackTrans
    
End Sub

Private Sub cmdBldNo_Click()
    
'    With frmBloodFind
'        .mode = mode
'        If .mode = modeEXPIRE And chkExpire.value = 1 Then
'            .HosExp = True
'        End If
'        .Show vbModal
'        If .isSelected = True Then
'            If chkBarCode.value = 1 Then
'                txtBldNo = .BldSrc & .BldYY & .BldNo & "10"
'            Else
'                txtBldNo = .BldSrc & "-" & .BldYY & "-" & .BldNo
'            End If
'            txtBldNoLostFocus
''            cboCompo.Text = medComboFind(cboCompo, .Compo)
'        End If
'    End With
'    Set frmBloodFind = Nothing
    
    Dim objSql As clsBldDelivery
    Dim strSQL As String
    
    Set objSql = New clsBldDelivery
    
    Set objBldList = New clsPopUpList
    With objBldList
        .Connection = DBConn
        .ColumnHeaderText = "혈액번호;환자ID;환자명;제재"
        .FormHeight = 4125
        If mode = modeEXPIRE And chkExpire.value = 1 Then '자체 폐기용 혈액 조회
            .FormCaption = "폐기 대상 리스트"
            .SortColumn = 1
            .ColumnHeaderWidth = "1230.236;0;0;2085.166"
            .FormWidth = 3765
        Else
            .FormCaption = "출고 대상 리스트"
            .SortColumn = 3
            .ColumnHeaderWidth = "1230.236;915.0237;929.7639;2085.166"
            .FormWidth = 5715
        End If
        .ColumnHeaderAlign = "0;2;2;0"
        .AutoGap = True
        If mode = modeEXPIRE And chkExpire.value = 1 Then '자체 폐기용 혈액 조회
            .SqlStmt = objSql.GetSQLBloodList(4)
        Else
            .SqlStmt = objSql.GetSQLBloodList(mode)
        End If
        .LoadPopUp
    End With
    
    Set objBldList = Nothing
    Set objSql = Nothing
End Sub

Private Sub cmdClear_Click()
    ClearAll
    Clear
    txtRcvId = ""
    lblRcvNm.Caption = ""
    txtBldNo.Text = ""
    cboCompo.Clear
    lblCompoCnt = "0"
    txtBldNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExpId_Click()
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "직원조회": .ColumnHeaderText = "사번;직원명"
        txtExpId.Text = "": lblExpNm.Caption = ""
        Call .LoadPopUp(GetSQLHisEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtExpId.Text = medGetP(.SelectedString, 1, ";")
            lblExpNm.Caption = medGetP(.SelectedString, 2, ";")
            
'            Call SetHisEmpToLisEmp(txtExpId.Text, lblExpNm.Caption)
        End If
    End With
    
    Set objMyList = Nothing
'    With objMyList
'        .BackColor = Me.BackColor
'        .Caption = "직원조회": .HeadName = "사번,직원명"
'        .Width = .Width + 300: .ColSize(0) = 1000
'        txtExpId.Text = "": lblExpNm.Caption = ""
'        Call .ListPop(GetEmpListSQL, 2350, 7650)
'        If .SelectedString <> "" Then
'            txtExpId.Text = medGetP(.SelectedString, 1, ";")
'            lblExpNm.Caption = medGetP(.SelectedString, 2, ";")
'
'        End If
'
'    End With

End Sub



Private Sub cmdRefresh_Click()
    Call SetDeliveryHistory(lblPtId.Caption)
End Sub

Private Sub cmdRetID_Click()
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "직원조회": .ColumnHeaderText = "사번;직원명"
        txtRetID.Text = "": lblRetNm.Caption = ""
        Call .LoadPopUp(GetSQLHisEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtRetID.Text = medGetP(.SelectedString, 1, ";")
            lblRetNm.Caption = medGetP(.SelectedString, 2, ";")
            
'            Call SetHisEmpToLisEmp(txtRetID.Text, lblRetNm.Caption)
        End If
    End With

'    Set objMyList = New clsS2DLP
'    With objMyList
'        .BackColor = Me.BackColor
'        .Caption = "직원조회": .HeadName = "사번,직원명"
'        .Width = .Width + 300: .ColSize(0) = 1000
'        txtRetID.Text = "": lblRetNm.Caption = ""
'        Call .ListPop(GetEmpListSQL, 2350, 7650)
'        If .SelectedString <> "" Then
'            txtRetID.Text = medGetP(.SelectedString, 1, ";")
'            lblRetNm.Caption = medGetP(.SelectedString, 2, ";")
'
'        End If
'
'    End With

    Set objMyList = Nothing
End Sub

Private Sub cmdRcvId_Click()
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "직원조회": .ColumnHeaderText = "사번;직원명"
        txtRcvId.Text = "": lblRcvNm.Caption = ""
        Call .LoadPopUp(GetSQLHisEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtRcvId.Text = medGetP(.SelectedString, 1, ";")
            lblRcvNm.Caption = medGetP(.SelectedString, 2, ";")
            
'            Call SetHisEmpToLisEmp(txtRcvId.Text, lblRcvNm.Caption)
        End If
    End With
    Set objMyList = Nothing
'
'    Set objMyList = New clsS2DLP
'    With objMyList
'        .BackColor = Me.BackColor
'        .Caption = "직원조회": .HeadName = "사번,직원명"
'        .Width = .Width + 300: .ColSize(0) = 1000
'        txtRcvId.Text = "": lblRcvNm.Caption = ""
'        Call .ListPop(GetEmpListSQL, 2350, 7650)
'        If .SelectedString <> "" Then
'            txtRcvId.Text = medGetP(.SelectedString, 1, ";")
'            lblRcvNm.Caption = medGetP(.SelectedString, 2, ";")
'        End If
'    End With
'
End Sub


Private Sub cmdSave_Click()

    If mode = EMode.modeEXPIRE Then
        If lblBldABO.Caption = "" Or txtBldNo = "" Then
            MsgBox "작업진행 혈액을 선택하신후 진행 하십시요.", vbCritical + vbOKOnly, "혈액선택"
            Exit Sub
        End If
    Else
        If txtBldNo = "" Or lblPtId.Caption = "" Then
            MsgBox "작업진행 혈액을 선택하신후 진행 하십시요.", vbCritical + vbOKOnly, "혈액선택"
            Exit Sub
        End If
    End If
    Select Case mode
        Case EMode.modeDELIVERY:
            If BldDelivery = True Then ClearAll: txtBldNo = "": txtBldNo.SetFocus                                   '출고
        Case EMode.modeRETURN:
            If BldReturn = True Then
                '반환요청서
                Call PrintBloodReturn
            
                Call Clear
                Call ClearAll
                txtBldNo = ""
                txtBldNo.SetFocus   '반환
                MsgBox "반환처리 되었습니다.", vbInformation + vbOKOnly, "혈액반환"
            End If
        Case EMode.modeEXPIRE:
            If BldExpire = True Then
                Call Clear
                Call ClearAll
                txtBldNo = ""
                txtBldNo.SetFocus   '폐기
                MsgBox "폐기처리 되었습니다.", vbInformation + vbOKOnly, "혈액폐기"
            End If
        Case EMode.modeBAGRETURN:
            If BldBag = True Then
                Call Clear
                Call ClearAll
                txtBldNo = ""
                txtBldNo.SetFocus   '회수
                MsgBox "회수처리 되었습니다.", vbInformation + vbOKOnly, "혈액Bag회수"
            End If
    End Select
    'txtBldNo = ""
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    
    If First = False Then Exit Sub
    
    First = False
    
    fraMode(1).Left = fraMode(0).Left
    fraMode(1).Top = fraMode(0).Top
    fraMode(2).Left = fraMode(0).Left
    fraMode(2).Top = fraMode(0).Top
    fraMode(3).Left = fraMode(0).Left
    fraMode(3).Top = fraMode(0).Top
    
    fraMode(0).Visible = False
    fraMode(1).Visible = False
    fraMode(2).Visible = False
    fraMode(3).Visible = False
    
    fraMode(mode).Visible = True
'    cmdF.Visible = True'전주 예수병원은 필터반환/폐기를 안하는 것으로 한다.
    cmdF.Visible = False
    Select Case mode
        Case EMode.modeBAGRETURN:
            Me.Caption = "혈액BAG회수"
            cmdSave.Caption = "회수(&S)"
        Case EMode.modeEXPIRE:
            Me.Caption = "혈액폐기"
            cmdSave.Caption = "폐기(&S)"
            cmdF.Caption = "필터폐기(&D)"
            LisLabel3.Caption = "필터폐기"
            cmdFSave.Caption = "필터폐기"
        Case EMode.modeRETURN:
            Me.Caption = "혈액반환"
            cmdSave.Caption = "반환(&S)"
            cmdF.Caption = "필터반환(&D)"
            LisLabel3.Caption = "필터반환"
            cmdFSave.Caption = "필터반환"
        Case Else
            cmdF.Visible = False
    End Select
    
    If mode = EMode.modeEXPIRE Then
        chkExpire.Visible = True
    Else
        chkExpire.Visible = False
    End If

    If mode = EMode.modeDELIVERY Then
        lblDelivery.Visible = False
        lblDeliveryDt.Visible = False
        lblDelivery1.Visible = False
        lblDeliveryTm.Visible = False
        lblDeliverySeq.Visible = False
    Else
        lblDelivery.Visible = True
        lblDeliveryDt.Visible = True
        lblDelivery1.Visible = True
        lblDeliveryTm.Visible = True
        lblDeliverySeq.Visible = True
    End If
    Clear
    ClearAll
    
End Sub

Private Sub Form_Load()
    Call SetExpireRsn

    First = True
    modeMsg(0) = "이미 출고되었거나, 준비되지 않은 혈액입니다"
    modeMsg(1) = "반환 혹은 폐기되었거나, 준비되지 않은 혈액입니다."
    modeMsg(2) = "반환 혹은 폐기되었거나, 준비되지 않은 혈액입니다."
    modeMsg(3) = "반환 혹은 폐기되었거나, 준비되지 않은 혈액입니다."
    
    chkBarCode.value = 1
    ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub lblDeliveryDt_Click()
    '혈액폐기,반환,회수시에만 사용
End Sub

Private Sub lblDeliverySeq_Click()
    '혈액폐기,반환,회수시에만 사용
End Sub


Private Sub lstNewTest_ItemCheck(Item As Integer)
    Dim i As Long
    
    If onPgm = True Then Exit Sub
    
    onPgm = True
    With lstNewTest
        For i = 0 To .ListCount - 1
            If i <> Item Then
                .Selected(i) = False
            End If
        Next i
    End With
    onPgm = False
End Sub

Private Sub objBldList_SelectedItem(ByVal pSelectedItem As String)
    If chkBarCode.value = 1 Then
        txtBldNo.Text = Replace(medGetP(pSelectedItem, 1, ";"), "-", "")
    Else
        txtBldNo.Text = medGetP(pSelectedItem, 1, ";")
    End If
    If txtBldNo.Text = "" Then Exit Sub
    txtBldNoLostFocus
End Sub

Private Sub txtBagID_Change()
    If lblBagNm.Caption <> "" Then
        lblBagNm.Caption = ""
    End If
End Sub

Private Sub txtBagID_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtBagID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBagID_Validate(Cancel As Boolean)
    Dim strBagNm As String
    
    If txtBagID.Text = "" Then Exit Sub
    
    strBagNm = GetEmpNm(txtBagID.Text)
    If strBagNm = "" Then
        Cancel = True
        MsgBox "등록되지 않은 사용자입니다.", vbExclamation
    Else
        lblBagNm.Caption = strBagNm
    End If
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtBldNo_Change()
    Dim lngLen As Long
    
    If lblCompoCd.Caption <> "" Then
        ClearAll
        Clear
        txtBldNo.tag = ""
        txtRcvId.Text = ""
        lblRcvNm.Caption = ""
        cboCompo.Clear
        lblCompoCnt.Caption = "0"
    End If
    
    If chkBarCode.value = 1 Then Exit Sub
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    SendKeys "{Home}+{End}"
    txtBldNo.tag = txtBldNo
End Sub

Private Sub txtBldNo_LostFocus()
    If chkBarCode <> 1 Then
        If Len(Trim(txtBldNo)) <= 6 Then Exit Sub
    End If
    If txtBldNo.tag = txtBldNo Then Exit Sub
    If txtBldNo.Text = "" Then Exit Sub
    Me.MousePointer = 11
    '--------- 자료조회 ----------
    Call txtBldNoLostFocus
    Me.MousePointer = 0
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
    If chkBarCode.value = 1 Then Exit Sub
    If Len(txtBldNo.Text) <> 3 Or Len(txtBldNo.Text) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
    
End Sub

Private Sub txtBldNoLostFocus()
    Dim DrRS           As Recordset
    Dim objBldDelivery As clsBldDelivery
    Dim BldSrc  As String
    Dim BldYY   As String
    Dim BldNo   As String
    
    Dim i As Long

    If chkBarCode.value = 1 Then
        If Len(txtBldNo) < 7 Then
            MsgBox "혈액번호를 확인하세요.", vbInformation + vbOKOnly, "혈액번호오류"
            txtBldNo.SelStart = 0
            txtBldNo.SelLength = Len(txtBldNo)
            txtBldNo.SetFocus
            Exit Sub
        End If
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Format(Mid(txtBldNo.Text, 5, 6), "0#####")
    Else
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = Format(medGetP(txtBldNo, 3, "-"), "0#####")
    End If
    
    Clear
    ClearAll
    
    If BldSrc = "" Or BldYY = "" Or BldNo = "" Then Exit Sub
    
    Set objBldDelivery = New clsBldDelivery
    
    '자체폐기시
    If mode = EMode.modeEXPIRE Then
        If chkExpire.value = 1 Then
            Set DrRS = objBldDelivery.GetExpireHospital(BldSrc, BldYY, BldNo)
        Else
            Set DrRS = objBldDelivery.GetBloodCompoList(BldSrc, BldYY, BldNo, mode)
        End If
    Else
        Set DrRS = objBldDelivery.GetBloodCompoList(BldSrc, BldYY, BldNo, mode)
    End If
    
    If DrRS Is Nothing Then
        Set objBldDelivery = Nothing
        Exit Sub
    End If
    
    With DrRS
        cboCompo.Clear
        lblCompoCnt = .RecordCount
        If .RecordCount = 1 Then
            cboCompo.AddItem .Fields("compocd").value & "" & COL_DIV & .Fields("componm").value & ""
            lblCompoCd.Caption = .Fields("compocd").value & ""
            lblCompoNm.Caption = .Fields("componm").value & ""
            cboCompo.ListIndex = 0
        ElseIf .RecordCount > 1 Then
            For i = 1 To .RecordCount
                cboCompo.AddItem .Fields("compocd").value & "" & COL_DIV & .Fields("componm").value & ""
                .MoveNext
            Next i
        Else
            MsgBox modeMsg(mode), vbCritical, Me.Caption
            txtBldNo = ""
            txtBldNo.SetFocus
        End If
    End With
    
    Set DrRS = Nothing
    Set objBldDelivery = Nothing
End Sub

Private Function BldBag() As Boolean
    Dim objBldDelivery As clsBldDelivery
    Dim BldSrc     As String
    Dim BldYY      As String
    Dim BldNo      As String
    Dim CompoCd    As String
    Dim deliverydt As String
    Dim Bagid      As String
    Dim Bagrcvid   As String
    
    
    If txtBagID.Text = "" Then
        MsgBox "회수요청자를 선택하신 후 회수처리 하십시오.", vbInformation + vbOKOnly, "회수요청자 선택"
        Exit Function
    End If
    '-----------------------------------------------------------------------
    '출고된 혈액의 회수처리를 한다.
    '-----------------------------------------------------------------------
    '혈액출고내역(BBS402)에 bagfg<-'1',bagdt<-'오늘일자'를 update한다.
    '혈액입고내역(BBS401)에 stscd를 BBSBloodStatus.stsBAG으로 만든다.
    '-----------------------------------------------------------------------
    
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
    Else
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = medGetP(txtBldNo, 3, "-")
    End If
    Bagid = ObjMyUser.EmpId
    Bagrcvid = txtBagID.Text
    deliverydt = Format(lblDeliveryDt.Caption, PRESENTDATE_FORMAT)
    
    Set objBldDelivery = New clsBldDelivery
'    Set objBldDelivery.DrDB = DBConn
    
    BldBag = objBldDelivery.BldBag(BldSrc, BldYY, BldNo, lblCompoCd.Caption, deliverydt, lblDeliverySeq.Caption, Bagid, Bagrcvid)
    
    Set objBldDelivery = Nothing
End Function

Private Function GetOrdCd(ordcd As String) As Boolean
    Dim i As Long
    
    If lstNewTest.ListCount < 1 Then
        ordcd = ""
        GetOrdCd = True
    Else
        For i = 0 To lstNewTest.ListCount - 1
            If lstNewTest.Selected(i) = True Then
                ordcd = medGetP(lstNewTest.List(i), 1, " ")
                Exit For
            End If
        Next i
        If ordcd = "" Then
            MsgBox "수가 계산용 코드가 선택되지 않았읍니다.", vbCritical, Me.Caption
            GetOrdCd = False
        Else
            GetOrdCd = True
        End If
    End If
End Function

Private Function BldExpire() As Boolean
    Dim BldSrc     As String
    Dim BldYY      As String
    Dim BldNo      As String
    Dim CompoCd    As String
    Dim deliverydt As String
'수가계산내역의 처방코드
    Dim ordcd      As String
    Dim expid As String '폐기 처리 당사자.(혈액은행 담당)
    Dim exprsncd   As String
    Dim expbilldiv As String
    
    Dim objBldDelivery As clsBldDelivery
    
    If txtExpId.Text = "" Then
        MsgBox "폐기요청자를 선택하신 후 폐기 처리 하십시오.", vbInformation + vbOKOnly, "폐기요청자 선택"
        Exit Function
    End If
    '-----------------------------------------------------------------------
    '출고된 혈액의 폐기처리를 한다.
    '-----------------------------------------------------------------------
    '혈액출고내역(BBS402)에 expfg<-'1',expdt<-'오늘일자'를 update한다.
    '혈액입고내역(BBS401)에 stscd를 BBSBloodStatus.stsEXPIRE로 만든다.
    '-----------------------------------------------------------------------
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Mid(txtBldNo, 5, 6)
    Else
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = medGetP(txtBldNo, 3, "-")
    End If
    deliverydt = Format(lblDeliveryDt.Caption, PRESENTDATE_FORMAT)
    
    If GetOrdCd(ordcd) = False Then
        BldExpire = False
        Exit Function
    End If
    
    Set objBldDelivery = New clsBldDelivery
    
    objBldDelivery.WorkSeq = lblWorkSeq.Caption
    
    expid = ObjMyUser.EmpId
    exprsncd = medGetP(cboExpRsnCd.Text, 1, " ")
    expbilldiv = medGetP(cboExpRsnCd.Text, 2, vbTab)
    
    If chkExpire.value = 1 And expbilldiv = "1" Then
        MsgBox "혈액은행내 자체폐기인경우는 환자부담으로 폐기할수 없습니다.", vbCritical + vbOKOnly, "혈액폐기"
        Exit Function
    End If
    
    If exprsncd = "" Then
        MsgBox "폐기사유를 선택한후 진행하십시요.", vbCritical + vbOKOnly, "혈액폐기"
        Set objBldDelivery = Nothing
        Exit Function
    End If
    
    If chkExpire.value = 1 Then
        If lblTestDt.Caption <> "" Then
        '자체폐기시(Assign 된 혈액의 폐기와, 그렇지 않은 혈액의 두가지 종류가 있다.
            BldExpire = objBldDelivery.BldHosExpire(BldSrc, BldYY, BldNo, lblCompoCd.Caption, txtExpId.Text, _
                                                  expid, exprsncd, expbilldiv, txtExpRsnRmk.Text, False)
        Else
        'Assign 되지 않은 혈액
            BldExpire = objBldDelivery.BldHosExpire(BldSrc, BldYY, BldNo, lblCompoCd.Caption, txtExpId.Text, _
                                                  expid, exprsncd, expbilldiv, txtExpRsnRmk.Text)
        End If
    Else
    '출고에 의한 폐기시
        BldExpire = objBldDelivery.BldExpire(BldSrc, BldYY, BldNo, lblCompoCd.Caption, _
                                             deliverydt, lblDeliverySeq.Caption, _
                                             txtExpId.Text, _
                                             expid, exprsncd, expbilldiv, txtExpRsnRmk.Text, lblPtId.Caption, ordcd)
    End If
    
    Set objBldDelivery = Nothing
End Function
Public Function BldHourChk(ByVal deliverydt As String, ByVal deliverytm As String) As Boolean
'혈액출고 시간을 체크한다.(30시간)
    Dim objBldDelivery As clsBldDelivery
    Dim lngStoreHour   As Long
    Dim strCompare     As String
    Dim strCompare1    As String
    Dim Possible       As Long
    
    Set objBldDelivery = New clsBldDelivery
'    Set objBldDelivery.DrDB = DBConn
    
    Possible = objBldDelivery.BldReturnHour
    
    strCompare1 = Format(GetSystemDate, PRESENTDATE_FORMAT) & Format(GetSystemDate, "HHmm")
    strCompare1 = Format(strCompare1, "####-##-## ##:##")
       
    strCompare = deliverydt & Mid(deliverytm, 1, 4)
    strCompare = Format(strCompare, "####-##-## ##:##")
        
    lngStoreHour = CLng(DateDiff("n", strCompare, strCompare1))
    If lngStoreHour <= Possible Then
        BldHourChk = True
    Else
        BldHourChk = False
    End If
End Function
Private Function BldReturn() As Boolean
    Dim deliverydt As String
    Dim deliverytm As String
    Dim BldSrc     As String
    Dim BldYY      As String
    Dim BldNo      As String
    Dim CompoCd    As String
    Dim ordcd      As String
    Dim strTmp     As String
    
    Dim objBldDelivery As clsBldDelivery
    
    If txtRetID.Text = "" Then
        MsgBox "반환요청자를 입력하신후 반환처리 하십시오.", vbInformation + vbOKOnly, "반환요청자 선택"
        Exit Function
    End If
    '-----------------------------------------------------------------------
    '출고된 혈액의 반환처리를 한다.
    '-----------------------------------------------------------------------
    '혈액출고내역(BBS402)에 retfg<-'1',retdt<-'오늘일자'를 update한다.
    '혈액입고내역(BBS401)에 stscd를 BBSBloodStatus.stsRETURN으로 만든다.
    '-----------------------------------------------------------------------
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Mid(txtBldNo, 5, 6)
    Else
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = medGetP(txtBldNo, 3, "-")
    End If
    deliverydt = Format(lblDeliveryDt.Caption, PRESENTDATE_FORMAT)
    If GetOrdCd(ordcd) = False Then
        BldReturn = False
        Exit Function
    End If
'    ordCd = medGetP(cboNewTest.Text, 1, " ")
    deliverytm = Mid(lblDeliveryTm.Caption, 1, 2) & Mid(lblDeliveryTm.Caption, 4)

    '반환시 30시간 체크를 하여 이미 경과 한경우는 반환할수 없다.
    If BldHourChk(deliverydt, deliverytm) = False Then
        strTmp = MsgBox("반환 가능 시간이 지난 혈액입니다. 반환처리 하시겠습니까?", vbYesNo + vbExclamation + vbDefaultButton2, "혈액반환")
        If strTmp = vbNo Then
            Exit Function
        End If
        '2001-11-27수정
        If Trim(txtRetRmk.Text) = "" Then
            MsgBox "반환사유를 반드시 입력하십시오", vbExclamation, "혈액반환"
            txtRetRmk.SetFocus
            Exit Function
        End If
    End If
    '반환자,반환수령자,리마크 추가(2001,02,09)
    Dim Retid As String
    Dim Retrcvid As String
    Dim Retrmk As String
    
    Retid = ObjMyUser.EmpId
    Retrcvid = txtRetID.Text
    Retrmk = txtRetRmk.Text

    Set objBldDelivery = New clsBldDelivery
    
    objBldDelivery.WorkSeq = lblWorkSeq.Caption
    
    BldReturn = objBldDelivery.BldReturn(BldSrc, BldYY, BldNo, lblCompoCd.Caption, deliverydt, _
                                         lblDeliverySeq.Caption, ordcd, lblPtId.Caption, _
                                         Retid, Retrcvid, Retrmk)
    
    Set objBldDelivery = Nothing
End Function

Private Function BldDelivery() As Boolean
    Dim today As Date
    Dim tmpStr As String
    Dim objBldDelivery As clsBldDelivery
    Dim ordcd As String
    
    '----------------------------------------------------------------------------
    '값을 String으로 넘긴다
    '----------------------------------------------------------------------------
    'bldsrc,bldyy,bldno,compocd,deliverydt,deliveryseq,deliverytm,deliveryid
    'rcvid,ptid,orddt,ordno,ordseq,rstseq,ordcd,localcd,irrafg,filter,retfg,expfg
    '----------------------------------------------------------------------------
        
    If GetOrdCd(ordcd) = False Then
        BldDelivery = False
        Exit Function
    End If
        
        
        
    today = GetSystemDate
    If chkBarCode.value = 1 Then
        tmpStr = Mid(txtBldNo, 1, 2) & COL_DIV & _
                Mid(txtBldNo, 3, 2) & COL_DIV & _
                Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2) & COL_DIV
    Else
        tmpStr = medGetP(txtBldNo, 1, "-") & COL_DIV & _
                medGetP(txtBldNo, 2, "-") & COL_DIV & _
                medGetP(txtBldNo, 3, "-") & COL_DIV
    End If
    
    tmpStr = tmpStr & _
                    lblCompoCd.Caption & COL_DIV & _
                    Format(today, PRESENTDATE_FORMAT) & COL_DIV & _
                    "" & COL_DIV & _
                    Format(today, "HHMMSS") & COL_DIV & _
                    ObjMyUser.EmpId & COL_DIV & _
                    txtRcvId & COL_DIV & _
                    C_WORKAREA & COL_DIV & _
                    medGetP(lblAccNo.Caption, 1, "-") & COL_DIV & _
                    medGetP(lblAccNo.Caption, 2, "-") & COL_DIV & _
                    lblRstSeq.Caption & COL_DIV & _
                    ordcd & COL_DIV & _
                    "" & COL_DIV & _
                    chkIrra.value & COL_DIV & _
                    chkFilter.value & COL_DIV & _
                    "" & COL_DIV & _
                    "" & COL_DIV & _
                    lblPtId.Caption & COL_DIV & _
                    Format(lblOrdDt.Caption, PRESENTDATE_FORMAT) & COL_DIV & _
                    lblOrdNo.Caption & COL_DIV & _
                    lblOrdSeq.Caption
                    
    Set objBldDelivery = New clsBldDelivery
    BldDelivery = objBldDelivery.BldDelivery(tmpStr)
    Set objBldDelivery = Nothing
End Function

Private Sub SetNewTest(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As String, ByVal CompoCd As String, ByVal volume As String, ByVal TestDiv As String)
    Dim Cnt As Long
    Dim aryOrdCd() As String
    Dim today As Date
    Dim objBldDelivery As clsBldDelivery
    Dim i As Long
    
    today = GetSystemDate
    
    Set objBldDelivery = New clsBldDelivery
'    Set objBldDelivery.DrDB = DBConn
    Cnt = objBldDelivery.GetOrdCd(BldSrc, BldYY, BldNo, CompoCd, Format(today, PRESENTDATE_FORMAT), volume, TestDiv, aryOrdCd)
    Set objBldDelivery = Nothing
    
'    cboNewTest.Clear
    lstNewTest.Clear
    If Cnt > 0 Then
        For i = 1 To Cnt
            lstNewTest.AddItem aryOrdCd(i - 1)
        Next i
        onPgm = True
        If lstNewTest.ListCount = 1 Then lstNewTest.Selected(0) = True
        onPgm = False
'        cboNewTest.ListIndex = 0
    End If
End Sub

Private Sub SetTransRsn(ByVal PtId As String, ByVal orddt As String, ByVal ordno As String)
    Dim objQOrder As clsQueryOrder
    
    Set objQOrder = New clsQueryOrder
    lblRsnNm.Caption = objQOrder.GetTransReason(PtId, orddt, ordno)
    Set objQOrder = Nothing
End Sub

Private Sub SetDeliveryHistory(ByVal PtId As String)
    Dim i As Long
    Dim DrRS As Recordset
    Dim objBldDelivery As clsBldDelivery
    
    tblDelivery.MaxRows = 0
    
    If PtId = "" Then Exit Sub
    
    Set objBldDelivery = New clsBldDelivery
'    Set objBldDelivery.DrDB = DBConn
    Set DrRS = objBldDelivery.GetDeliveryHistory(PtId)
    If DrRS Is Nothing Then
        Set objBldDelivery = Nothing
        Exit Sub
    End If
    
    With tblDelivery
        If DrRS.RecordCount > 0 Then
            For i = 1 To DrRS.RecordCount
                'If i > DrRS.RecordCount Then Exit For
                .MaxRows = i
                .Row = i
                
                .Col = TblColumn.tcDELIVERYDT:  .value = Format(DrRS.Fields("deliverydt").value & "", "####-##-##")
                .Col = TblColumn.tcBldNo:       .value = DrRS.Fields("bldsrc").value & "" & "-" & DrRS.Fields("bldyy").value & "" & "-" & Format(DrRS.Fields("bldno").value & "", "0#####")
                .Col = TblColumn.tcCOMPONENT:   .value = DrRS.Fields("componm").value & ""
                .Col = TblColumn.tcABO:         .value = DrRS.Fields("abo").value & "" & DrRS.Fields("rh").value & ""
                .Col = TblColumn.tcVOLUMN:      .value = DrRS.Fields("volumn").value & ""
                .Col = TblColumn.tcIRRADIATION: .value = DrRS.Fields("irrafg").value & ""
                .Col = TblColumn.tcFilter:      .value = DrRS.Fields("filter").value & ""
                .Col = TblColumn.tcRETURN:      .value = DrRS.Fields("retfg").value & ""
                .Col = TblColumn.tcEXPIRE:      .value = DrRS.Fields("expfg").value & ""
                
                DrRS.MoveNext
            Next i
        End If
    End With
    
    Set DrRS = Nothing
    Set objBldDelivery = Nothing
End Sub

Private Sub DetailSearch(PtId As String)
'혈액형,부작용,감염정보
    Dim ObjABO As New clsABO
    Dim objinfection As New clsInfection
    Dim objReaction As New clsReaction
    
    With ObjABO
        .PtId = PtId
        .GetABO
        lblABO.Caption = .ABO & .Rh
    End With
    With objinfection
        .PtId = PtId
        .GetInfection
        If .Infection = True Then
            lblInfection.Visible = True
        Else
            lblInfection.Visible = False
        End If
    End With
    
    With objReaction
        .PtId = PtId
        .GetReaction
        If .Reaction = True Then
            lblReaction.Visible = True
        Else
            lblReaction.Visible = False
        End If
    End With
    
    
    Set objReaction = Nothing
    Set objinfection = Nothing
    Set ObjABO = Nothing
End Sub

Private Sub SetBloodInfo(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As String, ByVal CompoCd As String)
    Dim objBldDelivery As clsBldDelivery
    Dim isOK           As Boolean
    Dim DrRS           As Recordset
    Dim RS             As Recordset
    Dim strSDA         As String            'Sex/Birth/Age
    
    
    DoEvents
    
    Set objBldDelivery = New clsBldDelivery
    '혈액자체폐기시
    If mode = EMode.modeEXPIRE Then
        If chkExpire.value = 1 Then
            Set RS = objBldDelivery.GetXMExpireHospital(BldSrc, BldYY, BldNo, CompoCd)
'            Clear
'            ClearAll
            If RS.RecordCount > 0 Then
            '이미 xm 검사가 진행된 혈액의 정보.
                Set DrRS = objBldDelivery.GetXMExpireOrdInfo(BldSrc, BldYY, BldNo, CompoCd)
                With DrRS
                    If Not DrRS.EOF Then
                        lblPtId.Caption = .Fields("ptid").value & ""
                        lblPtNm.Caption = .Fields("ptnm").value & ""
                        '혈액형,간염정보,부작용 정보를 조회한다.
                        DetailSearch lblPtId.Caption
                        
                        lblCompoCd.Caption = .Fields("compocd").value & ""
                        lblBldABO.Caption = .Fields("abo").value & "" & DrRS.Fields("rh").value & ""
                        lblVolumn.Caption = .Fields("volumn").value & ""
                        
                        strSDA = SDA_String(.Fields("ssn").value) & ""
                        lblSexAge.Caption = medGetP(strSDA, 1, COL_DIV) & "/" & medGetP(strSDA, 3, COL_DIV)
                        
                        lblDeptCd.Caption = .Fields("deptcd").value & ""
                        lblDeptNm.Caption = .Fields("deptnm").value & ""
                        lblWard.Caption = .Fields("wardid").value & "" & "-" & .Fields("hosilid").value & "" & "-" & .Fields("bedid").value & ""
                        
                        lblAccNo.Caption = Trim(.Fields("accdt").value & "") & "-" & .Fields("accseq").value & ""
                        lblOrdNo.Caption = .Fields("ordno").value & ""
                        lblOrdSeq.Caption = .Fields("ordseq").value & ""
                        
                        lblTestCd.Caption = .Fields("ordcd").value & ""
                        lblTestNm.Caption = .Fields("testnm").value & ""
                        lblTestDt.Caption = Format(.Fields("vfydt").value & "", "####-##-##")
                        lblRstSeq.Caption = .Fields("rstseq").value & ""
                        
                        lblIrra.Caption = .Fields("irradfg").value & ""
                        lblFilter.Caption = .Fields("filterfg").value & ""
                        lblUnitQty.Caption = .Fields("unitqty").value & ""
                        lblReqDtTm.Caption = Format(.Fields("reqdt").value & "", "####-##-##") & " " & Format(.Fields("reqtm").value & "", "##:##:##")
                        lblOrdDt.Caption = Format(.Fields("orddt").value & "", "####-##-##")
                        lblRsnNm.Caption = ""
                        
                        txtrmk = .Fields("mesg").value & ""
                        Call SetTransRsn(.Fields("ptid").value & "", .Fields("orddt").value & "", .Fields("ordno").value & "")
                        tblDelivery.MaxRows = 0
                    End If
                End With
                
            Else
                Set DrRS = objBldDelivery.GetExpireBloodInfo(BldSrc, BldYY, BldNo, CompoCd)
                If Not DrRS.EOF Then
                    lblBldABO.Caption = DrRS.Fields("abo").value & "" & DrRS.Fields("rh").value & ""
                    lblVolumn.Caption = DrRS.Fields("volumn").value & ""
                    lblCompoCd.Caption = DrRS.Fields("compocd").value & ""
                    lblCompoNm.Caption = DrRS.Fields("componm").value & ""
                End If
            End If
            Set objBldDelivery = Nothing
            Set DrRS = Nothing
            Exit Sub
        Else
            Set DrRS = objBldDelivery.GetBloodInfo(BldSrc, BldYY, BldNo, CompoCd, mode)
        End If
    Else
        Set DrRS = objBldDelivery.GetBloodInfo(BldSrc, BldYY, BldNo, CompoCd, mode)
    End If
    
    If DrRS Is Nothing Then
        Set objBldDelivery = Nothing
        Exit Sub
    End If
    
    With DrRS
        If DrRS.RecordCount < 1 Then
            MsgBox "내역을 찾을 수 없습니다", vbCritical, Me.Caption
            Call Clear
        Else
            lblPtId.Caption = .Fields("ptid").value & ""
            lblPtNm.Caption = GetPtNm(.Fields("ptid").value & "")
            '혈액형,간염정보,부작용 정보를 조회한다.
            Call DetailSearch(lblPtId.Caption)
            
            Call GetBBS_Ptinfo(lblPtId.Caption, strSDA)
            strSDA = SDA_String(strSDA)
            lblSexAge.Caption = medGetP(strSDA, 1, COL_DIV) & "/" & medGetP(strSDA, 3, COL_DIV)
            
            lblDeptCd.Caption = .Fields("deptcd").value & ""
            lblDeptNm.Caption = .Fields("deptcd").value & ""
                        
            Dim strDeptNm As String
                        
            strDeptNm = GetDeptNm(lblDeptCd.Caption)
            If strDeptNm <> "" Then lblDeptNm.Caption = strDeptNm
            
'            If ObjBBSComCode.DeptCd.Exists(lblDeptCd.Caption) Then
'                ObjBBSComCode.DeptCd.KeyChange lblDeptCd.Caption
'                lblDeptNm.Caption = ObjBBSComCode.DeptCd.Fields("deptnm")
'            End If
            
            lblWard.Caption = .Fields("wardid").value & "" & "-" & .Fields("hosilid").value & "" & "-" & .Fields("bedid").value & ""
            
            tblDelivery.MaxRows = 0
            
            Call SetDeliveryHistory(lblPtId.Caption)
            
            '-----------------------------------------------
            '혈액출고는 Assign된 혈액을 기준으로 작업
            '혈액반환은 출고된 혈액을 기준으로 작업
            '혈액폐기,혈액회수도 출고된 혈액을 기준으로 작업
            '-----------------------------------------------
            ' BBS401의 stscd로 판단.
            ' 0.대기 1.Assign 2.출고 3.폐기 4.회수
            '-----------------------------------------------
            Select Case mode
                Case EMode.modeDELIVERY:
                    isOK = (.Fields("stscd").value & "" = BBSBloodStatus.stsASSIGN)
                Case Else:
                    isOK = (.Fields("stscd").value & "" = BBSBloodStatus.stsDELIVERY)
            End Select
            
            If isOK Then
                lblBldABO.Caption = .Fields("abo").value & "" & .Fields("rh").value & ""
                lblVolumn.Caption = .Fields("volumn").value & ""
                lblTestDt.Caption = Format(.Fields("vfydt").value & "", "####-##-##")
                lblRstSeq.Caption = .Fields("rstseq").value & ""
                lblWorkSeq.Caption = .Fields("workseq").value & ""
                If mode <> EMode.modeDELIVERY Then
                    lblDeliveryDt.Caption = Format(.Fields("deliverydt").value & "", "####-##-##")
                    lblDeliveryTm.Caption = Format(Mid(.Fields("deliverytm").value & "", 1, 4), "00:00")
                    lblDeliverySeq.Caption = .Fields("deliveryseq").value & ""
                End If
                
                picStat.Visible = ((.Fields("stat").value & "") = "1")
                
                lblAccNo.Caption = Trim(.Fields("accdt").value & "") & "-" & .Fields("accseq").value & ""
                lblOrdNo.Caption = .Fields("ordno").value & ""
                lblOrdSeq.Caption = .Fields("ordseq").value & ""
                
                lblTestCd.Caption = .Fields("ordcd").value & ""
                lblTestNm.Caption = .Fields("testnm").value & ""
                lblUnitQty.Caption = .Fields("unitqty").value & ""
                lblReqDtTm.Caption = Format(.Fields("reqdt").value & "", "####-##-##") & " " & Format(.Fields("reqtm").value & "", "##:##:##")
                lblOrdDt.Caption = Format(.Fields("orddt").value & "", "####-##-##")
                lblRsnNm.Caption = ""
                txtrmk = .Fields("mesg").value & ""
                If .Fields("irrfg").value & "" = "1" Then
                    lblirrFg.Visible = True
                    lblirrstring.Caption = objBldDelivery.GetIRROrder
                    lblirrstring.Visible = True
                Else
                    lblirrFg.Visible = False
                    lblirrstring.Visible = False
                End If
                chkIrra.value = Val(.Fields("irrfg").value & "")
                lblIrra.Caption = .Fields("irradfg").value & ""
                lblFilter.Caption = .Fields("filterfg").value & ""
                
                txtRcvId = ""
                lblRcvNm.Caption = ""
                'chkIrra.value = 0
                chkFilter.value = 0
                
                Call SetTransRsn(.Fields("ptid").value & "", .Fields("orddt").value & "", .Fields("ordno").value & "")
                
                lblNewTestDiv = .Fields("newtestdiv").value & ""
                If mode = EMode.modeDELIVERY Then
                    lstNewTest.Clear
                    Call SetNewTest(BldSrc, BldYY, BldNo, CompoCd, .Fields("volumn").value & "", .Fields("testdiv").value & "")
                Else
'                    cboNewTest.Clear
'                    cboNewTest.AddItem .Fields("newordcd") & " " & .Fields("newordnm")
'                    cboNewTest.ListIndex = 0
                    lstNewTest.Clear
                    If .Fields("newordcd").value <> "" Then
                        lstNewTest.AddItem .Fields("newordcd").value & "" & " " & .Fields("newordnm").value & ""
                        onPgm = True
                        lstNewTest.Selected(0) = True
                        onPgm = False
                    End If
                End If
                
                Select Case mode
                    Case EMode.modeRETURN:    txtRetID.SetFocus
                    Case EMode.modeDELIVERY:  txtRcvId.SetFocus
                    Case EMode.modeEXPIRE:    txtExpId.SetFocus
                    Case EMode.modeBAGRETURN: txtBagID.SetFocus
                End Select
                
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
        End If
    End With
    
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    
    Set DrRS = Nothing
    Set objBldDelivery = Nothing
End Sub

Private Sub SetExpireRsn()
    Dim i As Long
    Dim RS As Recordset
    
    Set RS = ReadCom003(BC2_EXP_RESON)
    
    With RS
        cboExpRsnCd.Clear
        For i = 1 To .RecordCount
            cboExpRsnCd.AddItem .Fields("cdval1").value & "" & " " & .Fields("field1").value & "" & vbTab & .Fields("field2").value & ""
            .MoveNext
        Next i
    End With
    
    Set RS = Nothing
End Sub

Private Sub ClearAll()
    'txtBldNo = ""
    cboCompo.Clear
    lblCompoCnt = "0"
    
    lblCompoCd.Caption = ""
    lblCompoNm.Caption = ""
    
    lblBldABO.Caption = ""
    lblVolumn.Caption = ""
    lblTestDt.Caption = ""
    lblRstSeq.Caption = ""
    lblWorkSeq.Caption = ""
    lblDeliveryDt.Caption = ""
    lblDeliverySeq.Caption = ""
    lblDeliveryTm.Caption = ""
    
    
    lblirrFg.Visible = False
    lblirrstring.Visible = False
    
    Call ICSPatientMark
    chkIrra.value = 0
End Sub

Private Sub Clear()
    
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDeptCd.Caption = ""
    lblDeptNm.Caption = ""
    lblWard.Caption = ""
    lblABO.Caption = ""
    lblInfection.Visible = False
    lblReaction.Visible = False
    
    tblDelivery.MaxRows = 0
    
'    lblBldABO.Caption = ""
'    lblVolumn.Caption = ""
'    lblTestDt.Caption = ""
'    lblRstSeq.Caption = ""
'
'    lblDeliveryDt.Caption = ""
'    lblDeliveryTm.Caption = ""
'    lblDeliverySeq.Caption = ""
    
    picStat.Visible = False
    
    lblAccNo.Caption = ""
    lblOrdNo.Caption = ""
    lblOrdSeq.Caption = ""
    lblTestCd.Caption = ""
    lblTestNm.Caption = ""
    lblUnitQty.Caption = ""
    lblReqDtTm.Caption = ""
    lblOrdDt.Caption = ""
    
    lblRsnNm.Caption = ""
    txtrmk = ""
    lblIrra.Caption = ""
    lblFilter.Caption = ""
    
    
    chkIrra.value = 0
    chkFilter.value = 0
    
'    cboNewTest.Clear
    lstNewTest.Clear
    txtRetID = ""
    txtBagID = ""
    txtExpId = ""
    lblRetNm.Caption = ""
    lblBagNm.Caption = ""
    lblExpNm.Caption = ""
    lblExpBillDiv.Caption = ""
    txtExpRsnRmk = ""
    txtRetRmk = ""
End Sub

Private Sub txtExpId_Change()
    If lblExpNm.Caption = "" Then
        lblExpNm.Caption = ""
        cboExpRsnCd.ListIndex = -1
        lblExpBillDiv.Caption = ""
        txtExpRsnRmk.Text = ""
    End If
End Sub

Private Sub txtExpId_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtExpId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtExpId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtExpId_Validate(Cancel As Boolean)
    Dim strExpNm As String
    
    If txtExpId.Text = "" Then Exit Sub
    
    strExpNm = GetEmpNm(txtExpId.Text)
    
    If strExpNm = "" Then
        Cancel = True
        MsgBox "등록되지 않은 사용자 입니다.", vbExclamation
    Else
        lblExpNm.Caption = strExpNm
    End If
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtRcvId_Change()
    If lblRcvNm.Caption <> "" Then
        lblRcvNm.Caption = ""
    End If
End Sub

Private Sub txtRcvId_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRcvId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtRcvId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRcvId_Validate(Cancel As Boolean)
    Dim strRcvNm As String
    
    If txtRcvId.Text = "" Then Exit Sub
    
    strRcvNm = GetEmpNm(txtRcvId.Text)
    If strRcvNm = "" Then
        Cancel = True
        MsgBox "등록되지 않은 사용자입니다.", vbExclamation
    Else
        lblRcvNm.Caption = strRcvNm
    End If
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtRetID_Change()
    If lblRetNm.Caption <> "" Then
        lblRetNm.Caption = ""
        txtRetRmk.Text = ""
    End If
End Sub

Private Sub txtRetID_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRetID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtRetID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRetID_Validate(Cancel As Boolean)
    Dim strRetNm As String
    
    If txtRetID.Text = "" Then Exit Sub
    
    strRetNm = GetEmpNm(txtRetID.Text)
    
    If strRetNm = "" Then
        Cancel = True
        MsgBox "등록되지 않은 사용자입니다.", vbExclamation
    Else
        lblRetNm.Caption = strRetNm
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtBagID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBagID_LostFocus()
    Dim name As String
    
    If txtBagID = "" Then
        lblBagNm.Caption = ""
        Exit Sub
    End If
    
    name = GetEmpNm(txtBagID.Text)
    lblBagNm.Caption = name
    
'    If name <> "" Then
'        Call SetHisEmpToLisEmp(txtBagID.Text, lblBagNm.Caption)
'    End If
    
    cmdSave.SetFocus
End Sub


'출고혈액 반환요청서 출력물임...
Private Sub PrintBloodReturn()
'    Dim strPrint As String
'
'    strPrint = MsgBox("반환요청서를 출력하시겠습니까?", vbInformation + vbYesNo, "반환요청서 출력")
'    If strPrint = vbNo Then Exit Sub
'
'    Dim objPrint As New clsBBSPrint
'
'
'    Call objPrint.PrintBloodReturn(lblPtNm.Caption, lblPtId.Caption, lblSexAge.Caption, lblDeptNm.Caption, lblWard.Caption, _
'                              lblABO.Caption, txtBldNo.Text, lblDeliveryDt.Caption, lblDeliveryTm.Caption, Trim(Mid(cboCompo.Text, 3)), chkBarCode.value)
'    Set objPrint = Nothing
End Sub

Private Sub cmdClose_Click()
    fraFilter.Visible = False
End Sub

Private Sub cmdF_Click()
    Call FilterClear
    
    fraFilter.Visible = True
    fraFilter.ZOrder 0
    txtPtId.SetFocus
    If lblPtId.Caption <> "" Then
        txtPtId.Text = lblPtId.Caption
        lblPtnm1.Caption = lblPtNm.Caption
        cmdFQuery.SetFocus
    End If
    
End Sub

Private Sub FilterClear()
    txtPtId.Text = "": lblDelCnt.Caption = ""
    lblPtnm1.Caption = ""
    tblFilter.MaxRows = 0
    dtpFromDate.value = DateAdd("d", -3, GetSystemDate)
    dtpToDate.value = GetSystemDate
End Sub

'Filter처방조회
Private Sub cmdFQuery_Click()
    Dim RS          As Recordset
    Dim strPtid     As String
    Dim strFDate    As String
    Dim strTDate    As String
    Dim SSQL        As String
    
    strPtid = txtPtId.Text
    
    If strPtid = "" Then
        MsgBox "환자ID를 입력후 조회하세요.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    strFDate = Format(dtpFromDate.value, PRESENTDATE_FORMAT)
    strTDate = Format(dtpToDate.value, PRESENTDATE_FORMAT)
    
    SSQL = " SELECT a.ordcd,a.entdt,a.workseq,b.testnm " & _
           " FROM " & T_BBS001 & " b," & T_BBS304 & " a" & _
           " WHERE " & DBW("a.entdt>=", strFDate) & _
           " AND   " & DBW("a.entdt<=", strTDate) & " AND " & DBW("a.ptid=", strPtid) & _
           " AND   " & DBW("a.stscd=", BBSBloodStatus.stsDELIVERY) & _
           " AND   (a.retfg is null or " & DBW("a.retfg<>", "1") & ")" & _
           " AND   (a.expfg is null or " & DBW("a.expfg<>", "1") & ")" & _
           " AND a.ordcd=b.testcd" & _
           " AND (b.expdt='' or b.expdt is null)"

    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    With tblFilter
        .MaxRows = 0
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .RowHeight(.Row) = 13.3
                .Col = 1: .CellType = CellTypeCheckBox
                          .TypeHAlign = TypeHAlignCenter
                .Col = 2: .value = RS.Fields("ordcd").value & ""
                .Col = 3: .value = RS.Fields("testnm").value & ""
                .Col = 4: .value = Format(RS.Fields("entdt").value & "", "0###-##-##")
                .Col = 5: .value = RS.Fields("workseq").value & ""
                RS.MoveNext
            Loop
            Call tblFilter_Click(1, 1)
        End If
    End With
    
    Set RS = Nothing
    
End Sub

Private Sub cmdFSave_Click()
    Dim RS              As Recordset
    Dim strWorkArea     As String
    Dim strAccDt        As String
    Dim strAccSeq       As String
    Dim strRCnt         As String
    Dim strECnt         As String
    
    Dim strTestCd       As String
    Dim strTestNm       As String
    Dim strEntdt        As String
    Dim strEntTm        As String
    Dim strEntID        As String
    Dim strTmp          As String
    Dim blnDelivery     As Boolean
    Dim SSQL            As String
    Dim ii              As Integer
    
    Dim strWorkSeq      As String
    Dim strMode         As String
    
    
    strEntdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strEntTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strEntID = ObjSysInfo.EmpId
    
    Select Case mode
        Case EMode.modeEXPIRE:  strMode = BBSBloodStatus.stsEXPIRE
        Case EMode.modeRETURN:  strMode = BBSBloodStatus.stsRETURN
    End Select
    
On Error GoTo Errors
    DBConn.BeginTrans
    
    With tblFilter
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1
            If .CellType = CellTypeCheckBox And .value = 1 Then
                .Col = 2: strTestCd = .value
                .Col = 3: strTestNm = .value
                .Col = 5: strWorkSeq = .value
                
                strTmp = MsgBox("검사코드 : " & strTestCd & "[" & strTestNm & "]" & vbCrLf & _
                                " 의 필터를 반환(또는 폐기) 하시겠습니까?", vbYesNo + vbInformation, "Info")
                If strTmp = vbYes Then
                    Select Case strMode
                        Case BBSBloodStatus.stsEXPIRE
                        '전주 예수병원은 수가관련 테이블은 건들필요 없음
'                            SSQL = " UPDATE " & T_BBS903 & " SET " & DBW("orddiv", "4", 3) & _
'                                        "entdt=sysdate," & _
'                                         DBW("donefg", "0", 3) & DBW("mesg", "", 2) & _
'                                   " WHERE order_key=" & strWorkSeq
'                            DBConn.Execute SSQL
                            
                            SSQL = " UPDATE " & T_BBS304 & " SET " & _
                                     DBW("expfg", "1", 3) & DBW("expdt", strEntdt, 3) & DBW("exptm", strEntTm, 3) & DBW("expid", strEntID, 2) & _
                                   " WHERE " & DBW("WORKSEQ", strWorkSeq, 2)
                            DBConn.Execute SSQL
                                 
                        Case BBSBloodStatus.stsRETURN
                        '전주 예수병원은 수가관련 테이블은 건들필요 없음
'                            SSQL = " UPDATE " & T_BBS903 & " SET " & DBW("orddiv", "1", 3) & _
'                                        "entdt=sysdate," & _
'                                         DBW("donefg", "0", 3) & DBW("mesg", "", 2) & _
'                                   " WHERE order_key=" & strWorkSeq
'                            DBConn.Execute SSQL
                            
                            SSQL = " UPDATE " & T_BBS304 & " SET " & _
                                     DBW("retfg", "1", 3) & DBW("retdt", strEntdt, 3) & DBW("rettm", strEntTm, 3) & DBW("retid", strEntID, 2) & _
                                   " WHERE " & DBW("WORKSEQ=", strWorkSeq)
                            DBConn.Execute SSQL
                    End Select
                    
                    SSQL = " select a.workarea,a.accdt,a.accseq,b.retcnt,b.expcnt " & _
                           " from " & T_BBS203 & " b," & T_LAB102 & " a," & T_BBS304 & " c" & _
                           " where " & DBW("c.workseq=", strWorkSeq) & _
                           " and  c.ptid = a.ptid And c.orddt = a.orddt And c.ordno = a.ordno And c.Ordseq = a.Ordseq" & _
                           " and a.workarea=b.workarea and a.accdt=b.accdt and a.accseq=b.accseq"
                    Set RS = Nothing
                    Set RS = New Recordset
                    RS.Open SSQL, DBConn
                    
                    If Not RS.EOF Then
                        strWorkArea = RS.Fields("workarea").value & ""
                        strAccDt = RS.Fields("accdt").value & ""
                        strAccSeq = RS.Fields("accseq").value & ""
                        strRCnt = "0": strECnt = "0"
                        If strMode = BBSBloodStatus.stsRETURN Then
                            strRCnt = Val(RS.Fields("retcnt").value & "") + 1
                        End If
                        If strMode = BBSBloodStatus.stsRETURN Then
                            strECnt = Val(RS.Fields("expcnt").value & "") + 1
                        End If
                        SSQL = " update " & T_BBS203 & " set " & DBW("retcnt", strRCnt, 3) & DBW("expcnt", strECnt, 2) & _
                             " where " & DBW("workarea=", strWorkArea) & _
                             " and " & DBW("accdt=", strAccDt) & _
                             " and " & DBW("accseq=", strAccSeq)
                        DBConn.Execute SSQL
                    End If
                    
                    
                    blnDelivery = True
                End If
            End If
        Next
    End With
    
    If blnDelivery = True Then MsgBox "정상적으로 반환/폐기 되었습니다.", vbInformation + vbOKOnly, "Info"
    
    Call FilterClear
Skip:
    DBConn.CommitTrans
    Exit Sub
Errors:
    DBConn.RollbackTrans
End Sub

Private Function ConvToDate(ByVal argDate As String) As String
    ConvToDate = "To_Date('" & argDate & "', 'YYYYMMDD') "
End Function

Private Sub tblFilter_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    lblDelCnt.Caption = "1"
    With tblFilter
        .Row = Row: .Col = Col
        If .value = "" Then Exit Sub
    End With
End Sub

Private Function Search_PtInfo() As Boolean
    Dim objPt   As clsPtInformation
    Dim RS      As Recordset
    Dim ii      As Long
    Dim strLng  As String
    
    
    tblFilter.MaxRows = 0
    lblDelCnt.Caption = ""
    
    If txtPtId.Text = "" Then
        lblPtnm1.Caption = ""
        Search_PtInfo = True
    Else
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii

        If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
            txtPtId.Text = Format(txtPtId.Text, strLng & "#")
        End If

        Set objPt = New clsPtInformation
        Set RS = New Recordset
        RS.Open objPt.Get_Ptid(txtPtId.Text), DBConn
        
        If RS.EOF = False Then
            With objPt
                .BedPt_Chk txtPtId.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
                If .PtDiv = "BED" Then
                    lblPtnm1.Caption = .ptnm
                Else
                    lblPtnm1.Caption = .ptnm
                End If
            End With
            Search_PtInfo = True
        Else
            MsgBox "해당되는 환자가 없습니다. 확인후 조회하세요.", vbInformation + vbOKOnly, Me.Caption
            txtPtId.Text = ""
            lblPtnm1.Caption = ""
            Search_PtInfo = False
        End If
        Set RS = Nothing
        Set objPt = Nothing
    End If
    If Search_PtInfo Then cmdFQuery.SetFocus
End Function

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtPtId_LostFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
    Call Search_PtInfo
End Sub
