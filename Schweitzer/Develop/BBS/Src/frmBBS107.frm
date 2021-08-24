VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS107 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Out Patient Collection Or Accept"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   Icon            =   "frmBBS107.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   14625
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   3675
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   10785
      _ExtentX        =   19024
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
      Caption         =   "환자 기본 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   3570
      _ExtentX        =   6297
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
      Caption         =   "환자검색"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   10
      Left            =   3675
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2385
      Width           =   10785
      _ExtentX        =   19024
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
      Caption         =   "처방 정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8175
      Left            =   75
      TabIndex        =   6
      Top             =   270
      Width           =   3585
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Search"
         Height          =   630
         Left            =   60
         TabIndex        =   7
         Tag             =   "136"
         Top             =   540
         Width           =   3450
         Begin VB.TextBox txtSearchKey 
            Height          =   300
            Left            =   90
            MaxLength       =   10
            TabIndex        =   10
            Top             =   240
            Width           =   1470
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Name"
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
            Index           =   1
            Left            =   2205
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "15305"
            Top             =   270
            Width           =   810
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1680
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "15304"
            Top             =   285
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6855
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "환자ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "환자명"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "주민번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "생년월일"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "성별"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "접수일"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "접수번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "처방일자"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   330
         Left            =   1020
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   165
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
         Format          =   76873728
         UpDown          =   -1  'True
         CurrentDate     =   36342.5951388889
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   165
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
         Caption         =   "처방일"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   2085
      Left            =   3675
      TabIndex        =   14
      Top             =   300
      Width           =   10800
      Begin VB.TextBox txtReceptNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1185
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   810
         Width           =   2010
      End
      Begin VB.TextBox txtPtid 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1185
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   390
         Width           =   2010
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  '평면
         BackColor       =   &H00FEEFFE&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   1185
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1230
         Width           =   8910
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   4455
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   390
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   360
         Left            =   8010
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   4455
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   810
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   8010
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   810
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   390
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   3375
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   390
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "성명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   3375
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   6930
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   390
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   6930
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   105
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "영수증번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   105
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "Remark"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   4395
      Left            =   3675
      TabIndex        =   29
      Top             =   2625
      Width           =   10800
      Begin VB.CheckBox chkCollect 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   330
         Width           =   795
      End
      Begin VB.CheckBox chkCollect 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈및 접수"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   330
         Width           =   1275
      End
      Begin VB.CheckBox chkCollect 
         BackColor       =   &H00DBE6E6&
         Caption         =   "추가채혈"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   330
         Width           =   1095
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   3135
         Left            =   180
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   690
         Width           =   10425
         _Version        =   196608
         _ExtentX        =   18389
         _ExtentY        =   5530
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   10
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS107.frx":076A
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblReaction 
         Height          =   315
         Left            =   9420
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
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
         Left            =   9000
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   1305
      Left            =   3675
      TabIndex        =   36
      Top             =   7140
      Width           =   10800
      Begin VB.TextBox txtColNo 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   4740
         TabIndex        =   42
         Top             =   690
         Width           =   675
      End
      Begin VB.TextBox txtRowNo 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   4080
         TabIndex        =   41
         Top             =   690
         Width           =   675
      End
      Begin VB.CheckBox chkSPos 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체보관장소 자동지정"
         Height          =   555
         Left            =   30
         TabIndex        =   40
         Top             =   195
         Width           =   1395
      End
      Begin VB.TextBox txtColID 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   7230
         TabIndex        =   39
         Top             =   195
         Width           =   1035
      End
      Begin VB.CommandButton cmdPopUp 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   360
         Left            =   8310
         Style           =   1  '그래픽
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   195
         Width           =   360
      End
      Begin VB.ComboBox cboLeg 
         Height          =   300
         ItemData        =   "frmBBS107.frx":0E69
         Left            =   3060
         List            =   "frmBBS107.frx":0E6B
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   690
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   375
         Left            =   7230
         TabIndex        =   43
         Top             =   615
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd "
         Format          =   76873731
         CurrentDate     =   36851
      End
      Begin DRcontrol1.DrLabel lblColNm 
         Height          =   360
         Left            =   8730
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   195
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
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
         Border          =   1
         TextPosition    =   4
         Caption         =   ""
      End
      Begin MSComCtl2.DTPicker dtpColTm 
         Height          =   360
         Left            =   8730
         TabIndex        =   45
         Top             =   630
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   76873731
         UpDown          =   -1  'True
         CurrentDate     =   36851
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   6150
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   6150
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   615
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   780
         Index           =   12
         Left            =   1785
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   195
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1376
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
         Caption         =   "검체보관장소"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   480
         Index           =   13
         Left            =   3060
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   195
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   847
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
         Caption         =   "Rack"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   480
         Index           =   14
         Left            =   4095
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   195
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   847
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
         Caption         =   "Row"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   480
         Index           =   15
         Left            =   4755
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   195
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   847
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
         Caption         =   "Col"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents objLisCollectForm As clsLisCollectForm
Attribute objLisCollectForm.VB_VarHelpID = -1

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Me.Show
    Me.WindowState = 2
    Set objLisCollectForm = New clsLisCollectForm
    objLisCollectForm.EmpId = ObjSysInfo.EmpId
    Call objLisCollectForm.CollectionButtonClick("LIS206", Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objLisCollectForm = Nothing
End Sub

Private Sub objLisCollectForm_LastFormUnload()
    Unload Me
End Sub

'Option Explicit
'Private Enum TblColumn
'    tcORDDT = 1
'    tcORDNM
'    tcTRANSDT
'    tcREASON
'    tcDOCT
'    tcERCHECK
'    tcTESTLOCATION
'    tcREMARK
'    tcORDNO
'    tcORDSEQ
'End Enum
'
'
'Private WithEvents objMyList As clspopuplist
'Private WithEvents objListPop As clspopuplist
'
'Public OnlyCollection As Boolean    '채혈만 하는지?
'Private blnSearch As Boolean
'Private strDeptCd As String     '진료과구분
'Private strBldCd As String      '병동의 건물 코드
'Private strErbldcd As String    '응급일경우 검사할 건물코드
'Private strGbldcd As String     '일반일경우 검사할 건물코드
'Private strWardID As String
'Private strBussdiv As String
'Private strB_SpcNum As String   '바코드발행할 검체번호
'Private bln_BarCd As Boolean
'Private Bussdiv As String
'Private blnAdd_Col As Boolean
'Private strReqDt As String
'Private strStatFg As String
'
'Private Sub chkSPos_Click()
'    If chkSPos.value = 1 Then
'        txtRowNo = ""
'        txtColNo = ""
'        txtRowNo.Locked = True
'        txtColNo.Locked = True
'        txtRowNo.BackColor = Me.BackColor
'        txtColNo.BackColor = Me.BackColor
'    Else
'        txtRowNo.Locked = False
'        txtColNo.Locked = False
'        txtRowNo.BackColor = RGB(255, 255, 255)
'        txtColNo.BackColor = RGB(255, 255, 255)
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub chkSPos_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If chkSPos.value = 1 Then
'            chkSPos.value = 0
'        Else
'            chkSPos.value = 1
'            txtColID.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
'End Sub
'
'Private Sub Form_Load()
'    If OnlyCollection = True Then Me.Caption = "외래 채혈"
'
'    Dim objAccess As clsBBSAccess
'    Dim RS        As RECORDSET
'    Set objAccess = New clsBBSAccess
'
'    With objAccess
'        Set RS = OpenRecordSet(.Get_LegPos(ObjSysInfo.BuildingCd))
'        If RS.EOF = False Then
'            cboLeg.Clear
'            Do Until RS.EOF = True
'                cboLeg.AddItem RS.Fields("legcd").value & ""
'                RS.MoveNext
'            Loop
'        End If
'        If cboLeg.ListCount > 0 Then cboLeg.ListIndex = 0
'    End With
'    Set RS = Nothing
'    Set objAccess = Nothing
'
'    lvwPtList.ListItems.Clear
'
'    dtpToTime.value = Format(GetSystemDate, "yyyy-MM-dd")
'    dtpColdt.value = Format(GetSystemDate, "yyyy-MM-dd")
'    dtpColTm.value = Format(GetSystemDate, "HH:mm")
'    optSort(0).value = True
'    blnSearch = True
'    chkSPos.value = 1
'End Sub
'
'Private Sub cmdClear_Click()
'    Clear
'    lvwPtList.ListItems.Clear
'End Sub
'Private Sub Clear()
'    Call ICSPatientMark
'    tblResult.MaxRows = 0
'    txtPtid = ""
'    lblPtNm.Caption = ""
'    lblSexAge.Caption = ""
'    lblDeptNm.Caption = ""
'    lblOrdDt.Caption = ""
'    txtReceptNo.Text = ""
'    txtColNo.Text = ""
'    txtRowNo.Text = ""
'    txtColID.Text = ""
'    txtSearchKey.Text = ""
'    lblColNm.Caption = ""
'    lblInfection.Visible = False
'    lblReaction.Visible = False
'End Sub
'
'Private Sub cmdPopUp_Click()
'    Set objMyList = New clspopuplist
'    With objListPop
'        .BackColor = Me.BackColor
'        .Caption = "직원조회": .HeadName = "사번, 직원명"
'        .Width = .Width + 300: .ColSize(0) = 1000
'        Call .ListPop(getsqlemp, 2350, 7650)
'        If .SelectedString <> "" Then
'
'            txtColID.Text = medGetP(.SelectedString, 1, ";")
'            lblColNm.Caption = medGetP(.SelectedString, 2, ";")
'        End If
'    End With
'    Set objMyList = Nothing
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Call ICSPatientMark
'End Sub
'
'Private Sub optSort_Click(Index As Integer)
'    If Index = 0 Then
'        blnSearch = True
'    Else
'        blnSearch = False
'    End If
'End Sub
'Private Sub SearchColID()
'    lblColNm.Caption = getempnm(txtColID.Text)
'    If lblColNm.Caption = "" Then
'        MsgBox "해당되는 자료가 없습니다.확인후 입력하세요.", vbInformation + vbOKOnly, Me.Caption
'        txtColID.Text = ""
'        lblColNm.Caption = ""
'    End If
'End Sub
'
'Private Sub txtColID_GotFocus()
'    txtColID.SelStart = 0
'    txtColID.SelLength = Len(txtColID)
'    txtColID.tag = txtColID
'End Sub
'
'Private Sub txtColID_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtColID_LostFocus()
'    If Trim(txtColID) = "" Then Exit Sub
'    If txtColID.tag = txtColID Then Exit Sub
'
'    Call SearchColID
'End Sub
'Private Function Save_chk() As Boolean
'    If tblResult.MaxRows > 0 Then
'        Save_chk = True
'    Else
'        Save_chk = False
'    End If
'End Function
'Private Sub DetailSearch(ptid As String)
''부작용,감염정보를 조회한다.
'    Dim objinfection As New clsInfection
'    Dim objReaction As New clsReaction
'
'    With objinfection
'        .ptid = ptid
'        .GetInfection
'        If .Infection = True Then
'            lblInfection.Visible = True
'        Else
'            lblInfection.Visible = False
'        End If
'    End With
'
'    With objReaction
'        .ptid = ptid
'        .GetReaction
'        If .Reaction = True Then
'            lblReaction.Visible = True
'        Else
'            lblReaction.Visible = False
'        End If
'    End With
'
'
'    Set objReaction = Nothing
'    Set objinfection = Nothing
'End Sub
'
'Private Sub cmdSave_Click() '채혈/접수 수행
'
'    Dim objOutPatient As clsBBSCollection
'    Dim objOut        As clsDictionary
'    Dim buildcd       As String
'    Dim colid         As String
'    Dim coldt         As String
'    Dim ColTm         As String
'    Dim rcvid         As String
'    Dim rcvdt         As String
'    Dim rcvTm         As String
'    Dim strErChk      As String
'    Dim SavePos       As String
'    Dim i             As Integer
'
'
'    If txtPtid = "" Then Exit Sub
'    Set objOutPatient = New clsBBSCollection
'    Set objOut = New clsDictionary
'
'    If chkSPos.value = 0 Then
'        If Trim(cboLeg.Text) = "" Or Trim(txtRowNo) = "" Or Trim(txtColNo) = "" Then
'            MsgBox "검체보관장소를 지정하십시오.", vbCritical, Me.Caption
'            Exit Sub
'        End If
'    End If
'
'    With tblResult
'        For i = 1 To .MaxRows
'            .Row = i: .Col = 6
'            If .value = "응급" Then
'                strErChk = "Ok"
'                Exit For
'            End If
'        Next
'    End With
'
'    If chkSPos.value = 1 Then
'        SavePos = "Ok"
'    End If
'
'
'    buildcd = ObjSysInfo.BuildingCd
'    colid = txtColID.Text
'    rcvid = ObjSysInfo.EmpId
'    rcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
'    rcvTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
'    coldt = Format(dtpColdt.value, PRESENTDATE_FORMAT)
'    ColTm = Format(dtpColTm.value, PRESENTTIME_FORMAT)
'
'    objOut.Clear
'    objOut.FieldInialize "ptid", "buildcd,erchk,colid,coldt,coltm,rcvid,rcvdt,rcvtm,savechk,leg,row,col"
'    objOut.AddNew txtPtid.Text, Join(Array(buildcd, strErChk, colid, coldt, ColTm, rcvid, rcvdt, rcvTm, SavePos, cboLeg.Text, txtRowNo, txtColNo), COL_DIV)
'
'    If objOutPatient.Set_OutPatientCollect(objOut) Then
'        If objOutPatient.SpcNum <> "" Then
'            BarCode_Print objOutPatient.SpcNum
'        End If
'        lvwPtList.ListItems.Remove (lvwPtList.SelectedItem.Index)
'        Clear
'    End If
'    Set objOut = Nothing
'    Set objOutPatient = Nothing
'
'End Sub
'Private Sub BarCode_Print(ByVal SpcNumber As String)
'    Dim objBar     As clsBarcode
'    Dim strBuildNm As String        '건물이름
'    Dim strW_Dept  As String
'    Dim strColDt   As String
'    Dim strColTm   As String
'    Dim strAccSeq  As String         'SpcYy-SpcNo 형태의 검체번호
'
'    Set objBar = New clsBarcode
'    Set objBar.MyDb = DBConn
'    Set objBar.TableInfo = New clsTables
'
'    strW_Dept = strWardID
'    If strW_Dept = "" Then strW_Dept = strDeptCd
'
'    strColDt = Format(Mid(Format(GetSystemDate, PRESENTDATE_FORMAT), 5), "00/00")
'    strColTm = Format(dtpColTm.value, "HH:mm")
'
'    '검체번호 출력 : 2001.2.8 추가
'    strAccSeq = Mid(SpcNumber, 1, 2) & "-" & Format(Mid(SpcNumber, 3), "########0")
'    strAccSeq = Format(strAccSeq, String(11, "@"))
'    strBuildNm = BBSName
'
'    objBar.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strB_SpcNum, txtPtid.Text, _
'                                        lblPtNm.Caption, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
'                                        "", 1
'    Set objBar = Nothing
'End Sub
'
'Private Sub txtColNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
'End Sub
'
'
'Private Sub lvwPtList_DblClick()
'    Dim iTmx As ListItem
'    Dim strPtid As String
'    Dim strAccdt As String
'    Dim strAccSeq As String
'    Dim tmpOrdDt As String
'
'    If lvwPtList.ListItems.Count < 1 Then Exit Sub
'    With lvwPtList
'        Set iTmx = .ListItems(.SelectedItem.Index)
'        strPtid = .ListItems(.SelectedItem.Index).Text
'        strAccdt = iTmx.SubItems(3)
'        strAccSeq = iTmx.SubItems(4)
'        tmpOrdDt = iTmx.SubItems(5)
'    End With
'
'    Call ICSPatientMark(strPtid, enICSNum.BBS_ALL)
'
'    If strAccdt = "" Or strAccdt = "*" Then         '처방에 따른 정상적인 채혈
'        blnAdd_Col = True                           '(*)는 정상 채혈과 추가채혈이 동시존재한다.
'        If strAccdt = "*" Then
'            ptInfo strPtid, tmpOrdDt
'        Else
'        ptInfo strPtid, Format(dtpToTime.value, PRESENTDATE_FORMAT)
'        End If
'        PtDisplay strAccdt
'    Else                                            '검체추가에따를 채혈
'        blnAdd_Col = False
'        ptInfo strPtid, , strAccdt, strAccSeq
'        PtDisplay strAccdt, strAccSeq
'    End If
'    chkSPos.SetFocus
'End Sub
'Private Function Direct_Collect(searchkey As String, TF As Boolean) As Boolean
''채혈대상자를 조회시에 사용함.
''조회하고자 하는 문자를 입력한후 Enter(신규검체조회,추가검체 두가지를 구분하여 보여준다.
''처음 채혈하고자하는 채혈대상과 검체추가에 의한 채혈의 구분은
''리스트뷰 item 3,4,5 에 접수일자/접수번호를 가지고 구분한다.
'    Dim objGetSql As New clsBBSCollection
'    Dim DrRS As New RECORDSET
'    Dim strOrdDt As String
'    Dim blnEOF As Boolean
'    Dim blnEOF1 As Boolean
'    Dim iTmx As Object
'    Dim itmx2 As Object
'    Dim itmFound As ListItem
'
''    objGetSql.setDbConn DBConn
'
'    strOrdDt = Format(dtpToTime.value, PRESENTDATE_FORMAT)
'
'    lvwPtList.ListItems.Clear
'
'    Set DrRS = OpenRecordSet(objGetSql.Get_SideCollect(searchkey, TF, strOrdDt))
'    If DrRS.EOF = False Then
'        With lvwPtList
'            .ListItems.Clear
'            Do Until DrRS.EOF
'                Set itmx2 = .ListItems.Add(, , DrRS.Fields("ptid").value)
'                itmx2.SubItems(1) = DrRS.Fields("ptnm").value
'                itmx2.SubItems(2) = Mid(DrRS.Fields("SSN").value, 1, 6) & "-" & Mid(DrRS.Fields("ssn").value, 7)
'                DrRS.MoveNext
'            Loop
'        End With
'        blnEOF = True
'    End If
'    DrRS.RsClose:        Set DrRS = Nothing
'
'    '추가검체 조회
'    Set DrRS = OpenRecordSet(objGetSql.Get_AddSpcInFo(searchkey, TF))
'    If DrRS.EOF = False Then
'        With lvwPtList
'            Do Until DrRS.EOF
'                Set itmFound = .FindItem(DrRS.Fields("ptid").value, lvwText, , lvwPartial)
'                If itmFound Is Nothing Then
'                    Set iTmx = .ListItems.Add(, , DrRS.Fields("ptid").value)
'                    iTmx.ForeColor = vbBlue
'                    iTmx.SubItems(1) = DrRS.Fields("ptnm").value
'                    iTmx.ListSubItems(1).ForeColor = vbBlue
'                    iTmx.SubItems(2) = Mid(DrRS.Fields("SSN").value, 1, 6) & "-" & Mid(DrRS.Fields("ssn").value, 7)
'                    iTmx.ListSubItems(2).ForeColor = vbBlue
'                    iTmx.SubItems(3) = DrRS.Fields("accdt").value
'                    iTmx.SubItems(4) = DrRS.Fields("accno").value
'                    iTmx.SubItems(5) = DrRS.Fields("orddt").value
'                    'lblOrdDt.Caption = Format(DrRS.Fields("orddt"), "####-##-##")
'                Else
'                    '정상적인 채혈과 검체추가가 겹치는 경우
'                    .ListItems(itmFound.Index).SubItems(3) = "*"
'                    .ListItems(itmFound.Index).ForeColor = vbBlue
'                    .ListItems(itmFound.Index).ListSubItems(1).ForeColor = vbBlue
'                    .ListItems(itmFound.Index).ListSubItems(2).ForeColor = vbBlue
'
'                End If
'                DrRS.MoveNext
'            Loop
'        End With
'        blnEOF1 = True
'    End If
'
'    DrRS.RsClose
'    Set DrRS = Nothing
'
'    If blnEOF = False And blnEOF1 = False Then
'        Direct_Collect = False
'        Clear
'    Else
'        Direct_Collect = True
'    End If
'
'    Set objGetSql = Nothing
'End Function
'Private Sub ptInfo(ByVal ptid As String, Optional orddt As String = "", _
'                   Optional accdt As String = "", Optional accseq As String = "")
''리스트뷰에서 선택한 환자의 환자정보와 채혈내역에 저장될 건물코드를 조회한다.
'    Dim objGetSql As New clsGetSqlStatement
'    Dim objCollect As New clsBBSCollection
'    Dim DrRS      As New RECORDSET
'    Dim strTmp    As String
'    Dim strSDA    As String             'Sex/Birth/Age
'
'    Set DrRS = objGetSql.Get_PtInfo(ptid, BBSBUSSDIV.stsNotBed, orddt, accdt, accseq)
'    txtPtid = ptid
'    lblPtNm.Caption = DrRS.Fields("ptnm").value & ""
'
'    strSDA = SDA_String(DrRS.Fields("ssn").value & "")
'    lblSexAge.Caption = medGetP(strSDA, 1, COL_DIV) & "/" & medGetP(strSDA, 3, COL_DIV)
'
'    lblDeptNm.Caption = IIf(IsNull(DrRS.Fields("deptnm").value & "") = True, "", DrRS.Fields("deptnm").value & "")
'    strDeptCd = DrRS.Fields("deptcd").value & ""
'
'    strBldCd = ObjSysInfo.BuildingCd
'    strTmp = objCollect.TestBuildCd(strBldCd)
'    strErbldcd = medGetP(strTmp, 1, COL_DIV)
'    strGbldcd = medGetP(strTmp, 2, COL_DIV)
'
'    Set objGetSql = Nothing
'    Set objCollect = Nothing
'End Sub
'Private Sub PtDisplay(Optional ByVal accdt As String = "", Optional ByVal accseq As String = "")
'    '조회된 환자ID를 가지고 채혈등록시 필요한 자료를 가지고 온다.
'    Dim RS          As New RECORDSET
'    Dim iTmx        As ListItem
'    Dim strOrdDt    As String
'    Dim strReason   As String
'    Dim strTmp      As String
'    Dim blnStatFg   As Boolean
'    Dim i           As Integer
'
'    Dim objGetSql       As clsBBSCollection
'    Dim objTransReason  As clsQueryOrder
'
'    Set objGetSql = New clsBBSCollection
'    Set objTransReason = New clsQueryOrder
'
'    strOrdDt = Format(dtpToTime.value, PRESENTDATE_FORMAT)
'    i = 1
'
'    If accdt = "" Or accdt = "*" Then
'        '일반처방에 의한 채혈이거나, 추가요청과,일반처방이 같이 존재하는 경우
'        '처방을 저장하고, 추가채혈내역,번호부여정보,(일반처방과,추가채혈이 동시)
'        '처방저장,번호부여정보저장(일반처방)
'        Set RS = objGetSql.Get_Order_107(txtPtid.Text, strOrdDt, "107")
'    Else
'        '추가처방에 의한 채혈인경우(추가채혈내역,번호부여정보저장)
'        Set RS = objGetSql.Get_ADDSPC(txtPtid.Text, accdt, accseq)
'        If Not RS.EOF Then strReqDt = RS.Fields("reqdt1").value & ""
'    End If
'
'    '감염정보,부작용정보를 조회한다.
'    Call DetailSearch(txtPtid.Text)
'    '''
'    If Not RS.EOF = True Then
'        Select Case RS.Fields("donefg").value
'            Case "0": chkCollect(1).value = 1: chkCollect(0).value = 0: chkCollect(2).value = 0           '채혈및 접수 수행
'            Case "1": chkCollect(0).value = 1: chkCollect(1).value = 0: chkCollect(2).value = 0           '접수만 수행
'            Case "2": chkCollect(2).value = 1: chkCollect(0).value = 0: chkCollect(1).value = 0           '추가채혈(접수까지 수행)
'        End Select
'        If accdt = "*" Then chkCollect(2).value = 1
'
'        lblOrdDt.Caption = Format(strOrdDt, "####-##-##")     '처방일
'        strBussdiv = RS.Fields("bussdiv").value & ""                  '구분(외래:"1")
'        txtReceptNo.Text = IIf(IsNull(RS.Fields("receptno").value & "") = True, "", RS.Fields("receptno").value & "")
'
'        strBldCd = ObjSysInfo.BuildingCd                      '건물코드
'        strTmp = objGetSql.TestBuildCd(strBldCd)
'        strErbldcd = medGetP(strTmp, 1, COL_DIV)                             '응급검사 건물코드
'        strGbldcd = medGetP(strTmp, 2, COL_DIV)                              '일반검사 건물코드
'    End If
'
'    With tblResult
'        .ReDraw = False
'        Do Until RS.EOF = True
'            .MaxRows = RS.RecordCount
'            .Row = i
'            '추가요청내역의 요청일자.
'
'            .Col = TblColumn.tcORDDT: .value = Format(RS.Fields("orddt").value, "####-##-##")
'            .Col = TblColumn.tcORDNM: .value = Get_TestNm(RS.Fields("ordcd").value & "")
'            .Col = TblColumn.tcTRANSDT: .value = Format(RS.Fields("reqdt").value, "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value, 1, 4), "00:00")           '수혈예정일
'
'            strReason = objTransReason.GetTransReason(txtPtid, RS.Fields("orddt").value, RS.Fields("ordno").value)
'
'            .Col = TblColumn.tcREASON: .value = strReason                                                     '수혈사유
'            .Col = TblColumn.tcDOCT: .value = getdoctnm(RS.Fields("orddoct").value)
'            .Col = TblColumn.tcERCHECK: .value = Trim(IIf(RS.Fields("statfg").value = "1", "Y", ""))
'            .ForeColor = RGB(255, 0, 0)
'
'            Select Case RS.Fields("statfg").value
'                Case "1": .Col = TblColumn.tcTESTLOCATION: .value = objGetSql.TestBldNm(strErbldcd)
'                Case "0": .Col = TblColumn.tcTESTLOCATION: .value = objGetSql.TestBldNm(strGbldcd)
'            End Select
'
'            .Col = TblColumn.tcREMARK: .value = IIf(IsNull(RS.Fields("mesg").value) = True, "", RS.Fields("mesg").value)
'            .Col = TblColumn.tcORDNO: .value = Trim(RS.Fields("ordno").value)
'            .Col = TblColumn.tcORDSEQ: .value = Trim(RS.Fields("ordseq").value)
'
'            i = i + 1
'            RS.MoveNext
'        Loop
'        For i = 1 To .MaxRows
'            .Col = TblColumn.tcREMARK
'            If .value <> "" Then
'                txtRemark = txtRemark & .value & vbNewLine
'            End If
'            If blnStatFg = False Then
'                .Col = TblColumn.tcERCHECK
'                If .value = "Y" Then
'                    strStatFg = "1"
'                    blnStatFg = True
'                End If
'            End If
'        Next
'        If txtRemark <> "" Then
'            txtRemark = Mid(txtRemark, 1, Len(txtRemark) - 1)
'        End If
'        .ReDraw = True
'    End With
'
'    RS.RsClose
'    Set RS = Nothing
'    Set objGetSql = Nothing
'    Set objTransReason = Nothing
'End Sub
'Private Sub txtPtId_GotFocus()
'    txtPtid.tag = txtPtid
'    txtPtid.SelStart = 0
'    txtPtid.SelLength = Len(txtPtid)
'
'End Sub
'
'Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
'End Sub
'
'Private Sub txtPtId_LostFocus()
'    Dim ii        As Long
'    Dim strLng    As String
'
'    If txtPtid.tag = txtPtid Then Exit Sub
'    If txtPtid = "" Then Exit Sub
'
'
'    For ii = 1 To Val(BBS_PTID_LENGTH) - 1
'        strLng = strLng & "0"
'    Next ii
'
'    If Len(Trim(txtPtid.Text)) <> BBS_PTID_LENGTH Then
'        txtPtid.Text = Format(txtPtid.Text, strLng & "#")
'    End If
'
'    If Direct_Collect(txtPtid.Text, True) = True Then
'        Call lvwPtList_DblClick
'        chkSPos.SetFocus
'    Else
'        MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요.", vbInformation + vbOKOnly, "채혈접수대상선택"
'        Call Clear
'    End If
'End Sub
'
'Private Sub txtRowNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
'End Sub
'
'Private Sub txtSearchKey_GotFocus()
'    txtSearchKey.SelStart = 0
'    txtSearchKey.SelLength = Len(txtSearchKey)
'    txtSearchKey.tag = txtSearchKey
'End Sub
'
'Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
'End Sub
'Private Sub txtSearchKey_LostFocus()
'    Dim ii        As Long
'    Dim strLng    As String
'
'    If txtSearchKey.tag = txtSearchKey Then Exit Sub
'    If txtSearchKey = "" Then Exit Sub
'
'    If blnSearch = True Then
'        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
'            strLng = strLng & "0"
'        Next ii
'
'        If Len(Trim(txtSearchKey.Text)) <> BBS_PTID_LENGTH Then
'            txtSearchKey.Text = Format(txtSearchKey.Text, strLng & "#")
'        End If
'
'        If Direct_Collect(txtSearchKey.Text, blnSearch) = False Then
'            MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요.", vbInformation + vbOKOnly, "채혈접수대상선택"
'        Else
'            SendKeys "{TAB}"
'        End If
'    Else
'        If Direct_Collect(txtSearchKey, blnSearch) = False Then
'            MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요.", vbInformation + vbOKOnly, "채혈접수대상선택"
'        Else
'            SendKeys "{TAB}"
'        End If
'    End If
'
'    Call Clear
'    txtSearchKey = ""
'End Sub
