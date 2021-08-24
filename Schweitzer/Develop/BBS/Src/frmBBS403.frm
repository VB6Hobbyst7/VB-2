VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS403 
   BackColor       =   &H00DBE6E6&
   Caption         =   "문진등록"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS403.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   14700
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdCallTest 
      BackColor       =   &H00F4F0F2&
      Caption         =   "검사의뢰(&N)"
      Height          =   510
      Left            =   2265
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "문진취소"
      Height          =   510
      Left            =   6960
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2000-01-01"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDonorCd 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2280
      TabIndex        =   11
      Top             =   2310
      Width           =   9945
      Begin VB.ComboBox cboDonorCd 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmBBS403.frx":076A
         Left            =   1065
         List            =   "frmBBS403.frx":077A
         Locked          =   -1  'True
         Style           =   1  '단순 콤보
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   225
         Width           =   2055
      End
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00CFDCDE&
         Height          =   315
         Left            =   4575
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   225
         Width           =   1305
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   315
         Left            =   5895
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   225
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   10
         Left            =   75
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   225
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
         Caption         =   "헌혈종류"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   11
         Left            =   3570
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   225
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   8280
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9585
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
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
      BorderStyle     =   0
      Caption         =   "  기 본 정 보"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblAsk 
      Height          =   2730
      Left            =   2280
      TabIndex        =   8
      Top             =   3525
      Width           =   9915
      _Version        =   196608
      _ExtentX        =   17489
      _ExtentY        =   4815
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
      MaxCols         =   6
      MaxRows         =   50
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS403.frx":07A8
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   1695
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
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
      BorderStyle     =   0
      Caption         =   "  문 진 내 역"
      Appearance      =   0
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   2280
      TabIndex        =   22
      Top             =   6270
      Width           =   9930
      Begin VB.OptionButton optOk 
         BackColor       =   &H00DBF2FD&
         Caption         =   "적   격"
         Height          =   375
         Index           =   0
         Left            =   1380
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optOk 
         BackColor       =   &H00DBF2FD&
         Caption         =   "부적격"
         Height          =   375
         Index           =   1
         Left            =   1380
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtRmk 
         Height          =   825
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   240
         Width           =   6750
      End
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00DBE6E6&
         Caption         =   "보류"
         Height          =   255
         Left            =   540
         TabIndex        =   23
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 결과판정"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   2280
      TabIndex        =   27
      Top             =   2910
      Width           =   9930
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1080
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      End
      Begin MedControls1.LisLabel lblStsCd 
         Height          =   315
         Left            =   2325
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv1Nm 
         Height          =   315
         Left            =   3615
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      End
      Begin MedControls1.LisLabel lblOkDiv1Cd 
         Height          =   315
         Left            =   4560
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv2Nm 
         Height          =   315
         Left            =   5880
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      End
      Begin MedControls1.LisLabel lblOkDiv2Cd 
         Height          =   315
         Left            =   6825
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv3Nm 
         Height          =   315
         Left            =   8145
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      End
      Begin MedControls1.LisLabel lblOkDiv3Cd 
         Height          =   315
         Left            =   9105
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   75
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "현재상태"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   7
         Left            =   2625
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "접수결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   8
         Left            =   4875
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "문진결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   9
         Left            =   7140
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "검사결과"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   14
      Top             =   720
      Width           =   9945
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '평면
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Top             =   165
         Width           =   1485
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   330
         Left            =   4275
         TabIndex        =   15
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6630
         TabIndex        =   16
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
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
         Caption         =   "M/100"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   330
         Left            =   8940
         TabIndex        =   17
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCnt 
         Height          =   330
         Left            =   4275
         TabIndex        =   18
         Top             =   525
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTotVol 
         Height          =   330
         Left            =   6630
         TabIndex        =   19
         Top             =   510
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDonorID 
         Height          =   315
         Left            =   1020
         TabIndex        =   20
         Top             =   540
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSSN 
         Height          =   315
         Left            =   1800
         TabIndex        =   29
         Top             =   540
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   165
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
         Caption         =   "성   명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3285
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3285
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "헌혈횟수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   5640
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "성/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   5640
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   510
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "총 헌혈량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7950
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "cc"
         Height          =   180
         Left            =   7545
         TabIndex        =   21
         Top             =   660
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmBBS403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum TblColumn
    tcASK = 1
    tcYES
    tcNo
    tcISOK
    tcNORMAL
    tcASKCODE
End Enum

'2001-11-27추가
Private strSaveDonorId As String
Private strSaveDonorNm As String

'2001-11-27 추가
Private Sub cmdCallTest_Click()
    frmBBS411.Show
    frmBBS411.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS411.CallDonorNmLostFocus
End Sub

Private Sub cmdCancel_Click()
    Dim donorid As String
    Dim accdt As String
    Dim objSQL As clsBBSSQLStatement
    
    If tabAccDt.SelectedItem.Index > 1 Then
        '최종 접수일자가 아니다. 접수취소 할 수 없다.
        MsgBox "문진취소를 할 수 없습니다.", vbCritical, Me.Caption
        Exit Sub
    Else
        '헌혈자의 상태를 파악한다.
        If lblStsCd.Caption > DonorStatus.stsAskVerify Then
            MsgBox "문진취소를 할 수 없습니다.", vbCritical, Me.Caption
            Exit Sub
        End If
    End If
    
    donorid = lblDonorID.Caption
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Set objSQL = New clsBBSSQLStatement
'    objSql.setDbConn DBConn
    If objSQL.SetDonorStatus(donorid, accdt, DonorStatus.stsAskSave) = True Then
        FormInitialize
    End If
    Set objSQL = Nothing
End Sub

Private Sub cmdClear_Click()
    FormInitialize
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        FormInitialize
    End If
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    SetAsk
    FormInitialize
End Sub


Private Sub tabAccDt_Click()
    Dim donorid As String
    Dim donoraccdt As String
    Dim canEdit As Boolean
    
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    '이 헌혈자에 등록된 문진내역이 있으면 조회한다.
    Call SetDonorAsk(donorid, donoraccdt)
    Call SetDonorStatus(donorid, donoraccdt)
    Call SetDonorResult(donorid, donoraccdt)
    
    canEdit = GetCanEdit
    fraDonorCd.Enabled = canEdit
    fraResult.Enabled = canEdit
    With tblAsk
        .Col = -1
        .Row = -1
        .BlockMode = True
        tblAsk.Lock = Not canEdit
        .BlockMode = False
    End With
    
    cmdSave.Enabled = canEdit
    cmdCancel.Enabled = Not canEdit
End Sub

Private Function GetCanEdit() As Boolean
    '수정이 가능한지를 판단한다.
    If tabAccDt.SelectedItem.Index > 1 Then
        '최종 접수일자가 아니다. 수정할 수 없다.
        GetCanEdit = False
    Else
        Select Case lblStsCd.Caption
            Case DonorStatus.stsAccessSave
                GetCanEdit = False
            Case DonorStatus.stsAccessVerify
                GetCanEdit = (lblOkDiv1Cd.Caption = "1")
            Case DonorStatus.stsAskSave
                GetCanEdit = True
            Case DonorStatus.stsAskVerify
                GetCanEdit = False
            Case DonorStatus.stsDonation
                GetCanEdit = False
            Case DonorStatus.stsFinish
                GetCanEdit = False
            Case DonorStatus.stsPrint
                GetCanEdit = False
            Case Else
                GetCanEdit = False
        End Select
    End If
End Function

Private Sub SetDonorResult(ByVal donorid As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim DrRS As Recordset
    
    Set objDonor = New clsBBSSQLStatement
    Set DrRS = objDonor.GetDonorResult(donorid, accdt)
    Set objDonor = Nothing
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        If .RecordCount > 0 Then
            Select Case .Fields("okdiv2").value & ""
                Case 0:
                    optOk(1).value = True
                    optOk(0).value = False
                Case 1:
                    optOk(1).value = False
                    optOk(0).value = True
                Case Else:
                    optOk(1).value = False
                    optOk(0).value = False
            End Select
            txtrmk = .Fields("rmk2").value & ""
        End If
    End With
End Sub

Private Sub SetDonorStatus(ByVal donorid As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim strStatus As String
    Dim IsPhere As Boolean
    
    
    Set objDonor = New clsBBSSQLStatement
    strStatus = objDonor.GetDonorStatus(donorid, accdt, IsPhere)
    Set objDonor = Nothing
    
    lblStsNm.Caption = medGetP(strStatus, 1, vbTab)
    lblStsCd.Caption = medGetP(strStatus, 2, vbTab)
    lblOkDiv1Nm.Caption = medGetP(strStatus, 3, vbTab)
    lblOkDiv1Cd.Caption = medGetP(strStatus, 4, vbTab)
    lblOkDiv2Nm.Caption = medGetP(strStatus, 5, vbTab)
    lblOkDiv2Cd.Caption = medGetP(strStatus, 6, vbTab)
    lblOkDiv3Nm.Caption = medGetP(strStatus, 7, vbTab)
    lblOkDiv3Cd.Caption = medGetP(strStatus, 8, vbTab)
    
    If lblOkDiv1Nm.Caption = "부적격" Then
        lblOkDiv1Nm.ForeColor = vbRed
        lblOkDiv1Cd.ForeColor = vbRed
    Else
        lblOkDiv1Nm.ForeColor = vbBlack
        lblOkDiv1Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv2Nm.Caption = "부적격" Then
        lblOkDiv2Nm.ForeColor = vbRed
        lblOkDiv2Cd.ForeColor = vbRed
    Else
        lblOkDiv2Nm.ForeColor = vbBlack
        lblOkDiv2Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv3Nm.Caption = "부적격" Then
        lblOkDiv3Nm.ForeColor = vbRed
        lblOkDiv3Cd.ForeColor = vbRed
    Else
        lblOkDiv3Nm.ForeColor = vbBlack
        lblOkDiv3Cd.ForeColor = vbBlack
    End If
    
    
End Sub

Private Sub tblAsk_Click(ByVal Col As Long, ByVal Row As Long)
    Dim value As String
    Dim normal As String
    
    If Row < 1 Then Exit Sub
    
    
    With tblAsk
        
        If .Lock = True Then Exit Sub
    
        Select Case Col
            Case TblColumn.tcYES, TblColumn.tcNo:
                '정상치
                .Row = Row: .Col = TblColumn.tcNORMAL: normal = .value
                
                '체크박스 변경
                .Row = Row: .Col = Col: value = IIf(.value = 1, 0, 1)
                .value = value
                
                '정상인지 아닌지 OK,Not으로 표시
                .Row = Row: .Col = TblColumn.tcISOK:
                If value = 1 Then
                    If normal = "1" Then
                        .value = IIf(Col = TblColumn.tcYES, "OK", "Not")
                    Else
                        .value = IIf(Col = TblColumn.tcNo, "OK", "Not")
                    End If
                Else
                    .value = ""
                End If
                
                'Yes나 No중 하나만 선택할 수 있다.
                If value = "1" Then
                    .Row = Row
                    .Col = IIf(Col = TblColumn.tcNo, TblColumn.tcYES, TblColumn.tcNo)
                    .value = 0
                End If
        End Select
        
    End With
    
    '문진내역 전체에 대한 적격/부적격 여부 자동 셋팅
    Call SetOptOk
End Sub

Private Sub txtDonorNm_GotFocus()
    txtDonorNm.tag = txtDonorNm
End Sub

Private Sub txtDonorNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call DonorFind
        txtDonorNm.tag = txtDonorNm
    End If
End Sub

Private Sub txtDonorNm_LostFocus()
    If txtDonorNm.tag <> txtDonorNm Then
        Call DonorFind
    End If
End Sub

Private Sub DonorFind()
    Dim objDonor As clsBBSBldDonationBusi
    
    If txtDonorNm = "" Then Call FrameInitialize: Exit Sub
    
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor

        If .DonorFind(txtDonorNm) = True Then
            Call FrameInitialize
            
            lblDonorID.Caption = .mDonorID
            txtDonorNm = .mDonorNm
            '2001-11-27 추가
            strSaveDonorId = lblDonorID.Caption
            strSaveDonorNm = txtDonorNm.Text
            
            lblDOB.Caption = .mDOB
            lblSex.Caption = .mSEX
            lblABO.Caption = .mABO
            lblCnt.Caption = .Mcnt
            lblTotVol.Caption = .mTotVol
        
            Call ShowAccList
'            cmdNew.Enabled = True
        End If
    End With
    Set objDonor = Nothing
End Sub

Private Sub ShowAccList()
    Dim strAccDt As String
    Dim Rs As Recordset
    Dim objMySQL As clsBBSSQLStatement
    '헌혈자에 대해서 접수된 정보가 있을 경우에 접수 내역을 보여준다.

    Set objMySQL = New clsBBSSQLStatement

'    objMySQL.setDbConn DBConn
    Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    
    If Rs.EOF Then
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = False
    Else
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = True
        
        Do Until Rs.EOF
            strAccDt = Format(Rs.Fields("donoraccdt").value & "", "####-##-##")
            tabAccDt.Tabs.Add , , strAccDt
            Rs.MoveNext
        Loop
        
        cmdSave.Enabled = True
        Call tabAccDt_Click
    End If

End Sub

Private Sub SetAsk()
    Dim i As Long
    Dim Rs As Recordset
    
    
    tblAsk.MaxRows = 0
    
    Set Rs = ReadCom003(BC2_ASK)
    If Rs Is Nothing Then Exit Sub
    
    With tblAsk
        .MaxRows = Rs.RecordCount
        For i = 1 To Rs.RecordCount
            .Row = i
            .Col = TblColumn.tcASK:     .value = Rs.Fields("text1").value & ""
            .Col = TblColumn.tcNORMAL:  .value = Rs.Fields("field1").value & ""
            .Col = TblColumn.tcASKCODE: .value = Rs.Fields("cdval1").value & ""
            
            .RowHeight(i) = .MaxTextRowHeight(i)
            Rs.MoveNext
        Next i
    End With
        
    
    Set Rs = Nothing
End Sub

Private Sub SetOptOk()
    Dim r As Long
    Dim nosetcnt As Long
    
    With tblAsk
        nosetcnt = 0
        For r = 1 To .MaxRows
            .Row = r
            .Col = TblColumn.tcISOK
            If .value = "" Then
                nosetcnt = nosetcnt + 1
            ElseIf .value = "Not" Then
                optOk(1).value = True
                Exit Sub
            End If
        Next r
        
        If nosetcnt = .MaxRows Then
            optOk(0).value = False
            optOk(1).value = False
        Else
            optOk(0).value = True
        End If
    End With
    
End Sub

Private Sub SetDonorAsk(ByVal donorid As String, ByVal donoraccdt As String)
    Dim i As Long
    Dim r As Long
    
    Dim askcd As String
    Dim yesno As String
    Dim normal As String
    Dim okdiv As String
    
    Dim objDonorAsk As clsDonorAsk
    Dim DrRS As Recordset
    Dim DrRS1 As Recordset
    
    
    Dim objTestReq As clsBBSSQLStatement
    Dim RsTestReq As Recordset
    
    
    Call FrameInitialize
    
    
    
    Set objTestReq = New clsBBSSQLStatement
    With objTestReq
'        .setDbConn DBConn
        Set RsTestReq = .GetDonorAccHistory(donorid, donoraccdt)
    End With
    
    If Not RsTestReq.EOF Then
        '임시환자id
'        lblTmpPtId = RsTestReq.Fields("tmpid")
        
        '헌혈종류
        Select Case RsTestReq.Fields("donorcd").value & ""
            Case "0": cboDonorCd.ListIndex = 0
            Case "1": cboDonorCd.ListIndex = 1
            Case "2": cboDonorCd.ListIndex = 2
            Case "3": cboDonorCd.ListIndex = 3
            Case Else
                      cboDonorCd.ListIndex = -1
        End Select
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
        If txtReservedID <> "0" Then
            lblReservedNm.Caption = objTestReq.GetPtntNm(txtReservedID)
        Else
            lblReservedNm.Caption = ""
        End If
    End If
    Set RsTestReq = Nothing
    Set objTestReq = Nothing
    
    
    Set objDonorAsk = New clsDonorAsk
    Set DrRS = objDonorAsk.GetDonorAsk(donorid, donoraccdt)
    If DrRS Is Nothing Then
        Set objDonorAsk = Nothing
        Exit Sub
    End If
    
    If DrRS.RecordCount > 0 Then
        With tblAsk
            For i = 1 To DrRS.RecordCount
                '스프레트에 있는 문진크도와
                '데이터베이스에서 읽은 문진코드가
                '동일한 Row에 대하여 작업한다.
                For r = 1 To .MaxRows
                    .Row = r
                    .Col = TblColumn.tcASKCODE
                    askcd = Trim(.value)
                    
                    If askcd = Trim(DrRS.Fields("askcd").value & "") Then
                        yesno = DrRS.Fields("yesno").value & ""
                        okdiv = DrRS.Fields("okdiv").value & ""
                        
                        If yesno = "1" Then
                            .Col = TblColumn.tcYES: .value = 1
                            .Col = TblColumn.tcNo: .value = 0
                        ElseIf yesno = "0" Then
                            .Col = TblColumn.tcYES: .value = 0
                            .Col = TblColumn.tcNo: .value = 1
                        Else
                            .Col = TblColumn.tcYES: .value = 0
                            .Col = TblColumn.tcNo: .value = 0
                        End If
                        
                        If okdiv = "1" Then
                            .Col = TblColumn.tcISOK: .value = "OK"
                        ElseIf okdiv = "0" Then
                            .Col = TblColumn.tcISOK: .value = "Not"
                        Else
                            .Col = TblColumn.tcISOK: .value = ""
                        End If
                        
                        Exit For
                    End If
                Next r
                DrRS.MoveNext
            Next i
        End With
        
        Set DrRS1 = objDonorAsk.GetDonorOk(donorid, donoraccdt)
        If Not (DrRS1 Is Nothing) Then
            With DrRS1
                If .RecordCount > 0 Then
                    If .Fields("okdiv2").value & "" = 1 Then
                        optOk(0).value = True
                    Else
                        optOk(1).value = True
                    End If
                    txtrmk = .Fields("rmk2").value & ""
                End If
            End With
            Set DrRS1 = Nothing
        End If
    End If
    
    Set DrRS = Nothing
End Sub

Private Sub FormInitialize()
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
    
    txtDonorNm = ""
    lblDonorID.Caption = ""
    lblDOB.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""

    Call FrameInitialize
End Sub

Private Sub FrameInitialize()
    Dim r As Long
    
    txtReservedID = ""
    lblReservedNm.Caption = ""
    
    lblStsNm.Caption = ""
    lblStsCd.Caption = ""
    lblOkDiv1Nm.Caption = ""
    lblOkDiv1Cd.Caption = ""
    lblOkDiv2Nm.Caption = ""
    lblOkDiv2Cd.Caption = ""
    lblOkDiv3Nm.Caption = ""
    lblOkDiv3Cd.Caption = ""
    
    cboDonorCd.ListIndex = -1
    
    With tblAsk
        For r = 1 To .MaxRows
            .Row = r
            .Col = TblColumn.tcYES:  .value = 0
            .Col = TblColumn.tcNo:   .value = 0
            .Col = TblColumn.tcISOK: .value = ""
        Next r
    End With
    
    optOk(0).value = False
    optOk(1).value = False
    txtrmk = ""
    
End Sub

Private Function Save() As Boolean
    Dim donorid As String
    Dim donoraccdt As String
    Dim okdiv As String
    Dim rmk As String
    Dim r As Long
    Dim tmpOkDiv As String
    
    Dim askcd As String
    Dim yes As String, no As String
    Dim askyesno As String
    Dim askokdiv As String
    
    Dim strAsk As String
    
    Dim objDonorAsk As clsDonorAsk
    
    Dim IsHold   As Boolean
    Dim chk As Boolean
    
    '저장하기 위한 모든 값이 셋팅되었는지 점검한다.-----------------------------------------------
    With tblAsk
        chk = True
        tmpOkDiv = 1
        For r = 1 To .MaxRows
            .Row = r
            .Col = TblColumn.tcISOK
            If .value = "" Then
                chk = False
                Exit For
            ElseIf .value = "Not" Then
                tmpOkDiv = 0
            End If
        Next r
    End With
    
    If chk = False Then
        If MsgBox("결과가 입력되지 않는 문진내역이 있읍니다. 저장하시겠읍니까?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Save = False
            Exit Function
        End If
    End If
    
    '이런 경우는 없겠지만, 혹시나
    If optOk(0).value = False And optOk(1).value = False Then
        MsgBox "판정결과가 없습니다.", vbCritical, Me.Caption
        Save = False
        Exit Function
    End If
    '문진내역의 자동판정과 판정결과가 상이할때는 remark를 입력하여야 한다.
    If tmpOkDiv = 1 Then
        If optOk(1).value = True And Trim(txtrmk) = "" Then
            MsgBox "자동판정과 판정결과가 다르므로 Remark를 입력하여야 합니다", vbCritical, Me.Caption
            Save = False
            Exit Function
        End If
    Else
        If optOk(0).value = True And Trim(txtrmk) = "" Then
            MsgBox "자동판정과 판정결과가 다르므로 Remark를 입력하여야 합니다", vbCritical, Me.Caption
            Save = False
            Exit Function
        End If
    End If
    
    
    '저장하기 위한 값을 구한다.-----------------------------------------------------------------------
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    okdiv = IIf(optOk(0).value = True, "1", "0")
    rmk = txtrmk
    
    With tblAsk
        strAsk = ""
        For r = 1 To .MaxRows
            .Row = r
            .Col = TblColumn.tcASKCODE: askcd = .value
            .Col = TblColumn.tcYES:     yes = .value
            .Col = TblColumn.tcNo:      no = .value
                                        If yes = "1" Then
                                            askyesno = "1"
                                        ElseIf no = "1" Then
                                            askyesno = "0"
                                        Else
                                            askyesno = ""
                                        End If

            .Col = TblColumn.tcISOK:
                                        If .value = "OK" Then
                                            askokdiv = "1"
                                        ElseIf .value = "Not" Then
                                            askokdiv = "0"
                                        Else
                                            askokdiv = ""
                                        End If
            
            If strAsk = "" Then
                strAsk = askcd & COL_DIV & askyesno & COL_DIV & askokdiv
            Else
                strAsk = strAsk & LINE_DIV & _
                         askcd & COL_DIV & askyesno & COL_DIV & askokdiv
            End If
        Next r
    End With
    
    
    IsHold = (chkHold.value = 1)
    
    Set objDonorAsk = New clsDonorAsk
    
    Save = objDonorAsk.Save(donorid, donoraccdt, okdiv, rmk, strAsk, IsHold)
    
    Set objDonorAsk = Nothing
End Function

'2001-11-27추가
Public Sub CallDonorNmLostFocus()
    Call txtDonorNm_LostFocus
End Sub


