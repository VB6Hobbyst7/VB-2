VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS106 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Specimen Accept"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmBBS106.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Index           =   2
      Left            =   5010
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   9450
      _ExtentX        =   16669
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
      Caption         =   "처방정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1395
      Left            =   5010
      TabIndex        =   5
      Top             =   315
      Width           =   9465
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "BarCode 입력"
         Height          =   240
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtSpcNo 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   765
         Width           =   1755
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   765
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
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin VB.Label lblAddInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "추가 요청 검체입니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   3420
         TabIndex        =   9
         Top             =   900
         Width           =   2115
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Index           =   0
      Left            =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   75
      Width           =   4860
      _ExtentX        =   8573
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
      Caption         =   "검체 접수 옵션"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   6825
      Left            =   5010
      TabIndex        =   25
      Top             =   1650
      Width           =   9465
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "frmBBS106.frx":076A
         Top             =   1695
         Width           =   9180
      End
      Begin MedControls1.LisLabel lblPtID 
         Height          =   360
         Left            =   1185
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   450
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   3975
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   450
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MedControls1.LisLabel lblColDtTm 
         Height          =   360
         Left            =   1185
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   870
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MedControls1.LisLabel lblColNm 
         Height          =   360
         Left            =   3975
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   870
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   6765
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   450
         Width           =   855
         _ExtentX        =   1508
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
      Begin FPSpread.vaSpread tblOrderList 
         Height          =   3345
         Left            =   135
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "15109"
         Top             =   3390
         Width           =   9165
         _Version        =   196608
         _ExtentX        =   16166
         _ExtentY        =   5900
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         GrayAreaBackColor=   15003117
         MaxCols         =   11
         MaxRows         =   25
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS106.frx":079C
         VisibleRows     =   25
      End
      Begin MedControls1.LisLabel lblReaction 
         Height          =   360
         Left            =   6090
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   870
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   635
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
         Height          =   360
         Left            =   5700
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   870
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
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
         Index           =   0
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   450
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
         Left            =   2910
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   450
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
         Left            =   5700
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   450
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
         Left            =   2910
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   870
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
         Index           =   7
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   870
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
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2760
         _ExtentX        =   4868
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
         Caption         =   "Order REMARK"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   135
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3000
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
         Caption         =   "처방 정보"
         Appearance      =   0
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Height          =   855
         Left            =   7740
         TabIndex        =   43
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label lblABO 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "AB+"
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
         Left            =   8070
         TabIndex        =   42
         Top             =   660
         Width           =   840
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   2535
      Left            =   75
      TabIndex        =   11
      Top             =   315
      Width           =   4875
      Begin VB.ComboBox cboLeg 
         Height          =   300
         Left            =   2640
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1035
      End
      Begin VB.CheckBox chkAuto 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체번호 입력즉시 접수"
         Height          =   240
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox txtLeg 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   2100
         Width           =   675
      End
      Begin VB.TextBox txtRow 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   3300
         TabIndex        =   14
         Top             =   2100
         Width           =   675
      End
      Begin VB.TextBox txtCol 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   2100
         Width           =   675
      End
      Begin VB.CheckBox chkSPos 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체보관장소 자동지정"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack Select"
         Height          =   180
         Left            =   1560
         TabIndex        =   23
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Height          =   600
         Left            =   1560
         TabIndex        =   22
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   21
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Caption         =   "Row"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   20
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Caption         =   "Col"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   19
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보관장소"
         Height          =   180
         Left            =   1740
         TabIndex        =   18
         Top             =   2040
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00F4F0F2&
      Caption         =   "접수(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8565
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblSpcList 
      Height          =   5280
      Left            =   75
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "15109"
      Top             =   3195
      Width           =   4845
      _Version        =   196608
      _ExtentX        =   8546
      _ExtentY        =   9313
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15003117
      MaxCols         =   4
      MaxRows         =   21
      OperationMode   =   3
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS106.frx":0CBC
      VisibleRows     =   20
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2850
      Width           =   4860
      _ExtentX        =   8573
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
      Caption         =   "검체 접수 리스트"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcTESTNM = 1
    tcORDCD
    tcUNITQTY
    tcORDDT
    tcREQDT
    tcREASON
    tcACCDTSEQ
    tcORDNO
    tcORDSEQ
    tcDCFG
    tcORDDIV
End Enum
Private Enum TblColumn2
    tcSPCNM = 1
    tcSAVEPOS
    tcPTNM
    TcABO
End Enum
    
Private blnAdd  As Boolean '추가채혈 여부 판단(T:추가채혈,F:처방채혈)
Private onPgm   As Boolean
Private blnSpcCancel As Boolean '접수취소된 검체접수여부

Private Sub chkAuto_Click()
    If chkAuto.value = 1 Then
        chkSPos.Enabled = False
        cboLeg.Enabled = True
        txtLeg = ""
        txtRow = ""
        txtCol = ""
        txtLeg.Locked = True
        txtRow.Locked = True
        txtCol.Locked = True
        txtLeg.BackColor = Me.BackColor
        txtRow.BackColor = Me.BackColor
        txtCol.BackColor = Me.BackColor
        txtLeg = cboLeg.Text
    Else
        chkSPos.Enabled = True
        If chkSPos.value = 1 Then
            cboLeg.Enabled = True
            txtLeg = ""
            txtRow = ""
            txtCol = ""
            txtLeg.Locked = True
            txtRow.Locked = True
            txtCol.Locked = True
            txtLeg.BackColor = Me.BackColor
            txtRow.BackColor = Me.BackColor
            txtCol.BackColor = Me.BackColor
            txtLeg = cboLeg.Text
        Else
            cboLeg.Enabled = False
            txtLeg.Locked = False
            txtRow.Locked = False
            txtCol.Locked = False
            txtLeg.BackColor = RGB(255, 255, 255)
            txtRow.BackColor = RGB(255, 255, 255)
            txtCol.BackColor = RGB(255, 255, 255)
        End If
    End If
End Sub

Private Sub chkBar_Click()
    If onPgm = False Then txtSpcNo.SetFocus
End Sub

Private Sub chkSPos_Click()
    If chkAuto.value = 1 Then
        txtLeg = ""
        txtRow = ""
        txtCol = ""
        txtLeg.Locked = True
        txtRow.Locked = True
        txtCol.Locked = True
        txtLeg.BackColor = Me.BackColor
        txtRow.BackColor = Me.BackColor
        txtCol.BackColor = Me.BackColor
        txtLeg = cboLeg.Text
        cboLeg.Enabled = True
    Else
        If chkSPos.value = 1 Then
            txtLeg = ""
            txtRow = ""
            txtCol = ""
            txtLeg.Locked = True
            txtRow.Locked = True
            txtCol.Locked = True
            txtLeg.BackColor = Me.BackColor
            txtRow.BackColor = Me.BackColor
            txtCol.BackColor = Me.BackColor
            txtLeg = cboLeg.Text
            cboLeg.Enabled = True
        Else
            txtLeg.Locked = False
            txtRow.Locked = False
            txtCol.Locked = False
            txtLeg.BackColor = RGB(255, 255, 255)
            txtRow.BackColor = RGB(255, 255, 255)
            txtCol.BackColor = RGB(255, 255, 255)
            cboLeg.Enabled = False
        End If

    End If
End Sub

Private Sub cmdClear_Click()
    Clear
    txtSpcNo.SetFocus
End Sub

Private Sub Clear()
    tblOrderList.MaxRows = 0: tblOrderList.MaxRows = 20
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblColDtTm.Caption = ""
    lblColNm.Caption = ""
    lblABO.Caption = ""
    lblSexAge.Caption = ""
    lblAddInfo.Caption = ""
    txtSpcNo.Text = ""
    txtRemark.Text = ""
    txtLeg = ""
    txtRow = ""
    txtCol = ""
    lblInfection.Visible = False
    lblReaction.Visible = False
End Sub

Private Sub Form_Setting()
    Dim objAccess As clsBBSAccess
    Dim objNumbers As clsBBSNumbers
    Dim DrRS As New Recordset
    
    Set objAccess = New clsBBSAccess
    Set objNumbers = New clsBBSNumbers
    
    With objAccess
        DrRS.Open .Get_LegPos(ObjSysInfo.BuildingCd), DBConn
        If DrRS.EOF = False Then
            cboLeg.AddItem ""
            Do Until DrRS.EOF = True
                cboLeg.AddItem DrRS.Fields("legcd").value & ""
                DrRS.MoveNext
            Loop
        End If
    End With

    cmdCollect.Enabled = False
    Set objNumbers = Nothing
    Set objAccess = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    onPgm = True
    
    chkBar.value = 0
    chkAuto.value = 1
    chkSPos.value = 1
    chkBar.value = 1

    Clear
    Form_Setting
    
    onPgm = False
    
    
    Me.Show
    txtSpcNo.SetFocus
End Sub

Private Sub DetailSearch()
'혈액형,부작용,감염정보
    Dim ObjABO As New clsABO
    Dim objinfection As New clsInfection
    Dim objReaction As New clsReaction
    
    With ObjABO
        .PtId = lblPtId.Caption
        .GetABO
        lblABO.Caption = .ABO & .Rh
    End With
    With objinfection
        .PtId = lblPtId.Caption
        .GetInfection
        If .Infection = True Then
            lblInfection.Visible = True
        Else
            lblInfection.Visible = False
        End If
        
    End With
    
    With objReaction
        .PtId = lblPtId.Caption
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


Private Sub cmdCollect_Click()
    Dim spcyy As String
    Dim spcno As Long
    
    If chkBar.value = 1 Then
        spcyy = Mid(txtSpcNo, 1, 2)
        spcno = Val(Mid(txtSpcNo, 3, 9))
    Else
        spcyy = medGetP(txtSpcNo, 1, "-")
        spcno = Val(medGetP(txtSpcNo, 2, "-"))
    End If
    
    
    If Save_Position(spcyy, spcno) = True Then
        With tblSpcList
            If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .DataRowCnt + 1
            If chkBar.value = 0 Then
                .Col = TblColumn2.tcSPCNM: .value = txtSpcNo
            Else
                .Col = TblColumn2.tcSPCNM: .value = Mid(txtSpcNo, 1, 2) & "-" & Format(Mid(txtSpcNo, 3, 9), "#########")
            End If
            .Col = TblColumn2.tcSAVEPOS:  .value = txtLeg & "(" & txtRow & "," & txtCol & ")"
            .Col = TblColumn2.tcPTNM:     .value = lblPtNm.Caption
            .Col = TblColumn2.TcABO:      .value = lblABO.Caption
        End With
        
        Call cmdClear_Click
        MsgBox "접수되었습니다.", vbInformation + vbOKOnly, Me.Caption
        cmdCollect.Enabled = False
    Else
        MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub txtCol_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRow_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSpcNo_Change()
    Dim lngLen As Long
    
    If chkBar.value = 1 Then Exit Sub
    
    With txtSpcNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtSpcNo_GotFocus()
    txtSpcNo.tag = txtSpcNo
    txtSpcNo.SelStart = 0
    txtSpcNo.SelLength = Len(txtSpcNo)
End Sub

Private Sub txtSpcNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtSpcNo) <> 3 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtSpcNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtSpcNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtSpcNo = "" Then Exit Sub
        
        If Spc_ExistChk(txtSpcNo) = False Then
            txtSpcNo = ""
        Else
            cmdCollect.Enabled = True
            If chkAuto.value = 1 Then
                Call cmdCollect_Click
                txtSpcNo.Text = "": txtSpcNo.SetFocus
            End If
        End If
       
    End If
End Sub

Private Function Spc_ExistChk(ByVal SpcNum As String) As Boolean
    Dim Rs        As Recordset
    Dim QRS       As Recordset
    Dim objAccess As clsBBSAccess
    Dim spcyy     As String
    Dim spcno     As Long
    
    
    If chkBar.value = 1 Then
        spcyy = Mid(SpcNum, 1, 2)
        spcno = Val(Mid(SpcNum, 3, 9))
    Else
        spcyy = Mid(SpcNum, 1, 2)
        spcno = Val(Mid(SpcNum, 4, 9))
    End If
    
    Set objAccess = New clsBBSAccess
    tblOrderList.MaxRows = 0
    blnSpcCancel = False
    
    ' 1. 해당검체가 채혈내역에 존재하는 지를 판단.(True:있다,False:없다)
    ' 2. 접수취소가 되었을경우(예: 처방3건이 하나의 검체를 가지고 있을경우, 처방한건을 접수취소시켰다.)
    ' 3. 취소된처방에 대해서 검체접수를 하여야 한다.
    ' 방법: 1. 채혈내역에서 접수되지 않은검체가 있는지 조회한다.(rcvdt가 ''인경우)
    '       2. 검체번호에 따른 환자ID를 가지고, 처방접수내역(혈액은행,BBS202)에서 stscd='1' and  canclefg가 =1인경우를 찾는다
    '       3. BBS202에서 처방번호,처방Seq를 구하여 접수처리를 한다.
    '       3. 접수완료시 BBS202에 따른 취소 여부를 ''로 update 해준다.
    
      
    
    If objAccess.Spc_ExistChk(spcyy, spcno) = True Then
        '접수 안된 검체
        If GetAddSpcChk(spcyy, spcno) = False Then
            Set objAccess = Nothing
            Exit Function
        End If
        
        Call SpreadDisplay(, spcyy, spcno)
        Call DetailSearch
    Else
        '이미 접수된검체
        '접수취소가 되었을경우..
        Set Rs = objAccess.GetAccCancelSpcInfo(spcyy, spcno)
        Dim sPtid  As String
        Dim sRcvDt As String
        Dim sRcvTm As String
        
        blnSpcCancel = True
        If Not Rs.EOF Then
            If GetAddSpcChk(spcyy, spcno) = False Then
                Set Rs = Nothing: Set objAccess = Nothing
                Exit Function
            End If
            Call DetailSearch
            
            Do Until Rs.EOF
                sPtid = Rs.Fields("ptid").value & ""
                sRcvDt = Rs.Fields("rcvdt").value & ""
                sRcvTm = Rs.Fields("rcvtm").value & ""
                Call SpreadDisplay(objAccess.GetAccCancelOrdInfo(sPtid, sRcvDt, sRcvTm))
                Rs.MoveNext
            Loop
            
        Else
            Call Clear
            txtSpcNo = ""
            MsgBox "입력하신 검체가 존재하지 않습니다." & vbNewLine & _
                   "이미 접수하신 검체이거나 존재하지 않는 검체일수 있습니다", vbInformation + vbOKOnly, "검체접수"
            Set objAccess = Nothing
            Exit Function
        End If
    End If
    
    Spc_ExistChk = True
    Set objAccess = Nothing
End Function

Private Function GetAddSpcChk(ByVal spcyy As String, ByVal spcno As Long) As Boolean
    Dim objAcc      As New clsBBSAccess
    Dim Rs          As Recordset
    Dim strSDA      As String
    
    Set Rs = objAcc.Get_SpcInFormation(spcyy, spcno)
    
    If Not Rs.EOF Then
        With Rs
            '이 건물에서 검사할 수 있는지 검사한다.-------------------------------------
            If .Fields("buildcd").value & "" <> ObjSysInfo.BuildingCd Then
                MsgBox "다른 센터에서 검사해야하는 검체입니다.", vbCritical, "오류"
                Set Rs = Nothing
                Set objAcc = Nothing
                Exit Function
            End If
            
            lblPtId.Caption = .Fields("ptid").value & ""
            lblPtNm.Caption = .Fields("ptnm").value & ""
            lblColDtTm.Caption = Format(.Fields("coldt").value & "", "####-##-##") & " " & _
                                 Format(Mid(.Fields("coltm").value & "", 1, 4), "##:##")
            
            strSDA = SDA_String(.Fields("ssn").value & "")
            lblSexAge.Caption = medGetP(strSDA, 1, COL_DIV) & "/" & medGetP(strSDA, 3, COL_DIV)
            
            lblColNm.Caption = .Fields("empnm").value & ""
            
            If .Fields("addfg").value & "" = "1" Then
                lblAddInfo.Caption = "추가채혈된 검체입니다.."
                blnAdd = True
            Else
                lblAddInfo.Caption = ""
                blnAdd = False
            End If
        End With
    End If
    
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    
    GetAddSpcChk = True
    
    Set objAcc = Nothing
    Set Rs = Nothing

End Function

Private Sub SpreadDisplay(Optional ByVal tmpSQL As String = "", _
                          Optional ByVal spcyy As String = "", Optional ByVal spcno As Long)
    Dim objAcc         As New clsBBSAccess
    Dim objTransReason As clsQueryOrder
    
    Dim Rs      As Recordset
    Dim AccdtSeq As String
    Dim strReason As String
    
    
    
    If tmpSQL = "" Then
        If blnAdd = True Then
            AccdtSeq = objAcc.Get_AccDtSeq(spcyy, spcno)
            Set Rs = objAcc.Get_SpcOrderList(lblPtId.Caption, medGetP(AccdtSeq, 1, COL_DIV), medGetP(AccdtSeq, 2, COL_DIV))
        Else
            Set Rs = objAcc.Get_SpcOrderList(lblPtId.Caption)
        End If
    Else
        Set Rs = New Recordset
        Rs.Open tmpSQL, DBConn
    End If
    
    Set objTransReason = New clsQueryOrder
    If Not Rs.EOF Then
        With tblOrderList
            Do Until Rs.EOF
                .MaxRows = .DataRowCnt + 1
                .Row = .MaxRows
                strReason = objTransReason.GetTransReason(lblPtId.Caption, Rs.Fields("orddt").value & "", Rs.Fields("ordno").value & "")
                .Col = TblColumn.tcTESTNM:   .value = Rs.Fields("testnm").value & ""
                .Col = TblColumn.tcORDCD:    .value = Rs.Fields("ordcd").value & ""
                .Col = TblColumn.tcUNITQTY:  .value = CLng(Rs.Fields("unitqty").value & "")
                .Col = TblColumn.tcORDDT:    .value = Format(Rs.Fields("orddt").value & "", "####-##-##")
                .Col = TblColumn.tcREQDT:    .value = Format(Rs.Fields("reqdt").value & "", "####-##-##")
                .Col = TblColumn.tcREASON:   .value = strReason
                .Col = TblColumn.tcACCDTSEQ: .value = Rs.Fields("accdt").value & "" & "-" & Val(Rs.Fields("accseq").value & "")
                                             If .value = "-0" Then .value = ""
                .Col = TblColumn.tcORDNO:    .value = CLng(Rs.Fields("ordno").value & "")
                .Col = TblColumn.tcORDSEQ:   .value = CLng(Rs.Fields("ordseq").value & "")
                                             .ForeColor = vbRed
                .Col = TblColumn.tcDCFG:     .value = IIf(Rs.Fields("dcfg").value & "" = "1", "Y", "")
                                             .ForeColor = vbBlack
                .Col = TblColumn.tcORDDIV:   .value = Rs.Fields("orddiv").value & ""
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set objTransReason = Nothing
    Set objAcc = Nothing


End Sub

Private Function AddSpcChk(ByVal spcyy As String, ByVal spcno As Long) As Boolean
'
End Function

Private Function Save_Position(ByVal spcyy As String, ByVal spcno As Long) As Boolean
'저장장소를 자동지정 or 임의 지정여부로 저장한다.
    Dim objAcess As clsBBSAccess
    Dim strLeg   As String
    Dim lngRow   As Long
    Dim lngCol   As Long
    Dim SSQL     As String
    Dim ii       As Integer: ii = 0
    
    Set objAcess = New clsBBSAccess
    If blnSpcCancel = False Then
        With objAcess
    '        .setDbConn DBConn
            If chkAuto.value = 1 Then         '자동 접수시
                If .Get_Position(1, ObjSysInfo.BuildingCd, cboLeg.List(cboLeg.ListIndex)) Then
                    strLeg = .Leg(1)
                    lngRow = .Row(1)
                    lngCol = .Col(1)
                Else
                    Save_Position = False
                    Exit Function
                End If
            Else                              '자동접수가 아니면서 보관장소지정여부
                If chkSPos.value = 1 Then     '보관장소 자동지정
                    If .Get_Position(1, ObjSysInfo.BuildingCd, txtLeg) Then
                        strLeg = .Leg(1)
                        lngRow = .Row(1)
                        lngCol = .Col(1)
                    Else
                        Save_Position = False
                        Exit Function
                    End If
                Else                          '보관장소 지정
                    strLeg = txtLeg: lngRow = Val(txtRow): lngCol = Val(txtCol)
                    If .SavePointChk(strLeg, lngRow, lngCol, ObjSysInfo.BuildingCd) = False Then
                        Set objAcess = Nothing
                        Exit Function
                    End If
                End If
            End If
        End With
    End If
    
    
On Error GoTo Save_Spc_Error

    DBConn.BeginTrans
    If blnSpcCancel = False Then
        '채혈내역 update
        SSQL = objAcess.Set_UpdateB201(spcyy, spcno, ObjMyUser.EmpId, strLeg, lngRow, lngCol)
        DBConn.Execute SSQL
        
        '보관장소 update
        SSQL = objAcess.Set_UpdateB206(ObjSysInfo.BuildingCd, strLeg, lngRow, lngCol, spcyy, spcno)
        DBConn.Execute SSQL
    Else
        Call objAcess.GetSpcSavePosition(spcyy, spcno, strLeg, lngCol, lngRow)
    End If
    txtLeg = strLeg: txtRow = lngRow: txtCol = lngCol
    
    If blnAdd = False Then
        '처방에대한 update
        If Order_Acess(strLeg, lngRow, lngCol) = False Then GoTo Save_Spc_Error
    Else
        '추가 처방인경우 203의 환자에 대한 검체정보를 업데이트 해줘야 한다.
        Dim objCollect  As New clsBBSCollection
        Dim strSQL      As String
        
        strSQL = objCollect.Set_UpdateSQL_203(lblPtId.Caption, spcyy, CStr(spcno))
        If strSQL <> "" Then DBConn.Execute strSQL
        Set objCollect = Nothing
    End If
    
    DBConn.CommitTrans
    Save_Position = True
    Set objAcess = Nothing
    Exit Function
    
Save_Spc_Error:
    DBConn.RollbackTrans
    Error_TableClear
    Save_Position = False
    Set objAcess = Nothing
    MsgBox Err.Description, vbExclamation
End Function
Public Function Order_Acess(ByVal Leg As String, _
                            ByVal Row As Long, _
                            ByVal Col As Long) As Boolean
'처방을 update 해준다.
    Dim objAcess   As clsBBSAccess
    Dim objNum     As clsBBSNumbers
    Dim accdt      As String
    Dim accseq     As Long
    Dim PtId       As String
    Dim orddt      As String
    Dim ordno      As Long
    Dim Ordseq     As Long
    Dim SSQL       As String
    Dim BlnNum     As Boolean
    
    
    Dim ii As Integer
        
    Set objNum = New clsBBSNumbers
    Set objAcess = New clsBBSAccess
    
    With objNum
        accdt = .Get_AccdtFormat
        accseq = Val(.Get_AccDT_Seq(accdt))
        Set objNum = Nothing
    End With
    
    PtId = lblPtId.Caption
    
On Error GoTo Save_SpcOrder_Error
    '------------------------------------------------------
    '1:헤더저장
    '2:바디저장
    '3:처방접수내역저장
    '4:길병원의 경우 접수시에 BBS203을 먼저 생성해준다.
    '------------------------------------------------------
    With tblOrderList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcORDDT:  orddt = Replace(.value, "-", "")
            .Col = TblColumn.tcORDNO:  ordno = Val(.value)
            .Col = TblColumn.tcORDSEQ: Ordseq = Val(.value)
            
            .Col = TblColumn.tcORDDIV
            If .value <> "Z" Then
                .Col = TblColumn.tcACCDTSEQ
                If .value = "" Then
                        '1
                        SSQL = objAcess.Set_UpdateL101(PtId, orddt, ordno)
                        DBConn.Execute SSQL
                        .Col = TblColumn.tcDCFG
                        If .value = "" Then
                        '2
                            SSQL = objAcess.Set_UpdateL102(PtId, orddt, ordno, Ordseq, accdt, accseq)
                            DBConn.Execute SSQL
                            '3
                            If blnSpcCancel = False Then
                                '접수취소가 않된경우
                                SSQL = objAcess.Set_InsertB202(accdt, accseq, PtId, orddt, ordno, Ordseq, ObjMyUser.EmpId)
                                DBConn.Execute SSQL
                            Else
                                '접수취소후 접수하는경우
                                SSQL = objAcess.Set_UpdateB202(PtId, orddt, ordno, Ordseq)
                                DBConn.Execute SSQL
                            End If
                            '4

                            Dim objCollect As New clsBBSCollection
                            Dim SQLTmp As String
                            SQLTmp = objCollect.Set_AccUnitSQL_203(PtId, accdt, CStr(accseq))
                            SSQL = medGetP(SQLTmp, 1, COL_DIV)
                            DBConn.Execute SSQL
                            SSQL = medGetP(SQLTmp, 2, COL_DIV)
                            DBConn.Execute SSQL
                            Set objCollect = Nothing
                            
                            .Col = TblColumn.tcACCDTSEQ: .value = accdt & "-" & accseq
                            accseq = accseq + 1
'                            If OCSActingCheck(PtId, orddt, ordno, Ordseq) = False Then GoTo Save_SpcOrder_Error
                        End If
                    BlnNum = True
                End If
            End If
        Next
        
        If BlnNum = True Then
            '-------------
            '접수번호 저장
            '-------------
            SSQL = objAcess.Set_AccessUpdate(accdt, accseq - 1)
            DBConn.Execute SSQL
        End If
        
    End With
    
    Order_Acess = True
    Set objAcess = Nothing
    Exit Function
    
Save_SpcOrder_Error:
    Order_Acess = False
    Set objAcess = Nothing

End Function

'Private Function OCSActingCheck(ByVal strPtid As String, ByVal strOrdDt As String, _
'                                ByVal strOrdNo As String, ByVal strOrdSeq As String) As Boolean
'    Dim RS          As Recordset
'    Dim SqlStmt     As String
'    Dim strOcsOrdNo As String
'    Dim strBussdiv  As String
'
'On Error GoTo Errors
'
'    '접수시 OCS 관련 Table 에 Acting_Check를 해준다.
'
'    SqlStmt = " SELECT a.ocsordno,b.bussdiv " & _
'              " FROM " & T_LAB101 & " b," & T_LAB102 & " a" & _
'              " WHERE " & DBW("a.ptid =", strPtid) & _
'              " AND " & DBW("a.orddt=", strOrdDt) & _
'              " AND " & DBW("a.ordno=", strOrdNo) & _
'              " AND " & DBW("a.ordseq=", strOrdSeq) & _
'              " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
'
'    Set RS = New Recordset
'    RS.Open SqlStmt, DBConn
'
'    If Not RS.EOF Then
'        strOcsOrdNo = Val(Trim(RS.Fields("ocsordno").value & ""))
'        strBussdiv = Trim(RS.Fields("bussdiv").value & "")
'        If strOcsOrdNo <> "" Then
'            '병동은 ipd_order_dmc,ipd_order_update_dmc 업데이트
'            '외래는 opd_order_dmc 업데이트
'            If strBussdiv = enBussDiv.BussDiv_InPatient Then
'                SqlStmt = " UPDATE med_ocs.ipd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'                DBConn.Execute SqlStmt
'                SqlStmt = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            Else
'                SqlStmt = " UPDATE med_ocs.opd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'                DBConn.Execute SqlStmt
'            End If
'        End If
'    End If
'
'    Set RS = Nothing
'    OCSActingCheck = True
'    Exit Function
'
'Errors:
'    Set RS = Nothing
'    OCSActingCheck = False
'End Function

Private Sub Error_TableClear()
    Dim ii As Integer
    
    With tblOrderList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcACCDTSEQ: .value = ""
        Next
    End With
End Sub

