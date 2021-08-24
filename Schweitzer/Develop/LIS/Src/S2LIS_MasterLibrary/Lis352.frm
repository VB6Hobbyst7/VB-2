VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm352Specimen 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis352.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel lblTestName 
      Height          =   375
      Left            =   2925
      TabIndex        =   65
      Top             =   480
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   661
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
      Caption         =   ""
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00EBEBEB&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7770
      Style           =   1  '그래픽
      TabIndex        =   64
      Tag             =   "124"
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00EBEBEB&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7335
      Style           =   1  '그래픽
      TabIndex        =   63
      Tag             =   "124"
      Top             =   465
      Width           =   435
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
      Left            =   8235
      Style           =   1  '그래픽
      TabIndex        =   59
      Tag             =   "128"
      Top             =   330
      Width           =   1320
   End
   Begin VB.CommandButton cmdPopupList 
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
      Height          =   375
      Left            =   1200
      MousePointer    =   14  '화살표와 물음표
      Picture         =   "Lis352.frx":038A
      Style           =   1  '그래픽
      TabIndex        =   51
      Top             =   480
      Width           =   390
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
      Left            =   9570
      Style           =   1  '그래픽
      TabIndex        =   34
      Tag             =   "128"
      Top             =   330
      Width           =   1320
   End
   Begin VB.Frame fraRefRange 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4200
      Left            =   180
      TabIndex        =   31
      Tag             =   "35208"
      Top             =   4470
      Width           =   4590
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   405
         Left            =   255
         TabIndex        =   75
         Top             =   990
         Width           =   4080
         Begin MSComctlLib.TabStrip tabRefAppDt 
            Height          =   300
            Left            =   60
            TabIndex        =   76
            Top             =   75
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   529
            Style           =   2
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
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
         Begin VB.Shape Shape2 
            BorderColor     =   &H0085A3A3&
            BorderWidth     =   2
            Height          =   345
            Left            =   45
            Shape           =   4  '둥근 사각형
            Top             =   60
            Width           =   4020
         End
      End
      Begin VB.CommandButton cmdRefer 
         BackColor       =   &H00F4F0F2&
         Caption         =   "참고치 수정(&R)"
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
         Left            =   1920
         Style           =   1  '그래픽
         TabIndex        =   32
         Tag             =   "35206"
         Top             =   3600
         Width           =   2370
      End
      Begin MSComctlLib.ListView lvwReference 
         Height          =   2070
         Left            =   270
         TabIndex        =   33
         Top             =   1410
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3651
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15728382
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Range"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   3450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lis352.frx":0914
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lis352.frx":0A4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lis352.frx":0B84
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label14 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   195
         Left            =   690
         TabIndex        =   72
         Top             =   345
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00DDF0F5&
         FillStyle       =   0  '단색
         Height          =   390
         Index           =   1
         Left            =   285
         Shape           =   4  '둥근 사각형
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblItemSpec 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Ca - Whole Blood"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   285
         TabIndex        =   35
         Top             =   720
         Width           =   4005
      End
   End
   Begin VB.Frame fraSpcList 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   195
      TabIndex        =   29
      Tag             =   "35209"
      Top             =   930
      Width           =   4590
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&S)"
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
         Left            =   3045
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "121"
         Top             =   210
         Width           =   1320
      End
      Begin FPSpread.vaSpread tblSpcList 
         Height          =   2520
         Left            =   300
         TabIndex        =   1
         Tag             =   "35220"
         Top             =   780
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
         _ExtentY        =   4445
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   5
         MaxRows         =   7
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis352.frx":0CBC
         VirtualRows     =   7
      End
      Begin VB.Label Label22 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체리스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   195
         Left            =   540
         TabIndex        =   71
         Top             =   360
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00DDF0F5&
         FillStyle       =   0  '단색
         Height          =   390
         Index           =   0
         Left            =   270
         Shape           =   4  '둥근 사각형
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.TextBox txtTestCd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1590
      TabIndex        =   0
      Top             =   480
      Width           =   1350
   End
   Begin VB.Frame fraDetail 
      BackColor       =   &H00DBE6E6&
      Caption         =   "상세 정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7770
      Left            =   4875
      TabIndex        =   22
      Tag             =   "35207"
      Top             =   930
      Width           =   6000
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   405
         Left            =   435
         TabIndex        =   73
         Top             =   1620
         Width           =   5400
         Begin MSComctlLib.TabStrip tabAppDt 
            Height          =   300
            Left            =   60
            TabIndex        =   74
            Top             =   75
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   529
            Style           =   2
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H0085A3A3&
            BorderWidth     =   2
            Height          =   345
            Left            =   45
            Top             =   60
            Width           =   5115
         End
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0CFC2&
         Caption         =   "삭제(&D)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   330
         Style           =   1  '그래픽
         TabIndex        =   66
         Tag             =   "35301"
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Frame fraInform 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   5490
         Left            =   105
         TabIndex        =   38
         Top             =   2160
         Width           =   5490
         Begin VB.TextBox txtTatStat1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4530
            TabIndex        =   85
            Top             =   2130
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtTatNormal1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4530
            TabIndex        =   84
            Top             =   1755
            Width           =   945
         End
         Begin VB.TextBox txtOutStat1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4530
            TabIndex        =   83
            Top             =   2490
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtOutStat 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3525
            TabIndex        =   81
            Top             =   2490
            Width           =   945
         End
         Begin VB.TextBox txtArletToVal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3900
            TabIndex        =   79
            Top             =   4590
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtArletFrVal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1845
            TabIndex        =   78
            Top             =   4590
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CheckBox chkArletFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Arlet   Check"
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
            Left            =   240
            TabIndex        =   77
            Tag             =   "35202"
            Top             =   4605
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtTatNormal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3525
            TabIndex        =   68
            Top             =   1755
            Width           =   945
         End
         Begin VB.TextBox txtTatStat 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3525
            TabIndex        =   67
            Top             =   2130
            Width           =   945
         End
         Begin VB.TextBox txtDeltaVal2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3900
            TabIndex        =   19
            Top             =   3870
            Width           =   1305
         End
         Begin VB.TextBox txtLabelCnt 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1845
            TabIndex        =   20
            Text            =   "1"
            Top             =   4950
            Width           =   690
         End
         Begin VB.CheckBox chkRndFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "아침 채혈 여부"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   12
            Tag             =   "35203"
            Top             =   3420
            Width           =   1695
         End
         Begin VB.CheckBox chkStatFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "응급 검사 여부"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2250
            TabIndex        =   13
            Tag             =   "35204"
            Top             =   3435
            Width           =   1755
         End
         Begin VB.CheckBox chkPanicFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Panic Check"
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
            Left            =   240
            TabIndex        =   14
            Tag             =   "35202"
            Top             =   4245
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtPanicFrVal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1845
            TabIndex        =   15
            Top             =   4230
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CheckBox chkDeltaFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Delta Check"
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
            Left            =   240
            TabIndex        =   17
            Tag             =   "35201"
            Top             =   3900
            Width           =   1320
         End
         Begin VB.TextBox txtDeltaVal1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1845
            TabIndex        =   18
            Top             =   3870
            Width           =   1305
         End
         Begin VB.TextBox txtPanicToVal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3900
            TabIndex        =   16
            Top             =   4230
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtAvalVal 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1215
            TabIndex        =   5
            Text            =   "0"
            Top             =   810
            Width           =   525
         End
         Begin VB.ComboBox cboStoreCd 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "Lis352.frx":1146
            Left            =   1215
            List            =   "Lis352.frx":1150
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   1200
            Width           =   1770
         End
         Begin VB.ComboBox cboRstUnit 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "Lis352.frx":115A
            Left            =   1215
            List            =   "Lis352.frx":11AF
            TabIndex        =   4
            Top             =   435
            Width           =   1785
         End
         Begin VB.TextBox txtTatAvg 
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3840
            TabIndex        =   11
            Top             =   5310
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtTestCost 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   9
            Top             =   2505
            Width           =   1095
         End
         Begin VB.TextBox txtSpcUnit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   8
            Top             =   2130
            Width           =   1095
         End
         Begin VB.TextBox txtSpcQty 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   7
            Top             =   1755
            Width           =   1080
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   330
            Left            =   1740
            TabIndex        =   39
            Top             =   810
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "txtAvalVal"
            BuddyDispid     =   196644
            OrigLeft        =   2760
            OrigTop         =   1080
            OrigRight       =   3000
            OrigBottom      =   1410
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtSpcAppDt 
            Height          =   330
            Left            =   1215
            TabIndex        =   2
            Top             =   15
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   582
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
            Format          =   85131267
            CurrentDate     =   36328
         End
         Begin MSComCtl2.DTPicker txtSpcExpDt 
            Height          =   330
            Left            =   3750
            TabIndex        =   3
            Top             =   15
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
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
            CheckBox        =   -1  'True
            CustomFormat    =   "yyy-MM-dd"
            DateIsNull      =   -1  'True
            Format          =   85131267
            CurrentDate     =   36328
         End
         Begin MedControls1.LisLabel lblSpcGroup 
            Height          =   330
            Left            =   3900
            TabIndex        =   21
            Top             =   4950
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            BackColor       =   16576489
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   330
            Left            =   2550
            TabIndex        =   52
            Top             =   4950
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "txtLabelCnt"
            BuddyDispid     =   196636
            OrigLeft        =   3840
            OrigTop         =   330
            OrigRight       =   4080
            OrigBottom      =   645
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   330
            Left            =   2055
            TabIndex        =   56
            Top             =   2850
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "txtSeq"
            BuddyDispid     =   196651
            OrigLeft        =   3840
            OrigTop         =   330
            OrigRight       =   4080
            OrigBottom      =   645
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   10
            Text            =   "1"
            Top             =   2850
            Width           =   855
         End
         Begin VB.Label Label20 
            BackStyle       =   0  '투명
            Caption         =   "근무외시간"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   87
            Tag             =   "35116"
            Top             =   1530
            Width           =   900
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label19 
            BackStyle       =   0  '투명
            Caption         =   "근무내시간"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3540
            TabIndex        =   86
            Tag             =   "35116"
            Top             =   1530
            Width           =   900
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "외래"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2880
            TabIndex        =   82
            Tag             =   "35116"
            Top             =   2565
            Width           =   360
         End
         Begin VB.Label Label17 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   270
            Left            =   3300
            TabIndex        =   80
            Top             =   4635
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "일반"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2895
            TabIndex        =   70
            Tag             =   "35116"
            Top             =   1845
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "응급"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2880
            TabIndex        =   69
            Tag             =   "35116"
            Top             =   2205
            Width           =   360
         End
         Begin VB.Label Label3 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "%"
            Height          =   270
            Left            =   3105
            TabIndex        =   62
            Top             =   3915
            Width           =   285
         End
         Begin VB.Label Label11 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "(+)"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   3585
            TabIndex        =   61
            Top             =   3945
            Width           =   285
         End
         Begin VB.Label Label10 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "(-)"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   1560
            TabIndex        =   60
            Top             =   3945
            Width           =   285
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(9 :적용안함)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1995
            TabIndex        =   58
            Tag             =   "35210"
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "우선순위"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   57
            Tag             =   "35211"
            Top             =   2940
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검체군"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3315
            TabIndex        =   55
            Tag             =   "35126"
            Top             =   5040
            Width           =   540
         End
         Begin VB.Label lblCnt2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Barcode 출력장수"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   315
            TabIndex        =   54
            Tag             =   "35211"
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2835
            TabIndex        =   53
            Tag             =   "35211"
            Top             =   5040
            Width           =   180
         End
         Begin VB.Label lblUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "결과단위"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   50
            Tag             =   "35219"
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblDecPnt 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "유효숫자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   49
            Tag             =   "35213"
            Top             =   870
            Width           =   720
         End
         Begin VB.Label Label15 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "%"
            Height          =   270
            Left            =   5115
            TabIndex        =   48
            Top             =   3915
            Width           =   285
         End
         Begin VB.Label Label16 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            Height          =   270
            Left            =   3300
            TabIndex        =   47
            Top             =   4275
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAppDt 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "적용일"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   46
            Tag             =   "35210"
            Top             =   90
            Width           =   540
         End
         Begin VB.Label lblExpDt 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "폐기일"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3075
            TabIndex        =   45
            Tag             =   "35214"
            Top             =   105
            Width           =   540
         End
         Begin VB.Label lblMethod 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "보관방법"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   44
            Tag             =   "35216"
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "검사소요시간  (H:시간,M:분)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3105
            TabIndex        =   43
            Tag             =   "35116"
            Top             =   1290
            Width           =   2400
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "수가"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   42
            Tag             =   "35116"
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "단위"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   41
            Tag             =   "35116"
            Top             =   2220
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검체량"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   285
            TabIndex        =   40
            Tag             =   "35116"
            Top             =   1845
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소(&U)"
         Enabled         =   0   'False
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
         Left            =   4290
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   37
         Tag             =   "35106"
         Top             =   1050
         Width           =   1320
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "수정(&E)"
         Enabled         =   0   'False
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
         Left            =   2970
         Style           =   1  '그래픽
         TabIndex        =   36
         Tag             =   "35205"
         Top             =   1050
         Width           =   1320
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&A)"
         Enabled         =   0   'False
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
         Left            =   1650
         Style           =   1  '그래픽
         TabIndex        =   25
         Tag             =   "35205"
         Top             =   1050
         Width           =   1320
      End
      Begin VB.TextBox txtSpcName 
         BackColor       =   &H00D1D8D3&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   615
         Width           =   3645
      End
      Begin VB.TextBox txtSpcCd 
         BackColor       =   &H00D1D8D3&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   615
         Width           =   1590
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   435
         X2              =   5520
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   420
         X2              =   5520
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Label lblSpcNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 체  명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1980
         TabIndex        =   27
         Tag             =   "35218"
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lblSpcCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   345
         TabIndex        =   26
         Tag             =   "35217"
         Top             =   405
         Width           =   915
      End
   End
   Begin VB.Label lblItemCd 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사코드"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   28
      Tag             =   "35215"
      Top             =   615
      Width           =   855
   End
End
Attribute VB_Name = "frm352Specimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents objCodeList As clsCodeList
Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private MyItem As New clsItem
Private MySqlStmt As New clsLISSqlStatement    ' SQL 클래스
Private MySpecimens As New clsSpecimens  ' 검체 클래스
Private MySpc As New clsSpecimen

Private InsertFlag As Integer
Private UpdateFlag As Integer
Private SvApplyDt As String
Private Const Indicator = "▶"

Private Sub cmdAdd_Click()

    Dim tmpSql As String

    With tblSpcList
        .Col = 1
        .Row = -1
        .Action = ActionClearText
    
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows   ' .DataRowCnt + 1
        .Value = Indicator
        .ForeColor = &HFF&
    
        .Col = 2
        .Value = .Row
    End With

    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Specimen Code List.."
        .Tag = "Specimen"
        .Delimiter = ";"
        .ColumnHeaderText = "검체코드;검체명"
        tmpSql = MySqlStmt.SqlLAB032CodeList(LC3_Specimen, "cdval1, field3")
        .LoadPopUp tmpSql ', 3650, 7350)
        
        If .SelectedString = "" Then
            tblSpcList.MaxRows = tblSpcList.DataRowCnt - 1
            Call tblSpcList_Click(1, 1)
            Me.Enabled = True
        End If
        
    End With


End Sub

Private Sub cmdDelete_Click()

    Dim Resp As VbMsgBoxResult
    
    'If tabAppDt.Pages.Count <= 0 Then Exit Sub
    If tabAppDt.Tabs.Count <= 0 Then Exit Sub
    
    Resp = MsgBox("해당 적용일의 데이타를 모두 삭제하시겠습니까?", vbQuestion, "지정검체 등록")
    If Resp = vbNo Then Exit Sub
    
    Dim MySpecimen As New clsSpecimen

    With MySpecimen
        Call Lab004Move(MySpecimen)
        .SpcDelete
        MySpecimens.Remove Format(txtSpcAppDt.Value, CS_DateDbFormat)
    End With
    Set MySpecimen = Nothing

    
    Call txtTestCd_KeyPress(vbKeyReturn)

End Sub

Private Sub cmdFind_Click(Index As Integer)

    Dim i As Integer
    
    If txtTestCd.Text = "" Then Exit Sub
    If Not ConfirmExit Then Exit Sub

'    I = medListFind(lstItemList, txtTestCd.Text)
    If Not lstItemList.Exists(txtTestCd.Text) Then Exit Sub
    Call lstItemList.KeyChange(txtTestCd.Text)

'    If I < 0 Then Exit Sub
    Select Case Index
        Case 0:   'Previous
            'If I <= 0 Then Exit Sub
            'txtTestCd.Text = lstItemList.List(I - 1)
            lstItemList.MovePrevious
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
        Case 1:   'Next
'            If I >= lstItemList.ListCount - 1 Then Exit Sub
'            txtTestCd.Text = lstItemList.List(I + 1)
            lstItemList.MoveNext
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
    End Select
    Call txtTestCd_KeyPress(vbKeyReturn)

End Sub

Private Sub cmdCancel_Click()

    Call CancelRoutine
    
    'If tabAppDt.Pages.Count > 0 Then
    If tabAppDt.Tabs.Count > 0 Then
        tabAppDt.Tabs(1).Selected = True
        'tabAppDt.Value = 0: Call tabAppDt_Click(0)
    Else
        Call txtTestCd_KeyPress(vbKeyReturn)
    End If

End Sub

Private Sub CancelRoutine()

    If Not ConfirmExit Then Exit Sub

    InsertFlag = 0
    UpdateFlag = 0

    Call LockRtn(1, True)

    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdNew.Caption = "추가"
    cmdEdit.Caption = "수정"
    cmdCancel.Enabled = False
    cmdAdd.Enabled = True

End Sub

Private Sub cmdClear_Click()

    If Not ConfirmExit Then Exit Sub
    
    Call ClearRtn(2)
    cmdAdd.Enabled = True
    txtTestCd.Text = ""
    txtTestCd.SetFocus

End Sub

Private Sub cmdEdit_Click()

    Dim MySpecimen As New clsSpecimen

    If UpdateFlag = 1 Then  ' Update

        If P_CheckSugaCode Then
            If MySpc.CheckCostCd(txtTestCost.Text) = False Then
                MsgBox "수가코드를 다시 입력하세요.", vbCritical, "Messgae"
                txtTestCost.SetFocus
                Exit Sub
            End If
        End If
        
        cmdEdit.Caption = "수정"
        With MySpecimen
            Call Lab004Move(MySpecimen)
            .SpcUpdate
            MySpecimens.Update Format(txtSpcAppDt.Value, CS_DateDbFormat), MySpecimen
        End With
        Set MySpecimen = Nothing
        UpdateFlag = 0
        Call LockRtn(1, True)
        cmdNew.Enabled = True
        cmdCancel.Enabled = False
        cmdAdd.Enabled = True

    Else    ' Edit

        txtSpcAppDt.Enabled = False
        cmdEdit.Caption = "저장"
        UpdateFlag = 1
        Call LockRtn(2, False)
        cmdNew.Enabled = False
        cmdCancel.Enabled = True
        cmdAdd.Enabled = False
        cboRstUnit.SetFocus

    End If
    
    

End Sub

Private Sub cmdExit_Click()

    If Not ConfirmExit Then Exit Sub

    Unload Me
End Sub

Private Sub cmdNew_Click()

    Dim MySpecimen As New clsSpecimen

    If InsertFlag = 1 Then  ' Insert

        If SvApplyDt <> "" And SvApplyDt >= Format(txtSpcAppDt.Value, CS_DateDbFormat) Then
            MsgBox "적용일을 수정하세요.."
            txtSpcAppDt.SetFocus
            Exit Sub
        End If

        If P_CheckSugaCode Then
            If MySpc.CheckCostCd(txtTestCost.Text) = False Then
               MsgBox "수가코드를 다시 입력하세요.", vbCritical, "Messgae"
               txtTestCost.SetFocus
               Exit Sub
            End If
        End If
        
        cmdNew.Caption = "추가"

        With MySpecimen
             Call Lab004Move(MySpecimen)
    '         Set .DbConn = DbConn
             .SpcInsert
             MySpecimens.Add Format(txtSpcAppDt.Value, CS_DateDbFormat), MySpecimen
        End With

        
        Set MySpecimen = Nothing
        InsertFlag = 0
        Call tblSpcList_Click(1, tblSpcList.Row)
        Call LockRtn(1, True)
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        cmdAdd.Enabled = True
        SvApplyDt = ""
        'txtSpcAppDt.SetFocus
        
        

    Else    ' New

        cmdNew.Caption = "저장"
        InsertFlag = 1
        Call ClearRtn(1)
        Call LockRtn(1, False)
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdAdd.Enabled = False
        'If tabAppDt.Pages.Count > 0 Then
        If tabAppDt.Tabs.Count > 0 Then
            SvApplyDt = Format(txtSpcAppDt.Value, CS_DateDbFormat)
        Else
            SvApplyDt = ""
        End If

        Me.Enabled = True
        fraInform.Enabled = True
        txtSpcAppDt.Enabled = True
        txtSpcAppDt.Value = Format(Now, CS_DateLongFormat)
        txtSpcAppDt.SetFocus
    End If

End Sub


Private Sub cmdPopupList_Click()

    Dim tmpSql As String

    If Not ConfirmExit Then Exit Sub

    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Test Code List.."
        .Tag = "TestCode"
        .ColumnHeaderText = "검사코드;검사명"
        tmpSql = MySqlStmt.SqlLAB001CodeList
        .LoadPopUp tmpSql '(, Me.Top + txtTestCd.Top + txtTestCd.Height, Me.Left + txtTestCd.Left, lstItemList)
        'Call .ListPop(tmpSql, Me.Top + txtTestCd.Top + txtTestCd.Height, Me.Left + txtTestCd.Left)
    End With

End Sub

Private Sub cmdRefer_Click()

    With frm353Reference
        .Show
        Call SetParent(.hWnd, gParentWhnd)
        .txtTestCd.Text = txtTestCd.Text
        Call .Raise_TestCd_Keypress
        .cboSpcCd.ListIndex = medComboFind(.cboSpcCd, txtSpcCd.Text)
        .ZOrder 0
    End With

End Sub



Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    Me.WindowState = 2

End Sub

Private Sub Form_Deactivate()
    Set objCodeList = Nothing

End Sub

Private Sub Form_Load()

    txtSpcAppDt.Value = Format(Now, CS_DateLongFormat)
    txtSpcExpDt.Value = ""
    cmdAdd.Enabled = False

    Call MySpc.GetStoreCd(cboStoreCd)
'    Call MySpc.GetBuildings(lstStatFg)
'    Call MySpc.GetBuildings(lstTestFg)
    Call MyItem.GetItemList(lstItemList): DoEvents

    Call LockRtn(5, True)
    InsertFlag = 0
    UpdateFlag = 0
    
'    tabAppDt.Pages.Clear
'    tabRefAppDt.Pages.Clear
    tabAppDt.Tabs.Clear
    tabRefAppDt.Tabs.Clear
    
    cmdDelete.Enabled = ObjMyUser.isdeveloper 'gIsDeveloper

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set objCodeList = Nothing
    Set MyItem = Nothing
    Set MySqlStmt = Nothing
    Set MySpecimens = Nothing
    Set MySpc = Nothing

End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    Dim i As Long
    Dim tmpTag As String

    tmpTag = objCodeList.Tag
    Set objCodeList = Nothing

    Select Case tmpTag
        '검사항목
        Case "TestCode":
            If pSelectedItem = "" Then GoTo Skip
            txtTestCd.Text = medShift(pSelectedItem, ";")
            lblTestName.Caption = medShift(pSelectedItem, ";")
            Call txtTestCd_KeyPress(vbKeyReturn)
        '검체
        Case "Specimen":
            With tblSpcList
                If pSelectedItem = "" Then
                    .MaxRows = .DataRowCnt - 1
                    Call tblSpcList_Click(1, 1)
                    GoTo Skip
                End If
                .Row = .MaxRows   ' .DataRowCnt + 1
                .TopRow = .MaxRows - 6
                .Col = 3: '기 등록 여부 Check
                            For i = 1 To .MaxRows - 1
                                .Row = i
                                If Trim(.Value) = Trim(medGetP(pSelectedItem, 1, ";")) Then
                                    MsgBox "이미 등록된 검체입니다.."
                                    .MaxRows = .MaxRows - 1
                                    .TopRow = i
                                    Call tblSpcList_Click(1, i)
                                    GoTo Skip
                                End If
                            Next
                            .Row = .MaxRows   ' .DataRowCnt + 1
                            txtSpcCd.Text = Trim(medGetP(pSelectedItem, 1, ";"))
                            .Value = txtSpcCd.Text
                .Col = 4: txtSpcName.Text = Trim(medGetP(pSelectedItem, 2, ";"))
                         .Value = txtSpcName.Text
                InsertFlag = 0
                Call tblSpcList_Click(1, .Row)
                Call cmdNew_Click
                .Col = 2: .Value = .Row
                txtSeq.Text = .Value
            End With
    End Select

Skip:
    Me.Enabled = True
End Sub

'Private Sub objCodeList_SendCode(ByVal SelString As String)
''
'    Dim i As Long
'    Dim tmpTag As String
'
'    tmpTag = objCodeList.Tag
'    Set objCodeList = Nothing
'
'    Select Case tmpTag
'        '검사항목
'        Case "TestCode":
'            If SelString = "" Then GoTo Skip
'            txtTestCd.Text = medShift(SelString, ";")
'            lblTestName.Caption = medShift(SelString, ";")
'            Call txtTestCd_KeyPress(vbKeyReturn)
'        '검체
'        Case "Specimen":
'            With tblSpcList
'                If SelString = "" Then
'                    .MaxRows = .DataRowCnt - 1
'                    Call tblSpcList_Click(1, 1)
'                    GoTo Skip
'                End If
'                .Row = .MaxRows   ' .DataRowCnt + 1
'                .TopRow = .MaxRows - 6
'                .Col = 3: '기 등록 여부 Check
'                            For i = 1 To .MaxRows - 1
'                                .Row = i
'                                If Trim(.Value) = Trim(medGetP(SelString, 1, ";")) Then
'                                    MsgBox "이미 등록된 검체입니다.."
'                                    .MaxRows = .MaxRows - 1
'                                    .TopRow = i
'                                    Call tblSpcList_Click(1, i)
'                                    GoTo Skip
'                                End If
'                            Next
'                            .Row = .MaxRows   ' .DataRowCnt + 1
'                            txtSpcCd.Text = Trim(medGetP(SelString, 1, ";"))
'                            .Value = txtSpcCd.Text
'                .Col = 4: txtSpcName.Text = Trim(medGetP(SelString, 2, ";"))
'                         .Value = txtSpcName.Text
'                InsertFlag = 0
'                Call tblSpcList_Click(1, .Row)
'                Call cmdNew_Click
'                .Col = 2: .Value = .Row
'                txtSeq.Text = .Value
'            End With
'    End Select
'
'Skip:
'    Me.Enabled = True
'
'End Sub

'% 검체의 적용일을 선택하면 상세정보를 Display한다.
'Private Sub tabAppDt_Click(Index As Long)
Private Sub tabAppDt_Click()

    Dim tmpStr As String
    Dim tmpAppDt As String
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    Dim tmpTATS1 As String
    Dim tmpTATS2 As String
    Dim tmpTATS3 As String
    Dim tmpTATS4 As String
    Dim tmpTATS5 As String
    Dim tmpTATS6 As String
    Dim i As Integer, j As Integer

    Call CancelRoutine
    
    tmpAppDt = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
    With MySpecimens.Specimen(tmpAppDt)
        txtTestCd.Text = .TestCd
        txtSpcCd.Text = .SpcCd
        txtSpcAppDt.Value = Format(.ApplyDt, CS_DateMask)
        'cboSpcGrpCd.ListIndex = medListFind(cboSpcGrpCd, .SpcGrpCd)
        txtLabelCnt.Text = .LabelCnt
        cboRstUnit.Text = .RstUnit
        chkRndFg.Value = Val(.RndFg)
        chkStatFg.Value = Val(.StatFg)
    
    
'        tmpStatFg = medGetP(.StatFlags, 1, ";")
'        For I = 0 To lstStatFg.ListCount - 1
'            If Mid(tmpStatFg, I + 1, 1) = "1" Then
'                lstStatFg.Selected(I) = True
'            Else
'                lstStatFg.Selected(I) = False
'            End If
'        Next
'        tmpTestFg = medGetP(.StatFlags, 2, ";")
'        For I = 0 To lstTestFg.ListCount - 1
'            If Mid(tmpTestFg, I + 1, 1) = "1" Then
'                lstTestFg.Selected(I) = True
'            Else
'                lstTestFg.Selected(I) = False
'            End If
'        Next
    
        txtAvalVal.Text = .AvalVal
        chkPanicFg.Value = Val(.PanicFg)
        txtPanicFrVal.Text = .PanicFrVal
        txtPanicToVal.Text = .PanicToVal
        
        chkArletFg.Value = Val(.ArletFg)
        txtArletFrVal.Text = .ArletFrVal
        txtArletToVal.Text = .ArletToVal
        
        chkDeltaFg.Value = Val(.DeltaFg)
        txtDeltaVal1.Text = .DeltaVal1
        txtDeltaVal2.Text = .DeltaVal2
        txtTestCost.Text = .TestCost
        txtSeq.Text = .Seq
        cboStoreCd.ListIndex = medComboFind(cboStoreCd, .StoreCd)
        txtTatAvg.Text = .TatAvg
        txtSpcQty.Text = .SpcQty
        txtSpcUnit.Text = .SpcUnit
        If Trim(.ExpDt) = "" Then
            txtSpcExpDt.Value = ""
        Else
            txtSpcExpDt.Value = Format(.ExpDt, CS_DateMask)
        End If
        tmpTATS1 = medGetP(.TATS, 1, ";")
        tmpTATS2 = medGetP(.TATS, 2, ";")
        tmpTATS3 = medGetP(.TATS, 3, ";")
        tmpTATS4 = medGetP(.TATS, 4, ";")
'        tmpTATS5 = medGetP(.TATS, 5, ";")
'        tmpTATS6 = medGetP(.TATS, 6, ";")
    End With
    
    txtTatNormal.Text = tmpTATS1
    txtTatNormal1.Text = tmpTATS4
    txtTatStat.Text = tmpTATS2
'    txtTatStat1.Text = tmpTATS5
    txtOutStat.Text = tmpTATS3
'    txtOutStat1.Text = tmpTATS6
    
'    With tblTAT
'        .Row = 1: .Row2 = .MaxRows
'        .Col = 1: .Col2 = .MaxCols
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'        For I = 1 To .MaxRows
'            .Row = I
'            .Col = 1: .Value = medGetP(tmpTATS1, I, ":")
'            .Col = 2: .Value = medGetP(tmpTATS2, I, ":")
'        Next
'    End With
'    tabFlags.Tabs(1).Selected = True
    
    Call LockRtn(1, True)
    
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False

End Sub

'Private Sub tabFlags_Click()
'    'If Not fraInform.Enabled Then Exit Sub
'    If tabFlags.SelectedItem.Index = 1 Then
'        lstStatFg.Visible = True
'        lstTestFg.Visible = False
'        'lstStatFg.ZOrder 0
'        'lstStatFg.SetFocus
'    Else
'        lstStatFg.Visible = False
'        lstTestFg.Visible = True
'        'lstTestFg.ZOrder 0
'        'lstTestFg.SetFocus
'    End If
'End Sub

'Private Sub tabRefAppDt_Click(Index As Long)
Private Sub tabRefAppDt_Click()

    Dim tmpSql As String
    Dim tmpAppDt As String
    Dim tmpSpcCd As String
    
    'lstSpcName.ListIndex = cboSpcCd.ListIndex
    tmpAppDt = Format(tabRefAppDt.SelectedItem.Caption, CS_DateDbFormat)
    tmpSpcCd = txtSpcCd.Text
    tmpSql = MySqlStmt.SqlLAB005Read(txtTestCd.Text, tmpSpcCd, tmpAppDt)
    Call Lab005Show(tmpSql)

End Sub

'% 검체리스트에서 한 검체를 선택(클릭)하면 상세 정보 Display

Private Sub tblSpcList_Click(ByVal Col As Long, ByVal Row As Long)

    Static OldRow As Long
    Dim SqlStmt As String
    Dim tmpSeq As String

    If Row < 1 Then Exit Sub
    If InsertFlag = 1 Then Exit Sub

    With tblSpcList
    
        .Col = 1
        If OldRow > 0 Then
            .Row = OldRow
            .Value = ""
        End If
    
        .Row = Row
        .Value = Indicator
        .ForeColor = &HFF&
    
        .Col = 2: txtSeq.Text = .Value
        .Col = 3: txtSpcCd.Text = .Value
        .Col = 4: txtSpcName.Text = .Value
        .Col = 5: lblSpcGroup.Caption = .Value
    
        OldRow = .Row
    
    End With


    lblItemSpec.Caption = lblTestName.Caption & " - " & txtSpcName.Text
    SqlStmt = MySqlStmt.SqlLAB004Read(txtTestCd.Text, txtSpcCd.Text, txtSeq.Text)
    Call Lab004Load(SqlStmt)
    Call Lab004Show
    Call Lab005Load(SqlStmt)
    Call LockRtn(1, True)

    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False

End Sub

Private Sub txtSpcAppDt_LostFocus()

    Dim intNew As Integer


End Sub


'% 검사코드 입력후 검체정보 Display

Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objCodeList Is Nothing Then Call cmdPopupList_Click
'        Call objCodeList.SetFocus(2)
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim tmpSql As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If Not ConfirmExit Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text = "" Then Exit Sub
        Call ClearRtn(2)
    
        If lstItemList.Exists(Trim(txtTestCd.Text)) Then
            lstItemList.KeyChange (Trim(txtTestCd.Text))
            lblTestName.Caption = lstItemList.Fields("testnm")
            cmdAdd.Enabled = True
            Call LockRtn(5, False)
        Else
            Call LockRtn(5, True)
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        tmpSql = MySqlStmt.SqlSpecimenRead(txtTestCd.Text)
        Call LabSpecimenLoad(tmpSql)
        If tblSpcList.MaxRows > 0 Then
            Call tblSpcList_Click(1, 1)
        End If
    End If
    
End Sub


'% Sub Routine 3 : LabSpecimenLoad
'%                 지정검체명들을 Tab에 Display

Private Sub LabSpecimenLoad(ByVal SqlStmt As String)

    Dim objRs As Recordset
    Dim i As Integer

    Set objRs = New Recordset  'Sql 실행
    objRs.Open SqlStmt, DBConn
    
    With tblSpcList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 0
    
        .Row = 0
        While (objRs.EOF = False)
            If .Row = .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .Row + 1
            .Col = 2: .Value = Trim("" & objRs.Fields("Seq").Value)
            .Col = 3: .Value = Trim("" & objRs.Fields("SpcCd").Value)
            .Col = 4: .Value = Trim("" & objRs.Fields("SpcNm").Value)
            .Col = 5: .Value = Trim("" & objRs.Fields("SpcGrp").Value)
            objRs.MoveNext
        Wend
        .RowHeight(-1) = 13.3
    End With

    Set objRs = Nothing

End Sub


'% Sub Routine 4 : Lab004Load
'%                        Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                        클래스 clsItem의 Data Attribute에 저장한다.

Private Sub Lab004Load(ByVal SqlStmt As String)

    Dim MySpecimen As clsSpecimen
    Dim objRs As Recordset
    Dim tmpTATS As String

    On Error GoTo Error_Trap

    Set MySpecimen = New clsSpecimen
    Set objRs = New Recordset   'Sql 실행
    objRs.Open SqlStmt, DBConn
    
    MySpecimens.Clear
    With MySpecimen
        While (objRs.EOF = False)
            .TestCd = "" & objRs.Fields("TestCd").Value
            .SpcCd = "" & objRs.Fields("SpcCd").Value
            .Seq = Val("" & objRs.Fields("Seq").Value)
            .ApplyDt = "" & objRs.Fields("ApplyDt").Value
            .SpcGrpCd = "" & objRs.Fields("SpcGrpCd").Value
            .RstUnit = "" & objRs.Fields("RstUnit").Value
            .RndFg = "" & objRs.Fields("RndFg").Value
            .StatFg = "" & objRs.Fields("StatFg").Value
            .StatFlags = "" & objRs.Fields("StatFlags").Value
            .AvalVal = Val("" & objRs.Fields("AvalVal").Value)
            .LabelCnt = Val("" & objRs.Fields("LabelCnt").Value)
            .PanicFg = "" & objRs.Fields("PanicFg").Value
            .PanicFrVal = Val("" & objRs.Fields("PanicFrVal").Value)
            .PanicToVal = Val("" & objRs.Fields("PanicToVal").Value)
            .DeltaFg = "" & objRs.Fields("DeltaFg").Value
            .DeltaVal1 = Val("" & objRs.Fields("DeltaVal").Value)
            .DeltaVal2 = Val("" & objRs.Fields("DeltaVal2").Value)
            .TestCost = "" & objRs.Fields("TestCost").Value
            .StoreCd = "" & objRs.Fields("StoreCd").Value
            .TatAvg = Val("" & objRs.Fields("TatAvg").Value)
            .SpcQty = Val("" & objRs.Fields("SpcQty").Value)
            .SpcUnit = "" & objRs.Fields("SpcUnit").Value
            .ExpDt = "" & objRs.Fields("ExpDt").Value
            .TATS = "" & objRs.Fields("tats").Value
'            .ArletFg = "" & objRs.Fields("arletfg").Value
'            .ArletFrVal = Val("" & objRs.Fields("arletfrval").Value)
'            .ArletToVal = Val("" & objRs.Fields("arlettoval").Value)
            MySpecimens.Add objRs.Fields("ApplyDt").Value, MySpecimen
            objRs.MoveNext
        Wend
    End With

    Set objRs = Nothing
    Exit Sub

Error_Trap:
    If Err.Number <> 94 Then
        MsgBox Err.Number & "  " & Err.Description
        Set objRs = Nothing
        Exit Sub
    Else
        Resume Next
    End If

End Sub


'% 검체 마스터의 검체(SpcCd)와 적용일(ApplyDt)을 Tab에 Display

Private Sub Lab004Show()

    Dim i As Integer
    Dim strTmp As String

'    tabAppDt.Pages.Clear
    tabAppDt.Tabs.Clear
    If MySpecimens.Count = 0 Then Exit Sub
    For i = 1 To MySpecimens.Count
        strTmp = Format(MySpecimens.Specimen(i).ApplyDt, CS_DateMask)
        'tabAppDt.Pages.Add MySpecimens.Specimen(i).ApplyDt, strTmp, i - 1
        tabAppDt.Tabs.Add i, , strTmp
    Next
   If tabAppDt.Tabs.Count > 0 Then tabAppDt.Tabs(1).Selected = True
    'If tabAppDt.Pages.Count > 0 Then tabAppDt.Value = 0: Call tabAppDt_Click(0)

End Sub


Private Sub Lab004Move(ByRef MySpecimen As clsSpecimen)

    Dim i As Integer
    Dim tmpTATS1 As String
    Dim tmpTATS2 As String

    With MySpecimen
        .TestCd = txtTestCd.Text
        .SpcCd = txtSpcCd.Text
        .Seq = Val(txtSeq.Text)   ' tblSpcList.Row
        .ApplyDt = Format(txtSpcAppDt.Value, CS_DateDbFormat)
        .LabelCnt = Val(txtLabelCnt.Text)
        .SpcGrpCd = ""
        .RstUnit = cboRstUnit.Text
        .RndFg = chkRndFg.Value
        .StatFg = chkStatFg.Value
        .StatFlags = chkStatFg.Value & ";" & "1"
        
'        .StatFlags = ""
'        For I = 0 To lstStatFg.ListCount - 1
'            If lstStatFg.Selected(I) Then
'                .StatFlags = .StatFlags & "1"
'            Else
'                .StatFlags = .StatFlags & "0"
'            End If
'        Next
'        .StatFlags = .StatFlags & ";"
'        For I = 0 To lstTestFg.ListCount - 1
'            If lstTestFg.Selected(I) Then
'               .StatFlags = .StatFlags & "1"
'            Else
'               .StatFlags = .StatFlags & "0"
'            End If
'        Next
        .AvalVal = Val(txtAvalVal.Text)
        .PanicFg = chkPanicFg.Value
        .PanicFrVal = Val(txtPanicFrVal.Text)
        .PanicToVal = Val(txtPanicToVal.Text)
        
        .ArletFg = chkArletFg.Value
        .ArletFrVal = Val(txtArletFrVal.Text)
        .ArletToVal = Val(txtArletToVal.Text)
        
        .DeltaFg = chkDeltaFg.Value
        .DeltaVal1 = Val(txtDeltaVal1.Text)
        .DeltaVal2 = Val(txtDeltaVal2.Text)
        .TestCost = txtTestCost.Text
        If .TestCost = "" Then .TestCost = .TestCd
        .StoreCd = medGetP(cboStoreCd.Text, 1, " ")
        .TatAvg = Val(txtTatAvg.Text)
        .SpcQty = Val(txtSpcQty.Text)
        .SpcUnit = txtSpcUnit.Text

        If IsNull(txtSpcExpDt.Value) Then
            .ExpDt = ""
        Else
            .ExpDt = Format(txtSpcExpDt.Value, CS_DateDbFormat)
        End If

        .TATS = txtTatNormal.Text & ";" & txtTatStat.Text & ";" & txtOutStat.Text & ";" & txtTatNormal1.Text ' & ";" & txtTatStat1.Text & ";" & txtOutStat1.Text

'        tmpTATS1 = ""
'        tmpTATS2 = ""
'        For I = 1 To tblTAT.MaxRows
'            tblTAT.Row = I
'            tblTAT.Col = 1: tmpTATS1 = tmpTATS1 & tblTAT.Value & ":"
'            tblTAT.Col = 2: tmpTATS2 = tmpTATS2 & tblTAT.Value & ":"
'        Next
'        .TATS = tmpTATS1 & ";" & tmpTATS2
    End With

End Sub

Private Sub ClearRtn(ByVal intPart As Integer)

    Dim i As Integer

    ' intPart : 1-Specimen, 2-All
    Select Case intPart
    Case 0, 2: GoTo Clear0
    Case 1: GoTo Clear1
    End Select
    Exit Sub

Clear0:
      txtSpcCd.Text = ""
      txtSpcName.Text = ""
      tabRefAppDt.Tabs.Clear
      'tabRefAppDt.Pages.Clear
      lvwReference.ListItems.Clear
      lblItemSpec.Caption = ""
      lblTestName.Caption = ""
      tblSpcList.MaxRows = 0

      If intPart <> 2 Then Exit Sub

Clear1:
      'tabAppDt.Pages.Clear
      lblItemSpec.Caption = ""
      txtSpcAppDt.Value = Format(Now, CS_DateLongFormat)
      txtSpcExpDt.Value = ""
      cboRstUnit.Text = ""
      txtSeq.Text = ""
      txtAvalVal.Text = "9"
      cboStoreCd.ListIndex = -1
      txtSpcQty.Text = ""
      txtSpcUnit.Text = ""
      txtTestCost.Text = ""
      txtTatAvg.Text = ""
      chkRndFg.Value = 0
      chkStatFg.Value = 0
      chkPanicFg.Value = 0
      chkDeltaFg.Value = 0
      chkArletFg.Value = 0
      
      txtPanicFrVal.Text = ""
      txtPanicToVal.Text = ""
      
      txtArletFrVal.Text = ""
      txtArletToVal.Text = ""
      
      txtDeltaVal1.Text = ""
      txtDeltaVal2.Text = ""
      txtLabelCnt.Text = 1
'      For I = 0 To lstStatFg.ListCount - 1
'         lstStatFg.Selected(I) = False
'      Next
'      For I = 0 To lstTestFg.ListCount - 1
'         lstTestFg.Selected(I) = False
'      Next
      txtTatNormal.Text = ""
      txtTatStat.Text = ""
      txtOutStat.Text = ""
      
'      tblTAT.Row = 1: tblTAT.Row2 = tblTAT.MaxRows
'      tblTAT.Col = 1: tblTAT.Col2 = tblTAT.MaxCols
'      tblTAT.BlockMode = True
'      tblTAT.Action = ActionClearText
'      tblTAT.BlockMode = False

End Sub


Private Sub LockRtn(ByVal intPart As Integer, ByVal LockValue As Boolean)

    Dim EnableValue As Boolean

    ' intPart : 1-Item, 2-Clinical Notice, 3-Specimen, 4-Reference Range, 5-All
    If LockValue Then
       EnableValue = False
    Else: EnableValue = True
    End If

    If intPart = 1 Then txtSpcAppDt.Enabled = EnableValue
    fraInform.Enabled = EnableValue

    If intPart = 5 Then
        fraSpcList.Enabled = EnableValue
        fraRefRange.Enabled = EnableValue
        fraDetail.Enabled = EnableValue
    End If

End Sub

Public Sub Raise_TestCd_Keypress()
    Call txtTestCd_KeyPress(13)
End Sub

Private Sub Lab005AppDt(ByVal SqlStmt As String)

    Dim objRs As Recordset       'Oracle DynaSet
    Dim i As Integer
    Dim tmpKey As String
    Dim tmpCaption As String

    Set objRs = New Recordset   'Sql 실행
    objRs.Open SqlStmt, DBConn
    
    i = 0
    tabRefAppDt.Tabs.Clear
    'tabRefAppDt.Pages.Clear
    While (objRs.EOF = False)
        i = i + 1
        tmpKey = "" & objRs.Fields("ApplyDt").Value
        tmpCaption = Format(tmpKey, CS_DateMask)
        tabRefAppDt.Tabs.Add i, , tmpCaption
        'tabRefAppDt.Pages.Add , tmpCaption, i - 1
        objRs.MoveNext
    Wend

    Set objRs = Nothing
End Sub


Private Sub Lab005Load(ByVal SqlStmt As String)

    Dim i As Integer
    Dim MyReference As New clsReference


    Dim tmpSql As String
    Dim tmpSpcCd As String

    'tabRefAppDt.Pages.Clear
    tabRefAppDt.Tabs.Clear
    lvwReference.ListItems.Clear

    tmpSql = MySqlStmt.SqlLAB005AppDt(txtTestCd.Text, txtSpcCd.Text)
    Call Lab005AppDt(tmpSql)
    'If tabRefAppDt.Pages.Count > 0 Then
    If tabRefAppDt.Tabs.Count > 0 Then
        Call tabRefAppDt_Click
    End If

End Sub

Private Sub Lab005Show(ByVal SqlStmt As String)
    Dim objRs As Recordset
    Dim intFieldCount As Integer
    Dim aryTitle As Variant
    Dim aryWidth As Variant
    Dim itmx As ListItem
    Dim i As Long

    Set objRs = New Recordset 'Sql 실행
    objRs.Open SqlStmt, DBConn
    
    If objRs.EOF Then GoTo NoData

    i = 0
    intFieldCount = 5
    aryTitle = Array("성별", "일령", "기준치", "Auto Value", "Panic Value")
    aryWidth = Array(-85, -55, 50, -50, 0)
    lvwReference.ColumnHeaders.Clear
    With lvwReference
        For i = 1 To intFieldCount
            .ColumnHeaders.Add i, aryTitle(i - 1), aryTitle(i - 1), (.Width \ intFieldCount) _
                                + aryWidth(i - 1), vbLeftJustify
        Next i
        .View = lvwReport                          ' lvwReport = 3 (Report Style)
    End With
    lvwReference.ListItems.Clear
    lvwReference.SmallIcons = ImageList1
    lvwReference.Icons = ImageList1

    While (objRs.EOF = False)
        i = i + 1
        Select Case "" & objRs.Fields("ApplySex").Value
        Case "M":
            Set itmx = lvwReference.ListItems.Add(, , "남자", 1, 1)
        Case "F":
            Set itmx = lvwReference.ListItems.Add(, , "여자", 2, 2)
        Case "B":
            Set itmx = lvwReference.ListItems.Add(, , "Both")
        Case "U":
            Set itmx = lvwReference.ListItems.Add(, , "Unknown", 3)
        Case "Z":
            Set itmx = lvwReference.ListItems.Add(, , "Both", 3)
        End Select
        itmx.SubItems(1) = "" & objRs.Fields("AgeFrom").Value & " - " & objRs.Fields("AgeTo").Value & " days"
        '선린 : 2001-05-31(추가)
        If Val("" & objRs.Fields("RefValFrom").Value) = 0 And Val("" & objRs.Fields("RefValTo").Value) = 0 Then
            itmx.SubItems(2) = "" & objRs.Fields("RefCd").Value
        Else
            itmx.SubItems(2) = "" & objRs.Fields("RefValFrom").Value & " - " & objRs.Fields("RefValTo").Value
            If Len("" & objRs.Fields("RefCd").Value) Then
                itmx.SubItems(2) = itmx.SubItems(2) & "(" & objRs.Fields("RefCd").Value & ")"
            End If
        End If
        itmx.SubItems(3) = "" & objRs.Fields("aRefValFrom").Value & " - " & objRs.Fields("aRefValTo").Value
        itmx.SubItems(4) = "" & objRs.Fields("panicfrval").Value & " - " & objRs.Fields("panictoval").Value
        objRs.MoveNext
    Wend


NoData:
    Set objRs = Nothing
    Exit Sub

End Sub

Private Function ConfirmExit() As Boolean

    Dim intResp As VbMsgBoxResult

    ConfirmExit = True
    If InsertFlag = 1 Or UpdateFlag = 1 Then
        intResp = MsgBox("변경된 내용을 취소하시겠습니까 ? ", vbYesNo)
        If intResp = vbNo Then
            ConfirmExit = False
            Exit Function
        End If
        InsertFlag = 0
        UpdateFlag = 0
    End If

End Function

Private Sub txtTestCost_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTestCost_LostFocus()
    Dim ValCheck As Boolean

    If ActiveControl.Name = cmdCancel.Name Then Exit Sub

    If P_CheckSugaCode Then
        If txtTestCost.Text = "" Then
            MsgBox "수가코드를 반드시 입력하세요.", vbCritical, "Messgae"
            txtTestCost.SetFocus
            Exit Sub
        End If
        ValCheck = MySpc.CheckCostCd(txtTestCost.Text)
        If Not ValCheck Then
            MsgBox "등록되지 않은 수가코드입니다.. 다시 입력하세요.", vbCritical, "입력Error"
            txtTestCost.SetFocus
            Exit Sub
        End If
    End If
    
End Sub
