VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_Result_View 
   BorderStyle     =   0  '없음
   Caption         =   "검사결과조회"
   ClientHeight    =   8745
   ClientLeft      =   150
   ClientTop       =   1875
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8745
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin RichTextLib.RichTextBox txtRemark 
      Height          =   1185
      Left            =   45
      TabIndex        =   42
      Top             =   2010
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   15463915
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ANATO111.frx":0000
   End
   Begin RichTextLib.RichTextBox txtDiag 
      Height          =   6348
      Left            =   48
      TabIndex        =   41
      Top             =   1992
      Width           =   7932
      _ExtentX        =   13996
      _ExtentY        =   11192
      _Version        =   393217
      BackColor       =   15463915
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ANATO111.frx":0265
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00808000&
      Height          =   7380
      Left            =   80000
      ScaleHeight     =   7320
      ScaleWidth      =   7890
      TabIndex        =   30
      Top             =   435
      Visible         =   0   'False
      Width           =   7950
      Begin Threed.SSCommand SSCommand1 
         Height          =   465
         Left            =   -2775
         TabIndex        =   38
         Top             =   675
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   "취  소"
         ForeColor       =   255
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
      Begin FPSpread.vaSpread ssResult 
         Height          =   6615
         Left            =   200
         TabIndex        =   31
         Top             =   200
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13944
         _ExtentY        =   11668
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   11
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   8421376
         MaxCols         =   12
         MaxRows         =   600
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "ANATO111.frx":04CA
         VisibleCols     =   500
         VisibleRows     =   500
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   1065
      Left            =   60
      ScaleHeight     =   1005
      ScaleWidth      =   7845
      TabIndex        =   15
      Top             =   810
      Width           =   7905
      Begin VB.TextBox lblJDate 
         BackColor       =   &H00EBF5EB&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   3705
         TabIndex        =   37
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "접수일자"
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
         Left            =   2724
         TabIndex        =   36
         Top             =   588
         Width           =   768
      End
      Begin VB.Label lblDrName 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6255
         TabIndex        =   27
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblChiefName 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6255
         TabIndex        =   26
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label lblRoomCode 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3705
         TabIndex        =   25
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   38000
         TabIndex        =   24
         Top             =   630
         Width           =   1300
      End
      Begin VB.Label lblSname 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   23
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label lblPtNo 
         BackColor       =   &H00EBF5EB&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   22
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label10 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "환 자 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   252
         TabIndex        =   21
         Top             =   588
         Width           =   720
      End
      Begin VB.Label Label9 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "보고일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2724
         TabIndex        =   20
         Top             =   168
         Width           =   780
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "의뢰의사"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5340
         TabIndex        =   19
         Top             =   168
         Width           =   732
      End
      Begin VB.Label Label7 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
         Caption         =   "성별/병실"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   26500
         TabIndex        =   18
         Top             =   680
         Width           =   1100
      End
      Begin VB.Label Label5 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "검 사 자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5340
         TabIndex        =   17
         Top             =   588
         Width           =   732
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "등록번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   255
         TabIndex        =   16
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.ListBox lstPtList 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   8070
      TabIndex        =   13
      Top             =   1200
      Width           =   1905
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808000&
      Height          =   3360
      Left            =   10050
      ScaleHeight     =   3300
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   4350
      Width           =   1815
      Begin Threed.SSCommand cmdPrt 
         Height          =   825
         Left            =   0
         TabIndex        =   44
         Top             =   1650
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "출력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         AutoSize        =   1
         Picture         =   "ANATO111.frx":0BCF
      End
      Begin Threed.SSCommand cmdRemark 
         Height          =   825
         Left            =   0
         TabIndex        =   28
         Top             =   825
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "병력사항"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO111.frx":0EE9
      End
      Begin Threed.SSCommand cmdView 
         Height          =   825
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "조 회"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO111.frx":133B
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   825
         Left            =   0
         TabIndex        =   5
         Top             =   2475
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO111.frx":178D
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   3510
      Left            =   10050
      ScaleHeight     =   3450
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   780
      Width           =   1815
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   315
         Left            =   360
         TabIndex        =   40
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36313
      End
      Begin MSComCtl2.DTPicker dtFromDate 
         Height          =   315
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36313
      End
      Begin VB.OptionButton optSName 
         Caption         =   "환 자 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   2730
         Width           =   1335
      End
      Begin VB.OptionButton optPtNo 
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1995
         Width           =   1335
      End
      Begin VB.OptionButton optAnatNo 
         Caption         =   "병리번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   1275
         Width           =   1335
      End
      Begin VB.OptionButton optFromTo 
         Caption         =   "보고일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   345
         Width           =   1335
      End
      Begin VB.TextBox txtSeqnum 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1590
         Width           =   600
      End
      Begin VB.TextBox txtDateYY 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   500
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox txtClass 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "P"
         Top             =   1590
         Width           =   350
      End
      Begin VB.TextBox txtSName 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   10
         TabIndex        =   29
         ToolTipText     =   "한글두자리이상을 입력하십시요."
         Top             =   3045
         Width           =   1500
      End
      Begin VB.TextBox txtPtNo 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         MaxLength       =   8
         TabIndex        =   3
         Top             =   2310
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "조  건"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1740
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12045
      _Version        =   65536
      _ExtentX        =   21246
      _ExtentY        =   1244
      _StockProps     =   15
      Caption         =   "ANATOMIC   PATHOLOGY"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.76
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   9660
         ScaleHeight     =   315
         ScaleWidth      =   2115
         TabIndex        =   7
         Top             =   150
         Width           =   2175
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   945
            TabIndex        =   9
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "User:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   180
            TabIndex        =   8
            Top             =   60
            Width           =   450
         End
      End
   End
   Begin VB.Label lblRownum 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   8100
      TabIndex        =   43
      Top             =   7800
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00800000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "병 리 번 호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8070
      TabIndex        =   14
      Top             =   795
      Width           =   1905
   End
End
Attribute VB_Name = "Anato_Result_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
    Dim LbRemark            As Boolean
    Dim TabCheck            As Integer
     
    Dim i
    Dim LsClass             As String * 2
    Dim LsDateYY            As String * 4
    Dim LsSeqNum            As String * 5
    Dim LsDeptCode          As String
    
    Dim LSROWNUM            As Integer

    Dim LsDiag             As String

    Dim SLIDENO
    Dim LsElectro
    Dim LsFlow
    

Private Sub Txt_Clear()

'    txtFromDate = ""
'    txtToDate = ""
    
    txtClass.Text = ""
    txtDateYY.Text = ""
    txtSeqnum.Text = ""
    txtPtNo.Text = ""
    txtSName.Text = ""
    
End Sub


Private Sub cmdExit_Click()

    If GsGoFlag = "DIAG" Then
        Unload Me
        Anato_Result.Show
    Else
        Unload Me
    End If
    
    GsGoFlag = ""
    
End Sub

Private Sub cmdPrt_Click()
    
    Dim LsJNum
    
    Dim LsPtNo              As String * 8
    Dim LsSname             As String * 10
    
    Dim LsChiefName         As String * 10
    Dim LsDrName            As String * 8
    
    Dim LsJDate             As String * 10
    Dim LsDiagdate          As String * 10
    
    Dim LsDescr             As String
    
    LsJNum = Trim(LsClass) & "-" & LsDateYY & "-" & LsSeqNum
    
    LsPtNo = lblPtno.Caption
    LsSname = lblSname.Caption
    
    LsChiefName = lblChiefName
    LsDrName = lblDrName
    
    LsJDate = lblJDate
    LsDiagdate = lblRoomCode
    
    
    Printer.Print
    Printer.Print
                        
    Printer.FontName = "돋움체"
    Printer.FontSize = 28
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    If Mid(lblPtno.Caption, 1, 1) = "C" Then
        Printer.Print Tab(12); " 세포병리 검사보고"     '5
    Else
        Printer.Print Tab(12); " 조직병리 검사보고"     '5
    End If
    
    Printer.Print
    
    
    Printer.FontName = "돋움체" '"바탕체"
    Printer.FontSize = 11
    
    Printer.FontBold = True
    
    Printer.FontItalic = False
    Printer.FontUnderline = False
    Printer.Print
            
'    LsJNum = LsClass & LsDateYY & LsSeqNum
'
'    LsPtNo = lblPtno.Caption
'    LsSname = lblSname.Caption
'
'    LsChiefName = lblChiefName
'    LsDrName = lblDrName
'
'    LsJDate = lblJDate
'    LsDiagdate = lblRoomCode
    LsDescr = txtDiag.Text
    
    Printer.Print
    
    Printer.Print Tab(10); "병리번호 : " & LsJNum
    Printer.Print Tab(10); "등록번호 : " & LsPtNo
    Printer.Print Tab(10); "환 자 명 : " & LsSname
    Printer.Print Tab(10); "검 사 자 : " & LsChiefName
    Printer.Print Tab(10); "의뢰의사 : " & LsDrName;
    Printer.Print Tab(10); "접수일자 : " & LsJDate;
    Printer.Print Tab(10); "보고일자 : " & LsDiagdate;
    
    Printer.Print
    Printer.Print
    
'    Printer.Print Tab(7); LsDescr
    GoSub RESULT_PRINT
    
    Printer.EndDoc

    Exit Sub
    
RESULT_PRINT:
    Dim LsStr       As String
    Dim LsChr       As String
    Dim LsLineStr   As String
    Dim LiLen       As Integer
    Dim LiPos       As Integer
    Dim y           As Integer
    Dim LF
    Dim CR
    
    LF = Chr(10)
    CR = Chr(13)
    
    If Trim(LsDescr) = "" Then Return
    
    LiLen = LenB(LsDescr)
    LsStr = LsDescr
    LiPos = 0
    LsLineStr = ""
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    Do
        LiPos = LiPos + 1
        LsChr = Mid(LsStr, LiPos, 1)
        If LsChr <> CR And LsChr <> LF Then
            LsLineStr = LsLineStr & LsChr
        End If
        
        If LsChr = LF Then

            Printer.Print Tab(10); LsLineStr
            LsLineStr = ""
        
        ElseIf LenB(LsLineStr) > 180 Then      '74  150  ' max cols
            GoSub SUB_LINEFEED_PROC
        End If
    
    Loop Until (LiPos > LiLen)
    
    Printer.Print Tab(10); LsLineStr
    
    Return
    

SUB_LINEFEED_PROC:
    
    Dim LsTempChr       As String
    
    LsTempChr = RightB(LsLineStr, 1)
    
    If Trim(LsTempChr) = "" Then
        Printer.Print Tab(10); LsLineStr
    ElseIf LsTempChr >= "A" And LsTempChr <= "z" Then
        If Trim(MidH(LsStr, LiPos + 1, 1)) = "" Then
            Printer.Print Tab(10); LsLineStr
        Else
            Printer.Print Tab(10); LsLineStr & "-"
        End If
    Else
        Printer.Print Tab(10); LsLineStr
    End If
    
    LsLineStr = ""
    
    Return


End Sub

Private Sub cmdRemark_Click()
    
    If LbRemark = False Then
        txtRemark.Visible = True
        LbRemark = True
    Else
        txtRemark.Visible = False
        LbRemark = False
    End If
 
End Sub

Private Sub cmdView_Click()

    Dim rs                  As ADODB.Recordset
    
    gSFrDate = Format(dtFromDate.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    lblRownum.Caption = ""
    lstPtList.Clear
    lblPtno = ""
    lblSname = ""
    lblRoomCode = ""
    lblChiefName = ""
    lblDrName = ""
    lblSex = ""
    txtRemark.Text = ""
    txtDiag.Text = ""
    
    lblRownum.Caption = ""
    
    LsDiag = ""
    
    
    If optSname.Value = True And LenH(txtSName.Text) < 4 Then
        Exit Sub
    End If

'/------------ Option button Check -----------
'    intOPTION = 0
'    If optFromTo.Value = True Then intOPTION = intOPTION + 1
'    If optAnatNo.Value = True Then intOPTION = intOPTION + 1
'    If optPtNo.Value = True Then intOPTION = intOPTION + 1
'    If optSName.Value = True Then intOPTION = intOPTION + 1
'
'    If intOPTION = 0 Then Exit Sub   ' 선택된 조건이 하나도 없을때...
'/--------------------------------------------
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.*, ROWNUM "
    strSQL = strSQL & "   FROM TWANAT_DIAG A"
    strSQL = strSQL & "  WHERE GbResult >= '4'"         ' 9 => 4 변경
    
    If optFromTo.Value = True Then
        strSQL = strSQL & " AND   DiagDate >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd')"
        strSQL = strSQL & " AND   DiagDate <= TO_DATE('" & gSToDate & "','yyyy-MM-dd')"
    End If
    
    If optAnatNo.Value = True Then
        strSQL = strSQL & " AND   CLass  = '" & txtClass.Text & "'"
        strSQL = strSQL & " AND   Dateyy = '" & txtDateYY.Text & "'"
        strSQL = strSQL & " AND   Seqnum = '" & txtSeqnum.Text & "'"
    End If
    
    If optPtNo.Value = True Then
        strSQL = strSQL & " AND   Ptno   = '" & txtPtNo.Text & "'"
    End If
    
    If optSname.Value = True Then
        strSQL = strSQL & " AND   SName like '%" & txtSName.Text & "%'"
    End If
    
    strSQL = strSQL & " AND ROWNUM < 100 "
    strSQL = strSQL & " ORDER BY CLASS, DATEYY, SEQNUM DESC"
    
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        LsClass = rs.Fields("CLass").Value & ""
        LsDateYY = rs.Fields("DateYY").Value & ""
        LsSeqNum = rs.Fields("Seqnum").Value & ""
        
        lstPtList.AddItem LsClass & "-" & LsDateYY & "-" & LsSeqNum
        
        If LSROWNUM < rs.Fields("ROWNUM").Value Then
            LSROWNUM = rs.Fields("ROWNUM").Value
        End If
        rs.MoveNext
    Loop
        
    AdoCloseSet rs
    
    If LSROWNUM >= 99 Then
        lblRownum.Caption = "         조회결과가 99개 이상입니다.             결과범위를 줄여서 조회 바랍니다. "
    End If
    
    lstPtList.SetFocus
    lstPtList.ListIndex = 0
        
End Sub


Private Sub Form_Activate()

    optFromTo.Value = True
    
    cmdView.SetFocus

End Sub

Private Sub Form_Load()

    dtFromDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
'    optFromTo.Value = True
    
End Sub

Private Sub lstPtList_Click()

    LsClass = Mid(lstPtList.List(lstPtList.ListIndex), 1, 2)
    LsDateYY = Mid(lstPtList.List(lstPtList.ListIndex), 4, 4)
    LsSeqNum = Mid(lstPtList.List(lstPtList.ListIndex), 9, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT  a.*, a.RowID,"
    strSQL = strSQL & "         TO_CHAR(a.Jdate,    'YYYY-MM-DD') Jdate,"
    strSQL = strSQL & "         TO_CHAR(a.DiagDate, 'YYYY-MM-DD') DiagDate,"
    strSQL = strSQL & "         b.Deptnamek, c.Drname, d.Name"
    strSQL = strSQL & " FROM    TWANAT_Diag     a,"
    strSQL = strSQL & "         TWBAS_Dept      b,"
    strSQL = strSQL & "         TWBAS_Doctor    c,"
    strSQL = strSQL & "         TWBAS_PASS d"
    strSQL = strSQL & " WHERE   a.CLass    =  '" & LsClass & "'"
    strSQL = strSQL & " AND     a.Dateyy   =  '" & LsDateYY & "'"
    strSQL = strSQL & " AND     a.Seqnum   =  '" & LsSeqNum & "'"
    strSQL = strSQL & " AND     a.DeptCode =  b.Deptcode(+)"
    strSQL = strSQL & " AND     a.Drcode   =  c.Drcode(+)"
    strSQL = strSQL & " AND     a.Chief    =  d.IDNUMBER(+)"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lblPtno = rs.Fields("Ptno").Value & ""
    lblSname = rs.Fields("Sname").Value & ""
    lblRoomCode = rs.Fields("DiagDate").Value & ""
    lblSex = rs.Fields("SEX").Value & "  " & rs.Fields("RoomCode").Value & ""
    lblJDate = rs.Fields("Jdate").Value & ""
    LsDeptCode = rs.Fields("DeptCode").Value & ""
    lblDrName = rs.Fields("Drname").Value & ""
    
'    txtDiag.Text = rs.Fields("DESCR").Value & ""
    txtDiag.Text = (rs.Fields("DIAGEYE").Value & "") & (rs.Fields("DIAGPRE").Value & "") & (rs.Fields("DESCR").Value & "") & (rs.Fields("DIAGADD").Value & "")
    
    txtRemark.Text = rs.Fields("DrRemark").Value & ""
    lblChiefName = rs.Fields("Name").Value & ""
    
    
    AdoCloseSet rs
    
    Call Special_result
    
    txtDiag.Text = txtDiag.Text & vbCrLf & LsDiag
    
End Sub


Private Sub optAnatNo_Click()

'    optFromTo.Value = False
'    optAnatNo.Value = True
'    optPtNo.Value = False
'    optSName.Value = False
    
    Call Txt_Clear
    lstPtList.Clear

    
    txtClass.Text = "P"
    txtClass.SetFocus

End Sub

Private Sub optAnatNo_GotFocus()

    DoEvents
    txtClass.IMEMode = vbIMEModeAlpha
    
End Sub


Private Sub optFromTo_Click()

    dtFromDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

'    optFromTo.Value = True
'    optAnatNo.Value = False
'    optPtNo.Value = False
'    optSName.Value = False
        
    Call Txt_Clear
    
    lstPtList.Clear
    
    cmdView.SetFocus
    
End Sub

Private Sub optPtNo_Click()

'    optFromTo.Value = False
'    optAnatNo.Value = False
'    optPtNo.Value = True
'    optSName.Value = False
        
    Call Txt_Clear
    
    lstPtList.Clear
    txtPtNo.SetFocus

End Sub

Private Sub optSName_Click()

'    optFromTo.Value = False
'    optAnatNo.Value = False
'    optPtNo.Value = False
'    optSName.Value = True

'    cmdSearch.Enabled = True
        
    Call Txt_Clear
    
    lstPtList.Clear
    
    txtSName.SetFocus
    
End Sub


Private Sub txtClass_GotFocus()
 
    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)

End Sub


Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    
    
    If KeyAscii > 0 Then
        TabCheck = TabCheck + 1
    End If
    
    If TabCheck = 1 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If
    
'    If KeyAscii <> 13 Then Exit Sub
'    KeyAscii = 0
    
'    SendKeys "{tab}"

End Sub


Private Sub txtDateYY_GotFocus()
 
    txtDateYY.SelStart = 0
    txtDateYY.SelLength = Len(txtDateYY.Text)

End Sub


Private Sub txtDateYY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Len(txtPtNo.Text) = 4 Then

        SendKeys "{tab}"
    End If

End Sub

Private Sub txtDATEYY_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 0 And KeyAscii <> 8 Then
        TabCheck = TabCheck + 1
    Else
        TabCheck = TabCheck - 1
    End If
    
    If TabCheck = 4 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    

    SendKeys "{tab}"
    
End Sub




Private Sub txtPtNo_GotFocus()
 
    txtPtNo.SelStart = 0
    txtPtNo.SelLength = Len(txtPtNo.Text)

End Sub


Private Sub TxtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtPtNo_LostFocus()
    
    txtPtNo.Text = UCase(txtPtNo.Text)
    txtPtNo.Text = Format(txtPtNo, "00000000")

End Sub

Private Sub txtSeqnum_GotFocus()
 
    txtSeqnum.SelStart = 0
    txtSeqnum.SelLength = Len(txtSeqnum.Text)

End Sub


Private Sub txtSeqnum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
'    SendKeys "{tab}"
    cmdView.SetFocus

End Sub


Private Sub txtSName_GotFocus()

    DoEvents
    txtSName.IMEMode = vbIMEModeHangul
End Sub


Private Sub txtSName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
'    SendKeys "{tab}"
    cmdView.SetFocus

End Sub



Private Sub Special_result()

    Dim SpecialChar         As String
    Dim sSpecial(30)        As String
    Dim lsnumber


    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  Class  = '" & LsClass & "' "
    strSQL = strSQL & "   AND  DateYY  = '" & LsDateYY & "' "
    strSQL = strSQL & "   AND  SeqNum  = '" & LsSeqNum & "' "
                             
    Result = AdoOpenSet(rs, strSQL)
    If Result Then
        Do Until rs.EOF
            For j = 1 To 30
                SpecialChar = "SPECIAL" & Format(j, "00")
                sSpecial(j) = Trim(rs.Fields(SpecialChar).Value & "")
            Next j
            
            SLIDENO = Trim(rs.Fields("slid").Value & "")
            
            LsElectro = Trim(rs.Fields("ELECTROSCOPE").Value & "")
            LsFlow = Trim(rs.Fields("FLOW").Value & "")
            
            rs.MoveNext
        Loop
    End If
    
    
    Dim FlagSpecial1
    Dim FlagSpecial2
    Dim FlagSpecial3
    Dim FlagSpecial4
    Dim FlagSpecial5
    
    Dim Flag_Data1
    Dim Flag_Data2
    Dim Flag_Data3
    Dim Flag_Data4
    Dim Flag_Data5
    
    
    For j = 1 To 30
        Select Case sSpecial(j)
                Case "853001" To "853999"
                     FlagSpecial1 = 1
                     Flag_Data1 = Flag_Data1 & Special_Load(sSpecial(j)) & ", "
                Case "857001" To "857999"
                     FlagSpecial2 = 1
                     Flag_Data2 = Flag_Data2 & Special_Load(sSpecial(j)) & ", "
                Case "854001" To "854999"
                     FlagSpecial3 = 1
                     Flag_Data3 = Flag_Data3 & Special_Load(sSpecial(j)) & ", "
                Case "856001" To "856999"
                     FlagSpecial4 = 1
                     Flag_Data4 = Flag_Data4 & Special_Load(sSpecial(j)) & ", "
                Case "855001"
                     FlagSpecial5 = 1
                     Flag_Data5 = "Y "
        End Select
    Next j
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If FlagSpecial1 = 1 Then
         If Len(Flag_Data1) <= 60 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, Len(Flag_Data1) - 2) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 60 And Len(Flag_Data1) <= 120 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, Len(Flag_Data1) - 62) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 120 And Len(Flag_Data1) <= 180 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 121, Len(Flag_Data1) - 122) & vbCrLf
         
         ElseIf Len(Flag_Data1) > 181 Then
             LsDiag = LsDiag & "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 61, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 121, 60) & vbCrLf
             LsDiag = LsDiag & "              : " & Mid(Flag_Data1, 181, Len(Flag_Data1) - 182) & vbCrLf
         End If
    End If
    
    If FlagSpecial2 = 1 Then
        If Len(Flag_Data2) <= 60 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, Len(Flag_Data2) - 2) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 60 And Len(Flag_Data1) <= 120 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, Len(Flag_Data2) - 62) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 120 And Len(Flag_Data1) <= 180 Then
             LsDiag = LsDiag & "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 121, Len(Flag_Data2) - 122) & vbCrLf
        
        ElseIf Len(Flag_Data2) > 180 Then
             LsDiag = LsDiag & "Immunohistochemical stain : " & Mid(Flag_Data2, 1, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 61, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 121, 60) & vbCrLf
             LsDiag = LsDiag & "                          : " & Mid(Flag_Data2, 181, Len(Flag_Data2) - 182) & vbCrLf
        End If
         
    End If
    If FlagSpecial3 = 1 Then
         LsDiag = LsDiag & "IMMUNOFLUORESCENCE : " & Mid(Flag_Data3, 1, Len(Flag_Data3) - 2) & vbCrLf
    End If
    If FlagSpecial4 = 1 Then
         LsDiag = LsDiag & "ENZYME HISTOCHEMISTRY : " & Mid(Flag_Data4, 1, Len(Flag_Data4) - 2) & vbCrLf
    End If
    
    If LsElectro = "Y" Then
         LsDiag = LsDiag & "ELECTRON MICROSCOPIC EXAM : Y " & vbCrLf
    End If
    
    If LsFlow = "Y" Then
         LsDiag = LsDiag & "FLOW CYTOMETRY : Y " & vbCrLf
    End If
    
    AdoCloseSet rs
    
End Sub



'*************************************************************************************************************

'미사용
'Private Sub ssResult_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'    Dim CurrRow             As Integer
'    Dim i                   As Integer
'
'    lstPtList.Clear
'    lstPtList.Clear
'    lblPtNo = ""
'    lblSname = ""
'    lblRoomCode = ""
'    lblChiefName = ""
'    lblDrName = ""
'    lblSex = ""
'    txtRemark.Text = ""
'    txtDiag.Text = ""
'
'    CurrRow = Row
'    ssResult.Row = CurrRow
'    ssResult.Col = 3
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & " FROM   TWANAT_DIAG "
'    strSQL = strSQL & " WHERE  GBRESULT = '9'"
'    strSQL = strSQL & " AND    PTNO = '" & ssResult.Text & "'"
'
'    Result = AdoOpenSet(rs, strSQL)
'
'    If Result = False Then Exit Sub
'
'    Do Until rs.EOF
'        lstPtList.AddItem rs.Fields("CLass").Value & "-" & _
'                          rs.Fields("Dateyy").Value & "-" & _
'                          rs.Fields("SEQNUM").Value & ""
'        rs.MoveNext
'    Loop
'    AdoCloseSet rs
    
'    Picture5.Visible = False
'    cmdView.Enabled = True
'    cmdRemark.Enabled = True
'    cmdExit.Enabled = True

'End Sub


'Private Sub SSCommand1_Click()'
'
'    Picture5.Visible = False
'    cmdView.Enabled = True
'    cmdRemark.Enabled = True
'    cmdExit.Enabled = True
'
'End Sub



'Private Sub cmdSearch_Click()''
'
'    Picture5.Visible = True
'
'    Call SSInitialize(ssResult)
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT  a.*, a.RowID,"
'    strSQL = strSQL & "         TO_CHAR(a.Jdate,    'YYYY-MM-DD') Jdate,"
'    strSQL = strSQL & "         TO_CHAR(a.DiagDate, 'YYYY-MM-DD') DiagDate,"
'    strSQL = strSQL & "         b.Deptnamek, c.Drname"
'    strSQL = strSQL & " FROM    TWANAT_Diag  a,"
'    strSQL = strSQL & "         TWBAS_Dept   b,"
'    strSQL = strSQL & "         TWBAS_Doctor c "
'    strSQL = strSQL & " WHERE   a.GBRESULT = '9'"
'    strSQL = strSQL & " AND     a.SNAME  LIKE '" & Trim(txtSName) & "%'"
'    strSQL = strSQL & " AND     a.DeptCode = b.DeptCode(+)"
'    strSQL = strSQL & " AND     a.Drcode   = c.Drcode(+)"
'
'    Result = AdoOpenSet(rs, strSQL)
'
'    If Result = False Then
'        txtSName = ""
'        Exit Sub
'    End If
'
'    Do Until rs.EOF
'        ssResult.Row = ssResult.DataRowCnt + 1
'        ssResult.Col = 2:  ssResult.Text = rs.Fields("CLass").Value & "-" & _
'                                           rs.Fields("DateYY").Value & "-" & _
'                                           rs.Fields("Seqnum").Value & ""
'        ssResult.Col = 3:  ssResult.Text = rs.Fields("Ptno").Value & ""
'        ssResult.Col = 4:  ssResult.Text = rs.Fields("Sname").Value & ""
'        ssResult.Col = 5:  ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여")
'        ssResult.Col = 6:  ssResult.Text = rs.Fields("ageYY").Value & ""
'        ssResult.Col = 7:  ssResult.Text = rs.Fields("DiagDate").Value & ""
'        ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
'        ssResult.Col = 9:  ssResult.Text = rs.Fields("RoomCode").Value & ""
'        ssResult.Col = 10: ssResult.Text = rs.Fields("Deptnamek").Value & ""
'        ssResult.Col = 11: ssResult.Text = rs.Fields("Drname").Value & ""
'        rs.MoveNext
'    Loop
'    AdoCloseSet rs
'
'    cmdView.Enabled = False
'    cmdRemark.Enabled = False
'    cmdExit.Enabled = False'
'
'
'End Sub
'


