VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form Anato_OCS_Jeobsu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  '없음
   Caption         =   "Order 접수"
   ClientHeight    =   8370
   ClientLeft      =   1590
   ClientTop       =   1830
   ClientWidth     =   11850
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8370
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   7815
      Left            =   6540
      TabIndex        =   24
      Top             =   600
      Width           =   5355
      _Version        =   131072
      _ExtentX        =   9446
      _ExtentY        =   13785
      _StockProps     =   100
      TabsPerRow      =   3
      TabCount        =   3
      Tab             =   1
      OffsetFromClientTop=   -1  'True
      BookRingShowHole=   -1  'True
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      TabCaption      =   "ANATO117.frx":0000
      Begin VB.Frame Frame3 
         Caption         =   "검사정보"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5835
         Left            =   -20024
         TabIndex        =   38
         Top             =   -22694
         Width           =   4965
         Begin FPSpread.vaSpread ssItemJeobsu 
            Height          =   5076
            Left            =   180
            TabIndex        =   42
            Top             =   648
            Width           =   4608
            _Version        =   196608
            _ExtentX        =   8128
            _ExtentY        =   8954
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
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
            MaxCols         =   9
            MaxRows         =   50
            ScrollBars      =   2
            SelectBlockOptions=   4
            ShadowColor     =   12632256
            ShadowDark      =   8421504
            ShadowText      =   0
            SpreadDesigner  =   "ANATO117.frx":01C1
            UserResize      =   0
            VisibleCols     =   9
            VisibleRows     =   50
         End
         Begin VB.TextBox txtSeqNum 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2550
            MaxLength       =   5
            TabIndex        =   41
            Top             =   300
            Width           =   915
         End
         Begin VB.TextBox txtClass 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   40
            Top             =   300
            Width           =   465
         End
         Begin VB.TextBox txtDateYY 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1770
            MaxLength       =   4
            TabIndex        =   39
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label4 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "접수번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   270
            TabIndex        =   43
            Top             =   375
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "환자정보"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1350
         Left            =   -20024
         TabIndex        =   25
         Top             =   -16829
         Width           =   4965
         Begin VB.TextBox txtRoomcode 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Left            =   3510
            TabIndex        =   26
            Top             =   945
            Width           =   1185
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   270
            TabIndex        =   37
            Top             =   390
            Width           =   720
         End
         Begin VB.Label lblPtno 
            BackColor       =   &H00C0C0FF&
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
            Height          =   285
            Left            =   1230
            TabIndex        =   36
            Top             =   375
            Width           =   1185
         End
         Begin VB.Label Label3 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "성  명"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2790
            TabIndex        =   35
            Top             =   390
            Width           =   540
         End
         Begin VB.Label lblSname 
            BackColor       =   &H00C0C0FF&
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
            Height          =   285
            Left            =   3510
            TabIndex        =   34
            Top             =   375
            Width           =   1185
         End
         Begin VB.Label Label5 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "나    이"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   270
            TabIndex        =   33
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lblAge 
            BackColor       =   &H00C0C0FF&
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
            Height          =   285
            Left            =   1230
            TabIndex        =   32
            Top             =   660
            Width           =   1185
         End
         Begin VB.Label Label7 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "성  별"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2790
            TabIndex        =   31
            Top             =   690
            Width           =   540
         End
         Begin VB.Label lblSex 
            BackColor       =   &H00C0C0FF&
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
            Height          =   285
            Left            =   3510
            TabIndex        =   30
            Top             =   660
            Width           =   1185
         End
         Begin VB.Label Label9 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "진 료 과"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   270
            TabIndex        =   29
            Top             =   960
            Width           =   720
         End
         Begin VB.Label lblDeptName 
            BackColor       =   &H00C0C0FF&
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
            Height          =   285
            Left            =   1230
            TabIndex        =   28
            Top             =   945
            Width           =   1185
         End
         Begin VB.Label Label11 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "병  실"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2790
            TabIndex        =   27
            Top             =   975
            Width           =   540
         End
      End
      Begin Threed.SSPanel picRecept 
         Height          =   7185
         Left            =   180
         TabIndex        =   44
         Top             =   420
         Width           =   5145
         _Version        =   65536
         _ExtentX        =   9075
         _ExtentY        =   12674
         _StockProps     =   15
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin FPSpread.vaSpread ssRecept 
            Height          =   6630
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   4845
            _Version        =   196608
            _ExtentX        =   8546
            _ExtentY        =   11695
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
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
            MaxCols         =   12
            ScrollBars      =   2
            SelectBlockOptions=   8
            ShadowColor     =   10932443
            ShadowDark      =   4494504
            ShadowText      =   8404992
            SpreadDesigner  =   "ANATO117.frx":1363
            UserResize      =   0
            VisibleCols     =   12
            VisibleRows     =   500
            VScrollSpecial  =   -1  'True
         End
         Begin VB.Label lblReceptTitle 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   4845
         End
      End
      Begin Threed.SSPanel picRecept2 
         Height          =   7185
         Left            =   -20324
         TabIndex        =   47
         Top             =   -22604
         Width           =   5145
         _Version        =   65536
         _ExtentX        =   9075
         _ExtentY        =   12674
         _StockProps     =   15
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Enabled         =   0   'False
         Begin FPSpread.vaSpread ssRecept2 
            Height          =   6636
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   4848
            _Version        =   196608
            _ExtentX        =   8551
            _ExtentY        =   11705
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
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
            MaxCols         =   12
            ScrollBars      =   2
            SelectBlockOptions=   8
            ShadowColor     =   10932443
            ShadowDark      =   4494504
            ShadowText      =   8404992
            SpreadDesigner  =   "ANATO117.frx":2ED2
            UserResize      =   0
            VisibleCols     =   12
            VisibleRows     =   500
            VScrollSpecial  =   -1  'True
         End
         Begin VB.Label lblReceptTitle2 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Width           =   4845
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4560
      Left            =   5025
      ScaleHeight     =   4500
      ScaleWidth      =   1455
      TabIndex        =   7
      Top             =   2700
      Width           =   1515
      Begin Threed.SSCommand cmdExit 
         Height          =   900
         Left            =   0
         TabIndex        =   4
         Top             =   3600
         Width           =   1452
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "이전화면"
         ForeColor       =   -2147483630
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
         Picture         =   "ANATO117.frx":4A41
      End
      Begin Threed.SSCommand cmdOcsOrder 
         Height          =   900
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "검사대기자"
         ForeColor       =   -2147483630
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
         Picture         =   "ANATO117.frx":4D5B
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   900
         Left            =   0
         TabIndex        =   3
         Top             =   2700
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "접수취소"
         ForeColor       =   -2147483630
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
         Picture         =   "ANATO117.frx":5635
      End
      Begin Threed.SSCommand cmdJeobsuList 
         Height          =   900
         Left            =   0
         TabIndex        =   1
         Top             =   900
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "접수환자"
         ForeColor       =   -2147483630
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
         Picture         =   "ANATO117.frx":5A87
      End
      Begin Threed.SSCommand cmdJeobsuAdd 
         Height          =   900
         Left            =   0
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "접수등록"
         ForeColor       =   -2147483630
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
         Picture         =   "ANATO117.frx":5DA1
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   480
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   847
      _StockProps     =   15
      Caption         =   "선  택"
      ForeColor       =   8388608
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Font3D          =   2
   End
   Begin Threed.SSPanel ssTitle 
      Align           =   1  '위 맞춤
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11850
      _Version        =   65536
      _ExtentX        =   20902
      _ExtentY        =   952
      _StockProps     =   15
      Caption         =   "OCS ORDER 접수"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.Label Label6 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "사용자:"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9450
         TabIndex        =   9
         Top             =   105
         Width           =   915
      End
      Begin VB.Label lblExamName 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10350
         TabIndex        =   6
         Top             =   105
         Width           =   1635
      End
   End
   Begin VB.TextBox txtJdate 
      Height          =   285
      Left            =   5250
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order 정보"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7815
      Left            =   60
      TabIndex        =   13
      Top             =   600
      Width           =   4965
      Begin Threed.SSCommand cmdOrgan 
         Height          =   288
         Left            =   4368
         TabIndex        =   53
         Top             =   672
         Width           =   372
         _Version        =   65536
         _ExtentX        =   656
         _ExtentY        =   508
         _StockProps     =   78
         Caption         =   "&H"
      End
      Begin Threed.SSCommand cmdOp 
         Height          =   288
         Left            =   4368
         TabIndex        =   52
         Top             =   372
         Width           =   372
         _Version        =   65536
         _ExtentX        =   656
         _ExtentY        =   508
         _StockProps     =   78
         Caption         =   "&H"
      End
      Begin VB.TextBox txtRemark1 
         BackColor       =   &H00EBF5EB&
         Enabled         =   0   'False
         Height          =   288
         Left            =   1755
         TabIndex        =   51
         Top             =   972
         Width           =   2970
      End
      Begin VB.TextBox txtOrganPart 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   1755
         TabIndex        =   19
         Top             =   672
         Width           =   2610
      End
      Begin VB.TextBox txtOpname 
         BackColor       =   &H00C0E0FF&
         Height          =   288
         Left            =   1755
         TabIndex        =   18
         Top             =   372
         Width           =   2610
      End
      Begin FPSpread.vaSpread ssOrder 
         Height          =   5724
         Left            =   216
         TabIndex        =   14
         Top             =   1356
         Width           =   4512
         _Version        =   196608
         _ExtentX        =   7959
         _ExtentY        =   10097
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         MaxCols         =   22
         ScrollBars      =   2
         SelectBlockOptions=   4
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "ANATO117.frx":60BB
         UserResize      =   1
         VisibleCols     =   22
         VisibleRows     =   500
         VScrollSpecial  =   -1  'True
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3330
         TabIndex        =   22
         Top             =   7170
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdOrderJeobsu 
         Caption         =   "선택완료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1755
         TabIndex        =   21
         Top             =   7170
         Width           =   1395
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "모두선택"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         TabIndex        =   20
         Top             =   7170
         Width           =   1395
      End
      Begin VB.Label Label13 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   975
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "장기/검체"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "수술명/검사명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Top             =   375
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   1080
      Left            =   5025
      ScaleHeight     =   1020
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   1590
      Width           =   1515
      Begin VB.OptionButton optRefferal 
         BackColor       =   &H00FFC0C0&
         Caption         =   "REFFERAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   672
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.OptionButton optCytology 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CYTOLOGY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   375
         Width           =   1470
      End
      Begin VB.OptionButton optHistology 
         BackColor       =   &H00FFC0C0&
         Caption         =   "HISTOLOGY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   75
         Width           =   1470
      End
   End
   Begin MSComCtl2.DTPicker dtFromDate 
      Height          =   315
      Left            =   5100
      TabIndex        =   54
      Top             =   1230
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24444931
      CurrentDate     =   36526
   End
End
Attribute VB_Name = "Anato_OCS_Jeobsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LiOldRow                As Integer
Dim LsCmdOKflag             As String
Dim LiTime                  As Long
Dim LsJeobsuAdd             As Boolean
Dim LiSeqLimit              As Long
    
Dim LsjeobsuDt              As String
Dim LiJtimeHH               As Integer
Dim LiJtimeMM               As Integer
Dim LsSex                   As String * 1
Dim LiAgeYY                 As Integer
Dim LiAgeMM                 As Integer
Dim LsInDate                As String
Dim DeptCode                As String * 6
Dim LsGbio                  As String * 1
Dim LsBi                    As String * 2
Dim LsGbEr                  As String * 1
Dim LsGeomchcd              As String * 8
Dim LsGeomsaGu              As String * 1

Dim LsOpname                As String * 8
Dim LsOrganPart             As String * 8
Dim LsRemark1

Dim LsOrderDt               As String
Dim LiOrderNo               As Long
Dim LsCmDoctor              As String
Dim LsDrCode                As String
Dim LsDeptCode              As String

Dim CodeGuchk               As String


'''   미사용   ''''''
'Private Sub cmdAdd_Click()
'
'    Dim rs                  As ADODB.Recordset
'
'    Dim i                   As Integer
'    Dim j                   As Integer
'    Dim LiRowCnt            As Integer
'    Dim LsPtNo              As String
'
'
'    LsJeobsuAdd = True
'
'    Call SSInitialize(ssOrder)
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT a.*, a.RowID"
'    strSQL = strSQL & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
'    strSQL = strSQL & "        TO_CHAR(a.OrderDt, 'YYYY-MM-DD') OrderDt, "
'    strSQL = strSQL & "        b.iTemNM"
'    strSQL = strSQL & " FROM   TWEXAM_Order  a,"
'    strSQL = strSQL & "        TWEXAM_iTemML b "
'    strSQL = strSQL & " WHERE  a.Ptno     =  '" & lblPtno & "'"
'    strSQL = strSQL & " AND    a.Dateyy   =   " & GiExamNumb
'    strSQL = strSQL & " AND   (a.JeobsuYn = ' ' OR a.JeobsuYn IS NULL)"
'    strSQL = strSQL & " AND    a.iTemCD   = b.Codeky(+)"
'
'    Result = AdoOpenSet(rs, strSQL)
'
'    If Result = False Then Exit Sub
'
'    Do Until rs.EOF
'        ssOrder.Row = ssOrder.DataRowCnt + 1
'        ssOrder.Col = 2:  ssOrder.Text = rs.Fields("ITEMCD").Value & ""
'
'        If rs.Fields("iTemNM").Value & "" <> "" Then
'            ssOrder.Col = 3:   ssOrder.Text = rs.Fields("ITEMNM").Value & ""
'            ssOrder.Col = 6:   ssOrder.Text = "I"
'        Else
'           'SUB_EXAM_ITEM_PROC
'            Dim adoRoutine      As ADODB.Recordset
'
'            strSQL = " SELECT  RoutinNM  FROM TWEXAM_ROUTINE  WHERE RoutinCD =  '" & ssOrder.Text & "' "
'            Result = AdoOpenSet(adoRoutine, strSQL)
'            If Result = False Then Return
'            ssOrder.Col = 3:   ssOrder.Text = adoRoutine.Fields("ROUTINNM").Value & ""
'            ssOrder.Col = 6:   ssOrder.Text = "R"
'            AdoCloseSet adoRoutine
'        End If
'
'        ssOrder.Col = 4:  ssOrder.Text = Format(rs.Fields("JTIMEHH").Value, "00")
'                          ssOrder.Text = ssOrder.Text & ":" & Format(rs.Fields("JTIMEMM").Value, "00")
'        ssOrder.Col = 5:  ssOrder.Text = rs.Fields("RowID").Value & ""
'        ssOrder.Col = 7:  ssOrder.Text = rs.Fields("JTIMEHH").Value & ""
'        ssOrder.Col = 8:  ssOrder.Text = rs.Fields("JTIMEMM").Value & ""
'        ssOrder.Col = 9:  ssOrder.Text = rs.Fields("AGEMM").Value & ""
'        ssOrder.Col = 10: ssOrder.Text = rs.Fields("GBIO").Value & ""
'        ssOrder.Col = 11: ssOrder.Text = rs.Fields("GBER").Value & ""
'        ssOrder.Col = 12: ssOrder.Text = rs.Fields("GEOMCHCD").Value & ""
'        ssOrder.Col = 13: ssOrder.Text = rs.Fields("GEOMSAGU").Value & ""
'        ssOrder.Col = 14: ssOrder.Text = rs.Fields("ORDERDT").Value & ""
'        ssOrder.Col = 15: ssOrder.Text = rs.Fields("ORDERNO").Value & ""
'        ssOrder.Col = 16: ssOrder.Text = rs.Fields("CMDOCTOR").Value & ""
'                          txtRemark1.Text = ssOrder.Text
'        ssOrder.Col = 17: ssOrder.Text = rs.Fields("DRCODE").Value & ""
'        ssOrder.Col = 18: ssOrder.Text = rs.Fields("BI").Value & ""
'        rs.MoveNext
'    Loop
'    AdoCloseSet rs
'
'
'End Sub

Private Sub cmdAll_Click()
    '모두선택
    Dim i           As Integer
    
    For i = 1 To ssOrder.DataRowCnt
        ssOrder.Row = i
        ssOrder.Col = 1
        ssOrder.Text = "1"
    Next i
        
    cmdOrderJeobsu.SetFocus

End Sub

Private Sub cmdCancel_Click()
    '접수취소
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim LiCancelCount       As Integer
    Dim LiDataRowCnt        As Integer
    Dim Response            As Integer
    Dim LsItemCd            As String * 8
    Dim LsRowID             As String
    Dim TmpOrderNo
    Dim TmpItemcd
    Dim Click_Check         As Boolean
    
    LsItemCd = ""
    
'    For i = 1 To ssOrder.DataRowCnt
'        ssOrder.Col = 1
'        ssOrder.Row = i
'        If ssOrder.Text = "1" Then
'            Click_Check = True
'        End If
'    Next i
    
    For i = 1 To ssItemJeobsu.DataRowCnt
        ssItemJeobsu.Col = 1
        ssItemJeobsu.Row = i
        If ssItemJeobsu.Text = "1" Then
           Click_Check = True
        End If
    Next i
    
    If Click_Check = False Then Exit Sub
    
    If vbNo = MsgBox("선택한 검사항목을 취소합니까?", vbYesNo + vbExclamation + vbDefaultButton2, "알림") Then
        Exit Sub
    End If

    LiDataRowCnt = ssItemJeobsu.DataRowCnt
    
    For i = ssItemJeobsu.DataRowCnt To 1 Step -1
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 1
        If ssItemJeobsu.Text = "1" Then
            ''''''''''''''''''''''''''''
            GoSub SUB_JEOBSU_CANCEL_PROC
        End If
    Next i
    
    
'    LiDataRowCnt = ssItemJeobsu.DataRowCnt
'
'    For i = ssItemJeobsu.DataRowCnt To 1 Step -1
'        ssItemJeobsu.Row = i
'        ssItemJeobsu.Col = 1
'        If ssItemJeobsu.Text = "1" Then
'            ''''''''''''''''''''''''''''
'            GoSub SUB_JEOBSU_CANCEL_PROC
'        End If
'    Next i
    
    
    Exit Sub
    
    
'/------------------------------------------------------------------------------------------
SUB_JEOBSU_CANCEL_PROC:
    Dim sRowID              As String
    
    ssItemJeobsu.Col = 7
    If Trim(ssItemJeobsu.Text) = "1" Or Trim(ssItemJeobsu.Text) = "0" Then
        LiCancelCount = LiCancelCount + 1
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 1
        ssItemJeobsu.Action = SS_ACTION_DELETE_ROW
        ssItemJeobsu.Action = SS_ACTION_ACTIVE_CELL
        Return
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ssItemJeobsu.Col = 5
    sRowID = ssItemJeobsu.Text
'    strSQL = " DELETE FROM TWANAT_DIAG WHERE RowID = '" & sRowID & "'"
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & " SET    GbResult = 'X'"
    strSQL = strSQL & " WHERE  RowID = '" & sRowID & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Result Then
        adoConnect.CommitTrans
        
       'SUB_ORDER_UPDATE
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 2:   LsItemCd = Trim(ssItemJeobsu.Text)
        ssItemJeobsu.Col = 8:   TmpOrderNo = Val(ssItemJeobsu.Text)
       
        strSQL = ""
        strSQL = strSQL & " UPDATE TWEXAM_ORDER "
        strSQL = strSQL & " SET    JeobsuYn  = '#' "                 '미접수 ' ' , 접수취소 '#'
        strSQL = strSQL & " WHERE  Ptno    = '" & lblPtno & "'    "
        strSQL = strSQL & " AND    iTemCD  = '" & LsItemCd & "'   "
        strSQL = strSQL & " AND    Orderno = '" & TmpOrderNo & "' "
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
'            MsgBox " 작업이 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox " TWEXAM_ORDER TABLE " & vbCrLf & _
                   " Update Error가 발생하였습니다.", vbCritical, "오류"
        End If
        'SUB_ORDER_UPDATE END
        
        LiCancelCount = LiCancelCount + 1
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 1
        ssItemJeobsu.Action = SS_ACTION_DELETE_ROW
        ssItemJeobsu.Action = SS_ACTION_ACTIVE_CELL
    
    Else
        adoConnect.RollbackTrans
        MsgBox "삭제된 데이타가 없습니다.", vbCritical, "오류"
    End If
        
    Return
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    Anato_Main.Show

End Sub


Private Sub cmdJeobsuAdd_Click()
    '접수등록
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim Response            As Integer
    Dim LsItemCd            As String
    Dim LsJeobsuTm          As String
    Dim NQuantity           As Integer
    Dim InDate              As String
    Dim sRowID              As String
    
    On Error GoTo error1
    
'    InDate = Format(LsInDate, "DD-MMM-YYYY")
    InDate = Format(LsInDate, "YYYY-MM-DD")
    
    If Trim(txtClass) = "" Then
        Response = MsgBox("접수구분을 입력하세요.", vbYes + vbInformation, "알림")
        txtClass.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDateYY) = "" Then
        Response = MsgBox("접수년도를  입력하세요.", vbYes + vbInformation, "알림")
        txtDateYY.SetFocus
        Exit Sub
    End If
    
    If Trim(txtSeqNum) = "" Then
        Response = MsgBox("일렬번호를  입력하세요.", vbYes + vbInformation, "알림")
        txtSeqNum.SetFocus
        Exit Sub
    End If

'    GoSub SUB_PATIENT_TABLE_PROC             ' 환자정보갱신
    
    '/------------------------------------------------------------------------------------------
    'SUB_GENERAL_JEOBSU_PROC:                    ' 일반검사접수
    If ssItemJeobsu.DataRowCnt = 0 Then Exit Sub 'Return
    
    'SUB_DIAG_INSERT
    Dim siTemGb             As String
    Dim siTemCD             As String
    Dim sSeqnum             As String
    Dim sOrderNo            As String
    Dim sIndate             As String
    
    For i = 1 To ssItemJeobsu.DataRowCnt
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 1
        If ssItemJeobsu.Text = 1 Then
            
            ssItemJeobsu.Col = 2:  siTemCD = Trim(ssItemJeobsu.Text)
            ssItemJeobsu.Col = 8:  sOrderNo = Trim(ssItemJeobsu.Text)
            ssItemJeobsu.Col = 9:  sSeqnum = Trim(ssItemJeobsu.Text)
            LsJeobsuTm = Format(Time, "hh:mm")
            sIndate = IIf(Trim(InDate) = "", "", Format(InDate, "YYYY-MM-DD"))
            
            strSQL = ""
            strSQL = strSQL & " SELECT CODEGU "
            strSQL = strSQL & " FROM   TWEXAM_ITEMML "
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    '        strSQL = strSQL & " WHERE  CODEKY = '" & Trim(ssItemJeobsu.Text) & "' "
            strSQL = strSQL & " WHERE  CODEKY = '" & siTemCD & "' "
            
            Result = AdoOpenSet(rs, strSQL)
            
            If Result Then
                Select Case Val(rs.Fields("Codegu").Value & "")
                    Case 80:    siTemGb = "P"
                    Case 89:    siTemGb = "C"
                    Case 90:    siTemGb = "R"
                    Case Else: siTemGb = ""
                End Select
            Else
                siTemGb = ""
            End If
            AdoCloseSet rs
            
            Select Case siTemCD
                    Case "850001" To "851401", "855001", "859001" To "859999", "858001" To "858999"
                         'insert
                    'Case "859001" To "859999"   '세포병리
                    'Case "858001" To "858999"   '분자병리
                    
                            strSQL = ""
                            strSQL = strSQL & " INSERT INTO TWANAT_DIAG "
                            strSQL = strSQL & "  (  CLASS,         DATEYY,      SEQNUM,      JTIMEHH,  "
                            strSQL = strSQL & "     JTIMEMM,       PTNO,        SNAME,                 "
                            strSQL = strSQL & "     GEOMCHCD,      ORDERDT,     ORDERNO,               "
                            strSQL = strSQL & "     DRREMARK,      ROOMCODE,    DEPTCODE,    GBIO,     "
                            strSQL = strSQL & "     DRCODE,        INDATE,      ITEMCD,                "
                            strSQL = strSQL & "     DIAGDATE,      JDATE,                              "
                            strSQL = strSQL & "     CHIEF,         GBRESULT,    GBGROSS,               "
                            strSQL = strSQL & "     SEX,           AGEYY,       AGEMM,                 "
                            strSQL = strSQL & "     OPNAME,        ORGANPART,                          "
                            strSQL = strSQL & "     ITEMGB )    "
                            strSQL = strSQL & "VALUES ('" & Trim(txtClass.Text) & "',"
                            strSQL = strSQL & "         " & Val(txtDateYY.Text) & ","
                            strSQL = strSQL & "         " & sSeqnum & ","
                            
                            strSQL = strSQL & "         " & Val(Mid(LsJeobsuTm, 1, 2)) & ", "
                            strSQL = strSQL & "         " & Val(Mid(LsJeobsuTm, 4, 2)) & ", "
                            strSQL = strSQL & "        '" & lblPtno.Caption & "',"
                            strSQL = strSQL & "        '" & lblSname.Caption & "',"
                            strSQL = strSQL & "        '" & Trim(LsGeomchcd) & "',"
                            strSQL = strSQL & "             TO_DATE('" & Trim(LsOrderDt) & "','YYYY-MM-DD'),"
                    '        strSQL = strSQL & "        '" & Trim(LiOrderNo) & "',"
                            strSQL = strSQL & "         " & Val(Trim(sOrderNo)) & ","
                            
                            strSQL = strSQL & "        '" & Quot(Trim(LsCmDoctor)) & "',"
                            
                            strSQL = strSQL & "        '" & Trim(txtRoomcode.Text) & "',"
                            strSQL = strSQL & "        '" & Trim(LsDeptCode) & "',"
                            strSQL = strSQL & "        '" & Trim(LsGbio) & "',"
                            strSQL = strSQL & "        '" & Trim(LsDrCode) & "',"
                            If sIndate <> "" Then
                                strSQL = strSQL & "         TO_DATE('" & sIndate & "','YYYY-MM-DD'),"
                            Else
                                strSQL = strSQL & "         SYSDATE,"
                            End If
                            strSQL = strSQL & "        '" & Trim(siTemCD) & "',"
                            strSQL = strSQL & "        '',"
                            strSQL = strSQL & "             TO_DATE('" & txtJdate & "','YYYY-MM-DD'),"
                            strSQL = strSQL & "        ' ',"
                            strSQL = strSQL & "        '0',"
                            strSQL = strSQL & "        '0',"
                            strSQL = strSQL & "           '" & Trim(LsSex) & "',"
                            strSQL = strSQL & "            " & Val(LiAgeYY) & ","
                            strSQL = strSQL & "            " & Val(LiAgeMM) & ","
                            
                            strSQL = strSQL & "            '" & LsOpname & "',"
                            strSQL = strSQL & "            '" & LsOrganPart & "',"
                            
                            strSQL = strSQL & "        '" & siTemGb & "')"
                        
                            adoConnect.BeginTrans
                        
                            Result = AdoExecute(strSQL)
                            
                            If Result Then
                                adoConnect.CommitTrans
                                ssItemJeobsu.Col = 5
                                If Trim(ssItemJeobsu.Text) <> "" Then
                                    sRowID = ssItemJeobsu.Text
                                    ''''''''''''''''''''''
                                    
                                    'ORDER CODE 중 ROUTINE CODE 추가 UPDATE
                                    GoSub SUB_ROUTINE_UPDATE
                                    
                                    GoSub SUB_ORDER_UPDATE
                                End If
                            Else
                               adoConnect.RollbackTrans
                               MsgBox " 작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
                            End If
                    
                    Case Else                   '기타 특수염색검사등 : 타검사의 부속코드로서 하나의 Field만 차지함
                         'update
                         
            
            
            
            
            End Select
            
        End If
    Next i
    
'    Call SSInitialize(ssItemJeobsu)
    
'    LiDataRowCnt = ssItemJeobsu.DataRowCnt
    
    For i = ssItemJeobsu.DataRowCnt To 1 Step -1
        ssItemJeobsu.Row = i
        ssItemJeobsu.Col = 1
        If ssItemJeobsu.Text = "1" Then
            ssItemJeobsu.Row = i
            ssItemJeobsu.Col = 1
            ssItemJeobsu.Action = SS_ACTION_DELETE_ROW
            ssItemJeobsu.Action = SS_ACTION_ACTIVE_CELL
        End If
    Next i
    
    cmdOcsOrder.SetFocus
    
    
    Exit Sub


'/------------------------------------------------------------------------------------------
'/------------------------------------------------------------------------------------------
SUB_PATIENT_TABLE_PROC:
    Dim adoIDNO     As ADODB.Recordset
    
    strSQL = " SELECT * FROM TWEXAM_IDNOMST WHERE Ptno = '" & lblPtno & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        GoSub SUB_PATIENT_INSERT
    Else
         AdoCloseSet rs
        GoSub SUB_PATIENT_UPDATE
    End If
      
    Return

'/------------------------------------------------------------------------------------------
SUB_PATIENT_INSERT:
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWEXAM_IDNOMST "
    strSQL = strSQL & "       (  PtNo,        Sname,        Sex,         AgeYY, "
    strSQL = strSQL & "          AgeMM,       Indate,       DeptCode,    RoomCode, "
    strSQL = strSQL & "          DrCode,      Gbio,         Bi ) "
    strSQL = strSQL & " VALUES ('" & Trim(lblPtno) & "',"
    strSQL = strSQL & "         '" & Trim(lblSname) & "',"
    strSQL = strSQL & "         '" & Trim(LsSex) & "',"
    strSQL = strSQL & "          " & Val(lblAge) & ","
    strSQL = strSQL & "          " & Val(LiAgeMM) & ","
    strSQL = strSQL & "              TO_DATE('" & LsInDate & "','YYYY-MM-DD'),"
    strSQL = strSQL & "         '" & Trim(LsDeptCode) & "',"
    strSQL = strSQL & "         '" & Trim(txtRoomcode) & "',"
    strSQL = strSQL & "         '" & Trim(LsDrCode) & "',"
    strSQL = strSQL & "         '" & Trim(LsGbio) & "',"
    strSQL = strSQL & "         '" & Trim(LsBi) & "')"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return
 

'/------------------------------------------------------------------------------------------
SUB_PATIENT_UPDATE:
    strSQL = ""
    strSQL = strSQL & " UPDATE TWEXAM_IDNOMST SET "
    strSQL = strSQL & " SET    Sname    = '" & Trim(lblSname) & "',"
    strSQL = strSQL & "        Sex      = '" & Trim(LsSex) & "',"
    strSQL = strSQL & "        AgeYY    =  " & Val(lblAge) & ","
    strSQL = strSQL & "        AgeMM    =  " & Val(LiAgeMM) & ","
    strSQL = strSQL & "        Indate   =      TO_DATE('" & LsInDate & "','YYYY-MM-DD'),"
    strSQL = strSQL & "        DeptCode = '" & Trim(LsDeptCode) & "',"
    strSQL = strSQL & "        RoomCode = '" & Trim(txtRoomcode) & "',"
    strSQL = strSQL & "        Gbio     = '" & Trim(LsGbio) & "',"
    strSQL = strSQL & "        BI       = '" & Trim(LsBi) & "'"
    strSQL = strSQL & " WHERE  PtNo     = '" & Trim(lblPtno) & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return
   

'/------------------------------------------------------------------------------------------
SUB_ROUTINE_UPDATE:
    Dim SPECIALCNT      As Integer
    
'''    Return
    
    strSQL = ""
    strSQL = strSQL & " SELECT  ROUTINCD,CODEKY "
    strSQL = strSQL & "   FROM  TWEXAM_ROUTINE "
    strSQL = strSQL & "  WHERE  ROUTINCD = '" & siTemCD & "' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        'UPDATE
        'ROWINDICATOR
        Do Until rs.EOF
            
            SPECIALCNT = SPECIALCNT + 1
            
            strSQL = ""
            strSQL = strSQL & " UPDATE TWANAT_DIAG SET "
'            strSQL = strSQL & " SET    SPECIAL01  = '" & Trim(rs.Fields("CODEKY").Value & "") & "' "
            strSQL = strSQL & "        SPECIAL" & Format(SPECIALCNT, "00") & "  = '" & Trim(rs.Fields("CODEKY").Value & "") & "' "
            strSQL = strSQL & " WHERE  CLASS      = '" & Trim(txtClass.Text) & "'"
            strSQL = strSQL & "   AND  DATEYY     = '" & Trim(txtDateYY.Text) & "'"
            strSQL = strSQL & "   AND  SEQNUM     = '" & Trim(sSeqnum) & "'"
            
            adoConnect.BeginTrans
            
            Result = AdoExecute(strSQL)
            
            If Result = True And Rowindicator > 0 Then
                adoConnect.CommitTrans
'                MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
            Else
                adoConnect.RollbackTrans
                MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
            End If
            
            If SPECIALCNT >= 30 Then Exit Do
            rs.MoveNext
        Loop
    
    End If
    AdoCloseSet rs


    Return




'/------------------------------------------------------------------------------------------
SUB_ORDER_UPDATE:
    
    NQuantity = 0
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWEXAM_ORDER "
    strSQL = strSQL & " WHERE  Ptno    = '" & Trim(lblPtno) & "' "
    strSQL = strSQL & " AND    Indate  = TO_DATE('" & Trim(InDate) & "','YYYY-MM-DD') "
    strSQL = strSQL & " AND    Slipno1 = " & GiExamNumb     '91
    strSQL = strSQL & " AND    OrderNo = " & sOrderNo    'number   'LiOrderNo
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        NQuantity = Val(rs.Fields("QUANTITY").Value & "")
    End If
        
    AdoCloseSet rs
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  PTNO   =  '" & Trim(lblPtno) & "' "
    strSQL = strSQL & " AND    INDATE  = TO_DATE('" & Trim(InDate) & "','YYYY-MM-DD') "
    strSQL = strSQL & " AND    CLASS  =  '" & txtClass & "' "
    strSQL = strSQL & " AND    OrderNo = '" & sOrderNo & "' "   'character 'LiOrderNo
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        If rs.RecordCount >= 1 And NQuantity >= 3 Then
            
            strSQL = ""
            strSQL = strSQL & " UPDATE TWEXAM_Order Set QUANTITY = '2' "   '접수 Flag
            strSQL = strSQL & "  Where RowID = '" & sRowID & "'"
            
            adoConnect.BeginTrans
            
            Result = AdoExecute(strSQL)
            
            If Result = True And Rowindicator > 0 Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
                MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
            End If
        
        ElseIf rs.RecordCount >= 1 And NQuantity = 2 Then
            
            strSQL = ""
            strSQL = strSQL & " UPDATE TWEXAM_Order Set QUANTITY = '1' "   '접수 Flag
            strSQL = strSQL & "  Where RowID = '" & sRowID & "'"
            
            adoConnect.BeginTrans
            
            Result = AdoExecute(strSQL)
            
            If Result = True And Rowindicator > 0 Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
                MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
            End If
        
        ElseIf rs.RecordCount >= 1 And NQuantity = 1 Then
            
            strSQL = ""
            strSQL = strSQL & " UPDATE TWEXAM_Order Set JeobsuYN = '*' "   '접수 Flag
            strSQL = strSQL & "  Where RowID = '" & sRowID & "'"
            
            adoConnect.BeginTrans
            
            Result = AdoExecute(strSQL)
            
            If Result = True And Rowindicator > 0 Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
                MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
            End If
        
        End If
    End If
    AdoCloseSet rs
    
    Return

error1:
    On Error Resume Next


End Sub


Private Sub cmdJeobSuList_Click()
    '접수환자
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim Response            As Integer
    
    Anato_OCS_Jeobsu.MousePointer = vbHourglass
    
    Call SSInitialize(ssOrder)
    
    'GoSub SUB_FORM_INITIALIZE_PROC
    'SUB_FORM_INITIALIZE_PROC:
    txtClass.BackColor = RGB(235, 245, 235)
    txtDateYY.BackColor = RGB(235, 245, 235)
    txtSeqNum.BackColor = RGB(235, 245, 235)
    ssItemJeobsu.Col = 1:        ssItemJeobsu.Row = 0
    ssItemJeobsu.Row2 = -1:      ssItemJeobsu.Col2 = 5
    ssItemJeobsu.BlockMode = True
    ssItemJeobsu.BackColor = RGB(235, 245, 235)
    ssItemJeobsu.BlockMode = False
    lblPtno = ""
    lblAge = ""
    lblDeptName = ""
    lblSname = ""
    lblSex = ""
    txtRoomcode = ""
    txtSeqNum = ""
    Call SSInitialize(ssRecept2)
    Call SSInitialize(ssItemJeobsu)
    ssTitle = GsExamJong & " " & "Order List"
    cmdCancel.Enabled = True
    'Return
    
'    lblRemark = ""
'    lblGbInfo = ""
'    lblGeomchName = ""
    
    LsCmdOKflag = "JEOBSU_PATIENT"
    lblReceptTitle2 = "접 수 환 자 명 단"
    lblReceptTitle2.BackColor = &H800000
    lblReceptTitle2.ForeColor = &HFFFF&
    vaTabPro1.ActiveTab = 2
    
    
    If optHistology.Value = True Then
        GiExamNumb = 85 '61  ' 91
        txtClass = "P"
    ElseIf optCytology.Value = True Then
        GiExamNumb = 85  '61     '2  ' 92
        txtClass = "C"
    Else
        GiExamNumb = 85  '61     '2  ' 92
        txtClass = "R"
    End If
    
    
    strSQL = ""
    strSQL = strSQL & " SELECT distinct a.Ptno, a.AgeYY, a.SEQNUM, a.Sex, a.DeptCode, a.RoomCode,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,'YYYY-MM-DD') Jdate,"
    strSQL = strSQL & "        a.Class, a.Dateyy, b.Sname, c.Deptnamek"
    strSQL = strSQL & " FROM    TWANAT_DIAG   a, "
    strSQL = strSQL & "         TWBAS_PATIENT b, "
    strSQL = strSQL & "         TWBAS_DEPT    c  "
    strSQL = strSQL & " WHERE   a.Class    = '" & Trim$(txtClass) & "'"
    strSQL = strSQL & " AND     a.GbResult = '0' "
    strSQL = strSQL & " AND     a.Ptno     = b.Ptno(+)"
    strSQL = strSQL & " AND     a.Deptcode = c.Deptcode(+)"
'    strSQL = strSQL & " ORDER   BY DATEYY DESC , a.SEQNUM "
    strSQL = strSQL & " ORDER   BY a.SEQNUM desc"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    ssRecept2.MaxRows = Rowindicator + 1

    Do Until rs.EOF
        ssRecept2.Row = ssRecept2.DataRowCnt + 1
        ssRecept2.Col = 2:  ssRecept2.Text = rs.Fields("Jdate").Value & ""
        ssRecept2.Col = 3:  ssRecept2.Text = Trim$(rs.Fields("PTNO").Value & "")
        ssRecept2.Col = 4:  ssRecept2.Text = Trim$(rs.Fields("SNAME").Value & "")
        ssRecept2.Col = 5:  ssRecept2.Text = Trim$(rs.Fields("ROOMCODE").Value & "")
        ssRecept2.Col = 6:  ssRecept2.Text = Trim$(rs.Fields("Deptnamek").Value & "")
        ssRecept2.Col = 7:  ssRecept2.Text = Trim$(rs.Fields("SEX").Value & "")
        ssRecept2.Col = 8:  ssRecept2.Text = Trim$(rs.Fields("AGEYY").Value & "")
        ssRecept2.Col = 9:  ssRecept2.Text = Trim$(rs.Fields("DATEYY").Value & "")
'        ssRecept2.Col = 10: ssRecept2.Text = Trim$(rs.Fields("SEQNUM").Value & "")
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
'    If ssRecept2.DataRowCnt > 23 Then
'        ssRecept2.MaxRows = ssRecept2.DataRowCnt + 1
'    Else
'        ssRecept2.MaxRows = 23
'    End If

    Anato_OCS_Jeobsu.MousePointer = vbDefault
    
    Exit Sub


    
End Sub


Private Sub cmdOcsOrder_Click()
    '검사대기자
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim NumQuantity         As Integer
    Dim Response            As Integer
    Dim ClassChk            As String
    
    gSFrDate = Format(dtFromDate.Value, "yyyy-MM-dd")

    Anato_OCS_Jeobsu.MousePointer = vbHourglass
    
    Call SSInitialize(ssOrder)
    
    
    LsCmdOKflag = "GEOMSA_WAIT"
    lblReceptTitle = "검 사 대 기 자 명 단"
    lblReceptTitle.BackColor = &H808000
    lblReceptTitle.ForeColor = &HFFFFFF
    vaTabPro1.ActiveTab = 1
    
    If optHistology.Value = True Then
        GiExamNumb = 85   '61  ' 91
        txtClass = "P"
        ClassChk = "80"
    ElseIf optCytology.Value = True Then
        GiExamNumb = 85   ' 61   '2  ' 92
        txtClass = "C"
        ClassChk = "89"
    Else
        GiExamNumb = 85   ' 61   '2  ' 92
        txtClass = "R"
        ClassChk = "90"
    End If
    
    txtDateYY = Dual_Date_Get("YYYY")
    

'/------------------------------------------------------------------------------------------
    Dim adoSeq      As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & " SELECT MAX(SEQNUM) MAXSEQ "
    strSQL = strSQL & " FROM   TWANAT_DIAG"
    strSQL = strSQL & " WHERE  CLASS    = '" & txtClass & "'"
    strSQL = strSQL & " AND    DATEYY   =  " & Val(txtDateYY.Text)
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
       txtSeqNum = 1
    Else
       txtSeqNum = Val(rs.Fields("MAXSEQ").Value) + 1
       AdoCloseSet rs
    End If
    
    
    'GoSub SUB_FORM_INITIALIZE_PROC
    'SUB_FORM_INITIALIZE_PROC:
    txtClass.BackColor = RGB(250, 250, 225)
    txtDateYY.BackColor = RGB(250, 250, 225)
    txtSeqNum.BackColor = RGB(250, 250, 225)
    ssItemJeobsu.Col = 1:        ssItemJeobsu.Row = 0
    ssItemJeobsu.Row2 = -1:      ssItemJeobsu.Col2 = 5
    ssItemJeobsu.BlockMode = True
    ssItemJeobsu.BackColor = RGB(250, 250, 225)
    ssItemJeobsu.BlockMode = False
    lblPtno = ""
    lblAge = ""
    lblDeptName = ""
    lblSname = ""
    lblSex = ""
    txtRoomcode = ""
    
    Call SSInitialize(ssRecept)
    Call SSInitialize(ssItemJeobsu)
    '    ssTitle = GsExamJong & " " & "Order List"
    cmdCancel.Enabled = True
    'Return
    
    LsJeobsuAdd = False

   ' 수량대로 접수환자에 색깔별로 DISPLAY되도록 수정
    strSQL = ""
    strSQL = strSQL & " SELECT  TO_CHAR(O.JEOBSUDT, 'YYYY-MM-DD') JeobsuDT,"
    strSQL = strSQL & "         TO_CHAR(O.INDATE, 'YYYY-MM-DD') INDATE, "
    strSQL = strSQL & "         O.PTNO, O.QUANTITY, P.SNAME, "
    strSQL = strSQL & "         O.ROOMCODE, O.SEX,  O.AGEYY, "
    strSQL = strSQL & "         D.DEPTNAMEK, D.DEPTCODE "
    strSQL = strSQL & "   FROM  TWEXAM_ORDER  O, "
    strSQL = strSQL & "         TWexam_itemml L, "
    strSQL = strSQL & "         TWBAS_PATIENT P, "
    strSQL = strSQL & "         TWBAS_DEPT    D  "
    strSQL = strSQL & "  WHERE  O.SLIPNO1  = '" & GiExamNumb & "' "
    strSQL = strSQL & "    AND  O.itemcd   = L.codeky(+) "
    strSQL = strSQL & "    AND  L.codegu   = '" & ClassChk & "' "
    strSQL = strSQL & "    AND  O.PTNO     = P.PTNO(+)"
    strSQL = strSQL & "    AND  O.DEPTCODE = D.DEPTCODE(+)"
    strSQL = strSQL & "    AND  ( O.JEOBSUYN  IS NULL OR  O.JEOBSUYN = ' ' )"
    
    strSQL = strSQL & "    AND  ( O.JEOBSUYN  IS NULL OR  O.JEOBSUYN  <> '#' )"
    
    strSQL = strSQL & "    AND (O.ITEMCD BETWEEN '850001' AND '851499' "
    strSQL = strSQL & "         OR O.ITEMCD BETWEEN '859001' AND '859999') "
    
    strSQL = strSQL & "    AND   O.JEOBSUDT >= TO_DATE('" & gSFrDate & "','YYYY-MM-DD') "
    
    strSQL = strSQL & "  ORDER  BY P.SNAME "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    ssRecept.MaxRows = Rowindicator + 1
    
    i = 0
    Do Until rs.EOF
        NumQuantity = 0
        NumQuantity = Val(rs.Fields("Quantity").Value & "")
        ssRecept.Row = i + 1
        
        '수량
        Select Case NumQuantity
            
            ''''''''''''''''''''''''''''''''''''''''
            Case 1: 'GoSub BLOCK_DISPLAY
                    'BLOCK_DISPLAY:
                    ssRecept.Row = ssRecept.Row
                    ssRecept.Col = 2
                    ssRecept.Row2 = ssRecept.Row
                    ssRecept.Col2 = ssRecept.MaxCols
                    ssRecept.BlockMode = True
                    ssRecept.ForeColor = RGB(0, 0, 0)
                    ssRecept.BlockMode = False
                    ssRecept.Col = 1
                    ssRecept.Lock = False
                    'Return

            Case 2: 'GoSub BLUE_DISPLAY
                    'BLUE_DISPLAY:
                    ssRecept.Row = ssRecept.Row
                    ssRecept.Col = 2
                    ssRecept.Row2 = ssRecept.Row
                    ssRecept.Col2 = ssRecept.MaxCols
                    ssRecept.BlockMode = True
                    ssRecept.ForeColor = RGB(0, 0, 255)
                    ssRecept.BlockMode = False
                    ssRecept.Col = 1: ssRecept.Lock = False
                    'Return
            Case 3
                    'GoSub RED_DISPLAY
                    'RED_DISPLAY:
                    ssRecept.Row = ssRecept.Row
                    ssRecept.Col = 2
                    ssRecept.Row2 = ssRecept.Row
                    ssRecept.Col2 = ssRecept.MaxCols
                    ssRecept.BlockMode = True
                    ssRecept.ForeColor = RGB(255, 0, 0)
                    ssRecept.BlockMode = False
                    ssRecept.Col = 1
                    ssRecept.Lock = False
                    'Return
            Case 4 To 999
                    'GoSub RED_DISPLAY
                    'yellow_DISPLAY:
                    ssRecept.Row = ssRecept.Row
                    ssRecept.Col = 2
                    ssRecept.Row2 = ssRecept.Row
                    ssRecept.Col2 = ssRecept.MaxCols
                    ssRecept.BlockMode = True
                    ssRecept.ForeColor = RGB(200, 150, 0)
                    ssRecept.BlockMode = False
                    ssRecept.Col = 1
                    ssRecept.Lock = False
                    'Return
        
        End Select
        
        'GoSub Spread_Display
        'Spread_Display:
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("JEOBSUDT").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("PTNO").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("SNAME").Value & ""
        
        If Trim(rs.Fields("ROOMCODE") & "") <> "" Then
            ssRecept.Col = 5:  ssRecept.Text = rs.Fields("ROOMCODE").Value & ""
        Else
            ssRecept.Col = 5:  ssRecept.Text = rs.Fields("DEPTCODE").Value & ""
        End If
        
        ssRecept.Col = 6:  ssRecept.Text = rs.Fields("DEPTNAMEK").Value & ""
        ssRecept.Col = 7:  ssRecept.Text = rs.Fields("SEX").Value & ""
        ssRecept.Col = 8:  ssRecept.Text = rs.Fields("AGEYY").Value & ""
        ssRecept.Col = 9:  ssRecept.Text = rs.Fields("INDATE").Value & ""
        ssRecept.Col = 11: ssRecept.Text = rs.Fields("DEPTCODE").Value & ""
        ssRecept.Col = 12
        
        ' QUANTITY_DISPLAY
        Dim StrJeobsuDt     As String
        Dim StrPtno         As String
        Dim adoDiag         As ADODB.Recordset
        
        StrJeobsuDt = rs.Fields("JEOBSUDT").Value & ""
        StrPtno = rs.Fields("PTNO").Value & ""
        
        strSQL = ""
        strSQL = strSQL & " SELECT  *  "
        strSQL = strSQL & " FROM    TWANAT_DIAG "
        strSQL = strSQL & " WHERE   PtNo   =  '" & StrPtno & "'"
        strSQL = strSQL & " AND     InDate =  TO_DATE('" & StrJeobsuDt & "','YYYY-MM-DD')"
        strSQL = strSQL & " AND     Class  =  'C '"      'cytology
        
        Result = AdoOpenSet(adoDiag, strSQL)
        
        If Result Then
            ssRecept.Text = adoDiag.RecordCount
            AdoCloseSet adoDiag
        End If
        'Return
        
        rs.MoveNext: i = i + 1
    Loop
    AdoCloseSet rs
    
'    If ssRecept.DataRowCnt > 21 Then
'        ssRecept.MaxRows = ssRecept.DataRowCnt + 1
'    Else
'        ssRecept.MaxRows = 21
'    End If
    
    Anato_OCS_Jeobsu.MousePointer = vbDefault
    
Exit Sub
    
    
End Sub


Private Sub cmdOp_Click()

'    FrmViewIlls.Show 1
'    txtOpname = FrmViewIlls.Tag
    
    FrmViewopIlls.Show 1
    txtOpname = FrmViewopIlls.Tag
    
    
    'pnlPreDiagnosis = Get_IllName(txtPreDiagnosis)

End Sub


Private Sub cmdOrderJeobsu_Click()
    '선택완료
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim LsRowID             As String
    Dim LsOrderRowID        As String
'    Dim LiOrderNo           As Long
    Dim nSeqnum             As Long
    Dim FCheck              As Boolean
    
    
    If Trim(Mid(txtOpname.Text, 1, 8)) = "" Then MsgBox "수술명을 입력하십시요.": Exit Sub
    If Trim(Mid(txtOrganPart.Text, 1, 8)) = "" Then MsgBox "장기부위명을 입력하십시요.": Exit Sub
    
    For i = 1 To ssOrder.DataRowCnt
        ssOrder.Row = i
        ssOrder.Col = 1
        If ssOrder.Text = "1" Then
            FCheck = True
        End If
    Next i
    If FCheck = False Then Exit Sub
    
    If LsJeobsuAdd = False Then
        Call SSInitialize(ssItemJeobsu)
    End If
    
    strSQL = ""
    strSQL = strSQL & " SELECT max(Seqnum) "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  CLASS  = '" & txtClass & "' "
    strSQL = strSQL & " AND    Dateyy = '" & Trim$(txtDateYY.Text) & "' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        nSeqnum = 0
        AdoCloseSet rs
    Else
        nSeqnum = Val(rs.Fields("max(seqnum)").Value & "")
        
        AdoCloseSet rs
    End If
    
    For i = 1 To ssOrder.DataRowCnt
        ssOrder.Row = i
        ssOrder.Col = 1
        If ssOrder.Text = "1" Then
            ssOrder.Row = i
            ssOrder.Text = "0"
            
            ssOrder.Col = 5:        LsOrderRowID = Trim(ssOrder.Text)
            ssOrder.Col = 7:        LiJtimeHH = Val(ssOrder.Text)
            ssOrder.Col = 8:        LiJtimeMM = Val(ssOrder.Text)
            ssOrder.Col = 9:        LiAgeMM = Val(ssOrder.Text)
            ssOrder.Col = 10:       LsGbio = Trim(ssOrder.Text)
            ssOrder.Col = 11:       LsGbEr = Trim(ssOrder.Text)
            ssOrder.Col = 12:       LsGeomchcd = Trim(ssOrder.Text)
            ssOrder.Col = 13:       LsGeomsaGu = Trim(ssOrder.Text)
            ssOrder.Col = 14:       LsOrderDt = Trim(ssOrder.Text)
            ssOrder.Col = 15:       LiOrderNo = Val(ssOrder.Text)
            ssOrder.Col = 16:       LsCmDoctor = Trim(ssOrder.Text)
            
            ssOrder.Col = 17:       LsDrCode = Trim(ssOrder.Text)
            ssOrder.Col = 18:       LsBi = Trim(ssOrder.Text)
            
            ssOrder.Col = 23:       LsOpname = Trim(Mid(txtOpname.Text, 1, 8))
            ssOrder.Col = 24:       LsOrganPart = Trim(Mid(txtOrganPart.Text, 1, 8))
'            ssOrder.Col = 25:       LsRemark1 = Trim(txtRemark1.Text)
            
            nSeqnum = nSeqnum + 1
            ssOrder.Col = 6
            Select Case ssOrder.Text
                Case "II"
                    ssItemJeobsu.Row = ssItemJeobsu.DataRowCnt + 1
                    ssItemJeobsu.Col = 1: ssItemJeobsu.Text = "1"
                    ssItemJeobsu.Col = 2:  ssOrder.Col = 2: ssItemJeobsu.Text = ssOrder.Text
                    ssItemJeobsu.Col = 3:  ssOrder.Col = 3: ssItemJeobsu.Text = ssOrder.Text
                    ssItemJeobsu.Col = 7:  ssItemJeobsu.Text = "0"
                    ssItemJeobsu.Col = 5:  ssItemJeobsu.Text = LsOrderRowID
                    ssItemJeobsu.Col = 8:  ssItemJeobsu.Text = LiOrderNo
                    ssItemJeobsu.Col = 9:  ssItemJeobsu.Text = nSeqnum
                Case "IR"
                    ssOrder.Col = 5:   LsOrderRowID = ssOrder.Text
                    ssOrder.Col = 15:  LiOrderNo = ssOrder.Text
                    GoSub SUB_ITEMML_PROC
                Case "R"
                    GoSub SUB_ROUTINE_PROC
            End Select
            'Exit For
        End If
    Next i
    
    Call SSInitialize(ssOrder)
    
    txtOpname.Text = ""
    txtOrganPart.Text = ""
    txtRemark1.Text = ""
    
    cmdJeobsuAdd.SetFocus
    
    Exit Sub
    
'/------------------------------------------------------------------------------------------
SUB_ITEMML_PROC:
    Dim adoiTem             As ADODB.Recordset
    
    ssOrder.Row = i
    ssOrder.Col = 19
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
    strSQL = strSQL & " WHERE  SUBSTR(CODEKY,1,2) = '" & GiExamNumb & "'"
    strSQL = strSQL & " AND    SUGACD             = '" & Trim(ssOrder.Text) & "'"
    strSQL = strSQL & " ORDER  BY CODEKY  ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Return
    
    Do Until rs.EOF
        If Trim(rs.Fields("GbInput").Value & "") = "1" Then
            ssItemJeobsu.Row = ssItemJeobsu.DataRowCnt + 1
            ssItemJeobsu.Col = 2:   ssItemJeobsu.Text = rs.Fields("CODEKY").Value & ""
            ssItemJeobsu.Col = 3:   ssItemJeobsu.Text = rs.Fields("ITEMNM").Value & ""
            
            If i = 0 Then
                ssItemJeobsu.Col = 7:  ssItemJeobsu.Text = "0"
                ssItemJeobsu.Col = 5:  ssItemJeobsu.Text = LsOrderRowID
            Else
                ssItemJeobsu.Col = 7:  ssItemJeobsu.Text = "1"
            End If
            
            ssItemJeobsu.Col = 5:  ssItemJeobsu.Text = LsOrderRowID
            ssItemJeobsu.Col = 8:  ssItemJeobsu.Text = LiOrderNo
            ssItemJeobsu.Col = 9:  ssItemJeobsu.Text = nSeqnum
        
        End If
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
    Return

'/------------------------------------------------------------------------------------------
SUB_ROUTINE_PROC:
    Dim adoRTN              As ADODB.Recordset
    
    ssOrder.Row = i
    ssOrder.Col = 2
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWEXAM_ROUTINE "
    strSQL = strSQL & " WHERE  ROUTINCD = '" & Trim(ssOrder.Text) & "'"
    strSQL = strSQL & " ORDER  BY CODEKY  ASC "
    
    Result = AdoOpenSet(adoRTN, strSQL)
    
    If Result = False Then Return
    Do Until adoRTN.EOF
        ssItemJeobsu.Row = ssItemJeobsu.DataRowCnt + 1
        ssItemJeobsu.Col = 2:   ssItemJeobsu.Text = adoRTN.Fields("CODEKY").Value
        ssItemJeobsu.Col = 3:   ssItemJeobsu.Text = adoRTN.Fields("iTemnm").Value
        ssOrder.Col = 5:        LsRowID = ssOrder.Text
        ssOrder.Col = 2:        ssItemJeobsu.Col = 6:  ssItemJeobsu.Text = ssOrder.Text
        ssOrder.Col = 15:       ssItemJeobsu.Col = 8:  ssItemJeobsu.Text = ssOrder.Text
                                ssItemJeobsu.Col = 9:  ssItemJeobsu.Text = nSeqnum
        If i = 0 Then
            ssItemJeobsu.Col = 5:   ssItemJeobsu.Text = LsRowID
            ssItemJeobsu.Col = 7:   ssItemJeobsu.Text = "0"
        Else
            ssItemJeobsu.Col = 7:   ssItemJeobsu.Text = "1"
        End If
        
        adoRTN.MoveNext
    Loop
    AdoCloseSet adoRTN
    Return

End Sub

Private Sub cmdOrgan_Click()
    GDict = "T"
    
    Anato_Jindan_Code.Show vbModal
    
    txtOrganPart.Text = GJindan
'    txtDiagCodeName.Text = GPJindan

End Sub

Private Sub cmdRefresh_Click()
    '미사용 '2000.01.19
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim LiRowCnt            As Integer
    Dim LsPtNo              As String
    
    Call SSInitialize(ssOrder)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSQL = strSQL & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt "
    strSQL = strSQL & " FROM   TWEXAM_ORDER a, "
    strSQL = strSQL & "        TWEXAM_ITEMML c "
    strSQL = strSQL & " WHERE  a.Ptno     =  '" & Trim(lblPtno) & "' "
    strSQL = strSQL & " AND    a.JeobsuDt =  TO_DATE( '" & LsjeobsuDt & "','yyyy-MM-dd')"
    strSQL = strSQL & " AND    a.Slipno1  = '" & GiExamNumb & "'"
    strSQL = strSQL & " AND    a.itemcd   =  c.codeky(+) "
    strSQL = strSQL & " AND    c.codegu   = '" & CodeGuchk & "'"
    strSQL = strSQL & " AND   (a.JeobsuYN = ' ' OR  a.JeobsuYN IS NULL)"
    strSQL = strSQL & " ORDER BY JEOBSUT1, JEOBSUT2, ITEMCD ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssOrder.Row = ssOrder.DataRowCnt + 1
        ssOrder.Col = 2:   ssOrder.Text = Trim(rs.Fields("iTemCD").Value & "")
                           
        GoSub SUB_EXAM_ITEM_PROC
        
        ssOrder.Col = 4:   ssOrder.Text = Format(rs.Fields("Jeobsut1").Value, "00") & ":" & _
                                          Format(rs.Fields("Jeobsut2").Value, "00")
        ssOrder.Col = 5:   ssOrder.Text = rs.Fields("RowID").Value & ""
        ssOrder.Col = 7:   ssOrder.Text = rs.Fields("JEOBSUT1").Value & ""
        ssOrder.Col = 8:   ssOrder.Text = rs.Fields("JEOBSUT2").Value & ""
        ssOrder.Col = 9:   ssOrder.Text = rs.Fields("AGEMM").Value & ""
        ssOrder.Col = 10:  ssOrder.Text = rs.Fields("GBIO").Value & ""
        ssOrder.Col = 11:  ssOrder.Text = rs.Fields("GBER").Value & ""
        ssOrder.Col = 12:  ssOrder.Text = rs.Fields("GEOMCHCD").Value & ""
        ssOrder.Col = 13:  ssOrder.Text = rs.Fields("GEOMSAGU").Value & ""
        ssOrder.Col = 14:  ssOrder.Text = rs.Fields("ORDERDT").Value & ""
        ssOrder.Col = 15:  ssOrder.Text = rs.Fields("ORDERNO").Value & ""
        ssOrder.Col = 16:  ssOrder.Text = rs.Fields("CMDOCTOR").Value & ""
                           txtRemark1.Text = ssOrder.Text
        ssOrder.Col = 17:  ssOrder.Text = rs.Fields("DRCODE").Value & ""
        ssOrder.Col = 18:  ssOrder.Text = rs.Fields("BI").Value & ""
        ssOrder.Col = 20:  ssOrder.Text = rs.Fields("GBINFO").Value & ""
        
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
    Exit Sub
    
'/------------------------------------------------------------------------------------------
SUB_EXAM_ITEM_PROC:
    Dim adoiTem             As ADODB.Recordset
    
    strSQL = ""
''    strSQL = strSQL & " SELECT Codeky, iTemnm, sugaCD "
    strSQL = strSQL & " SELECT Codeky, iTemnm, sugaCD, Gbroutine "
    strSQL = strSQL & "   FROM TWEXAM_iTemML "
    strSQL = strSQL & "  WHERE Codeky = '" & Trim(ssOrder.Text) & "'"
    
    Result = AdoOpenSet(adoiTem, strSQL)
    
    If Result = False Then
    
        Dim adoExam     As ADODB.Recordset
        
        strSQL = ""
        strSQL = strSQL & " SELECT SugaCD, RoutinNM "
        strSQL = strSQL & "   FROM TWEXAM_ROUTINE "
        strSQL = strSQL & "  WHERE RoutinCD = '" & Trim(ssOrder.Text) & "'"
        
        Result = AdoOpenSet(adoExam, strSQL)
        If Result Then
            ssOrder.Col = 3:   ssOrder.Text = adoExam.Fields("ROUTINNM").Value & ""
            ssOrder.Col = 6:   ssOrder.Text = "R"
        End If
        AdoCloseSet adoExam
    Else
        If adoiTem.Fields("Gbroutine").Value & "" = "0" Then
            ssOrder.Col = 3:   ssOrder.Text = adoiTem.Fields("ITEMNM").Value & ""
            ssOrder.Col = 6:   ssOrder.Text = "IR"
            ssOrder.Col = 19:  ssOrder.Text = adoiTem.Fields("SUGACD").Value & ""
        Else
            ssOrder.Col = 3:   ssOrder.Text = adoiTem.Fields("ITEMNM").Value & ""
            ssOrder.Col = 6:   ssOrder.Text = "II"
            ssOrder.Col = 19:  ssOrder.Text = adoiTem.Fields("SUGACD").Value & ""
        End If
    End If
    
    AdoCloseSet adoiTem
    
    Return
    
End Sub

Private Sub Form_Load()
 
    dtFromDate.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")), "yyyy-MM-dd")
    
    lblExamName = Trim(GstrPassName)
    
    txtJdate = Dual_Date_Get("YYYY-MM-DD")
'    txtClass = "P"
    vaTabPro1.ActiveTab = 0
    
    LsSex = ""
    DeptCode = ""
    LsGbio = ""
    LsBi = ""
    LsGbEr = ""
    LsGeomchcd = ""
    LsGeomsaGu = ""
    
    optHistology.Value = True
    optCytology.Value = False

End Sub


Private Sub optHistology_Click()
'    optHistology.Value = True
    CodeGuchk = "80"  '1
    
End Sub

Private Sub optCytology_Click()
'    optCytology.Value = True
    CodeGuchk = "89"  '2

End Sub

Private Sub optRefferal_Click()
'    optRefferal.Value = True
    CodeGuchk = "90" '3

End Sub

Private Sub optSpecial_Click()
'''    optSpecial.Value = True
'''    CodeGuchk = "81"

End Sub

Private Sub ssItemJeobsu_Click(ByVal Col As Long, ByVal Row As Long)
    If Col >= 2 Then Exit Sub
    ssItemJeobsu.Col = 4
    ssItemJeobsu.Row = Row
    
    txtSeqNum.Text = ssItemJeobsu.Text

End Sub


Private Sub ssorder_Click(ByVal Col As Long, ByVal Row As Long)
    
'    lblRemark = ""              'Check
'    lblGbInfo = ""
'    lblGeomchName = ""          'Check
    
    If Row > ssOrder.DataRowCnt Then Exit Sub
    
    ssOrder.Col = 2:           ssOrder.Col2 = ssOrder.MaxCols
    ssOrder.Row = LiOldRow:    ssOrder.Row2 = LiOldRow
    ssOrder.BlockMode = True
    ssOrder.BackColor = RGB(235, 245, 235)
    ssOrder.ForeColor = RGB(0, 0, 0)
    ssOrder.BlockMode = False
    
    ssOrder.Col = 2:           ssOrder.Col2 = ssOrder.MaxCols
    ssOrder.Row = Row:         ssOrder.Row2 = Row
    ssOrder.BlockMode = True
    ssOrder.BackColor = RGB(0, 0, 128)
    ssOrder.ForeColor = RGB(255, 255, 255)
    ssOrder.BlockMode = False
    LiOldRow = Row
    
    ssOrder.Row = Row
    ssOrder.Col = 2  '20
'    lblGbInfo = ssOrder.Text
    
End Sub


Private Sub ssorder_LostFocus()

    ssOrder.Col = 2:           ssOrder.Col2 = ssOrder.MaxCols
    ssOrder.Row = LiOldRow:    ssOrder.Row2 = LiOldRow
    ssOrder.BlockMode = True
    ssOrder.BackColor = RGB(235, 245, 235)
    ssOrder.ForeColor = RGB(0, 0, 0)
    ssOrder.BlockMode = False

End Sub



Private Sub ssRecept_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    '검사대기자 명단
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim LiRowCnt            As Integer
    Dim LiDateyy            As Integer
    Dim LsPtNo              As String
    Dim LsSlipNo2           As String
    Dim LsRowID             As String
    
    LsDeptCode = ""
    txtOpname.Text = ""
    txtOrganPart.Text = ""
    txtRemark1.Text = ""
    
    
    If LsCmdOKflag = "GEOMSA_WAIT" Then
        GoSub SUB_GEOMSA_WAIT_PROC
    ElseIf LsCmdOKflag = "JEOBSU_PATIENT" Then
        GoSub SUB_JEOBSU_PATIENT_PROC         '미사용
    End If
    
    txtOpname.SetFocus
    
    Exit Sub
    
       
'------------------------------------------------------------------------------------------
SUB_GEOMSA_WAIT_PROC:
    ssRecept.Row = Row
    ssRecept.Col = 2:   LsjeobsuDt = Format(ssRecept.Text, "YYYY-MM-DD")
' TWEXMA_DIAG의 INDATE에 저장하기 위해
    
    ssRecept.Col = 9:   LsInDate = ssRecept.Text
    
    ssRecept.Col = 3:   LsPtNo = ssRecept.Text:     lblPtno = ssRecept.Text
    ssRecept.Col = 4:   lblSname = ssRecept.Text
    ssRecept.Col = 5:   txtRoomcode = ssRecept.Text
'    ssRecept.Col = 6:   lblDeptName = ssRecept.Text
    ssRecept.Col = 7:   lblSex = ssRecept.Text:     LsSex = ssRecept.Text
    ssRecept.Col = 8:   lblAge = ssRecept.Text:     LiAgeYY = Val(ssRecept.Text)
    ssRecept.Col = 11:  LsDeptCode = Trim(ssRecept.Text)
    
    
    vaTabPro1.ActiveTab = 0
    Call SSInitialize(ssOrder)
    
    strSQL = ""
    strSQL = strSQL & " SELECT  a.*, a.RowID,"
    strSQL = strSQL & "         TO_CHAR(a.Jeobsudt, 'YYYY-MM-DD') JeobsuDt,"
    strSQL = strSQL & "         TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt, "
    strSQL = strSQL & "         b.Deptnamek"
    strSQL = strSQL & " FROM    TWEXAM_ORDER a, "
    strSQL = strSQL & "         TWBAS_DEPT   b, "
    strSQL = strSQL & "         TWEXAM_ITEMML  c "
    strSQL = strSQL & " WHERE   a.Ptno     =  '" & LsPtNo & "'"
    strSQL = strSQL & " AND     a.JeobsuDt =  TO_DATE('" & LsjeobsuDt & "','YYYY-MM-DD')"
''    strSQL = strSQL & " AND     a.inDate =  TO_DATE('" & LsInDate & "','YYYY-MM-DD')"
    strSQL = strSQL & " AND     a.Slipno1  =  " & GiExamNumb
    strSQL = strSQL & " AND     a.itemcd   =  c.codeky(+) "
''    strSQL = strSQL & " AND     c.codegu   =  '" & CodeGuchk & "'"
    strSQL = strSQL & " AND    (a.JeobsuYN =  ' ' OR JeobsuYN IS NULL)"
    strSQL = strSQL & " AND     a.Deptcode =  b.Deptcode(+)"
    strSQL = strSQL & " ORDER BY JEOBSUT1, JEOBSUT2, ITEMCD ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    If Result = False Then Return
    
    ssOrder.ReDraw = False
    Do Until rs.EOF
        ssOrder.Row = ssOrder.DataRowCnt + 1
        ssOrder.Col = 2:  ssOrder.Text = rs.Fields("ITEMCD").Value & ""
           
           'SUB_EXAM_ITEM_PROC
            Dim adoCode     As ADODB.Recordset
            
            strSQL = ""
            strSQL = strSQL & " SELECT Codeky, iTemnm, SugaCD, GbRoutine "
            strSQL = strSQL & "   From TWEXAM_iTemML "
            strSQL = strSQL & "  WHERE Codeky = '" & ssOrder.Text & "'"
            
            Result = AdoOpenSet(adoCode, strSQL)
            
            If Result = False Then
                strSQL = " SELECT SugaCD, Routinnm FROM TWEXAM_Routine WHERE RoutinCD = '" & ssOrder.Text & "'"
                If AdoOpenSet(adoCode, strSQL) Then
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ROUTINNM").Value & ""
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''spread 1cell에 복수 line 표시
'''''''''''''''''''''''''''''''''''''''''''''''''''속성에서 multi lines선택
                    ssOrder.RowHeight(ssOrder.Row) = ssOrder.MaxTextRowHeight(ssOrder.Row)
                    
                    ssOrder.Col = 6:   ssOrder.Text = "R"
                    AdoCloseSet adoCode
                End If
            Else
                If adoCode.Fields("GBROUTINE").Value & "" = "0" Then
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ITEMNM").Value & ""
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''
                    ssOrder.RowHeight(ssOrder.Row) = ssOrder.MaxTextRowHeight(ssOrder.Row)
                    
                    ssOrder.Col = 6:   ssOrder.Text = "IR"
                    ssOrder.Col = 19:  ssOrder.Text = adoCode.Fields("SUGACD").Value & ""
                Else
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ITEMNM").Value & ""

'''''''''''''''''''''''''''''''''''''''''''''''''''
                    ssOrder.RowHeight(ssOrder.Row) = ssOrder.MaxTextRowHeight(ssOrder.Row)
                    
                    ssOrder.Col = 6:   ssOrder.Text = "II"
                    ssOrder.Col = 19:  ssOrder.Text = adoCode.Fields("SUGACD").Value & ""
                End If
                AdoCloseSet adoCode
            End If
'            Return
        
        ssOrder.Col = 4:  ssOrder.Text = Format(rs.Fields("JEOBSUT1").Value, "00")
                          ssOrder.Text = ssOrder.Text & ":" & Format(rs.Fields("JEOBSUT2").Value, "00")
        ssOrder.Col = 5:  ssOrder.Text = rs.Fields("ROWID").Value & ""
        ssOrder.Col = 7:  ssOrder.Text = rs.Fields("JEOBSUT1").Value & ""
        ssOrder.Col = 8:  ssOrder.Text = rs.Fields("JEOBSUT2").Value & ""
        ssOrder.Col = 9:  ssOrder.Text = rs.Fields("AGEMM").Value & ""
        ssOrder.Col = 10: ssOrder.Text = rs.Fields("GBIO").Value & ""
        ssOrder.Col = 11: ssOrder.Text = rs.Fields("GBER").Value & ""
        ssOrder.Col = 12: ssOrder.Text = rs.Fields("GEOMCHCD").Value & ""
        ssOrder.Col = 13: ssOrder.Text = rs.Fields("GEOMSAGU").Value & ""
        ssOrder.Col = 14: ssOrder.Text = rs.Fields("ORDERDT").Value & ""
        ssOrder.Col = 15: ssOrder.Text = rs.Fields("ORDERNO").Value & ""
        ssOrder.Col = 16: ssOrder.Text = rs.Fields("CMDOCTOR").Value & ""
                          txtRemark1.Text = ssOrder.Text
        ssOrder.Col = 17: ssOrder.Text = rs.Fields("DRCODE").Value & ""
        ssOrder.Col = 18: ssOrder.Text = rs.Fields("BI").Value & ""
        ssOrder.Col = 20: ssOrder.Text = rs.Fields("GBINFO").Value & ""
        ssOrder.Col = 21: ssOrder.Text = rs.Fields("DEPTCODE").Value & ""
        ssOrder.Col = 22: ssOrder.Text = rs.Fields("Deptnamek").Value & ""
                                         lblDeptName = rs.Fields("Deptnamek").Value & ""
        rs.MoveNext
    Loop
    ssOrder.ReDraw = True
    
    AdoCloseSet rs
    Return
        

'---------------------------------------------------------------------------------------
'- 미사용
'---------------------------------------------------------------------------------------
SUB_JEOBSU_PATIENT_PROC:
    ssRecept.Row = Row
    ssRecept.Col = 2:   LsjeobsuDt = ssRecept.Text
' TWEXMA_DIAG의 INDATE에 저장하기 위해
    ssRecept.Col = 2:   LsInDate = ssRecept.Text
    ssRecept.Col = 3:   LsPtNo = ssRecept.Text:     lblPtno = ssRecept.Text
    ssRecept.Col = 4:   lblSname = ssRecept.Text
    ssRecept.Col = 5:   txtRoomcode = ssRecept.Text
    ssRecept.Col = 6:   lblDeptName = ssRecept.Text
    ssRecept.Col = 7:   lblSex = ssRecept.Text:     LsSex = ssRecept.Text
    ssRecept.Col = 8:   lblAge = ssRecept.Text:     LiAgeYY = Val(ssRecept.Text)
    ssRecept.Col = 9:   txtDateYY = ssRecept.Text
    ssRecept.Col = 10:  txtSeqNum = ssRecept.Text
    

    vaTabPro1.ActiveTab = 0
    
    Call SSInitialize(ssItemJeobsu)
    
    strSQL = ""
    strSQL = strSQL & " SELECT  a.ITEMCD, a.GBRESULT, a.ORDERNO, a.ROWID, b.ItemNM "
    strSQL = strSQL & " FROM    TWANAT_DIAG   a,"
    strSQL = strSQL & "         TWEXAM_ITEMML b "
    strSQL = strSQL & " WHERE   a.GBRESULT   = '0'"
    strSQL = strSQL & " AND     a.DATEYY   = '" & txtDateYY & "' "
    strSQL = strSQL & " AND     a.CLASS    = '" & txtClass & "' "                'Histology
    strSQL = strSQL & " AND     a.PTNO     = '" & LsPtNo & "' "
    strSQL = strSQL & " AND     a.ItemCD   = b.Codeky(+)"
 
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssItemJeobsu.Row = ssItemJeobsu.DataRowCnt + 1
        ssItemJeobsu.Col = 2:  ssItemJeobsu.Text = rs.Fields("ITEMCD").Value & ""
        ssItemJeobsu.Col = 3:  ssItemJeobsu.Text = rs.Fields("iTemNM").Value & ""
'''        ssItemJeobsu.Col = 4:  ssItemJeobsu.Text = rs.Fields("Verify").Value & ""
        ssItemJeobsu.Col = 5:  ssItemJeobsu.Text = rs.Fields("RowID").Value & ""
'''        ssItemJeobsu.Col = 6:  ssItemJeobsu.Text = rs.Fields("RoutinCD").Value & ""
        ssItemJeobsu.Col = 8:  ssItemJeobsu.Text = rs.Fields("ORDERNO").Value & ""
        
        rs.MoveNext
    Loop
    AdoCloseSet rs
 
    Return

    
End Sub


Private Sub ssRecept2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    '접수환자명단
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim LiRowCnt            As Integer
    Dim LiDateyy            As Integer
    Dim LsPtNo              As String
    Dim LsSlipNo2           As String
    Dim LsRowID             As String
    
    LsDeptCode = ""
    
    If LsCmdOKflag = "GEOMSA_WAIT" Then
        GoSub SUB_GEOMSA_WAIT_PROC
    ElseIf LsCmdOKflag = "JEOBSU_PATIENT" Then
        GoSub SUB_JEOBSU_PATIENT_PROC
    End If
    
    Exit Sub
    
'------------------------------------------------------------------------------------------
'- 미사용
'------------------------------------------------------------------------------------------
SUB_GEOMSA_WAIT_PROC:
    ssRecept2.Row = Row
    ssRecept2.Col = 2:   LsjeobsuDt = Format(ssRecept2.Text, "YYYY-MM-DD")
' TWEXMA_DIAG의 INDATE에 저장하기 위해
    ssRecept2.Col = 2:   LsInDate = ssRecept2.Text
    ssRecept2.Col = 3:   LsPtNo = ssRecept2.Text:     lblPtno = ssRecept2.Text
    ssRecept2.Col = 4:   lblSname = ssRecept2.Text
    ssRecept2.Col = 5:   txtRoomcode = ssRecept2.Text
'    ssRecept2.Col = 6:   lblDeptName = ssRecept2.Text
    ssRecept2.Col = 7:   lblSex = ssRecept2.Text:     LsSex = ssRecept2.Text
    ssRecept2.Col = 8:   lblAge = ssRecept2.Text:     LiAgeYY = Val(ssRecept2.Text)
    ssRecept2.Col = 11:  LsDeptCode = Trim(ssRecept2.Text)
    
    vaTabPro1.ActiveTab = 0
    Call SSInitialize(ssOrder)
    
    strSQL = ""
    strSQL = strSQL & " SELECT  a.*, a.RowID,"
    strSQL = strSQL & "         TO_CHAR(a.Jeobsudt, 'YYYY-MM-DD') JeobsuDt,"
    strSQL = strSQL & "         TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt, "
    strSQL = strSQL & "         b.Deptnamek"
    strSQL = strSQL & " FROM    TWEXAM_ORDER a, "
    strSQL = strSQL & "         TWBAS_DEPT   b, "
    strSQL = strSQL & "         TWEXAM_ITEMML  c "
    strSQL = strSQL & " WHERE   a.Ptno     =  '" & LsPtNo & "'"
    strSQL = strSQL & " AND     a.JeobsuDt =       TO_DATE('" & LsjeobsuDt & "','YYYY-MM-DD')"
    strSQL = strSQL & " AND     a.Slipno1  =  " & GiExamNumb
    strSQL = strSQL & " AND     a.itemcd   =  c.codeky "
    strSQL = strSQL & " AND     c.codegu   =  '" & CodeGuchk & "' "
    strSQL = strSQL & " AND    (a.JeobsuYN =  ' ' OR JeobsuYN IS NULL) "
    strSQL = strSQL & " AND     a.Deptcode =  b.Deptcode(+)"
    strSQL = strSQL & " ORDER BY JEOBSUT1, JEOBSUT2, ITEMCD ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    If Result = False Then Return
    
    Do Until rs.EOF
        ssOrder.Row = ssOrder.DataRowCnt + 1
        ssOrder.Col = 2:  ssOrder.Text = rs.Fields("ITEMCD").Value & ""
'        GoSub SUB_EXAM_ITEM_PROC
           'SUB_EXAM_ITEM_PROC:
            Dim adoCode     As ADODB.Recordset
            
            strSQL = ""
            strSQL = strSQL & " SELECT Codeky, iTemnm, SugaCD, GbRoutine "
            strSQL = strSQL & "   From TWEXAM_iTemML "
            strSQL = strSQL & "  WHERE Codeky = '" & ssOrder.Text & "'"
            
            Result = AdoOpenSet(adoCode, strSQL)
            
            If Result = False Then
                strSQL = " SELECT SugaCD, Routinnm FROM TWEXAM_Routine WHERE RoutinCD = '" & ssOrder.Text & "'"
                If AdoOpenSet(adoCode, strSQL) Then
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ROUTINNM").Value & ""
                    ssOrder.Col = 6:   ssOrder.Text = "R"
                    AdoCloseSet adoCode
                End If
            Else
                If adoCode.Fields("GBROUTINE").Value & "" = "0" Then
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ITEMNM").Value & ""
                    ssOrder.Col = 6:   ssOrder.Text = "IR"
                    ssOrder.Col = 19:  ssOrder.Text = adoCode.Fields("SUGACD").Value & ""
                Else
                    ssOrder.Col = 3:   ssOrder.Text = adoCode.Fields("ITEMNM").Value & ""
                    ssOrder.Col = 6:   ssOrder.Text = "II"
                    ssOrder.Col = 19:  ssOrder.Text = adoCode.Fields("SUGACD").Value & ""
                End If
                AdoCloseSet adoCode
            End If
            
            'Return
        
        ssOrder.Col = 4:  ssOrder.Text = Format(rs.Fields("JEOBSUT1").Value, "00")
                          ssOrder.Text = ssOrder.Text & ":" & Format(rs.Fields("JEOBSUT2").Value, "00")
        ssOrder.Col = 5:  ssOrder.Text = rs.Fields("ROWID").Value & ""
        ssOrder.Col = 7:  ssOrder.Text = rs.Fields("JEOBSUT1").Value & ""
        ssOrder.Col = 8:  ssOrder.Text = rs.Fields("JEOBSUT2").Value & ""
        ssOrder.Col = 9:  ssOrder.Text = rs.Fields("AGEMM").Value & ""
        ssOrder.Col = 10: ssOrder.Text = rs.Fields("GBIO").Value & ""
        ssOrder.Col = 11: ssOrder.Text = rs.Fields("GBER").Value & ""
        ssOrder.Col = 12: ssOrder.Text = rs.Fields("GEOMCHCD").Value & ""
        ssOrder.Col = 13: ssOrder.Text = rs.Fields("GEOMSAGU").Value & ""
        ssOrder.Col = 14: ssOrder.Text = rs.Fields("ORDERDT").Value & ""
        ssOrder.Col = 15: ssOrder.Text = rs.Fields("ORDERNO").Value & ""
        ssOrder.Col = 16: ssOrder.Text = rs.Fields("CMDOCTOR").Value & ""
                          txtRemark1.Text = ssOrder.Text
        ssOrder.Col = 17: ssOrder.Text = rs.Fields("DRCODE").Value & ""
        ssOrder.Col = 18: ssOrder.Text = rs.Fields("BI").Value & ""
        ssOrder.Col = 20: ssOrder.Text = rs.Fields("GBINFO").Value & ""
        ssOrder.Col = 21: ssOrder.Text = rs.Fields("DEPTCODE").Value & ""
        ssOrder.Col = 22: ssOrder.Text = rs.Fields("Deptnamek").Value & ""
                          lblDeptName = rs.Fields("Deptnamek").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    Return

'---------------------------------------------------------------------------------------
SUB_JEOBSU_PATIENT_PROC:
    ssRecept2.Row = Row
    ssRecept2.Col = 2:   LsjeobsuDt = ssRecept2.Text
' TWEXMA_DIAG의 INDATE에 저장하기 위해
    ssRecept2.Col = 2:   LsInDate = ssRecept2.Text
    ssRecept2.Col = 3:   LsPtNo = ssRecept2.Text:     lblPtno = ssRecept2.Text
    ssRecept2.Col = 4:   lblSname = ssRecept2.Text
    ssRecept2.Col = 5:   txtRoomcode = ssRecept2.Text
    ssRecept2.Col = 6:   lblDeptName = ssRecept2.Text
    ssRecept2.Col = 7:   lblSex = ssRecept2.Text:     LsSex = ssRecept2.Text
    ssRecept2.Col = 8:   lblAge = ssRecept2.Text:     LiAgeYY = Val(ssRecept2.Text)
    ssRecept2.Col = 9:   txtDateYY = ssRecept2.Text
    ssRecept2.Col = 10:  txtSeqNum = ssRecept2.Text
    
    vaTabPro1.ActiveTab = 0
    
    Call SSInitialize(ssItemJeobsu)

    strSQL = ""
    strSQL = strSQL & " SELECT  a.ITEMCD, a.GBRESULT, a.ORDERNO, a.seqnum, a.ROWID, b.ItemNM "
    strSQL = strSQL & " FROM    TWANAT_DIAG   a,"
    strSQL = strSQL & "         TWEXAM_ITEMML b "
    strSQL = strSQL & " WHERE   a.GBRESULT   = '0'"
    strSQL = strSQL & " AND     a.DATEYY   = '" & txtDateYY & "' "
    strSQL = strSQL & " AND     a.CLASS    = '" & txtClass & "' "
    strSQL = strSQL & " AND     a.PTNO     = '" & LsPtNo & "'"
    strSQL = strSQL & " AND     a.ItemCD   = b.Codeky(+)"
 
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssItemJeobsu.Row = ssItemJeobsu.DataRowCnt + 1
        ssItemJeobsu.Col = 2:  ssItemJeobsu.Text = rs.Fields("ITEMCD").Value & ""
        ssItemJeobsu.Col = 3:  ssItemJeobsu.Text = rs.Fields("iTemNM").Value & ""
'''        ssItemJeobsu.Col = 4:  ssItemJeobsu.Text = rs.Fields("Verify").Value & ""
        ssItemJeobsu.Col = 4:  ssItemJeobsu.Text = rs.Fields("SeqNum").Value & ""
        If txtSeqNum.Text = "" Then
            txtSeqNum.Text = ssItemJeobsu.Text
        End If
        
        ssItemJeobsu.Col = 5:  ssItemJeobsu.Text = rs.Fields("RowID").Value & ""
        ssItemJeobsu.ColHidden = True '@@@@@@ hidden

'''        ssItemJeobsu.Col = 6:  ssItemJeobsu.Text = rs.Fields("RoutinCD").Value & ""
        ssItemJeobsu.Col = 8:  ssItemJeobsu.Text = rs.Fields("ORDERNO").Value & ""
        
        rs.MoveNext
    Loop
    AdoCloseSet rs
 
    Return

End Sub


Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    SendKeys "{tab}"

End Sub


Private Sub txtDATEYY_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    SendKeys "{tab}"

End Sub


Private Sub txtOpname_GotFocus()
    txtOpname.SelStart = 0
    txtOpname.SelLength = Len(txtOpname.Text)

End Sub

Private Sub txtOpname_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If

End Sub


Private Sub txtOrganPart_GotFocus()
    txtOrganPart.SelStart = 0
    txtOrganPart.SelLength = Len(txtOrganPart.Text)

End Sub

Private Sub txtOrganPart_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtOrganPart_LostFocus()
   txtOrganPart.Text = UCase(txtOrganPart.Text)

End Sub

Private Sub vaTabPro1_TabActivate(TabToActivate As Integer)
    If TabToActivate = 0 Then
        'clear
'        lblPtno = ""
'        lblAge = ""
'        lblDeptName = ""
'        lblSname = ""
'        lblSex = ""
'        txtRoomcode = ""
        
    ElseIf TabToActivate = 1 Then
        Call cmdOcsOrder_Click
    
    ElseIf TabToActivate = 2 Then
        Call cmdJeobSuList_Click
    
    End If


End Sub
