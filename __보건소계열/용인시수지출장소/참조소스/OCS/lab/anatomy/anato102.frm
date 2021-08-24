VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Menual_Jeobsu 
   BorderStyle     =   0  '없음
   Caption         =   "접수 처리"
   ClientHeight    =   8760
   ClientLeft      =   315
   ClientTop       =   1980
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8760
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   7575
      Left            =   9975
      ScaleHeight     =   7515
      ScaleWidth      =   1635
      TabIndex        =   32
      Top             =   765
      Width           =   1695
      Begin Threed.SSCommand cmdSelect 
         Height          =   945
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         Picture         =   "ANATO102.frx":0000
      End
      Begin Threed.SSCommand cmdErase 
         Height          =   945
         Left            =   0
         TabIndex        =   43
         Top             =   1890
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "취 소"
         ForeColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         Picture         =   "ANATO102.frx":0452
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   945
         Left            =   0
         TabIndex        =   19
         Top             =   3780
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO102.frx":076C
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   945
         Left            =   0
         TabIndex        =   17
         Top             =   945
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "접수/수정"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO102.frx":0A86
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   945
         Left            =   0
         TabIndex        =   18
         Top             =   2835
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "삭 제"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO102.frx":0DA0
      End
   End
   Begin VB.ListBox lstDisp 
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
      Height          =   7275
      Left            =   5610
      TabIndex        =   31
      Top             =   1095
      Width           =   4245
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   7620
      Left            =   120
      ScaleHeight     =   7560
      ScaleWidth      =   5265
      TabIndex        =   24
      Top             =   765
      Width           =   5325
      Begin VB.TextBox txtindate 
         BackColor       =   &H00C0C0FF&
         Height          =   345
         Left            =   1350
         TabIndex        =   14
         Top             =   6300
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpNameH 
         Caption         =   "H"
         Height          =   300
         Left            =   4530
         TabIndex        =   49
         Top             =   1890
         Width           =   225
      End
      Begin VB.TextBox txtOpName 
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1350
         TabIndex        =   48
         Top             =   1890
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Height          =   468
         Left            =   1368
         TabIndex        =   44
         Top             =   432
         Width           =   3525
         Begin VB.OptionButton optClass 
            Caption         =   "Refferal"
            Height          =   252
            Index           =   2
            Left            =   2340
            TabIndex        =   52
            Top             =   180
            Width           =   1125
         End
         Begin VB.OptionButton optClass 
            Caption         =   "Histology"
            Height          =   252
            Index           =   0
            Left            =   90
            TabIndex        =   46
            Top             =   168
            Width           =   1125
         End
         Begin VB.OptionButton optClass 
            Caption         =   "Cytology"
            Height          =   252
            Index           =   1
            Left            =   1260
            TabIndex        =   45
            Top             =   168
            Width           =   1125
         End
      End
      Begin VB.TextBox txtAgeMM 
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
         Height          =   300
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   10
         Top             =   4500
         Width           =   850
      End
      Begin VB.ComboBox cboGeomch 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   2250
         Width           =   3405
      End
      Begin VB.ComboBox cboItem 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   1500
         Width           =   3405
      End
      Begin VB.TextBox txtOrderDt 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   16
         Top             =   7170
         Width           =   1800
      End
      Begin VB.ComboBox cboSex 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   8
         Top             =   4020
         Width           =   1800
      End
      Begin VB.TextBox txtJdate 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   15
         Top             =   6750
         Width           =   1800
      End
      Begin VB.ComboBox cboDoctor 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   5940
         Width           =   3405
      End
      Begin VB.ComboBox cboDept 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   5460
         Width           =   3405
      End
      Begin VB.ComboBox cboGbio 
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
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   3540
         Width           =   1800
      End
      Begin VB.TextBox txtRoomCode 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   4
         TabIndex        =   11
         Top             =   4980
         Width           =   1800
      End
      Begin VB.TextBox txtSeqNum 
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
         Height          =   300
         Left            =   2955
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1020
         Width           =   1780
      End
      Begin VB.TextBox txtDateYY 
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
         Height          =   300
         Left            =   2145
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1020
         Width           =   600
      End
      Begin VB.TextBox txtPtNo 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   5
         Top             =   2730
         Width           =   1800
      End
      Begin VB.TextBox txtClass 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   0
         Top             =   1020
         Width           =   600
      End
      Begin VB.TextBox txtName 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3060
         Width           =   1800
      End
      Begin VB.TextBox txtAge 
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
         Height          =   300
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   9
         Top             =   4500
         Width           =   850
      End
      Begin VB.Label Label18 
         Caption         =   "내원일"
         Height          =   195
         Left            =   690
         TabIndex        =   51
         Top             =   6330
         Width           =   585
      End
      Begin VB.Label Label17 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "수술명"
         Height          =   195
         Left            =   30
         TabIndex        =   47
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label16 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   435
         TabIndex        =   42
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label15 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "검사명코드"
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
         Left            =   315
         TabIndex        =   41
         Top             =   1575
         Width           =   900
      End
      Begin VB.Label Label9 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ORDER일자"
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
         Left            =   390
         TabIndex        =   40
         Top             =   7230
         Width           =   870
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   495
         TabIndex        =   39
         Top             =   6780
         Width           =   720
      End
      Begin VB.Label Label14 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "병    실"
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
         Left            =   495
         TabIndex        =   38
         Top             =   5040
         Width           =   720
      End
      Begin VB.Label Label13 
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
         Left            =   495
         TabIndex        =   37
         Top             =   6000
         Width           =   720
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   495
         TabIndex        =   36
         Top             =   5535
         Width           =   720
      End
      Begin VB.Label Label11 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "내원구분"
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
         Left            =   495
         TabIndex        =   35
         Top             =   3615
         Width           =   720
      End
      Begin VB.Label Label10 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "검사종류"
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
         Left            =   495
         TabIndex        =   34
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "환자번호"
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
         Left            =   495
         TabIndex        =   30
         Top             =   2775
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "병리번호"
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
         Left            =   495
         TabIndex        =   29
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "성    별"
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
         Left            =   495
         TabIndex        =   28
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label Label5 
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
         Left            =   495
         TabIndex        =   27
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   495
         TabIndex        =   26
         Top             =   4560
         Width           =   720
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "환 자 정 보 등 록"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   25
         Top             =   30
         Width           =   5280
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   675
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12060
      _Version        =   65536
      _ExtentX        =   21273
      _ExtentY        =   1191
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
         Height          =   405
         Left            =   9660
         ScaleHeight     =   345
         ScaleWidth      =   2115
         TabIndex        =   21
         Top             =   120
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
            TabIndex        =   23
            Top             =   90
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
            TabIndex        =   22
            Top             =   90
            Width           =   450
         End
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "환 자 접 수 현 황"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5610
      TabIndex        =   33
      Top             =   795
      Width           =   4245
   End
End
Attribute VB_Name = "Anato_Menual_Jeobsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Dim LsGbio(0 To 10)         As String * 1
' Dim LsSex(0 To 2)           As String * 1
' Dim LsDeptCode(0 To 200)    As String * 4
' Dim LsDrCode(0 To 500)      As String * 6
' Dim LsItem(0 To 50)         As String * 6
' Dim LsGeomch(0 To 100)      As String * 8
 
    Dim LsGbio(-1 To 10)     As String * 1
    Dim LsSex(-1 To 2)        As String * 1
    Dim LsDeptCode(-1 To 200) As String * 4
    Dim LsDrCode(-1 To 500)   As String * 6
    Dim LsItem(-1 To 400)      As String * 6
    Dim LsGeomch(-1 To 700)   As String * 8
 
 '주의 memory check
    Dim LsMetrix1(-1 To 200) As String
    Dim LsMetrix2(-1 To 200) As String
    Dim LsMetrix3(-1 To 200) As String
  
    Dim TempGbio            As String
    Dim TempSex             As String
    Dim TempDeptCode        As String
    Dim TempItem            As String
    Dim TempGeomch          As String
    Dim LiIdNoFlg           As Integer      '환자 Db Read Check


Private Sub cmdErase_Click()
    
    Dim Response

    Response = MsgBox(" 입력된 DATA를 지우시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "접수취소")
    
    If Response = vbYes Then
        Call Item_Clear
    End If

End Sub


Private Sub cmdOpNameH_Click()
    
    FrmViewopIlls.Show 1
    txtOpname.Text = FrmViewopIlls.Tag
    

End Sub


Sub IDNO_SELECT()

    Dim rs                  As ADODB.Recordset
    
    Dim LsDrTemp            As String
    Dim i                   As Integer
    Dim sYYYY               As String
    Dim BYYYY               As String
    Dim SMM                 As String
    Dim BMM                 As String
    Dim AgeYY               As Integer
    Dim AgeMM               As Integer
    
    LiIdNoFlg = 0
    
    sYYYY = Trim(Dual_Date_Get("YYYY"))
    SMM = Trim(Dual_Date_Get("MM"))
    
    
'환자Date Base Select ------------------------------------------------------------
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.*, TO_CHAR(A.BIRTHDATE,'YYYYMM') BDAY "
    strSQL = strSQL & " FROM   TWBAS_PATIENT A "
    strSQL = strSQL & " WHERE  PtNo =  '" & txtPtNo & "' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtName.Enabled = True
        strSQL = ""
        strSQL = strSQL & " SELECT A.*, TO_CHAR(A.BIRTHDAY,'YYYYMM') BDAY "
        strSQL = strSQL & " FROM   TWEXAM_IDNOMST A "
        strSQL = strSQL & " WHERE  PtNo =  '" & txtPtNo & "' "
        
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Call Item_Clear
            txtPtNo.SetFocus
            MsgBox "등록된 환자가 아닙니다."
            
            Exit Sub
        End If
        
        BYYYY = Trim$(Left$(rs.Fields("BDAY").Value, 4))
        BMM = Trim$(Right$(rs.Fields("BDAY").Value, 2))
        
        TempGbio = rs.Fields("Gbio").Value & ""
        For i = 0 To 5
            If Trim(TempGbio) = Trim(LsGbio(i)) Then
               cboGbio.ListIndex = i
               cboGbio.Refresh
               Exit For
            End If
        Next i
        
        AgeYY = sYYYY - BYYYY
        AgeMM = SMM - BMM
        If AgeMM < 0 Then
            AgeYY = AgeYY - 1
            AgeMM = 12 - AgeMM
        End If
        
'        txtAge = AgeYY

        txtRoomcode = rs.Fields("RoomCode").Value & ""
    Else
        BYYYY = Trim$(Left$(rs.Fields("BDAY").Value, 4))
        BMM = Trim$(Right$(rs.Fields("BDAY").Value, 2))
        
        AgeYY = sYYYY - BYYYY
        AgeMM = SMM - BMM
        If AgeMM < 0 Then
            AgeYY = AgeYY - 1
            AgeMM = 12 - AgeMM
        End If
        
        txtAge = AgeYY
        LiIdNoFlg = 1
        txtName.Enabled = False
        
'        cboGbio.ListIndex = 3
    
    End If

    txtName = rs.Fields("Sname").Value & ""
    TempSex = rs.Fields("Sex").Value & ""
    For i = 0 To 1
        If Trim(TempSex) = Trim(LsSex(i)) Then
           cboSex.ListIndex = i
           cboSex.Refresh
           Exit For
        End If
    Next i
              
    TempDeptCode = rs.Fields("Deptcode").Value & ""
    For i = 0 To 200
        If Trim(TempDeptCode) = Trim(LsDeptCode(i)) Then
           cboDept.ListIndex = i
           cboDept.Refresh
           Exit For
        End If
    Next i
    
    LsDrTemp = rs.Fields("DrCode").Value & ""
    Call DOCTOR_SELECT
    For i = 0 To 500
        If Trim(LsDrTemp) = Trim(LsDrCode(i)) Then
           cboDoctor.ListIndex = i
           cboDoctor.Refresh
           Exit For
        End If
    Next i
         
    txtindate.Text = ""
    txtJdate = ""
    txtOrderDt = ""
    AdoCloseSet rs
     
End Sub


Sub Item_Clear()

    txtDateYY = ""
    txtSeqNum = ""
    
    cboItem.ListIndex = -1
    cboGeomch.ListIndex = -1
    txtPtNo = ""
    txtName = ""
    txtOpname.Text = ""
    
    cboGbio.ListIndex = -1
    
    cboSex.ListIndex = -1
    txtAge = ""
    txtAgeMM = ""
    txtRoomcode = ""
    cboDept.ListIndex = -1
    cboDoctor.ListIndex = -1
    
    txtindate.Text = ""
    txtJdate = ""
    txtOrderDt = ""

End Sub


Sub JEOBSU_PATIENT()

    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim LsSeqNum            As String * 8
    Dim LsPtNo              As String * 12
    Dim LsSname             As String * 20
    Dim adoPT               As ADODB.Recordset
    Dim sYYYY               As String
    Dim LsOpname
    
    lstDisp.Clear
'-------------------------------------------------'
' 환자접수현황 DISPLAY                            '
'-------------------------------------------------'
        
    sYYYY = Dual_Date_Get("yyyy")
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DIAG "
    strSQL = strSQL & " WHERE  GBRESULT =  '0' "
    strSQL = strSQL & " AND    CLASS    =  '" & txtClass & "' "
    strSQL = strSQL & " AND    DATEYY   =  '" & sYYYY & "'"
    strSQL = strSQL & " ORDER  BY CLASS, DATEYY, SEQNUM "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
'        i = i + 1
        LsSeqNum = rs.Fields("SEQNUM").Value & ""
        LsPtNo = rs.Fields("PTNO").Value & ""
        LsOpname = rs.Fields("OPNAME").Value & ""
        LsSname = rs.Fields("SNAME").Value & ""
        lstDisp.AddItem rs.Fields("CLASS").Value & "-" & _
                        rs.Fields("DATEYY").Value & "-" & _
                        LsSeqNum & LsPtNo & LsSname
        LsMetrix1(lstDisp.ListIndex) = rs.Fields("PTNO").Value & ""
        LsMetrix2(lstDisp.ListIndex) = rs.Fields("ITEMCD").Value & ""
        LsMetrix3(lstDisp.ListIndex) = rs.Fields("ORDERNO").Value & ""
        
        rs.MoveNext
    Loop
    lstDisp.ListIndex = lstDisp.ListCount - 1
    lstDisp.Refresh
    
    AdoCloseSet rs
    
End Sub

Private Sub cboDept_Click()
    
    Call DOCTOR_SELECT

End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"
    
End Sub


Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub cboGbio_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub cboGeomch_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub cboItem_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub cmdCancel_Click()
    '삭제
    Dim Response

    If txtSeqNum = "" Then Exit Sub
    
    Response = MsgBox(" 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "접수취소")
    
    If Response = vbYes Then
        
        strSQL = ""
        strSQL = strSQL & " UPDATE TWANAT_DIAG "
        strSQL = strSQL & " SET    GBRESULT = 'X' "
        strSQL = strSQL & " WHERE  CLASS   = '" & txtClass & "' "
        strSQL = strSQL & " AND    DATEYY  = '" & txtDateYY & "' "
        strSQL = strSQL & " AND    SEQNUM  = '" & txtSeqNum & "' "
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
        
        
            If LsMetrix3(lstDisp.ListIndex) <> "" Then
            
                strSQL = ""
                strSQL = strSQL & " UPDATE TWEXAM_ORDER"
                strSQL = strSQL & " SET    JeobsuYn  = ' '"                 '미접수 ' ' , 접수취소 'X'
                strSQL = strSQL & " WHERE  Ptno    = '" & LsMetrix1(lstDisp.ListIndex) & "'    "
                strSQL = strSQL & " AND    iTemCD  = '" & LsMetrix2(lstDisp.ListIndex) & "'   "
                strSQL = strSQL & " AND    Orderno = '" & LsMetrix3(lstDisp.ListIndex) & "" & "' "
                
                adoConnect.BeginTrans
                
                Result = AdoExecute(strSQL)
                
                If Result = True And Rowindicator > 0 Then
                    adoConnect.CommitTrans
                Else
                    adoConnect.RollbackTrans
                End If
           Else
                MsgBox " 수동접수한 DATA를 삭제하였습니다. ", vbCritical, "Information"
           
           End If
           
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    End If
        
'    Call JEOBSU_PATIENT
    Call Item_Clear
    
    Exit Sub

End Sub

Private Sub cmdExit_Click()
    
    Unload Me

End Sub


Private Sub cmdSave_Click()

    Dim rs                  As ADODB.Recordset
    
    Dim LsJeobsuTm
    Dim sGbGross            As String
    Dim siTemGb             As String
    Dim siTemCD             As String
    Dim sGeomChCD           As String
    Dim sGbIO               As String
    Dim sSex                As String
    Dim sAgeYY              As String
    Dim sAgeMM              As String
    Dim sRoomCode           As String
    Dim sDeptcode           As String
    Dim sDrCode             As String
    
'    On Error GoTo Error_Save
    
    If Trim(txtClass) = "" Then Exit Sub
    If Trim(txtDateYY) = "" Then Exit Sub
    If Trim(txtSeqNum) = "" Then Exit Sub

    If Trim(txtindate.Text) = "" And optClass(2).Value = False Then
        MsgBox " 내원일을 입력하십시요."
        Exit Sub
    End If
    
    If Trim(txtJdate.Text) = "" And optClass(2).Value = False Then
        MsgBox " 접수일을 입력하십시요."
        Exit Sub
    End If
    
    If Trim(txtOrderDt.Text) = "" And optClass(2).Value = False Then
        MsgBox " Order일을 입력하십시요."
        Exit Sub
    End If
    
    If Trim(txtPtNo.Text) = "" And optClass(2).Value = False Then
        MsgBox " 환자번호에 의뢰 병원명을 한글 5자 이내로 입력하십시요."
        Exit Sub
    End If
    
    If Trim(txtName.Text) = "" And optClass(2).Value = False Then
        MsgBox " 환자명을 입력하십시요."
        Exit Sub
    End If
    
    
'--------환자Master처리-------------------------------------
    
    If LiIdNoFlg <> 1 Then
        strSQL = ""
        strSQL = strSQL & " SELECT  *  "
        strSQL = strSQL & "   FROM TWEXAM_IDNOMST "
        
        strSQL = strSQL & "  WHERE PtNo =  '" & txtPtNo & "'"
        
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
           GoSub IDNOMST_INSERT
        Else
           AdoCloseSet rs
           GoSub IDNOMST_UPDATE
        End If
    End If

'-----------------------------------------------------------
    LsJeobsuTm = Format(Time, "hh:mm")

    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWANAT_DIAG"
    strSQL = strSQL & " WHERE  CLASS   = '" & txtClass & "'"
    strSQL = strSQL & " AND    DATEYY  = '" & txtDateYY & "'"
    strSQL = strSQL & " AND    SEQNUM  = '" & txtSeqNum & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        AdoCloseSet rs
        GoSub PATIENT_UPDATE
    Else
        GoSub PATIENT_INSERT
    End If
    
'    Call JEOBSU_PATIENT
'    cboClass.SetFocus
    If optClass(0).Value = True Then
        optClass(0).SetFocus
    ElseIf optClass(1).Value = True Then
        optClass(1).SetFocus
    ElseIf optClass(2).Value = True Then
        optClass(2).SetFocus
    End If
    
    Exit Sub
    
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
PATIENT_INSERT:
    
    siTemCD = IIf(cboItem.ListIndex = -1, "", Trim(LsItem(cboItem.ListIndex)))
    sGeomChCD = IIf(cboGeomch.ListIndex = -1, "", Trim(LsGeomch(cboGeomch.ListIndex)))
    sGbIO = IIf(cboGbio.ListIndex = -1, "", Trim(LsGbio(cboGbio.ListIndex)))
    sSex = IIf(cboSex.ListIndex = -1, "", Trim(LsSex(cboSex.ListIndex)))
    sAgeYY = IIf(txtAge.Text = "", "", Trim(txtAge.Text))
    sAgeMM = IIf(txtAgeMM.Text = "", "", Trim(txtAgeMM.Text))
    sRoomCode = IIf(txtRoomcode.Text = "", "", Trim(txtRoomcode.Text))
    sDeptcode = IIf(cboDept.ListIndex = -1, "", Trim(LsDeptCode(cboDept.ListIndex)))
    sDrCode = IIf(cboDoctor.ListIndex = -1, "", Trim(LsDrCode(cboDoctor.ListIndex)))
    
    strSQL = ""
'    strSQL = strSQL & " SELECT MinCham "
    strSQL = strSQL & " SELECT CodeGu "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    strSQL = strSQL & " WHERE Codeky = '" & Trim(LsItem(cboItem.ListIndex)) & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        Select Case rs.Fields("CodeGu").Value & ""
            Case "80":  siTemGb = "P"
            Case "89":  siTemGb = "C"
            Case Else: siTemGb = "R"
        End Select
        AdoCloseSet rs
    Else
        siTemGb = ""
        AdoCloseSet rs
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWANAT_DIAG "
    strSQL = strSQL & "  (GBRESULT, GBGROSS,   CLASS,     DATEYY,   SEQNUM, "
    strSQL = strSQL & "   ITEMCD,   GEOMCHCD,  PTNO,      SNAME,    GBIO, "
    strSQL = strSQL & "   SEX,      AGEYY,     AGEMM,     ROOMCODE, DEPTCODE, "
    strSQL = strSQL & "   DRCODE,   OPNAME,    ORGANPART, INDATE,   JDATE,    JTIMEHH,  "
    strSQL = strSQL & "   JTIMEMM,  ORDERDT,   ITEMGB ) "
    strSQL = strSQL & " VALUES ('0',"
    strSQL = strSQL & "         '0',"
    strSQL = strSQL & "           '" & txtClass.Text & "',"
    strSQL = strSQL & "            " & Val(txtDateYY.Text) & ","
    strSQL = strSQL & "            " & Val(txtSeqNum.Text) & ","
    strSQL = strSQL & "           '" & Trim(siTemCD) & "',"
    strSQL = strSQL & "           '" & Trim(sGeomChCD) & "',"
    strSQL = strSQL & "           '" & txtPtNo.Text & "',"
    strSQL = strSQL & "           '" & txtName.Text & "',"
    strSQL = strSQL & "           '" & sGbIO & "',"
    strSQL = strSQL & "           '" & sSex & "',"
    strSQL = strSQL & "           '" & sAgeYY & "',"
    strSQL = strSQL & "           '" & sAgeMM & "',"
    strSQL = strSQL & "           '" & sRoomCode & "',"
    strSQL = strSQL & "           '" & sDeptcode & "',"
    strSQL = strSQL & "           '" & sDrCode & "',"
    strSQL = strSQL & "           '" & Trim(txtOpname.Text) & "',"
    strSQL = strSQL & "           '" & Trim(sGeomChCD) & "',"
    strSQL = strSQL & "   TO_DATE('" & txtindate.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "   TO_DATE('" & txtJdate.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "           '" & Mid(LsJeobsuTm, 1, 2) & "',"
    strSQL = strSQL & "           '" & Mid(LsJeobsuTm, 4, 2) & "',"
    strSQL = strSQL & "   TO_DATE('" & txtOrderDt.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "           '" & siTemGb & "')"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        Call Item_Clear
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return

'----------------------------------------------------------------------------
PATIENT_UPDATE:
    siTemCD = IIf(cboItem.ListIndex = -1, "", Trim(LsItem(cboItem.ListIndex)))
    sGeomChCD = IIf(cboGeomch.ListIndex = -1, "", Trim(LsGeomch(cboGeomch.ListIndex)))
    sGbIO = IIf(cboGbio.ListIndex = -1, "", Trim(LsGbio(cboGbio.ListIndex)))
    sSex = IIf(cboSex.ListIndex = -1, "", Trim(LsSex(cboSex.ListIndex)))
    sAgeYY = IIf(txtAge.Text = "", "", Trim(txtAge.Text))
    sAgeMM = IIf(txtAgeMM.Text = "", "", Trim(txtAgeMM.Text))
    sRoomCode = IIf(txtRoomcode.Text = "", "", Trim(txtRoomcode.Text))
    sDeptcode = IIf(cboDept.ListIndex = -1, "", Trim(LsDeptCode(cboDept.ListIndex)))
    sDrCode = IIf(cboDoctor.ListIndex = -1, "", Trim(LsDrCode(cboDoctor.ListIndex)))
    
    If Trim(txtClass) = "P" Then sGbGross = "0"
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG "
    strSQL = strSQL & " SET     PTNO       = '" & txtPtNo.Text & "',"
    strSQL = strSQL & "         SNAME      = '" & txtName.Text & "',"
    strSQL = strSQL & "         ITEMCD     = '" & siTemCD & "',"
    strSQL = strSQL & "         GEOMCHCD   = '" & sGeomChCD & "',"
    strSQL = strSQL & "         GBIO       = '" & sGbIO & "',"
    strSQL = strSQL & "         SEX        = '" & sSex & "',"
    strSQL = strSQL & "         AGEYY      = " & sAgeYY & ","
    strSQL = strSQL & "         AGEMM      = '" & sAgeMM & "',"
    strSQL = strSQL & "         ROOMCODE   = '" & sRoomCode & "',"
    strSQL = strSQL & "         DEPTCODE   = '" & sDeptcode & "',"
    strSQL = strSQL & "         DRCODE     = '" & sDrCode & "',"
    
    strSQL = strSQL & "         OPNAME     = '" & Trim(txtOpname.Text) & "',"
    strSQL = strSQL & "         ORGANPART  = '" & Trim(sGeomChCD) & "',"
    
    strSQL = strSQL & "         INDATE     = TO_DATE('" & txtindate.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "         JDATE      = TO_DATE('" & txtJdate.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "         JTIMEHH    = " & Mid(LsJeobsuTm, 1, 2) & ","
    strSQL = strSQL & "         JTIMEMM    = " & Mid(LsJeobsuTm, 4, 2) & ","
    strSQL = strSQL & "         ORDERDT    = TO_DATE('" & txtOrderDt.Text & "','YYYY-MM-DD'),"
    strSQL = strSQL & "         GBRESULT   = '0',"
    strSQL = strSQL & "         GBGROSS    = '" & sGbGross & "'"
    strSQL = strSQL & " WHERE   CLASS      =  '" & txtClass.Text & "'"
    strSQL = strSQL & " AND     DATEYY     =  " & txtDateYY.Text & ""
    strSQL = strSQL & " AND     SEQNUM     =  " & txtSeqNum.Text & ""
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        Call Item_Clear
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "TWANAT_DIAG"
    End If
    
    Return

'----------------------------------------------------------------------------------------'

IDNOMST_INSERT:
    sGbIO = IIf(cboGbio.ListIndex = -1, "", Trim(LsGbio(cboGbio.ListIndex)))
    sSex = IIf(cboSex.ListIndex = -1, "", Trim(LsSex(cboSex.ListIndex)))
    sAgeYY = IIf(txtAge.Text = "", "", MidH(txtAge, 1, 2))
    sAgeMM = IIf(txtAgeMM.Text = "", "", Trim(txtAgeMM))
    sRoomCode = IIf(txtRoomcode.Text = "", "", Trim(txtRoomcode.Text))
    sDeptcode = IIf(cboDept.ListIndex = -1, "", Trim(LsDeptCode(cboDept.ListIndex)))
    sDrCode = IIf(cboDoctor.ListIndex = -1, "", Trim(LsDrCode(cboDoctor.ListIndex)))
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWEXAM_IDNOMST                                            "
    strSQL = strSQL & "       (  PtNo,        Sname,        Sex,         AgeYY,     "
    strSQL = strSQL & "          AgeMM,                     DeptCode,    RoomCode , "
    strSQL = strSQL & "          DrCode,      Gbio,         Bi )                    "
    strSQL = strSQL & " VALUES ('" & txtPtNo.Text & "',"
    strSQL = strSQL & "         '" & txtName.Text & "',"
    strSQL = strSQL & "         '" & sSex & "',"
    strSQL = strSQL & "         '" & sAgeYY & "',"
    strSQL = strSQL & "         '" & sAgeMM & "',"
    strSQL = strSQL & "         '" & sDeptcode & "',"
    strSQL = strSQL & "         '" & sRoomCode & "',"
    strSQL = strSQL & "         '" & sDrCode & "',"
    strSQL = strSQL & "         '" & sGbIO & "',"
    strSQL = strSQL & "         '')"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        Exit Sub
    End If
    
    Return

'----------------------------------------------------------------------------
IDNOMST_UPDATE:
    sGbIO = IIf(cboGbio.ListIndex = -1, "", Trim(LsGbio(cboGbio.ListIndex)))
    sSex = IIf(cboSex.ListIndex = -1, "", Trim(LsSex(cboSex.ListIndex)))
    sAgeYY = IIf(txtAge.Text = "", "", MidH(txtAge, 1, 2))
    sAgeMM = IIf(txtAgeMM.Text = "", "", Trim(txtAgeMM))
    sRoomCode = IIf(txtRoomcode.Text = "", "", Trim(txtRoomcode.Text))
    sDeptcode = IIf(cboDept.ListIndex = -1, "", Trim(LsDeptCode(cboDept.ListIndex)))
    sDrCode = IIf(cboDoctor.ListIndex = -1, "", Trim(LsDrCode(cboDoctor.ListIndex)))
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWEXAM_IDNOMST "
    strSQL = strSQL & " SET    Sname    = '" & txtName.Text & "',"
    strSQL = strSQL & "        Sex      = '" & sSex & "',"
    strSQL = strSQL & "        AgeYY    = '" & sAgeYY & "',"
    strSQL = strSQL & "        AgeMM    = '" & sAgeMM & "',"
    strSQL = strSQL & "        DeptCode = '" & sDeptcode & "',"
    strSQL = strSQL & "        DrCode   = '" & sDrCode & "',"
    strSQL = strSQL & "        RoomCode = '" & sRoomCode & "',"
    strSQL = strSQL & "        Bi       = '',"
    strSQL = strSQL & "        Gbio     = '" & sGbIO & "'"
    strSQL = strSQL & " WHERE  PtNo     = '" & txtPtNo.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        Exit Sub
    End If
    
    Return

'Error_Save:

'    On Error Resume Next
    

End Sub


Sub DOCTOR_SELECT()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim adoDr               As ADODB.Recordset
    
    If cboDept.ListIndex = -1 Then
       Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " SELECT  Drcode, Drname "
    strSQL = strSQL & " FROM    TWBAS_DOCTOR "
    strSQL = strSQL & " WHERE ( Drdept1 =  '" & Trim(LsDeptCode(cboDept.ListIndex)) & "'   "
    strSQL = strSQL & "     OR  Drdept2 =  '" & Trim(LsDeptCode(cboDept.ListIndex)) & "' ) "
    strSQL = strSQL & " AND     GBOUT   =  'N' "
    strSQL = strSQL & " ORDER   BY Drname "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        cboDoctor.Clear
        Exit Sub
    Else
        cboDoctor.Clear
        Do Until rs.EOF
            LsDrCode(i) = rs.Fields("Drcode").Value & ""
            cboDoctor.AddItem Right(LsDrCode(i), 4) & " " & rs.Fields("Drname").Value & ""
            rs.MoveNext
            i = i + 1
        Loop
        AdoCloseSet rs
    End If
        
     
End Sub


Private Sub cmdSelect_Click()
    
    Dim rs                  As ADODB.Recordset
    Dim i                   As Integer

    Call Item_Clear
    
    If optClass(0).Value = True Then
        txtClass = "P"
    ElseIf optClass(1).Value = True Then
        txtClass = "C"
    ElseIf optClass(2).Value = True Then
        txtClass = "R"
    End If

'-------------------------------------------------'
' 검사항목 DB READ                                '
'-------------------------------------------------'
    cboItem.Clear
    
    strSQL = ""
    strSQL = " SELECT * FROM TWEXAM_ITEMML"
    
'    If optClass(0).Value = True Then
        strSQL = strSQL & " WHERE SubStr(CODEKY,1,2) = '85' " '91
        strSQL = strSQL & "   AND  (CODEGU = '80' OR CODEGU = '89') "
        strSQL = strSQL & " ORDER BY CODEKY                 "
'    Else
'        strSQL = strSQL & " WHERE SubStr(CODEKY,1,2) = '85' "
'        strSQL = strSQL & "   AND CodeGu = '89' "
'        strSQL = strSQL & " ORDER BY ITEMNM                 "
'    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    i = 0
   ' cboItem.Clear
    
    Do Until rs.EOF
        LsItem(i) = rs.Fields("CODEKY").Value & ""
'        cboItem.AddItem rs.Fields("CODEKY").Value & "" & rs.Fields("ITEMNM").Value & ""
        cboItem.AddItem rs.Fields("CODEKY").Value & "" & rs.Fields("YAGEO").Value & ""
        
        rs.MoveNext
        i = i + 1
    Loop
    AdoCloseSet rs
    Call JEOBSU_PATIENT


End Sub

Private Sub Form_Activate()
'    optClass(0).SetFocus

End Sub

Private Sub Form_Load()
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    Dim LsSeqNum            As String * 8
    Dim LsPtNo              As String * 12
    Dim LsSname             As String * 20
    
    lblUser = GstrPassName
    
    GoSub Activate_Init
    GoSub Dept_Select
    
    GoSub Item_Read
    GoSub Geomch_Read
    
    
'    cboClass.Clear
'    cboClass.AddItem "진단병리"   '61
'    cboClass.AddItem "Histology"       '61
'    cboClass.AddItem "Cytology"     '62
'    cboClass.ListIndex = 0
    
    optClass(0).Value = True
    
    txtClass = "P"
    
    Exit Sub
    
'/------------------------------------------
Activate_Init:
    txtClass = ""
    txtDateYY = ""
    txtSeqNum = ""
    txtPtNo = ""
    txtName = ""
    txtAge = ""
    txtAgeMM = ""
    txtRoomcode = ""
    
    txtindate.Text = ""
    txtJdate = ""
    txtOrderDt = ""

    LsGbio(0) = "I"
    LsGbio(1) = "O"
    LsGbio(2) = "G"
    LsGbio(3) = "E"
    
    cboGbio.AddItem "입 원"
    cboGbio.AddItem "외 래"
    cboGbio.AddItem "검 진"
    cboGbio.AddItem "기 타"
    
    LsSex(0) = "M"
    LsSex(1) = "F"
    
    cboSex.AddItem "남 자"
    cboSex.AddItem "여 자"
    
    Return
        
'/--------------------------------------------
Dept_Select:
    strSQL = ""
    strSQL = strSQL & " SELECT DeptCode, DeptNameK "         '진료과선택
    strSQL = strSQL & " FROM   TWBAS_DEPT "
    strSQL = strSQL & " WHERE  GbJupsu =  '0' "
    strSQL = strSQL & " ORDER  BY Printranking "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Return
    
    Do Until rs.EOF
        LsDeptCode(i) = rs.Fields("DeptCode").Value & ""
        cboDept.AddItem rs.Fields("DeptNameK").Value & ""
        rs.MoveNext
        i = i + 1
    Loop
    AdoCloseSet rs
    Return
    
'--------------------------------------------------------------------
Item_Read:
'검사항목
    cboItem.Clear
    
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_ITEMML "
    strSQL = strSQL & " WHERE  SubStr(CODEKY,1,2) = '85' "  '91
    strSQL = strSQL & "   AND  (CODEGU = '80' OR CODEGU = '89') "
    strSQL = strSQL & " ORDER  BY CODEKY "
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Return
    i = 0
    Do Until rs.EOF
'        cboItem.AddItem rs.Fields("CODEKY").Value & rs.Fields("ITEMNM").Value & ""
        cboItem.AddItem rs.Fields("CODEKY").Value & "" & rs.Fields("YAGEO").Value & ""
        LsItem(i) = rs.Fields("CODEKY").Value & ""
        rs.MoveNext
        i = i + 1
    Loop
    
    AdoCloseSet rs
    Return
    
Geomch_Read:
' 검체 DB READ                                    '
    cboGeomch.Clear
    
'    strSQL = ""
'    strSQL = strSQL & " SELECT Codeky,   Codenm   "
'    strSQL = strSQL & " FROM   TWEXAM_SPECODE     "
'    strSQL = strSQL & " WHERE  Codegu = '13'      "
'    strSQL = strSQL & " ORDER  BY Codenm          "
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM   TWANAT_DICT            "
    strSQL = strSQL & " WHERE  SUBSTR(CODE,1,1) = 'T' "
    strSQL = strSQL & " ORDER  BY DXDICT                "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Return
    
    i = 0
    Do Until rs.EOF
        LsGeomch(i) = rs.Fields("Code").Value & ""
        cboGeomch.AddItem rs.Fields("DXDICT").Value & ""
        rs.MoveNext
        i = i + 1
    Loop
    AdoCloseSet rs
    Return


End Sub


Private Sub lstDisp_Click()
    
    Dim FindCLASS           As String * 2
    Dim FindDATEYY          As String * 4
    Dim FindSEQNUM          As String * 5
    Dim LsDrTemp            As String
    Dim i                   As Integer
    
    Dim rs                  As ADODB.Recordset
    
    FindCLASS = MidH(lstDisp.List(lstDisp.ListIndex), 1, 2)
    FindDATEYY = MidH(lstDisp.List(lstDisp.ListIndex), 4, 4)
    FindSEQNUM = MidH(lstDisp.List(lstDisp.ListIndex), 9, 5)
     
    If Trim(FindCLASS) = "" Then Exit Sub
    If Trim(FindDATEYY) = "" Then Exit Sub
    If Trim(FindSEQNUM) = "" Then Exit Sub
          
    cboItem.ListIndex = -1
    cboGeomch.ListIndex = -1
    cboGbio.ListIndex = -1
    cboSex.ListIndex = -1
    cboDept.ListIndex = -1
    cboDoctor.ListIndex = -1
    
    txtAge = ""
    txtAgeMM = ""
    txtRoomcode = ""
    txtPtNo = ""
    txtName = ""
    
    txtindate.Text = ""
    txtJdate = ""
    txtOrderDt = ""
          
    strSQL = ""
    strSQL = strSQL & " SELECT  CLASS, DATEYY, SEQNUM, PTNO, SNAME, GBIO,Opname,ORGANPART,         "
    strSQL = strSQL & "         SEX, AGEYY, AGEMM, ROOMCODE, DEPTCODE, DRCODE,ITEMCD, "
    strSQL = strSQL & "         GEOMCHCD, "
    strSQL = strSQL & "         TO_CHAR(INDATE, 'YYYY-MM-DD') INDATE, "
    strSQL = strSQL & "         TO_CHAR(JDATE, 'YYYY-MM-DD') JDATE, "
    strSQL = strSQL & "         TO_CHAR(ORDERDT,'YYYY-MM-DD') OrderDt "
    strSQL = strSQL & " FROM    TWANAT_DIAG             "
    strSQL = strSQL & " WHERE   CLASS  =  '" & FindCLASS & "'                    "
    strSQL = strSQL & " AND     DATEYY =  '" & FindDATEYY & "'                  "
    strSQL = strSQL & " AND     SEQNUM =  '" & FindSEQNUM & "'                  "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    LsDrTemp = rs.Fields("DRCODE").Value & ""
    TempDeptCode = rs.Fields("DEPTCODE").Value & ""
    txtClass = rs.Fields("CLASS").Value & ""
    txtDateYY = rs.Fields("DATEYY").Value & ""
    txtSeqNum = rs.Fields("SEQNUM").Value & ""
    txtPtNo = rs.Fields("PTNO").Value & ""
    txtName = rs.Fields("SNAME").Value & ""
    
    txtOpname.Text = rs.Fields("Opname").Value & ""
'    txtOrganPart.Text = rs.Fields("SNAME").Value & ""
    
    TempSex = rs.Fields("SEX").Value & ""
    TempGbio = rs.Fields("GBIO").Value & ""
    txtRoomcode = rs.Fields("ROOMCODE").Value & ""
    TempItem = rs.Fields("ITEMCD").Value & ""
    
'    TempGeomch = rs.Fields("GEOMCHCD").Value & ""
    TempGeomch = rs.Fields("OrganPart").Value & ""
    
    txtindate.Text = Format(rs.Fields("INDATE").Value & "", "YYYY-MM-DD")
    txtJdate = Format(rs.Fields("JDATE").Value & "", "YYYY-MM-DD")
    txtOrderDt = Format(rs.Fields("ORDERDT").Value & "", "YYYY-MM-DD")
    
    
    If Trim(txtClass.Text) = "P" Then
        optClass(0).Value = True
    ElseIf Trim(txtClass.Text) = "C" Then
        optClass(1).Value = True
    ElseIf Trim(txtClass.Text) = "R" Then
        optClass(2).Value = True
    End If
       
    For i = 0 To 5
        If Trim(TempGbio) = Trim(LsGbio(i)) Then
            cboGbio.ListIndex = i
            cboGbio.Refresh
        End If
    Next i
    
    For i = 0 To 1
        If Trim(TempSex) = Trim(LsSex(i)) Then
            cboSex.ListIndex = i
            cboSex.Refresh
        End If
    Next i
    
    txtAge = rs.Fields("AGEYY").Value & ""
    txtAgeMM = rs.Fields("AGEMM").Value & ""
              
    For i = 0 To 200
        If Trim(TempDeptCode) = Trim(LsDeptCode(i)) Then
            cboDept.ListIndex = i
            cboDept.Refresh
            Exit For
        End If
    Next i
    
    Call DOCTOR_SELECT
    For i = 0 To 500
        If Trim(LsDrTemp) = Trim(LsDrCode(i)) Then
            cboDoctor.ListIndex = i
            cboDoctor.Refresh
            Exit For
        End If
    Next i
    
    For i = 0 To 400  ' 50
        If Trim(TempItem) = Trim(LsItem(i)) Then
            cboItem.ListIndex = i
            cboItem.Refresh
            Exit For
        End If
    Next i
    
'    For I = 1 To 100
'
'        Debug.Print cboGeomch.ListIndex = I
'    Next I
    
    For i = 0 To 700
        If Trim(TempGeomch) = Trim(LsGeomch(i)) Then
            cboGeomch.ListIndex = i
            cboGeomch.Refresh
            Exit For
        End If
    Next i
     
     AdoCloseSet rs
    
End Sub


Private Sub optClass_Click(Index As Integer)
    
    If optClass(0).Value = True Then
        txtClass = "P"
    ElseIf optClass(1).Value = True Then
        txtClass = "C"
    ElseIf optClass(2).Value = True Then
        txtClass = "R"
    End If


End Sub

Private Sub txtAge_GotFocus()
 
    txtAge.SelStart = 0
    txtAge.SelLength = Len(txtAge.Text)

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtAgeMM_GotFocus()
 
    txtAgeMM.SelStart = 0
    txtAgeMM.SelLength = Len(txtAgeMM.Text)


End Sub


Private Sub txtAgeMM_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtClass_GotFocus()
 
    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)

End Sub

Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtDateYY_GotFocus()
     
'    txtDateYY = Format(Date, "YY")
'    txtDateYY = Format(Date, "YYYY")
    txtDateYY = Dual_Date_Get("yyyy")
    
    txtDateYY.SelStart = 0
    txtDateYY.SelLength = Len(txtDateYY.Text)

End Sub

Private Sub txtDATEYY_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    
    If Len(txtDateYY.Text) < 4 Then
        MsgBox "년도는 4자리를 입력하여 주시기 바랍니다."
        txtDateYY.SetFocus
        Exit Sub
    End If
    
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtDateYY_LostFocus()
    
    Dim rs                  As ADODB.Recordset
        
    strSQL = ""
'    strSQL = strSQL & " SELECT  SEQNUM"
    strSQL = strSQL & " SELECT  max(SEQNUM) maxseq"
    strSQL = strSQL & " FROM    TWANAT_DIAG"
    strSQL = strSQL & " WHERE   CLASS    = '" & txtClass & "'"
    strSQL = strSQL & " AND     DATEYY   = " & txtDateYY.Text & ""
    strSQL = strSQL & " ORDER   BY SEQNUM   DESC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtSeqNum = 1
    Else
'        txtSeqnum = Val(rs.Fields("SEQNUM").Value) + 1
        txtSeqNum = Val(rs.Fields("maxseq").Value) + 1
        AdoCloseSet rs
    End If

End Sub

Private Sub txtindate_GotFocus()
    txtindate = Dual_Date_Get("yyyy-MM-dd")
    
    txtindate.SelStart = 0
    txtindate.SelLength = Len(txtindate.Text)

End Sub

Private Sub txtindate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"
    
End Sub

Private Sub txtJdate_GotFocus()
 
'    txtJdate = Format(Date, "yyyy-mm-dd")
    txtJdate = Dual_Date_Get("yyyy-MM-dd")
    
    txtJdate.SelStart = 0
    txtJdate.SelLength = Len(txtJdate.Text)

End Sub

Private Sub txtJdate_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtName_GotFocus()

    DoEvents
    txtName.IMEMode = vbIMEModeHangul
    
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtName_LostFocus()
    
    DoEvents
    txtName.IMEMode = vbIMEModeAlpha

End Sub

Private Sub txtOpName_Change()
    txtOpname.SelStart = 0
    txtOpname.SelLength = Len(txtOpname.Text)

End Sub

Private Sub txtOpname_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtOrderDt_GotFocus()
 
'    txtOrderDt = Format(Date, "yyyy-mm-dd")
    txtOrderDt = Dual_Date_Get("yyyy-MM-dd")
    
    txtOrderDt.SelStart = 0
    txtOrderDt.SelLength = Len(txtOrderDt.Text)


End Sub

Private Sub txtOrderDt_KeyPress(KeyAscii As Integer)
    
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

    If txtPtNo = "" Then Exit Sub
    
    If optClass(2).Value = True Then Exit Sub
    
    txtPtNo = Format(txtPtNo, "00000000")
    
    Call IDNO_SELECT
    
End Sub

Private Sub txtRoomCode_GotFocus()
 
    txtRoomcode.SelStart = 0
    txtRoomcode.SelLength = Len(txtRoomcode.Text)

End Sub

Private Sub txtRoomCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtSeqnum_GotFocus()

    txtSeqNum.SelStart = 0
    txtSeqNum.SelLength = Len(txtSeqNum.Text)

End Sub

Private Sub txtSeqnum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtSeqNum_LostFocus()

    Dim rs                  As ADODB.Recordset
    
    Dim LsDrTemp            As String
    Dim i                   As Integer
   

    strSQL = ""
    strSQL = strSQL & " SELECT a.*, "
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') Jdate1, "
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt1 "
    strSQL = strSQL & "   FROM TWANAT_DIAG  a "
    strSQL = strSQL & "  WHERE CLASS = '" & txtClass & "' "
    strSQL = strSQL & "    AND DATEYY = '" & txtDateYY & "' "
    strSQL = strSQL & "    AND SEQNUM = '" & txtSeqNum & "' "
    
    txtPtNo = ""
    txtName = ""
    cboItem.ListIndex = -1
    cboGeomch.ListIndex = -1
    cboGbio.ListIndex = -1
    cboSex.ListIndex = -1
    txtAge = ""
    txtAgeMM = ""
    txtRoomcode = ""
    cboDept.ListIndex = -1
    cboDoctor.ListIndex = -1
    txtJdate = ""
    txtOrderDt = ""
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
     
    txtClass.Text = Trim(rs.Fields("CLASS").Value & "")

       If txtClass.Text = "P " Then
'           cboClass.Text = "Histology"
           optClass(0).Value = True
       ElseIf txtClass.Text = "C " Then
'           cboClass.Text = "Cytology"
           optClass(1).Value = True
       ElseIf txtClass.Text = "R " Then
'           cboClass.Text = "Cytology"
           optClass(2).Value = True
       End If
       
    txtDateYY = rs.Fields("DATEYY").Value & ""
    txtSeqNum = rs.Fields("SEQNUM").Value & ""
    
    txtPtNo = rs.Fields("PTNO").Value & ""
    txtName = rs.Fields("SNAME").Value & ""
    
    TempGbio = rs.Fields("GBIO").Value & ""
    For i = 0 To 5
        If Trim(TempGbio) = Trim(LsGbio(i)) Then
           cboGbio.ListIndex = i
           cboGbio.Refresh
           Exit For
        End If
    Next i
    
    TempSex = rs.Fields("SEX").Value & ""
    For i = 0 To 1
        If Trim(TempSex) = Trim(LsSex(i)) Then
           cboSex.ListIndex = i
           cboSex.Refresh
           Exit For
        End If
    Next i
    
    If Val(rs.Fields("AGEYY").Value & "") = 0 Then
        txtAge = ""
    Else
        txtAge = rs.Fields("AGEYY").Value & ""
    End If
    
    If Val(rs.Fields("AGEMM").Value & "") = 0 Then
        txtAgeMM = ""
    Else
        txtAgeMM = rs.Fields("AGEMM").Value & ""
    End If
    
    txtRoomcode = rs.Fields("ROOMCODE").Value & ""
    
    TempDeptCode = rs.Fields("DEPTCODE").Value & ""
    For i = 0 To 200
        If Trim(TempDeptCode) = Trim(LsDeptCode(i)) Then
           cboDept.ListIndex = i
           cboDept.Refresh
           Exit For
        End If
    Next i
    
    LsDrTemp = rs.Fields("DRCODE").Value & ""
    Call DOCTOR_SELECT
    For i = 0 To 500
        If Trim(LsDrTemp) = Trim(LsDrCode(i)) Then
           cboDoctor.ListIndex = i
           cboDoctor.Refresh
           Exit For
        End If
    Next i
    
    TempItem = rs.Fields("ITEMCD").Value & ""
    For i = 0 To 50
        If Trim(TempItem) = Trim(LsItem(i)) Then
           cboItem.ListIndex = i
           cboItem.Refresh
           Exit For
        End If
    Next i
    
'    TempGeomch = rs.Fields("GEOMCHCD").Value & ""
'    For i = 0 To 100
'         If Trim(TempGeomch) = Trim(LsGeomch(i)) Then
'            cboGeomch.ListIndex = i
'            cboGeomch.Refresh
'            Exit For
'         End If
'     Next i
     
    TempGeomch = rs.Fields("ORGANPART").Value & ""
    For i = 0 To 700
         If Trim(TempGeomch) = Trim(LsGeomch(i)) Then
            cboGeomch.ListIndex = i
            cboGeomch.Refresh
            Exit For
         End If
     Next i
    
    txtJdate = rs.Fields("JDATE1").Value & ""
    txtOrderDt = rs.Fields("ORDERDT1").Value & ""
    
    AdoCloseSet rs
       
End Sub

