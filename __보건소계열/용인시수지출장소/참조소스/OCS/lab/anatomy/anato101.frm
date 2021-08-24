VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "조직병리검사관리"
   ClientHeight    =   8250
   ClientLeft      =   870
   ClientTop       =   1350
   ClientWidth     =   11940
   Icon            =   "ANATO101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8250
   ScaleWidth      =   11940
   WindowState     =   2  '최대화
   Begin VB.PictureBox pctLogin 
      Appearance      =   0  '평면
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   4608
      ScaleHeight     =   2085
      ScaleWidth      =   2685
      TabIndex        =   40
      Top             =   3048
      Visible         =   0   'False
      Width           =   2715
      Begin VB.TextBox txtUserId 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   42
         Top             =   345
         Width           =   1080
      End
      Begin VB.TextBox txtPassWord 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  '사용 못함
         Left            =   1305
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   735
         Width           =   1080
      End
      Begin VB.TextBox txtExamDate 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   41
         Top             =   4305
         Width           =   1650
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   435
         Left            =   2010
         TabIndex        =   45
         Top             =   1530
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "취소"
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
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   435
         Left            =   1320
         TabIndex        =   44
         Top             =   1530
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "확인"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   165
         TabIndex        =   50
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   195
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         BorderStyle     =   1  '단일 고정
         Caption         =   " Login Change......"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   2685
      End
      Begin VB.Label lblUserName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1290
         TabIndex        =   47
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         Caption         =   "검사  일자"
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
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   4350
         Width           =   1065
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   300
         Picture         =   "ANATO101.frx":030A
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.PictureBox picPanel 
      Height          =   5808
      Left            =   8730
      ScaleHeight     =   5745
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   1500
      Width           =   2172
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   10
         Left            =   720
         TabIndex        =   9
         ToolTipText     =   "검사코드관리"
         Top             =   4320
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":0614
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   17
         Left            =   0
         TabIndex        =   29
         ToolTipText     =   "OCS ORDER 접수"
         Top             =   717
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":0A66
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   3
         Left            =   1416
         TabIndex        =   26
         ToolTipText     =   "건양대병원.."
         Top             =   0
         Visible         =   0   'False
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":0D80
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   21
         Left            =   1410
         TabIndex        =   25
         ToolTipText     =   "보관SLIDE조회"
         Top             =   2880
         Visible         =   0   'False
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":2512
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   18
         Left            =   0
         TabIndex        =   24
         ToolTipText     =   "검사아이템관리"
         Top             =   5040
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":282C
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   23
         Left            =   720
         TabIndex        =   23
         ToolTipText     =   "보관검체조회"
         Top             =   2880
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":2B46
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   9
         Left            =   0
         TabIndex        =   22
         ToolTipText     =   "병리사전관리"
         Top             =   4317
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":2E60
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   12
         Left            =   1416
         TabIndex        =   21
         ToolTipText     =   "접수환자출력"
         Top             =   720
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":373A
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   6
         Left            =   720
         TabIndex        =   20
         ToolTipText     =   "사용자관리"
         Top             =   0
         Visible         =   0   'False
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":3A54
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   15
         Left            =   1416
         TabIndex        =   19
         ToolTipText     =   "특수검사환자조회"
         Top             =   2160
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":3EA6
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   13
         Left            =   720
         TabIndex        =   18
         ToolTipText     =   "진단명조회"
         Top             =   2160
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":41C1
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   14
         Left            =   1416
         TabIndex        =   17
         ToolTipText     =   "특수검사접수"
         Top             =   1440
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":44DB
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   5
         Left            =   720
         TabIndex        =   16
         ToolTipText     =   "상용구절관리"
         Top             =   3600
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":5C6D
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   4
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "매크로관리"
         Top             =   3597
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":60BF
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   11
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "검사결과조회"
         Top             =   2157
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":6999
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   7
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "결과입력"
         Top             =   1437
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":6CB3
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   705
         Index           =   1
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "로그 인"
         Top             =   0
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":7105
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   19
         Left            =   900
         TabIndex        =   11
         ToolTipText     =   "면역염색"
         Top             =   4320
         Visible         =   0   'False
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":79DF
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   16
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "육안사진환자조회"
         Top             =   2877
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":7CF9
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   8
         Left            =   720
         TabIndex        =   8
         ToolTipText     =   "결과지출력"
         Top             =   1440
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":8013
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   22
         Left            =   1092
         TabIndex        =   7
         ToolTipText     =   "전자현미경검사"
         Top             =   4320
         Visible         =   0   'False
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":832D
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   20
         Left            =   1272
         TabIndex        =   6
         ToolTipText     =   "면역형광염색"
         Top             =   4320
         Visible         =   0   'False
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":8647
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   2
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "의뢰접수등록"
         Top             =   720
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO101.frx":8961
      End
      Begin Threed.SSCommand cmdWork 
         Height          =   708
         Index           =   24
         Left            =   1416
         TabIndex        =   4
         ToolTipText     =   "작업종료"
         Top             =   5040
         Width           =   708
         _Version        =   65536
         _ExtentX        =   1249
         _ExtentY        =   1249
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "ANATO101.frx":8DB3
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7350
      Left            =   570
      TabIndex        =   0
      Top             =   930
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   12965
      _StockProps     =   15
      ForeColor       =   8388608
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.Timer Timer1 
         Left            =   180
         Top             =   150
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1080
         Index           =   1
         Left            =   3330
         ScaleHeight     =   1050
         ScaleWidth      =   3255
         TabIndex        =   32
         Top             =   6180
         Visible         =   0   'False
         Width           =   3285
         Begin FPSpread.vaSpread ssOrderInfo 
            Height          =   690
            Left            =   105
            TabIndex        =   33
            Top             =   300
            Width           =   3045
            _Version        =   196608
            _ExtentX        =   5371
            _ExtentY        =   1217
            _StockProps     =   64
            BackColorStyle  =   1
            DisplayColHeaders=   0   'False
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
            GridColor       =   15463915
            MaxCols         =   2
            MaxRows         =   3
            ScrollBars      =   0
            ShadowColor     =   12632256
            ShadowDark      =   8421504
            ShadowText      =   0
            SpreadDesigner  =   "ANATO101.frx":90CD
            UserResize      =   1
            VisibleCols     =   2
            VisibleRows     =   3
         End
         Begin VB.Label Label3 
            BackColor       =   &H00800000&
            Caption         =   "  Order    Information"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '투명
         Caption         =   "건양대학교병원"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   21.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   432
         Left            =   180
         TabIndex        =   31
         Top             =   6780
         Width           =   3024
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "조직병리"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   960
         Left            =   1476
         TabIndex        =   30
         Top             =   216
         Width           =   3852
      End
      Begin VB.Image Image1 
         Appearance      =   0  '평면
         Height          =   7176
         Left            =   804
         Picture         =   "ANATO101.frx":93FA
         Stretch         =   -1  'True
         Top             =   84
         Width           =   5028
      End
   End
   Begin Threed.SSPanel PanelIcon 
      Height          =   7350
      Left            =   8100
      TabIndex        =   1
      Top             =   915
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   12965
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.Label LabelMsg 
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "작 업 내 용"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   225
         TabIndex        =   35
         Top             =   6660
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00800000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "작 업 내 용"
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
         Height          =   312
         Index           =   0
         Left            =   636
         TabIndex        =   2
         Top             =   240
         Width           =   2172
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   648
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   11940
      _Version        =   65536
      _ExtentX        =   21061
      _ExtentY        =   1143
      _StockProps     =   15
      Caption         =   "TISSUE PATHOLOGY"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
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
      Begin VB.PictureBox Picture1 
         Height          =   405
         Index           =   0
         Left            =   9660
         ScaleHeight     =   345
         ScaleWidth      =   2115
         TabIndex        =   37
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   90
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox picTip 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E1FAFA&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   90
      ScaleHeight     =   645
      ScaleWidth      =   1485
      TabIndex        =   27
      Top             =   -30
      Visible         =   0   'False
      Width           =   1485
      Begin Threed.SSPanel panTip 
         Height          =   330
         Left            =   90
         TabIndex        =   28
         Top             =   315
         Visible         =   0   'False
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "자료등록"
         ForeColor       =   12583104
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Outline         =   -1  'True
         FloodColor      =   12582912
         Autosize        =   1
      End
   End
End
Attribute VB_Name = "Anato_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Dim LbUserId            As Boolean
    'Dim LbPassWord          As Boolean
    'Dim LsPassword          As String
    Dim LiPassChkCount      As Integer
    Dim LiIndex(0 To 100)   As Integer


Private Sub Form_Activate()
    GbTimmerOn = True
    LabelMsg = ""
    
'    Call SPECODE
'    Call Timer1_Timer

End Sub


Private Sub Form_Load()

    Call DbAdoConnect("TW_MIS_EXAM", "HOSPITAL", "kuh2")
    
    GsLoginOk = "LOGOUT"
    
'    GbTimmerOn = True
'    picPanel.Enabled = False
    
    FrmIdPass.Show vbModal
    
    GsLoginOk = "LOGIN"
    
    lblUser = GstrPassName

End Sub


Private Sub cmdWork_Click(Index As Integer)
    
    Dim i                   As Integer
    Dim Response            As Integer

    If Index <> 1 And Index <> 24 And GsLoginOk = "LOGOUT" Then
        GoSub SUB_NOTLOGIN_MESS
        Exit Sub
    End If
    
    
    GbTimmerOn = False
    
    Select Case Index
        Case 1
            GoSub SUB_LOGIN_PROC                    '로그인
        Case 2
            Anato_Menual_Jeobsu.Show vbModal        '의뢰 접수 등록
        Case 4
            Anato_Macro.Left = 500
            Anato_Macro.Top = 1200     '1550
            Anato_Macro.Show vbModal                '매크로관리
        Case 5
            Anato_Use_Word.Left = 370
            Anato_Use_Word.Top = 1200    ' 1500
            Anato_Use_Word.Show vbModal             '상용구절 관리
'        Case 6
'            Anato_User.Left = 370
'            Anato_User.Top = 1200   '1530
'            Anato_User.Show vbModal                 '사용자 관리
        Case 7
            Anato_Result.Show                       '진단서 작성
            Me.Hide
        Case 8
            Anato_Result_Print.Show vbModal         '결과 출력
        Case 9
            Anato_Dict.Left = 500
            Anato_Dict.Top = 1200 '1550
            Anato_Dict.Show vbModal                 '병리사전관리
        Case 10
            GCodegu = "50"
            Anato_Code.Left = 370
            Anato_Code.Top = 1200   ' 1430
            Anato_Code.Show vbModal                 '검사코드관리
        Case 11
            Anato_Result_View.Show                  '검사결과조회
        Case 12
            Anato_Jeobsu_Print.Show                 '접수환자출력
        Case 13
            Anato_DiagName_Search.Show              '진단명 검색
        Case 14                                     '특수검사접수
            Anato_Special_Diag.Top = 1200   '6650
            Anato_Special_Diag.Left = 580  '90
            Anato_Special_Diag.Show 1
        
        Case 15
            Anato_Dyeing_Persons.Show               '특수검사환자조회
        Case 16
            Anato_EyePhoto_View.Show                '육안사진환자조회
        Case 17
            Anato_OCS_Jeobsu.Show                   'OCS ORDER 접수
            Me.Hide
        Case 18
            Anato_ItemCode.Show                     '검사아이템관리
'        Case 19
'            GCodegu = "87" '56
'            Anato_Code.Left = 370
'            Anato_Code.Top = 1200   ' 1430
'            Anato_Code.Show vbModal                 '면역염색
'        Case 20
'            GCodegu = "84" '57
'            Anato_Code.Left = 370
'            Anato_Code.Top = 1200   ' 1430
'            Anato_Code.Show vbModal                 '면역형광염색
        Case 21
            Anato_Slide_View.Show                   '보관SLIDE 조회
'        Case 22
'            GCodegu = "86" '58
'            Anato_Code.Left = 370
'            Anato_Code.Top = 1200   ' 1430
'            Anato_Code.Show vbModal                 '전자현미경검사
        Case 23
            Anato_Specimen_View.Show                '보관검체조회
        Case 24
            Unload Me                               '종료
            End
        Case Else
            GoSub WORKINGMESS
    End Select

    Exit Sub
    
'----------------------------------------------------------------------------------------'
WORKINGMESS:
    Response = MsgBox("현재 개발작업 중 입니다.", vbOKOnly + vbInformation, "진단병리")
    
    Return
'----------------------------------------------------------------------------------------'
SUB_NOTLOGIN_MESS:
    Response = MsgBox("올바른 사용자가 아닙니다. 사용자 확인을 하세요.", vbOKOnly + vbCritical, "진단병리")
    
    Return

'----------------------------------------------------------------------------------------'
SUB_LOGIN_PROC:
    If GsLoginOk = "LOGIN" Then
        Response = MsgBox("현재 '" & Trim(GstrPassName) & "' 사용자가 로그인 중입니다. 다른 사용자로 로그인 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2, "진단병리")
        If Response = vbNo Then
            Return
        End If
    End If

    GsLoginOk = "LOGOUT"
    txtUserId = ""
    lblUser = GstrPassName
    txtPassword = ""
    pctLogin.Visible = True
    txtUserId.SetFocus

'    txtExamDate = Dual_Date_Get("yyyy-MM-dd")
'    lblUserName = ""
'    LbUserId = False
'    LbPassWord = False
'    picPanel.Enabled = False
    
    
    
    
'    If Trim(txtUserId) = "" Then Exit Sub
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & "   FROM TWBAS_PASS "
'    strSQL = strSQL & "  WHERE IDNUMBER = '" & txtUserId & "'"

'    Result = AdoOpenSet(rs, strSQL)
    
'    If Result Then
'        lblUserName = Trim(rs.Fields("Name").Value & "")
'        LsPassword = Trim(rs.Fields("Password").Value & "")
'        LsDeptNO = Trim(rs.Fields("Deptcode").Value & "")
'        LbUserId = True
'        LbPassWord = True
'    Else
'        lblUserName = ""
'        LsPassword = ""
'        LbUserId = False
'    End If
    
'    AdoCloseSet rs
    
    
    
    
    Return
    
End Sub


Private Sub cmdWork_LostFocus(Index As Integer)

    'picTip.Visible = False
    LabelMsg = ""
    
End Sub


Private Sub cmdWork_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim LsTipMess(1 To 24)  As String
        
    LsTipMess(1) = "로그 인"
    LsTipMess(2) = "의뢰접수등록"
    LsTipMess(3) = "건양대병원.."
    LsTipMess(4) = "매크로관리"
    LsTipMess(5) = "상용구절관리"
    LsTipMess(6) = "사용자관리"
    LsTipMess(7) = "결과입력"               ' "진단서작성"
    LsTipMess(8) = "결과지출력"             ' "진단서출력"
    LsTipMess(9) = "병리사전관리"
    LsTipMess(10) = "검사코드관리"
    LsTipMess(11) = "검사결과조회"
    LsTipMess(12) = "접수환자출력"
    LsTipMess(13) = "진단명조회"
    LsTipMess(14) = "특수검사접수"
    LsTipMess(15) = "특수염색환자조회"
    LsTipMess(16) = "육안사진환자조회"
    LsTipMess(17) = "OCS ORDER 접수"
    LsTipMess(18) = "검사아이템관리"
    LsTipMess(19) = "면역염색"
    LsTipMess(20) = "면역형광염색"
    LsTipMess(21) = "보관SLIDE조회"
    LsTipMess(22) = "전자현미경검사"
    LsTipMess(23) = "보관검체조회"
    LsTipMess(24) = "작업종료"
    
    cmdWork(Index).SetFocus
    
'    picTip.Visible = True
'    panTip.Visible = True

    If LsTipMess(Index) = "" Then Exit Sub
    
'    picTip.Top = picPanel.Top + cmdWork(Index).Top + 380
'    picTip.Left = (picPanel.Left + cmdWork(Index).Left) - panTip.Width
'    panTip.Caption = LsTipMess(Index)
    
'    picTip.Width = panTip.Width
'    picTip.Height = panTip.Height
'    panTip.Left = 0
'    panTip.Top = 0

     LabelMsg = LsTipMess(Index)

End Sub


Private Sub SSPanel3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LabelMsg = ""

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If Cancel = 1 Then
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DbAdoDisConnect
    
'    pctLogin.Visible = False
    End

End Sub



'Private Sub Timer1_Timer()
'    Dim rs                  As ADODB.Recordset
'
'    Dim i                   As Integer
'    Dim LiSlipNo1           As Integer
'
'    On Error Resume Next
'
'    If GbTimmerOn = False Then Exit Sub
'
'    '-------------------------------------------------------------
'    ' Spread_Col2_Clear
'
'    ssOrderInfo.BlockMode = True
'
'    ssOrderInfo.Row = 1
'    ssOrderInfo.Row2 = 3        'ssOrderInfo.DataRowCnt
'
'    ssOrderInfo.Col = 2
'    ssOrderInfo.Col2 = 2
'
'    ssOrderInfo.Text = ""
'    ssOrderInfo.BlockMode = False
'
'
'    '-------------------------------------------------------------
'    ' Get_Exam_Order & Count Display
''    strSQL = ""
''    strSQL = strSQL & " SELECT  *"
''    strSQL = strSQL & " FROM    TWEXAM_ORDER "
''    strSQL = strSQL & " WHERE   SLIPNO1  = 61 "        '>=91
''    strSQL = strSQL & " AND    (JeobsuYN  = ' ' OR  JeobsuYN IS NULL)"
'
'''    strSQL = ""
'''    strSQL = strSQL & " SELECT * "
'''    strSQL = strSQL & "   FROM TWEXAM_ORDER O, "
'''    strSQL = strSQL & "        TWEXAM_ITEMML L "
'''    strSQL = strSQL & "  WHERE O.SLIPNO1 = 85 "
'''    strSQL = strSQL & "    AND O.ITEMCD = L.CODEKY "
'''    strSQL = strSQL & "    AND L.CODEGU BETWEEN '80' AND '90' "
'''    strSQL = strSQL & "    AND (O.JeobsuYN  = ' ' OR  O.JeobsuYN IS NULL) "
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT PTNO, JEOBSUDT, ITEMCD , CODEGU "
'    strSQL = strSQL & "   FROM TWEXAM_ORDER O, "
'    strSQL = strSQL & "        TWEXAM_ITEMML L "
'    strSQL = strSQL & "  Where O.SLIPNO1 = 85 "
'    strSQL = strSQL & "    AND O.ITEMCD = L.CODEKY "
'    strSQL = strSQL & "    AND L.CODEGU BETWEEN '80' AND '90' "
'    strSQL = strSQL & "    AND (O.JeobsuYN  = ' ' OR  O.JeobsuYN IS NULL) "
'    strSQL = strSQL & "    AND (O.ITEMCD BETWEEN '850001' AND '851499' "
'    strSQL = strSQL & "         OR O.ITEMCD BETWEEN '859001' AND '859999') "
'
'    Result = AdoOpenSet(rs, strSQL)
'
'    If Result = False Then
'        ssOrderInfo.Col = 2
'        ssOrderInfo.Row = 1
'        ssOrderInfo.Text = "0"
'        ssOrderInfo.TypeHAlign = SS_CELL_H_ALIGN_LEFT
'
'        ssOrderInfo.Row = 2
'        ssOrderInfo.Text = "0"
'        ssOrderInfo.TypeHAlign = SS_CELL_H_ALIGN_LEFT
'
'        Exit Sub
'    End If
'
'    i = 0
'    Do Until rs.EOF
'
'        If Trim(rs.Fields("CodeGu").Value & "") = "80" Then  '1
'            ssOrderInfo.Row = 1
'            ssOrderInfo.Col = 2
'            ssOrderInfo.Text = Val(ssOrderInfo.Text) + 1'
'
'            ssOrderInfo.TypeHAlign = SS_CELL_H_ALIGN_LEFT
'        ElseIf Trim(rs.Fields("CodeGu").Value & "") = "89" Then
'            ssOrderInfo.Row = 2
'            ssOrderInfo.Col = 2
'            ssOrderInfo.Text = Val(ssOrderInfo.Text) + 1
'            ssOrderInfo.TypeHAlign = SS_CELL_H_ALIGN_LEFT
''        ElseIf Trim(rs.Fields("CodeGu").Value & "") = "81" Then
''            ssOrderInfo.Row = 3
''            ssOrderInfo.Col = 2
''            ssOrderInfo.Text = Val(ssOrderInfo.Text) + 1
''            ssOrderInfo.TypeHAlign = SS_CELL_H_ALIGN_LEFT
'        End If
'
'        rs.MoveNext: i = i + 1
'    Loop
'
'    AdoCloseSet rs
'
'End Sub


'Private Sub SPECODE()
'    Dim rs                  As ADODB.Recordset
'
'    Dim LiSlipNo1           As Integer'
'
'    '-------------------------------------------------------------
'    ' Get_Specode12
'    strSQL = ""
'    strSQL = strSQL & " SELECT  * "
'    strSQL = strSQL & " FROM    TWEXAM_SPECODE "
'    strSQL = strSQL & " WHERE   Codegu = '12' "
'    strSQL = strSQL & " AND     Substr(Codeky,1,2) >= '85' "
'    strSQL = strSQL & " ORDER   BY CODEKY ASC "
'
'    Result = AdoOpenSet(rs, strSQL)
'
'    If Result = False Then Exit Sub
'
'    i = 0
'    Do Until rs.EOF
'        LiSlipNo1 = Mid(rs.Fields("CODEKY").Value, 1, 2)
'        LiIndex(LiSlipNo1) = i + 1
'        ssOrderInfo.Row = i + 1
'        ssOrderInfo.Col = 1
'        ssOrderInfo.Text = rs.Fields("CODENM").Value & ""
'
'        rs.MoveNext: i = i + 1
'    Loop
'    AdoCloseSet rs'
'
'End Sub


Private Sub CmdOK_Click()

    Dim Response            As Integer
    
    strSQL = ""
    strSQL = strSQL & "SELECT Name, PassWord, Class, Grade, Part, DeptCode "
    strSQL = strSQL & "  FROM TWBAS_PASS "
    strSQL = strSQL & " WHERE IDnumber  = '" & Trim(txtUserId.Text) & "'"
    strSQL = strSQL & "   AND PassWord  = '" & Trim(txtPassword.Text) & "'"
    strSQL = strSQL & "   AND DeptCode  = 'AP'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = True And Rowindicator <> 0 Then
        GstrPassWord = AdoGetString(rs, "PassWord", 0)
        GstrPassName = AdoGetString(rs, "Name", 0)
        GstrPassClass = AdoGetString(rs, "Class", 0)
        GstrPassGrade = AdoGetString(rs, "Grade", 0)
        GstrPassPart = AdoGetString(rs, "Part", 0)
        GstrPassDept = AdoGetString(rs, "DeptCode", 0)
        GstrIdnumber = Format(txtUserId.Text, "000000")
        GstrPassIDnumber = Format(txtUserId.Text, "000000")
        
        GsExDate = Dual_Date_Get("yyyy-MM-dd")
        
        pctLogin.Visible = False
        Response = MsgBox("'" & Trim(GstrPassName) & "'" & " 사용자가 로그인 하였습니다.", vbOKOnly + vbInformation, "진단병리")
        
        lblUser = Trim(GstrPassName)
    
        GsLoginOk = "LOGIN"
    Else
        LiPassChkCount = LiPassChkCount + 1
        If LiPassChkCount > 2 Then
            Response = MsgBox(" 이프로그램을 사용할 수 없습니다. ", vbOKOnly + vbInformation, "진단병리")
            LiPassChkCount = 0
            pctLogin.Visible = False
        Else
            txtUserId.SetFocus
        End If
    End If
    
    AdoCloseSet rs
 
End Sub


Private Sub cmdCancel_Click()
    pctLogin.Visible = False

End Sub

Private Sub txtUserId_GotFocus()

    txtUserId.SelStart = 0
    txtUserId.SelLength = Len(txtUserId.Text)

End Sub


Private Sub txtUserId_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

    SendKeys "{tab}"

End Sub


Private Sub txtPassword_GotFocus() '
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0

    SendKeys "{tab}"

End Sub

Private Sub txtPassword_LostFocus()
    
    txtPassword.Text = UCase(txtPassword.Text)

End Sub



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'Private Sub txtExamDate_GotFocus()'
'
'    txtExamDate.SelStart = 0
'    txtExamDate.SelLength = Len(txtExamDate.Text)'
'
'End Sub
'
'
'Private Sub txtExamDate_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii <> 13 Then Exit Sub
'    KeyAscii = 0
'
'    SendKeys "{tab}"
'
'End Sub


'Private Sub txtPassword_LostFocus()
'
'    If Trim(txtPassWord) = "" Then Exit Sub
'
'    If UCase(Trim(txtPassWord)) = Trim(LsPassword) Then
'        LbPassWord = True
'    Else
'        LbPassWord = False
'    End If
'
'End Sub


'Private Sub txtUserId_LostFocus()
'    Dim rs                  As ADODB.Recordset
'
'    If Trim(txtUserId) = "" Then Exit Sub
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & "   FROM TWBAS_PASS "
'    strSQL = strSQL & "  WHERE IDNUMBER = '" & txtUserId & "'"

'    Result = AdoOpenSet(rs, strSQL)
    
'    If Result Then
'        lblUserName = Trim(rs.Fields("Name").Value & "")
'        LsPassword = Trim(rs.Fields("Password").Value & "")
'        LsDeptNO = Trim(rs.Fields("Deptcode").Value & "")
'        LbUserId = True
'        LbPassWord = True
'    Else
'        lblUserName = ""
'        LsPassword = ""
'        LbUserId = False
'    End If
    
'    AdoCloseSet rs

'End Sub

'Private Sub CmdOK_Click()
'
'    Dim Response        As Integer
    
'    If LbUserId = True And LbPassWord = True Then
'        GsExDate = txtExamDate
'        GstrPassIDnumber = txtUserId
''        GsUserName = lblUserName
''           lblUser = lblUserName
'        pctLogin.Visible = False
'        Response = MsgBox("'" & Trim(lblUser) & "'" & " 사용자가 로그인 하였습니다.", vbOKOnly + vbInformation, "진단병리")
'       picPanel.Enabled = True
        
'        GstrPassIDnumber = Trim$(txtUserId)
''        GstrPassDept = LsDeptNO
'        GstrPassClass = "OCS"
        
'    Else
'        LiPassChkCount = LiPassChkCount + 1
'        If LiPassChkCount > 2 Then
'            LiPassChkCount = 0
'            pctLogin.Visible = False
'        Else
'            txtUserId.SetFocus
'        End If
'    End If
 
'End Sub

'Private Sub CmdCancel_Click()
'    Call DbAdoDisConnect
'
'    pctLogin.Visible = False
'    End
'End Sub

