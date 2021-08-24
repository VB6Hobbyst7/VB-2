VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form FrmViewopIlls 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "수술코드조회"
   ClientHeight    =   8325
   ClientLeft      =   5700
   ClientTop       =   1860
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8325
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread SS1 
      Height          =   1812
      Left            =   6288
      TabIndex        =   24
      Top             =   936
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1588
      _ExtentY        =   3196
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "opma08.frx":0000
   End
   Begin Threed.SSPanel PanelSearch 
      Height          =   840
      Left            =   60
      TabIndex        =   9
      Top             =   1740
      Width           =   6072
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   1482
      _StockProps     =   15
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Begin VB.CommandButton CmdSearch 
         Caption         =   "A"
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
         Index           =   0
         Left            =   72
         TabIndex        =   51
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "B"
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
         Index           =   1
         Left            =   444
         TabIndex        =   50
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "C"
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
         Index           =   2
         Left            =   828
         TabIndex        =   49
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "D"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   48
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "E"
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
         Index           =   4
         Left            =   1572
         TabIndex        =   47
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "F"
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
         Index           =   5
         Left            =   1944
         TabIndex        =   46
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "G"
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
         Index           =   6
         Left            =   2328
         TabIndex        =   45
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "H"
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
         Index           =   7
         Left            =   2700
         TabIndex        =   44
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "I"
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
         Index           =   8
         Left            =   3072
         TabIndex        =   43
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "J"
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
         Index           =   9
         Left            =   3444
         TabIndex        =   42
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "K"
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
         Index           =   10
         Left            =   3828
         TabIndex        =   41
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "L"
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
         Index           =   11
         Left            =   4200
         TabIndex        =   40
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "M"
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
         Index           =   12
         Left            =   4572
         TabIndex        =   39
         Top             =   24
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "N"
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
         Index           =   13
         Left            =   72
         TabIndex        =   38
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "O"
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
         Index           =   14
         Left            =   444
         TabIndex        =   37
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "P"
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
         Index           =   15
         Left            =   828
         TabIndex        =   36
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Q"
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
         Index           =   16
         Left            =   1200
         TabIndex        =   35
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "R"
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
         Index           =   17
         Left            =   1572
         TabIndex        =   34
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S"
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
         Index           =   18
         Left            =   1944
         TabIndex        =   33
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "T"
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
         Index           =   19
         Left            =   2328
         TabIndex        =   32
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "U"
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
         Index           =   20
         Left            =   2700
         TabIndex        =   31
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "V"
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
         Index           =   21
         Left            =   3072
         TabIndex        =   30
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "W"
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
         Index           =   22
         Left            =   3444
         TabIndex        =   29
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "X"
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
         Index           =   23
         Left            =   3828
         TabIndex        =   28
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Y"
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
         Index           =   24
         Left            =   4200
         TabIndex        =   27
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Z"
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
         Index           =   25
         Left            =   4572
         TabIndex        =   26
         Top             =   420
         Width           =   360
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0000FFFF&
         Caption         =   "ALL"
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
         Index           =   26
         Left            =   4980
         TabIndex        =   10
         Top             =   30
         Width           =   1050
      End
      Begin Threed.SSCommand cmdFav 
         Height          =   360
         Left            =   4968
         TabIndex        =   52
         Top             =   408
         Width           =   1056
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "개인등록"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel Panel 
      Height          =   372
      Left            =   60
      TabIndex        =   12
      Top             =   36
      Width           =   6072
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.99
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   36
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&1.개인"
         ForeColor       =   0
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
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   1
         Left            =   1056
         TabIndex        =   1
         Top             =   36
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&2.과별"
         ForeColor       =   0
         BackColor       =   16777152
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
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   2
         Top             =   36
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&3.전체"
         ForeColor       =   0
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
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   3
         Left            =   3036
         TabIndex        =   3
         Top             =   36
         Visible         =   0   'False
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&4.계통별"
         ForeColor       =   0
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
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   4
         Left            =   4020
         TabIndex        =   4
         Top             =   36
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&5.찾기"
         ForeColor       =   0
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
      Begin Threed.SSPanel PanelMenus 
         Height          =   300
         Index           =   5
         Left            =   5016
         TabIndex        =   5
         Top             =   36
         Width           =   1008
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "&9.종료"
         ForeColor       =   0
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
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   570
      TabIndex        =   8
      Top             =   4080
      Width           =   4992
   End
   Begin Threed.SSPanel PanelFind 
      Height          =   840
      Left            =   60
      TabIndex        =   13
      Top             =   2976
      Width           =   6072
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   1482
      _StockProps     =   15
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Begin VB.TextBox TxtFind 
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
         Left            =   1980
         TabIndex        =   6
         Top             =   240
         Width           =   2625
      End
      Begin VB.CommandButton CmdFindOK 
         BackColor       =   &H0000FFFF&
         Caption         =   "찾기 확인"
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
         Left            =   4710
         TabIndex        =   7
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label 
         Caption         =   "찾고자하는 상병명의 단어를 입력하세요."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1815
      End
   End
   Begin Threed.SSPanel PanelSet 
      Height          =   372
      Left            =   60
      TabIndex        =   15
      Top             =   456
      Width           =   6072
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.99
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSPanel PanelSort 
         Height          =   288
         Index           =   0
         Left            =   4512
         TabIndex        =   20
         Top             =   48
         Width           =   732
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "코드순"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   0
         Alignment       =   8
      End
      Begin Threed.SSPanel PanelSort 
         Height          =   288
         Index           =   1
         Left            =   5268
         TabIndex        =   21
         Top             =   48
         Width           =   732
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "수술순"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   288
         Left            =   48
         TabIndex        =   23
         Top             =   48
         Width           =   3480
         _Version        =   65536
         _ExtentX        =   6138
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "찾고자하는 상병을 선택하세요."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin Threed.SSCheck ROCheck 
         Height          =   225
         Left            =   3560
         TabIndex        =   22
         Top             =   75
         Visible         =   0   'False
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "R/O"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   19
         Top             =   90
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "양측"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   2
         Left            =   1960
         TabIndex        =   18
         Top             =   90
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "좌측"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   17
         Top             =   90
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "우측"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptSets 
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   90
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   " None"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin MSOutl.Outline OutlineIlls 
      Height          =   6420
      Left            =   96
      TabIndex        =   25
      Top             =   984
      Width           =   6060
      _Version        =   65536
      _ExtentX        =   10689
      _ExtentY        =   11324
      _StockProps     =   77
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
   Begin VB.Label LabelName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   60
      TabIndex        =   11
      Top             =   7740
      Width           =   6075
   End
End
Attribute VB_Name = "FrmViewopIlls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strSQL              As String

Dim i, j, k             As Integer
Dim nSELECT             As Integer  '1.상용, 2.과상용, 3.전체상병
Dim nSet                As Integer  '0.None, 1. R/O,   2.좌측,   3.우측,   4.양측
Dim nLoadOutLine        As Integer  'OutLine View 에 상병 Load Flag
Dim nSort               As Integer  '0:코드순, 1:상병순

Dim strKorEng           As String   '한글,영문 토글

Dim nIllsIndex          As Integer
Dim GnIllSort           As Integer
Dim GstrIllSort         As String

    Dim cillcodeF
    Dim cillcodeT

    Dim cDeptDr
    Dim cIllCode1
    Dim cIllCode2
    Dim cIllCode3


Sub OutLineIlls_Init()
    
    Dim i               As Integer
    
    OutlineIlls.Clear
    
    For i = 1 To SS1.DataRowCnt
        SS1.Row = i
        SS1.Col = 1
        If Trim(SS1.Text) = "0" Then
            Select Case strKorEng
                Case "KOR": SS1.Col = 2     '한글명
                Case Else:  SS1.Col = 5     '영문명
            End Select
            OutlineIlls.AddItem " " & Trim(SS1.Text), -1
            OutlineIlls.ItemData(OutlineIlls.ListCount - 1) = SS1.Row
        End If
        SS1.Col = 1
        If Trim(SS1.Text) > "0" Then Exit For
    Next i
    

End Sub


Private Sub Read_SubTitle()

    Dim nIndent             As Integer
    Dim nItemData           As Integer
    Dim nListIndex          As Integer
    Dim i                   As Integer
    Dim nCNT                As Integer
    Dim strFrom             As String * 3
    Dim strTo               As String * 3
    Dim strIllName          As String * 199
    
    On Error Resume Next
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    nListIndex = OutlineIlls.ListIndex
    nIndent = OutlineIlls.Indent(nListIndex)
    nItemData = OutlineIlls.ItemData(nListIndex)
        
    If OutlineIlls.HasSubItems(nListIndex) = False Then
        Select Case nIndent
            Case 1
                GoSub Read_Indent_1
            Case 2
                GoSub Read_Indent_2
            Case 3
                GoSub Read_Indent_3
            Case 4
                Exit Sub
        End Select
    End If
    
    If OutlineIlls.Expand(nListIndex) = True Then
        Do While OutlineIlls.Indent(nListIndex + 1) = (nIndent + 1)
            OutlineIlls.RemoveItem (nListIndex + 1)
        Loop
        OutlineIlls.PictureType(nListIndex) = outClosed
    Else
        OutlineIlls.Expand(nListIndex) = True
        OutlineIlls.PictureType(nListIndex) = outOpen
    End If
    

Exit Sub


'/-------------------------------------------------------------------------------------------/

Read_Indent_1:
    
    nCNT = 0
    
    For i = 1 To SS1.DataRowCnt
        SS1.Row = i
        SS1.Col = 1
        If Val(Trim(SS1.Text)) = nItemData Then
            Select Case strKorEng
                Case "KOR": SS1.Col = 2     '한글명
                Case Else:  SS1.Col = 5     '영문명
            End Select
            OutlineIlls.AddItem " " & Trim(SS1.Text)
            nCNT = nCNT + 1
            OutlineIlls.ItemData(nListIndex + nCNT) = SS1.Row
        End If
        SS1.Col = 1
        If Val(Trim(SS1.Text)) > nItemData Then Return
    Next i
    
    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_2:

    SS1.Row = OutlineIlls.ItemData(nListIndex)
    SS1.Col = 3:    strFrom = Trim(SS1.Text)
    SS1.Col = 4:    strTo = Trim(SS1.Text)
    
    cillcodeF = strFrom & "   "
    cillcodeT = strTo & "ZZZ"
    
    strSQL = ""
    strSQL = strSQL & " SELECT IllNameK, IllNameE, IllCode "
    strSQL = strSQL & "   FROM TWBAS_ILLS "
    strSQL = strSQL & "  WHERE IllClass = '3'         "
    strSQL = strSQL & "    AND IllCode >= '" & cillcodeF & "' "
    strSQL = strSQL & "    AND IllCode <= '" & cillcodeT & "' "
'    strSQL = strSQL & "    AND SUBSTR(IllCode, 1, 1) BETWEEN  '0' AND '9' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
    Do Until rs.EOF
        Select Case strKorEng
            Case "KOR": strIllName = rs.Fields("IllNameK").Value & ""   '한글명
            Case Else:  strIllName = rs.Fields("IllNameE").Value & ""   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & rs.Fields("IllCode").Value & "" 'IllCode 201
        rs.MoveNext
    Loop
    
    End If
    
    Return


'/-------------------------------------------------------------------------------------------/

Read_Indent_3:

    
    strFrom = MidH(OutlineIlls.List(nListIndex), 201)
    strTo = MidH(OutlineIlls.List(nListIndex), 201)
    
    cillcodeF = strFrom & "   "
    cillcodeT = strTo & "ZZZ"

    strSQL = " SELECT IllNameK, IllNameE, IllCode "
    strSQL = strSQL & " FROM TWBAS_ILLS "
    strSQL = strSQL & "WHERE IllClass = '3'         "
    strSQL = strSQL & "  AND IllCode >= '" & cillcodeF & "' "
    strSQL = strSQL & "  AND IllCode <= '" & cillcodeT & "' "
'    strSQL = strSQL & "  AND SUBSTR(IllCode, 1, 1) BETWEEN  '0' AND '9' "
    
    Result = AdoOpenSet(rs, strSQL)
        
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        Select Case strKorEng
            Case "KOR": strIllName = rs.Fields("IllNameK").Value   '한글명
            Case Else:  strIllName = rs.Fields("IllNameE").Value   '영문명
        End Select
        OutlineIlls.AddItem " " & strIllName & rs.Fields("IllCode").Value
        OutlineIlls.PictureType(nListIndex + i + 1) = outLeaf
        rs.MoveNext
    Loop
    
    Return

End Sub

Sub Read_Ills(argIndex, ArgDeptDr)

    Dim i                   As Integer
    Dim strDeptDr           As String * 6
    Dim strIllCode          As String * 8
    
    
    List1.Clear

    GoSub Option_Sql_Made
    GoSub Read_Ill
    
    Exit Sub
    

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    
    strDeptDr = ArgDeptDr
'    If Trim(strDeptDr) = "GY" Then strDeptDr = "OB"
    cDeptDr = strDeptDr
    cIllCode1 = CmdSearch(argIndex).Caption & "%"
    cIllCode2 = LCase(CmdSearch(argIndex).Caption) & "%"
    
    Select Case nSELECT
        Case 3:
            If nSort = 0 Then
                    strSQL = " SELECT IllCode, IllNameE      "
                    strSQL = strSQL & " FROM TWBAS_ILLS "
                    strSQL = strSQL & "WHERE ILLCLASS = '3' "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND  (IllCode Like '" & cIllCode1 & "' "
                    strSQL = strSQL & "   OR   IllCode Like '" & cIllCode2 & "') "
                End If
            Else
                    strSQL = " SELECT IllCode, IllNameE      "
                    strSQL = strSQL & " FROM TWBAS_ILLS "
                    strSQL = strSQL & "WHERE ILLCLASS = '3' "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND  (IllNameE Like '" & cIllCode1 & "' "
                    strSQL = strSQL & "   OR   IllNameE Like '" & cIllCode2 & "') "
                End If
            End If
        Case Else
            If nSort = 0 Then
                    strSQL = " SELECT A.IllCode, B.IllNameE     "
                    strSQL = strSQL & " FROM TWOCS_OILLDEF A,          "
                    strSQL = strSQL & "      TWBAS_ILLS B  "
'                    strSQL = strSQL & "WHERE A.DeptDr   = '" & cDeptDr & "' "
                    strSQL = strSQL & "WHERE A.DeptDr   = '" & GstrIdnumber & "' "
                    strSQL = strSQL & "  AND IllClass = '3'         "
                    strSQL = strSQL & "  AND A.IllCode  > ' '          "
                    strSQL = strSQL & "  AND A.IllCode  = B.IllCode(+) "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( B.IllCode Like '" & cIllCode1 & "' "
                    strSQL = strSQL & "   OR   B.IllCode Like '" & cIllCode2 & "') "
                End If
            Else
                    strSQL = " SELECT A.IllCode, B.IllNameE     "
                    strSQL = strSQL & " FROM TWOCS_OILLDEF A,          "
                    strSQL = strSQL & "      TWBAS_ILLS B  "
'                    strSQL = strSQL & "WHERE A.DeptDr   = '" & cDeptDr & "' "
                    strSQL = strSQL & "WHERE A.DeptDr   = '" & GstrIdnumber & "' "
                    strSQL = strSQL & "  AND IllClass = '3'         "
                    strSQL = strSQL & "  AND A.IllCode  > ' '          "
                    strSQL = strSQL & "  AND A.IllCode  = B.IllCode(+) "
                If CmdSearch(argIndex).Caption <> "ALL" Then
                    strSQL = strSQL & "  AND ( B.IllNameE Like '" & cIllCode1 & "' "
                    strSQL = strSQL & "   OR   B.IllNameE Like '" & cIllCode2 & "') "
                End If
            End If
    End Select
    
    If nSort = 0 Then
        strSQL = strSQL & " ORDER BY IllCode "
    Else
        strSQL = strSQL & " ORDER BY IllNameE "
    End If
            
    Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

'    Result = dosql(strSQL)
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        If Trim$(rs.Fields("IllNameE").Value & "") > "" Then
            strIllCode = rs.Fields("IllCode").Value & ""
            List1.AddItem strIllCode & rs.Fields("IllNameE").Value & ""
        End If
        rs.MoveNext
    Loop
    
    Return
    
        
End Sub


Private Sub CmdFav_Click()
    Dim strIllCode          As String
    Dim Msg, Style, Title, Response
    
    If List1.ListIndex = -1 Then Exit Sub
    
    strIllCode = Left$(Trim(List1.List(List1.ListIndex)), 6)
    If Trim(strIllCode) = "" Then Exit Sub
    
    Response = MsgBox("개인 코드로 등록 하시겠습니까??", _
             vbYesNo + vbDefaultButton2, "알림")
    
    If Response = vbYes Then
        strSQL = "INSERT INTO TWOCS_OILLDEF VALUES ( "
        strSQL = strSQL & "'" & Trim(GstrIdnumber) & "', '" & Trim(strIllCode) & "' ) "
 
        adoConnect.BeginTrans
        Result = AdoExecute(strSQL)
        If Result = True And Rowindicator >= 1 Then
            adoConnect.CommitTrans
        Else
            MsgBox "    Data가 입력되지 않았습니다.    "
            adoConnect.RollbackTrans
            Exit Sub
        End If
    End If
    
End Sub

Private Sub CmdFindOK_Click()
    
    Dim strIllCode          As String * 8
    Dim strFind             As String
    
    strFind = Trim$(TxtFind.Text)
    
    If Len(strFind) <= 1 Then MsgBox " 영숫자를 2자 이상 입력하십시요.": Exit Sub

'    Result = Execsql("Open Scope")
    List1.Clear
    
    GoSub Option_Sql_Made
    GoSub Read_Ill
    
'    Result = Execsql("Close Scope")
    
Exit Sub
    

'/----------------------------------------------------------------------------------------/

Option_Sql_Made:
    
    cIllCode1 = "%" & UCase(LeftH(strFind, 1)) & LCase(MidH(strFind, 2)) & "%"
    cIllCode2 = "%" & LCase(strFind) & "%"
    cIllCode3 = "%" & strFind & "%"
    
    strSQL = " SELECT Distinct A.IllCode, IllNameE "
    strSQL = strSQL & " FROM TWOCS_OILLDEF A, TWBAS_ILLS B "
    strSQL = strSQL & "WHERE ( IllNameE Like '" & cIllCode1 & "' "
    strSQL = strSQL & "   OR   IllNameE Like '" & cIllCode2 & "' "
    strSQL = strSQL & "   OR   IllNameE Like '" & cIllCode3 & "') "
    strSQL = strSQL & "  AND   B.IllClass = '3' "
    strSQL = strSQL & "  AND   A.IllCode  = B.IllCode     "
    strSQL = strSQL & " ORDER BY IllCode "
           
'    strSQL = " SELECT Distinct IllCode, IllNameE "
'    strSQL = strSQL & " FROM  TWBAS_ILLS B "
'    strSQL = strSQL & "WHERE ( IllNameE Like '" & cIllCode1 & "' "
'    strSQL = strSQL & "   OR   IllNameE Like '" & cIllCode2 & "' "
'    strSQL = strSQL & "   OR   IllNameE Like '" & cIllCode3 & "') "
'    strSQL = strSQL & "  AND SUBSTR(IllCode, 1, 1) BETWEEN  '0' AND '9' "
'    strSQL = strSQL & " ORDER BY IllCode "

Return
    

'/----------------------------------------------------------------------------------------/

Read_Ill:

'    Result = dosql(strSQL)
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        strIllCode = rs.Fields("IllCode").Value
        List1.AddItem strIllCode & rs.Fields("IllNameE").Value
        rs.MoveNext
    Loop
    
Return
    
End Sub

Private Sub cmdSearch_Click(Index As Integer)

    GnIllSort = Index
    
    List1.Clear
    
    Select Case nSELECT
        Case 1:     Call Read_Ills(Index, GstrIdnumber)
        Case 2:     Call Read_Ills(Index, GstrDeptCode)
        Case 3:     Call Read_Ills(Index, " ")
    End Select
    
End Sub

Private Sub Form_Activate()
    
    nSet = 0
    'OptSets(0).Value = True
'    GstrSELECTIllcode = ""
    Me.Refresh
'    GstrDeptCode = FrmAppoint.cboDeptCode
    
End Sub

Private Sub Form_Load()
    
    Me.Top = 100
    Me.Left = 5610
    Me.Width = 6300
    Me.Height = 8790
    
    OutlineIlls.Top = 885
    PanelSearch.Top = 885
    PanelFind.Top = 885
    
    OutlineIlls.Left = Panel.Left
    PanelSearch.Left = Panel.Left
    PanelFind.Left = Panel.Left
    
    List1.Top = 1800
    List1.Left = OutlineIlls.Left
    List1.Width = OutlineIlls.Width
    List1.Height = 5910
    
    OutlineIlls.Visible = False
    PanelFind.Visible = False
    nLoadOutLine = False
    SS1.Visible = False
    
    If Trim(GstrDeptCode) = "IM" Or Trim(GstrDeptCode) = "CS" Then
        Me.Caption = "수술코드조회 : 개인 상용상병중 조회"
        strKorEng = "ENG"   '영문 기본
        nSELECT = 1         '과   상용
        nSet = 0            '기본 조회
    
        Call PanelMenus_Click(0)
        PanelMenus(0).BackColor = RGB(128, 255, 255)
    Else
        Me.Caption = "수술코드조회 : 과 상용상병중 조회"
        strKorEng = "ENG"   '영문 기본
        nSELECT = 2         '과   상용
        nSet = 0            '기본 조회
        
    End If
    
    GnIllSort = 0
    
    If GstrIllSort = "명순" Then
        Call PanelSort_Click(1)
    Else
        Call PanelSort_Click(0)
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set FrmViewopIlls = Nothing
    
    Unload Me

End Sub

Private Sub LabelName_DblClick()
    
    If OutlineIlls.Visible = True Then
        Select Case strKorEng
            Case "KOR": strKorEng = "ENG"
            Case Else:  strKorEng = "KOR"
        End Select
        
        Call OutLineIlls_Init
    End If
    
End Sub


Private Sub List1_Click()
    
    If List1.ListIndex = -1 Then Exit Sub
    
    LabelName.Caption = List1.List(List1.ListIndex)
    
End Sub

Private Sub List1_DblClick()
    
    Dim i                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    
    If List1.ListIndex = -1 Then Exit Sub
    
    strIllCode = Trim(LeftH(List1.List(List1.ListIndex), 6))
    strIllNameE = Trim(MidH(List1.List(List1.ListIndex), 9, 80))
    
    FrmViewopIlls.Tag = strIllCode
    
    GstrOpillName = strIllNameE
    
    Me.Hide
    
    
End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call List1_DblClick

End Sub


Private Sub OptSets_Click(Index As Integer, Value As Integer)
    
    For i = 0 To 3
        OptSets(i).ForeColor = RGB(0, 0, 0)
    Next i
    
    OptSets(Index).ForeColor = RGB(0, 0, 255)
    nSet = Index
    
End Sub

Private Sub OutlineIlls_Click()

    Dim strIlls             As String * 10
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    strIlls = MidH(OutlineIlls.List(OutlineIlls.ListIndex), 201)
    
    If Trim(strIlls) = "" Then
        strIlls = "SESSION : "
    Else
        strIlls = Left(strIlls, 7) & " : "
    End If
    
    LabelName.Caption = strIlls & OutlineIlls.List(OutlineIlls.ListIndex)

End Sub

Private Sub OutlineIlls_DblClick()
    
    Dim i                   As Integer
    Dim strSetCode          As String * 6
    Dim strIllCode          As String * 6
    Dim strIllNameE         As String * 80
    
    If OutlineIlls.ListIndex = -1 Then Exit Sub
    
    Call Read_SubTitle
    
    If OutlineIlls.PictureType(OutlineIlls.ListIndex) = outLeaf Then GoSub Data_Send: Exit Sub
    If OutlineIlls.Indent(OutlineIlls.ListIndex) > 2 Then
        If OutlineIlls.PictureType(OutlineIlls.ListIndex) = outOpen Then
            If OutlineIlls.Indent(OutlineIlls.ListIndex) = OutlineIlls.Indent(OutlineIlls.ListIndex + 1) Then
                GoSub Data_Send
            End If
        End If
    End If
    
    Exit Sub
    
    
'/--------------------------------------------------------------------------------------/

Data_Send:

    strIllCode = Trim(MidH(OutlineIlls.List(OutlineIlls.ListIndex), 201, 6))
    
    FrmViewopIlls.Tag = strIllCode
    Me.Hide
    
    Return
    
'/-------------------------------------------------------------------------------------------/
End Sub


Private Sub OutlineIlls_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Call OutlineIlls_DblClick
    
End Sub


Private Sub PanelMenus_Click(Index As Integer)
    
    Dim strIllCode              As String * 6
    Dim nYESNO                  As Integer
    
    If Index = 3 And PanelMenus(Index).BackColor = RGB(128, 255, 255) Then
        Call OutLineIlls_Init
        Exit Sub
    End If
    
    If PanelMenus(Index).BackColor = RGB(128, 255, 255) Then Exit Sub
    LabelName.Caption = ""
    
    If Index < 5 Then
        PanelMenus(0).BackColor = RGB(192, 192, 192)
        PanelMenus(1).BackColor = RGB(192, 192, 192)
        PanelMenus(2).BackColor = RGB(192, 192, 192)
        PanelMenus(3).BackColor = RGB(192, 192, 192)
        PanelMenus(4).BackColor = RGB(192, 192, 192)
        PanelMenus(Index).BackColor = RGB(128, 255, 255)
    End If
    
    Select Case Index
        Case 0: GoSub Menu_Search_1     '개인 상용상병 조회
        Case 1: GoSub Menu_Search_2     '과별 상용상병 조회
        Case 2: GoSub Menu_Search_3     '전체 상병코드 조회
        Case 3: GoSub Menu_Search_4     '계통별 상병   조회
        Case 4: GoSub Menu_Search_5     '상병 단어별   찾기
        Case 5: Me.Hide
    End Select
    
    
    Exit Sub
    
'/---------------------------------------------------------------------/
Menu_Search_1:      '개인 상용상병 조회
    nSELECT = 1
    Me.Caption = "수술코드조회 : 개인 상용상병중 조회"
    cmdFav.Enabled = False
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    
    'GnIllSort = 26
    Call Read_Ills(GnIllSort, GstrIdnumber)
    
Return


Menu_Search_2:      '과별 상용상병 조회
    nSELECT = 2
    Me.Caption = "수술코드조회 : 과별 상용상병중 조회"
    cmdFav.Enabled = True
    CmdSearch(26).Enabled = True
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
    'GnIllSort = 26
    Call Read_Ills(GnIllSort, GstrDeptCode)
Return


Menu_Search_3:      '전체 상병코드 조회
    nSELECT = 3
    Me.Caption = "수술코드조회 : 전체 수술중 조회"
    cmdFav.Enabled = True
    CmdSearch(26).Enabled = False
    OutlineIlls.Visible = False
    PanelSearch.Visible = True
    PanelFind.Visible = False
    List1.Visible = True
'    GnIllSort = 0
    Call Read_Ills(GnIllSort, " ")
Return


Menu_Search_4:      '계통별 상병   조회
    Me.Caption = "수술코드조회 : 계통별 조회"
    PanelSearch.Visible = False
    OutlineIlls.Visible = True
    PanelFind.Visible = False
    List1.Visible = False
    Call OutLineIlls_Init
Return


Menu_Search_5:      '상병 단어별   찾기
    Me.Caption = "수술코드조회 : 수술 단어별 찾기"
    PanelSearch.Visible = False
    OutlineIlls.Visible = False
    PanelFind.Visible = True
    List1.Visible = True
    List1.Clear
    TxtFind.SetFocus
Return


End Sub

Private Sub PanelMenus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If PanelMenus(Index).BackColor <> RGB(128, 255, 255) Then
        PanelMenus(Index).BackColor = RGB(255, 255, 0)
    End If
    
End Sub


Private Sub PanelMenus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If PanelMenus(Index).BackColor <> RGB(128, 255, 255) Then
        PanelMenus(Index).BackColor = RGB(192, 192, 192)
    End If
    
End Sub

Private Sub PanelSort_Click(Index As Integer)

    nSort = Index
    If Index = 0 Then
        PanelSort(0).BackColor = &HFFFFC0
        PanelSort(1).BackColor = &HC0C0C0
        GstrIllSort = "코드순"
        CmdSearch(0).Caption = "0"
        CmdSearch(1).Caption = "1"
        CmdSearch(2).Caption = "2"
        CmdSearch(3).Caption = "3"
        CmdSearch(4).Caption = "4"
        CmdSearch(5).Caption = "5"
        CmdSearch(6).Caption = "6"
        CmdSearch(7).Caption = "7"
        CmdSearch(8).Caption = "8"
        CmdSearch(9).Caption = "9"
        CmdSearch(10).Visible = False
        CmdSearch(11).Visible = False
        CmdSearch(12).Visible = False
        CmdSearch(13).Visible = False
        CmdSearch(14).Visible = False
        CmdSearch(15).Visible = False
        CmdSearch(16).Visible = False
        CmdSearch(17).Visible = False
        CmdSearch(18).Visible = False
        CmdSearch(19).Visible = False
        CmdSearch(20).Visible = False
        CmdSearch(21).Visible = False
        CmdSearch(22).Visible = False
        CmdSearch(23).Visible = False
        CmdSearch(24).Visible = False
        CmdSearch(25).Visible = False
    
    Else
        PanelSort(1).BackColor = &HFFFFC0
        PanelSort(0).BackColor = &HC0C0C0
        GstrIllSort = "명순"
        CmdSearch(0).Caption = "A"
        CmdSearch(1).Caption = "B"
        CmdSearch(2).Caption = "C"
        CmdSearch(3).Caption = "D"
        CmdSearch(4).Caption = "E"
        CmdSearch(5).Caption = "F"
        CmdSearch(6).Caption = "G"
        CmdSearch(7).Caption = "H"
        CmdSearch(8).Caption = "I"
        CmdSearch(9).Caption = "J"
        CmdSearch(10).Visible = True
        CmdSearch(11).Visible = True
        CmdSearch(12).Visible = True
        CmdSearch(13).Visible = True
        CmdSearch(14).Visible = True
        CmdSearch(15).Visible = True
        CmdSearch(16).Visible = True
        CmdSearch(17).Visible = True
        CmdSearch(18).Visible = True
        CmdSearch(19).Visible = True
        CmdSearch(20).Visible = True
        CmdSearch(21).Visible = True
        CmdSearch(22).Visible = True
        CmdSearch(23).Visible = True
        CmdSearch(24).Visible = True
        CmdSearch(25).Visible = True
    
    End If
    
'    If Trim(GstrDeptCode) = "IM" Then
'        Call Read_Ills(GnIllSort, GstrIdnumber)
'    Else
        Call Read_Ills(GnIllSort, GstrDeptCode)   'GstrDrCode)
'    End If
    
End Sub

Private Sub TxtFind_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
    
End Sub


