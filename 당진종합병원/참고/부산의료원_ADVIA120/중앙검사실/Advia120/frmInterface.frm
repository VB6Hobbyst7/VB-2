VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Interface Program"
   ClientHeight    =   10440
   ClientLeft      =   240
   ClientTop       =   645
   ClientWidth     =   15240
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   15240
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   90
      TabIndex        =   10
      Top             =   750
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16960
      _Version        =   393216
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Interface"
      TabPicture(0)   =   "frmInterface.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "통계"
      TabPicture(2)   =   "frmInterface.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   9120
         Left            =   -74850
         TabIndex        =   36
         Top             =   360
         Width           =   14760
         Begin FPSpread.vaSpread vasSumTemp 
            Height          =   2535
            Left            =   2640
            TabIndex        =   50
            Top             =   4230
            Visible         =   0   'False
            Width           =   9465
            _Version        =   393216
            _ExtentX        =   16695
            _ExtentY        =   4471
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
            SpreadDesigner  =   "frmInterface.frx":0496
         End
         Begin VB.Frame Frame6 
            Caption         =   "[통계조회]"
            Height          =   735
            Left            =   180
            TabIndex        =   37
            Top             =   210
            Width           =   14385
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   9210
               Top             =   150
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdCSV 
               Caption         =   "Excel File 변환"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   7140
               TabIndex        =   40
               Top             =   210
               Width           =   1905
            End
            Begin VB.CommandButton cmdSugaClear 
               Caption         =   "화면초기화"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   5820
               TabIndex        =   39
               Top             =   210
               Width           =   1275
            End
            Begin VB.CommandButton cmdSumSch 
               Caption         =   "결과조회"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4620
               TabIndex        =   38
               Top             =   210
               Width           =   1155
            End
            Begin MSComCtl2.DTPicker dtpSumSDate 
               Height          =   315
               Left            =   1110
               TabIndex        =   41
               Top             =   270
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   67305473
               CurrentDate     =   40780
            End
            Begin MSComCtl2.DTPicker dtpSumEDate 
               Height          =   315
               Left            =   2940
               TabIndex        =   42
               Top             =   270
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   67305473
               CurrentDate     =   40780
            End
            Begin VB.Label Label7 
               Caption         =   "-"
               Height          =   225
               Left            =   2730
               TabIndex        =   44
               Top             =   330
               Width           =   135
            End
            Begin VB.Label Label6 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   180
               TabIndex        =   43
               Top             =   330
               Width           =   915
            End
         End
         Begin FPSpread.vaSpread vasSum 
            Height          =   7875
            Left            =   180
            TabIndex        =   45
            Top             =   1080
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   13891
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":06DA
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9120
         Left            =   -74850
         TabIndex        =   25
         Top             =   360
         Width           =   14760
         Begin FPSpread.vaSpread vasResTemp 
            Height          =   2355
            Left            =   420
            TabIndex        =   49
            Top             =   6120
            Visible         =   0   'False
            Width           =   11265
            _Version        =   393216
            _ExtentX        =   19870
            _ExtentY        =   4154
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
            SpreadDesigner  =   "frmInterface.frx":3697
         End
         Begin VB.CommandButton cmdVasListWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   48
            Top             =   1110
            Width           =   405
         End
         Begin VB.CheckBox ChkAll 
            Height          =   255
            Left            =   720
            TabIndex        =   35
            Top             =   1170
            Width           =   225
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   7875
            Left            =   180
            TabIndex        =   47
            Top             =   1080
            Width           =   6375
            _Version        =   393216
            _ExtentX        =   11245
            _ExtentY        =   13891
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":38DB
         End
         Begin VB.Frame Frame4 
            Caption         =   "[검사결과조회]"
            Height          =   735
            Left            =   180
            TabIndex        =   26
            Top             =   210
            Width           =   14385
            Begin VB.TextBox txtBarcode 
               Height          =   315
               Left            =   11760
               TabIndex        =   56
               Top             =   270
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.ComboBox cmbTransGubun 
               Height          =   315
               ItemData        =   "frmInterface.frx":6C1F
               Left            =   3330
               List            =   "frmInterface.frx":6C2C
               TabIndex        =   30
               Text            =   "전체"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton cmdCall 
               Caption         =   "데이터 불러오기"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4860
               TabIndex        =   29
               Top             =   210
               Width           =   1815
            End
            Begin VB.CommandButton cmdListClear 
               Caption         =   "화면초기화"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   6720
               TabIndex        =   28
               Top             =   210
               Width           =   1275
            End
            Begin VB.CommandButton cmdListTrans 
               Caption         =   "검사결과수동전송"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8040
               TabIndex        =   27
               Top             =   210
               Width           =   1905
            End
            Begin MSComCtl2.DTPicker dtpExamDate 
               Height          =   315
               Left            =   1110
               TabIndex        =   31
               Top             =   270
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   67305473
               CurrentDate     =   40780
            End
            Begin VB.Label Label4 
               Caption         =   "Barcode 검색"
               Height          =   225
               Left            =   10380
               TabIndex        =   34
               Top             =   330
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label Label2 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   180
               TabIndex        =   33
               Top             =   330
               Width           =   915
            End
            Begin VB.Label Label3 
               Caption         =   "구분"
               Height          =   225
               Left            =   2820
               TabIndex        =   32
               Top             =   330
               Width           =   555
            End
         End
         Begin FPSpread.vaSpread vasListRes 
            Height          =   7875
            Left            =   6750
            TabIndex        =   54
            Top             =   1080
            Width           =   7815
            _Version        =   393216
            _ExtentX        =   13785
            _ExtentY        =   13891
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   8
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":6C44
         End
      End
      Begin VB.Frame Frame3 
         Height          =   9120
         Left            =   150
         TabIndex        =   16
         Top             =   360
         Width           =   14760
         Begin VB.CommandButton cmdEquipConnect 
            Caption         =   "장비연결"
            Height          =   405
            Left            =   3780
            TabIndex        =   60
            Top             =   270
            Width           =   1695
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1215
            Left            =   8160
            TabIndex        =   59
            Top             =   2700
            Visible         =   0   'False
            Width           =   3435
            _Version        =   393216
            _ExtentX        =   6059
            _ExtentY        =   2143
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
            SpreadDesigner  =   "frmInterface.frx":7682
         End
         Begin VB.TextBox txtReceBarcode 
            Height          =   315
            Left            =   9840
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.TextBox txtData 
            Height          =   1215
            Left            =   11580
            TabIndex        =   53
            Top             =   6600
            Visible         =   0   'False
            Width           =   2715
         End
         Begin FPSpread.vaSpread vasOrderBuf 
            Height          =   1215
            Left            =   7200
            TabIndex        =   52
            Top             =   6600
            Visible         =   0   'False
            Width           =   4395
            _Version        =   393216
            _ExtentX        =   7752
            _ExtentY        =   2143
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
            SpreadDesigner  =   "frmInterface.frx":78C6
         End
         Begin FPSpread.vaSpread vasOrder 
            Height          =   1215
            Left            =   7200
            TabIndex        =   51
            Top             =   5400
            Visible         =   0   'False
            Width           =   4395
            _Version        =   393216
            _ExtentX        =   7752
            _ExtentY        =   2143
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
            SpreadDesigner  =   "frmInterface.frx":BD8C
         End
         Begin VB.CommandButton cmdVasIDWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   46
            Top             =   810
            Width           =   405
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   1155
            Left            =   12720
            TabIndex        =   22
            Top             =   7800
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Height          =   1125
            Left            =   7380
            TabIndex        =   21
            Top             =   7860
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.TextBox txtBuff 
            Height          =   1215
            Left            =   11580
            TabIndex        =   20
            Top             =   5400
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "화면초기화"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   180
            TabIndex        =   19
            Top             =   270
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Trans 
            Caption         =   "검사결과수동전송"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1650
            TabIndex        =   18
            Top             =   270
            Width           =   2085
         End
         Begin VB.CheckBox chkA 
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   870
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   8175
            Left            =   180
            TabIndex        =   23
            Top             =   780
            Width           =   6375
            _Version        =   393216
            _ExtentX        =   11245
            _ExtentY        =   14420
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":10252
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8175
            Left            =   6750
            TabIndex        =   24
            Top             =   780
            Width           =   7815
            _Version        =   393216
            _ExtentX        =   13785
            _ExtentY        =   14420
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   8
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":135A6
         End
         Begin VB.Label lblConnect 
            Height          =   285
            Left            =   5610
            TabIndex        =   61
            Top             =   390
            Width           =   3165
         End
         Begin VB.Label lblMT 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8970
            TabIndex        =   58
            Top             =   420
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label5 
            Caption         =   "BARCODE : "
            Height          =   285
            Left            =   8790
            TabIndex        =   55
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
         End
      End
   End
   Begin Threed.SSPanel sspMode 
      Height          =   525
      Left            =   2040
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   926
      _StockProps     =   15
      Caption         =   "전송모드"
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   10.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2700
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      RThreshold      =   1
      EOFEnable       =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   15030
      _Version        =   65536
      _ExtentX        =   26511
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "INTERFACE"
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.Timer Timer2 
         Left            =   6990
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   30000
         Left            =   5700
         Top             =   60
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   9000
         TabIndex        =   15
         Top             =   180
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8640
         Picture         =   "frmInterface.frx":13FE4
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   14
         Top             =   180
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   11760
         TabIndex        =   13
         Top             =   120
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21364736
         CurrentDate     =   40778
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Height          =   195
         Left            =   10740
         TabIndex        =   12
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
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
         Height          =   225
         Left            =   5310
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6585
      Left            =   -840
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   8835
      Begin VB.TextBox txtMsg 
         ForeColor       =   &H000000C0&
         Height          =   825
         Left            =   7830
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   3300
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtErr 
         Height          =   1035
         Left            =   4440
         TabIndex        =   8
         Top             =   5100
         Width           =   1935
      End
      Begin VB.TextBox txtDate 
         Height          =   405
         Left            =   330
         TabIndex        =   5
         Top             =   1260
         Width           =   2325
      End
      Begin VB.TextBox txtAll 
         Height          =   375
         Left            =   300
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txtTemp 
         Height          =   375
         Left            =   300
         TabIndex        =   3
         Top             =   450
         Width           =   2055
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   4455
         Left            =   4020
         TabIndex        =   6
         Top             =   0
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":1456E
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   2535
         Left            =   150
         TabIndex        =   7
         Top             =   2130
         Visible         =   0   'False
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   4471
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
         SpreadDesigner  =   "frmInterface.frx":18A98
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "메인"
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuConf 
      Caption         =   "설정"
      Begin VB.Menu mnuCodeConfig 
         Caption         =   "코드설정"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "통신설정"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "전송"
      Begin VB.Menu mnuAuto 
         Caption         =   "자동전송"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuManual 
         Caption         =   "수동전송"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu subUp 
         Caption         =   "검체번호 수정"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu subDel 
         Caption         =   "검체결과 삭제"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colPID = 3
Const colPName = 4
Const colRcnt = 5
Const colState = 6
Const colRStart = 6

' 장비코드 검사코드 검사명 수치결과 문자결과 seq
Const colEquipExam = 1
Const colExamCode = 2
Const colExamName = 3
Const colResValue = 4
Const colResult = 5
Const colSeq = 6
Const colResDate = 7
Const colResTime = 8

Public gRow As Long
Dim sOrder As String
Dim ConfirmData As String
Dim sSampleType As String
Dim lsFlag As String
Dim llRow As Long
Dim gMT As String           'Message Toggle
Dim gErrState As Long
Dim gComState As Long

Function LRC(ByVal asData As String) As String
'Longitudinal Redundancy Check

    Dim i As Integer
    Dim a
    
    a = Asc(Left(asData, 1))
    
    For i = 2 To Len(asData)
        a = a Xor Asc(Mid(asData, i, 1))
    Next i
    
    If a = 3 Then a = 127
    
    LRC = Chr(a)
End Function

Function Advia_IDSet(asID As String) As String
    '14자리
    '문경제일병원 - 바코드 12자리
    Advia_IDSet = "00" & asID
End Function

Function Advia_NoOrder(asID As String) As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    
    lsData = gMT & "N R " & Advia_IDSet(asID) & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 3
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    'save_raw_Data "[TX]" & lsData
End Function

Function Advia_Init() As String
'Initialization

    Dim lsData As String
    
    lblConnect.Caption = "연결대기중.."
    
    gMT = "0"
    gErrState = 0
    
    lsData = gMT & "I " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 0
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    Save_Raw_Data "[Tx]" & lsData
End Function

Function Advia_Token() As String
'Token Message

    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    If gMT = "" Then gMT = "0"
    
    lsData = gMT & "S          " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 1
    
    lblMT.Caption = gMT
    DoSleep 1
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    Save_Raw_Data "[Tx]" & lsData
End Function

Function Advia_Token_1() As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    If gMT = "" Then gMT = "0"
    
    lsData = "S          " & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 1
    
    lblMT.Caption = gMT
    DoSleep 1
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    Save_Raw_Data "[Tx]" & lsData
End Function

Function Advia_ResValid() As String
    Dim lsData As String
    
    gMT = Chr(Asc(gMT) + 1)
    If gMT > "Z" Then gMT = "0"
    
    lsData = gMT & "Z   " & Space(6) & " " & Space(6) & " " & " 0" & chrCR & chrLF
    lsData = chrSTX & lsData & LRC(lsData) & chrETX
    
    gComState = 4
    
    MSComm1.Output = lsData
    Timer1.Enabled = True
    'save_raw_Data "[TX]" & lsData
End Function


Private Sub chkA_Click()
    Dim iRow As Integer
    
    If chkA.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkA.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub ChkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmd_Trans_Click()
'선택전송
Dim VasidRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If txtUID.Text = "" Then
        MsgBox "사용자 확인을 해 주십시오"
        txtUID.SetFocus
        Exit Sub
    End If
    
    If vasID.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
    Connect_Server
    For VasidRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = VasidRow
        
        If vasID.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
            If Barcode_Gubun(Trim(GetText(vasID, VasidRow, colBarCode))) = "Q" Then
                liRet = Insert_QC_Data(vasID, VasidRow)
            Else
                liRet = Insert_Data(VasidRow)
            End If
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", VasidRow, colState
            Else
                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", VasidRow, colState
            End If
            vasID.Col = 1
            vasID.Row = VasidRow
            vasID.Value = 0
        Else
        
        End If
    Next VasidRow
    
End Sub

Function Insert_Data(argSpcRow As Integer) As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim lsID        As String
    Dim sResult     As String
    Dim sResult1    As String       '수치결과
    Dim sResult2    As String       '문자결과
    
    Dim iPos        As Integer
    Dim iPos1       As Integer
    Dim sORD_CD     As String
    Dim sSPCCD      As String
    Dim sSEQ_NO     As String
    
    Dim sDecision   As String
    Dim sPanicFlag  As String
    Dim sDeltaFlag  As String
    Dim sDPA_GB     As String       'DELTA/PANIC 동시 발생시 ('DP'로 변경)
    
    Dim sCnt        As String
    
    Dim sResultCD   As String
    Dim sAllResult  As String
    Dim sEquipCode  As String
    Dim sReceCode   As String
    Dim sTransDate As String
    Dim sTransTime As String
    
    Dim sRsltSqno As String
    Dim sResValue As String
    Dim sRcpnSqno As String
    Dim sExamCode As String
    Dim sResGubun As String
    Dim sTransRes As String
        
    
    Insert_Data = -1
    
    lsID = ""
    lsID = Trim(GetText(vasID, argSpcRow, colBarCode))
    
    sTransDate = Format(GetDateFull, "yyyymmdd")
    sTransTime = Format(GetDateFull, "hhmmss")
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasTemp
    
    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun " & vbCrLf & _
          " From pat_res a, equipexam b " & vbCrLf & _
          " Where a.equipno = b.equipno " & vbCrLf & _
          " And a.examcode = b.examcode " & vbCrLf & _
          " And a.equipcode = b.equipcode " & vbCrLf & _
          " And a.equipno = '" & gEquip & "' " & vbCrLf & _
          " And a.barcode = '" & lsID & "' "
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    sCnt = ""
        
    cn_Ser.BeginTrans
    '서버로 결과값 저장하기
    For i = 1 To vasTemp.DataRowCnt
        
            
        sExamCode = Trim(GetText(vasTemp, i, 2))
        sResValue = Trim(GetText(vasTemp, i, 3))
        sResult = Trim(GetText(vasTemp, i, 4))
        sResGubun = Trim(GetText(vasTemp, i, 5))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResValue & "(" & sResult & ")"
            
        Else
            sTransRes = sResValue
            sResult = ""
        End If
        
        
        If sExamCode <> "" And sResValue <> "" Then

            SQL = "SELECT A.SPCM_NO, AA.RSLT_SQNO, A.RCPN_SQNO " & vbCrLf & _
                  "FROM MS.MSLRCPT A " & vbCrLf & _
                  "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
                  "WHERE A.SPCM_NO = '" & lsID & "' " & vbCrLf & _
                  "  AND AA.EXMN_CD = '" & sExamCode & "'"
            
            res = db_select_Col(gServer, SQL)
            If res = -1 Then
                Save_Raw_Data "[QueryErr]" & SQL
                Exit Function
                
            End If
            
            If res > 0 Then
            
                sRsltSqno = Trim(gReadBuf(1))
                sRcpnSqno = Trim(gReadBuf(2))
                '/아래 조건이 어긋나면 전송 취소
                If Trim(sRsltSqno) <> "" And Trim(sRcpnSqno) <> "" Then
                
                    SQL = "select eqpm_rslt_valu from mslintrslt " & vbCrLf & _
                              " where rslt_sqno = '" & sRsltSqno & "' "
                        res = db_select_Col(gServer, SQL)
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                        
                     If res > 0 Then
                                        SQL = "UPDATE MSLINTRSLT"
                         SQL = SQL & vbCrLf & "   SET EQPM_RSLT_VALU = '" & sResValue & "'"
                         SQL = SQL & vbCrLf & "      ,INIT_EQPM_RSLT_VALU = '" & sTransRes & "'"
                         SQL = SQL & vbCrLf & "      ,RSLT_PRGR_STAT_CD = '07'"
                         SQL = SQL & vbCrLf & "      ,LAST_UPDT_USID = '" & gExamUID & "'"
                         SQL = SQL & vbCrLf & "      ,LAST_UDDT = SYSDATE "
                         SQL = SQL & vbCrLf & " WHERE RSLT_SQNO = '" & sRsltSqno & "'"
                         SQL = SQL & vbCrLf & "   AND EQPM_RCPN_SQNO = '" & sRcpnSqno & "'"
                         SQL = SQL & vbCrLf & "   AND RSLT_PRGR_STAT_CD < '11'"
        
                         res = SendQuery(gServer, SQL)
                         If res = -1 Then
                             Save_Raw_Data "[QueryErr]" & SQL
                             cn_Ser.RollbackTrans
                             Exit Function
                         End If
                     ElseIf res = 0 Then
                         SQL = "insert into mslintrslt (rslt_sqno, rslt_trms_date, rslt_trms_time, eqpm_cd, eqpm_rslt_valu, " & vbCrLf & _
                               "eqpm_rslt_dvcd, err_valu, init_eqpm_rslt_valu, updt_eqpm_rslt_valu, eqpm_rslt_rmrk, " & vbCrLf & _
                               "eqpm_rcpn_sqno, rslt_prgr_stat_cd, frst_rgst_usid, frst_rgdt, last_updt_usid, last_uddt) " & vbCrLf & _
                               "values( " & vbCrLf & _
                               "'" & sRsltSqno & "','" & sResValue & "','" & sTransTime & "', " & vbCrLf & _
                               "'" & gEquip & "','" & sTransRes & "', " & vbCrLf & _
                               "'','','" & sTransRes & "', " & vbCrLf & _
                               "'','', " & vbCrLf & _
                               "'" & sRcpnSqno & "','07', '" & gExamUID & "', " & vbCrLf & _
                               "SYSDATE,'" & gExamUID & "',SYSDATE " & vbCrLf & _
                               ") "
                         res = SendQuery(gServer, SQL)
                         If res = -1 Then
                             Save_Raw_Data "[QueryErr]" & SQL
                             cn_Ser.RollbackTrans
                             Exit Function
                             
                         End If
                    End If
                    
                    SQL = "UPDATE MS.MSLGNRLRSLT " & vbCrLf & _
                          "SET    RSLT_PRGR_STAT_CD = '07',  --결과저장(예비결과)  " & vbCrLf & _
                          "       NMVL_RSLT_VALU = '" & sResValue & "',  " & vbCrLf & _
                          "       TXT_RSLT_VALU = '" & sTransRes & "', " & vbCrLf & _
                          "       NRML_DVCD = '', " & vbCrLf & _
                          "       DELT_YN = '', " & vbCrLf & _
                          "       PANC_YN = '', " & vbCrLf & _
                          "       ALRT_YN = '', " & vbCrLf & _
                          "       EXMN_RSLT_STOR_DATE = TO_CHAR(SYSDATE, 'YYYYMMDD'), " & vbCrLf & _
                          "       EXMN_RSLT_STOR_TIME = TO_CHAR(SYSDATE, 'HH24MISS'), " & vbCrLf & _
                          "       EXMN_RSLT_STOR_PRSN_ID = '" & gExamUID & "', " & vbCrLf & _
                          "       LAST_UPDT_USID = '" & gExamUID & "', " & vbCrLf & _
                          "       LAST_UDDT = SYSTIMESTAMP, EXMN_EQPM_CD = '" & gEquip & "'  " & vbCrLf & _
                          " WHERE RSLT_SQNO = '" & sRsltSqno & "'  AND RSLT_PRGR_STAT_CD <> '11' "
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                        
                    End If
                    
                    SQL = "UPDATE MS.MSLRCPT " & vbCrLf & _
                          " SET   exmn_prgr_stat_cd = '07', " & vbCrLf & _
                          "        last_updt_usid = '" & gExamUID & "', " & vbCrLf & _
                          "        last_uddt = SYSTIMESTAMP " & vbCrLf & _
                          "  WHERE RCPN_SQNO = '" & sRcpnSqno & "' "
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                        
                    End If
                End If
            End If
            
        End If
        DoSleep 50
    Next i
    cn_Ser.CommitTrans
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data = 1
    
End Function

Function Insert_Data_1(argSpcRow As Integer) As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim lsID        As String
    Dim sResult     As String
    Dim sResult1    As String       '수치결과
    Dim sResult2    As String       '문자결과
    
    Dim iPos        As Integer
    Dim iPos1       As Integer
    Dim sORD_CD     As String
    Dim sSPCCD      As String
    Dim sSEQ_NO     As String
    
    Dim sDecision   As String
    Dim sPanicFlag  As String
    Dim sDeltaFlag  As String
    Dim sDPA_GB     As String       'DELTA/PANIC 동시 발생시 ('DP'로 변경)
    
    Dim sCnt        As String
    
    Dim sResultCD   As String
    Dim sAllResult  As String
    Dim sEquipCode  As String
    Dim sReceCode   As String
    Dim sTransDate As String
    Dim sTransTime As String
    
    Dim sRsltSqno As String
    Dim sResValue As String
    Dim sRcpnSqno As String
    Dim sExamCode As String
    Dim sResGubun As String
    Dim sTransRes As String
        
    
    Insert_Data_1 = -1
    
    lsID = ""
    lsID = Trim(GetText(vasList, argSpcRow, colBarCode))
    
    sTransDate = Format(GetDateFull, "yyyymmdd")
    sTransTime = Format(GetDateFull, "hhmmss")
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasTemp
    
    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun " & vbCrLf & _
          " From pat_res a, equipexam b " & vbCrLf & _
          " Where a.equipno = b.equipno " & vbCrLf & _
          " And a.examcode = b.examcode " & vbCrLf & _
          " And a.equipcode = b.equipcode " & vbCrLf & _
          " And a.equipno = '" & gEquip & "' " & vbCrLf & _
          " And a.barcode = '" & lsID & "' "
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    sCnt = ""
    cn_Ser.BeginTrans
    '서버로 결과값 저장하기
    For i = 1 To vasTemp.DataRowCnt
        
            
        sExamCode = Trim(GetText(vasTemp, i, 2))
        sResValue = Trim(GetText(vasTemp, i, 3))
        sResult = Trim(GetText(vasTemp, i, 4))
        sResGubun = Trim(GetText(vasTemp, i, 5))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResValue & "(" & sResult & ")"
            
        Else
            sTransRes = sResValue
            sResult = ""
        End If
        
        
        If sExamCode <> "" And sResValue <> "" Then

            SQL = "SELECT A.SPCM_NO, AA.RSLT_SQNO, A.RCPN_SQNO " & vbCrLf & _
                  "FROM MS.MSLRCPT A " & vbCrLf & _
                  "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
                  "WHERE A.SPCM_NO = '" & lsID & "' " & vbCrLf & _
                  "  AND AA.EXMN_CD = '" & sExamCode & "'"
            
            res = db_select_Col(gServer, SQL)
            If res = -1 Then
                Save_Raw_Data "[QueryErr]" & SQL
                Exit Function
                
            End If
            
            If res > 0 Then
            
                sRsltSqno = Trim(gReadBuf(1))
                sRcpnSqno = Trim(gReadBuf(2))
                '/아래 조건이 어긋나면 전송 취소
                If Trim(sRsltSqno) <> "" And Trim(sRcpnSqno) <> "" Then
                
                    SQL = "select eqpm_rslt_valu from mslintrslt " & vbCrLf & _
                              " where rslt_sqno = '" & sRsltSqno & "' "
                        res = db_select_Col(gServer, SQL)
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                        
                     If res > 0 Then
                                        SQL = "UPDATE MSLINTRSLT"
                         SQL = SQL & vbCrLf & "   SET EQPM_RSLT_VALU = '" & sResValue & "'"
                         SQL = SQL & vbCrLf & "      ,INIT_EQPM_RSLT_VALU = '" & sTransRes & "'"
                         SQL = SQL & vbCrLf & "      ,RSLT_PRGR_STAT_CD = '07'"
                         SQL = SQL & vbCrLf & "      ,LAST_UPDT_USID = '" & gExamUID & "'"
                         SQL = SQL & vbCrLf & "      ,LAST_UDDT = SYSDATE "
                         SQL = SQL & vbCrLf & " WHERE RSLT_SQNO = '" & sRsltSqno & "'"
                         SQL = SQL & vbCrLf & "   AND EQPM_RCPN_SQNO = '" & sRcpnSqno & "'"
                         SQL = SQL & vbCrLf & "   AND RSLT_PRGR_STAT_CD < '11'"
        
                         res = SendQuery(gServer, SQL)
                         If res = -1 Then
                             Save_Raw_Data "[QueryErr]" & SQL
                             cn_Ser.RollbackTrans
                             Exit Function
                         End If
                     ElseIf res = 0 Then
                         SQL = "insert into mslintrslt (rslt_sqno, rslt_trms_date, rslt_trms_time, eqpm_cd, eqpm_rslt_valu, " & vbCrLf & _
                               "eqpm_rslt_dvcd, err_valu, init_eqpm_rslt_valu, updt_eqpm_rslt_valu, eqpm_rslt_rmrk, " & vbCrLf & _
                               "eqpm_rcpn_sqno, rslt_prgr_stat_cd, frst_rgst_usid, frst_rgdt, last_updt_usid, last_uddt) " & vbCrLf & _
                               "values( " & vbCrLf & _
                               "'" & sRsltSqno & "','" & sResValue & "','" & sTransTime & "', " & vbCrLf & _
                               "'" & gEquip & "','" & sTransRes & "', " & vbCrLf & _
                               "'','','" & sTransRes & "', " & vbCrLf & _
                               "'','', " & vbCrLf & _
                               "'" & sRcpnSqno & "','07', '" & gExamUID & "', " & vbCrLf & _
                               "SYSDATE,'" & gExamUID & "',SYSDATE " & vbCrLf & _
                               ") "
                         res = SendQuery(gServer, SQL)
                         If res = -1 Then
                             Save_Raw_Data "[QueryErr]" & SQL
                             cn_Ser.RollbackTrans
                             Exit Function
                             
                         End If
                    End If
                    
                    SQL = "UPDATE MS.MSLGNRLRSLT " & vbCrLf & _
                          "SET    RSLT_PRGR_STAT_CD = '07',  --결과저장(예비결과)  " & vbCrLf & _
                          "       NMVL_RSLT_VALU = '" & sResValue & "',  " & vbCrLf & _
                          "       TXT_RSLT_VALU = '" & sTransRes & "', " & vbCrLf & _
                          "       NRML_DVCD = '', " & vbCrLf & _
                          "       DELT_YN = '', " & vbCrLf & _
                          "       PANC_YN = '', " & vbCrLf & _
                          "       ALRT_YN = '', " & vbCrLf & _
                          "       EXMN_RSLT_STOR_DATE = TO_CHAR(SYSDATE, 'YYYYMMDD'), " & vbCrLf & _
                          "       EXMN_RSLT_STOR_TIME = TO_CHAR(SYSDATE, 'HH24MISS'), " & vbCrLf & _
                          "       EXMN_RSLT_STOR_PRSN_ID = '" & gExamUID & "', " & vbCrLf & _
                          "       LAST_UPDT_USID = '" & gExamUID & "', " & vbCrLf & _
                          "       LAST_UDDT = SYSTIMESTAMP, EXMN_EQPM_CD = '" & gEquip & "'  " & vbCrLf & _
                          " WHERE RSLT_SQNO = '" & sRsltSqno & "' AND RSLT_PRGR_STAT_CD <> '11' "
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                        
                    End If
                    
                    SQL = "UPDATE MS.MSLRCPT " & vbCrLf & _
                          " SET   exmn_prgr_stat_cd = '07', " & vbCrLf & _
                          "        last_updt_usid = '" & gExamUID & "', " & vbCrLf & _
                          "        last_uddt = SYSTIMESTAMP " & vbCrLf & _
                          "  WHERE RCPN_SQNO = '" & sRcpnSqno & "' "
                    res = SendQuery(gServer, SQL)
                    
                    If res = -1 Then
                        Save_Raw_Data "[QueryErr]" & SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                        
                    End If
                End If
            End If
            
        End If
        DoSleep 50
    Next i
    cn_Ser.CommitTrans
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data_1 = 1
    
End Function

Private Sub cmdCall_Click()
    Dim i As Long
    Dim varSendFlag
    Dim j As Long
    Dim x As Long
    Dim strResult As String
    
    
    ClearSpread vasList
    
    varSendFlag = cmbTransGubun.ListIndex

    SQL = "select '', barcode, pid, pname, count(result), sendflag from pat_res " & vbCrLf & _
          " where equipno = '" & gEquip & "' and examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    Else
        SQL = SQL & " and sendflag <> '0' "
    End If
    
    SQL = SQL & vbCrLf & " group by barcode, pid, pname,  sendflag"
    res = db_select_Vas(gLocal, SQL, vasList)

    
    vasList.MaxRows = vasList.DataRowCnt
    For i = 1 To vasList.DataRowCnt
        If GetText(vasList, i, colState) = "1" Then
            SetText vasList, "Result", i, colState
            
        ElseIf GetText(vasList, i, colState) = "2" Then
            SetText vasList, "Trans", i, colState
            SetBackColor vasList, i, i, colBarCode, colState, 255, 255, 180
        End If
    Next
    
    ClearSpread vasResTemp
    
    SQL = "select barcode, equipcode, resvalue, result from pat_res " & vbCrLf & _
          " where equipno = '" & gEquip & "' and examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    Else
        SQL = SQL & " and sendflag <> '0' "
    End If
    
    SQL = SQL & vbCrLf & " group by barcode, equipcode, resvalue, result"
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
'''    gArr_Exam(i, 1) = i    '순서
'''    gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
'''    gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
    
    For i = 1 To vasResTemp.DataRowCnt
        For j = 1 To vasList.DataRowCnt
            If Trim(GetText(vasResTemp, i, 1)) = Trim(GetText(vasList, j, colBarCode)) Then
                For x = 1 To vasList.MaxCols - colRStart
                    If Trim(GetText(vasResTemp, i, 2)) = Trim(gArr_Exam(x, 2)) Then
                        If gArr_Exam(x, 4) = "0" Then
                            strResult = Trim(GetText(vasResTemp, i, 3))
                        ElseIf gArr_Exam(x, 4) = "1" Then
                            strResult = Trim(GetText(vasResTemp, i, 4)) & "(" & Trim(GetText(vasResTemp, i, 3)) & ")"
                        End If
                        
                        SetText vasList, strResult, j, colRStart + CCur(gArr_Exam(x, 1))
                        Exit For
                    End If
                Next x
                Exit For
            End If
        Next j
    Next i

End Sub

Private Sub cmdClear_Click()
Dim iNumber As Integer
Dim i As Integer
    
    txtMsg.Text = ""
    
'''    ClearSpread vasID, 1, 1
'''    vasID.MaxRows = 0
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
    
    For i = vasID.DataRowCnt To 1 Step -1
        vasID.Col = colCheckBox
        vasID.Row = i
        If vasID.Value = 1 Then
            DeleteRow vasID, i, i
        End If
    Next
    
'''    Advia_Init
End Sub

Private Sub cmdCSV_Click()
    Dim i As Long
    Dim j As Long
    Dim strCSV As String
    Dim strFileName As String
    Dim FilNum
    
    CommonDialog1.Filter = "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    
    strFileName = CommonDialog1.FileName
    strCSV = ""
    If Trim(strFileName) <> "" Then
        For i = 0 To vasSum.DataRowCnt
            For j = 1 To vasSum.MaxCols
                strCSV = strCSV & Trim(GetText(vasSum, i, j)) & ","
            Next j
            strCSV = strCSV & vbCrLf
            
        Next i
        
        FilNum = FreeFile
        Open strFileName For Output As FilNum
        
        Print #FilNum, strCSV
        Close FilNum
    
    End If
    
    
    
    
End Sub

Private Sub cmdEquipConnect_Click()
    Advia_Init
End Sub

Private Sub cmdListClear_Click()
    Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasList, 1, 1
    vasList.MaxRows = 0
    ClearSpread vasListRes, 1, 1
    vasListRes.MaxRows = 0
End Sub

Private Sub cmdListTrans_Click()
'선택전송
Dim VasidRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If txtUID.Text = "" Then
        MsgBox "사용자 확인을 해 주십시오"
        txtUID.SetFocus
        Exit Sub
    End If
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
    Connect_Server
    For VasidRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = VasidRow
        
        If vasList.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
'''            liRet = Insert_Data(VasidRow)
            If Barcode_Gubun(Trim(GetText(vasList, VasidRow, colBarCode))) = "Q" Then
                liRet = Insert_QC_Data(vasList, VasidRow)
            Else
                liRet = Insert_Data_1(VasidRow)
            End If
            
            
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasList, VasidRow, VasidRow, colBarCode, colState, 255, 255, 180
                SetText vasList, "Trans", VasidRow, colState
            Else
                SetBackColor vasList, VasidRow, VasidRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasList, "Failed", VasidRow, colState
            End If
            vasList.Col = 1
            vasList.Row = VasidRow
            vasList.Value = 0
        Else
        
        End If
    Next VasidRow
    
End Sub


Private Sub cmdSugaClear_Click()
    ClearSpread vasSum
    vasSum.MaxRows = 0
End Sub

Private Sub cmdSumSch_Click()
    Dim i As Long
    Dim j As Long
    Dim x As Long
    
    Dim iSumRow As Integer
    Dim iSum As Long
    
    
    ClearSpread vasSum
    
    SQL = "select distinct examdate from pat_res " & vbCrLf & _
          "where examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "'"
    res = db_select_Vas(gLocal, SQL, vasSum)
    
    iSumRow = vasSum.DataRowCnt + 1
    vasSum.MaxRows = iSumRow
    
    SetText vasSum, "합계", iSumRow, 1
    
    For i = 1 To vasSum.DataRowCnt
        For j = 2 To vasSum.MaxCols
            SetText vasSum, "0", i, j
            
        Next
    Next
    
    
    ClearSpread vasSumTemp
    SQL = "select examdate, barcode, equipcode, resvalue from pat_res " & vbCrLf & _
          "where examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "'" & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "group by examdate, equipcode, resvalue, barcode"
    res = db_select_Vas(gLocal, SQL, vasSumTemp)
    

'''    gArr_Exam(i, 1) = i    '순서
'''    gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
'''    gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
'''    gArr_Exam(i, 4) = Trim(GetText(vasTemp, i, 4))    '결과구분
            
    For i = 1 To vasSumTemp.DataRowCnt
        For j = 1 To iSumRow - 1
            If Trim(GetText(vasSumTemp, i, 1)) = Trim(GetText(vasSum, j, 1)) Then
                For x = 2 To vasSum.MaxCols
                    If Trim(GetText(vasSumTemp, i, 3)) = Trim(gArr_Exam(x - 1, 2)) Then
                        SetText vasSum, CCur(Trim(GetText(vasSum, j, x))) + 1, j, x
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Next
    
    For i = 2 To vasSum.MaxCols
        iSum = 0
        For j = 1 To iSumRow - 1
            iSum = iSum + CCur(Trim(GetText(vasSum, j, i)))
        Next
        SetText vasSum, iSum, iSumRow, i
        
    Next
    
    
    ClearSpread vasSumTemp
    
End Sub

Private Sub cmdVasIDWidth_Click()
    Dim i As Integer
    
    
    If cmdVasIDWidth.Caption = ">>" Then
        vasID.Width = 14385
        cmdVasIDWidth.Caption = "<<"
        
        vasID.Visible = False
        For i = colRStart + 1 To vasID.MaxCols
            vasID.Col = i
            vasID.ColHidden = False
        Next
        vasID.Visible = True
        vasID.ScrollBars = ScrollBarsBoth
    Else
        vasID.Width = 6375
        cmdVasIDWidth.Caption = ">>"
        vasID.Visible = False
        For i = colRStart + 1 To vasID.MaxCols
            vasID.Col = i
            vasID.ColHidden = True
        Next
        vasID.Visible = True
        vasID.ScrollBars = ScrollBarsVertical
    End If
End Sub

Private Sub cmdVasListWidth_Click()
    Dim i As Integer
    
    If cmdVasListWidth.Caption = ">>" Then
        vasList.Width = 14385
        cmdVasListWidth.Caption = "<<"
        
        vasList.Visible = False
        For i = colRStart + 1 To vasList.MaxCols
            vasList.Col = i
            vasList.ColHidden = False
        Next
        vasList.Visible = True
        vasList.ScrollBars = ScrollBarsBoth
    Else
        vasList.Width = 6375
        cmdVasListWidth.Caption = ">>"
        vasList.Visible = False
        For i = colRStart + 1 To vasList.MaxCols
            vasList.Col = i
            vasList.ColHidden = True
        Next
        vasList.Visible = True
        vasList.ScrollBars = ScrollBarsVertical
    End If
End Sub

Private Sub Command1_Click()
    Dim S As String
    Dim i As Long
    
    
    For i = 1 To Len(Text1.Text)
    
    
        S = Mid(Text1, i, 1)
        
        Timer1.Enabled = False
        
        If gErrState = 1 Then
            Advia_Init
            Exit Sub
        End If
                
                If gComState = 0 Then               'Initialize 신호 던진 후
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                Advia_Token                 'Token Message 던지기
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 1 Then           'Token Message 던진 후 ACK
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
                'Advia_Token
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 3 Then           'Order 혹은 No Order 를 보낸 후
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 4 Then           'Result Validation Message 보내고 난 뒤
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        Else
            Select Case S
            Case chrNACK
                gErrState = 1
    '''            Advia_Token_1
                Advia_Init
                Exit Sub
            Case chrSTX
                txtBuff.Text = txtBuff.Text & chrSTX
            Case chrETX
                txtBuff.Text = txtBuff.Text & S
    
                Save_Raw_Data "[Rx]" & txtBuff.Text
    
                If Mid(txtBuff.Text, 2, 1) < "0" Or Mid(txtBuff.Text, 2, 1) > "Z" Then
                    gErrState = 1
                    Advia_Token_1
    '                Advia_Init
                    Exit Sub
    
                Else
                    gMT = Mid(txtBuff.Text, 2, 1)
                    MSComm1.Output = gMT
                    Timer1.Enabled = True
    
                    'save_raw_Data "[Tx]" & gMT
    '''                If gMT = "Q" Or gMT = "R" Then      'Query와 Result인 경우만 로그 남기기
    '''                    Save_Raw_Data "[Rx]" & txtBuff.Text
    '''                End If
                End If
    
                Advia Mid(txtBuff.Text, 2)
                txtBuff.Text = ""
            Case Else
                txtBuff.Text = txtBuff.Text & S
            End Select
        End If

    
    Next
    Text1.Text = ""
End Sub

Sub Var_Clear()
    gOrderMessage = ""
    
    gBarCode = ""
'''    sBarCode = ""
'''    sSeqNo = ""
'''    sDiskno = ""
'''    sPosno = ""
    sSampleType = ""
'''    txtpat = ""
    llRow = -1
End Sub

Private Function Result_Set(asExamCode As String, asResult As String) As String
    Dim strRefH As String
    Dim strRefM As String
    Dim strRefL As String
    Dim cRefH As String
    Dim cRefL As String
    Dim strResGubun As String
    Dim strLEquil As String
    Dim strHEquil As String
    Dim i As Integer
    Dim strRespRec As String
    Dim strPointFormat As String
    Dim cRepH As String
    Dim cRepL As String
    Dim strGiho As String
    Dim strResult As String
    Dim strResValue As String
    
    On Error GoTo ErrRes:
    
    Result_Set = ""
    
    strResValue = asResult
    
    If IsNumeric(strResValue) = False Then
        Result_Set = strResValue & "/" & strResValue
        Exit Function
    End If
    
    SQL = "SELECT REPLOW, REPHIGH, REFLOW, REFHIGH, LSTRING, MSTRING, HSTRING, LEQUIL, HEQUIL, RESPREC, RESGUBUN " & vbCrLf & _
          "FROM EQUIPEXAM WHERE EQUIPNO = '" & gEquip & "' AND EXAMCODE = '" & asExamCode & "'"
    res = db_select_Col(gLocal, SQL)
    
    cRepL = Trim(gReadBuf(0))
    cRepH = Trim(gReadBuf(1))
    cRefL = Trim(gReadBuf(2))
    cRefH = Trim(gReadBuf(3))
    strRefL = Trim(gReadBuf(4))
    strRefM = Trim(gReadBuf(5))
    strRefH = Trim(gReadBuf(6))
    strLEquil = Trim(gReadBuf(7))
    strHEquil = Trim(gReadBuf(8))
    strRespRec = Trim(gReadBuf(9))
    strResGubun = Trim(gReadBuf(10))
    
    If IsNumeric(cRepL) = True Then
        If CCur(cRepL) > CCur(strResValue) Then
            strGiho = "<"
            strResValue = cRepL
        End If
    End If
    
    If IsNumeric(cRepH) = True Then
        If CCur(cRepH) < CCur(strResValue) Then
            strGiho = ">"
            strResValue = cRepH
        End If
    End If
    
    If strResGubun = "1" Then '문자
        If IsNumeric(cRefL) = True Then
            If strLEquil = "1" Then
                If CCur(cRefL) >= CCur(strResValue) Then
                    strResult = strRefL
                End If
            Else
                If CCur(cRefL) > CCur(strResValue) Then
                    strResult = strRefL
                End If
            End If
        End If
        
        If IsNumeric(cRefH) = True Then
            If strHEquil = "1" Then
                If CCur(cRefH) <= CCur(strResValue) Then
                    strResult = strRefH
                End If
            Else
                If CCur(cRefH) < CCur(strResValue) Then
                    strResult = strRefH
                End If
            End If
        End If
        
        If IsNumeric(cRefL) = True And IsNumeric(cRefH) = True Then
            If strLEquil = "1" And strHEquil = "1" Then
                If CCur(cRefL) <= CCur(strResValue) And CCur(cRefL) >= CCur(strResValue) Then
                    strResult = strRefM
                End If
            ElseIf strLEquil = "1" And strHEquil = "0" Then
                If CCur(cRefL) <= CCur(strResValue) And CCur(cRefL) > CCur(strResValue) Then
                    strResult = strRefM
                End If
            ElseIf strLEquil = "0" And strHEquil = "1" Then
                If CCur(cRefL) < CCur(strResValue) And CCur(cRefL) >= CCur(strResValue) Then
                    strResult = strRefM
                End If
            Else
                If CCur(cRefL) < CCur(strResValue) And CCur(cRefL) > CCur(strResValue) Then
                    strResult = strRefM
                End If
                
            End If
        End If
        
    End If
    
    
    If IsNumeric(strRespRec) = True Then
        strPointFormat = ""
        For i = 1 To CInt(strRespRec)
            
            If i = 1 Then
                strPointFormat = ".0"
            Else
                strPointFormat = strPointFormat & "0"
            End If
        Next
        
        strPointFormat = "##0" & strPointFormat
        
        strResValue = Format(strResValue, strPointFormat)
        
    Else
        strResValue = strResValue
    End If
    
    Result_Set = strGiho & strResValue & "/" & strResult
    Exit Function
    
ErrRes:
    
    Result_Set = strResValue & "/" & strResValue
    Exit Function
    
End Function

Private Sub Init_Form()
    frmInterface.Caption = gEquipName & " Interface Program"
    SSPanel1.Caption = "     " & gEquipName & "  INTERFACE"
End Sub

Private Sub Command9_Click()

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    
    '1. 화면 및 변수 초기화
    '2. 데이타베이스에 Connect 하기 - Local - Server
    '3. Ini 내용 불러오기    GetSetup
    '4. Comport Open
    
    'Timer interval = 3000 -> 10000
    
    Me.Left = 0
    Me.Top = 0
    
    
        
    GetSetup    'ini에서 DB정보 불러오기
    
    Init_Form
    
    If Not Connect_Server Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If

    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = "True"
    MSComm1.DTREnable = "True"
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
   
    lblUser.Caption = gExamUID
    txtUID.Text = gExamUID

    raw_data = ""

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    cmdClear_Click
    
    dtpToday = GetDateFull
    dtpExamDate = GetDateFull
    dtpSumSDate = Format(GetDateFull, "yyyy/mm")
    dtpSumEDate = GetDateFull
    
    
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(dtpToday), -gLocalExpDate), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    res = SendQuery(gLocal, SQL)
    '===================================================================
    
    '검사코드 가져오기
    GetExamCode

    ClearSpread vasCode

    vasID.MaxRows = 1
    vasID.ColsFrozen = 6
    vasRes.MaxRows = 20
    vasList.MaxRows = 1
    
    vasList.ColsFrozen = 6
    
    vasListRes.MaxRows = 20
    
    vasSum.MaxRows = 20
    vasSum.ColsFrozen = 1
    
'''    vasID.Visible = False
    For i = colRStart + 1 To vasID.MaxCols
        vasID.Col = i
        vasID.ColHidden = True
    Next
'''    vasID.Visible = True
    
'''    vasList.Visible = False
    For i = colRStart + 1 To vasList.MaxCols
        vasList.Col = i
        vasList.ColHidden = True
    Next
'''    vasID.Visible = True

    SSTab1.Tab = 0
    Advia_Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    WritePrivateProfileString "config", "UID", txtUID.Text, App.Path & "\interface.ini"
    
    DisConnect_Server
    DisConnect_Local
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    
    gAllExam = ""
    gOrderExam = ""
    gReceExam = ""
    
    
    For i = 1 To 500
        gArr_Exam(i, 1) = ""
        gArr_Exam(i, 2) = ""
        gArr_Exam(i, 3) = ""
    Next i
    
    ClearSpread vasTemp
    
    SQL = "Select SeqNo, EquipCode, ExamName, resgubun From EquipExam where Equipno = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "GROUP BY SeqNo, EquipCode, ExamName, resgubun "
    SQL = SQL & vbCrLf & " Order by SeqNo"
    res = db_select_Vas(gLocal, SQL, vasTemp)
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    vasID.MaxCols = colRStart + vasTemp.DataRowCnt
    vasList.MaxCols = colRStart + vasTemp.DataRowCnt
    vasSum.MaxCols = vasTemp.DataRowCnt + 1
    
    For i = 1 To vasTemp.DataRowCnt
        If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
            gArr_Exam(i, 1) = i    '순서
            gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
            gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
            gArr_Exam(i, 4) = Trim(GetText(vasTemp, i, 4))    '결과구분
            
            SetText vasID, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
            SetText vasList, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
            SetText vasSum, Trim(GetText(vasTemp, i, 3)), 0, i + 1
            
        End If
        
    Next i
    
'''    For i = 1 To 100
'''        gArr_Exam(i, 1) = ""
'''        gArr_Exam(i, 2) = ""
'''        gArr_Exam(i, 3) = ""
'''    Next i
    

    
    
    ClearSpread vasTemp
    
    SQL = "Select ExamCode From EquipExam where Equipno = '" & gEquip & "' "
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For i = 1 To vasTemp.DataRowCnt

        If Trim(GetText(vasTemp, i, 1)) <> "" Then
            If gAllExam = "" Then
                gAllExam = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
            Else
                gAllExam = gAllExam & ",'" & Trim(GetText(vasTemp, i, 1)) & "'"
            End If
        End If
    Next i
    
End Sub


Private Sub mnuAuto_Click()
    mnuManual.Checked = False
    mnuAuto.Checked = True
End Sub

Private Sub mnuCodeConfig_Click()
    frmEquipExam.SSPanel1.Caption = "  " & gEquipName & " 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub mnuConfig_Click()
    frmConfig.SSPanel_machine.Caption = gEquipName
    frmConfig.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuManual_Click()
    mnuManual.Checked = True
    mnuAuto.Checked = False
End Sub

Private Sub MSComm1_OnComm()
    
    Dim S As String
    Dim sAll As String
    Dim i As Integer
    
    

    sAll = MSComm1.Input
    Save_Raw_Data "[AA]" & sAll
    
    If Trim(sAll) = chrNACK Then
        Advia_Init
        Exit Sub
    End If
    
    
    For i = 1 To Len(sAll)
    
        S = Mid(sAll, i, 1)
        Timer1.Enabled = False
        
        If gErrState = 1 Then
            Advia_Init
            Exit Sub
        End If

        If gComState = 0 Then               'Initialize 신호 던진 후
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                lblConnect.Caption = "연결!"
                Advia_Token                 'Token Message 던지기
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 1 Then           'Token Message 던진 후 ACK
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
                'Advia_Token
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 3 Then           'Order 혹은 No Order 를 보낸 후
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        ElseIf gComState = 4 Then           'Result Validation Message 보내고 난 뒤
            'save_raw_Data "[Rx]" & S
            If S = gMT Then
                gComState = 2
            ElseIf S = chrSTX Then
                txtBuff.Text = chrSTX
                gComState = 2
            End If
        Else
            Select Case S
            Case chrNACK
                gErrState = 1
    '''            Advia_Token_1
                Advia_Init
                Exit Sub
            Case chrSTX
                txtBuff.Text = txtBuff.Text & chrSTX
            Case chrETX
                txtBuff.Text = txtBuff.Text & S
    
                Save_Raw_Data "[Rx]" & txtBuff.Text
    
                If Mid(txtBuff.Text, 2, 1) < "0" Or Mid(txtBuff.Text, 2, 1) > "Z" Then
                    gErrState = 1
                    Advia_Token_1
    '                Advia_Init
                    Exit Sub
    
                Else
                    gMT = Mid(txtBuff.Text, 2, 1)
                    MSComm1.Output = gMT
                    Timer1.Enabled = True
    
                    'save_raw_Data "[Tx]" & gMT
    '''                If gMT = "Q" Or gMT = "R" Then      'Query와 Result인 경우만 로그 남기기
    '''                    Save_Raw_Data "[Rx]" & txtBuff.Text
    '''                End If
                End If
    
                Advia Mid(txtBuff.Text, 2)
                txtBuff.Text = ""
            Case Else
                txtBuff.Text = txtBuff.Text & S
            End Select
        End If
    Next
    
    
'''    Timer1.Enabled = False
'''
''''''    If S = gMT Then
'''''''''        Save_Raw_Data "[Rx]" & S
'''''''''
'''''''''        Advia_Token
''''''
''''''    Else
''''''        txtBuff.Text = S
''''''        Save_Raw_Data "[Rx]" & S
''''''        If Mid(S, 2, 1) < "0" Or Mid(S, 2, 1) > "Z" Then
''''''            gErrState = 1
''''''            Advia_Token_1
'''''''                Advia_Init
''''''            Exit Sub
''''''        Else
'''''''''            If gMT = "Q" Or gMT = "R" Then      'Query와 Result인 경우만 로그 남기기
'''''''''            Save_Raw_Data "[Rx]" & S
'''''''''            End If
''''''
''''''            gMT = Mid(S, 2, 1)
''''''            MSComm1.Output = gMT
''''''
''''''            Save_Raw_Data "[Tx]" & gMT
''''''
''''''        End If
''''''
''''''        Advia Mid(S, 2)
''''''    End If
'''
'''
'''
'''    If gErrState = 1 Then
'''        Advia_Init
'''        Exit Sub
'''    End If
'''
'''    If gComState = 0 Then               'Initialize 신호 던진 후
'''        'save_raw_Data "[Rx]" & S
'''        If S = gMT Then
'''            Advia_Token                 'Token Message 던지기
'''        ElseIf S = chrSTX Then
'''            txtBuff.Text = chrSTX
'''            gComState = 2
'''        End If
'''    ElseIf gComState = 1 Then           'Token Message 던진 후 ACK
'''        'save_raw_Data "[Rx]" & S
'''        If S = gMT Then
'''            gComState = 2
'''            'Advia_Token
'''        ElseIf S = chrSTX Then
'''            txtBuff.Text = chrSTX
'''            gComState = 2
'''        End If
'''    ElseIf gComState = 3 Then           'Order 혹은 No Order 를 보낸 후
'''        'save_raw_Data "[Rx]" & S
'''        If S = gMT Then
'''            gComState = 2
'''        ElseIf S = chrSTX Then
'''            txtBuff.Text = chrSTX
'''            gComState = 2
'''        End If
'''    ElseIf gComState = 4 Then           'Result Validation Message 보내고 난 뒤
'''        'save_raw_Data "[Rx]" & S
'''        If S = gMT Then
'''            gComState = 2
'''        ElseIf S = chrSTX Then
'''            txtBuff.Text = chrSTX
'''            gComState = 2
'''        End If
'''    Else
'''        Select Case S
'''        Case chrSTX
'''            txtBuff.Text = chrSTX
'''        Case chrETX
'''            txtBuff.Text = txtBuff.Text & S
'''
'''            Save_Raw_Data "[Rx]" & txtBuff.Text
'''
'''            If Mid(txtBuff.Text, 2, 1) < "0" Or Mid(txtBuff.Text, 2, 1) > "Z" Then
'''                gErrState = 1
'''                Advia_Token_1
''''                Advia_Init
'''                Exit Sub
'''            ElseIf Mid(txtBuff.Text, 2, 1) = chrETX Then
'''                gErrState = 1
'''                Advia_Token_1
''''                Advia_Init
'''                Exit Sub
'''            Else
'''                gMT = Mid(txtBuff.Text, 2, 1)
'''                MSComm1.Output = gMT
'''                Timer1.Enabled = True
'''
'''                'save_raw_Data "[Tx]" & gMT
''''''                If gMT = "Q" Or gMT = "R" Then      'Query와 Result인 경우만 로그 남기기
''''''                    Save_Raw_Data "[Rx]" & txtBuff.Text
''''''                End If
'''            End If
'''
'''            Advia Mid(txtBuff.Text, 2)
'''
'''        Case Else
'''            txtBuff.Text = txtBuff.Text & S
'''        End Select
'''    End If
    
'''    Next
    
End Sub

Sub Advia(asVar As String)
    Dim lsData As String
    Dim lsIDCode As String
    Dim lsID As String
    Dim lsTube As String
    Dim lsSeq As String
    Dim lsTest As String
    Dim lsRes As String
    Dim lsFlag As String
    
    Dim i, j As Integer
    Dim iRow As Integer
    
    Dim MyVar As String
    Dim MyRet As String
    Dim MyRetTmp As String
    
    Dim sDisk As String
    Dim sPosition As String
    Dim sBarcode As String
    
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsAlarm As String
    Dim lsExamDate As String
    Dim lsExamTime As String
'''    Dim lsSeq As String
    
    Dim sExamCode1 As String
    Dim RCnt As String
    
    
    lsExamDate = Format(Date, "yyyymmdd")
    lsExamTime = Format(Date, "hhmmss")
    
    lsData = asVar
    lsIDCode = Mid(lsData, 2, 1)
    
    Select Case lsIDCode
    Case "I"    'Initialization
    
    Case "S"    'Token Transfer
        lblMT.Caption = Mid(lsData, 1, 1)
        DoSleep 1
        Timer1.Enabled = False
        
        Advia_Token
    Case "Y"    'Workorder
    
    Case "E"    'Workorder Validation
    
    Case "R"    'Result
        sDisk = ""
        sPosition = ""
        sBarcode = ""
        
        sDisk = Mid(lsData, 19, 3)
        sPosition = Mid(lsData, 23, 2)
        
        sBarcode = Right(Mid(lsData, 4, 14), 12)
        gSpecID = sBarcode
        
        If Trim(sBarcode) = "" Then
            Exit Sub
        End If
    
        lsData = Mid(lsData, 53)
        i = InStr(1, lsData, chrLF)
        If i < 1 Then Exit Sub
        
        MyRet = Mid(lsData, i + 1)
        
        glRow = -1
        For iRow = vasID.DataRowCnt To 1 Step -1
            If Trim(GetText(vasID, iRow, colBarCode)) = gSpecID Then
                glRow = iRow
                Exit For
            End If
        Next iRow
        
        If glRow = -1 Then
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow
            End If
        End If
        
        vasActiveCell vasID, glRow, colBarCode
            
        ClearSpread vasRes, 1, 1
        
'''        SetText vasID, sDisk & "-" & sPosition, glRow, colSeqNo
        SetText vasID, gSpecID, glRow, colBarCode

        '환자정보 불러오기
        If Trim(GetText(vasID, glRow, colPID)) = "" Then
            If Barcode_Gubun(gSpecID) = "Q" Then
                Get_QC_Info vasID, glRow
                
            Else
                Get_Sample_Info glRow
            End If
        End If
        
        '검사코드만큼 Row의 갯수를 설정
        SQL = "Select count(ExamCode) From EquipExam" & vbCrLf & _
                  " Where Equipno = '" & gEquip & "' "
        res = db_select_Col(gLocal, SQL)
        vasRes.MaxRows = Trim(gReadBuf(0))
        
        RCnt = 0
        
        j = 0
        
        '결과
        MyRetTmp = MyRet

        Do While Trim(MyRetTmp) <> ""
            lsAlarm = ""
            lsResult = ""
            lsEquipCode = ""
            
            If Len(MyRetTmp) < 8 Then
                Exit Do
            End If
            
            lsEquipCode = Trim(Left(MyRetTmp, 3))
            lsEquipCode = Format(lsEquipCode, "00#")
            
            lsResult = Trim(Mid(MyRetTmp, 4, 5))
            lsAlarm = Trim(Mid(MyRetTmp, 9, 1))
                    
            gReadBuf(0) = "0"
            
            sExamCode1 = ""
            lsExamCode = ""
            
            SQL = "select examcode from pat_res " & vbCrLf & _
                  " where Equipno = '" & gEquip & "' " & vbCrLf & _
                  "   And barcode = '" & gSpecID & "' and equipcode = '" & lsEquipCode & "'"
            res = db_select_Col(gLocal, SQL)
            
            lsExamCode = Trim(gReadBuf(0))
            
            If lsExamCode = "" Then
            
                SQL = "Select examcode From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And EquipCode = '" & lsEquipCode & "'"
                res = db_select_Row(gLocal, SQL)
                
                For i = 1 To res
                    If sExamCode1 = "" Then
                        sExamCode1 = "'" & Trim(gReadBuf(i - 1)) & "'"
                    Else
                        sExamCode1 = sExamCode1 & ", '" & Trim(gReadBuf(i - 1)) & "'"
                    End If
                Next i
                
                If sExamCode1 = "" Then sExamCode1 = "''"
                
                
                If Barcode_Gubun(lsID) = "Q" Then
                    SQL = Select_QC_Exam(lsID, sExamCode1)
                    
                Else
                    SQL = "SELECT AA.EXMN_CD " & vbCrLf & _
                          "FROM MS.MSLRCPT A " & vbCrLf & _
                          "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
                          "WHERE A.SPCM_NO = '" & Trim(gSpecID) & "' " & vbCrLf & _
                          "AND AA.RSLT_PRGR_STAT_CD in ('05', '07', '09', '12') " & vbCrLf & _
                          "AND AA.EXMN_CD IN (" & sExamCode1 & ")"
                End If
                
                res = db_select_Col(gServer, SQL)
                
                lsExamCode = ""
                lsExamCode = gReadBuf(0)
            End If
            
            If lsExamCode = "" Then
                SQL = "Select equipcode, examcode, examname, seqno From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And EquipCode = '" & lsEquipCode & "'"
                res = db_select_Col(gLocal, SQL)
            Else
                SQL = "Select equipcode, examcode, examname, seqno From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And ExamCode = '" & lsExamCode & "'"
                res = db_select_Col(gLocal, SQL)
            End If
            
            
            
            If (res = 1) And (gReadBuf(0) <> "") Then
                lsExamCode = Trim(gReadBuf(1))
                lsExamName = Trim(gReadBuf(2))
                lsSeq = Trim(gReadBuf(3))
                
                RCnt = RCnt + 1
                j = j + 1
                
                lsResult = Result_Set(lsExamCode, lsResult)
                i = InStr(1, lsResult, "/")
                If i > 0 Then
                    lsResult = Mid(lsResult, 1, i - 1)
                Else
                    lsResult = lsResult
                End If
                
                SetText vasRes, lsEquipCode, j, colEquipExam
                SetText vasRes, lsExamCode, j, colExamCode
                SetText vasRes, lsExamName, j, colExamName
                SetText vasRes, lsSeq, j, colSeq
                SetText vasRes, lsResult, j, colResult
                SetText vasRes, lsResult, j, colResValue
                SetText vasRes, lsExamDate, j, colResDate
                SetText vasRes, lsExamTime, j, colResTime
                SetPositionResult glRow, lsEquipCode, lsResult
                
'''                '결과에 대한 판정
'''                Check_Result gSpecID, Trim(GetText(vasID, glRow, colPID)), lsExamCode, CStr(lsResult), j, Trim(GetText(vasID, glRow, colPSex))
                
                '결과 Local에 저장
                Save_Local_One glRow, j, "1"
            End If
            
            MyRetTmp = Mid(MyRetTmp, 10)
        Loop

        '결과갯수
        SetText vasID, CStr(RCnt), glRow, colRcnt
        SetText vasID, "Result", glRow, colState
        SetForeColor vasID, glRow, glRow, 0, 0, 0
        
        If mnuAuto.Checked = True Then
            res = -1
            If Barcode_Gubun(Trim(GetText(vasID, glRow, colBarCode))) = "Q" Then
                res = Insert_QC_Data(vasID, glRow)
            Else
                res = Insert_Data(CInt(glRow))
            End If
            If res = 1 Then
                'db_Commit gServer
                
                SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", glRow, colState
            Else
                SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", glRow, colState
            End If
            
        End If
        
        
        Timer1.Enabled = False
        Advia_ResValid
        
    Case "Q"    'Query
        sBarcode = Right(Mid(lsData, 4, 14), 12)
        gOrderMessage = ""
        
        'Order 생성
        res = Proc_Order_Advia(sBarcode)
        If res < 1 Then
            MSComm1.Output = Advia_NoOrder(sBarcode)
        Else
            MSComm1.Output = gOrderMessage
            Timer1.Enabled = True
            Save_Raw_Data "[TX]" & gOrderMessage
            gComState = 3
        End If
                
    Case "N"    'No Order
    
    Case "Z"    'Result Vaidation
    
'    Case Else
'        MSComm1.Output = chrNACK
'        save_raw_Data "[TX]" & chrNACK
    End Select
End Sub

Function Proc_Order_Advia(asID As String) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iRow As Integer
    
    Dim sOCnt As String
    
    Dim lsID As String
    Dim lsExamCode As String
    Dim sExamCode As String

    Dim retInfo As String
    Dim retOrder As String
    Dim lsEquipCode As String

    Dim sExamFlag As String
    Dim sPSex As String
    Dim sDate As String
    
On Error GoTo ErrHandle

    retOrder = ""
    retInfo = ""
    
    gOrderMessage = ""
    
    Proc_Order_Advia = -1

    lsID = Trim(asID)
    
    If Trim(lsID) = "" Then
        gOrderMessage = Advia_NoOrder(lsID)
        Exit Function
    End If
    
    glRow = -1
    For iRow = vasID.DataRowCnt To 1 Step -1
        If Trim(GetText(vasID, iRow, colBarCode)) = lsID Then
            glRow = iRow
            Exit For
        End If
    Next iRow
    
    If glRow = -1 Then
        glRow = vasID.DataRowCnt + 1
        If glRow > vasID.MaxRows Then
            vasID.MaxRows = glRow
        End If
    End If
    
    SetText vasID, lsID, glRow, colBarCode
    
    vasActiveCell vasID, glRow, colPID
    
    If Trim(GetText(vasID, glRow, colPID)) = "" Then
        If Barcode_Gubun(lsID) = "Q" Then
            Get_QC_Info vasID, glRow
            
        Else
            Get_Sample_Info glRow
        End If
    End If
    
    ClearSpread vasRes, 1, 1

'''    SetText vasID, "", glRow, colSeq

    SetForeColor vasID, glRow, glRow, 0, 0, 0
    
    retOrder = ""
    lsExamCode = ""
    
    sExamFlag = ""
    
    If Barcode_Gubun(lsID) = "Q" Then
        SQL = Select_QC_Exam(lsID)
        
    Else
        SQL = "SELECT A.SPCM_NO, B.SEX_CD " & vbCrLf & _
              "FROM MS.MSLRCPT A " & vbCrLf & _
              "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
              "INNER JOIN HO.PCPPATIENT B ON A.PID = B.PID " & vbCrLf & _
              "WHERE A.SPCM_NO = '" & Trim(lsID) & "' " & vbCrLf & _
              "AND AA.RSLT_PRGR_STAT_CD in ('05', '07', '09', '12') " & vbCrLf & _
              "AND AA.EXMN_CD IN (" & gAllExam & ")"
    End If
    res = db_select_Col(gServer, SQL)
    sPSex = Trim(gReadBuf(1))
    If sPSex = "" Then
        sPSex = "M"
    End If
    
    If Trim(lsID) = Trim(gReadBuf(0)) Then
        sExamFlag = "A"
    Else
        sExamFlag = " "
    End If
    
    sDate = SeperatorCls(Format(Date, "yyyy/mm/dd"))
    
    '검사코드 가져오기
    ClearSpread vasCode
    
    If Barcode_Gubun(lsID) = "Q" Then
        SQL = Select_QC_Exam(lsID)
        
    Else
        SQL = "SELECT distinct AA.EXMN_CD " & vbCrLf & _
              "FROM MS.MSLRCPT A " & vbCrLf & _
              "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
              "WHERE A.SPCM_NO = '" & Trim(lsID) & "' " & vbCrLf & _
              "AND AA.RSLT_PRGR_STAT_CD in ('05', '07', '09', '12') " & vbCrLf & _
              "AND AA.EXMN_CD IN (" & gAllExam & ")"
    End If
    
    res = db_select_Vas(gServer, SQL, vasCode)
    
    If res = 0 Then
        SetText vasID, "0", glRow, colRcnt              'Order 없음
        SetForeColor vasID, glRow, glRow, 255, 0, 0
        
        Proc_Order_Advia = 0
        retOrder = ""

        gOrderMessage = Advia_NoOrder(lsID)

        Exit Function

    Else
        lsExamCode = ""
        For i = 0 To vasCode.DataRowCnt
            If lsExamCode = "" Then
                lsExamCode = "'" & Trim(GetText(vasCode, i, 1)) & "'"
            Else
                lsExamCode = lsExamCode & ",'" & Trim(GetText(vasCode, i, 1)) & "'"
            End If
        Next i
        
        Save_Raw_Data lsID & " : " & lsExamCode
        
        retInfo = "   " & sExamFlag & " " & Advia_IDSet(lsID) & "                         " & _
                  SetSpace(Trim(GetText(vasID, glRow, colPID)), 14, 1) & "   " & Space(30) & " " & _
                  "          " & " " & sPSex & " " & _
                  Mid(sDate, 5, 2) & "/" & Mid(sDate, 7, 2) & "/" & Mid(sDate, 3, 2) & " " & "    " & " " & _
                  Space(6) & " " & Space(6) & " " & chrCR & chrLF
                  
                  
    End If
    
    'Order
    sOCnt = 1
    
    If lsExamCode <> "" Then
        For i = 1 To vasCode.DataRowCnt
            sExamCode = Trim(GetText(vasCode, i, 1))
    
            '검사코드로 장비코드 불러오기
            lsEquipCode = ""
            lsEquipCode = GetEquip_ExamCode(sExamCode)
            
            SetPositionResult glRow, lsEquipCode, "*"

            If lsEquipCode <> "" Then
                retOrder = retOrder & Format(CLng(lsEquipCode), "000")
            
                sOCnt = sOCnt + 1
                
                SQL = "select barcode from pat_res where barcode = '" & Trim(lsID) & "' and examcode = '" & Trim(sExamCode) & "'"
                res = db_select_Col(gLocal, SQL)
                If res = 0 Then
                    SQL = "select examname, seqno from equipexam " & vbCrLf & _
                          "where equipno = '" & gEquip & "' and examcode = '" & sExamCode & "' "
                    res = db_select_Col(gLocal, SQL)
                    
                    SQL = "insert into pat_res(equipno, examdate, barcode, examcode, equipcode, result, pname, pid, " & vbCrLf & _
                          "                    seqno, page, examname) " & vbCrLf & _
                          " values('" & gEquip & "', '" & Format(Date, "yyyymmdd") & "', " & vbCrLf & _
                          " '" & Trim(lsID) & "','" & Trim(sExamCode) & "','" & Trim(lsEquipCode) & "', '', " & vbCrLf & _
                          "'" & Trim(GetText(vasID, glRow, colPName)) & "', '" & Trim(GetText(vasID, glRow, colPID)) & "', " & vbCrLf & _
                          "'" & Trim(gReadBuf(1)) & "', 0, '" & Trim(gReadBuf(0)) & "')"
                    res = SendQuery(gLocal, SQL)
                End If
                
            End If
        Next i
    Else
         Proc_Order_Advia = 0
    End If
'=======================================================================

    If lsExamCode = "" Then
        gOrderMessage = Advia_NoOrder(lsID)
    Else
        Proc_Order_Advia = 1
        gMT = Chr(Asc(gMT) + 1)
        If gMT > "Z" Then gMT = "0"
        retOrder = gMT & "Y" & retInfo & retOrder & chrCR & chrLF
        gOrderMessage = chrSTX & retOrder & LRC(retOrder) & chrETX
    End If

    SetText vasID, sOCnt, glRow, colRcnt
    If sOCnt = 0 Then
        SetText vasID, "없음", glRow, colState
        SetForeColor vasID, glRow, glRow, 255, 0, 0
    Else
        SetText vasID, CStr(sOCnt - 1), glRow, colRcnt
        SetText vasID, "Order", glRow, colState
        SetForeColor vasID, glRow, glRow, 0, 0, 0
    End If

    vasActiveCell vasID, glRow, 1

    Exit Function

ErrHandle:
    Proc_Order_Advia = -1
    SaveQuery SQL
    Resume Next
End Function


Private Sub SetPositionResult(asRow As Long, asEquipCode As String, asResult As String)
    Dim strEquipCode As String
    Dim strResult As String
    Dim lngRow As Long
    Dim i As Integer
    
    lngRow = asRow
    strEquipCode = asEquipCode
    strResult = asResult

    For i = colRStart + 1 To vasID.MaxCols
        If Trim(gArr_Exam(i - colRStart, 2)) = Trim(strEquipCode) Then
            SetText vasID, strResult, lngRow, i
            Exit For
        End If
    Next
End Sub

Public Function GetExamCode_Equip(argCode As String, argReceNo As String, argDate As String) As Integer
'검체번호에 존재하는 장비번호 해당하는 검사코드 가져오기

    Dim i As Integer
    Dim sExamCode As String
     
    sExamCode = ""
    GetExamCode_Equip = -1
    ClearSpread frmInterface.vaSpread1
    
    If argCode = "" Then
        Exit Function
    End If
    
    sExamCode = ""
    SQL = "Select ExamCode From EquipExam" & vbCrLf & _
          "Where Equip = '" & gEquip & "'" & vbCrLf & _
          "  And EquipCode = '" & argCode & "' "
    res = db_select_Vas(gServer, SQL, frmInterface.vaSpread1)
    
    For i = 1 To frmInterface.vaSpread1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        End If
    Next i
     
    gAllExam1 = sExamCode
    
    GetExamCode_Equip = 1
    
End Function


Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sID As String
    
    Dim lsPID As String
    Dim lsPname As String
    Dim lsDate As String
    
    '환자정보 가져오기
    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    lsDate = Format(Date, "yyyymmdd")
    
    If sID = "" Then
        Exit Function
    End If
    
    '바코드, 병록번호, 환자명, 검체코드, 검체명
    
    SQL = "SELECT A.SPCM_NO, A.PID , B.PT_NM , A.SPCM_CD , c.SPCM_ENM " & vbCrLf & _
          "FROM MS.MSLRCPT A " & vbCrLf & _
          "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
          "INNER JOIN HO.PCPPATIENT B ON A.PID = B.PID " & vbCrLf & _
          "INNER JOIN MS.MSLSPCMM C ON A.SPCM_CD = C.SPCM_CD " & vbCrLf & _
          "WHERE A.SPCM_NO = '" & sID & "' " & vbCrLf & _
          "AND AA.EXMN_CD IN (" & gAllExam & ") " & vbCrLf & _
          "GROUP BY A.SPCM_NO, A.PID, B.PT_NM, A.SPCM_CD, C.SPCM_ENM"
    res = db_select_Col(gServer, SQL)
    
    If res = 1 Then
        lsPID = Trim(gReadBuf(1))
        lsPname = Trim(gReadBuf(2))
        
        SetText vasID, lsPID, asRow, colPID
        SetText vasID, lsPname, asRow, colPName
    End If
    
End Function

Private Sub SSPanel1_Click()
    If Text1.Visible = True Then
        Text1.Visible = False
    Else
        Text1.Visible = True
    End If
    
    If Command1.Visible = True Then
        Command1.Visible = False
    Else
        Command1.Visible = True
    End If
End Sub

Private Sub sspMode_Click()
    If sspMode.Caption = "수정모드" Then
        sspMode.Caption = "전송모드"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "전송모드" Then
        sspMode.Caption = "수정모드"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False

    Advia_Init
End Sub

Private Sub Timer2_Timer()
    If dtpToday <> Date Then
        dtpToday = Date
    End If
End Sub


Private Sub txtReceBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim iRow As Integer
    Dim lsBarcode As String
    If KeyCode = 13 Then
        lsBarcode = Trim(txtReceBarcode.Text)
        iRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = lsBarcode Then
                iRow = i
                Exit For
            End If
        Next
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If iRow > vasID.MaxRows Then
                vasID.MaxRows = iRow
            End If
        End If
        SetText vasID, lsBarcode, iRow, colBarCode
        If Trim(GetText(vasID, iRow, colPID)) = "" Then
            Get_Sample_Info iRow
            SetText vasID, "Order", iRow, colState
            
        End If
        txtReceBarcode.Text = ""
    End If
    
End Sub

Private Sub txtUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gExamUID = txtUID.Text
        Call WritePrivateProfileString("CONFIG", "UID", txtUID.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    Dim lsTempBarCode As String
    Dim lsPID As String
    Dim lsPname As String
    Dim lsSex As String
    Dim lsAge As String
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    ClearSpread vasRes
    vasRes.MaxRows = 0
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
        
    
    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE  " & vbCrLf & _
          "  equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND Barcode = '" & Trim(GetText(vasID, Row, colBarCode)) & "' " & vbCrLf & _
          "  order by seqno, equipcode"

    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sExamCode As String
    
    sExamDate = ""
    sExamDate = Trim(GetText(vasRes, asRow2, colResDate))
    sExamTime = Trim(GetText(vasRes, asRow2, colResTime))
    sExamCode = Trim(GetText(vasRes, asRow2, colExamCode))
    
    If Trim(sExamDate) = "" Then
        sExamDate = Format(Date, "yyyymmdd")
    End If
    
    
    SQL = "select examcode FROM pat_res " & vbCrLf & _
          "WHERE equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "

    res = db_select_Row(gLocal, SQL)
    
    If res > 0 Then
        SQL = "update pat_res set resvalue = '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
              "result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "sendflag = '" & asSend & "', " & vbCrLf & _
              "examdate = '" & sExamDate & "', examtime = '" & sExamTime & "', " & vbCrLf & _
              "EXAMCODE = '" & sExamCode & "' " & vbCrLf & _
              "WHERE equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
              "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "
        res = SendQuery(gLocal, SQL)
        
    Else
        SQL = "insert into pat_res(examdate, equipno, barcode, equipcode, examcode, " & vbCrLf & _
              "refflag, sendflag, seqno, examname, resvalue, " & vbCrLf & _
              "result, examtime, pid, pname) " & vbCrLf & _
              "values('" & sExamDate & "', '" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
              "'', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "'" & sExamTime & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "') "
        res = SendQuery(gLocal, SQL)
    End If
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Private Sub vasID_KeyPress(KeyAscii As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long
    Dim i As Integer
    
    Dim sExamCode   As String
    
    If KeyAscii = 13 Then

        llRow = vasID.ActiveRow
        sSpecID = Trim(GetText(vasID, llRow, colBarCode))

        '샘플의 환자 정보 가져오기
        If Barcode_Gubun(sSpecID) = "Q" Then 'QC
            Get_QC_Info vasID, llRow
            
        Else
            Get_Sample_Info llRow
        End If
        
        For iRow = 1 To vasRes.DataRowCnt
            '/장비코드로 검사코드 불러오기
            sExamCode = ""
            SQL = "Select examcode From EquipExam" & vbCrLf & _
                  " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                  "  And EquipCode = '" & Trim(GetText(vasRes, iRow, colEquipExam)) & "'"
            res = db_select_Row(gLocal, SQL)
            
            For i = 1 To res
                If sExamCode = "" Then
                    sExamCode = "'" & Trim(gReadBuf(i - 1)) & "'"
                Else
                    sExamCode = sExamCode & ", '" & Trim(gReadBuf(i - 1)) & "'"
                End If
            Next i
            
            If sExamCode = "" Then sExamCode = "''"
            
            
            If Barcode_Gubun(sSpecID) = "Q" Then
                SQL = Select_QC_Exam(sSpecID, sExamCode)
            Else
                SQL = "SELECT AA.EXMN_CD " & vbCrLf & _
                      "FROM MS.MSLRCPT A " & vbCrLf & _
                      "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
                      "WHERE A.SPCM_NO = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      "AND AA.RSLT_PRGR_STAT_CD in ('05', '07', '09', '12') " & vbCrLf & _
                      "AND AA.EXMN_CD IN (" & sExamCode & ")"
            End If
            res = db_select_Col(gServer, SQL)
            
            SetText vasRes, gReadBuf(0), iRow, colExamCode
            
            Save_Local_One llRow, iRow, "1"
        Next
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If

'    PopupMenu mnuPop
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    Dim lsTempBarCode As String
    Dim lsPID As String
    Dim lsPname As String
    Dim lsSex As String
    Dim lsAge As String
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    ClearSpread vasListRes
    vasListRes.MaxRows = 0
    
    lsID = Trim(GetText(vasList, Row, colBarCode))
    
    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE  " & vbCrLf & _
          "  equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND Barcode = '" & Trim(GetText(vasList, Row, colBarCode)) & "' " & vbCrLf & _
          "  order by seqno, equipcode"


    res = db_select_Vas(gLocal, SQL, vasListRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If


End Sub

Private Sub vasres_rightclick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim VasidRow As Integer
    Dim VasResRow As Integer
    
    VasidRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If VasidRow < 1 Or VasidRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop

End Sub

Private Sub subDel_Click()
    Dim i As Long
    Dim VasidRow As Integer
    Dim VasResRow As Integer
    Dim x As Long
    Dim j As Long
    Dim c, r, c2, r2

    VasidRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If VasidRow < 1 Or VasidRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If

    If vasRes.IsBlockSelected Or vasRes.SelectionCount Then

        vasRes.BlockMode = True
'        db_BeginTran gLocal
        
        For x = 0 To vasRes.SelectionCount - 1
            vasRes.GetSelection x, c, r, c2, r2
            vasRes.Col = c
            vasRes.Col2 = c2
            vasRes.Row = r
            vasRes.Row2 = r2
            If IsNumeric(r) = True And IsNumeric(r2) = True Then
                If CInt(r) > 0 And CInt(r2) > 0 Then
                    For j = r To r2
                        SQL = "Delete from pat_res where barcode = '" & Trim(GetText(vasID, VasidRow, colBarCode)) & "' " & vbCrLf & _
                              "and equipcode = '" & Trim(GetText(vasRes, j, colEquipExam)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                    Next
                End If
            End If
        Next x
        vasRes.BlockMode = False
'        db_Commit gLocal
        

    End If

'    SQL = "Delete from pat_res where barcode = '" & Trim(GetText(vasID, VasidRow, colBarCode)) & "' " & vbCrLf & _
'          "and equipcode = '" & Trim(GetText(vasRes, VasResRow, colEquipExam)) & "' "
'    res = SendQuery(gLocal, SQL)
    
    vasID_Click colBarCode, VasidRow
    vasRes_Click 3, 1
End Sub

'Private Sub subResDel_Click()
'    Dim i As Long
'    i = vasID.ActiveRow
'    vasID.DeleteRows i, 1
'    If i > vasID.DataRowCnt Then
'        i = vasID.DataRowCnt
'    End If
'    vasID.MaxRows = vasID.DataRowCnt
'    vasActiveCell vasID, i, colBarCode
'    vasID.SetFocus
'End Sub


Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim Response, Help
'''    Dim VasResRow As Long
'''    Dim vasResCol As Long
'''    Dim VasidRow As Long
'''
'''    VasResRow = vasRes.ActiveRow
'''    vasResCol = vasRes.ActiveCol
'''    If KeyCode = vbKeyReturn Then
'''        VasidRow = vasID.ActiveRow
'''        If vasResCol = colResult And _
'''           Trim(GetText(vasRes, VasResRow, colResult)) <> Trim(GetText(vasRes, VasResRow, colResult1)) Then
'''
'''            Response = MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
'''            If Response = vbYes Then
'''                '판정, 델타, 패닉 수정
'''                Check_Result Trim(GetText(vasID, VasidRow, colBarCode)), _
'''                             Trim(GetText(vasID, VasidRow, colPID)), _
'''                             Trim(GetText(vasRes, VasResRow, colExamCode)), _
'''                             Trim(GetText(vasRes, VasResRow, colResult)), _
'''                             VasResRow, Trim(GetText(vasID, VasidRow, colPSex))
'''
'''                SQL = " Update pat_res " & vbCrLf & _
'''                      " Set result = '" & Trim(GetText(vasRes, VasResRow, colResult)) & "', " & vbCrLf & _
'''                      " refFlag = '" & Trim(GetText(vasRes, VasResRow, colRCheck)) & "', " & vbCrLf & _
'''                      " panicFlag = '" & Trim(GetText(vasRes, VasResRow, colPCheck)) & "', " & vbCrLf & _
'''                      " deltaFlag = '" & Trim(GetText(vasRes, VasResRow, colDCheck)) & "' " & vbCrLf & _
'''                      " WHERE examdate = '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
'''                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'''                      "  AND equipcode = '" & Trim(GetText(vasRes, VasResRow, colEquipExam)) & "'" & vbCrLf & _
'''                      "  AND barcode = '" & Trim(GetText(vasID, VasidRow, colBarCode)) & "' "
'''                res = SendQuery(gLocal, SQL)
'''
'''                SetText vasRes, Trim(GetText(vasRes, VasResRow, colResult)), VasResRow, colResult1
'''
'''            End If
'''        End If
'''
'''    End If
End Sub

'''Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
'''                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
'''    Dim sDiffRet, sDiffRet1 As String
'''    Dim PreResult   As String
'''
'''    Dim sResClassCode As String     '결과종류
'''    Dim sLow        As String       '참조치
'''    Dim sHigh       As String
'''    Dim RefRet      As String
'''    Dim sPanicGubun As String
'''    Dim sPanicLow   As String       'Panic
'''    Dim sPanicHigh  As String
'''    Dim PanicRet    As String
'''    Dim sDeltaGubun As String
'''    Dim sDeltaLow   As String       'Delta
'''    Dim sDeltaHigh  As String
'''    Dim DeltaRet    As String
'''
'''    Dim sTmpRece1, sTmpRet1 As String
'''    Dim sTmpRece2, sTmpRet2 As String
'''    Dim sMax_ReceNo As String
'''    Dim i           As Integer
'''    Dim sReceNo     As String
'''    Dim sPID        As String
'''
'''    Dim sTmpStr As String
'''
'''    Check_Result = -1
'''
'''    If argBarCode = "" Then
'''        Exit Function
'''    End If
'''
'''    If argExamCode = "" Then
'''        Exit Function
'''    End If
'''
'''
'''    RefRet = ""
'''    PanicRet = ""
'''    DeltaRet = ""
'''
'''    sDiffRet = argResult
'''    If sDiffRet = "" Then
'''        Check_Result = -1
'''        Exit Function
'''    End If
'''
'''    SQL = " Select ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
'''          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
'''          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
'''          "From ExamMaster " & CR & _
'''          " Where HID = '115' " & CR & _
'''          " And ExamCode = '" & Trim(argExamCode) & "' "
'''    res = db_select_Col(gServer, SQL)
'''
'''    sResClassCode = Trim(gReadBuf(0))
'''    Save_Raw_Data "ErrorPoint 9"
'''    If sResClassCode = "1" Then '숫자
''''참조치 체크
'''        sLow = ""
'''        sHigh = ""
'''
'''        '숫자인지 아닌지 확인
'''        If IsNumeric(sDiffRet) = False Then
'''           'MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
'''           Check_Result = -1
'''           Exit Function
'''        End If
'''
'''        If IsNumeric(gReadBuf(13)) Then
'''            If CInt(gReadBuf(13)) > 0 Then
'''                sTmpStr = "#0."
'''                For i = 1 To CInt(gReadBuf(13))
'''                    sTmpStr = sTmpStr & "0"
'''                Next i
'''            Else
'''                sTmpStr = "#0"
'''            End If
'''            sDiffRet = Format(sDiffRet, sTmpStr)
'''            SetText vasRes, sDiffRet, argRow, colResult
'''            SetText vasRes, sDiffRet, argRow, colResult1
'''        End If
'''        Save_Raw_Data "ErrorPoint 10"
'''        Select Case asSex
'''        Case "M", ""
'''            sLow = Trim(gReadBuf(1))
'''            sHigh = Trim(gReadBuf(2))
'''        Case "F"
'''            sLow = Trim(gReadBuf(3))
'''            sHigh = Trim(gReadBuf(4))
'''        End Select
'''
'''        If sLow = "" And sHigh = "" Then
'''            RefRet = ""
'''        ElseIf sLow = "" And sHigh <> "" And IsNumeric(sHigh) = True And IsNumeric(sDiffRet) = True Then  '이상
'''            If CCur(sHigh) < CCur(sDiffRet) Then
'''                RefRet = "H"
'''            End If
'''        ElseIf sLow <> "" And sHigh = "" And IsNumeric(sLow) = True And IsNumeric(sDiffRet) = True Then   '이하
'''            If CCur(sLow) > CCur(sDiffRet) Then
'''                RefRet = "L"
'''            End If
'''        Else
'''            If IsNumeric(sLow) = True And IsNumeric(sHigh) = True And IsNumeric(sDiffRet) = True Then
'''                If CCur(sLow) > CCur(sDiffRet) Then
'''                    RefRet = "L"
'''                ElseIf CCur(sHigh) < CCur(sDiffRet) Then
'''                    RefRet = "H"
'''                ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
'''                    RefRet = ""
'''                End If
'''            End If
'''        End If
'''        Save_Raw_Data "ErrorPoint 11"
'''
''''Panic 체크
'''        sPanicLow = ""
'''        sPanicHigh = ""
'''
'''        sPanicGubun = Trim(gReadBuf(5))
'''
'''        Select Case asSex
'''        Case "M", ""
'''            sPanicLow = Trim(gReadBuf(6))
'''            sPanicHigh = Trim(gReadBuf(7))
'''        Case "F"
'''            sPanicLow = Trim(gReadBuf(8))
'''            sPanicHigh = Trim(gReadBuf(9))
'''        End Select
'''
'''        If sPanicGubun = "0" Then '상한/하한
'''            If sPanicLow = "" Or sPanicHigh = "" Then
'''                PanicRet = ""
'''            Else
'''                If CCur(sPanicLow) > CCur(sDiffRet) Then
'''                    PanicRet = "L"
'''                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
'''                    PanicRet = "H"
'''                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
'''                    PanicRet = ""
'''                End If
'''            End If
'''            Save_Raw_Data "ErrorPoint 12"
'''        ElseIf sPanicGubun = "1" Then 'percent
'''            If sPanicLow = "" Then
'''                PanicRet = ""
'''            Else
'''                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
'''                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'''                        PanicRet = "L"
'''                    Else
'''                        PanicRet = ""
'''                    End If
'''                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
'''                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'''                        PanicRet = "H"
'''                    Else
'''                        PanicRet = ""
'''                    End If
'''                Else
'''                    PanicRet = ""
'''                End If
'''            End If
'''        End If
'''        Save_Raw_Data "ErrorPoint 13"
'''
''''Delta 체크
'''        sDeltaLow = ""
'''        sDeltaHigh = ""
'''
'''        sTmpRece1 = ""
'''        sTmpRet1 = ""
'''        sTmpRece2 = ""
'''        sTmpRet2 = ""
'''        PreResult = ""
'''
'''        sMax_ReceNo = ""
''''        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
'''        sReceNo = argBarCode
'''
''''        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
''''              " Where HID = '115' " & vbCrLf & _
''''              " And PID = '" & Trim(argPID) & "' " & CR & _
''''              " And ReceNo < '" & argBarCode & "' " & CR & _
''''              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
''''              " Group By Result"
'''
'''        '2004/12/30 이상은 - 정렬부분 추가
'''        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'''              " Where HID = '115' " & CR & _
'''              " And PID = '" & Trim(argPID) & "' " & CR & _
'''              " And ReceNo < '" & argBarCode & "' " & CR & _
'''              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'''              " Group By Result" & CR & _
'''              " Order by 2 desc "
'''        res = db_select_Col(gServer, SQL)
'''        Save_Raw_Data "ErrorPoint 14"
'''        If res > 0 And gReadBuf(0) <> "" Then
'''            PreResult = gReadBuf(0)
'''        Else
'''            PreResult = ""
'''        End If
'''
'''        If PreResult <> "" And IsNumeric(PreResult) Then
'''          'PreResult = Trim(gReadBuf(0))
'''          sDeltaGubun = Trim(gReadBuf(10))
'''
'''          sDeltaLow = Trim(gReadBuf(11))
'''          sDeltaHigh = Trim(gReadBuf(12))
'''          Save_Raw_Data "ErrorPoint 15"
'''            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
''''            sDiffRet = PreResult - sDiffRet
'''            sDiffRet1 = sDiffRet - PreResult
'''            If sDeltaGubun = "0" Then '상한/하한
'''                If sDeltaLow = "" Or sDeltaHigh = "" Then
'''                    DeltaRet = ""
'''                Else
'''                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
'''                        DeltaRet = "L"
'''                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
'''                        DeltaRet = "H"
'''                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
'''                        DeltaRet = ""
'''                    End If
'''                End If
''''            Save_Raw_Data "ErrorPoint 16"
'''            ElseIf sDeltaGubun = "1" Then 'percent
'''               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
'''                  DeltaRet = ""
'''               Else
'''                   If sDeltaLow = "" Then
'''                        DeltaRet = ""
'''                    Else
'''                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
'''                            DeltaRet = "D"
'''                        Else
'''                            DeltaRet = ""
'''                        End If
'''                    End If
'''               End If
'''            End If
'''        End If
''''        Save_Raw_Data "ErrorPoint 17"
'''    ElseIf sResClassCode = "2" Then '문자
'''
'''    End If
'''
'''    SetText vasRes, RefRet, argRow, colRCheck
'''    SetText vasRes, PanicRet, argRow, colPCheck
'''    SetText vasRes, DeltaRet, argRow, colDCheck
'''
'''
'''    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
'''    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
'''    If RefRet = "L" Then
'''        vasRes.Row = argRow
'''        vasRes.Col = colRCheck
'''        vasRes.ForeColor = RGB(65, 105, 225)
'''    Else
'''        vasRes.Row = argRow
'''        vasRes.Col = colRCheck
'''        vasRes.ForeColor = RGB(205, 55, 0)
'''    End If
'''
'''    If PanicRet = "L" Then
'''        vasRes.Row = argRow
'''        vasRes.Col = colPCheck
'''        vasRes.ForeColor = RGB(65, 105, 225)
'''    Else
'''        vasRes.Row = argRow
'''        vasRes.Col = colPCheck
'''        vasRes.ForeColor = RGB(205, 55, 0)
'''    End If
'''
'''    If DeltaRet = "L" Then
'''        vasRes.Row = argRow
'''        vasRes.Col = colDCheck
'''        vasRes.ForeColor = RGB(65, 105, 225)
'''    ElseIf DeltaRet = "D" Then
'''        vasRes.Row = argRow
'''        vasRes.Col = colDCheck
'''        vasRes.ForeColor = RGB(65, 105, 225)
'''    Else
'''        vasRes.Row = argRow
'''        vasRes.Col = colDCheck
'''        vasRes.ForeColor = RGB(205, 55, 0)
'''    End If
'''    Save_Raw_Data "ErrorPoint 18"
'''    '2006/11/06 이상은 - 인증심사로 인해 추가함
'''    '205,55,0
'''    Select Case PanicRet
'''    Case "H", "L"
'''        SetBackColor vasRes, argRow, argRow, 1, vasRes.MaxCols, 255, 255, 100
'''        Exit Function
'''    Case Else
'''        SetBackColor vasRes, argRow, argRow, 1, vasRes.MaxCols, 255, 255, 255
'''    End Select
'''
'''    Select Case DeltaRet
'''    Case "D"
'''        SetBackColor vasRes, argRow, argRow, 1, vasRes.MaxCols, 255, 255, 100
'''        Exit Function
'''    Case Else
'''        SetBackColor vasRes, argRow, argRow, 1, vasRes.MaxCols, 255, 255, 255
'''    End Select
'''
'''    Check_Result = 1
'''    Save_Raw_Data "ErrorPoint 19"
'''End Function

'''Public Function QC_Result(argBarCode As String, argExamCode As String, _
'''                            argResult As String, ByVal argRow As Integer, argRRow As Integer) As Integer
'''    Dim sDiffRet, sDiffRet1 As String
'''    Dim PreResult   As String
'''
'''    Dim sResClassCode As String     '결과종류
'''    Dim sLow        As String       '참조치
'''    Dim sHigh       As String
'''    Dim RefRet      As String
'''
'''    Dim sPart       As String
'''    Dim sEquip      As String
'''    Dim sLevel      As String
'''    Dim sLotNo      As String
'''
'''    Dim sTmpRece1, sTmpRet1 As String
'''    Dim sTmpRece2, sTmpRet2 As String
'''    Dim i           As Integer
'''    Dim sReceNo     As String
'''    Dim sPID        As String
'''
'''    Dim sTmpStr As String
'''
'''    QC_Result = -1
'''
'''    If argBarCode = "" Then
'''        Exit Function
'''    End If
'''
'''    If argExamCode = "" Then
'''        Exit Function
'''    End If
'''
'''
'''    RefRet = ""
'''
'''    sDiffRet = argResult
'''    If sDiffRet = "" Then
'''        QC_Result = -1
'''        Exit Function
'''    End If
'''    sPart = Trim(GetText(vasID, argRow, colJumin))
'''    sEquip = gEquip
'''    sLevel = Trim(GetText(vasID, argRow, colPName))
'''    sLotNo = Trim(GetText(vasID, argRow, colPID))
'''
'''    SQL = "Select Max(q.AppDate), e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh   " & vbCrLf & _
'''          "From QCInItem q, ExamMaster e " & vbCrLf & _
'''          "Where q.LabCode = '" & sPart & "' " & vbCrLf & _
'''          "  and q.EquipCode = '" & sEquip & "' " & vbCrLf & _
'''          "  and q.QCInLevel = '" & sLevel & "' " & vbCrLf & _
'''          "  and q.LotNo = '" & sLotNo & "' " & vbCrLf & _
'''          "  and q.QCBarcode = '" & argBarCode & "' " & vbCrLf & _
'''          "  and q.ExamCode = '" & argExamCode & "' " & vbCrLf & _
'''          "  and q.AppDate >= '1900-01-01' " & vbCrLf & _
'''          "  and e.AppDate = (select Max(c.AppDate) from ExamMaster c Where c.AppDate >= '1900-01-01' and c.ExamCode = q.ExamCode)" & vbCrLf & _
'''          "  and e.ExamCode = q.ExamCode " & vbCrLf & _
'''          "Group by e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh"
'''    res = db_select_Col(gServer, SQL)
'''    sResClassCode = Trim(gReadBuf(1))
'''
'''    If sResClassCode = "1" Then '숫자
'''        '참조치 체크
'''        sLow = ""
'''        sHigh = ""
'''
'''        '숫자인지 아닌지 확인
'''        If IsNumeric(sDiffRet) = False Then
'''           'MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
'''           QC_Result = -1
'''           Exit Function
'''        End If
'''
'''        If IsNumeric(gReadBuf(2)) Then
'''            If CInt(gReadBuf(2)) > 0 Then
'''                sTmpStr = "#0."
'''                For i = 1 To CInt(gReadBuf(2))
'''                    sTmpStr = sTmpStr & "0"
'''                Next i
'''            Else
'''                sTmpStr = "#0"
'''            End If
'''            sDiffRet = Format(sDiffRet, sTmpStr)
'''            SetText vasRes, sDiffRet, argRRow, colResult
'''            SetText vasRes, sDiffRet, argRRow, colResult1
'''        End If
'''
'''        sLow = Trim(gReadBuf(3))
'''        sHigh = Trim(gReadBuf(4))
'''
'''        If sLow = "" And sHigh = "" Then
'''            RefRet = ""
'''        ElseIf sLow = "" And sHigh <> "" Then   '이상
'''            If CCur(sHigh) < CCur(sDiffRet) Then
'''                RefRet = "H"
'''            End If
'''        ElseIf sLow <> "" And sHigh = "" Then   '이하
'''            If CCur(sLow) > CCur(sDiffRet) Then
'''                RefRet = "L"
'''            End If
'''        Else
'''            If CCur(sLow) > CCur(sDiffRet) Then
'''                RefRet = "L"
'''            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
'''                RefRet = "H"
'''            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
'''                RefRet = ""
'''            End If
'''        End If
'''
'''
'''
'''    ElseIf sResClassCode = "2" Then '문자
'''
'''    End If
'''
'''    SetText vasRes, RefRet, argRRow, colRCheck
'''
'''    If RefRet = "L" Then
'''        vasRes.Row = argRRow
'''        vasRes.Col = colRCheck
'''        vasRes.ForeColor = RGB(65, 105, 225)
'''    Else
'''        vasRes.Row = argRRow
'''        vasRes.Col = colRCheck
'''        vasRes.ForeColor = RGB(205, 55, 0)
'''    End If
'''
'''    QC_Result = 1
'''
'''End Function

