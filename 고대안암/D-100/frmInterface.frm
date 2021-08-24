VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Interface Program"
   ClientHeight    =   10440
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   15225
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
   ScaleWidth      =   15225
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   90
      TabIndex        =   10
      Top             =   720
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Interface"
      TabPicture(0)   =   "frmInterface.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   9120
         Left            =   150
         TabIndex        =   24
         Top             =   360
         Width           =   14760
         Begin FPSpread.vaSpread vasResPrint 
            Height          =   7845
            Left            =   1080
            TabIndex        =   59
            Top             =   2490
            Visible         =   0   'False
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   13838
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            EditModeReplace =   -1  'True
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
            MaxCols         =   17
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":047A
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1215
            Left            =   7590
            TabIndex        =   45
            Top             =   3810
            Visible         =   0   'False
            Width           =   4785
            _Version        =   393216
            _ExtentX        =   8440
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
            SpreadDesigner  =   "frmInterface.frx":23AE
         End
         Begin FPSpread.vaSpread vasResTemp 
            Height          =   2355
            Left            =   420
            TabIndex        =   37
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
            SpreadDesigner  =   "frmInterface.frx":2615
         End
         Begin VB.CommandButton cmdVasListWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   36
            Top             =   1110
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CheckBox ChkAll 
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   1170
            Width           =   225
         End
         Begin VB.Frame Frame4 
            Caption         =   "[검사결과조회]"
            Height          =   735
            Left            =   180
            TabIndex        =   25
            Top             =   210
            Width           =   14385
            Begin VB.CommandButton cmdResDel 
               Caption         =   "결과삭제"
               Height          =   405
               Left            =   9750
               TabIndex        =   56
               Top             =   210
               Width           =   1095
            End
            Begin VB.CommandButton cmdResPrint 
               Caption         =   "출력"
               Height          =   405
               Left            =   10890
               TabIndex        =   55
               Top             =   210
               Width           =   825
            End
            Begin VB.ComboBox comEquipNo 
               Height          =   315
               Left            =   3030
               TabIndex        =   54
               Text            =   "All"
               Top             =   270
               Width           =   795
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
               Left            =   10890
               TabIndex        =   50
               Top             =   0
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.TextBox txtBarcode 
               Height          =   315
               Left            =   12480
               TabIndex        =   43
               Top             =   240
               Width           =   1785
            End
            Begin VB.ComboBox cmbTransGubun 
               Height          =   315
               ItemData        =   "frmInterface.frx":287C
               Left            =   4380
               List            =   "frmInterface.frx":2889
               TabIndex        =   29
               Text            =   "전체"
               Top             =   270
               Width           =   1065
            End
            Begin VB.CommandButton cmdCall 
               Caption         =   "데이터 조회"
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
               Left            =   5490
               TabIndex        =   28
               Top             =   210
               Width           =   1365
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
               Left            =   6900
               TabIndex        =   27
               Top             =   210
               Width           =   1245
            End
            Begin VB.CommandButton cmdListTrans 
               Caption         =   "검사결과전송"
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
               Left            =   8190
               TabIndex        =   26
               Top             =   210
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker dtpExamDate 
               Height          =   315
               Left            =   1020
               TabIndex        =   30
               Top             =   270
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107282433
               CurrentDate     =   40780
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   13470
               Top             =   600
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label12 
               Caption         =   "상태"
               Height          =   285
               Left            =   3900
               TabIndex        =   60
               Top             =   330
               Width           =   495
            End
            Begin VB.Label Label11 
               Caption         =   "장비"
               Height          =   285
               Left            =   2580
               TabIndex        =   53
               Top             =   330
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "검색"
               Height          =   225
               Left            =   11970
               TabIndex        =   33
               Top             =   330
               Width           =   555
            End
            Begin VB.Label Label2 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   120
               TabIndex        =   32
               Top             =   330
               Width           =   915
            End
            Begin VB.Label Label3 
               Caption         =   "구분"
               Height          =   225
               Left            =   3120
               TabIndex        =   31
               Top             =   -60
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin FPSpread.vaSpread vasListRes 
            Height          =   7875
            Left            =   8490
            TabIndex        =   41
            Top             =   3000
            Visible         =   0   'False
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
            MaxCols         =   9
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":28A1
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   7845
            Left            =   180
            TabIndex        =   58
            Top             =   1080
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   13838
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            EditModeReplace =   -1  'True
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
            MaxCols         =   17
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":3381
         End
      End
      Begin VB.Frame Frame3 
         Height          =   9120
         Left            =   -74850
         TabIndex        =   16
         Top             =   360
         Width           =   14760
         Begin FPSpread.vaSpread vasVres 
            Height          =   2835
            Left            =   5070
            TabIndex        =   116
            Top             =   2040
            Visible         =   0   'False
            Width           =   4365
            _Version        =   393216
            _ExtentX        =   7699
            _ExtentY        =   5001
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
            SpreadDesigner  =   "frmInterface.frx":5278
         End
         Begin FPSpread.vaSpread vasTux 
            Height          =   2775
            Left            =   -30
            TabIndex        =   61
            Top             =   6210
            Visible         =   0   'False
            Width           =   7905
            _Version        =   393216
            _ExtentX        =   13944
            _ExtentY        =   4895
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
            SpreadDesigner  =   "frmInterface.frx":54DF
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "결과삭제"
            Height          =   405
            Left            =   3570
            TabIndex        =   57
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox txtBuff2 
            Height          =   1215
            Left            =   11340
            TabIndex        =   52
            Top             =   4110
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   465
            Left            =   11400
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   1155
            Left            =   11940
            TabIndex        =   22
            Top             =   7860
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Height          =   1125
            Left            =   8550
            TabIndex        =   21
            Top             =   6420
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   615
            Left            =   9630
            TabIndex        =   49
            Top             =   5460
            Visible         =   0   'False
            Width           =   2145
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2025
            Left            =   -1020
            TabIndex        =   48
            Top             =   6930
            Visible         =   0   'False
            Width           =   11955
            _Version        =   393216
            _ExtentX        =   21087
            _ExtentY        =   3572
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
            SpreadDesigner  =   "frmInterface.frx":5746
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   4575
            Left            =   -2010
            TabIndex        =   47
            Top             =   6120
            Visible         =   0   'False
            Width           =   12135
            _Version        =   393216
            _ExtentX        =   21405
            _ExtentY        =   8070
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
            SpreadDesigner  =   "frmInterface.frx":59AD
         End
         Begin VB.TextBox txtReceBarcode 
            Height          =   315
            Left            =   9240
            TabIndex        =   44
            Top             =   300
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.TextBox txtData 
            Height          =   1215
            Left            =   11580
            TabIndex        =   40
            Top             =   6600
            Visible         =   0   'False
            Width           =   2715
         End
         Begin FPSpread.vaSpread vasOrderBuf 
            Height          =   1215
            Left            =   7200
            TabIndex        =   39
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
            SpreadDesigner  =   "frmInterface.frx":5C14
         End
         Begin FPSpread.vaSpread vasOrder 
            Height          =   1215
            Left            =   7200
            TabIndex        =   38
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
            SpreadDesigner  =   "frmInterface.frx":A0FD
         End
         Begin VB.CommandButton cmdVasIDWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   210
            TabIndex        =   35
            Top             =   810
            Visible         =   0   'False
            Width           =   405
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
            Top             =   240
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Trans 
            Caption         =   "검사결과전송"
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
            Top             =   240
            Width           =   1875
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
            TabIndex        =   46
            Top             =   780
            Width           =   14385
            _Version        =   393216
            _ExtentX        =   25374
            _ExtentY        =   14420
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            EditModeReplace =   -1  'True
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
            MaxCols         =   17
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":E5E6
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8175
            Left            =   6750
            TabIndex        =   23
            Top             =   780
            Visible         =   0   'False
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
            MaxCols         =   9
            MaxRows         =   100
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":104DD
         End
         Begin VB.Label Label5 
            Caption         =   "BARCODE : "
            Height          =   285
            Left            =   7830
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   1005
         End
      End
   End
   Begin Threed.SSPanel sspMode 
      Height          =   525
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   926
      _Version        =   131072
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
      Caption         =   "전송모드"
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2670
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      InputLen        =   1
      RThreshold      =   1
      EOFEnable       =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   979
      _Version        =   131072
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
      Caption         =   "INTERFACE"
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.TextBox txtUName 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   270
         Left            =   9000
         TabIndex        =   117
         Top             =   150
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5700
         Top             =   60
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   270
         Left            =   6480
         TabIndex        =   15
         Top             =   90
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8640
         Picture         =   "frmInterface.frx":10FCF
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   14
         Top             =   150
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
         Format          =   107282432
         CurrentDate     =   40778
      End
      Begin MSCommLib.MSComm MSComm2 
         Left            =   3180
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InBufferSize    =   4096
         InputLen        =   1
         RThreshold      =   1
         EOFEnable       =   -1  'True
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
      Left            =   -30
      TabIndex        =   2
      Top             =   6750
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
         SpreadDesigner  =   "frmInterface.frx":11559
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
         SpreadDesigner  =   "frmInterface.frx":15AA6
      End
   End
   Begin VB.Frame Frame5 
      Height          =   9120
      Left            =   60
      TabIndex        =   62
      Top             =   1770
      Visible         =   0   'False
      Width           =   14760
      Begin VB.Frame Frame6 
         Caption         =   "[QC조회]"
         Height          =   735
         Left            =   180
         TabIndex        =   103
         Top             =   210
         Width           =   14415
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
            Left            =   7200
            TabIndex        =   108
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
            Left            =   6000
            TabIndex        =   107
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton cmdQCPrint 
            Caption         =   "결과출력"
            Height          =   405
            Left            =   8520
            TabIndex        =   106
            Top             =   210
            Width           =   1395
         End
         Begin VB.CommandButton cmdQCDel 
            Caption         =   "결과삭제"
            Height          =   405
            Left            =   9960
            TabIndex        =   105
            Top             =   210
            Width           =   1335
         End
         Begin VB.ComboBox cmbQCEquipNo 
            Height          =   315
            ItemData        =   "frmInterface.frx":15D0D
            Left            =   5100
            List            =   "frmInterface.frx":15D0F
            TabIndex        =   104
            Text            =   "1"
            Top             =   240
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpSumSDate 
            Height          =   315
            Left            =   1110
            TabIndex        =   109
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   107282433
            CurrentDate     =   40780
         End
         Begin MSComCtl2.DTPicker dtpSumEDate 
            Height          =   315
            Left            =   2940
            TabIndex        =   110
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   107282433
            CurrentDate     =   40780
         End
         Begin VB.Label Label7 
            Caption         =   "-"
            Height          =   225
            Left            =   2730
            TabIndex        =   113
            Top             =   330
            Width           =   135
         End
         Begin VB.Label Label6 
            Caption         =   "검사일자"
            Height          =   225
            Left            =   180
            TabIndex        =   112
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label9 
            Caption         =   "장비"
            Height          =   225
            Left            =   4620
            TabIndex        =   111
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "[Instrument I]"
         Height          =   7965
         Left            =   150
         TabIndex        =   90
         Top             =   1020
         Width           =   14445
         Begin VB.Frame Frame8 
            Caption         =   "[Low]"
            Height          =   1095
            Left            =   240
            TabIndex        =   95
            Top             =   6810
            Width           =   4125
            Begin VB.Label lblL1 
               Caption         =   "Mean :"
               Height          =   255
               Left            =   150
               TabIndex        =   98
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label lblSDL1 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   97
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblCVL1 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   96
               Top             =   810
               Width           =   2055
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "[High]"
            Height          =   1095
            Left            =   4530
            TabIndex        =   91
            Top             =   6810
            Width           =   4125
            Begin VB.Label lblH1 
               Caption         =   "Mean :"
               Height          =   195
               Left            =   150
               TabIndex        =   94
               Top             =   240
               Width           =   1995
            End
            Begin VB.Label lblSDH1 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   93
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblCVH1 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   92
               Top             =   810
               Width           =   2055
            End
         End
         Begin FPSpread.vaSpread vasEquipL1 
            Height          =   6165
            Left            =   240
            TabIndex        =   99
            Top             =   630
            Width           =   4095
            _Version        =   393216
            _ExtentX        =   7223
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":15D11
         End
         Begin FPSpread.vaSpread vasEquipH1 
            Height          =   6165
            Left            =   4530
            TabIndex        =   100
            Top             =   630
            Width           =   4095
            _Version        =   393216
            _ExtentX        =   7223
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":176A2
         End
         Begin VB.Label Label8 
            Caption         =   "[Low Control]"
            Height          =   315
            Left            =   210
            TabIndex        =   102
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "[High Control]"
            Height          =   315
            Left            =   4560
            TabIndex        =   101
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "[Instrument I]"
         Height          =   7995
         Left            =   9510
         TabIndex        =   76
         Top             =   4770
         Visible         =   0   'False
         Width           =   4755
         Begin VB.Frame Frame10 
            Caption         =   "[High]"
            Height          =   1125
            Left            =   2400
            TabIndex        =   81
            Top             =   6810
            Width           =   2265
            Begin VB.Label lblCVH2 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   84
               Top             =   810
               Width           =   2055
            End
            Begin VB.Label lblSDH2 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   83
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblH2 
               Caption         =   "Mean :"
               Height          =   195
               Left            =   150
               TabIndex        =   82
               Top             =   240
               Width           =   1995
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "[Low]"
            Height          =   1125
            Left            =   90
            TabIndex        =   77
            Top             =   6810
            Width           =   2265
            Begin VB.Label lblCVL2 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   80
               Top             =   810
               Width           =   2055
            End
            Begin VB.Label lblSDL2 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   79
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblL2 
               Caption         =   "Mean :"
               Height          =   255
               Left            =   150
               TabIndex        =   78
               Top             =   240
               Width           =   2055
            End
         End
         Begin FPSpread.vaSpread vasEquipL2 
            Height          =   6165
            Left            =   90
            TabIndex        =   85
            Top             =   630
            Width           =   2265
            _Version        =   393216
            _ExtentX        =   3995
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":19033
         End
         Begin FPSpread.vaSpread vasEquipH2 
            Height          =   6165
            Left            =   2400
            TabIndex        =   86
            Top             =   630
            Width           =   2265
            _Version        =   393216
            _ExtentX        =   3995
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1A9D6
         End
         Begin VB.Label Label16 
            Caption         =   "[High Control]"
            Height          =   315
            Left            =   2430
            TabIndex        =   88
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "[Low Control]"
            Height          =   315
            Left            =   120
            TabIndex        =   87
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "[Instrument I]"
         Height          =   7995
         Left            =   11370
         TabIndex        =   63
         Top             =   5040
         Visible         =   0   'False
         Width           =   4755
         Begin VB.Frame Frame15 
            Caption         =   "[Low]"
            Height          =   1125
            Left            =   90
            TabIndex        =   68
            Top             =   6810
            Width           =   2265
            Begin VB.Label lblL3 
               Caption         =   "Mean :"
               Height          =   255
               Left            =   150
               TabIndex        =   71
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label lblSDL3 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   70
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblCVL3 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   69
               Top             =   810
               Width           =   2055
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "[High]"
            Height          =   1125
            Left            =   2400
            TabIndex        =   64
            Top             =   6810
            Width           =   2265
            Begin VB.Label lblH3 
               Caption         =   "Mean :"
               Height          =   195
               Left            =   150
               TabIndex        =   67
               Top             =   240
               Width           =   1995
            End
            Begin VB.Label lblSDH3 
               Caption         =   "SD :"
               Height          =   225
               Left            =   150
               TabIndex        =   66
               Top             =   510
               Width           =   2055
            End
            Begin VB.Label lblCVH3 
               Caption         =   "CV :"
               Height          =   195
               Left            =   150
               TabIndex        =   65
               Top             =   810
               Width           =   2055
            End
         End
         Begin FPSpread.vaSpread vasEquipL3 
            Height          =   6165
            Left            =   90
            TabIndex        =   72
            Top             =   630
            Width           =   2265
            _Version        =   393216
            _ExtentX        =   3995
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1C379
         End
         Begin FPSpread.vaSpread vasEquipH3 
            Height          =   6165
            Left            =   2400
            TabIndex        =   73
            Top             =   630
            Width           =   2265
            _Version        =   393216
            _ExtentX        =   3995
            _ExtentY        =   10874
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1DD1C
         End
         Begin VB.Label Label25 
            Caption         =   "[Low Control]"
            Height          =   315
            Left            =   120
            TabIndex        =   75
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "[High Control]"
            Height          =   315
            Left            =   2430
            TabIndex        =   74
            Top             =   330
            Width           =   1695
         End
      End
      Begin FPSpread.vaSpread vasQCPrint 
         Height          =   6105
         Left            =   11010
         TabIndex        =   89
         Top             =   4680
         Visible         =   0   'False
         Width           =   8505
         _Version        =   393216
         _ExtentX        =   15002
         _ExtentY        =   10769
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":1F6BF
      End
      Begin FPSpread.vaSpread vasSumTemp 
         Height          =   1785
         Left            =   11430
         TabIndex        =   114
         Top             =   6000
         Visible         =   0   'False
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   3149
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
         SpreadDesigner  =   "frmInterface.frx":21062
      End
      Begin FPSpread.vaSpread vasSum 
         Height          =   4815
         Left            =   10920
         TabIndex        =   115
         Top             =   5400
         Visible         =   0   'False
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   8493
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
         SpreadDesigner  =   "frmInterface.frx":212C9
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
Const colEquipNum = 5
Const colRack = 6

Const colRcnt = 7
Const colState = 8
Const colEquipCode = 9
Const colExamCode = 10
Const colResult = 11
Const colResFlag = 12

Const colIFCC = 13
Const colEAG = 14
Const colArea = 15

Const colExamDTTM = 16
Const colRemark = 17

Const colRStart = 8
' 장비코드 검사코드 검사명 수치결과 문자결과 seq
Const colEquipExam = 1
'''Const colExamCode = 2
Const colExamName = 3
Const colResValue = 4
'''Const colResult = 5
Const colSeq = 6
Const colResDate = 7
Const colResTime = 8

Public gRow As Long
Dim sOrder As String
Dim ConfirmData As String
Dim sSampleType As String
Dim lsFlag As String
Dim llRow As Long

Dim gRemark As String
Dim gTotalRes As Boolean
Dim gTotalRes1 As Boolean
Dim gP3Res As Boolean
Dim gP4Res As Boolean
Dim gEDS As Boolean
Dim gA1c As Boolean
Dim gAoRes As Boolean
Dim gA1cTime As Boolean
Dim gLA1cRes As Boolean
Dim gHbFRes As Boolean
Dim gFArea As Boolean
Dim gUnknown As Boolean
Dim giTotalRes As Integer




Dim gQCGubun As String
Dim gQCLevel As String
Dim gQCRes As String
Dim gAreaRes As String
Dim gUnknownCnt As Integer


Dim strA1a As String
Dim strA1b As String
Dim strA1c As String

Dim strgFRes As String
Dim strgVRes As String
Dim strgP3Res As String
Dim strgTotalRes As String
Dim strgUnknown As String



Dim in_spc_no$, spc_no$(), tst_cd$(), tst_nm$()
Dim spc_cd$(), tst_frct_cd$(), tst_frct_nm$()
Dim tst_dte$(), tst_time$(), work_no$()
Dim pt_no$(), pt_nm$(), Sex$(), birthday$(), Intbase$()
Dim b_dept$(), b_ord_site$()
Dim mmftp As New clsFTP     'FTP관련

Dim acpt_no$()




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
'''    Connect_Server
    For VasidRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = VasidRow

        If vasID.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then

            liRet = -1
            liRet = Insert_Data(VasidRow, vasID)
            
            If liRet = 1 Then
                'db_Commit gServer

                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "Trans", VasidRow, colState
            Else
                SetBackColor vasID, VasidRow, VasidRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "Failed", VasidRow, colState
            End If
            vasID.Col = 1
            vasID.Row = VasidRow
            vasID.Value = 0
        Else

        End If
    Next VasidRow
    
End Sub

Function Insert_Data(argSpcRow As Integer, argSpread As vaSpread) As Integer
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
    Dim sResDateTime As String
    Dim sResIFCC As String
    Dim sResEAG As String
    Dim iResCnt As Integer
    Dim sEquipName As String
    Dim iResState As Integer
    Dim strTrnasRes As String
    Dim strErrRemark As String
    Dim strExamSeq As String
    Dim strResTemp As String
    Dim strResCnt As String
    Dim oerrmsg$
    Dim ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), i_qcOrdcd$(), igubun$()
    Dim MyFileName As String
    Dim FileNumber
'''    Dim argSpcRow As Integer
    Dim intArgRow As Integer
    Dim strAllStr   As String
    Dim strResFlag As String
    Dim strPSex As String
    Dim strReceDate As String
    Dim strReceNo As String
    Dim strRefLow As String
    Dim strRefHigh As String
    Dim msgres
    Dim strNGSP As String
    Dim strIFCC As String
    Dim strResStr As String
    Dim blRemark As Boolean
    Dim streAG As String
    Dim sQCURL As String
    Dim strCode As String
    Dim intSendCnt As Integer
    Dim strSendRes As String
    Dim strSendMsg As String
    Dim strReturnMsg As String
    Dim iRes As Integer
    Dim iiRes As Integer
    Dim vSub As Variant
    Dim vRes As Variant
'''    Dim iRow As Integer
    Dim vTrans As Variant
'''    Dim iResCnt As Integer
    Dim bTrans As Boolean
    Dim strPName As String
    Dim vID As String * 10
    Dim vExamID As String * 5
    Dim vEquip As String * 4
    Dim vResData(10) As String * 255
    Dim vsRes As String * 255
    
    
    strAllStr = ""
    
            
    Insert_Data = -1
    
    blRemark = False
    bTrans = False
    
    lsID = ""
    lsID = Trim(GetText(argSpread, argSpcRow, colBarCode))
    vID = lsID
    
    sTransDate = Format(Date, "yyyymmdd")
    
    sResDateTime = Trim(GetText(argSpread, argSpcRow, colExamDTTM))
    
    strPName = Trim(GetText(argSpread, argSpcRow, colPName))
    
'''    If IsNumeric(lsID) = False Then Exit Function
    If Len(lsID) < 9 Then Exit Function
'''            If Trim(GetText(argSpread, argSpcRow, colRemark)) <> "" Then Exit Function
    
    If InStr(1, UCase(lsID), "UNKNOWN") > 0 Then
    
    Else
        
'''        MsgBox "1"
        
'''        iRes = TuxedoInit("kuh_test", "tuxedo", "01")
'''        MsgBox "2"
        
'''        DoSleep 100
        
        iiRes = ExaminfoList("", vID, "", vRes)
        
        If iiRes = -1 Then
            iiRes = ExaminfoList("", vID, "", vRes)
        End If
        
'''        vRes = vResData
        
'''        If InStr(1, strPName, "HbA1c") > 0 Or InStr(1, lsID, "C2") Then
'''            res = TuxedoTerm
'''            res = TuxedoInit("kuh_test", "tuxedo", "01")
'''        End If
        
        
'''        iRes = TuxedoTerm
'''        iiRes = 15

'''        MsgBox "3"

'''        Save_Raw_Data CStr(iiRes)
        
'''        ReDim vRes(iiRes - 1)
        
'''        MsgBox "4"
'''    vRes(0) = "01035371|장석인|M|58|2031|160|2016-11-11|BL3340|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(1) = "01035371|장석인|M|58|2030|160|2016-11-11|BL3360|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(2) = "01035371|장석인|M|58|2026|216|2016-11-11|BG2243|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(3) = "01035371|장석인|M|58|2024|646|2016-11-11|BG2200|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(4) = "01035371|장석인|M|58|2019|646|2016-11-11|BG2210|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(5) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2411|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(6) = "01035371|장석인|M|58|2033|646|2016-11-11|BM3825|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(7) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2441A|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(8) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2442|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(9) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2443|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(10) = "01035371|장석인|M|58|2022|646|2016-11-11|BJ2570|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(11) = "01035371|장석인|M|58|2021|646|2016-11-11|BJ2580|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(12) = "01035371|장석인|M|58|2023|646|2016-11-11|BJ2602|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(13) = "01035371|장석인|M|58|2020|646|2016-11-11|BM3720|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(14) = "01035371|장석인|M|58|2029|646|2016-11-11|BM3780|2016-11-11|885|DM|101|Serum|/|N||"

        
        'Local에서 환자별로 결과값 가져오기
        ClearSpread vasTemp
        
        SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun, a.result_ifcc, a.result_eag, a.equipnum, a.errremark, b.seqno, a.refflag " & vbCrLf & _
              " From pat_res a, equipexam b " & vbCrLf & _
              " Where a.equipno = b.equipno " & vbCrLf & _
              " And a.examcode = b.examcode " & vbCrLf & _
              " And a.equipcode = b.equipcode " & vbCrLf & _
              " And a.equipno = '" & gEquip & "' " & vbCrLf & _
              " And a.barcode = '" & lsID & "' and resdate = '" & sResDateTime & "' "
              
        SQL = SQL & "group by a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun, a.result_ifcc, a.result_eag, a.equipnum, a.errremark, b.seqno, a.refflag"
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
'''        SQL = "SELECT PERSON_SEX FROM TRUST WHERE REQUEST_DATE = '" & strReceDate & "' AND EXAM_NO = '" & strReceNo & "' "
'''        res = db_select_Col(gServer, SQL)
'''
'''        strPSex = Trim(gReadBuf(0))
'''
'''        If strPSex <> "F" Then strPSex = "M"
'''
'''        sCnt = ""
'''    '''    cn_Ser.BeginTrans
'''        '서버로 결과값 저장하기
'''        strResTemp = ""
        intSendCnt = 0
        
        iResCnt = 0
        
        
        For i = 1 To vasTemp.DataRowCnt
'''            iResCnt = i
            strResStr = ""
            
            sEquipName = gEquip ''& Trim(GetText(vasTemp, i, 8))
            
            sExamCode = Trim(GetText(vasTemp, i, 2))
            sResValue = Trim(GetText(vasTemp, i, 3))
            sResult = Trim(GetText(vasTemp, i, 4))
            sResGubun = Trim(GetText(vasTemp, i, 5))
            sResIFCC = Trim(GetText(vasTemp, i, 6))
            sResEAG = Trim(GetText(vasTemp, i, 7))
            strErrRemark = Trim(GetText(vasTemp, i, 9))
            If Trim(strErrRemark) = "" Then
'''                strErrRemark = " "
                blRemark = True
                strErrRemark = ""
            Else
                strErrRemark = "*"
                
            End If
            
            strExamSeq = Trim(GetText(vasTemp, i, 10))
            strResFlag = Trim(GetText(vasTemp, i, 11))
            If sResGubun = "1" Then '문자
                sTransRes = sResValue & "(" & sResult & ")"

            Else
                sTransRes = sResValue
                sResult = ""
            End If
            
            If iiRes > 0 Then
                For j = 0 To iiRes - 1
                    vSub = Split(CStr(vRes(j)), "|")
                    
                    Save_Raw_Data lsID & " - 처방정보 - " & CStr(vRes(j))
                    If UBound(vSub) > 10 Then
                        If Trim(vSub(7)) = sExamCode Then
                            strSendMsg = ""
                            '검체번호|환자번호|처방일자|처방순번|검사코드|결과|검사특기사항|DeltaYn|PanicYn|정상치Flag|장비플래그|
                            strSendMsg = lsID & "|" & Trim(vSub(0)) & "|" & Trim(vSub(8)) & "|" & Trim(vSub(4)) & "|" & Trim(vSub(7))
                            strSendMsg = strSendMsg & "|" & sTransRes & "|||||" & strErrRemark & "|"
                            
                            Save_Raw_Data strSendMsg
                            
'''                            iResCnt = 0
                            If iResCnt > 0 Then
                                ReDim Preserve vTrans(iResCnt)
                            Else
                                ReDim vTrans(iResCnt)
                            End If
''                            vsRes = strSendMsg
                            
                            vTrans(iResCnt) = strSendMsg
                            bTrans = True
                            iResCnt = iResCnt + 1
                            Exit For
                            
                        End If
    
                    End If
                Next
            End If
        
'''            If sExamCode <> "" And sTransRes <> "" Then
'''                intSendCnt = intSendCnt + 1
'''                strSendRes = strSendRes & sExamCode & vbTab & sTransRes & vbLf
'''
'''                If strCode = "" Then
'''                    strCode = "'" & sExamCode & "'"
'''                Else
'''                    strCode = strCode & ", '" & sExamCode & "'"
'''                End If
'''
'''            End If
        Next i

        If bTrans = False Then
            Exit Function
            
        End If
        
        strSendMsg = ""
'''        strSendMsg = sEquipName & vbTab & gExamUID & vbTab & sTransDate & vbTab & Mid(lsID, 1, 10) & vbTab & "0" & vbTab & CStr(intSendCnt) & vbTab & strSendRes
''''''        Save_Raw_Data sEquipName & vbTab & gExamUID & vbTab & sTransDate & vbTab & Mid(lsID, 1, 10) & vbTab & "0" & vbTab & CStr(intSendCnt) & vbTab & strSendRes
'''
''''''        send string = 'XXX' + TAB + userid + TAB + 검사일자 + TAB +gubn + TAB + cnt + TAB + 처방코드+ TAB + 결과 + NL
'''
'''        strReturnMsg = InsertResultTux(strSendMsg)
        
'''        MsgBox "1"
        
        
'''        iRes = TuxedoInit("kuh_test", "tuxedo", "01")
        
'''        DoSleep 100
        
        
        
        
        vExamID = gExamUID
        vEquip = vEquip
        
        
        iiRes = ResultList("3", vExamID, vTrans, vEquip, "N")
        
        Save_Raw_Data "Result - " & iiRes
        
        If iiRes < 0 Then
            Insert_Data = -9
            Exit Function
        End If
        
'''        DoSleep 100
'''        MsgBox "2"
'''        iRes = TuxedoTerm
    
        
'''        If InStr(1, strPName, "HbA1c") > 0 Or InStr(1, lsID, "C2") Then
'''            res = TuxedoTerm
'''            res = TuxedoInit("kuh_test", "tuxedo", "01")
'''
'''        End If
        
        
        SQL = "update pat_res " & vbCrLf & _
              " set sendflag = '2' " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' and resdate = '" & sResDateTime & "' "
        res = SendQuery(gLocal, SQL)


    End If
                    
    Insert_Data = 1
    
End Function

Function Insert_Data_1(intRow As Integer) As Integer
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
    lsID = Trim(GetText(vasList, intRow, colBarCode))
    
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
                
    '''            SQL = "select eqpm_rslt_valu from mslintrslt " & vbCrLf & _
    '''                  " where rslt_sqno = '" & sRsltSqno & "' "
    '''            res = db_select_Col(gServer, SQL)
    '''            If res = -1 Then
    '''                Save_Raw_Data "[QueryErr]" & SQL
    '''                Exit Function
    '''
    '''            End If
                    
                
    '''            db_BeginTran gServer
                
    '''            If res > 0 Then
    '''                SQL = "update mslintrslt " & vbCrLf & _
    '''                      "set " & vbCrLf & _
    '''                      "  , " & vbCrLf & _
    '''                      "where rslt_sqno = '" & sRsltSqno & "' "
    '''                res = SendQuery(gServer, SQL)
    '''                If res = -1 Then
    '''                    Save_Raw_Data "[QueryErr]" & SQL
    '''                    db_RollBack
    '''                    Exit Function
    '''
    '''                End If
    '''
    '''            Else
                SQL = "insert into mslintrslt (rslt_sqno, rslt_trms_date, rslt_trms_time, eqpm_cd, eqpm_rslt_valu, " & vbCrLf & _
                      "eqpm_rslt_dvcd, err_valu, init_eqpm_rslt_valu, updt_eqpm_rslt_valu, eqpm_rslt_rmrk, " & vbCrLf & _
                      "eqpm_rcpn_sqno, rslt_prgr_stat_cd, frst_rgst_usid, frst_rgdt, last_updt_usid, last_uddt) " & vbCrLf & _
                      "values ( " & vbCrLf & _
                      "'" & sRsltSqno & "','" & sTransDate & "','" & sTransTime & "', " & vbCrLf & _
                      "'" & gEquip & "','" & sResValue & "', " & vbCrLf & _
                      "'','','" & sResValue & "', " & vbCrLf & _
                      "'','', " & vbCrLf & _
                      "'" & sRcpnSqno & "','09', '" & gExamUID & "', " & vbCrLf & _
                      "SYSDATE,'" & gExamUID & "',SYSDATE " & vbCrLf & _
                      ")"
                res = SendQuery(gServer, SQL)
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    db_RollBack gServer
                    Exit Function
                    
                End If
    '''            End If
                
                SQL = "UPDATE MS.MSLGNRLRSLT " & vbCrLf & _
                      "SET   RSLT_PRGR_STAT_CD = '07',  --결과저장(예비결과)  " & vbCrLf & _
                      "       NMVL_RSLT_VALU = '" & sResValue & "',  " & vbCrLf & _
                      "       TXT_RSLT_VALU = '" & sResValue & "', " & vbCrLf & _
                      "       NRML_DVCD = '', " & vbCrLf & _
                      "       DELT_YN = '', " & vbCrLf & _
                      "       PANC_YN = '', " & vbCrLf & _
                      "       ALRT_YN = '', " & vbCrLf & _
                      "       EXMN_RSLT_STOR_DATE = TO_CHAR(SYSDATE, 'YYYYMMDD'), " & vbCrLf & _
                      "       EXMN_RSLT_STOR_TIME = TO_CHAR(SYSDATE, 'HH24MISS'), " & vbCrLf & _
                      "       EXMN_RSLT_STOR_PRSN_ID = '" & gExamUID & "', " & vbCrLf & _
                      "       LAST_UPDT_USID = '" & gExamUID & "', " & vbCrLf & _
                      "       LAST_UDDT = SYSTIMESTAMP, EXMN_EQPM_CD = '" & gEquip & "' " & vbCrLf & _
                      " WHERE RSLT_SQNO = '" & sRsltSqno & "' AND RSLT_PRGR_STAT_CD <> '11' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    db_RollBack gServer
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
                    db_RollBack gServer
                    Exit Function
                    
                End If
            
                db_Commit gServer
            End If
            
        End If
'''        DoSleep 50
    Next i
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, intRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data_1 = 1
    
End Function

Private Sub cmdCall_Click()
    Dim i As Long
    Dim varSendFlag
    Dim j As Long
    Dim x As Long
    Dim strResult As String
    Dim strPastRes As String

    ClearSpread vasList

    varSendFlag = cmbTransGubun.ListIndex

    SQL = "select '', barcode, pid, pname, equipnum, diskno, count(result), sendflag, equipcode, examcode, resvalue, refflag, result_ifcc, result_eag, refvalue, resdate, errremark  from pat_res " & vbCrLf & _
          " where equipno = '" & gEquip & "' and examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    Else
        SQL = SQL & " and sendflag <> '0' "
    End If
    If comEquipNo.ListIndex = 0 Then
    Else
        SQL = SQL & " and equipnum = '" & comEquipNo.Text & "' "
    End If
    
    SQL = SQL & vbCrLf & " group by resdate, barcode, pid, equipnum, pname,  sendflag,  diskno, equipcode, examcode, resvalue, result_ifcc, result_eag, refvalue, errremark, refflag"
    res = db_select_Vas(gLocal, SQL, vasList)


    vasList.maxrows = vasList.DataRowCnt
    For i = 1 To vasList.DataRowCnt
        strPastRes = ""
        Select Case Trim(GetText(vasList, i, colEquipNum))
        Case "1"
            SetForeColor vasList, i, i, 0, 0, 0
        Case "2"
            SetForeColor vasList, i, i, 50, 50, 250
        Case "3"
            SetForeColor vasList, i, i, 20, 150, 20
        End Select
        
        
        If GetText(vasList, i, colState) = "1" Then
            SetText vasList, "Result", i, colState
'''            SetForeColor vasList, i, i, 230, 0, 0
        ElseIf GetText(vasList, i, colState) = "2" Then
            SetText vasList, "Trans", i, colState
            SetBackColor vasList, i, i, colCheckBox, colCheckBox, 202, 255, 112
        End If
        
        If Trim(GetText(vasList, i, colRemark)) <> "" Then
            SetBackColor vasList, i, i, colRemark, colRemark, 240, 180, 180
        End If

'''        If i > 1 Then
'''            For j = i - 1 To 1 Step -1
'''                If Trim(GetText(vasList, i, colBarCode)) = Trim(GetText(vasList, j, colBarCode)) Then
'''                    SetText vasList, Trim(GetText(vasList, j, colResult)), i, colPastRes
'''
'''                    Exit For
'''                End If
'''            Next
'''        End If
        SQL = "select result from pat_res "
        SQL = SQL & " where barcode = '" & Trim(GetText(vasList, i, colBarCode)) & "' "
        SQL = SQL & "   and resdate < '" & Trim(GetText(vasList, i, colExamDTTM)) & "' order by resdate desc"
        res = db_select_Col(gLocal, SQL)
        strPastRes = Trim(gReadBuf(0))
        If strPastRes <> "" Then
            SetText vasList, "R(" & strPastRes & ")", i, colResFlag
        End If

    Next

'''    For i = vasList.DataRowCnt To 1 Step -1
'''        For j = i - 1 To 1 Step -1
'''            If Trim(GetText(vasList, i, colBarCode)) = Trim(GetText(vasList, j, colBarCode)) Then
'''                DeleteRow vasList, j, j
'''            End If
'''
'''        Next
'''    Next
'''
'''    For i = 1 To vasList.DataRowCnt
'''        If Trim(GetText(vasList, i, colPastRes)) <> "" Then
'''            SetBackColor vasList, i, i, colBarCode, colPastRes, 255, 255, 180
'''        End If
'''    Next
'''
    vasList.SetSelection colBarCode, vasList.DataRowCnt, colBarCode, vasList.DataRowCnt


End Sub


Private Sub cmdClear_Click()
Dim iNumber As Integer
Dim i As Integer
    
    txtMsg.Text = ""
    
'''    ClearSpread vasID, 1, 1
'''    vasID.MaxRows = 0
    ClearSpread vasRes, 1, 1
    vasRes.maxrows = 0
    
    For i = vasID.DataRowCnt To 1 Step -1
        vasID.Col = colCheckBox
        vasID.Row = i
        If vasID.Value = 1 Then
            DeleteRow vasID, i, i
        End If
    Next
End Sub

Private Sub cmdCSV_Click()
    Dim i As Long
    Dim j As Long
    Dim strCSV As String
    Dim strFileName As String
    Dim FilNum
    
    CommonDialog1.Filter = "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowSave

    strFileName = CommonDialog1.filename

'''    strFileName = App.Path & "\Res.csv"
    
    strCSV = ""
    If Trim(strFileName) <> "" Then
        For i = 0 To vasList.DataRowCnt
            For j = 1 To vasList.MaxCols
                If j = 1 Or j = 8 Or j = 9 Or j = 10 Then
                Else
                    strCSV = strCSV & Trim(GetText(vasList, i, j)) & ","
                End If
            Next j
            strCSV = strCSV & vbCrLf
            
        Next i
        
        FilNum = FreeFile
        Open strFileName For Output As FilNum
        
        Print #FilNum, strCSV
        Close FilNum
    
    End If
    
'''    Call ShellExecute(Me.hwnd, "OPEN", strFileName, vbNullString, vbNullString, 5)
    
'''    Shell strFileName
    
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer
    Dim iRt
    
    iRt = MsgBox("선택한 결과를 삭제하시겠습니까?", vbYesNo, "결과삭제")
    If iRt = 7 Then Exit Sub
    
    For i = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 1 Then
            SQL = "delete from pat_res where resdate = '" & Trim(GetText(vasID, i, colExamDTTM)) & "' "
            SQL = SQL & " and barcode = '" & Trim(GetText(vasID, i, colBarCode)) & "' "
            SQL = SQL & " and equipnum = '" & Trim(GetText(vasID, i, colEquipNum)) & "'"
            res = SendQuery(gLocal, SQL)
        End If
        
    Next
    
    cmdClear_Click
End Sub

Private Sub cmdListClear_Click()
    Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasList, 1, 1
    vasList.maxrows = 0
    ClearSpread vasListRes, 1, 1
    vasListRes.maxrows = 0
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
'''    Connect_Server
    For VasidRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = VasidRow

        If vasList.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then

            liRet = -1
            liRet = Insert_Data(VasidRow, vasList)
            
            If liRet = 1 Then
                'db_Commit gServer
                SetBackColor vasList, VasidRow, VasidRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasList, "Trans", VasidRow, colState
'''                SetForeColor vasList, VasidRow, VasidRow, 0, 0, 0
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


Private Sub cmdQCDel_Click()
    Dim i As Integer
    Dim sMsg
    
    
    sMsg = MsgBox("선택한 QC결과를 삭제하시겠습니까", vbYesNo)
    
    If sMsg = vbNo Then
        Exit Sub
    End If
    
    
    For i = 1 To vasEquipL1.DataRowCnt
        vasEquipL1.Col = 3
        vasEquipL1.Row = i
        If vasEquipL1.Value = 1 Then
            SQL = "delete from qc_res " & vbCrLf & _
                  "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'LC' " & vbCrLf & _
                  "  and examdatetime = '" & Trim(GetText(vasEquipL1, i, 1)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
    Next
    
    For i = 1 To vasEquipH1.DataRowCnt
        vasEquipH1.Col = 3
        vasEquipH1.Row = i
        If vasEquipH1.Value = 1 Then
            SQL = "delete from qc_res " & vbCrLf & _
                  "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'HC' " & vbCrLf & _
                  "  and examdatetime = '" & Trim(GetText(vasEquipH1, i, 1)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
    Next
    
    
    cmdSumSch_Click
End Sub

Private Sub cmdQCPrint_Click()
    Dim i As Integer
    Dim iMeanRow As Integer
    Dim iCVRow As Integer
    Dim iSDRow As Integer
    Dim j As Integer
    
    
    ClearSpread vasQCPrint
    
    vasQCPrint.maxrows = vasEquipL1.DataRowCnt + 4
    If vasQCPrint.maxrows < vasEquipH1.DataRowCnt + 4 Then
        vasQCPrint.maxrows = vasEquipH1.DataRowCnt + 4
    End If
    For i = 1 To vasEquipL1.DataRowCnt
        SetText vasQCPrint, Trim(GetText(vasEquipL1, i, 1)), i, 1
        SetText vasQCPrint, Trim(GetText(vasEquipL1, i, 2)), i, 2
    
    Next
    
    iMeanRow = vasEquipL1.DataRowCnt + 2
    iSDRow = vasEquipL1.DataRowCnt + 3
    iCVRow = vasEquipL1.DataRowCnt + 4
    
    j = InStr(1, lblL1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "Mean", iMeanRow, 1
        SetText vasQCPrint, Trim(Mid(lblL1.Caption, j + 1)), iMeanRow, 2
    End If
    
    j = InStr(1, lblSDL1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "SD", iSDRow, 1
        SetText vasQCPrint, Trim(Mid(lblSDL1.Caption, j + 1)), iSDRow, 2
    End If
    
    j = InStr(1, lblCVL1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "CV", iCVRow, 1
        SetText vasQCPrint, Trim(Mid(lblCVL1.Caption, j + 1)), iCVRow, 2
    End If
    
    
    For i = 1 To vasEquipH1.DataRowCnt
        SetText vasQCPrint, Trim(GetText(vasEquipH1, i, 1)), i, 3
        SetText vasQCPrint, Trim(GetText(vasEquipH1, i, 2)), i, 4
    Next
    
    iMeanRow = vasEquipH1.DataRowCnt + 2
    iSDRow = vasEquipH1.DataRowCnt + 3
    iCVRow = vasEquipH1.DataRowCnt + 4
    
    j = InStr(1, lblH1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "Mean", iMeanRow, 3
        SetText vasQCPrint, Trim(Mid(lblH1.Caption, j + 1)), iMeanRow, 4
    End If
    
    j = InStr(1, lblSDH1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "SD", iSDRow, 3
        SetText vasQCPrint, Trim(Mid(lblSDH1.Caption, j + 1)), iSDRow, 4
    End If
    
    j = InStr(1, lblCVH1.Caption, ":")
    If j > 0 Then
        SetText vasQCPrint, "CV", iCVRow, 3
        SetText vasQCPrint, Trim(Mid(lblCVH1.Caption, j + 1)), iCVRow, 4
    End If
    
    
    Dim iRow As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    
    
    If vasQCPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    
    sCurDate = Format(Date, "yyyy-mm-dd")
    
    sSerDate = Format(Date, "yyyy-mm-dd")
    
    vasQCPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasQCPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasQCPrint.PrintJobName = "QC 결과 출력"
    
    sHead = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "Variant II No." & cmbQCEquipNo.Text & " Control Data" & " /n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 한국필의료재단"
    
    vasQCPrint.PrintHeader = sHead
    vasQCPrint.PrintFooter = sFoot

    vasQCPrint.PrintMarginTop = 500
    vasQCPrint.PrintMarginBottom = 500
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasQCPrint.PrintMarginLeft = 700
    vasQCPrint.PrintMarginRight = 700
    
    vasQCPrint.PrintColor = True
    vasQCPrint.PrintGrid = True
    
'Set printing range
    vasQCPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasQCPrint.PrintShadows = True

    vasQCPrint.Action = 13 'SS_ACTION_PRINT
    
    
    
End Sub

Private Sub cmdResDel_Click()
    Dim i As Integer
    Dim iRt
    
    iRt = MsgBox("선택한 결과를 삭제하시겠습니까?", vbYesNo, "결과삭제")
    If iRt = 7 Then Exit Sub
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            SQL = "delete from pat_res where resdate = '" & Trim(GetText(vasList, i, colExamDTTM)) & "' "
            SQL = SQL & " and barcode = '" & Trim(GetText(vasList, i, colBarCode)) & "' "
            SQL = SQL & " and equipnum = '" & Trim(GetText(vasList, i, colEquipNum)) & "'"
            res = SendQuery(gLocal, SQL)
        End If
        
    Next
    
    cmdCall_Click
End Sub

Private Sub cmdResPrint_Click()
    Dim i As Integer
    Dim iMeanRow As Integer
    Dim iCVRow As Integer
    Dim iSDRow As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    
    ClearSpread vasResPrint
    
    For i = 1 To vasList.DataRowCnt
        For j = 1 To vasList.MaxCols
            SetText vasResPrint, Trim(GetText(vasList, i, j)), i, j
            
        Next
    Next
    
    If vasResPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    
    sCurDate = Format(Date, "yyyy-mm-dd")
    
    sSerDate = Format(Date, "yyyy-mm-dd")
    
    vasResPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasResPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasResPrint.PrintJobName = "QC 결과 출력"
    
    sHead = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "Variant II Result Data" & " /n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 한국필의료재단"
    
    vasResPrint.PrintHeader = sHead
    vasResPrint.PrintFooter = sFoot

    vasResPrint.PrintMarginTop = 500
    vasResPrint.PrintMarginBottom = 500
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasResPrint.PrintMarginLeft = 700
    vasResPrint.PrintMarginRight = 700
    
    vasResPrint.PrintColor = True
    vasResPrint.PrintGrid = True
    
'Set printing range
    vasResPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasResPrint.PrintShadows = True

    vasResPrint.Action = 13 'SS_ACTION_PRINT
    
    
End Sub

Private Sub cmdSugaClear_Click()
    ClearSpread vasSum
    vasSum.maxrows = 0
End Sub

Private Sub cmdSumSch_Click()
    Dim i As Long
    Dim sMean As Currency
    Dim sSum As Currency
    Dim sDlb As Currency
    Dim sSDSum As Currency
    Dim sCVSum As Currency
    
    
    ClearSpread vasEquipL1
    ClearSpread vasEquipL2
    ClearSpread vasEquipH3
    ClearSpread vasEquipH1
    ClearSpread vasEquipH2
    ClearSpread vasEquipH3
    
    
    'Equip 1 ===========================================================================================================
    
    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
          "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'LC' " & vbCrLf & _
          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
          "order by examdatetime "
    res = db_select_Vas(gLocal, SQL, vasEquipL1)
    
    sSum = 0
    sMean = 0
    For i = 1 To vasEquipL1.DataRowCnt
        If IsNumeric(GetText(vasEquipL1, i, 2)) = True Then
            sSum = sSum + CCur(GetText(vasEquipL1, i, 2))
        End If
    Next
    
    If vasEquipL1.DataRowCnt > 0 Then
        sMean = sSum / vasEquipL1.DataRowCnt
        sMean = Format(sMean, "##0.0")
        lblL1.Caption = "Mean : " & CStr(sMean)
    End If
    
    'SD, CV 결과 구하기 ============================================================================================
    If IsNumeric(sMean) = True Then
        sSDSum = 0
        SQL = "select stdev(result) from qc_res " & vbCrLf & _
              "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'LC' " & vbCrLf & _
              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
        res = db_select_Col(gLocal, SQL)
        If IsNumeric(gReadBuf(0)) = True Then
            sSDSum = Trim(gReadBuf(0))
            
    '''        For i = 1 To vasEquipL1.DataRowCnt
    '''            If IsNumeric(GetText(vasEquipL1, i, 2)) = True Then
    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipL1, i, 2)) - sMean) ^ 2
    '''            End If
    '''
    '''        Next
            
            
            If vasEquipL1.DataRowCnt > 0 And IsNumeric(sSDSum) = True Then
    '''            sSDSum = sSDSum / vasEquipL1.DataRowCnt
                sCVSum = 0
                If sMean <> 0 Then
                    sCVSum = sSDSum / sMean * 100
                End If
                
                sSDSum = Format(sSDSum, "##0.00")
                sCVSum = Format(sCVSum, "##0.00")
                lblSDL1.Caption = "SD : " & CStr(sSDSum)
                lblCVL1.Caption = "CV : " & CStr(sCVSum)
                
            End If
        End If
    End If
    
    '===========================================================================================================

    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
          "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'HC' " & vbCrLf & _
          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
          "order by examdatetime "
    res = db_select_Vas(gLocal, SQL, vasEquipH1)
    
    sSum = 0
    sMean = 0
    
    For i = 1 To vasEquipH1.DataRowCnt
        If IsNumeric(GetText(vasEquipH1, i, 2)) = True Then
            sSum = sSum + CCur(GetText(vasEquipH1, i, 2))
        End If
    Next
    
    If vasEquipH1.DataRowCnt > 0 Then
        sMean = sSum / vasEquipH1.DataRowCnt
        sMean = Format(sMean, "##0.0")
        lblH1.Caption = "Low : " & CStr(sMean)
    End If
    
    
    'SD, CV 결과 구하기 ============================================================================================
    If IsNumeric(sMean) = True Then
        sSDSum = 0
        SQL = "select stdev(result) from qc_res " & vbCrLf & _
              "where equipno = '" & cmbQCEquipNo.Text & "' and levelname = 'HC' " & vbCrLf & _
              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
        res = db_select_Col(gLocal, SQL)
        If IsNumeric(gReadBuf(0)) = True Then
            sSDSum = Trim(gReadBuf(0))
            
    '''        For i = 1 To vasEquipH1.DataRowCnt
    '''            If IsNumeric(GetText(vasEquipH1, i, 2)) = True Then
    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipH1, i, 2)) - sMean) ^ 2
    '''            End If
    '''
    '''        Next
            
            If vasEquipH1.DataRowCnt > 0 Then
    '''            sSDSum = sSDSum / vasEquipH1.DataRowCnt
                sCVSum = 0
                If sMean <> 0 Then
                    sCVSum = sSDSum / sMean * 100
                End If
                
                sSDSum = Format(sSDSum, "##0.00")
                sCVSum = Format(sCVSum, "##0.00")
                lblSDH1.Caption = "SD : " & CStr(sSDSum)
                lblCVH1.Caption = "CV : " & CStr(sCVSum)
                
            End If
        End If
    End If
    
    '===========================================================================================================
    
    
    sSum = 0
    sMean = 0
    
''' 'Equip 2 ===========================================================================================================
'''        SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '2' and levelname = 'LC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipL1)
'''
'''    sSum = 0
'''    sMean = 0
'''    For i = 1 To vasEquipL2.DataRowCnt
'''        If IsNumeric(GetText(vasEquipL2, i, 2)) = True Then
'''            sSum = sSum + CCur(GetText(vasEquipL2, i, 2))
'''        End If
'''    Next
'''
'''    If vasEquipL2.DataRowCnt > 0 Then
'''        sMean = sSum / vasEquipL2.DataRowCnt
'''        sMean = Format(sMean, "##0.0")
'''        lblL2.Caption = "Mean : " & CStr(sMean)
'''    End If
'''
'''    'SD, CV 결과 구하기 ============================================================================================
'''    If IsNumeric(sMean) = True Then
'''        sSDSum = 0
'''        SQL = "select stdev(result) from qc_res " & vbCrLf & _
'''              "where equipno = '2' and levelname = 'LC' " & vbCrLf & _
'''              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
'''        res = db_select_Col(gLocal, SQL)
'''        If IsNumeric(gReadBuf(0)) = True Then
'''            sSDSum = Trim(gReadBuf(0))
'''
'''    '''        For i = 1 To vasEquipL1.DataRowCnt
'''    '''            If IsNumeric(GetText(vasEquipL1, i, 2)) = True Then
'''    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipL1, i, 2)) - sMean) ^ 2
'''    '''            End If
'''    '''
'''    '''        Next
'''
'''
'''            If vasEquipL2.DataRowCnt > 0 And IsNumeric(sSDSum) = True Then
'''    '''            sSDSum = sSDSum / vasEquipL1.DataRowCnt
'''                sCVSum = 0
'''                If sMean <> 0 Then
'''                    sCVSum = sSDSum / sMean * 100
'''                End If
'''
'''                sSDSum = Format(sSDSum, "##0.00")
'''                sCVSum = Format(sCVSum, "##0.00")
'''                lblSDL2.Caption = "SD : " & CStr(sSDSum)
'''                lblCVL2.Caption = "CV : " & CStr(sCVSum)
'''
'''            End If
'''        End If
'''    End If
'''
'''    '===========================================================================================================
'''
'''    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '2' and levelname = 'HC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipH2)
'''
'''    sSum = 0
'''    sMean = 0
'''
'''    For i = 1 To vasEquipH2.DataRowCnt
'''        If IsNumeric(GetText(vasEquipH2, i, 2)) = True Then
'''            sSum = sSum + CCur(GetText(vasEquipH2, i, 2))
'''        End If
'''    Next
'''
'''    If vasEquipH2.DataRowCnt > 0 Then
'''        sMean = sSum / vasEquipH2.DataRowCnt
'''        sMean = Format(sMean, "##0.0")
'''        lblH2.Caption = "Low : " & CStr(sMean)
'''    End If
'''
'''    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '2' and levelname = 'LC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipL2)
'''
'''
'''    'SD, CV 결과 구하기 ============================================================================================
'''    If IsNumeric(sMean) = True Then
'''        sSDSum = 0
'''        SQL = "select stdev(result) from qc_res " & vbCrLf & _
'''              "where equipno = '2' and levelname = 'HC' " & vbCrLf & _
'''              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
'''        res = db_select_Col(gLocal, SQL)
'''        If IsNumeric(gReadBuf(0)) = True Then
'''            sSDSum = Trim(gReadBuf(0))
'''
'''    '''        For i = 1 To vasEquipH1.DataRowCnt
'''    '''            If IsNumeric(GetText(vasEquipH1, i, 2)) = True Then
'''    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipH1, i, 2)) - sMean) ^ 2
'''    '''            End If
'''    '''
'''    '''        Next
'''
'''            If vasEquipH2.DataRowCnt > 0 Then
'''    '''            sSDSum = sSDSum / vasEquipH1.DataRowCnt
'''                sCVSum = 0
'''                If sMean <> 0 Then
'''                    sCVSum = sSDSum / sMean * 100
'''                End If
'''
'''                sSDSum = Format(sSDSum, "##0.00")
'''                sCVSum = Format(sCVSum, "##0.00")
'''                lblSDH2.Caption = "SD : " & CStr(sSDSum)
'''                lblCVH2.Caption = "CV : " & CStr(sCVSum)
'''
'''            End If
'''        End If
'''    End If
'''
'''    '===========================================================================================================
'''
'''
'''    sSum = 0
'''    sMean = 0
''' 'Equip 3 ===========================================================================================================
'''        SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '3' and levelname = 'LC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipL3)
'''
'''    sSum = 0
'''    sMean = 0
'''    For i = 1 To vasEquipL3.DataRowCnt
'''        If IsNumeric(GetText(vasEquipL3, i, 2)) = True Then
'''            sSum = sSum + CCur(GetText(vasEquipL3, i, 2))
'''        End If
'''    Next
'''
'''    If vasEquipL3.DataRowCnt > 0 Then
'''        sMean = sSum / vasEquipL1.DataRowCnt
'''        sMean = Format(sMean, "##0.0")
'''        lblL3.Caption = "Mean : " & CStr(sMean)
'''    End If
'''
'''    'SD, CV 결과 구하기 ============================================================================================
'''    If IsNumeric(sMean) = True Then
'''        sSDSum = 0
'''        SQL = "select stdev(result) from qc_res " & vbCrLf & _
'''              "where equipno = '3' and levelname = 'LC' " & vbCrLf & _
'''              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
'''        res = db_select_Col(gLocal, SQL)
'''        If IsNumeric(gReadBuf(0)) = True Then
'''            sSDSum = Trim(gReadBuf(0))
'''
'''    '''        For i = 1 To vasEquipL1.DataRowCnt
'''    '''            If IsNumeric(GetText(vasEquipL1, i, 2)) = True Then
'''    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipL1, i, 2)) - sMean) ^ 2
'''    '''            End If
'''    '''
'''    '''        Next
'''
'''
'''            If vasEquipL3.DataRowCnt > 0 And IsNumeric(sSDSum) = True Then
'''    '''            sSDSum = sSDSum / vasEquipL1.DataRowCnt
'''                sCVSum = 0
'''                If sMean <> 0 Then
'''                    sCVSum = sSDSum / sMean * 100
'''                End If
'''
'''                sSDSum = Format(sSDSum, "##0.00")
'''                sCVSum = Format(sCVSum, "##0.00")
'''                lblSDL3.Caption = "SD : " & CStr(sSDSum)
'''                lblCVL3.Caption = "CV : " & CStr(sCVSum)
'''
'''            End If
'''        End If
'''    End If
'''
'''    '===========================================================================================================
'''
'''    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '3' and levelname = 'HC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipH3)
'''
'''    sSum = 0
'''    sMean = 0
'''
'''    For i = 1 To vasEquipH3.DataRowCnt
'''        If IsNumeric(GetText(vasEquipH3, i, 2)) = True Then
'''            sSum = sSum + CCur(GetText(vasEquipH3, i, 2))
'''        End If
'''    Next
'''
'''    If vasEquipH3.DataRowCnt > 0 Then
'''        sMean = sSum / vasEquipH3.DataRowCnt
'''        sMean = Format(sMean, "##0.0")
'''        lblH3.Caption = "Low : " & CStr(sMean)
'''    End If
'''
'''    SQL = "select examdatetime, result from qc_res " & vbCrLf & _
'''          "where equipno = '3' and levelname = 'LC' " & vbCrLf & _
'''          "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "order by examdatetime "
'''    res = db_select_Vas(gLocal, SQL, vasEquipL3)
'''
'''
'''    'SD, CV 결과 구하기 ============================================================================================
'''    If IsNumeric(sMean) = True Then
'''        sSDSum = 0
'''        SQL = "select stdev(result) from qc_res " & vbCrLf & _
'''              "where equipno = '3' and levelname = 'HC' " & vbCrLf & _
'''              "and examdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' "
'''        res = db_select_Col(gLocal, SQL)
'''        If IsNumeric(gReadBuf(0)) = True Then
'''            sSDSum = Trim(gReadBuf(0))
'''
'''    '''        For i = 1 To vasEquipH1.DataRowCnt
'''    '''            If IsNumeric(GetText(vasEquipH1, i, 2)) = True Then
'''    '''                sSDSum = sSDSum + (CCur(GetText(vasEquipH1, i, 2)) - sMean) ^ 2
'''    '''            End If
'''    '''
'''    '''        Next
'''
'''            If vasEquipH3.DataRowCnt > 0 Then
'''    '''            sSDSum = sSDSum / vasEquipH1.DataRowCnt
'''                sCVSum = 0
'''                If sMean <> 0 Then
'''                    sCVSum = sSDSum / sMean * 100
'''                End If
'''
'''                sSDSum = Format(sSDSum, "##0.00")
'''                sCVSum = Format(sCVSum, "##0.00")
'''                lblSDH3.Caption = "SD : " & CStr(sSDSum)
'''                lblCVH3.Caption = "CV : " & CStr(sCVSum)
'''
'''            End If
'''        End If
'''    End If
'''
'''    '===========================================================================================================
    
    
    sSum = 0
    sMean = 0
    
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
    Dim s As String
    Dim i As Integer
    
    VariantIIAll Text1.Text
    
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
            strPointFormat = strPointFormat & "0"
        Next
        If strRespRec = "0" Then
            strPointFormat = "##0"
        Else
            strPointFormat = "##0." & strPointFormat
        End If
    
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


Private Sub Command3_Click()
    Dim strS As String
    Dim vaS As Variant
    Dim i As Integer
    Dim iRes As Integer
    
'''    res = ExaminfoList2("1160609037", vaS)
    
    
    iRes = TuxedoInit(sEMRUser, sEMRID, sEMRPW)
    res = ExaminfoList("", "1160609704", "", vaS)
'''    Save_Raw_Data CStr(res)
    iRes = TuxedoTerm
    
    If res > 0 Then
        For i = 0 To res - 1
            Save_Raw_Data CStr(vaS(i))
            
'''            Save_Raw_Data CStr(vaS(i))
        Next
    End If
    
'''    Dim sWorkListURL As String
'''
'''
'''''    Online_Sch gURL.Order & txtReceBarcode.Text
'''
'''    SetText vasID, txtReceBarcode.Text, 1, colBarCode
'''
'''    Get_Sample_Info 1
'''
'''
'''
'''''    sWorkListURL = gURL.Order & txtReceBarcode.Text
'''''
'''''    URLstart sWorkListURL
'''''
'''''
''''''''    ControlURL1.Start "D:\인터페이스_진행\서울건국대학교병원\GyeolGwaIF.xml"
'''''    Save_Raw_Data "[XML]" & gStrXML
    
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
    
    cmdClear_Click
        
    GetSetup    'ini에서 DB정보 불러오기
    
    Init_Form

'''    If Not Connect_Server Then
'''        MsgBox "연결되지 않았습니다."
'''        Exit Sub
'''    End If
    
'''    cn_Server_Flag = dce_setenv("0VAR1.env", "", "")

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    res = TuxedoInit(sEMRUser, sEMRID, sEMRPW)
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = "True"
    MSComm1.DTREnable = "True"
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
   
    lblUser.Caption = gExamUID
    txtUID.Text = gExamUID
    txtUName.Text = gIFName
    
    raw_data = ""

    If gSetup.gUse = "Y" Then
    
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
    End If
    
'''    MSComm2.CommPort = gSetup2.gPort
'''    MSComm2.RTSEnable = "True"
'''    MSComm2.DTREnable = "True"
'''    MSComm2.Settings = gSetup2.gSpeed & "," & gSetup2.gParity & "," & gSetup2.gDataBit & "," & gSetup2.gStopBit
'''
'''    If gSetup2.gUse = "Y" Then
'''        If MSComm2.PortOpen = False Then
'''            MSComm2.PortOpen = True
'''        End If
'''    End If
    dtpToday = Date
    dtpExamDate = Date
    dtpSumSDate = Format(Date, "yyyy/mm")
    dtpSumEDate = Date


    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(dtpToday), -gLocalExpDate), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    res = SendQuery(gLocal, SQL)
    '===================================================================
    
'''    SQL = "update equipexam set equipno = 'H15' "
'''    res = SendQuery(gLocal, SQL)
    
    
    '검사코드 가져오기
    GetExamCode

    ClearSpread vasCode

    vasID.maxrows = 1
    vasID.ColsFrozen = 6
    vasRes.maxrows = 20
    vasList.maxrows = 1
    
    vasList.ColsFrozen = 6
    
    vasListRes.maxrows = 20
    
    vasSum.maxrows = 20
    vasSum.ColsFrozen = 1
    
    cmbQCEquipNo.AddItem "1", 0
    cmbQCEquipNo.AddItem "2", 1
    cmbQCEquipNo.AddItem "3", 2
    
    cmbQCEquipNo.ListIndex = 0
    
    comEquipNo.AddItem "All", 0
    comEquipNo.AddItem "1", 1
    comEquipNo.AddItem "2", 2
    comEquipNo.AddItem "3", 3
    
    comEquipNo.ListIndex = 0
    txtUID.Text = gExamUID
    
    

'''    For i = colRStart + 1 To vasID.MaxCols
'''        vasID.Col = i
'''        vasID.ColHidden = True
'''    Next
'''
'''    For i = colRStart + 1 To vasList.MaxCols
'''        vasList.Col = i
'''        vasList.ColHidden = True
'''    Next


    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    WritePrivateProfileString "config", "UID", txtUID.Text, App.Path & "\interface.ini"
    
'''    DisConnect_Server
    res = TuxedoTerm
    DisConnect_Local
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    
    gAllExam = ""
    gOrderExam = ""
    gReceExam = ""
    
    
    For i = 1 To 100
        gArr_Exam(i, 1) = ""
        gArr_Exam(i, 2) = ""
        gArr_Exam(i, 3) = ""
    Next i
    
    ClearSpread vasTemp
    
    SQL = "Select SeqNo, EquipCode, ExamName, resgubun From EquipExam where Equipno = '" & gEquip & "' " & vbCrLf & _
          " group by SeqNo, EquipCode, ExamName, resgubun"
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
'''    vasID.MaxCols = colRStart + vasTemp.DataRowCnt
'''    vasList.MaxCols = colRStart + vasTemp.DataRowCnt
'''    vasSum.MaxCols = vasTemp.DataRowCnt + 1
    
    For i = 1 To vasTemp.DataRowCnt
        If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
            gArr_Exam(i, 1) = i    '순서
            gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
            gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
            gArr_Exam(i, 4) = Trim(GetText(vasTemp, i, 4))    '결과구분
            
'''            SetText vasID, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
'''            SetText vasList, Trim(GetText(vasTemp, i, 3)), 0, colRStart + i
'''            SetText vasSum, Trim(GetText(vasTemp, i, 3)), 0, i + 1
            
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

Private Sub Label1_DblClick()
    If txtReceBarcode.Visible = False Then
    
        txtReceBarcode.Visible = True
        Command3.Visible = True
        Text1.Visible = True
        Command1.Visible = True
    Else
        txtReceBarcode.Visible = False
        Command3.Visible = False
        Text1.Visible = False
        Command1.Visible = False
        
    End If
    
    
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
    
    Dim s As String

    
    s = MSComm1.Input
    
    Select Case s
  
    Case chrENQ
        Save_Raw_Data "[RX]" & s
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
        txtBuff.Text = s
        
'''    Case chrSTX
'''        txtBuff.Text = chrSTX
        
    Case chrLF
        txtBuff.Text = txtBuff.Text & s
'''        Save_Raw_Data "[RX]" & txtBuff.Text
        
'''        VariantIIAll txtBuff.Text
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
       
    Case chrEOT     '자료수신 완료
        txtBuff.Text = txtBuff.Text & s
        Save_Raw_Data "[RX]" & txtBuff.Text
        VariantIIAll txtBuff.Text
        
        Save_Raw_Data "[RX]" & s
        
        txtBuff.Text = ""
        
    Case Else
        txtBuff.Text = txtBuff.Text & s
    End Select
End Sub


Private Sub MSComm2_OnComm()
    
    Dim s As String

    
    s = MSComm2.Input
    
    Select Case s
  
    Case chrENQ
        Save_Raw_Data "[RX]" & s
        
        MSComm2.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
        txtBuff2.Text = s
'''    Case chrSTX
'''        txtBuff2.Text = chrSTX
        
    Case chrLF
        txtBuff2.Text = txtBuff2.Text & s
'''        Save_Raw_Data "[RX]" & txtBuff2.Text
        
'''        VariantIIAll txtBuff2.Text
        
        MSComm2.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
       
    Case chrEOT     '자료수신 완료
        txtBuff2.Text = txtBuff2.Text & s
        Save_Raw_Data "[RX]" & txtBuff2.Text
        VariantIIAll txtBuff2.Text, "3"
        
        Save_Raw_Data "[RX]" & s
        
        txtBuff2.Text = ""
        
    Case Else
        txtBuff2.Text = txtBuff2.Text & s
    End Select
End Sub


Private Sub VariantIIAll(asData As String, Optional asEquip As String = "1")
    Dim strData As String
    Dim lsTemp As String
    
    Dim i As Long
    
    
    strData = asData
    
    strData = Replace(strData, chrENQ, "")
    strData = Replace(strData, chrEOT, "")
    
    i = InStr(1, strData, chrSTX)
    
    While i > 0
        strData = Mid(strData, 1, i - 1) & Mid(strData, i + 2)
        i = InStr(1, strData, chrSTX)
    Wend
    
    i = InStr(1, strData, chrLF)
    
    While i > 0
        strData = Mid(strData, 1, i - 4) & Mid(strData, i + 1)
        i = InStr(1, strData, vbLf)
    Wend
    
    
    strData = Replace(strData, chrETB, "")
    strData = Replace(strData, chrETX, "")
    strData = strData & chrCR
    
    i = InStr(1, strData, Chr(13))
    Do While i > 0
        lsTemp = Mid(strData, 1, i - 1)
        strData = Mid(strData, i + 1)
        
        
        
'''        Select Case Left(lsTemp, 1)
'''        Case "Q"
'''            lsMSGflag = "Q"
'''        Case "O"
'''            lsMSGflag = "O"
'''        End Select
        
        VariantII lsTemp, asEquip
        
        i = InStr(1, strData, chrCR)
''        If i = 0 Then
''            i = InStr(1, strData, chrETX)
''        End If
    Loop
    

End Sub
Private Sub VariantII(asData As String, Optional asEquipNo As String = "1")
    Dim sdata As String
    Dim i As Integer
    Dim j As Integer
    Dim sSubData(1 To 40) As String
    Dim sResData(1 To 25) As String
    Dim iRes As Integer
    Dim iResStr As Long
    Dim iSub As Integer
    Dim iStr As Long
    Dim iRow As Integer
    Dim strSampleNo As String
    Dim strEquipCode As String
    Dim strExamCode As String
    Dim strExamName As String
    Dim strSeqNo As String
    Dim strResult As String
    Dim strResValue As String
    Dim strTransDate As String
    Dim strTransTime As String
    Dim iResRow As Integer
    Dim liRet As Integer
    Dim strA1 As String
    Dim strReceCode As String
    Dim strRackPos As String
    Dim sResult1 As String
    Dim sIFCC As String
    Dim sEAG As String
    Dim strRefLow As String
    Dim strRefHigh As String
    Dim strResFlag As String
    Dim strPSex As String
    Dim strReceDate As String
    Dim strReceNo As String
    Dim strPastRes As String

    

    iRes = 1
    For iResStr = 1 To 15
        sResData(iResStr) = ""
    Next
    For iResStr = 1 To Len(asData)
        If Mid(asData, iResStr, 1) = "|" Then
            iRes = iRes + 1
            If iRes > 25 Then
                Exit For
            End If
            sResData(iRes) = ""
            
        Else
            sResData(iRes) = sResData(iRes) & Mid(asData, iResStr, 1)
        End If
    Next
    
    Select Case Mid(sResData(1), 1, 1)
    Case "H"
        strA1a = ""
        strA1b = ""
        strA1c = ""
        gInterfaceTime = ""
        gInterfaceTime = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss")
        gEquipNum = ""
        
        
        
    Case "P"
    Case "O"
        gRemark = ""
        gTotalRes = False
        gTotalRes1 = False
        gP3Res = False
        gP4Res = False
        gEDS = False
        gA1c = False
        gAoRes = False
        gA1cTime = False
        gLA1cRes = False
        gHbFRes = False
        gFArea = False
        gUnknown = False
        giTotalRes = 0
        
        strgFRes = ""
        strgVRes = ""
        strgP3Res = ""
        strgTotalRes = ""
        strgUnknown = ""

        
        gQCGubun = "P"
        gQCLevel = ""
        gUnknownCnt = 0
        
        ClearSpread vasRes
        glRow = -1
        
        strSampleNo = sResData(3)
        j = InStr(1, strSampleNo, "^")
        strRackPos = Mid(strSampleNo, j + 1)
        strRackPos = Replace(strRackPos, "^", "-")
        
        strSampleNo = Mid(strSampleNo, 1, j - 1)
        
        If asEquipNo = "3" Then
            gEquipNum = "3"
        Else
            gEquipNum = Trim(sResData(25))
        End If
        
        If IsNumeric(gEquipNum) = False Then
            gEquipNum = "1"
        End If
        
        
        If Mid(strSampleNo, 1, 2) = "HC" Then
            gQCGubun = "Q"
            gQCLevel = Mid(strSampleNo, 3, 1)
            
            Exit Sub
            
        End If
        
        
        j = InStr(1, strRackPos, "-")
        If j > 0 Then
            strRackPos = Mid(strRackPos, j + 1) & "-" & Mid(strRackPos, 1, j - 1)
            
        End If
        
        
        
        
        iRow = -1
'''        For j = 1 To vasID.DataRowCnt
'''            If Trim(GetText(vasID, j, colBarCode)) = strSampleNo Then
'''                iRow = j
'''                Exit For
'''            End If
'''        Next
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If iRow > vasID.maxrows Then
                vasID.maxrows = iRow
            End If
        End If
        
        glRow = iRow
        
        vasID.SetSelection colBarCode, glRow, colBarCode, glRow
        
        SetText vasID, strSampleNo, iRow, colBarCode
        SetText vasID, strRackPos, iRow, colRack
        SetText vasID, gInterfaceTime, iRow, colExamDTTM
        SetText vasID, gEquipNum, iRow, colEquipNum
        
        
        
        If Trim(GetText(vasID, iRow, colPName)) = "" Then
            Get_Sample_Info iRow
            
'''            SetText vasID, Format(Date, "yyyymmdd") & "-" & strSampleNo, iRow, colBarCode
        End If
        
'''        strTransDate = Mid(sResData(15), 1, 8)
'''        strTransTime = Mid(sResData(15), 9, 6)
        
    Case "R"
        
        i = InStr(1, sResData(3), "^^^")
        If i > 0 Then
            strEquipCode = Mid(sResData(3), i + 3)
        End If
        
        strResult = Trim(sResData(4))
        
'''        If strEquipCode = "A1a^AREA" Then
'''            strA1a = strResult
'''        ElseIf strEquipCode = "A1b^AREA" Then
'''            strA1b = strResult
'''        ElseIf strEquipCode = "A1c^AREA" Then
'''            strA1c = strResult
'''        End If

        If strEquipCode = "A1c^AREA" And gQCGubun = "Q" Then
            SQL = "insert into qc_res(equipno, examdatetime, levelname, result, examdate) " & vbCrLf & _
                  "values('" & gEquipNum & "', '" & gInterfaceTime & "', '" & gQCLevel & "', '" & strResult & "', '" & Format(Date, "yyyymmdd") & "') "
            res = SendQuery(gLocal, SQL)
            
        End If
        
        If gQCGubun = "Q" Then Exit Sub
        
        
        If strEquipCode = "TOTAL^AREA" Then
            gAreaRes = ""
            gAreaRes = strResult
            If IsNumeric(strResult) = True Then
                If CCur(strResult) > 350000 Or CCur(strResult) < 50000 Then
                    gTotalRes = True
                   
                    strgTotalRes = strResult
                    
        
                End If
            End If
'''        ElseIf strEquipCode = "A1c^TIME" Then
'''            If IsNumeric(strResult) = True Then
'''                If CCur(strResult) < 0.45 Or CCur(strResult) > 0.55 Then
'''                    gA1cTime = True
'''
'''                End If
'''            End If
'''        ElseIf strEquipCode = "Ao^AREA" Then
'''            If IsNumeric(strResult) = True Then
'''                If CCur(strResult) < 60 Then
'''                    gAoRes = True
'''
'''                End If
'''            End If
        ElseIf strEquipCode = "P3^AREA" Then
            If IsNumeric(strResult) = True Then
                If CCur(strResult) >= 10 Then
                    gP3Res = True
                
                    strgP3Res = strResult
        
                End If
                
            End If
        
        ElseIf strEquipCode = "LA1c^AREA" Then
            If IsNumeric(strResult) = True Then
                If CCur(strResult) > 7 Then

                    gLA1cRes = True
                End If

            End If
'''
'''        ElseIf strEquipCode = "HbF^AREA" Then
'''            If IsNumeric(strResult) = True Then
'''                If CCur(strResult) > 25 Then
'''
'''                    gHbFRes = True
'''                End If
'''
'''            End If
'''
'''        ElseIf strEquipCode = "P4^AREA" Then
'''            If IsNumeric(strResult) = True Then
'''                If CCur(strResult) <= 10 Then
'''
'''                    gP4Res = True
'''                End If
'''
'''            End If
            
        ElseIf strEquipCode = "F^AREA" Then
            If IsNumeric(strResult) = True Then
                If CCur(strResult) >= 30 Then
                    gFArea = True
                    strgFRes = strResult
        
                End If
            End If
        ElseIf strEquipCode = "Unknown^AREA" Then
            If IsNumeric(strResult) = True Then
                If CCur(strResult) >= 10 Then
                    gUnknown = True
                    strgUnknown = strResult
        
                End If
            End If
            
        End If
        
        If strEquipCode = "E^AREA" Or strEquipCode = "D^AREA" Or strEquipCode = "S^AREA" Or strEquipCode = "C^AREA" Then    'E.D.S 항목 추가해야 함
            If IsNumeric(strResult) = True Then
                If CCur(strResult) > 40 Then
                    gEDS = True
                    strgVRes = strResult
        
                End If
            End If
            
            
        End If

        SQL = "SELECT EXAMCODE FROM EQUIPEXAM " & vbCrLf & _
              "WHERE EQUIPNO = '" & gEquip & "' AND EQUIPCODE = '" & strEquipCode & "' "
        res = db_select_Row(gLocal, SQL)
        strReceCode = ""
        strExamCode = Trim(gReadBuf(0))
        
        If res > 0 Then
            strReceCode = ""
            If Trim(GetText(vasID, glRow, colExamCode)) = "" Then
                strExamCode = Trim(gReadBuf(0))
            Else
                strExamCode = Trim(GetText(vasID, glRow, colExamCode))
            End If
            
            
            SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO FROM EQUIPEXAM " & vbCrLf & _
                  "WHERE EQUIPNO = '" & gEquip & "' AND EXAMCODE = '" & strExamCode & "'"
            res = db_select_Col(gLocal, SQL)
            
            strExamName = Trim(gReadBuf(2))
            strSeqNo = Trim(gReadBuf(3))
            

            If strEquipCode = "A1c^AREA" And IsNumeric(strResult) = True Then
'''                strResult = CCur(strResult) + 0.1
                
'''
'''                sIFCC = CStr((CCur(strResult) - 2.15) / 0.0915)  '(NGSP -2.15) / 0.0915 이전식
                sIFCC = CStr((CCur(strResult) * 10.93) - 23.5)   '10.93 * NGSP - 23.50 최근식 2014/01/03
                sIFCC = Format(sIFCC, "##0")
                
                sEAG = CStr((CCur(strResult) * 28.7) - 46.7) '28.7 X NGSP -46.7
                sEAG = Format(sEAG, "##0.0")
                
                If CCur(strResult) < 3.5 Or CCur(strResult) > 19 Then
                    gA1c = True
                    If CCur(strResult) > 19 Then
                        giTotalRes = 2
                    ElseIf CCur(strResult) < 3.5 Then
                        giTotalRes = 1
                    End If
                    
                End If
                
            End If
            
            sResult1 = Result_Set(strExamCode, strResult)
                        
            j = InStr(1, sResult1, "/")
            
            strResValue = Mid(sResult1, 1, j - 1)
            strResult = Mid(sResult1, j + 1)
            
            SetText vasID, strResValue, glRow, colResult
            SetText vasID, sIFCC, glRow, colIFCC
            SetText vasID, sEAG, glRow, colEAG
            SetText vasID, strEquipCode, glRow, colEquipCode
            SetText vasID, strExamCode, glRow, colExamCode
            
            '판정처리해야함
            strResFlag = ""
            strPSex = ""
            
            
            
            
            strSampleNo = Trim(GetText(vasID, glRow, colBarCode))

            Save_Local_One glRow, iResRow, "1"
            
            
            SQL = "select result from pat_res "
            SQL = SQL & " where barcode = '" & Trim(GetText(vasID, glRow, colBarCode)) & "' "
            SQL = SQL & "   and resdate <> '" & Trim(GetText(vasID, glRow, colExamDTTM)) & "' order by resdate desc"
            res = db_select_Col(gLocal, SQL)
            strPastRes = Trim(gReadBuf(0))
            If strPastRes <> "" Then
                SetText vasID, "R(" & strPastRes & ")", glRow, colResFlag
            End If
            
        End If
    
    Case "L"
        gRemark = ""
        
        If gQCGubun = "Q" Then Exit Sub
        
        If gEDS = True Then
            gRemark = gRemark & "Variant Windows,"
        End If
        
'''        If gAoRes = True Then
'''            gRemark = gRemark & "[Ao<60]"
'''        End If
        
        If gTotalRes = True Then
            
            gRemark = gRemark & "Total Area,"
        End If
        
        If gA1c = True Then
'''            gRemark = gRemark & "[A1c]"
            
            If giTotalRes = 1 Then
                gRemark = gRemark & "Low Area,"
            Else
                gRemark = gRemark & "High Area,"
            End If
            
        End If
        
        If gFArea = True Then
            gRemark = gRemark & "F Peak,"
        End If
        
        If gP3Res = True Then
            gRemark = gRemark & "P3 Area,"
        End If
        
        If gUnknown = True Then
            gRemark = gRemark & "Unknown Area,"
        End If

'''        If gP4Res = False Then
'''            gRemark = gRemark & "[P4]"
'''        End If
        
'''        If gA1cTime = True Then
'''            gRemark = gRemark & "[A1c RT]"
'''        End If
'''
        If gLA1cRes = True Then
            gRemark = gRemark & "LA1c Area,"
        End If
        
        
        If gRemark <> "" Then
            gRemark = Mid(gRemark, 1, Len(gRemark) - 1)
        End If
        
'''
'''        If gHbFRes = True Then
'''            gRemark = gRemark & "[HbF]"
'''        End If
        
        
'''        If gUnknownCnt >= 3 Then
'''            gRemark = gRemark & "[Unknown >3]"
'''        End If
        
        
'''        If gRemark <> "" Then
            SetText vasID, gRemark, glRow, colRemark
            SetText vasID, gAreaRes, glRow, colArea
            
            Save_Local_One glRow, iResRow, "1"
            
        gAreaRes = ""
'''        End If
        
        SetText vasID, "Result", glRow, colState
        If Trim(gRemark) <> "" Then
            SetBackColor vasID, glRow, glRow, colRemark, colRemark, 240, 180, 180
            
        End If
        
        Select Case Trim(GetText(vasID, glRow, colEquipNum))
        Case "1"
            SetForeColor vasID, glRow, glRow, 0, 0, 0
        Case "2"
            SetForeColor vasID, glRow, glRow, 50, 50, 250
        Case "3"
            SetForeColor vasID, glRow, glRow, 20, 150, 20
        End Select
        
        
        If mnuAuto.Checked = True Then
'''        If mnuAuto.Checked = True And gTotalRes1 = False Then
            liRet = -1
            liRet = Insert_Data(CInt(glRow), vasID)

            If liRet = 1 Then
                'db_Commit gServer

                SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "Trans", glRow, colState
            Else
                SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "Failed", glRow, colState
            End If
            
            
            If glRow = 1 Then
                liRet = -1
                liRet = Insert_Data(CInt(glRow), vasID)
    
                If liRet = 1 Then
                    'db_Commit gServer
    
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 202, 255, 112
                    SetText vasID, "Trans", glRow, colState
                Else
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 255, 0, 0
                    SetText vasID, "Failed", glRow, colState
                End If
            
            End If
            
            
            If liRet = -9 Then
                res = TuxedoTerm
                res = TuxedoInit(sEMRUser, sEMRID, sEMRPW)
                
                liRet = -1
                liRet = Insert_Data(CInt(glRow), vasID)
    
                If liRet = 1 Then
                    'db_Commit gServer
    
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 202, 255, 112
                    SetText vasID, "Trans", glRow, colState
                Else
                    SetBackColor vasID, glRow, glRow, colCheckBox, colCheckBox, 255, 0, 0
                    SetText vasID, "Failed", glRow, colState
                End If

            End If
            
        End If
    End Select

End Sub
Sub SendOrder()
Dim sSendOrder As String
'''
'''    gOrderCnt = 1
    
    If Len(gOrderMessage) > 240 Then
        
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If

        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
        gOrderMessage = Mid(gOrderMessage, 241)

        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
        SaveQuery sSendOrder, 1

        gOrderCnt = gOrderCnt + 1
        comSend = "stENQ"

        gPreMsg = sSendOrder
        Save_Raw_Data "[TX]" & sSendOrder

        MSComm1.Output = sSendOrder

    Else
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
                
        gOrderMessage = ""
        comSend = "stOrder"
        
        gPreMsg = sSendOrder
        Save_Raw_Data "[TX]" & sSendOrder
        
        MSComm1.Output = sSendOrder
    End If
End Sub

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

    Dim lsID As String
    Dim strReceDate As String
    Dim strReceNo As String
    Dim sWorkListURL As String
    Dim sRece() As String
    Dim sSub() As String
    Dim sSub2() As String
    Dim i As Integer
    Dim j As Integer
    Dim strExamCode As String
    Dim iRow As String
    Dim strS As String
    Dim vRes As Variant
    Dim vSub As Variant
    Dim iRes As Integer
    Dim iiRes As Integer
    Dim sReceCode As String
    Dim vID As String * 10
    Dim vResData(10) As String * 255
    
    
    
'''    Dim j As Integer
'''    Dim iRow As Integer
    
    
    On Error GoTo ErrOcc:
    
    
    
    '환자 정보 가져오기
    
    
    
    ClearSpread vasTux
    
    lsID = Trim(GetText(vasID, asRow, colBarCode))
    vID = lsID
'''    If IsNumeric(lsID) = False Then Exit Function
    If Len(lsID) < 9 Then Exit Function
    
    
'''    iRes = TuxedoInit("kuh_test", "tuxedo", "01")
'''    DoSleep 100
    
    iiRes = ExaminfoList("", vID, "", vRes)
    If iiRes = -1 Then
        iiRes = ExaminfoList("", vID, "", vRes)
    End If
    
'''    vRes = vResData
    
'''    Save_Raw_Data CStr(res)
'''    iRes = TuxedoTerm
'''    iiRes = 14
'''    ReDim vRes(iiRes - 1)
    
'''    vRes(0) = "01035371|장석인|M|58|2031|160|2016-11-11|BL3340|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(1) = "01035371|장석인|M|58|2030|160|2016-11-11|BL3360|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(2) = "01035371|장석인|M|58|2026|216|2016-11-11|BG2243|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(3) = "01035371|장석인|M|58|2024|646|2016-11-11|BG2200|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(4) = "01035371|장석인|M|58|2019|646|2016-11-11|BG2210|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(5) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2411|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(6) = "01035371|장석인|M|58|2033|646|2016-11-11|BM3825|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(7) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2441A|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(8) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2442|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(9) = "01035371|장석인|M|58|2033|646|2016-11-11|BH2443|2016-11-11|885|DM|101|Serum|/|N|GCHE1031|"
'''    vRes(10) = "01035371|장석인|M|58|2022|646|2016-11-11|BJ2570|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(11) = "01035371|장석인|M|58|2021|646|2016-11-11|BJ2580|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(12) = "01035371|장석인|M|58|2023|646|2016-11-11|BJ2602|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(13) = "01035371|장석인|M|58|2020|646|2016-11-11|BM3720|2016-11-11|885|DM|101|Serum|/|N||"
'''    vRes(14) = "01035371|장석인|M|58|2029|646|2016-11-11|BM3780|2016-11-11|885|DM|101|Serum|/|N||"
    
'''    Save_Raw_Data "Sample Info : " & CStr(iiRes)
    
    
    If iiRes > 0 Then
        For i = 0 To iiRes - 1
            
            Save_Raw_Data CStr(vRes(i))
            
            vSub = Split(CStr(vRes(i)), "|")
            If UBound(vSub) > 10 Then
                iRow = vasTux.DataRowCnt + 1
                If iRow > vasTux.maxrows Then
                    iRow = vasTux.maxrows + 1
                End If
                For j = 1 To UBound(vSub)
                    SetText vasTux, Trim(vSub(j - 1)), iRow, j
                Next
                
            End If
'''            Save_Raw_Data CStr(vaS(i))
        Next
    End If
    
    
    If vasTux.DataRowCnt = 0 Then
        SetText vasID, "없음", asRow, colState
    Else
        
        For i = 1 To vasTux.DataRowCnt
            If i = 1 Then
                sReceCode = "'" & Trim(GetText(vasTux, i, 8)) & "'"
            Else
                sReceCode = sReceCode & ", '" & Trim(GetText(vasTux, i, 8)) & "'"
            End If
        Next
        If sReceCode = "" Then sReceCode = "''"
        
        gReadBuf(0) = ""
        SQL = "select examcode from equipexam "
        If sReceCode <> "" Then
            SQL = SQL & " where examcode in (" & sReceCode & ")"
        End If
        
        res = db_select_Col(gLocal, SQL)

        SetText vasID, Trim(gReadBuf(0)), asRow, colExamCode
        
        SetText vasID, Trim(GetText(vasTux, 1, 1)), asRow, colPID
        SetText vasID, Trim(GetText(vasTux, 1, 2)), asRow, colPName
        SetText vasID, "", asRow, colState
    End If
    

    Exit Function
    
ErrOcc:
    SetText vasID, "없음", asRow, colState
    Exit Function
    
End Function

Function Get_Sample_Info_List(ByVal asRow As Long) As Integer

    Dim lsID As String
    Dim strReceDate As String
    Dim strReceNo As String
    
    '환자 정보 가져오기
    lsID = Trim(GetText(vasList, asRow, colBarCode))
    
    If IsNumeric(lsID) = False Then Exit Function
    If Len(lsID) < 13 Then Exit Function
    
    strReceDate = Mid(lsID, 1, 8)
    strReceNo = CInt(Mid(lsID, 9))
    
    
'''    SQL = " Select 병록번호, 품목코드, 성명 " & CR & _
'''          " From 검사검체1V "
'''    SQL = SQL & CR & " Where 검체번호 = '" & lsID & "' "
'''    SQL = SQL & CR & " And 품목코드 IN (" & gAllExam & ") "

    SQL = "SELECT A.PERSON_NAME, B.EXAM_CODE "
    SQL = SQL & vbCrLf & "  FROM TRUST A, TRURES B "
    SQL = SQL & vbCrLf & " WHERE A.REQUEST_DATE = B.REQUEST_DATE  "
    SQL = SQL & vbCrLf & "   AND A.EXAM_NO = B.EXAM_NO  "
    SQL = SQL & vbCrLf & "   AND A.REQUEST_DATE = '" & strReceDate & "'  "
    SQL = SQL & vbCrLf & "   AND A.EXAM_NO = '" & strReceNo & "'  "
    SQL = SQL & vbCrLf & "   AND B.EXAM_CODE IN (" & gAllExam & ")  "
    SQL = SQL & vbCrLf & " GROUP BY A.PERSON_NAME, B.EXAM_CODE "
    res = db_select_Col(gServer, SQL)
    
    If res < 1 Then
        SetText vasList, "없음", asRow, colState
    Else
'''        SetText vasList, Trim(gReadBuf(0)), asRow, colPID
        SetText vasList, Trim(gReadBuf(1)), asRow, colExamCode
        SetText vasList, Trim(gReadBuf(0)), asRow, colPName
        SetText vasList, "", asRow, colState
        
    End If
    
End Function

Function SetResult(asResult As String, asExamCode As String) As String
'DB에서 불러오기
'    Dim iFloat As Integer
    Dim iFloat As String
    
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

'    Select Case aiItem
'    Case 7, 16
'        iFloat = 2
'    Case 14
'        iFloat = 0
'    Case Else
'        iFloat = 1
'    End Select
'
'    If iFloat = 0 Then
'        SetResult = CStr(CCur(asResult))
'    Else
'        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
'    End If
 
    gReadBuf(0) = ""
    
    SQL = " Select Point From ExamMaster " & vbCrLf & _
          " Where HID = '115' " & vbCrLf & _
          " And ExamCode = '" & Trim(asExamCode) & "' " & vbCrLf & _
          " And UseFlag = 'Y' "
    res = db_select_Col(gServer, SQL)
    
    iFloat = gReadBuf(0)
    
    '2004/05/31 이상은
    'ASO 관리자에는 소수점 2자리로 셋팅되어 있으나 1자리로 할 것
    If asExamCode = "C4633AJ" Then   'ASO
        iFloat = 1
    End If
    
    Select Case iFloat
    Case 0
        SetResult = Format(asResult, "#,##0")
    Case 1
        SetResult = Format(asResult, "#,##0.0")
    Case 2
        SetResult = Format(asResult, "#,##0.00")
    Case 3
        SetResult = Format(asResult, "#,##0.000")
    Case Else
    
    End Select
    
End Function




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


'Private Sub subUp_Click()
'Dim sValue As String
'Dim sTmp As String
'Dim i As Integer
'Dim j As Integer
'
'    sTmp = ""
'
'    vasID.Row = vasID.ActiveRow
'    vasID.Col = vasID.ActiveCol
'
'    sTmp = vasID.Text
'
'    sValue = InputBox("변경할 검체번호를 입력하세요")
'
'    If Trim(sValue) <> "" Then
'        If MsgBox("" & sTmp & "를 " & sValue & "로 수정하시겠습니까?", vbYesNo, "확인") = vbYes Then
'            SetText vasID, sValue, vasID.Row, vasID.Col
'
'            If Trim(GetText(vasID, vasID.Row, colBarCode)) <> "" Then
'                Get_Sample_Info vasID.Row
'
'                For i = 1 To vasRes.DataRowCnt
'                    Save_Local_One vasID.Row, i, "A"
'                Next
'            End If
'        End If
'    End If
'End Sub

'''Private Sub txtToday_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim i As Integer
'''
'''    If KeyCode = vbKeyReturn Then
'''
'''    SQL = "select barcode, receno, pid, pname, pjumin, psex, page, '', sendflag from pat_res " & vbCrLf & _
'''          "where examdate = '" & Format(Trim(txtToday), "yyyymmdd") & "' and equipno = '0025' " & vbCrLf & _
'''          "group by barcode, receno, pid, pname, pjumin, psex, page,  sendflag"
'''    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
'''
'''    For i = 1 To vasID.DataRowCnt
'''        If GetText(vasID, i, colState) = "A" Then
'''            SetText vasID, "수신완료", i, colState
'''            SetBackColor vasID, i, i, colCheckBox, colCheckBox, 100, 122, 255
'''        ElseIf GetText(vasID, i, colState) = "B" Then
'''            SetText vasID, "전송완료", i, colState
'''            SetBackColor vasID, i, i, colCheckBox, colCheckBox, 202, 255, 112
'''        End If
'''    Next
'''    End If
'''End Sub

Private Sub Timer1_Timer()
    If dtpToday <> Date Then
        dtpToday = Date
    End If
    
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim varSendFlag
    Dim j As Long
    Dim x As Long
    Dim strResult As String
    Dim strBarcode As String
    Dim strPastRes As String
    
        
    If KeyAscii = 13 Then
        
        If IsNumeric(txtBarcode) = False Then Exit Sub
        
        If Len(txtBarcode) <= 5 Then
            strBarcode = Format(Date, "yyyymmdd") & Format(txtBarcode, "00000")
        Else
            strBarcode = Mid(txtBarcode, 1, 8) & Format(CInt(Mid(txtBarcode, 9)), "00000")
        End If
        txtBarcode = ""
        ClearSpread vasList
    
        varSendFlag = cmbTransGubun.ListIndex
    
        SQL = "select '', barcode, pid, pname, equipnum, diskno, count(result), sendflag, equipcode, examcode, resvalue, refflag, result_ifcc, result_eag, refvalue, resdate, errremark  from pat_res " & vbCrLf & _
              " where equipno = '" & gEquip & "' and barcode = '" & strBarcode & "' "
        
        SQL = SQL & vbCrLf & " group by resdate, barcode, pid, equipnum, pname,  sendflag,  diskno, equipcode, examcode, resvalue, result_ifcc, result_eag, refvalue, errremark, refflag"
        res = db_select_Vas(gLocal, SQL, vasList)
    
    
        vasList.maxrows = vasList.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            If GetText(vasList, i, colState) = "1" Then
                SetText vasList, "Result", i, colState
    '''            SetForeColor vasList, i, i, 230, 0, 0
            ElseIf GetText(vasList, i, colState) = "2" Then
                SetText vasList, "Trans", i, colState
                SetBackColor vasList, i, i, colCheckBox, colCheckBox, 202, 255, 112
            End If
            
            If Trim(GetText(vasList, i, colRemark)) <> "" Then
                SetBackColor vasList, i, i, colRemark, colRemark, 240, 180, 180
            End If
            
            SQL = "select result from pat_res "
            SQL = SQL & " where barcode = '" & Trim(GetText(vasList, i, colBarCode)) & "' "
            SQL = SQL & "   and resdate < '" & Trim(GetText(vasList, i, colExamDTTM)) & "' order by resdate desc"
            res = db_select_Col(gLocal, SQL)
            strPastRes = Trim(gReadBuf(0))
            If strPastRes <> "" Then
                SetText vasList, "R(" & strPastRes & ")", i, colResFlag
            End If
        
        Next

        vasList.SetSelection colBarCode, vasList.DataRowCnt, colBarCode, vasList.DataRowCnt
    End If
End Sub

Private Sub txtReceBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim i As Integer
'''    Dim iRow As Integer
'''    Dim lsBarcode As String
'''    If KeyCode = 13 Then
'''        lsBarcode = Trim(txtReceBarcode.Text)
'''        iRow = -1
''''''        For i = vasID.DataRowCnt To 1 Step -1
''''''            If Trim(GetText(vasID, i, colBarCode)) = lsBarcode Then
''''''                DeleteRow vasID, i, i
''''''
''''''                iRow = i
''''''                Exit For
''''''            End If
''''''        Next
'''
'''        If iRow = -1 Then
'''            iRow = vasID.DataRowCnt + 1
'''            If iRow > vasID.maxrows Then
'''                vasID.maxrows = iRow
'''            End If
'''        End If
'''        SetText vasID, lsBarcode, iRow, colBarCode
'''        If Trim(GetText(vasID, iRow, colPID)) = "" Then
'''            Get_Sample_Info iRow
'''            SetText vasID, "Order", iRow, colState
'''
'''        End If
'''        txtReceBarcode.Text = ""
'''    End If
    
End Sub

Private Sub txtUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gExamUID = txtUID.Text
        Call WritePrivateProfileString("CONFIG", "UID", txtUID.Text, App.Path & "\Interface.ini")
    End If
End Sub

'Private Sub Timer1_Timer()
'    Dim lRow As Long
'    Dim lCnt As Long
'    Dim sID As String
'    Dim sCode As String
'    Dim sDate As String
'    Dim sRack As String
'    Dim sTube As String
'    Dim sNew As String
'    Dim i As Long
'    Dim X As Integer
'
'    If ComState = False Then
'        Exit Sub
'    End If
'
''    Save_Raw_Data "[OrderCnt]" & vasCode.DataRowCnt
'    For i = 1 To vasCode.DataRowCnt
'        sID = Trim(GetText(vasCode, i, 3))
'        sCode = Trim(GetText(vasCode, i, 2))
'        sDate = Trim(GetText(vasCode, i, 4))
'        sRack = Trim(GetText(vasCode, i, 5))
'        sTube = Trim(GetText(vasCode, i, 6))
'        sNew = Trim(GetText(vasCode, i, 7))
'        If sCode <> "" And sID <> "" Then
'            Save_Raw_Data "[TimerCnt]" & vasCode.DataRowCnt
'            Integra800_Order_Entry sID, sDate, sCode, sRack, sTube, sNew
'            DeleteRow vasCode, i, i
'
'            Exit Sub
'        Else
'            DeleteRow vasCode, i, i
'            i = i - 1
'        End If
'    Next i
'
'    If Host_BC = "09" Then
'        For lRow = 1 To vasID.DataRowCnt
'            If InStr(1, Trim(GetText(vasID, lRow, 6)), "수신완료") > 0 Then
'                lCnt = lCnt + 1
'            Else
'            End If
'        Next lRow
'        If lCnt < vasID.DataRowCnt Then
'            Integra800_Res_Req
'            Integra800_QCRes_Req
'        Else
'            Integra800_OrderID_Req
'        End If
''            Integra800_QCRes_Req
'    ElseIf Left(Host_BC, 2) = "60" Or Host_BC = "00" Then
'        For lRow = 1 To vasID.DataRowCnt
'            If InStr(1, Trim(GetText(vasID, lRow, 6)), "수신완료") > 0 Then
'            'If InStr(1, Trim(GetText(vasID, lRow, colState)), "수신완료") > 0 Then
'                lCnt = lCnt + 1
'            Else
'            End If
'        Next lRow
'        If lCnt < vasID.DataRowCnt Then
'            Integra800_Res_Req
'            Integra800_QCRes_Req
'
'        Else
'            Integra800_OrderID_Req
''            Integra800_Res_Req
''            Integra800_QCRes_Req
'        End If
''            Integra800_QCRes_Req
'    ElseIf Host_BC = "10" Then
'
'        If vasCode.DataRowCnt < 1 Then
'            Integra800_OrderID_Req
'        End If
''    Else
''        Integra800_OrderID_Req
''        Integra800_QCRes_Req
'    End If
'End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
'''    Dim lsID As String
'''    Dim i As Integer
'''
'''    Dim lsTempBarCode As String
'''    Dim lsPID As String
'''    Dim lsPname As String
'''    Dim lsSex As String
'''    Dim lsAge As String
'''
'''    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
'''
'''    ClearSpread vasRes
'''    vasRes.maxrows = 0
'''
'''    lsID = Trim(GetText(vasID, Row, colBarCode))
'''
'''
'''    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime " & vbCrLf & _
'''          "FROM pat_res " & vbCrLf & _
'''          "WHERE  " & vbCrLf & _
'''          "  equipno = '" & gEquip & "' " & vbCrLf & _
'''          "  AND Barcode = '" & Trim(GetText(vasID, Row, colBarCode)) & "' "
'''        SQL = SQL & vbCrLf & "AND diskno = '" & Trim(GetText(vasID, Row, colRack)) & "' "
'''        SQL = SQL & vbCrLf & "AND resdate = '" & Trim(GetText(vasID, Row, colExamDTTM)) & "' "
'''        SQL = SQL & vbCrLf & "  order by seqno, equipcode"
'''
'''
'''    res = db_select_Vas(gLocal, SQL, vasRes)
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Sub
'''    End If

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim sExamTime As String
    
    
    sExamDate = ""
'''    sExamDate = Trim(GetText(vasRes, asRow2, colResDate))
'''    sExamTime = Trim(GetText(vasRes, asRow2, colResTime))
    
    If Trim(sExamDate) = "" Then
        sExamDate = Format(Date, "yyyymmdd")
    End If
    
    
    SQL = "select examcode FROM pat_res " & vbCrLf & _
          "WHERE equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & Trim(GetText(vasID, asRow1, colEquipCode)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' and resdate = '" & Trim(GetText(vasID, asRow1, colExamDTTM)) & "'"
    SQL = SQL & vbCrLf & "AND diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "' and equipnum = '" & Trim(GetText(vasID, asRow1, colEquipNum)) & "' "
'''    SQL = SQL & vbCrLf & "AND posno = '" & Trim(GetText(vasID, asRow1, colExamDTTM)) & "' "
    res = db_select_Row(gLocal, SQL)
    
    If res > 0 Then
        SQL = "update pat_res set resvalue = '" & Trim(GetText(vasID, asRow1, colResult)) & "', " & vbCrLf & _
              "result = '" & Trim(GetText(vasID, asRow1, colResult)) & "', result_ifcc = '" & Trim(GetText(vasID, asRow1, colIFCC)) & "', result_eag = '" & Trim(GetText(vasID, asRow1, colEAG)) & "', " & vbCrLf & _
              "sendflag = '" & asSend & "', " & vbCrLf & _
              "examdate = '" & sExamDate & "', examtime = '" & sExamTime & "', " & vbCrLf & _
              "errremark = '" & Trim(GetText(vasID, asRow1, colRemark)) & "', refvalue = '" & Trim(GetText(vasID, asRow1, colArea)) & "' " & vbCrLf & _
              "WHERE equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND equipcode = '" & Trim(GetText(vasID, asRow1, colEquipCode)) & "'" & vbCrLf & _
              "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' and equipnum = '" & Trim(GetText(vasID, asRow1, colEquipNum)) & "' "
        SQL = SQL & vbCrLf & "AND diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "' "
        SQL = SQL & vbCrLf & "AND resdate = '" & Trim(GetText(vasID, asRow1, colExamDTTM)) & "' "
        res = SendQuery(gLocal, SQL)
        
    Else
        SQL = "insert into pat_res(examdate, equipno, barcode, equipcode, examcode, " & vbCrLf & _
              "refflag, sendflag, seqno, examname, resvalue, " & vbCrLf & _
              "result, examtime, pid, pname, diskno, resdate, result_ifcc, result_eag, errremark, equipnum, refvalue) " & vbCrLf & _
              "values('" & sExamDate & "', '" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colEquipCode)) & "', '" & Trim(GetText(vasID, asRow1, colExamCode)) & "', " & vbCrLf & _
              "'" & Mid(Trim(GetText(vasID, asRow1, colResFlag)), 1, 1) & "', '" & asSend & "', '', " & vbCrLf & _
              "'', '" & Trim(GetText(vasID, asRow1, colResult)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colResult)) & "', " & vbCrLf & _
              "'" & sExamTime & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colExamDTTM)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colIFCC)) & "', '" & Trim(GetText(vasID, asRow1, colEAG)) & "', '" & Trim(GetText(vasID, asRow1, colRemark)) & "', '" & Trim(GetText(vasID, asRow1, colEquipNum)) & "', '" & Trim(GetText(vasID, asRow1, colArea)) & "' ) "
        res = SendQuery(gLocal, SQL)
    End If
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

'''Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
'''    Dim strBarcode
'''    Dim strPBarcode As String
'''
'''    If Col = colBarCode Then
'''
'''        strBarcode = InputBox("새로운 검체번호를 입력하세요.", "검체번호변경")
'''
'''        If strBarcode = "" Then Exit Sub
'''
'''        strPBarcode = Trim(GetText(vasID, Row, colBarCode))
'''
'''        SetText vasID, strBarcode, Row, colBarCode
'''        Get_Sample_Info Row
'''
'''        SQL = "update pat_res set barcode = '" & strBarcode & "', " & vbCrLf & _
'''              "pid = '" & Trim(GetText(vasID, Row, colPID)) & "', pname = '" & Trim(GetText(vasID, Row, colPName)) & "' " & vbCrLf & _
'''              "where equipno = '" & gEquip & "' and barcode = '" & strPBarcode & "'"
'''        res = SendQuery(gLocal, SQL)
'''
'''    End If
'''
'''
'''End Sub

Private Sub vasID_KeyPress(KeyAscii As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long
    Dim i As Integer
    Dim ii As Integer

    If KeyAscii = 13 Then
    
'''        For i = 1 To vasID.DataRowCnt
'''            vasID.Row = i
'''            vasID.Col = 1
'''            vasID.Value = 0
'''
'''        Next
'''
        llRow = vasID.ActiveRow
        sSpecID = Trim(GetText(vasID, llRow, colBarCode))
        
        If IsNumeric(sSpecID) = False Then Exit Sub
        
        If Len(sSpecID) <= 5 Then
            sSpecID = Format(Date, "yyyymmdd") & Format(sSpecID, "00000")
        Else
            sSpecID = Mid(sSpecID, 1, 8) & Format(Mid(sSpecID, 9), "00000")
        End If
        SetText vasID, sSpecID, llRow, colBarCode
        
        
        
        '샘플의 환자 정보 가져오기
        Get_Sample_Info llRow
        
        SQL = "update pat_res set barcode = '" & sSpecID & "', " & vbCrLf & _
              "pid = '" & Trim(GetText(vasID, llRow, colPID)) & "', pname = '" & Trim(GetText(vasID, llRow, colPName)) & "' " & vbCrLf & _
              ", examcode = '" & Trim(GetText(vasID, llRow, colExamCode)) & "' " & vbCrLf & _
              "where equipno = '" & gEquip & "' and resdate = '" & Trim(GetText(vasID, llRow, colExamDTTM)) & "'"
        res = SendQuery(gLocal, SQL)
        
'''        vasID.Row = llRow
'''        vasID.Col = 1
'''        vasID.Value = 1
'''
'''        ii = llRow + 1
'''
'''        If ii > 0 And ii < vasID.DataRowCnt Then
'''            vasID.SetSelection colBarCode, ii, colBarCode, ii
'''
'''        End If
        
    End If
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim llRow As Integer
'''    If KeyCode = vbKeyUp Then
'''        llRow = vasID.ActiveRow - 1
'''        If llRow < 1 Then
'''            llRow = 1
'''        End If
'''
'''        vasID_Click colBarCode, llRow
'''    ElseIf KeyCode = vbKeyDown Then
'''        llRow = vasID.ActiveRow + 1
'''        If llRow < vasID.maxrows Then
'''            llRow = vasID.maxrows
'''        End If
'''
'''        vasID_Click colBarCode, llRow
'''    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If

'    PopupMenu mnuPop
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
'''    Dim lsID As String
'''    Dim i As Integer
'''
'''    Dim lsTempBarCode As String
'''    Dim lsPID As String
'''    Dim lsPname As String
'''    Dim lsSex As String
'''    Dim lsAge As String
'''
'''    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
'''
'''    ClearSpread vasListRes
'''    vasListRes.maxrows = 0
'''
'''    lsID = Trim(GetText(vasList, Row, colBarCode))
'''
'''
'''    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime " & vbCrLf & _
'''          "FROM pat_res " & vbCrLf & _
'''          "WHERE  " & vbCrLf & _
'''          "  equipno = '" & gEquip & "' " & vbCrLf & _
'''          "  AND Barcode = '" & Trim(GetText(vasList, Row, colBarCode)) & "' "
'''    SQL = SQL & vbCrLf & "AND diskno = '" & Trim(GetText(vasList, Row, colRack)) & "' "
'''    SQL = SQL & vbCrLf & "AND resdate = '" & Trim(GetText(vasList, Row, colExamDTTM)) & "' "
'''    SQL = SQL & vbCrLf & "  order by seqno, equipcode"
'''
'''    res = db_select_Vas(gLocal, SQL, vasListRes)
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Sub
'''    End If


End Sub

'''Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
'''    Dim strBarcode
'''    Dim strPBarcode As String
'''    Dim strTestTime As String
'''
'''    If Col = colBarCode Then
'''
'''        strBarcode = InputBox("새로운 검체번호를 입력하세요.", "검체번호변경")
'''
'''        If strBarcode = "" Then Exit Sub
'''
'''        strPBarcode = Trim(GetText(vasList, Row, colBarCode))
'''        strTestTime = Trim(GetText(vasList, Row, colExamDTTM))
'''
'''        SetText vasList, strBarcode, Row, colBarCode
'''        Get_Sample_Info2 Row
'''
'''        SQL = "update pat_res set barcode = '" & strBarcode & "', " & vbCrLf & _
'''              "pid = '" & Trim(GetText(vasList, Row, colPID)) & "', pname = '" & Trim(GetText(vasList, Row, colPName)) & "' " & vbCrLf & _
'''              "where equipno = '" & gEquip & "' and barcode = '" & strPBarcode & "' and resdate = '" & strTestTime & "'"
'''        res = SendQuery(gLocal, SQL)
'''
'''    End If
'''End Sub

Function Get_Sample_Info2(ByVal asRow As Long) As Integer
    Dim sID As String
    
    Dim lsPID As String
    Dim lsPname As String
    Dim lsDate As String
    
    '환자정보 가져오기
    sID = Trim(GetText(vasList, asRow, colBarCode))   '샘플 바코드 번호
    lsDate = Format(Date, "yyyymmdd")
    
    If sID = "" Then
        Exit Function
    End If
    
    '바코드, 병록번호, 환자명, 검체코드, 검체명
    SQL = "select a.spcid, a.patno, b.patname  from SLXWORKT a, appatbat b"
    SQL = SQL & vbCrLf & "where a.SPCID = '" & sID & "' and a.patno = b.patno"
    SQL = SQL & vbCrLf & "group by a.spcid, a.patno, b.patname"
'''    SQL = "SELECT A.SPCM_NO, A.PID , B.PT_NM , A.SPCM_CD , c.SPCM_ENM " & vbCrLf & _
'''          "FROM MS.MSLRCPT A " & vbCrLf & _
'''          "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
'''          "INNER JOIN HO.PCPPATIENT B ON A.PID = B.PID " & vbCrLf & _
'''          "INNER JOIN MS.MSLSPCMM C ON A.SPCM_CD = C.SPCM_CD " & vbCrLf & _
'''          "WHERE A.SPCM_NO = '" & sID & "' " & vbCrLf & _
'''          "AND AA.EXMN_CD IN (" & gAllExam & ") " & vbCrLf & _
'''          "GROUP BY A.SPCM_NO, A.PID, B.PT_NM, A.SPCM_CD, C.SPCM_ENM"
    res = db_select_Col(gServer, SQL)
    
    If res = 1 Then
        lsPID = Trim(gReadBuf(1))
        lsPname = Trim(gReadBuf(2))
        
        SetText vasList, lsPID, asRow, colPID
        SetText vasList, lsPname, asRow, colPName
    End If
    
End Function

Private Sub vasList_KeyPress(KeyAscii As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long
    Dim i As Integer
    Dim ii As Integer

    If KeyAscii = 13 Then
    
'''        For i = 1 To vaslist.DataRowCnt
'''            vaslist.Row = i
'''            vaslist.Col = 1
'''            vaslist.Value = 0
'''
'''        Next
'''
        llRow = vasList.ActiveRow
        sSpecID = Trim(GetText(vasList, llRow, colBarCode))
        
        If IsNumeric(sSpecID) = False Then Exit Sub
        
        If Len(sSpecID) <= 5 Then
            sSpecID = Format(Date, "yyyymmdd") & Format(sSpecID, "00000")
        Else
            sSpecID = Mid(sSpecID, 1, 8) & Format(Mid(sSpecID, 9), "00000")
        End If
        SetText vasList, sSpecID, llRow, colBarCode
        
        '샘플의 환자 정보 가져오기
        Get_Sample_Info_List llRow
        
        SQL = "update pat_res set barcode = '" & sSpecID & "', " & vbCrLf & _
              "pid = '" & Trim(GetText(vasList, llRow, colPID)) & "', pname = '" & Trim(GetText(vasList, llRow, colPName)) & "' " & vbCrLf & _
              ", examcode = '" & Trim(GetText(vasList, llRow, colExamCode)) & "' " & vbCrLf & _
              "where equipno = '" & gEquip & "' and resdate = '" & Trim(GetText(vasList, llRow, colExamDTTM)) & "'"
        res = SendQuery(gLocal, SQL)
        
'''        vaslist.Row = llRow
'''        vaslist.Col = 1
'''        vaslist.Value = 1
'''
'''        ii = llRow + 1
'''
'''        If ii > 0 And ii < vaslist.DataRowCnt Then
'''            vaslist.SetSelection colBarCode, ii, colBarCode, ii
'''
'''        End If
        
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

'''Private Sub vasSuga_Click(ByVal Col As Long, ByVal Row As Long)
'''    ClearSpread vasExamCnt
'''    SQL = "select kitcode, equipcode, count(equipcode) from pat_res " & vbCrLf & _
'''          "where equipno = '" & gEquip & "'" & vbCrLf & _
'''          "  and testdate between '" & Format(dtpSumSDate, "yyyymmdd") & "' and '" & Format(dtpSumEDate, "yyyymmdd") & "' " & vbCrLf & _
'''          "  and kitcode = '" & Trim(GetText(vasSuga, Row, 1)) & "' " & vbCrLf & _
'''          "group by kitcode, equipcode"
'''    res = db_select_Vas(gLocal, SQL, vasExamCnt)
'''
'''    vasExamCnt.MaxRows = vasExamCnt.DataRowCnt
'''
'''End Sub
'''
