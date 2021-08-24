VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " MicroScan Interface "
   ClientHeight    =   10680
   ClientLeft      =   330
   ClientTop       =   825
   ClientWidth     =   15585
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
   LockControls    =   -1  'True
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10680
   ScaleWidth      =   15585
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   2895
      Left            =   270
      TabIndex        =   25
      Top             =   3030
      Visible         =   0   'False
      Width           =   12645
      Begin VB.TextBox Text1 
         Height          =   2055
         Left            =   3540
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   450
         Width           =   5445
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   1650
         Left            =   0
         TabIndex        =   57
         Top             =   1230
         Width           =   12225
         _Version        =   393216
         _ExtentX        =   21564
         _ExtentY        =   2910
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2085
         Left            =   120
         TabIndex        =   63
         Top             =   3570
         Visible         =   0   'False
         Width           =   9465
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1545
            Left            =   1170
            TabIndex        =   64
            Top             =   240
            Width           =   8160
            _Version        =   393216
            _ExtentX        =   14393
            _ExtentY        =   2725
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
            MaxCols         =   9
            SpreadDesigner  =   "frmInterface.frx":44DE
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   1245
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   1065
            _Version        =   393216
            _ExtentX        =   1879
            _ExtentY        =   2196
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
            SpreadDesigner  =   "frmInterface.frx":5F86
         End
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1245
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   11955
         _Version        =   393216
         _ExtentX        =   21087
         _ExtentY        =   2196
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":61CD
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   45
         Top             =   315
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   1614
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
         SpreadDesigner  =   "frmInterface.frx":AE4A
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   43
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtData 
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
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   32
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   735
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   1320
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6705
         TabIndex        =   27
         Top             =   1350
         Visible         =   0   'False
         Width           =   1335
         Begin MSCommLib.MSComm MSComm1 
            Left            =   135
            Top             =   300
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   975
         Left            =   6780
         TabIndex        =   26
         Top             =   240
         Width           =   315
         _Version        =   393216
         _ExtentX        =   556
         _ExtentY        =   1720
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
         SpreadDesigner  =   "frmInterface.frx":B091
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3195
         TabIndex        =   33
         Top             =   180
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":B2D8
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   34
         Top             =   180
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":B51F
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   660
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Width           =   5970
         _Version        =   393216
         _ExtentX        =   10530
         _ExtentY        =   1164
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   5
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   5
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmInterface.frx":B766
         UserResize      =   0
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   37
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   36
         Top             =   1410
         Width           =   915
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   9315
      Left            =   75
      TabIndex        =   6
      Top             =   840
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "진행상태"
      TabPicture(0)   =   "frmInterface.frx":BCAC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":BCC8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8775
         Left            =   -74820
         TabIndex        =   16
         Top             =   360
         Width           =   14625
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   8460
            TabIndex        =   38
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   44
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   42
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label4 
               Caption         =   "환자명 :"
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
               Left            =   3150
               TabIndex        =   41
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   40
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "바코드번호 :"
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
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "EXCEL"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13050
            TabIndex        =   24
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "로컬결과조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3060
            TabIndex        =   23
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   300
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21364736
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   780
            TabIndex        =   19
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdRTrans 
            Caption         =   "결과수동전송"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5460
            TabIndex        =   18
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "화면초기화"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6900
            TabIndex        =   17
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   8055
            _Version        =   393216
            _ExtentX        =   14208
            _ExtentY        =   13785
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   13
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":BCE4
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7275
            Left            =   8460
            TabIndex        =   21
            Top             =   1260
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   12832
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   7
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":C787
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin FPSpread.vaSpread spdResult3 
            Height          =   7830
            Left            =   8490
            TabIndex        =   48
            Top             =   720
            Width           =   6030
            _Version        =   393216
            _ExtentX        =   10636
            _ExtentY        =   13811
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            ColsFrozen      =   6
            DisplayRowHeaders=   0   'False
            EditEnterAction =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   25
            ScrollBars      =   0
            SpreadDesigner  =   "frmInterface.frx":10551
            UserResize      =   2
         End
         Begin VB.TextBox txtToSeq 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   345
            Left            =   5190
            MaxLength       =   4
            TabIndex        =   61
            Text            =   "9999"
            Top             =   330
            Width           =   525
         End
         Begin VB.TextBox txtFrSeq 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   345
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   60
            Text            =   "0001"
            Top             =   330
            Width           =   525
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   59
            Top             =   4770
            Width           =   225
         End
         Begin VB.TextBox txtRemark 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   9450
            MaxLength       =   50
            TabIndex        =   58
            Text            =   "MRSA(resistant to all beta-lactams)"
            Top             =   8130
            Width           =   5040
         End
         Begin FPSpread.vaSpread vasResult 
            Height          =   3855
            Left            =   150
            TabIndex        =   56
            Top             =   4680
            Width           =   8235
            _Version        =   393216
            _ExtentX        =   14526
            _ExtentY        =   6800
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   15
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":10BC3
            UserResize      =   2
         End
         Begin VB.ComboBox cboSlip 
            Height          =   315
            Left            =   6300
            TabIndex        =   54
            Top             =   330
            Width           =   855
         End
         Begin VB.CommandButton cmdWorkList 
            Caption         =   " 조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7200
            TabIndex        =   51
            Top             =   330
            Width           =   1125
         End
         Begin VB.CommandButton cmdLogSend 
            Caption         =   "로그파일로전송"
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
            Left            =   9180
            TabIndex        =   49
            Top             =   270
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   6060
            TabIndex        =   12
            Top             =   4950
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   1680
            TabIndex        =   11
            Top             =   4800
            Visible         =   0   'False
            Width           =   4125
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "화면초기화"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13050
            TabIndex        =   15
            Top             =   270
            Width           =   1395
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "결과저장"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11550
            TabIndex        =   14
            Top             =   270
            Width           =   1395
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   10
            Top             =   780
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   3855
            Left            =   135
            TabIndex        =   13
            Top             =   720
            Width           =   8235
            _Version        =   393216
            _ExtentX        =   14526
            _ExtentY        =   6800
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
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
            MaxCols         =   14
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1172A
            UserResize      =   2
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1815
            Left            =   9045
            TabIndex        =   8
            Top             =   6060
            Visible         =   0   'False
            Width           =   5970
            Begin VB.TextBox txtErrLog 
               Appearance      =   0  '평면
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   9
               Top             =   240
               Width           =   5775
            End
         End
         Begin MSComCtl2.DTPicker dtpFrDt 
            Height          =   315
            Left            =   1140
            TabIndex        =   50
            Top             =   330
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21364737
            CurrentDate     =   40739
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   2850
            TabIndex        =   52
            Top             =   330
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21364737
            CurrentDate     =   40739
         End
         Begin VB.TextBox txtComm 
            Appearance      =   0  '평면
            Height          =   615
            Left            =   8490
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   7920
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.Label Label8 
            Caption         =   "Remark"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   8520
            TabIndex        =   68
            Top             =   8160
            Width           =   1065
         End
         Begin VB.Label Label7 
            Caption         =   "~"
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
            Left            =   2670
            TabIndex        =   66
            Top             =   390
            Width           =   195
         End
         Begin VB.Label Label6 
            Caption         =   "~"
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
            Left            =   4980
            TabIndex        =   62
            Top             =   390
            Width           =   195
         End
         Begin VB.Label Label5 
            Caption         =   "SLIP"
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
            Left            =   5760
            TabIndex        =   55
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "조회기간"
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
            Left            =   180
            TabIndex        =   53
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label lblSpecimen 
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
            Height          =   285
            Left            =   9180
            TabIndex        =   47
            Top             =   270
            Width           =   1965
         End
         Begin VB.Label lbl2 
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "검 체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   8490
            TabIndex        =   46
            Tag             =   "157"
            Top             =   330
            Visible         =   0   'False
            Width           =   585
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10305
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2011-12-01"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 5:00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Service Center (02)6205-1751"
            TextSave        =   "Service Center (02)6205-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   15045
      _Version        =   65536
      _ExtentX        =   26538
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     MicroScan INTERFACE"
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4785
         Picture         =   "frmInterface.frx":1225F
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   195
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   12120
         TabIndex        =   2
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364736
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   11190
         TabIndex        =   5
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5190
         TabIndex        =   4
         Top             =   255
         Width           =   1185
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "메인"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'vasid, vasrid colum
'Const colCheckBox = 1
'Const colBarcode = 2
'Const colRack = 3
'Const colPos = 4
'Const colPID = 5
'Const colPName = 6
'Const colSex = 7
'Const colAge = 8
'Const colJumin = 9
'Const colOCnt = 10
'Const colHospital = 11
'Const colState = 12


Const colCheckBox = 1
Const colSpecNo = 2
Const colBarcode = 3
'-- add
Const colTestCd = 4
'-- edit start
Const colRack = 5
Const colPos = 6
Const colPID = 7
Const colPName = 8
Const colSex = 9
Const colAge = 10
Const colOCnt = 11
Const colRCnt = 12
Const colState = 13
Const colA1c = 14
Const colCmt = 15
Const colIFCC = 16
Const coleAg = 18
'-- edit end

'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colFLAG = 6

Dim gRow As Long
Dim gsBarCode As String
Dim gsSampleType As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String
Dim gsFlag As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim gIFCC1 As String
Dim gIFCC2 As String
Dim geAg1 As String
Dim geAg2 As String
Dim gADD_IFCC As String
Dim gADD_eAg As String

Dim strBuffer As String

Public gENQFlag As Integer
Public gNAKFlag As Integer

'===============================
Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const rs  As String = ""
Const GS  As String = ""


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer

Dim OrgSort_Flag    As Integer
Dim RsltSort_Flag    As Integer

Private Sub chkAll_Click(Index As Integer)
    Dim iRow As Long
    
    If Index = 0 Then
        With vasID
            If .DataRowCnt = 0 Then Exit Sub
            .RowHeight(-1) = 12
            If chkAll(Index).Value = 1 Then
                For iRow = 1 To .DataRowCnt
                    .Row = iRow: .Col = 1
                    .Value = 1
                Next iRow
                .Col = 1: .Col2 = .MaxCols
                .Row = 1: .Row2 = .DataRowCnt
                .BlockMode = True
                .FontBold = True
                .BlockMode = False
            
            ElseIf chkAll(Index).Value = 0 Then
                For iRow = 1 To vasID.DataRowCnt
                    .Row = iRow: .Col = 1
                    .Value = 0
                Next iRow
                .Col = 1: .Col2 = .MaxCols
                .Row = 1: .Row2 = .DataRowCnt
                .BlockMode = True
                .FontBold = False
                .BlockMode = False
            End If
            .RowHeight(-1) = 12
            .SetFocus
        End With
    Else
        With vasResult
            If .DataRowCnt = 0 Then Exit Sub
            .RowHeight(-1) = 12
            If chkAll(Index).Value = 1 Then
                For iRow = 1 To .DataRowCnt
                    .Row = iRow: .Col = 1
                    .Value = 1
                Next iRow
                .Col = 1: .Col2 = .MaxCols
                .Row = 1: .Row2 = .DataRowCnt
                .BlockMode = True
                .FontBold = True
                .BlockMode = False
                
            ElseIf chkAll(Index).Value = 0 Then
                For iRow = 1 To .DataRowCnt
                    .Row = iRow: .Col = 1
                    .Value = 0
                Next iRow
                .Col = 1: .Col2 = .MaxCols
                .Row = 1: .Row2 = .DataRowCnt
                .BlockMode = True
                .FontBold = False
                .BlockMode = False
                
            End If
            .RowHeight(-1) = 12
            .SetFocus
        End With
    
    End If


End Sub

'Dim mOrder.NoOrder  As Boolean
'Dim mOrder.Order    As String
'Dim mOrder.IsSending As Boolean

'===============================

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub


Private Sub chkRAll_Click()
    Dim iRow As Long
    
    If chkRAll.Value = 1 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.Value = 1
        Next iRow
    ElseIf chkRAll.Value = 0 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmdExcel_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    
    
    
    ClearSpread vasPrint

    j = 1

    For iRow = 1 To vasRID.DataRowCnt
        vasRID.Row = iRow
        vasRID.Col = 1

        If vasRID.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasRID, iRow, colSpecNo)), j, 1
            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 2
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 3
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 4
            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 5
            'SetText vasPrint, Trim(GetText(vasRID, iRow, colHospital)), j, 5
            
            SQL = "SELECT RESULT " & vbCrLf & _
                  "FROM PAT_RES " & vbCrLf & _
                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
                  "ORDER BY SEQNO"
            res = db_select_Vas(gLocal, SQL, vasPrintBuf)
            
            sA1c = GetText(vasPrintBuf, 1, 1)
            sIFCC = GetText(vasPrintBuf, 2, 1)
            seAg = GetText(vasPrintBuf, 3, 1)

            ClearSpread vasPrintBuf, 1, 1

            SetText vasPrint, sA1c, j, 7
            SetText vasPrint, sIFCC, j, 8
            SetText vasPrint, seAg, j, 9
            
            '"GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, JUMIN, Hospital, SENDFLAG"
            
'            SetText vasprint, Trim(GetText(vasrid, iRow, vasrid.MaxCols)), j, 8
'            SetText vasprint, Trim(GetText(vasrid, iRow, 10)), j, 9
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        
    End If
End Sub
Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    With vasResult
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult2
        .MaxRows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With

    With spdResult3
        .MaxRows = 24
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now + 1
    
    txtRemark.Text = ""
    
    txtFrSeq.Text = "0001"
    txtToSeq.Text = "9999"
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasResult.DataRowCnt
        vasResult.Row = lRow
        vasResult.Col = 1
        If vasResult.Value = 1 Then
            res = Insert_Data_MIC(lRow)
        
            If res = -1 Then
                SetForeColor vasResult, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasResult, "Failed", lRow, colState
            Else
                vasResult.Row = lRow
                vasResult.Col = 1
                vasResult.Value = 1
                
                SetBackColor vasResult, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasResult, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasResult, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasResult.Row = lRow
            vasResult.Col = 1
            vasResult.Value = 0
        End If
    Next lRow
End Sub



Private Sub cmdRClear_Click()
    Dim i As Integer

'    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasRID, 1, vasRID.MaxRows, 1, vasRID.MaxCols, 0, 0, 0
    SetForeColor vasRRes, 1, vasRRes.MaxRows, 1, vasRRes.MaxCols, 0, 0, 0
    
    vasRID.MaxRows = 0
    vasRRes.MaxRows = 0
    
    dtpExamDate = Date
    
End Sub

Private Sub cmdRSch_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    Call chkRAll_Click
    
    SQL = "SELECT '', RECENO, BARCODE, EXAMCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND SENDFLAG IN ('1', '2') " & vbCrLf & _
          "GROUP BY BARCODE, RECENO, EXAMCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG"
    res = db_select_Vas(gLocal, SQL, vasRID)
    
          '"  AND SENDFLAG IN ('1', '2') "
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRID.DataRowCnt
        Select Case Trim(GetText(vasRID, iRow, colState))
        Case "2"
            SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasRID, "완료", iRow, colState
'        Case "0"
'            SetText vasID, "오더", iRow, colState
        Case "1"
            SetText vasRID, "결과", iRow, colState
        End Select
    Next iRow
End Sub

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            res = Insert_Data_MIC(lRow)
        
            If res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "Failed", lRow, colState
            ElseIf res = 0 Then
            
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasRID.Row = lRow
            vasRID.Col = 1
            vasRID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdWorkList_Click()
    
    Screen.MousePointer = 11
    
    Call GetWorkList(dtpFrDt.Value, dtpToDt.Value)

    Screen.MousePointer = 0
    
End Sub

Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim j As Integer
    Dim rs As ADODB.Recordset
    Dim sSpecNo As String
    Dim sWorkNo As String
    Dim buff As String
    
    buff = "0.7"
    
    vasID.MaxRows = 0
    
    '-- 검사대상자 가져오기
          SQL = "Select Distinct a.SPCM_NO, a.EXMN_CD, a.SPCM_CD "
    SQL = SQL & "  From SPSLHRRST a , SPSLMJBDI b"
'    SQL = SQL & " Where a.RCPN_DT between '" & Format(pFrDt, "yyyymmdd") & "' and '" & Format(pToDt, "yyyymmdd") & "'" & vbCrLf
    SQL = SQL & " Where a.RCPN_DT between TO_DATE(" & Format(pFrDt, "yyyymmdd") & ",'yyyymmdd') + 0.000000 "
    SQL = SQL & "    and TO_DATE(" & Format(pToDt, "yyyymmdd") & ",'yyyymmdd') + 0.999999 " & vbCrLf
    
'   AND B.RCPN_DT BETWEEN TO_DATE(20111108, 'yyyymmdd') + 0.000000
'                                     AND TO_DATE(20111108, 'yyyymmdd') + 0.999999
    
    SQL = SQL & "   AND SUBSTR(a.SPCM_NO,12,4) BETWEEN '" & Format(txtFrSeq.Text, "0000") & "' and '" & Format(txtToSeq.Text, "0000") & "'"
    SQL = SQL & "   AND a.RSLT_NO IS NOT NULL"
    SQL = SQL & "   AND a.RSLT_STAT <> '3' "
'    SQL = SQL & "   AND a.RSLT_STAT >= '1' "
    SQL = SQL & "   AND SUBSTR(a.WORK_NO,9,3) = '" & Trim(cboSlip.Text) & "' "
    SQL = SQL & "   AND a.SPCM_NO = b.SPCM_NO    "
    SQL = SQL & "   AND a.WORK_NO = b.WORK_NO    "
    SQL = SQL & "   AND a.EXMN_CD = b.EXMN_CD    "
    SQL = SQL & "   AND substr(a.EXMN_CD,1,3) <> 'L40'    "
    SQL = SQL & " Order By SPCM_NO, EXMN_CD "
    
    Set rs = cn_Ser.Execute(SQL, , 1)
          
    Do Until rs.EOF
        SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(rs.Fields(0)) & "') FROM DUAL "
        res = db_select_Col(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
              SQL = "SELECT WORK_NO"
        SQL = SQL & "  FROM SPSLMJBDI d, SPSLMJBBI b "
        SQL = SQL & " Where d.SPCM_NO = b.SPCM_NO"
        SQL = SQL & "   AND d.SPCM_NO = '" & Trim(rs.Fields(0)) & "'"
        SQL = SQL & "   AND b.slip_cd = '" & Trim(cboSlip.Text) & "'"
        res = db_select_Col(gServer, SQL)
        
        If Trim(gReadBuf(0)) <> "" Then
            sSpecNo = Trim(rs.Fields(0))
            sWorkNo = Trim(gReadBuf(0))
            sWorkNo = Mid(sWorkNo, 1, 11) & Mid(sWorkNo, 15, 4)
            
            SQL = "SELECT PID, PT_NM, SEX, AGE "
            SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
            SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & Trim(rs.Fields(0)) & "' "
            SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
            SQL = SQL & vbCrLf & "  AND RSLT_STAT <> '3' "
            res = db_select_Col(gServer, SQL)
            
            If Trim(gReadBuf(0)) <> "" Then
                j = j + 1
                vasID.MaxRows = j
                SetText vasID, sWorkNo, j, colSpecNo     '2
                SetText vasID, sSpecNo, j, colBarcode     '3
                SetText vasID, Trim(rs.Fields(1)), j, colTestCd    '4
                SetText vasID, Trim(gReadBuf(0)), j, colPID    '6
                SetText vasID, Trim(gReadBuf(1)), j, colPName  '7
                SetText vasID, Trim(gReadBuf(2)), j, colSex    '8
                SetText vasID, Trim(gReadBuf(3)), j, colAge    '9
                SetText vasID, Trim(rs.Fields(2)), j, 14    'SPCMCD 검체코드
            End If
        End If
        rs.MoveNext
    
    Loop
    
    vasID.RowHeight(-1) = 12
    
    'Call vasID_DblClick(2, 0)
    Call vasID_DblClick(2, 0)
    

End Sub

Private Sub GetWorkList_Result(ByVal strSpcmNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim j As Integer
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim sSpecNo As String
    Dim sWorkNo As String
    Dim buff As String
    Dim strBarNo As String
    Dim strWorkNo As String
    
    '-- 바코드번호로 SPCM_NO 찾아오기
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, 2)) = strSpcmNo Then
            intRow = i
            
            'strSpcmNo = Trim(GetText(vasID, i, 3))
            strSpcmNo = Trim(GetText(vasID, i, 2))
            Exit For
        End If
        
        If Trim(GetText(vasID, i, 2)) = "" Then
            intRow = i
            Exit For
        End If
    Next
    
    
    '-- 검사대상자 가져오기
'    SQL = "Select SPCM_NO, EXMN_CD, SPCM_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & strSpcmNo & "'" & _
          "   and EXMN_CD = '" & mResult.TestCd & "'" & _
          "   and rslt_no IS NOT NULL" & _
          "   and RSLT_STAT <> '3' "
    
    strWorkNo = Mid(strSpcmNo, 1, 11) & "00I" & Mid(strSpcmNo, 12, 4)
    '-- 검사대상자 가져오기
    SQL = "Select SPCM_NO, EXMN_CD, SPCM_CD From SPSLHRRST " & CR & _
          " Where WORK_NO = '" & strWorkNo & "'" & _
          "   and EXMN_CD = '" & mResult.TestCd & "'" & _
          "   and rslt_no IS NOT NULL" & _
          "   and RSLT_STAT <> '3' "
          
    Set rs = cn_Ser.Execute(SQL, , 1)
    
    With vasResult
        .MaxRows = .MaxRows + 1
    
        If Not rs.EOF Then
            Do Until rs.EOF
                SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(rs.Fields(0)) & "') FROM DUAL "
                res = db_select_Col(gServer, SQL)
                sSpecNo = Trim(gReadBuf(0))
    
                      SQL = "SELECT WORK_NO FROM SPSLMJBDI"
                SQL = SQL & " WHERE brcd_labl_no = '" & Trim(sSpecNo) & "'"
                res = db_select_Col(gServer, SQL)
                sWorkNo = Trim(gReadBuf(0))
                sWorkNo = Mid(sWorkNo, 1, 11) & Mid(sWorkNo, 15, 4)
    
                SQL = "SELECT PID, PT_NM, SEX, AGE "
                SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
                SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & Trim(rs.Fields(0)) & "' "
                SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
                SQL = SQL & vbCrLf & "  AND RSLT_STAT <> '3' "
                res = db_select_Col(gServer, SQL)
    
                SetText vasResult, sWorkNo, .MaxRows, 2  '2 검체번호
                'SetText vasResult, strSpcmNo, .MaxRows, 3             '3 바코드번호
                SetText vasResult, Trim(rs.Fields(0)), .MaxRows, 3             '3 바코드번호
                SetText vasResult, mResult.TestCd, .MaxRows, 4   '4 검사코드
                SetText vasResult, Trim(gReadBuf(0)), .MaxRows, 5   '5 환자번호
                SetText vasResult, Trim(gReadBuf(1)), .MaxRows, 6   '6 환자명
                SetText vasResult, mResult.MnmCd, .MaxRows, 7       '7 균코드
                SetText vasResult, mResult.MnmNm, .MaxRows, 8       '8 균명
                SetText vasResult, mResult.MCnt, .MaxRows, 9        '9 항생제수
                SetText vasResult, Trim(gReadBuf(2)), .MaxRows, 10  '10 성별
                SetText vasResult, Trim(gReadBuf(3)), .MaxRows, 11  '11 나이
                'SetText vasResult, Trim(rs.Fields(2)), .MaxRows, 15  '15 SPCMCD(검체코드)
    
                rs.MoveNext
            Loop
    
        Else
            '-- 검체번호 가져오기
                  SQL = "SELECT DISTINCT SPCM_NO  FROM SPSLHRRST"
'            SQL = SQL & "  Where  SUBSTR(WORK_NO,1,11) = '" & Mid(Trim(strSpcmNo), 1, 11) & "'"
'            SQL = SQL & "    AND SUBSTR(WORK_NO,15,4) = '" & Mid(Trim(strSpcmNo), 12, 4) & "'"
            SQL = SQL & "  Where  WORK_NO  = '" & strWorkNo & "'"
            
            'Set rs1 = New ADODB.Recordset
            Set rs1 = cn_Ser.Execute(SQL, , 1)
            Do Until rs1.EOF
                mResult.BarNo = Trim(rs1.Fields(0).Value) & ""
                rs1.MoveNext
            Loop
            
            
            SetText vasResult, strSpcmNo, .MaxRows, 2           '2 검체번호
            SetText vasResult, mResult.BarNo, .MaxRows, 3       '3 바코드번호
            SetText vasResult, mResult.TestCd, .MaxRows, 4      '4 검사코드
            SetText vasResult, mResult.PatNo, .MaxRows, 5       '5 환자번호
            SetText vasResult, "", .MaxRows, 6                  '6 환자명
            SetText vasResult, mResult.MnmCd, .MaxRows, 7       '7 균코드
            SetText vasResult, mResult.MnmNm, .MaxRows, 8       '8 균명
            SetText vasResult, mResult.MCnt, .MaxRows, 9        '9 항생제수
             
            '-- 임시 테스트 용
'            SetText vasResult, "20110831L4B0003" & vasResult.MaxRows, vasResult.MaxRows, 2     '2
'
'            SetText vasResult, "123456789" & vasResult.MaxRows, vasResult.MaxRows, colBarcode     '2
           ' SetText vasResult, "L41000", vasResult.MaxRows, colTestCd     '2
        
        End If
        
        .RowHeight(-1) = 12

    End With

End Sub

Private Sub cmdLOgSend_Click()
    
    Dim wkbuf As String
    
'    Open App.Path & "\log\long.log" For Input As #3
'    Open App.Path & "\log\multi.log" For Input As #3
'    Open App.Path & "\log\microscan.log" For Input As #3
    Open App.Path & "\log\1128.log" For Input As #3
    
    wkbuf = ""
    
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

    Close #3

    strBuffer = wkbuf
    
    Call MSComm1_OnComm
    
End Sub

Private Sub Label3_DblClick()

    If FrmHideControl.Visible = True Then
        FrmHideControl.Visible = False
    Else
        FrmHideControl.Visible = True
    End If

End Sub


Private Sub Label5_DblClick()

    If cmdLogSend.Visible = True Then
        cmdLogSend.Visible = False
    Else
        cmdLogSend.Visible = True
    End If
    
End Sub

Private Sub lblclear_Click()
    lblChangeBar.Caption = ""
    lblBarcode.Caption = ""
    lblChangePID.Caption = ""
    lblPname.Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
    
    
    
    
''            strBuffer = "1H|\^&|||ABL835^||||||||1|20070529193753" & vbCrLf
''strBuffer = strBuffer & "9C" & vbCrLf
'
'strBuffer = strBuffer & ""H","","","LabPro","3.01",",","""","N","","19",15,"Y","",""<CR><LF>
'P,"L","10363567","20110808L4B0026",,,,,,,"FM",,,,,,,,,N,,,,0,<CR><LF>
'B,"L","20110808L4B0026","10363567",,,"L41001","blood 1",N,20110811,,,,,,,F,0,Y,,,,,<CR><LF>
'F,"L",B,"20110808L4B0026::::20110811130114"<CR><LF>
'R,"L","01","20110808L4B0026","PBC28","Pos Breakpoint Combo 28",20110812,,N,,,"175","Staphylococcus hominis subsp. hominis","302064",P,,,,P,,,,,,N,,,,,28,,N,,2,2,,0,F,,<CR><LF>
'M,"1","AM","Ampicillin","4",N,,"BLAC",,,,,,,,,,,,,,,,,,N<CR><LF>
'M,"2","AUG","Amox/K Clav","<=4/2",N,,"S",,,,,,,,,,,,,,,,,,N<CR><LF>
'M,"26","TE","Tetracycline","<=4",N,,"S",,,,,,,,,,,,,,,,,,N<CR><LF>
'M,"27","TEI","Teicoplanin","<=4",N,,"S",,,,,,,,,,,,,,,,,,N<CR><LF>
'M,"L","VA","Vancomycin","<=0.5",N,,"S",,,,,,,,,,,,,,,,,,N<CR><LF>
'L,"L",Y,0<CR><LF>
'<EOT>
'
'Call MSComm1_OnComm
'
'    Exit Sub
    
    
'    For i = 1 To Len(txtTest)
'        lsChar = Mid(txtTest, i, 1)
'
'        Select Case lsChar
'        Case chrSTX
'            txtData.Text = lsChar
'
'        Case chrETX
'            SaveData "[RX]" & txtData.Text & lsChar
'
'            URISCAN_PRO txtData  '한 레코드 받으면 처리
'
'        Case Else
'            txtData.Text = txtData.Text & lsChar
'        End Select
'    Next i
'
'    txtTest = ""

End Sub

Private Sub URISCAN_PRO(asData As String)
    Dim MyVar As String
    Dim MyRet As String
          
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim iRow As Integer
    Dim jRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResult As String
    
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sDate As String
    
    Dim lsSeq As String
    Dim lsCnt As String
    
    If Trim(asData) = "" Then
        Exit Sub
    End If
    
    MyVar = Trim(asData)
         
    sDate = Format(dtpToday, "yyyymmdd")
    
    i = InStr(1, MyVar, "Date")
    If i > 0 Then
        sDate = Format(CDate(Trim(Mid(MyVar, i + 6, 20))), "yyyy-mm-dd hh:nn:ss")
    End If
    
    i = InStr(1, MyVar, "ID_NO")
    sSeqNo = CStr(CLng(Trim(Mid(MyVar, i + 6, 4))))

    sBarcode = CStr(Trim(Mid(MyVar, i + 11, 12)))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, 0)) = sSeqNo Then
            llRow = iRow
            Exit For
        End If
        
        If Trim(GetText(vasID, iRow, 0)) = "" Then
            llRow = iRow
            Exit For
        End If
    Next iRow

    If llRow = -1 Then
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow
        End If
    End If
    
    ClearSpread vasRes, 1, 1

    SetText vasID, sSeqNo, llRow, 0
    'SetText vasID, sExamDate, llRow, colDate
    'SetText vasID, sDate, llRow, colTime
    SetText vasID, sBarcode, llRow, colBarcode
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '샘플의 환자 정보 가져오기
    Get_Sample_Info llRow
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))

    
    lsSeq = ""
    lsCnt = ""
        
    
    '결과 잘라 넣기
    j = 0
    For j = 1 To vasRes.MaxRows
        sExamName = Trim(GetText(vasCode, j, 1))
        
        Select Case sExamName
        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"
            i = InStr(1, MyVar, Trim(sExamName))
            sResult = Trim(Mid(MyVar, i + 3, 8))

        Case "p.H"
            i = InStr(1, MyVar, "p.H")
            sResult = Trim(Mid(MyVar, i + 3, 14))

        Case "S.G"
            i = InStr(1, MyVar, "S.G")

            If Mid(MyVar, i) = "<=" Or Mid(MyVar, i) = ">=" Then
                sResult = Trim(Mid(MyVar, i + 3, 9))
            Else
                sResult = Trim(Mid(MyVar, i + 3, 12))
            End If
        End Select
        
        Select Case sResult
        Case "-"
            sResult = "Negatvie"
        End Select
        
        ClearSpread vasTemp
        
        SQL = "Select examcode, '', examname From EquipExam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  And EquipCode = '" & Trim(sExamName) & "'"
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        For jRow = 1 To vasTemp.DataRowCnt
            sExamCode = Trim(GetText(vasTemp, jRow, 1))
            sSeqNo = Trim(GetText(vasTemp, jRow, 2))
            sExamName = Trim(GetText(vasTemp, jRow, 3))
        
            SetText vasRes, Trim(sExamName), j, colEquipCode '장비코드
            SetText vasRes, sExamCode, j, colExamCode '검사코드
            SetText vasRes, sExamName, j, colExamName '검사명
            SetText vasRes, Trim(sResult), j, colResult   '검사결과
            SetText vasRes, sSeqNo, j, colSeq        '순번(서브코드)
            Trim (GetText(vasID, llRow, 0))
            Save_Local_One llRow, j, "1", CStr(Trim(sResult))
        Next jRow
    Next j
    gReadBuf(0) = ""
    
    '수신중========================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    

End Sub




Private Sub E411(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsSubCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    
    Dim lsresult_IFCC As String
    Dim lsresult_eAg As String
    
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String

    If asData = "" Then
        Exit Sub
    End If
    X = 0
    TablePtr = 1
    
'    For j = 1 To Len(asData)
'        If (Mid(asData, j, 1) = chrETX) Then
'            TablePtr = TablePtr + 1
'            ResultTbl(TablePtr) = " "
'        Else
'            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
'        End If
'    Next j
    
    Select Case Mid(asData, 2, 1)
    Case "H":       'Header Record
            Var_Clear
            gsSampleType = ""
            iCnt = 0
            
            For i = 1 To Len(asData)
                If Mid(asData, i, 1) = "|" Then
                    iCnt = iCnt + 1
    
                    Select Case iCnt
                        Case 11
                            gsSampleType = Mid(asData, i + 1, 1)
                        Case 13
                            gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                    End Select
                End If
            Next i
    Case "P":
    Case "O":
            gsBarCode = Trim$(mGetP(ResultTbl(1), 4, "|"))
            gsPosNo = ""
            gsRackNo = ""
            gsSeqNo = ""
            
            gRow = -1
            For i = 1 To vasID.DataRowCnt
                If gsBarCode <> "" Then
                    If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                        gRow = i
                        Exit For
                    End If
    '            ElseIf sSampleType = "Q" Then
    
                End If
            Next i
            
            If gRow < 0 Then
                gRow = vasID.DataRowCnt + 1
                If vasID.MaxRows < gRow Then
                    vasID.MaxRows = gRow
                End If
            End If
            
            SetText vasID, gsBarCode, gRow, colBarcode
            SetText vasID, gsRackNo, gRow, colRack
            SetText vasID, gsPosNo, gRow, colPos
            
            vasActiveCell vasID, gRow, colBarcode
            ClearSpread vasRes
            
            '샘플정보 가져오기
            If gsSampleType = "Q" Then
                SetText vasID, "QC", gRow, colPName
            Else
                If Trim(GetText(vasID, gRow, colPID)) = "" And gsBarCode <> "" And Mid(gsBarCode, 1, 1) <> "U" Then
                    Get_Sample_Info gRow
                End If
            End If
    Case "R":
            gOrderMessage = "R"
            
    
            lsTestID = Trim$(mGetP(ResultTbl(1), 3, "|"))    '장비코드
            lsTestID = Trim$(mGetP(lsTestID, 4, "^"))    '장비코드
            lsResult = Trim$(mGetP(ResultTbl(1), 4, "|"))            '결과
            
            If lsTestID = "" Then: Exit Sub
            
            ClearSpread vasTemp
    
            SQL = "Select examcode, examname, seqno From equipexam" & vbCrLf & _
                  "Where equipno = '" & gEquip & "' " & vbCrLf & _
                  "And equipcode = '" & lsTestID & "' " ' & vbCrLf & _
                  "and examcode in (" & gOrderExam & ") "
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 Then
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                lsSeqNo = Trim(gReadBuf(2))
                
                '숫자만 디스플레이 하기
                If IsNumeric(lsResult) = False Then
                    For ii = 1 To Len(lsResult)
                        If Mid(lsResult, ii, 1) = "?" Then
                            lsResult = Mid(lsResult, ii + 1)
                            
                            Exit For
                        End If
                    Next ii
                End If
                
                lsResRow = vasRes.DataRowCnt + 1
                If vasRes.MaxRows < lsResRow Then
                    vasRes.MaxRows = lsResRow
                End If
                
                '소수점 처리, 결과 형태 처리
                
                lsEquipRes = lsResult
                lsResult = SetResult(lsResult, lsTestID)
                lsResult_Buff = lsResult
                
                SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
                SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                SetText vasRes, lsResult, lsResRow, colResult            '결과
                
                SetText vasID, lsResult, gRow, colA1c                    '결과
                SetText vasID, gsFlag, gRow, colA1c + 1                  'Flag
                
                SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                SetText vasRes, gsFlag, lsResRow, 7                      'Flag
                
                
                Save_Local_One gRow, lsResRow, "1", CLng(lsEquipRes)
                            
                If IsNumeric(lsResult) = False Then
                    Exit Sub
                End If
    
                lsResult_Buff = ""
                    
            End If
    Case "L":
            gOrderExam = ""
            If MnTransAuto.Checked = True Then
                res = Insert_Data(gRow)
                
                If res = -1 Then
                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                    SetText vasID, "Failed", gRow, colState
                Else
                   
                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                    SetText vasID, "Trans", gRow, colState
                    
                    SQL = " Update pat_res Set " & vbCrLf & _
                          " sendflag = '2' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                End If
                
            End If
        
            SetText vasID, "Result", gRow, colState
    End Select
    
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click

    
    GetSetup
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If

'    -- osw 추가
    For i = 1 To 3
        If Not Connect_PRServer Then
            cn_cnt = cn_cnt + 1
            If cn_cnt = 3 Then
                If Not Connect_DRServer Then
                    MsgBox "연결되지 않았습니다."
                    cn_Server_Flag = False
                    Exit Sub
                Else
                    cn_Server_Flag = True
                End If
            End If
        Else
            cn_Server_Flag = True
        End If
    Next


    GetExamCode
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now + 1
    
    txtRemark.Text = ""
    
    txtFrSeq.Text = "0001"
    txtToSeq.Text = "9999"
    '==============================
    
'SLIP코드 조회
    cboSlip.Clear

    SQL = "SELECT slipcd " & CR & _
          "  From sliptable " & CR & _
          " order by seq "
    
    res = db_select_Row(gLocal, SQL)
'    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            cboSlip.AddItem Trim(gReadBuf(i)) & ""
        Else
            Exit For
        End If
    Next
    
    cboSlip.ListIndex = 0
    
End Sub



Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From equipexam " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasCode)
    If res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 6)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasCode.DataRowCnt
        If i = 1 Then
            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
        End If
        
        gArrEquip(i, 1) = i
        For j = 1 To 5
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
        
        
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1
    
End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '송신할 데이터
    Dim intCnt      As Integer
    Dim intOrdCnt   As Integer
    
    With vasID
        For intCnt = 1 To .DataRowCnt
            .Col = 1
            .Row = intCnt
            If .Value = "1" Then
                Select Case intSndPhase
                    Case 1  '## P
                        '<STX>P,"L","10881529","20110808L4B0039",,,,,,,"GIM",,,"MICU",,,,,,,,,,,<CR><LF><ETX>
                                    strOutput = STX & "P,"
                        strOutput = strOutput & """" & "L" & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 7) & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 2) & """,,,,,,,"
                        strOutput = strOutput & ",,,"
                        strOutput = strOutput & ",,,,,,,,,,,"
                        strOutput = strOutput & vbCr & vbLf & ETX
                        intSndPhase = 2
                        
                    Case 2  '## B
                        '<STX>B,"L","20110808L4B0039","10881529",,,"L41001",,,20110816,,,,,,,,,,"MICU",,,,<CR><LF><ETX>
                                    strOutput = STX & "B,"
                        strOutput = strOutput & """" & "L" & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 2) & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 7) & """,,,"
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 4) & """,,,"
                        strOutput = strOutput & Format(dtpToday, "yyyymmdd")
                        strOutput = strOutput & ",,,,,,,,,,"
                        strOutput = strOutput & ",,,,"
                        strOutput = strOutput & vbCr & vbLf & ETX
                        intSndPhase = 3
                        
                    Case 3  '## F
                        '<STX>F,"L","B","20110808L4B0039::::20110816105058:<CR><LF><ETX>
                                    strOutput = STX & "F,"
                        strOutput = strOutput & """" & "L" & ""","
                        strOutput = strOutput & """" & "B" & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 2) & "::::"
                        strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & ":"
                        strOutput = strOutput & vbCr & vbLf & ETX
                        intSndPhase = 4
                        
                    Case 4  '## R
                        '<STX>R,"L","1","20110808L4B0039","",,20110816,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,<CR><LF><ETX>
                                    strOutput = STX & "R,"
                        strOutput = strOutput & """" & "L" & ""","
                        strOutput = strOutput & """" & "1" & ""","
                        strOutput = strOutput & """" & GetText(vasID, intCnt, 2) & ""","
                        strOutput = strOutput & """" & """,,"
                        strOutput = strOutput & Format(dtpToday, "yyyymmdd")
                        strOutput = strOutput & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
                        strOutput = strOutput & vbCr & vbLf & ETX
                        intSndPhase = 5
                        
                    Case 5  '## L
                        '<STX>L,"L",N,<CR><LF><ETX>
                                    strOutput = STX & "L,"
                        strOutput = strOutput & """" & "L" & ""","
                        strOutput = strOutput & "N,"
                        strOutput = strOutput & vbCr & vbLf & ETX
                        
                        .Col = 1
                        .Row = intCnt
                        .Value = "0"
                        
                        SetBackColor vasID, intCnt, intCnt, 1, colState, 234, 255, 154
                        SetText vasID, "Send", intCnt, colState
                        
                        DoEvents

''                        For intOrdCnt = 1 To .DataRowCnt
''                            .Col = 1
''                            .Row = intOrdCnt
''                            If .Value = "1" Then
''                                intSndPhase = 1
''                                Exit For
'''                            Else
'''                                intSndPhase = 6
''                            End If
''                            intSndPhase = 6
''                        Next
                                                
                        If intCnt = .DataRowCnt Then
                            intSndPhase = 6
                        Else
                            intSndPhase = 1
                        End If
                                                
                                                
'                    Case 6  '## EOT
'
'                        intSndPhase = 1
'
'                        MSComm1.Output = EOT
'                        Save_Raw_Data "[Tx]" & EOT
'                        Debug.Print EOT
'
'                        Exit Sub
                End Select
                
                Exit For
            
            End If
        Next
    End With
    
    
    If intSndPhase = 6 Then
        'Call Sleep(500)
        intSndPhase = 1
        MSComm1.Output = EOT
        Debug.Print EOT
        Save_Raw_Data "[Tx]" & EOT
    End If
    
    MSComm1.Output = strOutput
    Debug.Print strOutput
    Save_Raw_Data "[Tx]" & strOutput
    
    

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-- 지금날짜와 검사일자 비교한다
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function


Private Sub MSComm1_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    Dim strDate As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
    dtpToday.Value = Format(strDate, "####-##-##")
    DoEvents
    
    If strBuffer <> "" And cmdLogSend.Visible = True Then
        Buffer = strBuffer
        strBuffer = ""
        GoTo Rst
    End If
    
    Select Case MSComm1.CommEvent
        Case comEvReceive
            Screen.MousePointer = 11
            strBuffer = ""
            Buffer = MSComm1.Input
Rst:
            Save_Raw_Data "[Rx]" & Buffer
            'txtComm.Text = txtComm.Text & Buffer
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

            'If BufChar = "" Then Stop
                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                intPhase = 2
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case ACK
                                If strState = "Q" Then Call SendOrder
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case STX
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                            Case ETB
                                blnIsETB = True
                                intPhase = 3
                            Case ETX
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 3
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case vbCr, vbLf
                            Case Else
                                If blnIsETB = False Then
                                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                                Else
                                    blnIsETB = False
                                End If
                        End Select
                    Case 3      '## Transfer Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
'                            Case vbCr
'                                intPhase = 4
'                                MSComm1.Output = ACK
'                                Save_Raw_Data "[Tx]" & ACK
'                            Case vbLf
'                                intPhase = 4
'                                MSComm1.Output = ACK
'                                Save_Raw_Data "[Tx]" & intPhase & ACK
                            Case vbCr, vbLf
                            Case ETX
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
'                                intPhase = 3
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            
                            Case EOT
                                Call EditRcvData
                                If strState = "Q" Then
                                    intSndPhase = 1
                                    intFrameNo = 1
                                    intPhase = 1
                                    MSComm1.Output = ENQ
                                    Save_Raw_Data "[Tx]" & ENQ
                                End If
                                intPhase = 1
                        
                        End Select
'                    Case 4      '## Termination Phase
'                        Select Case BufChar
'                            Case STX
'                                intPhase = 2
'                            Case EOT
'                                Call EditRcvData
'                                If strState = "Q" Then
'                                    intSndPhase = 1
'                                    intFrameNo = 1
'                                    'MSComm1.Output = ENQ
''                                    Save_Raw_Data "[Tx]" & ENQ
'                                End If
'                                intPhase = 1
'                        End Select
                End Select
            Next i
            Screen.MousePointer = 0
            vasResult.SetFocus

        Case comEvSend
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select


End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mOrder.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mOrder.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    strItems = GetEquipExamCode_E411(gEquip, pBarNo)

    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = ""
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    

End Sub

'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    
    Call ClearSpread(vasRes)
    Call ClearSpread(spdResult2)
    Call ClearSpread(spdResult3)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    gRow = intRow
    
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)

End Sub

Private Sub SetPatInfo_SPCMNO(ByVal pSpcmNo As String, ByVal pMnmCd As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasResult.DataRowCnt
        If Trim(GetText(vasResult, i, colSpecNo)) = pSpcmNo And Trim(GetText(vasResult, i, 7)) = pMnmCd Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasResult.MaxRows < intRow Then
            vasResult.MaxRows = intRow
        End If
    End If
    
    
'    Call SetText(vasResult, pSpcmNo, intRow, colSpecNo)  '3
'    Call SetText(vasResult, mResult.RackNo, intRow, colRack)       '4
'    Call SetText(vasResult, mResult.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasResult, intRow, colSpecNo)
    
    Call ClearSpread(vasRes)
    Call ClearSpread(spdResult2)
    Call ClearSpread(spdResult3)
    'Call Get_Sample_Info_SPCMNO(intRow)                        '2,6,7,8,9
    
    gRow = intRow
    
    gOrderExam = GetOrderExamCode_MIC(gEquip, pSpcmNo)

End Sub
'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetMnmInfo(ByVal pBarNo As String)
                

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strBarNo As String
    
    intRow = -1
    For i = 1 To spdResult2.DataRowCnt
        If Trim(GetText(spdResult2, i, colSpecNo)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = spdResult2.DataRowCnt + 1
        If spdResult2.MaxRows < intRow Then
            spdResult2.MaxRows = intRow
        End If
    End If
    
    
    strItems = Trim(GetText(frmInterface.vasResult, i, colTestCd))
    '-- 임시 테스트용
'    strItems = "L41000"
    If strItems = "" Then
        Exit Sub
    End If
    '바코드번호로 검체번호 불러오기FN_LABCVTPRTBCNO(SPCM_NO) --> 바코드라벨번호 리턴

    SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(pBarNo) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    strBarNo = Trim(gReadBuf(0))
    
    intRow = 1
    
    Call SetText(spdResult2, pBarNo, intRow, 1)
    Call SetText(spdResult2, strBarNo, intRow, 2)
    Call SetText(spdResult2, mResult.MnmCd, intRow, 3)
    Call SetText(spdResult2, mResult.MnmNm, intRow, 4)
    Call SetText(spdResult2, mResult.MCnt, intRow, 5)
    
    Call ClearSpread(spdResult3)

End Sub
'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetDrugInfo(ByVal pBarNo As String, ByVal strMachDrug As String, ByVal strDrug As String, _
                        ByVal strSensi As String, ByVal strVol As String)
                

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strCol1, strCol2, strCol3, strCol4 As String
    
'    intRow = -1
'    For i = 1 To spdResult2.DataRowCnt
'        If Trim(GetText(spdResult2, i, colBarcode)) = pBarNo Then
'            intRow = i
'            Exit For
'        End If
'    Next i
'
'    If intRow < 0 Then
'        intRow = spdResult3.DataRowCnt + 1
''        If spdResult3.MaxRows < intRow Then
''            spdResult3.MaxRows = intRow
''        End If
'    End If
    
    With spdResult3
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If Trim(.Text) = "" Then
                strCol1 = 1
                strCol2 = 2
                strCol3 = 3
                strCol4 = 4
                
                intRow = i
                Exit For
            End If
        Next
    
        If intRow = 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = 6
                If Trim(.Text) = "" Then
                    strCol1 = 6
                    strCol2 = 7
                    strCol3 = 8
                    strCol4 = 9
                
                    intRow = i
                    Exit For
                End If
            Next
        End If
    End With
    
    Call SetText(spdResult3, strMachDrug, intRow, strCol1)
    Call SetText(spdResult3, strDrug, intRow, strCol2)
    Call SetText(spdResult3, strVol, intRow, strCol3)
    Call SetText(spdResult3, strSensi, intRow, strCol4)
    
    If strSensi = "R" Then
        spdResult3.Row = intRow
        spdResult3.Col = strCol4
        spdResult3.ForeColor = vbRed
        spdResult3.FontBold = True
    Else
        spdResult3.Row = intRow
        spdResult3.Col = strCol4
        spdResult3.ForeColor = vbBlack
        spdResult3.FontBold = False
    End If
    spdResult3.RowHeight(-1) = 12
End Sub

Private Function GetDrug(ByVal pDrug As String) As String
    Dim Svr_Rs As ADODB.Recordset
    Dim strSQL As String
    
             strSQL = "select ANTB_ABBR_NM from SPSLMFMAT"
    strSQL = strSQL & " where ANTB_CD = '" & pDrug & "' "   '항생제코드:구분코드
'    strSQL = strSQL & "   and USE_STR_DT = '"
                
    
    Set Svr_Rs = cn_Ser.Execute(strSQL, , adCmdText)
    
    If Svr_Rs.EOF Then
        GetDrug = pDrug
    Else
        GetDrug = Svr_Rs.Fields("ANTB_ABBR_NM").Value & ""
    End If
    
    Set Svr_Rs = Nothing
End Function


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackno    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strIntBase2  As String   '수신한 장비기준 검사명(두뱐째 채널)
    Dim strResult    As String   '수신한 결과
    Dim strResult1   As String   '수신한 결과
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    
    Dim aryRcvBuf() As String
    Dim strWorkNo As String '작업번호 (접수일자 & 접수Seq)
    Dim strSpecNo     As String   'Specimen no
    Dim strPatNo     As String
    Dim strTestCd    As String
    Dim strMachMnmcd As String  '장비균명 코드
    Dim strMnmcd As String  '균명 코드
    Dim strMnmNm As String  '균명
    Dim strMCnt As String
    Dim strESBLVal As String  'ESBL 판정값
    Dim strScnt As String   '항생제 결과 수
    Static strRcvBufs As String
    Dim blnRst As Boolean
    Dim blnRst1 As Boolean
    Dim lngRstCnt As Long
    Dim i, j As Integer
    Dim strSndData  As String
    Dim rs_mic As ADODB.Recordset
    Dim strOrgTestCd As String
    
    strOrgTestCd = ""
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strRcvBuf = ReplaceVal(strRcvBuf)
        strType = Mid$(strRcvBuf, 1, 1)
        
        Select Case strType
            Case "H"    'Site/Header Record
                strRcvBufs = ""
            Case "P"    'Patient Record
                strBarNo = ReplaceVal(mGetP(strRcvBuf, 3, ",")) 'barcode No
                strWorkNo = ReplaceVal(mGetP(strRcvBuf, 4, ",")) 'WorkNo
            
            Case "A"    '오더요청 ==> 체크된 항목만 전송한다.
                strSpecNo = ReplaceVal(mGetP(strRcvBuf, 3, ",")) 'Specimen No
                With mOrder
                    .NoOrder = False
                End With
                
                strState = "Q"
                
            Case "B"    'Battery/Specimen Record
                strSpecNo = ReplaceVal(mGetP(strRcvBuf, 3, ","))    'Specimen No
                strPatNo = ReplaceVal(mGetP(strRcvBuf, 4, ","))
                strTestCd = ReplaceVal(mGetP(strRcvBuf, 7, ","))
                
                With mResult
                    .SpcmNo = strSpecNo     'Specimen No
                    .PatNo = strPatNo       'PatNo
                    .TestCd = strTestCd     'TestCd
                End With
            
            Case "R"    '## Result/Isolate Record
                strWorkNo = ReplaceVal(mGetP(strRcvBuf, 4, ","))    'WorkNo
                strMnmcd = ReplaceVal(mGetP(strRcvBuf, 12, ","))    '균명 코드
                strMnmNm = ReplaceVal(mGetP(strRcvBuf, 13, ","))    '균명
                strMCnt = ReplaceVal(mGetP(strRcvBuf, 30, ","))     '항생제 수
                strESBLVal = ReplaceVal(mGetP(strRcvBuf, 28, ","))  'ESBL 결과
                
                If strMnmcd <> "" Then
                    Set rs_mic = New ADODB.Recordset

                          SQL = "SELECT horgcd From orgtable "
                    SQL = SQL & " WHERE morgcd = '" & strMnmcd & "' "
                    Set rs_mic = cn.Execute(SQL)
                    Do Until rs_mic.EOF
                        strMnmcd = rs_mic.Fields(0).Value & ""
                        'mResult.MnmCd = strMnmcd
                        rs_mic.MoveNext
                    Loop

                    Set rs_mic = Nothing

                    If strMnmcd <> "" Then
                        Set rs_mic = New ADODB.Recordset
    
                              SQL = "SELECT DISTINCT bctr_cd From SPSLMFMBA "
                        SQL = SQL & " WHERE bctr_cd = '" & strMnmcd & "' "
                        SQL = SQL & " Union all "
                        SQL = SQL & "SELECT DISTINCT bctr_cd From SPSLMFMBA "
                        SQL = SQL & " WHERE bctr_itcn_cd = '" & strMnmcd & "' "
                        Set rs_mic = cn_Ser.Execute(SQL)
                        Do Until rs_mic.EOF
                            strMnmcd = rs_mic.Fields(0).Value & ""
                            mResult.MnmCd = strMnmcd
                            rs_mic.MoveNext
                        Loop
    
                        Set rs_mic = Nothing
                    End If
                    
                    With mResult
                        .MnmCd = strMnmcd
                        .MnmNm = strMnmNm
                        .MCnt = strMCnt '항생제 수
                    End With
                    
                    Call GetWorkList_Result(strWorkNo)
                    Call SetPatInfo_SPCMNO(strWorkNo, strMnmcd)
                    
                    If InStr(strWorkNo, "L4B") > 0 Then
                        strOrgTestCd = mResult.TestCd
                        mResult.TestCd = "L4100101"
                        Call GetWorkList_Result(strWorkNo)
                        Call SetPatInfo_SPCMNO(strWorkNo, strMnmcd)
                    
                        mResult.TestCd = "L4100102"
                        Call GetWorkList_Result(strWorkNo)
                        Call SetPatInfo_SPCMNO(strWorkNo, strMnmcd)
                        mResult.TestCd = strOrgTestCd
                    End If
                End If
                lblSpecimen.Caption = ""
                
            Case "M"    'MIC/Therapy/Dosage Record
                Dim strDrug As String
                Dim strMachDrug As String
                Dim strSensi As String
                Dim strVol As String
                Dim blnESBL As Boolean
                
                strState = "M"
                
                strDrug = Trim(ReplaceVal(mGetP(strRcvBuf, 3, ",")))    '-- 항생제코드
                strMachDrug = strDrug                                   '-- 항생제코드[장비]
                
                strIntBase = ""
                
                Set rs_mic = New ADODB.Recordset
    
                      SQL = "SELECT hanticd From antitable "
                SQL = SQL & " WHERE manticd = '" & strDrug & "' "
                Set rs_mic = cn.Execute(SQL)
                Do Until rs_mic.EOF
                    strIntBase = rs_mic.Fields(0).Value & ""
                    rs_mic.MoveNext
                Loop

                Set rs_mic = Nothing
                
                If strIntBase <> "" Then
                    Set rs_mic = New ADODB.Recordset

                          SQL = "SELECT DISTINCT antb_cd From SPSLMFMAT "
                    SQL = SQL & " WHERE antb_cd = '" & strIntBase & "' "
                    SQL = SQL & " Union all "
                    SQL = SQL & "SELECT DISTINCT antb_cd From SPSLMFMAT "
                    SQL = SQL & " WHERE antb_itcn_cd = '" & strIntBase & "' "
                    
                    Set rs_mic = cn_Ser.Execute(SQL)
                    Do Until rs_mic.EOF
                        strIntBase = rs_mic.Fields(0).Value & ""
                        rs_mic.MoveNext
                    Loop
                    Set rs_mic = Nothing
                Else
                    If strSensi <> "" Then
                        strIntBase = "----"
                        Call SetDrugInfo(strBarNo, strMachDrug, strIntBase, strSensi, strVol)
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        
                        SetText vasRes, strMachDrug, lsResRow, 15                '
    
                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                        
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c          '결과
                            SetText vasRes, "L4100101", lsResRow, colExamCode       '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c          '결과
                            SetText vasRes, "L4100102", lsResRow, colExamCode       '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                        
                        lsResult_Buff = ""
                        strIntBase = ""
                    End If
                End If
                
                strSensi = Trim(ReplaceVal(mGetP(strRcvBuf, 8, ",")))           '-- 감수성결과[R,R*,S...]
                
                If UCase(strSensi) <> "ESBL" And Len(strSensi) = 2 And InStr(strSensi, "*") > 0 Then
                    strSensi = Mid(strSensi, 1, 1)
                End If
                
                If strIntBase <> "" And Len(strIntBase) <= 5 And strSensi <> "" And strSensi <> "N/R" Then
                    strDrug = strIntBase
                    
                    If UCase(Trim(strSensi)) = "BLAC" Then
                        strSensi = "R"
                    End If
                    strResult = strSensi
                    strVol = Trim(ReplaceVal(mGetP(strRcvBuf, 5, ",")))         '-- 투여량[<=8/4..]

                    strComm = ""
                    
                    If UCase(strSensi) = "ESBL" Then
                        blnESBL = True
                    End If
                    
                    '=== ESBL ================================================================================
                    If (UCase(strSensi) = "ESBL" Or UCase(strSensi) = "R*") And UCase(strESBLVal) = "POS" Then
                        strComm = "ESBN"
                        strSensi = "R"
                        strResult = strSensi
                        
                        Call SetDrugInfo(strBarNo, strMachDrug, strDrug, strSensi, strVol)
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        SetText vasRes, strMachDrug, lsResRow, 15                '
    
                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                        
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                            SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                            SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                    
                        strMachDrug = "ESBL"
                        strDrug = "esb"
                        strIntBase = "esb"
                    End If
                    '=== ESBL ================================================================================
                    
                    
                    '=== MRSA ================================================================================
                    '-- 세균이 '150' => 'staaur' => 'sau' 이고
                    '-- 항생제가 'oxasillin' => 'ox' => 'oxs' => 'oxa' 이면서 결과값이 'R'이면 리마크 값을 넣는다.
                    If UCase(mResult.MnmCd) = "SAU" And UCase(strDrug) = "OXA" And strSensi = "R" Then
                        
                        Call SetDrugInfo(strBarNo, strMachDrug, strDrug, strSensi, strVol)
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        SetText vasRes, strMachDrug, lsResRow, 15                '
    
                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                        
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                            SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                            SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                        
                        strMachDrug = "MRSA"
                        strComm = "MRSA"
                        strDrug = "mrs"
                        strIntBase = "mrs"
                    End If
                    '=== MRSA ================================================================================
                    
                    '=== VREN ================================================================================
                    '-- 세균이 '125' => 'stafae' => 'efa' 이고
                    '-- 항생제가 'Vancomycin' => 'VA' => 'va' => 'van' 이면서 결과값이 'R'이면 리마크 값을 넣는다.
                    If mResult.MnmCd = "sau" And UCase(strDrug) = "VAN" And strSensi = "R" Then
                        Call SetDrugInfo(strBarNo, strMachDrug, strDrug, strSensi, strVol)
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        SetText vasRes, strMachDrug, lsResRow, 15                '
    
                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                        
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                            SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                            SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                        
                        strMachDrug = "VREN"
                        strComm = "VREN"
                        strDrug = "vre"
                        strIntBase = "vre"
                    End If
                    '=== VREN ================================================================================
                    
                    
                    '=== NORMAL ==============================================================================
                    Call SetDrugInfo(strBarNo, strMachDrug, strDrug, strSensi, strVol)
                    
                    lsResRow = vasRes.DataRowCnt + 1
                    If vasRes.MaxRows < lsResRow Then
                        vasRes.MaxRows = lsResRow
                    End If

                    '소수점 처리, 결과 형태 처리
                    lsEquipRes = strResult
                    lsResult_Buff = strResult

                    SetText vasResult, strResult, gRow, colA1c                   '결과
                    SetText vasResult, strComm, gRow, colCmt                  'Flag
                    
                    gOrderExam = Replace(gOrderExam, "'", "")
                    SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                    SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                    SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                    SetText vasRes, strResult, lsResRow, colResult           '결과
                    SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                    SetText vasRes, strResult1, lsResRow, 7                  'Flag
                    SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                    SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                    SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                    SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                    SetText vasRes, strVol, lsResRow, 12                     'antsize
                    SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                    SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                    SetText vasRes, strMachDrug, lsResRow, 15                '

                    Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                    
                    If InStr(strWorkNo, "L4B") > 0 Then
                        SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                        SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                        SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                        Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                        
                        SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                        SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                        SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                        Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                    End If
                    '=== NORMAL ==============================================================================
                    
                    '=== AM/AMP ==============================================================================
                    If UCase(strDrug) = "AM" Then
                        strDrug = "amp"
                        strIntBase = "amp"
                        
                        Call SetDrugInfo(strBarNo, strMachDrug, strDrug, strSensi, strVol)
                    
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        
                        SetText vasRes, strMachDrug, lsResRow, 15                '
    
                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                    
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                            SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                            SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                    End If
                    '=== AM/AMP ==============================================================================

                    lsResult_Buff = ""
                End If
                

            Case "L"    'End Of Block Record
                If strState = "M" Then
                
                    '=== ESBL/Pos ============================================================================
                    If blnESBL = True And UCase(strESBLVal) = "POS" Then
                        strDrug = "ESBL"
                        strSensi = "Pos"
                        strVol = ""
                        strComm = ""
                        
                        Call SetDrugInfo(strBarNo, strDrug, strDrug, strSensi, strVol)
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        strResult = strSensi
                        strIntBase = strDrug
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        lsResult_Buff = strResult
    
                        SetText vasResult, strResult, gRow, colA1c                   '결과
                        SetText vasResult, strComm, gRow, colA1c + 1                  'Flag
                        
                        gOrderExam = Replace(gOrderExam, "'", "")
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, mResult.TestCd, lsResRow, colExamCode    '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strResult1, lsResRow, 7                  'Flag
                        SetText vasRes, mResult.MnmCd, lsResRow, 8               'isocd
                        SetText vasRes, mResult.MnmNm, lsResRow, 9               'isonm
                        SetText vasRes, mResult.MCnt, lsResRow, 10               'antcnt
                        SetText vasRes, strIntBase, lsResRow, 11                 'antcd
                        SetText vasRes, strVol, lsResRow, 12                     'antsize
                        SetText vasRes, strSensi, lsResRow, 13                   'antrslt
                        SetText vasRes, GetText(vasResult, gRow, 15), lsResRow, 14                '
                        SetText vasRes, strIntBase, lsResRow, 15                '

                        Call Save_Local_One(gRow, lsResRow, "1", lsEquipRes)
                        
                        If InStr(strWorkNo, "L4B") > 0 Then
                            SetText vasResult, strResult, gRow + 1, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 1, colCmt                'Flag
                            SetText vasRes, "L4100101", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 1, lsResRow, "1", lsEquipRes)
                            
                            SetText vasResult, strResult, gRow + 2, colA1c                 '결과
                            SetText vasResult, strComm, gRow + 2, colCmt                'Flag
                            SetText vasRes, "L4100102", lsResRow, colExamCode    '검사코드
                            Call Save_Local_One(gRow + 2, lsResRow, "1", lsEquipRes)
                        End If
                        
                        lsResult_Buff = ""
                    End If
                    '=== ESBL/Pos ============================================================================
                    
                    If MnTransAuto.Checked = True Then
                        res = Insert_Data_MIC(gRow)

                        If res = -1 Then
                            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                            SetText vasID, "Failed", gRow, colState
                        Else
                            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                            SetText vasID, "Trans", gRow, colState

                            SQL = " Update pat_res Set " & vbCrLf & _
                                  " sendflag = '2' " & vbCrLf & _
                                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                                  " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If

                        End If

                    End If

                    SetText vasID, "Result", gRow, colState
                    strState = ""
                End If

        End Select
    Next

End Sub


Sub VARIANTII(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsSubCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    
    Dim lsresult_IFCC As String
    Dim lsresult_eAg As String
    
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String

    If asData = "" Then
        Exit Sub
    End If
    X = 0
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        gsSampleType = ""
        iCnt = 0
        
        For i = 1 To Len(asData)
            If Mid(asData, i, 1) = "|" Then
                iCnt = iCnt + 1

                Select Case iCnt
                    Case 11
                        gsSampleType = Mid(asData, i + 1, 1)
                    Case 13
                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                End Select
            End If
        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "O" Then
        If gsSampleType <> "P" Then: Exit Sub '/////QC데이터 안나와도 됨
        
        
        
        sTmp = Trim(ResultTbl(3))      'Barcode, Rack, Pos
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            If gsSampleType = "P" Then
                    If InStr(1, sTmp, "^") > 0 Then
                        iCnt = InStr(1, sTmp, "^")
                            gsBarCode = Trim(Mid(sTmp, 1, iCnt - 1))    'Barcode
                            If IsNumeric(gsBarCode) = True And Len(gsBarCode) > 12 Then
                                gsBarCode = Trim(Mid(gsBarCode, 1, 12))
                            End If
                        sTmp = Mid(sTmp, i + 1)
                        iCnt = InStr(1, sTmp, "^")
                            gsPosNo = Mid(sTmp, 1, iCnt - 1)       'Rack
                        sTmp = Mid(sTmp, 1)
                        iCnt = InStr(1, sTmp, "^")
                            gsRackNo = Mid(sTmp, iCnt + 1)     'pos
                    End If
'                If InStr(1, gsBarCode, "U") > 0 Then '////// Unknown 이 있을시에는
'                    gsBarCode = ""
'                End If
          
            ElseIf gsSampleType = "HC" Or gsSampleType = "LC" Then
                sLotNo = Trim(ResultTbl(16)) 'lotno
                i = InStr(1, sLotNo, "")
                If i > 0 Then
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
                i = InStr(1, sLotNo, "^")
                If i > 0 Then
'                    sLevel = Mid(sLotNo, 1, i - 1)
'                    sLotNo = Mid(sLotNo, i + 1)
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
            End If
        End If
        
        sTmp = Trim(ResultTbl(5))
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            i = InStr(1, sTmp, "^")
            sTmp = Mid(sTmp, i + 1)
            i = InStr(1, sTmp, "^")
            sTmp = Mid(sTmp, i + 1)
            i = InStr(1, sTmp, "^")
            gsSeqNo = Mid(sTmp, i + 1)
        End If
        
        
        
        
        gRow = -1
        For i = 1 To vasID.DataRowCnt
            If gsBarCode <> "" Then
                If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                    gRow = i
                    Exit For
                End If
'            ElseIf sSampleType = "Q" Then

            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
            End If
        End If
        
        SetText vasID, gsBarCode, gRow, colBarcode
        SetText vasID, gsRackNo, gRow, colRack
        SetText vasID, gsPosNo, gRow, colPos
        
        vasActiveCell vasID, gRow, colBarcode
        ClearSpread vasRes
        
        '샘플정보 가져오기
        If gsSampleType = "Q" Then
            SetText vasID, "QC", gRow, colPName
        Else
            If Trim(GetText(vasID, gRow, colPID)) = "" And gsBarCode <> "" And Mid(gsBarCode, 1, 1) <> "U" Then
                Get_Sample_Info gRow
            End If
        End If
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "P") Then          'Test Order Record
        
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        If Trim(GetText(vasID, gRow, colPName)) <> "" Then
        
            gOrderExam = ""
            If MnTransAuto.Checked = True Then
                res = Insert_Data(gRow)
                
                If res = -1 Then
                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                    SetText vasID, "Failed", gRow, colState
                Else
                   
                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                    SetText vasID, "Trans", gRow, colState
                    
                    SQL = " Update pat_res Set " & vbCrLf & _
                          " sendflag = '2' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                End If
                
            End If
            
        End If
    SetText vasID, "Result", gRow, colState
    End If
    

    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Left(sTmp, i - 1)    '장비코드
        i = InStr(1, sTmp, "^")
        lsSubCode = Mid(sTmp, i + 1)
        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
'        gsResDateTime = ResultTbl(10)    'result time
    
'        If Trim(gOrderExam) = "" Then
'            Exit Sub
'        End If
        If lsSubCode <> "AREA" Then: Exit Sub
        
        ClearSpread vasTemp

        SQL = "Select examcode, examname, seqno From equipexam" & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "And equipcode = '" & lsTestID & "' " ' & vbCrLf & _
              "and examcode in (" & gOrderExam & ") "
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
            lsExamCode = Trim(gReadBuf(0))
            lsExamName = Trim(gReadBuf(1))
            lsSeqNo = Trim(gReadBuf(2))
            
            '숫자만 디스플레이 하기
            If IsNumeric(lsResult) = False Then
                For ii = 1 To Len(lsResult)
                    If Mid(lsResult, ii, 1) = "?" Then
                        lsResult = Mid(lsResult, ii + 1)
                        
                        Exit For
                    End If
                Next ii
            End If
            
            lsResRow = vasRes.DataRowCnt + 1
            If vasRes.MaxRows < lsResRow Then
                vasRes.MaxRows = lsResRow
            End If
            
            '소수점 처리, 결과 형태 처리
            
            lsEquipRes = lsResult
            lsResult = SetResult(lsResult, lsTestID)
            lsResult_Buff = lsResult
            
            SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
            SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
            SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
            SetText vasRes, lsResult, lsResRow, colResult            '결과
            
            SetText vasID, lsResult, gRow, colA1c                    '결과
            SetText vasID, gsFlag, gRow, colA1c + 1                  'Flag
            
            SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
            SetText vasRes, gsFlag, lsResRow, 7                      'Flag
            
            
            Save_Local_One gRow, lsResRow, "1", CLng(lsEquipRes)
                        
            If IsNumeric(lsResult) = False Then
                Exit Sub
            End If
'//// IFCC, eAg 체크시
'''            For i = 1 To 2
'''                lsResRow = vasRes.DataRowCnt + 1
'''                If vasRes.MaxRows < lsResRow Then
'''                    vasRes.MaxRows = lsResRow
'''                End If
'''
'''                'IFCC,eAg 결과  처리
'''                If i = 1 Then
'''                    If gADD_IFCC = "-" Then
'''                        lsResult = CStr((CCur(gIFCC1) * CCur(lsResult_Buff)) - CCur(gIFCC2))
'''                    ElseIf gADD_IFCC = "+" Then
'''                        lsResult = CStr((CCur(gIFCC1) * CCur(lsResult_Buff)) + CCur(gIFCC2))
'''                    End If
'''                    lsResult = Format(lsResult, "####")
'''                    lsTestID = "IFCC"
'''                    lsExamCode = "B312002"
'''                    lsExamName = "IFCC"
'''                    lsSeqNo = "2"
'''                    lsResult = SetResult(lsResult, lsTestID)
'''                    SetText vasRes, lsResult, lsResRow, colResult           '결과
'''                    SetText vasID, lsResult, gRow, colIFCC              '결과
'''                    SetText vasID, gsFlag, gRow, colIFCC + 1          'Flag
'''                    SetText vasRes, gsFlag, lsResRow, 7          'Flag
'''                Else
'''                    If gADD_eAg = "-" Then
'''                        lsResult = CStr((CCur(geAg1) * CCur(lsResult_Buff)) - CCur(geAg2))
'''                    ElseIf gADD_eAg = "+" Then
'''                        lsResult = CStr((CCur(geAg1) * CCur(lsResult_Buff)) + CCur(geAg2))
'''                    End If
'''                    lsResult = Format(lsResult, "####")
'''                    lsTestID = "eAg"
'''                    lsExamCode = "B312003"
'''                    lsExamName = "eAg"
'''                    lsSeqNo = "3"
'''                    lsResult = SetResult(lsResult, lsTestID)
'''                    SetText vasRes, lsResult, lsResRow, colResult           '결과
'''                    SetText vasID, lsResult, gRow, coleAg               '결과
'''                    SetText vasID, gsFlag, gRow, coleAg + 1           'Flag
'''                    SetText vasRes, gsFlag, lsResRow, 7          'Flag
'''                End If
'''
'''                SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
'''                SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
'''                SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
'''                SetText vasRes, lsResult, lsResRow, colResult            '결과
'''                SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
'''
'''
'''                Save_Local_One gRow, lsResRow, "1"
'''            Next i
            
            lsResult_Buff = ""
                        
        End If
            
            
    End If
    
End Sub

Function SetResult(asResult As String, asEquipCode As String)
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
'    If IsNumeric(sEquipRes) = False Then
'        Exit Function
'    End If
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)
    
    If IsNumeric(gReadBuf(0)) = True Then
        sPoint = CInt(gReadBuf(0))
        sResType = ""
        For i = 0 To sPoint
            If i = 0 Then
                sResType = "#0"
            ElseIf i = 1 Then
                sResType = sResType & ".0"
            Else
                sResType = sResType & "0"
            End If
        Next
        
        sResult = Format(sEquipRes, sResType)
    Else
        sResult = sEquipRes
    End If
    
''    If IsNumeric(gReadBuf(1)) = True Then
''        sLVal = gReadBuf(1)
''        If CCur(sLVal) > CCur(sEquipRes) Then
''            sResFlag = "H"
''        End If
''    End If
''
''    If IsNumeric(gReadBuf(2)) = True Then
''        sHVal = gReadBuf(2)
''        If CCur(sHVal) < CCur(sEquipRes) Then
''            sResFlag = ">"
''        End If
''    End If
    
    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
        sLVal = gReadBuf(1)
        sHVal = gReadBuf(2)
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
    End If
    gsFlag = sResFlag
    'sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
          SQL = "DELETE FROM PAT_RES " & vbCrLf
    SQL = SQL & "WHERE EXAMDATE  = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "  AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
'    SQL = SQL & "  AND BARCODE   = '" & Trim(GetText(vasResult, asRow1, colBarcode)) & "' " & vbCrLf
    SQL = SQL & "  AND RECENO    = '" & Trim(GetText(vasResult, asRow1, colSpecNo)) & "' " & vbCrLf
    SQL = SQL & "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    SQL = SQL & "  and examcode  = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
    SQL = SQL & "  and isocd     = '" & Trim(GetText(vasRes, asRow2, 8)) & "'"
    SQL = SQL & "  and antcd     = '" & Trim(GetText(vasRes, asRow2, 11)) & "'"
    SQL = SQL & "  and antmachcd     = '" & Trim(GetText(vasRes, asRow2, 15)) & "'"
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
        
    SQL = "INSERT INTO PAT_RES"
    SQL = SQL & "(EQUIPNO, BARCODE, DISKNO,   POSNO,    PID,     PNAME,       PSEX,   PAGE,      EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
                " SEQNO,   RESULT,  EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, SAMPLESEQ, isocd, isonm, antcnt, antcd, antsize, antrslt,exmncd,antmachcd) " & vbCrLf
    SQL = SQL & "VALUES("
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, colBarcode)) & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 5)) & "',"    'PID
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 6)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 10)) & "', "
'    SQL = SQL & Trim(GetText(vasResult, asRow1, 11)) & ", "
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 11)) & "', "
    SQL = SQL & "'" & Trim(sExamDate) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf
    SQL = SQL & "'" & asSend & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 7)) & "', "
    SQL = SQL & "'" & Trim(asEquipResult) & "', "
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, colSpecNo)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 0)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 8)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 9)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 10)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 11)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 12)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 13)) & "',"
    SQL = SQL & "'" & Trim(GetText(vasResult, asRow1, 15)) & "',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, 15)) & "')"
    
    res = SendQuery(gLocal, SQL)

End Function

Function Save_Local_One_Micro(ByVal asRow1 As Long, ByVal strExamCode As String, ByVal strDrug As String, _
                              ByVal strSensi As String, ByVal strVol As String)
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
          SQL = "Delete From PAT_RES "
    SQL = SQL & " Where examdate  = '" & Format(dtpToday, "YYYYMMDD") & "' "
    SQL = SQL & "   and equipno   = '" & gEquip & "' "
    SQL = SQL & "   and barcode   = '" & Trim(mResult.BarNo) & "' "
    SQL = SQL & "   and equipcode = '" & Trim(strDrug) & "'"
    SQL = SQL & "   and examcode  = '" & Trim(strExamCode) & "'"
    SQL = SQL & "   and isocd     = '" & Trim(mResult.MnmCd) & "'"
    SQL = SQL & "   and antcd     = '" & Trim(strDrug) & "'"
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
          SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, POSNO, "
    SQL = SQL & "                    PID, PNAME, PSEX, PAGE, " & vbCrLf & _
                "                    EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
                "                    SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, SAMPLESEQ, " & vbCrLf & _
                "                    isocd, isonm, antcnt, antcd, antsize, antrslt) "
                
    SQL = SQL & " VALUES('" & gEquip & "', "
    SQL = SQL & "        '" & Trim(mResult.BarNo) & "',"
    SQL = SQL & "        '', "
    SQL = SQL & "        '', "
    SQL = SQL & "        '" & Trim(GetText(vasID, asRow1, colPID)) & "', "
    SQL = SQL & "        '" & Trim(GetText(vasID, asRow1, colPName)) & "',"
    SQL = SQL & "        '" & Trim(GetText(vasID, asRow1, colSex)) & "', "
    SQL = SQL & "        0, "
    SQL = SQL & "        '" & Trim(sExamDate) & "', "
    SQL = SQL & "        '" & Trim(strDrug) & "', "
    SQL = SQL & "        '" & Trim(strExamCode) & "', "
    SQL = SQL & "        '" & Trim(GetText(vasRes, asRow1, colSeq)) & "', "
    SQL = SQL & "        '" & Trim(strVol) & "', "
    SQL = SQL & "        '" & Trim(GetText(vasRes, asRow1, colExamName)) & "', "
    SQL = SQL & "        '1', "
    SQL = SQL & "        '" & Trim(GetText(vasRes, asRow1, 7)) & "', "
    SQL = SQL & "        '" & Trim(strVol) & "', "
    SQL = SQL & "        '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', "
    SQL = SQL & "        '" & Trim(GetText(vasID, asRow1, 0)) & "', "
    SQL = SQL & "        '" & Trim(mResult.MnmCd) & "', "
    SQL = SQL & "        '" & Trim(mResult.MnmNm) & "', "
    SQL = SQL & "        '" & Trim(24) & "', "
    SQL = SQL & "        '" & Trim(strDrug) & "', "
    SQL = SQL & "        '" & Trim(strSensi) & "', "
    SQL = SQL & "        '" & Trim(strVol) & "')"
    res = SendQuery(gLocal, SQL)

    
End Function

Sub Var_Clear()
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""
End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")

End Sub

'Private Sub spdResult3_KeyPress(KeyAscii As Integer)
'
'    With spdResult3
'        If KeyAscii = vbKeyReturn Then
'            If .ActiveCol = 3 Then
'                Call EditAntVal(GetText(spdResult3, .ActiveRow, 1), GetText(spdResult3, .ActiveRow, .ActiveCol))
'            ElseIf .ActiveCol = 7 Then
'                Call EditAntVal(GetText(spdResult3, .ActiveRow, 5), GetText(spdResult3, .ActiveRow, .ActiveCol))
'            End If
'        End If
'    End With
'
'
'End Sub

Private Sub EditAntVal(ByVal strAntCd As String, ByVal strAntVal As String)

    If strAntCd <> "" And strAntVal <> "" Then
              SQL = "UPDATE PAT_RES "
        SQL = SQL & "   SET RESULT      = '" & Trim(strAntVal) & "', "
        SQL = SQL & "       EQUIPRESULT = '" & Trim(strAntVal) & "', "
        SQL = SQL & "       ANTRSLT     = '" & Trim(strAntVal) & "' "
        SQL = SQL & " WHERE EXAMDATE    = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
        SQL = SQL & "   AND EQUIPNO     = '" & gEquip & "'"
        SQL = SQL & "   AND EQUIPCODE   = '" & strAntCd & "'"
        SQL = SQL & "   AND BARCODE     = '" & GetText(spdResult2, 1, 2) & "'"
        SQL = SQL & "   AND RECENO      = '" & GetText(spdResult2, 1, 1) & "'"
        SQL = SQL & "   AND ISOCD       = '" & GetText(spdResult2, 1, 3) & "'"
        SQL = SQL & "   AND ANTCD       = '" & strAntCd & "'"
        
        cn.Execute SQL
        
        Call vasResult_Click(1, vasResult.ActiveRow)
    
    End If
    
End Sub


Private Sub spdResult3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim strAnti As String
Dim strSeq  As String
Dim strMachAnti As String

    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 세균 삭제")

    With spdResult3
        Select Case lMenuChosen
            Case 1
                .Row = Row
                
                If Col = 5 Then
                    Exit Sub
                ElseIf Col <= 4 Then
                    strMachAnti = GetText(spdResult3, Row, 1)
                    strAnti = GetText(spdResult3, Row, 2)
                Else
                    strMachAnti = GetText(spdResult3, Row, 6)
                    strAnti = GetText(spdResult3, Row, 7)
                End If
                
                Call DelAntiVal(strAnti, strMachAnti)
                
        End Select
    End With
End Sub



Private Sub DelAntiVal(ByVal strAnti As String, Optional ByVal strMachAnti As String)

    If strAnti <> "" Then
'              SQL = "DELETE FROM PAT_RES "
'        SQL = SQL & " WHERE EXAMDATE  = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
'        SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "'"
'        SQL = SQL & "   AND EQUIPCODE = '" & strAnti & "'"
'        SQL = SQL & "   AND BARCODE   = '" & GetText(vasResult, 1, 3) & "'"
'        SQL = SQL & "   AND RECENO    = '" & GetText(vasResult, 1, 2) & "'"
'        SQL = SQL & "   AND ISOCD     = '" & GetText(vasResult, 1, 7) & "'"
'        SQL = SQL & "   AND ANTCD     = '" & strAnti & "'"

              SQL = "DELETE FROM PAT_RES "
        SQL = SQL & " WHERE EXAMDATE  = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
        SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "'"
        SQL = SQL & "   AND EQUIPCODE = '" & strAnti & "'"
        SQL = SQL & "   AND BARCODE   = '" & GetText(vasResult, vasResult.ActiveRow, 3) & "'"
        SQL = SQL & "   AND RECENO    = '" & GetText(vasResult, vasResult.ActiveRow, 2) & "'"
        SQL = SQL & "   AND ISOCD     = '" & GetText(vasResult, vasResult.ActiveRow, 7) & "'"
        SQL = SQL & "   AND ANTCD     = '" & strAnti & "'"
        SQL = SQL & "   AND ANTMACHCD     = '" & strMachAnti & "'"
    
        cn.Execute SQL
        
        Call vasResult_Click(1, vasResult.ActiveRow)
    
    End If
    
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasID, Row, colPos)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
    
End Sub

'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasID.ActiveRow
'    If KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'              " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasID, iRow, iRow
'        vasRes.MaxRows = 0
'    ElseIf KeyCode = 13 Then
'
'        Get_Sample_Info (iRow)
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
'              "  FROM EQUIPEXAM " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " ORDER BY SEQNO "
'
'        res = db_select_Vas(gLocal, SQL, vasTemp)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
'                  "POSNO, PID, PNAME, " & vbCrLf & _
'                  "JUMIN, PSEX, PAGE, " & vbCrLf & _
'                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
'                  "SEQNO, RESULT, EXAMNAME, " & vbCrLf & _
'                  "SENDFLAG, Hospital, refflag) " & vbCrLf & _
'                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, iRow, colBarcode)) & "', '" & Trim(GetText(vasID, iRow, colRack)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colPos)) & "', '" & Trim(GetText(vasID, iRow, colPID)) & "', '" & Trim(GetText(vasID, iRow, colPName)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colJumin)) & "', '" & Trim(GetText(vasID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
'                  "'" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 2)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasTemp, i, 4)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 3)) & "', " & vbCrLf & _
'                  "'1', '" & Trim(GetText(vasID, iRow, colHospital)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2))) & "')"
'                res = SendQuery(gLocal, SQL)
'            Next i
'
'            SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                  " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
'                  " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
'                  " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'                  " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'                  " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'            res = SendQuery(gLocal, SQL)
'
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "UPDATE PAT_RES "
'                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, iRow, colBarcode)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasTemp, i, 2)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' "
'                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' "
'                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'                res = SendQuery(gLocal, SQL)
'            Next i
'        End If
'        SetText vasID, "Result", gRow, colState
'
'    End If
'
'
'End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If OrgSort_Flag = 1 Then
            Call SpreadSheetSort(vasID, Col, 2)
            OrgSort_Flag = 2
        Else
            Call SpreadSheetSort(vasID, Col, 1)
            OrgSort_Flag = 1
        End If
    End If

End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_Click colBarcode, lRow
    End If
End Sub

'Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
'    Dim sResDateTime As String
'    Dim sControl As String
'    Dim sLotNo As String
'
'    Dim sRefLow As String
'    Dim sRefHigh As String
'    Dim sRefFlag As String
'
'    Dim sCnt As String
'
'    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
'    'sControl = Trim(Left(asBarcode, 2))
'    'sLotNo = Trim(Mid(asBarcode, 3))
'    sControl = asBarcode
'    sRefFlag = ""
'
'    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
'            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
'            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
'            If CCur(sRefHigh) < CCur(asRes2) Then
'                sRefFlag = "H"
'            End If
'            If CCur(sRefLow) > CCur(asRes2) Then
'                sRefFlag = "L"
'            End If
'        End If
'    End If
'
'    sCnt = ""
'    SQL = "Select count(*) from qc_res " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        db_RollBack gLocal
'        Exit Function
'    End If
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'    If Not IsNumeric(sCnt) Then sCnt = "0"
'
'    If CInt(sCnt) > 0 Then
'        SQL = "delete from qc_res " & vbCrLf & _
'              "where equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
'              "  and levelname = '" & sControl & "' " & vbCrLf & _
'              "  and equipcode = '" & asExamCode & "' "
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            'db_RollBack gLocal
'            SaveQuery SQL
'            Exit Function
'        End If
'    End If
'    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
'          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        'db_RollBack gLocal
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function

Private Sub vasResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = 1 Then
        With vasResult
            .RowHeight(-1) = 12
            .Col = 1: .Col2 = .MaxCols
            .Row = Row: .Row2 = Row
            .BlockMode = True
            If .FontBold = True Then
                .FontBold = False
            Else
                .FontBold = True
            End If
            .BlockMode = False
            
            .RowHeight(-1) = 12
            Exit Sub
        End With
    End If

End Sub

Private Sub vasResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim recNo As String
    Dim strTestCd As String
    Dim strIsoCd As String
    Dim i As Integer
    
    Dim adors As ADODB.Recordset
    
    
    If Row < 1 Or Row > vasResult.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasResult, Row, colBarcode))
    recNo = Trim(GetText(vasResult, Row, colSpecNo))
    strTestCd = Trim(GetText(vasResult, Row, 4))
    strIsoCd = Trim(GetText(vasResult, Row, 7))
    
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasResult, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG,'',isocd, isonm, antcnt, antcd, antsize, antrslt, exmncd, antmachcd  " & vbCrLf
    SQL = SQL & "  FROM PAT_RES " & vbCrLf
    SQL = SQL & "WHERE EQUIPNO = '" & gEquip & "' "
'    SQL = SQL & "  AND BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "  AND RECENO  = '" & recNo & "' " & vbCrLf
    SQL = SQL & "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf
    SQL = SQL & "  AND EXAMCODE  = '" & strTestCd & "' " & vbCrLf
    SQL = SQL & "  AND ISOCD  = '" & strIsoCd & "' " & vbCrLf
    SQL = SQL & "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG,isocd, isonm, antcnt, antcd, antsize, antrslt, exmncd,antmachcd "
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
    
    With spdResult3
        .MaxRows = 24
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    
          SQL = "SELECT isocd, isonm, antcnt, antcd, antsize, antrslt, antmachcd " & vbCrLf
    SQL = SQL & "FROM PAT_RES " & vbCrLf
    SQL = SQL & "WHERE EQUIPNO = '" & gEquip & "' "
'    SQL = SQL & "  AND BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "  AND EXAMCODE  = '" & strTestCd & "' " & vbCrLf
    SQL = SQL & "  AND ISOCD  = '" & strIsoCd & "' " & vbCrLf
    SQL = SQL & "  AND RECENO = '" & recNo & "' " & vbCrLf
    SQL = SQL & "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "

    Set adors = New ADODB.Recordset
    adors.CursorLocation = adUseClient
    adors.Open SQL, cn

    If Not adors.EOF Then
        Call SetText(spdResult2, Trim(GetText(vasResult, Row, colSpecNo)), 1, 1)
        Call SetText(spdResult2, lsID, 1, 2)
        Call SetText(spdResult2, Trim(adors("isocd").Value & ""), 1, 3)
        Call SetText(spdResult2, Trim(adors("isonm").Value & ""), 1, 4)
        Call SetText(spdResult2, Trim(adors("antcnt").Value & ""), 1, 5)

        Do While Not adors.EOF
            Call SetDrugInfo(lsID, Trim(adors("antmachcd").Value & ""), Trim(adors("antcd").Value & ""), Trim(adors("antrslt").Value & ""), Trim(adors("antsize").Value & ""))
            adors.MoveNext
        Loop

    End If

End Sub


Private Sub vasResult_DblClick(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If RsltSort_Flag = 1 Then
            Call SpreadSheetSort(vasResult, Col, 2)
            RsltSort_Flag = 2
        Else
            Call SpreadSheetSort(vasResult, Col, 1)
            RsltSort_Flag = 1
        End If
    End If
    
End Sub

Private Sub vasResult_KeyPress(KeyAscii As Integer)

    With vasResult
        If KeyAscii = vbKeyReturn Then
'            If .ActiveCol = 2 Or .ActiveCol = 4 Or .ActiveCol = 7 Then
            If .ActiveCol = 2 Or .ActiveCol = 7 Then
                Call EditMICVal(.ActiveCol, .ActiveRow)
            End If
        End If
    End With

End Sub

Private Sub EditMICVal(ByVal lngCol As Long, ByVal lngRow As Long)
    Dim rs_orgnm As ADODB.Recordset
    Dim strOrgNm As String
    Dim strWorkNo As String
    Dim strOrgWorkNo As String
    Dim strOrgBarNo As String
    Dim strOrgExmnCd As String
    Dim strNewExmnCd As String
    Dim intRow As Integer
    Dim intRow2 As Integer
    Dim intRow3 As Integer
    Dim varTmp As Variant
    
    If lngRow <> 0 Then
        If lngCol = 2 Then
            strNewExmnCd = ""
            strOrgBarNo = GetText(vasResult, lngRow, 3)
            strOrgExmnCd = GetText(vasResult, lngRow, 4)
            strWorkNo = GetText(vasResult, lngRow, 2)
            strWorkNo = Mid(strWorkNo, 1, 11) & "00I" & Mid(strWorkNo, 12, 4)
            '-- 검사코드 가져오기
                  SQL = "Select EXMN_CD From SPSLHRRST "
            SQL = SQL & " Where WORK_NO = '" & strWorkNo & "'"
            SQL = SQL & "   and substr(EXMN_CD,1,3) <> 'L40'"
            SQL = SQL & "   and RSLT_NO IS NOT NULL"
            'SQL = SQL & "   and RSLT_STAT <> '3' "
            'MsgBox strWorkNo
            If InStr(strWorkNo, "L4B") > 0 Then
                If Len(strOrgExmnCd) = 6 Then
                    SQL = SQL & "   and EXMN_CD = '" & strOrgExmnCd & "' "
                ElseIf Len(strOrgExmnCd) = 8 Then
                    SQL = SQL & "   and EXMN_CD = '" & Mid(strOrgExmnCd, 1, 6) & "' "
                End If
            End If
            
            Set rs_orgnm = cn_Ser.Execute(SQL)
            intRow = 0
            intRow2 = 0
            Do Until rs_orgnm.EOF
                strNewExmnCd = strNewExmnCd & "'" & rs_orgnm.Fields(0).Value & "',"
                SetText vasResult, "", lngRow, 3
                SetText vasResult, "", lngRow, 4
                
                rs_orgnm.MoveNext
            Loop
            
            Set rs_orgnm = Nothing
            
            If strNewExmnCd <> "" Then
                strNewExmnCd = Mid(strNewExmnCd, 1, Len(strNewExmnCd) - 1)
            Else
                Exit Sub
            End If
            
            '-- 검사대상자 가져오기
            SQL = "Select SPCM_NO,PID From SPSLHRRST "
            SQL = SQL & " Where WORK_NO = '" & strWorkNo & "'"
'            SQL = SQL & "   and EXMN_CD = '" & GetText(vasResult, lngRow, 4) & "'"
            SQL = SQL & "   and EXMN_CD in (" & strNewExmnCd & ")"
            SQL = SQL & "   and RSLT_NO IS NOT NULL"
'            SQL = SQL & "   and RSLT_STAT <> '3' "

'Text1.Text = SQL
            Set rs_orgnm = cn_Ser.Execute(SQL)
            varTmp = Split(strNewExmnCd, ",")
            Do Until rs_orgnm.EOF
                strOrgNm = rs_orgnm.Fields(0).Value & ""
                strNewExmnCd = Replace(strNewExmnCd, "'", "")
                SetText vasResult, strOrgNm, lngRow, 3
                
                If InStr(strWorkNo, "L4B") > 0 Then
                    SetText vasResult, strOrgExmnCd, lngRow, 4
                Else
                    SetText vasResult, strNewExmnCd, lngRow, 4
                End If
            
                SetText vasResult, rs_orgnm.Fields(1).Value & "", lngRow, 5
                
                rs_orgnm.MoveNext
            Loop
            
            Set rs_orgnm = Nothing
            
                  SQL = "UPDATE PAT_RES "
            SQL = SQL & "   SET RECENO   = '" & Trim(GetText(vasResult, lngRow, 2)) & "', "
            SQL = SQL & "       BARCODE  = '" & Trim(GetText(vasResult, lngRow, 3)) & "', "
'            SQL = SQL & "       EXAMCODE = '" & strNewExmnCd & "'"
            If InStr(strWorkNo, "L4B") > 0 Then
                SQL = SQL & "       EXAMCODE = '" & strOrgExmnCd & "'"
            Else
                SQL = SQL & "       EXAMCODE = '" & strNewExmnCd & "'"
            End If
            
            SQL = SQL & " WHERE EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
            SQL = SQL & "   AND EQUIPNO  = '" & gEquip & "'"
            SQL = SQL & "   AND BARCODE  = '" & strOrgBarNo & "'"
            'SQL = SQL & "   AND RECENO   = '" & GetText(vasResult, lngRow, 2) & "'"
            SQL = SQL & "   AND EXAMCODE = '" & strOrgExmnCd & "'"
            SQL = SQL & "   AND ISOCD    = '" & GetText(vasResult, lngRow, 7) & "'"
            cn.Execute SQL
                        
                  SQL = "DELETE FROM PAT_RES "
            SQL = SQL & " WHERE EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
            SQL = SQL & "   AND EQUIPNO  = '" & gEquip & "'"
            SQL = SQL & "   AND BARCODE  = '" & strOrgBarNo & "'"
            SQL = SQL & "   AND EXAMCODE = '" & strOrgExmnCd & "'"
            SQL = SQL & "   AND ISOCD    = '" & GetText(vasResult, lngRow, 7) & "'"
            cn.Execute SQL
                      
        ElseIf lngCol = 4 Then '-- 검사코드
            Exit Sub
            
                  SQL = "UPDATE PAT_RES "
            SQL = SQL & "   SET EXAMCODE = '" & Trim(GetText(vasResult, lngRow, 4)) & "' "
            SQL = SQL & " WHERE EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
            SQL = SQL & "   AND EQUIPNO  = '" & gEquip & "'"
            SQL = SQL & "   AND BARCODE  = '" & GetText(vasResult, lngRow, 3) & "'"
            SQL = SQL & "   AND RECENO   = '" & GetText(vasResult, lngRow, 2) & "'"
            SQL = SQL & "   AND ISOCD    = '" & GetText(vasResult, lngRow, 7) & "'"
            cn.Execute SQL

        ElseIf lngCol = 7 Then  '-- 균코드
            Set rs_orgnm = New ADODB.Recordset
            
                  SQL = "SELECT orgnm From orgtable "
            SQL = SQL & " WHERE morgcd = '" & Trim(GetText(vasResult, lngRow, 7)) & "' "
            Set rs_orgnm = cn.Execute(SQL)
            Do Until rs_orgnm.EOF
                'Call vasResult.SetText(lngRow, lngCol + 1, rs_orgnm.Fields(0).Value & "")
                strOrgNm = rs_orgnm.Fields(0).Value & ""
                SetText vasResult, strOrgNm, lngRow, 8
                rs_orgnm.MoveNext
            Loop
            
            Set rs_orgnm = Nothing
            
            Dim strMnmcd As String
            
            Set rs_orgnm = New ADODB.Recordset

                  SQL = "SELECT horgcd From orgtable "
            SQL = SQL & " WHERE morgcd = '" & Trim(GetText(vasResult, lngRow, 7)) & "' "
            Set rs_orgnm = cn.Execute(SQL)
            Do Until rs_orgnm.EOF
                strMnmcd = rs_orgnm.Fields(0).Value & ""
                SetText vasResult, strMnmcd, lngRow, 7
                'mResult.MnmCd = strMnmcd
                rs_orgnm.MoveNext
            Loop

            Set rs_orgnm = Nothing

            If strMnmcd <> "" Then
                Set rs_orgnm = New ADODB.Recordset

                      SQL = "SELECT DISTINCT bctr_cd From SPSLMFMBA "
                SQL = SQL & " WHERE bctr_cd = '" & strMnmcd & "' "
                SQL = SQL & " Union all "
                SQL = SQL & "SELECT DISTINCT bctr_cd From SPSLMFMBA "
                SQL = SQL & " WHERE bctr_itcn_cd = '" & strMnmcd & "' "
                Set rs_orgnm = cn_Ser.Execute(SQL)
                Do Until rs_orgnm.EOF
                    strMnmcd = rs_orgnm.Fields(0).Value & ""
                    'mResult.MnmCd = strMnmcd
                    SetText vasResult, strMnmcd, lngRow, 7
                    rs_orgnm.MoveNext
                Loop

                Set rs_orgnm = Nothing
            End If
            
            Set rs_orgnm = Nothing
            
                  SQL = "UPDATE PAT_RES "
            SQL = SQL & "   SET ISOCD    = '" & Trim(GetText(vasResult, lngRow, 7)) & "', "
            SQL = SQL & "       ISONM    = '" & strOrgNm & "' "
            SQL = SQL & " WHERE EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' "
            SQL = SQL & "   AND EQUIPNO  = '" & gEquip & "'"
            SQL = SQL & "   AND BARCODE  = '" & GetText(vasResult, lngRow, 3) & "'"
            SQL = SQL & "   AND EXAMCODE = '" & Trim(GetText(vasResult, lngRow, 4)) & "'"
            SQL = SQL & "   AND RECENO   = '" & GetText(vasResult, lngRow, 2) & "'"
            
            cn.Execute SQL
        End If
       
        
        Call vasResult_Click(1, vasResult.ActiveRow)
    End If
    
End Sub


Private Sub vasResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
'    If Col <> 1 Or NewCol <> 1 Then
    If Row <> NewRow Then
        Call vasResult_Click(NewCol, NewRow)
    End If

End Sub


Private Sub vasResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim strAnti As String
Dim strSeq  As String

    strAnti = GetText(vasResult, Row, 1)
    If strAnti <> "1" Then Exit Sub
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 세균 삭제")

    With vasResult
        Select Case lMenuChosen
            Case 1
                .Row = Row
                
'                If Col = 4 Then
'                    Exit Sub
'                ElseIf Col <= 3 Then
'                    strAnti = GetText(spdResult3, Row, 1)
'                Else
'                    strAnti = GetText(spdResult3, Row, 5)
'                End If
'
'                Call DelAntiVal(strAnti)
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With
    
End Sub

Private Sub vasRID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.Value = 0 Then
        vasRID.Value = 1
        Else
        vasRID.Value = 0
        End If
    Next i
End Sub

Private Sub vasRID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasRID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblBarcode.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
    lblPname.Caption = Trim(GetText(vasRID, Row, colPName))
    lblRrow.Caption = Row
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
          " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG , EQUIPRESULT"
    
    res = db_select_Vas(gLocal, SQL, vasRRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRRes.MaxRows = vasRRes.DataRowCnt
    
    For i = 1 To vasRRes.MaxRows
        If Trim(GetText(vasRRes, i, colFLAG)) = "H" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 255, 0, 0
        ElseIf Trim(GetText(vasRRes, i, colFLAG)) = "L" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 0, 255, 0
        End If
    Next i
End Sub

Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsPid As String
    Dim i As Integer
    
    iRow = vasRID.ActiveRow
    
    If KeyCode = 13 Then
        
        Get_Sample_InfoR (iRow)
        
        lsID = Trim(GetText(vasRID, iRow, colBarcode))
        
        'Local에서 불러오기
        ClearSpread vasTemp
        
        '장비코드, 검사코드, 검사명, 결과, 순번
        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
              "FROM PAT_RES " & vbCrLf & _
              "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
              "  AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf & _
              "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "

        res = db_select_Vas(gLocal, SQL, vasTemp)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        If lsID <> lblChangeBar.Caption Then
            For i = 1 To vasRRes.DataRowCnt
                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
                  "POSNO, PID, PNAME, " & vbCrLf & _
                  " PSEX, PAGE, " & vbCrLf & _
                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
                  "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, RECENO, EQUIPRESULT) " & vbCrLf & _
                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasRID, iRow, colBarcode)) & "', '" & Trim(GetText(vasRID, iRow, colRack)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colPos)) & "', '" & Trim(GetText(vasRID, iRow, colPID)) & "', '" & Trim(GetText(vasRID, iRow, colPName)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
                  "'" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasRRes, i, 1)) & "', '" & Trim(GetText(vasRRes, i, 2)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRRes, i, 5)) & "', '" & Trim(GetText(vasRRes, i, 4)) & "', '" & Trim(GetText(vasRRes, i, 3)) & "', " & vbCrLf & _
                  "'1', '" & Trim(GetText(vasRRes, i, colFLAG)) & "','" & Trim(GetText(vasRID, iRow, colSpecNo)) & "', '" & Trim(GetText(vasRRes, i, 7)) & "')"
                res = SendQuery(gLocal, SQL)
            Next i
            
                SQL = " DELETE FROM PAT_RES " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
                      " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
                      " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
                      " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
                      " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
                res = SendQuery(gLocal, SQL)
        ElseIf lsID = lblChangeBar.Caption Then
            For i = 1 To vasRRes.DataRowCnt
                SQL = "UPDATE PAT_RES "
                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasRRes, i, 4)) & "' "
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasRRes, i, 2)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, 1)) & "' "
                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasRID, iRow, colPID)) & "' "
                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' "
                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' "
                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
                res = SendQuery(gLocal, SQL)
            Next i
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasRID, iRow, colBarcode))
        lsPid = Trim(GetText(vasRID, iRow, colPID))
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
              " AND PID = '" & lsPid & "' " & vbCrLf & _
              " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
              " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
              " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasRID, iRow, iRow
        vasRRes.MaxRows = 0
        
    End If
End Sub

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub

Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then: vasRID_KeyDown KeyCode, 0
End Sub
