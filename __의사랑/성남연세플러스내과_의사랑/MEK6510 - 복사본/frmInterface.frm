VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "MEK6510 Interface "
   ClientHeight    =   10680
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   23985
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   10680
   ScaleWidth      =   23985
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   15180
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   240
         TabIndex        =   48
         Top             =   1620
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
         SpreadDesigner  =   "frmInterface.frx":14F5
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1275
         Left            =   330
         TabIndex        =   77
         Top             =   360
         Width           =   8055
         _Version        =   393216
         _ExtentX        =   14208
         _ExtentY        =   2249
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":2FA6
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   6375
      Left            =   15150
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   1170
         TabIndex        =   18
         Top             =   5010
         Visible         =   0   'False
         Width           =   4155
         Begin VB.CheckBox chkSave 
            Caption         =   "저장포함"
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
            Height          =   405
            Left            =   2640
            TabIndex        =   68
            Top             =   360
            Width           =   705
         End
         Begin MSCommLib.MSComm comEqp 
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
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   435
         Left            =   7290
         TabIndex        =   59
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTest 
         Height          =   675
         Left            =   4440
         TabIndex        =   58
         Top             =   1080
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   405
         Left            =   5370
         TabIndex        =   55
         Top             =   2580
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1500
         Picture         =   "frmInterface.frx":7602
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   5910
         Width           =   285
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1545
         Left            =   240
         TabIndex        =   35
         Top             =   3270
         Width           =   3495
         _Version        =   393216
         _ExtentX        =   6165
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
         SpreadDesigner  =   "frmInterface.frx":7B8C
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   150
         TabIndex        =   33
         Top             =   5760
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
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   23
         Top             =   5220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   4830
         TabIndex        =   22
         Top             =   5160
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
         Left            =   4830
         TabIndex        =   21
         Top             =   5655
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1770
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   20
         Top             =   5220
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
         Left            =   3660
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   5190
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1755
         Left            =   3870
         TabIndex        =   17
         Top             =   3180
         Width           =   4425
         _Version        =   393216
         _ExtentX        =   7805
         _ExtentY        =   3096
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
         SpreadDesigner  =   "frmInterface.frx":7DDC
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1425
         Left            =   240
         TabIndex        =   24
         Top             =   330
         Width           =   3435
         _Version        =   393216
         _ExtentX        =   6059
         _ExtentY        =   2514
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
         SpreadDesigner  =   "frmInterface.frx":802C
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2175
         Left            =   3750
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   4395
         _Version        =   393216
         _ExtentX        =   7752
         _ExtentY        =   3836
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
         SpreadDesigner  =   "frmInterface.frx":827C
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   270
         TabIndex        =   60
         Top             =   1860
         Width           =   3375
         _Version        =   393216
         _ExtentX        =   5953
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
         SpreadDesigner  =   "frmInterface.frx":84CC
      End
      Begin FPSpread.vaSpread vasExcel 
         Height          =   1185
         Left            =   6000
         TabIndex        =   65
         Top             =   5100
         Visible         =   0   'False
         Width           =   1935
         _Version        =   393216
         _ExtentX        =   3413
         _ExtentY        =   2090
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
         SpreadDesigner  =   "frmInterface.frx":871C
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
         Left            =   1800
         TabIndex        =   49
         Top             =   5910
         Width           =   1185
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3210
         TabIndex        =   27
         Top             =   5850
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   4050
         TabIndex        =   26
         Top             =   5820
         Width           =   915
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10185
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   17965
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
      TabCaption(0)   =   "WorkList"
      TabPicture(0)   =   "frmInterface.frx":896C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "받은결과"
      TabPicture(1)   =   "frmInterface.frx":8988
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   9675
         Left            =   -74820
         TabIndex        =   9
         Top             =   360
         Width           =   14625
         Begin VB.ComboBox cboPartR 
            Height          =   315
            Left            =   4380
            TabIndex        =   79
            Text            =   "Combo1"
            Top             =   300
            Width           =   1005
         End
         Begin VB.CommandButton cmdRIFTrans 
            Caption         =   "&Save"
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
            Left            =   8040
            TabIndex        =   78
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox chkQC 
            Caption         =   "QC만조회"
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
            Height          =   405
            Left            =   10530
            TabIndex        =   67
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdXmlDel 
            Caption         =   "XML정리"
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
            Left            =   11610
            TabIndex        =   61
            Top             =   270
            Width           =   1395
         End
         Begin VB.Frame Frame5 
            Height          =   825
            Left            =   10410
            TabIndex        =   28
            Top             =   690
            Width           =   4035
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   34
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   1500
               TabIndex        =   32
               Top             =   540
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
               Left            =   360
               TabIndex        =   31
               Top             =   540
               Width           =   1005
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1515
               TabIndex        =   30
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "챠트번호 :"
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
               Left            =   150
               TabIndex        =   29
               Top             =   240
               Width           =   1230
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
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
            Left            =   6750
            TabIndex        =   15
            Top             =   240
            Width           =   1275
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
            Left            =   5460
            TabIndex        =   14
            Top             =   240
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   13
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   133103617
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   11
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "&Clear"
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
            Left            =   13020
            TabIndex        =   10
            Top             =   270
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7965
            Left            =   10410
            TabIndex        =   12
            Top             =   1605
            Width           =   4005
            _Version        =   393216
            _ExtentX        =   7064
            _ExtentY        =   14049
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
            MaxCols         =   6
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":89A4
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   165
            TabIndex        =   46
            Top             =   750
            Width           =   10155
            _Version        =   393216
            _ExtentX        =   17912
            _ExtentY        =   15531
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
            MaxCols         =   8
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":C75D
            UserResize      =   2
         End
         Begin MSComCtl2.DTPicker dtpExamDate1 
            Height          =   315
            Left            =   2850
            TabIndex        =   63
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   133103617
            CurrentDate     =   40457
         End
         Begin VB.Label Label10 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2640
            TabIndex        =   62
            Top             =   390
            Width           =   105
         End
         Begin VB.Label Label9 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   45
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9645
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton cmdWDown 
            BackColor       =   &H00C0FFFF&
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7020
            Style           =   1  '그래픽
            TabIndex        =   75
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdWUp 
            BackColor       =   &H00C0FFFF&
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6630
            Style           =   1  '그래픽
            TabIndex        =   74
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdUp 
            BackColor       =   &H00C0FFFF&
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            Style           =   1  '그래픽
            TabIndex        =   73
            Top             =   4950
            Width           =   525
         End
         Begin VB.CommandButton cmdDown 
            BackColor       =   &H00C0FFFF&
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5370
            Style           =   1  '그래픽
            TabIndex        =   72
            Top             =   4950
            Width           =   525
         End
         Begin VB.CommandButton cmdWPrint 
            Caption         =   "출력"
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
            Left            =   11280
            TabIndex        =   71
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "출력"
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
            Left            =   6390
            TabIndex        =   70
            Top             =   4950
            Width           =   1005
         End
         Begin VB.ComboBox cboPart 
            Height          =   315
            Left            =   3720
            TabIndex        =   69
            Text            =   "Combo1"
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5940
            TabIndex        =   66
            Text            =   "0"
            Top             =   4980
            Width           =   435
         End
         Begin VB.CommandButton cmdDownLoad 
            Caption         =   "Down"
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
            Left            =   5670
            TabIndex        =   57
            Top             =   180
            Width           =   915
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "조회"
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
            Left            =   4740
            TabIndex        =   50
            Top             =   180
            Width           =   915
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "&Clear"
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
            Left            =   12360
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "&Save"
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
            Left            =   13440
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7530
            TabIndex        =   39
            Top             =   630
            Width           =   6945
            Begin VB.Label Label8 
               Caption         =   "챠트번호 :"
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
               TabIndex        =   44
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   0
               Left            =   1605
               TabIndex        =   43
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label6 
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
               TabIndex        =   42
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4200
               TabIndex        =   41
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   40
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   930
            TabIndex        =   5
            Top             =   780
            Width           =   255
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8070
            Left            =   7530
            TabIndex        =   6
            Top             =   1425
            Width           =   6915
            _Version        =   393216
            _ExtentX        =   12197
            _ExtentY        =   14235
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
            MaxCols         =   6
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":D000
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1815
            Left            =   8505
            TabIndex        =   3
            Top             =   6720
            Visible         =   0   'False
            Width           =   5970
            Begin VB.TextBox txtErrLog 
               Appearance      =   0  '평면
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   4
               Top             =   240
               Width           =   5775
            End
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   8640
            TabIndex        =   36
            Top             =   270
            Width           =   2505
            _ExtentX        =   4419
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
            Format          =   133103616
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   4155
            Left            =   180
            TabIndex        =   51
            Top             =   690
            Width           =   7215
            _Version        =   393216
            _ExtentX        =   12726
            _ExtentY        =   7329
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
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
            MaxCols         =   9
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":10D3E
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1080
            TabIndex        =   52
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   133103617
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2460
            TabIndex        =   53
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   133103617
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   4125
            Left            =   180
            TabIndex        =   76
            Top             =   5400
            Width           =   7215
            _Version        =   393216
            _ExtentX        =   12726
            _ExtentY        =   7276
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
            MaxCols         =   16
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":117E3
            UserResize      =   2
         End
         Begin VB.Label Label7 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2310
            TabIndex        =   56
            Top             =   300
            Width           =   105
         End
         Begin VB.Label Label5 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7620
            TabIndex        =   54
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "조회일자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   37
            Top             =   300
            Width           =   780
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   23985
      _ExtentX        =   42307
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13124
            MinWidth        =   13124
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "오후 3:47"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2017-12-05"
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
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
      Begin VB.Menu MnTransAuto 
         Caption         =   "Auto"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
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
Const colSpecNo = 0 '미사용
Const colCheckBox = 1
Const colBarcode = 2
Const colRack = 3
Const colDISK = 3
Const colPos = 4
Const colBarNo = 4
Const colPID = 5
Const colPName = 6
Const colSex = 7
Const colAge = 8
Const colOCnt = 9
Const colRCnt = 10
Const colState = 11

Const colA1c = 12
Const colIFCC = 13
Const coleAg = 14




'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
'Const colMachResult = 4
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
Const RS  As String = ""
Const GS  As String = ""
Const SB As String = ""  'Chr(11)
Const EB As String = ""   'Chr(28)

Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer

Dim blnSameRecord As Boolean

Private Type typeXMLInData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLInData As typeXMLInData

Private Type typeXMLOutData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLOutData As typeXMLOutData

Public RS1 As ADODB.Recordset
Const FieldCnt As Integer = 60
Const Fld_Div As String = ""
'===============================

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

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

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, vasID.MaxCols, lRow, 1, lRow + 1
    SetActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdDownLoad_Click()
    Dim intRow As Integer
    Dim j  As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                Call vasWorkList_DblClick(1, intRow)
                
'                vasID.MaxRows = vasID.MaxRows + 1
'
'                SetText vasID, "1", vasID.MaxRows, 1
'                .Col = 2
'                SetText vasID, Trim(.Text), vasID.MaxRows, 2
'
'                .Col = 4
'                SetText vasID, Trim(.Text), vasID.MaxRows, 4
'                Call GetSampleInfoW(vasID.MaxRows)                                '5,6,7,8
                
                
                'Call .DeleteRows(intRow, intRow)
                '.MaxRows = .MaxRows - 1
                '.Action = ActionDeleteRow
'                .MaxRows = .MaxRows - 1

                txtNum = txtNum + 1
                
                .Col = 1
                .Value = "0"
                
            End If
        Next
        vasID.RowHeight(-1) = 12
    End With
End Sub

Private Sub cmdExcel_Click()
    Dim iRow As Integer
    Dim j As Integer
    Dim k As Integer
    
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasExcel.MaxRows = vasRID.MaxRows
    vasExcel.MaxCols = vasRID.MaxCols
    
    For iRow = 1 To vasRID.DataRowCnt
        vasRID.Row = iRow
        vasRID.Col = 1
            
        If vasRID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasRID.MaxCols
                    SetText vasExcel, Trim(GetText(vasRID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To vasRID.MaxCols
                SetText vasExcel, Trim(GetText(vasRID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasExcel.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasExcel
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
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
    
    SetForeColor vasWorkList, 1, vasWorkList.MaxRows, 1, vasWorkList.MaxCols, 0, 0, 0
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasWorkList.MaxRows = 0
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            'If Mid(Trim(GetText(vasID, lRow, 3)), 1, 2) = "99" Then
            '    res = SaveTransDataW_QC(gRow)
            'Else
                Res = SaveTransDataW(gRow)
            'End If
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " TRANSYN = '2' " & vbCrLf & _
                      " WHERE EXAMTYPE = 'H' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdPrint_Click()
Dim iRow As Integer
    Dim j As Integer
    
    vasPrintBuf.MaxRows = 1
    vasPrintBuf.MaxRows = 2
    
    With vasID
        For iRow = 1 To .DataRowCnt
            If iRow = 1 Then
                j = 1
                
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 2)), 0, 1:    vasPrintBuf.ColWidth(1) = 14
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 5)), 0, 2
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 6)), 0, 3
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 7)), 0, 4
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 8)), 0, 5
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 9)), 0, 6
                SetText vasPrintBuf, "비고", 0, 7:                              vasPrintBuf.ColWidth(7) = 30
                
                vasPrintBuf.RowHeight(0) = 24
                
            End If
            
            

            vasPrintBuf.MaxRows = vasPrintBuf.MaxRows + 1
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 2)), iRow, 1
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 5)), iRow, 2
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 6)), iRow, 3
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 7)), iRow, 4
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 8)), iRow, 5
            SetText vasPrintBuf, Trim(GetText(vasID, iRow, 9)), iRow, 6
            SetText vasPrintBuf, " ", iRow, 7
                
            vasPrintBuf.RowHeight(iRow) = 24
                
            
        Next iRow
        
        'vasPrintbuf.RowHeight(-1) = 40

    End With
    
    If vasPrintBuf.DataRowCnt < 1 Then
        MsgBox "출력할 자료를 선택하세요", , "알 림"
        Exit Sub
    Else
        vasPrintBuf.PrintOrientation = PrintOrientationPortrait
        vasPrintBuf.Action = 13
    End If
    
    
'    vasWorkList.PrintOrientation = PrintOrientationPortrait '세로출력
'    vasWorkList.Action = 13
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
    dtpExamDate1 = Date
    
End Sub


Private Sub cmdRIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            Res = SaveTransDataR(lRow)
        
            If Res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "Failed", lRow, colState
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " TRANSYN = '2' " & vbCrLf & _
                      " WHERE EXAMTYPE = 'C' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, 3)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
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

Private Sub cmdRSch_Click()
    Dim RS1     As ADODB.Recordset
    Dim iRow    As Long
    Dim intCol  As Long
    Dim blnSame As Boolean
    Dim strDate As String
    Dim strChart As String
    Dim i As Integer
    
    ClearSpread vasRID
    ClearSpread vasRRes
    
    Call chkRAll_Click
    'Res = GetDBSelectVas(gLocal, SQL, vasRID)
    
          SQL = "SELECT RSLTDATE, BARCODE, CHARTNO, PATNAME, PATSEX, PATAGE " & vbCrLf
    SQL = SQL & "  FROM PAT_RES " & vbCrLf
    SQL = SQL & " WHERE RSLTDATE BETWEEN '" & Format(dtpExamDate, "YYYYMMDD") & "' AND '" & Format(dtpExamDate1, "YYYYMMDD") & "' " & vbCrLf
    'SQL = SQL & "   AND RESULT <> '' "
    SQL = SQL & "   AND EXAMTYPE = 'H' "
    
    If cboPartR.ListIndex = 1 Then
        SQL = SQL & "  and mid(CHARTNO,1,1) <> 'G'  "
    ElseIf cboPartR.ListIndex = 2 Then
        SQL = SQL & "  and mid(CHARTNO,1,1) = 'G' "
    End If
        
    SQL = SQL & " GROUP BY RSLTDATE, BARCODE,CHARTNO, PATNAME, PATSEX, PATAGE "
    
    cmdSQL.CommandText = SQL
    Set RS1 = cmdSQL.Execute
  
    If RS1.EOF = True Or RS1.BOF = True Then
        Exit Sub
    End If
    
    With vasRID
        While Not RS1.EOF
            For i = 1 To .DataRowCnt
                strDate = GetText(vasRID, i, 2)
                strChart = GetText(vasRID, i, 3)
                
                If Trim(RS1("RSLTDATE")) = strDate And Trim(RS1("CHARTNO")) = strChart Then
                    blnSame = True
                End If
                
            Next
            
            If blnSame = False Then
                iRow = iRow + 1
                .MaxRows = iRow
                .SetText colCheckBox, iRow, "1"
                .SetText 2, iRow, Trim(RS1.Fields("RSLTDATE").Value) & ""
                .SetText 3, iRow, Trim(RS1.Fields("BARCODE").Value) & ""
                .SetText 4, iRow, Trim(RS1.Fields("CHARTNO").Value) & ""
                .SetText 5, iRow, Trim(RS1.Fields("PATNAME").Value) & ""
                .SetText 6, iRow, Trim(RS1.Fields("PATSEX").Value) & ""
                .SetText 7, iRow, Trim(RS1.Fields("PATAGE").Value) & ""
                
            End If
            
            blnSame = False
            RS1.MoveNext
        Wend
        .RowHeight(-1) = 12
    End With
    
    Set RS1 = Nothing
    
End Sub

Private Function getResultCnt(ByVal strBarNo As String) As String

    Dim RS1     As ADODB.Recordset
    
    getResultCnt = "0"
    
    
          SQL = "SELECT COUNT(*) as CNT " & vbCrLf
    SQL = SQL & "  FROM PAT_RES " & vbCrLf
    SQL = SQL & " WHERE TRANSDT = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "   AND RESULT <> '' "
    SQL = SQL & "   AND EXAMTYPE = 'H' "
    SQL = SQL & "   AND BARCODE = '" & strBarNo & "' "
    
    cmdSQL.CommandText = SQL
    Set RS1 = cmdSQL.Execute
  
    If RS1.EOF = True Or RS1.BOF = True Then
        Exit Function
    End If
    
    While Not RS1.EOF
        getResultCnt = Trim(RS1.Fields("CNT").Value) & ""
        RS1.MoveNext
    Wend

    Set RS1 = Nothing

End Function

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            Res = SaveTransDataR(lRow)
        
            If Res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "Failed", lRow, colState
            ElseIf Res = 0 Then
            
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      "        TRANSYN = '2' " & vbCrLf & _
                      "  WHERE BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' " & _
                      "    AND EXAMTYPE = 'H' "
                      
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
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


Private Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = "C:\UBCare\SINAI\IF\ExamIF_In.xml"

    
    '1라인씩 가져오기 MSDN내용
    Dim TextLine
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    
    Next
    
'    f_subSet_XMLWorkList = Split(strTmp, "</검사>")
    strTmp = Replace(strTmp, "<검사>", ""): strTmp = Replace(strTmp, "</검사>", "|")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<요양기관번호>", ""): strTmp = Replace(strTmp, "</요양기관번호>", ",")
    strTmp = Replace(strTmp, "<차트번호>", ""): strTmp = Replace(strTmp, "</차트번호>", ",")
    strTmp = Replace(strTmp, "<수진자명>", ""): strTmp = Replace(strTmp, "</수진자명>", ",")
    strTmp = Replace(strTmp, "<주민등록번호>", ""): strTmp = Replace(strTmp, "</주민등록번호>", ",")
    strTmp = Replace(strTmp, "<내원번호>", ""): strTmp = Replace(strTmp, "</내원번호>", ",")
    strTmp = Replace(strTmp, "<의뢰일>", ""): strTmp = Replace(strTmp, "</의뢰일>", ",")
    strTmp = Replace(strTmp, "<검사번호>", ""): strTmp = Replace(strTmp, "</검사번호>", ",")
    strTmp = Replace(strTmp, "<검사ID>", ""): strTmp = Replace(strTmp, "</검사ID>", ",")
    strTmp = Replace(strTmp, "<업체검사ID>", ""): strTmp = Replace(strTmp, "</업체검사ID>", ",")
    strTmp = Replace(strTmp, "<검체>", ""): strTmp = Replace(strTmp, "</검체>", ",")
    strTmp = Replace(strTmp, "<결과치>", ""): strTmp = Replace(strTmp, "</결과치>", ",")
    strTmp = Replace(strTmp, "<참조치>", ""): strTmp = Replace(strTmp, "</참조치>", ",")
    strTmp = Replace(strTmp, "<소견>", ""): strTmp = Replace(strTmp, "</소견>", ",")
    strTmp = Replace(strTmp, "<결과일>", ""): strTmp = Replace(strTmp, "</결과일>", ",")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<입원외래구분>", ""): strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    blnSameRecord = True
    
    Kill strPath
    
    Screen.MousePointer = 0

    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


Private Function SeqSearch_New(ByVal brspread As Object, ByVal brSeq As String, ByVal brSeq2 As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch_New = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 3
                If Trim(.Text) = brSeq2 Then
                    SeqSearch_New = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            End If
        Next sCnt
    End With

End Function


Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 5
                SeqSearch = .Row
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub cmdSearch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim i, X As Long
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim FilNum
    Dim TxtString As String
    Dim TxtRece As String
    Dim PChartNum As String
    Dim PNAME As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAGE As String
    Dim PSEX As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    Dim PExamname As String
    Dim PEquipCode As String
    Dim pEqipType  As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim TxtPat As String
    Dim TestNum, IOGubun As String
    Dim FindFile As String
    Dim StartDate As String
    Dim EndDate As String
    Dim varXML      As Variant
    Dim varTmp      As Variant
    Dim strBarNo As String
    Dim intCnt As Integer
    Dim pGrid_Point As Integer
    Dim sList As Integer
    Dim strBarNum As String
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim RSX  As ADODB.Recordset
    Dim InsNo   As String
        
    Screen.MousePointer = 11
    vasWorkList.ReDraw = False
    
    ClearSpread vasWorkList

          SQL = "select distinct commdate,chartno,patname,patsex,patage,remark from pat_res "
    SQL = SQL & " where commdate between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'"
    If cboPart.ListIndex = 1 Then
        SQL = SQL & "  and mid(chartno,1,1) <> 'G'  "
    ElseIf cboPart.ListIndex = 2 Then
        SQL = SQL & "  and mid(chartno,1,1) = 'G' "
    'ElseIf cboPart.ListIndex = 3 Then
    '    SQL = SQL & "  and mid(chartno,1,1) = 'C' "
    End If


'        SQL = SQL & "   and (result = '' or result is null)"
    
    If chkSave.Value = "0" Then
        SQL = SQL & "   and result = '' "
    End If
    
    SQL = SQL & " Order by commdate,remark "
    
    
    Set RSX = cn.Execute(SQL)
    Do Until RSX.EOF
        With vasWorkList
            pGrid_Point = SeqSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)

            If pGrid_Point = 0 Then
                pGrid_Point = SeqNullSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                .RowHeight(-1) = 12
            End If
            
            .SetText 1, pGrid_Point, "1"
            .SetText 2, pGrid_Point, Format(Trim(RSX.Fields("COMMDATE")), "####-##-##")
            .SetText 3, pGrid_Point, "C"
            strBarNum = Mid(Format(Trim(RSX.Fields("COMMDATE")), "########"), 5, 4) & Format(Trim(RSX.Fields("CHARTNO")), "0000000000")
            .SetText 4, pGrid_Point, strBarNum
            .SetText 5, pGrid_Point, Trim(RSX.Fields("CHARTNO"))
            .SetText 6, pGrid_Point, Trim(RSX.Fields("PATNAME"))
            .SetText 7, pGrid_Point, Trim(RSX.Fields("PATSEX"))
            .SetText 8, pGrid_Point, Trim(RSX.Fields("PATAGE"))
            .SetText 9, pGrid_Point, "Order"
        
        End With
        RSX.MoveNext
    Loop
    RSX.Close

    '-- XML 일기
    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    If blnSameRecord = False Then
        'MsgBox "검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    End If
    
    If UBound(varXML) < 1 Then
        'MsgBox "검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarNo = ""

        With vasWorkList
            For intCnt = 0 To UBound(varXML) - 1
                varTmp = Split(varXML(intCnt), ",")
                                
                '-- 장비채널값찾기
                SQL = ""
                SQL = SQL & " SELECT EQUIPCODE "
                SQL = SQL & "   FROM EQUIPEXAM"
                SQL = SQL & "  WHERE EXAMCODE = '" & Trim(varTmp(8)) & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                XMLInData.ComExamID = ""
                
                '-- 오더 있을 경우
                If Res > 0 Then
                    
                    XMLInData.ComExamID = Trim(gReadBuf(0))
                
                    XMLInData.Company = varTmp(0)
                    XMLInData.HospCode = varTmp(1)
                    XMLInData.ChartNo = varTmp(2)
                    XMLInData.PatName = varTmp(3)
                    XMLInData.PatJumin = varTmp(4)
                    XMLInData.PatNo = varTmp(5)
                    XMLInData.CommDate = varTmp(6)
                    XMLInData.ExamNo = varTmp(7)
                    XMLInData.ExamID = varTmp(8)
                    'XMLInData.ComExamID = varTmp(9)
                    XMLInData.Specimen = varTmp(10)
                    XMLInData.Result = varTmp(11)
                    XMLInData.Reference = varTmp(12)
                    XMLInData.Remark = varTmp(13)
                    XMLInData.RsltDate = varTmp(14)
                    XMLInData.IOFlag = varTmp(15)
                    
                    
                    SQL = ""
                    SQL = SQL & "select equipno, equipcode, examname, examtype "
                    SQL = SQL & "  from equipexam "
                    SQL = SQL & " where examcode = '" & XMLInData.ExamID & "' "
                    Res = db_select_Col(gLocal, SQL)
                    If Res > 0 Then
                        PEquipno = gReadBuf(0)
                        PEquipCode = gReadBuf(1)
                        PExamname = gReadBuf(2)
                                        
                        If strBarNo <> XMLInData.ChartNo Or pEqipType <> gReadBuf(3) Then
                            pEqipType = gReadBuf(3)
                            
                            pGrid_Point = SeqSearch_New(vasWorkList, XMLInData.ChartNo, pEqipType, 5)
        
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(vasWorkList, XMLInData.ChartNo, 5)
                                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                                .RowHeight(-1) = 12
                            End If
                            
                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, Format(XMLInData.CommDate, "####-##-##")
                            .SetText 3, pGrid_Point, pEqipType
                            strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                            .SetText 4, pGrid_Point, strBarNum
                            .SetText 5, pGrid_Point, XMLInData.ChartNo
                            .SetText 6, pGrid_Point, XMLInData.PatName
                                        PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                                        Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
                            .SetText 7, pGrid_Point, gPatGen.Sex
                            .SetText 8, pGrid_Point, gPatGen.Age
                            .SetText 9, pGrid_Point, "Order"
    
                            InsNo = getMaxTestNum(XMLInData.CommDate)
                        End If
                              SQL = "Select ChartNo from pat_res "
                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                        SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                        SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                        
                        Res = db_select_Col(gLocal, SQL)
                        
                        If Res = 0 Then
                                  SQL = " insert into pat_res("
                            SQL = SQL & "Company,HospCode,ChartNo, "
                            SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
                            SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
                            SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
                            SQL = SQL & " values ("
                            SQL = SQL & "'" & XMLInData.Company & "',"
                            SQL = SQL & "'" & XMLInData.HospCode & "',"
                            SQL = SQL & "'" & XMLInData.ChartNo & "',"
                            SQL = SQL & "'" & XMLInData.PatName & "',"
                            SQL = SQL & "'" & gPatGen.Sex & "',"
                            SQL = SQL & "'" & gPatGen.Age & "',"
                            SQL = SQL & "'" & XMLInData.PatJumin & "',"
                            SQL = SQL & "'" & XMLInData.PatNo & "',"
                            SQL = SQL & "'" & XMLInData.CommDate & "',"
                            SQL = SQL & "'" & XMLInData.ExamNo & "',"
                            SQL = SQL & "'" & XMLInData.ExamID & "',"
                            SQL = SQL & "'" & XMLInData.ComExamID & "',"
                            SQL = SQL & "'" & XMLInData.Specimen & "',"
                            SQL = SQL & "'" & XMLInData.Result & "',"
                            SQL = SQL & "'" & XMLInData.Reference & "',"
                            'SQL = SQL & "'" & XMLInData.Remark & "',"
                            SQL = SQL & "'" & InsNo & "',"
                            SQL = SQL & "'" & XMLInData.RsltDate & "',"
                            SQL = SQL & "'" & XMLInData.IOFlag & "',"
                            SQL = SQL & "'" & strBarNum & "',"
                            SQL = SQL & "'" & pEqipType & "')"
                            
                            Res = SendQuery(gLocal, SQL)
                            
                            If Res = -1 Then
                                SaveQuery SQL
                            End If
                        
                        '-- 속도향상을 위해 쿼리문 지우기
                        Else
                                  SQL = " Update pat_res Set "
                            SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
                            SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
                            SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                            SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                            SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                            SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                            SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                            
                            Res = SendQuery(gLocal, SQL)
                        End If
                        
                        strBarNo = XMLInData.ChartNo
                    End If
                End If
                
                XMLInData.ComExamID = ""
            Next
            
        End With
    End If
    
    If chkSave.Value = "0" Then
        Call SaveCheck
    End If
    
    vasWorkList.ReDraw = True
    Screen.MousePointer = 0

End Sub


'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- 결과업데이트
          SQL = "SELECT MAX(REMARK) as SEQ FROM PAT_RES  "
    SQL = SQL & " WHERE MID(COMMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If Trim(gReadBuf(0)) = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(gReadBuf(0)) + 1
        End If
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
End Function

Private Sub SaveCheck()
    Dim i           As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim RS1         As ADODB.Recordset
    
    For i = 1 To vasWorkList.MaxRows
        strDate = Format(GetText(vasWorkList, i, 2), "yyyymmdd")
        strChart = GetText(vasWorkList, i, 5)
        
        '-- 기존에 있는지 조회
              SQL = "Select Count(*) from pat_res "
        SQL = SQL & " Where ChartNo  = '" & strChart & "' "
        SQL = SQL & "   and CommDate = '" & strDate & "'"
        'SQL = SQL & "   AND RSLTDATE BETWEEN '" & Format(dtpExamDate, "YYYYMMDD") & "' AND '" & Format(dtpExamDate1, "YYYYMMDD") & "' " & vbCrLf
        SQL = SQL & "   and RESULT = '' "
        
        cmdSQL.CommandText = SQL
        Set RS1 = cmdSQL.Execute
               
        If Not (RS1.EOF Or RS1.BOF) Then
            'rs.MoveFirst
        End If
        
        Do While Not RS1.EOF
            If RS1.Fields.Item(0).Value = 0 Then
                Call vasWorkList.DeleteRows(i, i)
                vasWorkList.MaxRows = vasWorkList.MaxRows - 1
            Else
                Exit Do
            End If
            RS1.MoveNext
            Exit Do
        Loop
        RS1.Close
    Next
    
End Sub

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
'    Case gServer1
'        Set cmdSQL.ActiveConnection = cn_Ser1
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS1 = cmdSQL.Execute
           
    If Not (RS1.EOF Or RS1.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        RS1.Close
        Exit Function
    End If
    
    
    Do While Not RS1.EOF
        For i = 0 To RS1.Fields.Count - 1
            If IsNull(RS1.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(RS1.Fields.Item(i).Value))
            End If
        Next i
        
        db_select_Col = 1
        
        RS1.MoveNext
        Exit Do
    Loop
    
    RS1.Close
    
    Exit Function
    
ErrHandle:
    ''MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
End Function

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, vasID.MaxCols, lRow, 1, lRow - 1
    SetActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1

End Sub

Public Function SetActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'특정 Cell 지정
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function


Private Sub cmdWDown_Click()
    Dim lRow As Long
    
    lRow = vasWorkList.ActiveRow
    
    vasWorkList.SwapRange 1, lRow, vasWorkList.MaxCols, lRow, 1, lRow + 1
    SetActiveCell vasWorkList, lRow + 1, 2
'    vasWorkList_Click 2, lRow + 1
End Sub

Private Sub cmdWPrint_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    vasPrintBuf.MaxRows = 0
    
    With vasWorkList
        For iRow = 1 To .DataRowCnt
            If iRow = 1 Then
                j = 1
                
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 2)), 0, 1:    vasPrintBuf.ColWidth(1) = 14
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 5)), 0, 2
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 6)), 0, 3
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 7)), 0, 4
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 8)), 0, 5
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 9)), 0, 6
                SetText vasPrintBuf, "비고", 0, 7:                              vasPrintBuf.ColWidth(7) = 30
                
                vasPrintBuf.RowHeight(0) = 30
                
            End If
            
            
            If GetText(vasWorkList, iRow, 1) = "1" Then
                vasPrintBuf.MaxRows = vasPrintBuf.MaxRows + 1
                
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 2)), vasPrintBuf.MaxRows, 1
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 5)), vasPrintBuf.MaxRows, 2
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 6)), vasPrintBuf.MaxRows, 3
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 7)), vasPrintBuf.MaxRows, 4
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 8)), vasPrintBuf.MaxRows, 5
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 9)), vasPrintBuf.MaxRows, 6
                SetText vasPrintBuf, " ", iRow, 7
                    
                vasPrintBuf.RowHeight(vasPrintBuf.MaxRows) = 23
            End If
            
        Next iRow
        
    End With
    
    If vasPrintBuf.DataRowCnt < 1 Then
        MsgBox "출력할 자료를 선택하세요", , "알 림"
        Exit Sub
    Else
        vasPrintBuf.PrintOrientation = PrintOrientationPortrait
        vasPrintBuf.Action = 13
    End If
    
    
End Sub

Private Sub cmdWUp_Click()
    Dim lRow As Long
    
    lRow = vasWorkList.ActiveRow
    
    vasWorkList.SwapRange 1, lRow, vasWorkList.MaxCols, lRow, 1, lRow - 1
    SetActiveCell vasWorkList, lRow - 1, 2
'    vasWorkList_Click 2, lRow - 1
End Sub

Private Sub cmdXmlDel_Click()
    Dim FindFile As String

    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
    If FindFile <> "" Then
        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
    End If

End Sub

Private Sub Command1_Click()

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblBarcode(1).Caption = ""
    lblPname(0).Caption = ""
    lblPname(1).Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
        
    strBuffer = ""
    strBuffer = strBuffer & "1H|\^&||||||||||P||05" & vbCrLf
    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||3B" & vbCrLf
    strBuffer = strBuffer & "3O|1|11208647111|807^00042^3^^SAMPLE^NORMAL|ALL|R|20111205092128|||||X||||||||||||||O|||||38" & vbCrLf
    strBuffer = strBuffer & "4R|1|^^^321^^0|>100.0|ng/ml|0.000^4.00|>||F|||20111205092406|20111205094226|CF" & vbCrLf
    strBuffer = strBuffer & "5C|1|I|51^Above measuring range|I04" & vbCrLf
    strBuffer = strBuffer & "6R|2|^^^391^^0|13.78|ng/ml|^|N||F|||20111205092448|20111205094308|14" & vbCrLf
    strBuffer = strBuffer & "7L|140" & vbCrLf
    strBuffer = strBuffer & "" & vbCrLf
    
    '-- Seq
    'strBuffer = "D 000701 6826      01201206196826    E001    24  102    15  003    46  104  7.59  005  4.96  106  0.13  007    56  108   427  009   144  110  95.4  011    47  112  8.97  013  1.08  114  3.32  015   178  116   147  017    48  118  9.66  019  2.59  120  1.08   23   101   24     2   25     8   26     1   27     3  "
    '-- 바코드
    strBuffer = "D 000201 039903073000126             E01   226H 02    85H 11   9.1  13   141H 18  13.2  21   0.7  24  20.2H "
    
    strBuffer = "DERERBDB"
    strBuffer = "R 003201 0018          1013002058"
    
    strBuffer = "D 003401 0019          1013002058    E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
    
    strBuffer = "R 000101 00011013002042"
    
    strBuffer = "D 000101 00011013002042    E012    18  017   129  018    26  "
    
    strBuffer = ""
    strBuffer = strBuffer & "R 000603 0013100825024755"
    
    
    'strBuffer = "D 003401 0019          100825024755  E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
'    strBuffer = "D 000103 0100060000034263    0001   7.8  002   4.9  003   2.9  004   1.7  005  1.25  006  0.28  007  0.97  008   117  009    87  010    71  011   250  012   123  013  11.6  014  1.04  015  11.2  016   128  017    84  018    17  019    57  020   257  097    54  "
                 
    strBuffer = "D 001408 006407102000164356    E001   7.5  002   4.7  004    12  005    14  012    80  013   8.1  014  87.2  015   4.8  017   5.1  019   9.4  099   101  "
'
'    strBuffer = "D 001201 00042000186446    E004    22  005    13  008   190  009  76.2  010  35.7  011 139.1  013   0.8  014  98.4  020 124.3  "
'
'    strBuffer = "D 001508 003806272000188274    E004      %?005      %?008      %?009      %?010      %?"
'
    strBuffer = ":n     1   1                              7  1    96   4   185   5   143   6    49  11    18  12    18  14    11 "
    
    strBuffer = ":n    17  17                              8  1   100   4   266   5   232   6    54   9  1.05  11    32  12    35  14    11 "

    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
        
    Else
        cn_Local_Flag = True
    End If
    
'    cboGbn.Clear
'    cboGbn.AddItem "전체"
'    cboGbn.AddItem "검진"
'    cboGbn.AddItem "진료"
'    cboGbn.ListIndex = 0
    
    '-- osw 추가
'''    For i = 1 To 1
'''        If Not Connect_PRServer Then
'''            'Cn_Cnt = Cn_Cnt + 1
'''            'If Cn_Cnt = 3 Then
'''            '    If Not Connect_DRServer Then
'''                    MsgBox "연결되지 않았습니다."
'''                    cn_Server_Flag = False
'''                    Exit Sub
'''             '   Else
'''             '       cn_Server_Flag = True
'''             '   End If
'''            'End If
'''        Else
'''            cn_Server_Flag = True
'''        End If
'''    Next
    
    GetExamCode
    
'    SetExamCode
    
    cboPart.AddItem "전체"
    cboPart.AddItem "외래"
    cboPart.AddItem "검진"
'    cboPart.AddItem "채용"
    cboPart.ListIndex = 0
    
    cboPartR.AddItem "전체"
    cboPartR.AddItem "외래"
    cboPartR.AddItem "검진"
    cboPartR.ListIndex = 0
    
    
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -1), "yyyymmdd")
    
    SQL = "delete from pat_res where transdt < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
'    dtpStartDt.Value = DateAdd("D", -30, Now)
    dtpStartDt.Value = Now 'DateAdd("D", -30, Now)
    dtpStopDt.Value = Now
    
    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    
    
'    Winsock1.LocalPort = CInt(5600)
'    Winsock1.Listen
    
End Sub


Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasRID
        .MaxCols = 7 + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter

            Call SetText(vasRID, gArrEquip(i + 1, 4), 0, 7 + (i + 1))
            .ColWidth(7 + (i + 1)) = 6
            
        Next
        
    End With
    
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
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
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
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    'frmTestSet.Show
    frmTestSet.Show
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
    Dim strOutput As String     '송신할 데이터
    
'''                    Case 0      'Message Header
'''                        MHead = "1H|\^&||||||||||P"
'''                        brCom.Output = STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''                        SendCount = SendCount + 1
'''                        Debug.Print "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf & Chr(13) + Chr(10);
'''                        MHead = ""
'''                    Case 1      'patient information
'''                        Pinfo = "2P|1||" & PatientID & "|||||||||||||||||||||||||||||||"
'''                        brCom.Output = STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''                        SendCount = SendCount + 1
'''                        Debug.Print "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf & Chr(13) + Chr(10);
'''                        Pinfo = ""
''''                        PatientID = ""
'''                    Case 2      'Test Order
'''                        SendCount = SendCount + 1
'''                        Call OrderingTheDataElecsys(brCom, com_sTemp, brSpread, brChannel, brItemdeci)

'''                        Orderoutput = "3O" & "|1|" & PatientID & "|" & PatientSeq & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
'''                        OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf

'''                    Case 3      'Message Terminator
'''                        SendCount = SendCount + 1
'''                        brCom.Output = STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''                        Debug.Print "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf & Chr(13) + Chr(10);
'''                    Case Else
'''                        brCom.Output = EOT
'''                        Debug.Print "[HOST] " & EOT
'''                        Print #1, "[HOST] " & EOT & Chr(13) + Chr(10);
'''                        SendCount = 0
'''                        Flag_HQL = ""
    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 4
            'strOutput = intFrameNo & "P|1|||||||||||||||||||||||||||||||||" & vbCr & ETX
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                '## 접수정보가 없을경우
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & _
                            "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
                            "|R||||||C||||||||||||||Q" & vbCr & ETX
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & _
                                "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
                                
                                '3O|1|9905300211|1^00014^1^^SAMPLE^NORMAL|ALL|R|20110613090006|||||X||||||||||||||O|||||
                                '90
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
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

Private Sub comEqp_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    Dim strDate As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Static ChkSumCnt As Long
    
    Select Case comEqp.CommEvent
        Case comEvReceive
            
            Buffer = comEqp.Input
            'Buffer = strBuffer
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        strBuffer = ""
                        
                    Case ETX
                         Call EditRcvData
                         strBuffer = ""
                    
                    Case Else
                        strBuffer = strBuffer & BufChar
                End Select
            Next i

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


Public Sub sndMore()
    Dim strSndMsg As String
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Public Sub sndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) & GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & Chr(13)
    
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
        If Trim(GetText(vasID, i, colBarNo)) = pBarNo Then
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
    
    Call SetText(vasID, "Order", intRow, colState)
    Call SetText(vasID, pBarNo, intRow, colBarNo)
    Call vasActiveCell(vasID, intRow, colBarNo)
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                            '5,6,7,8
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
    
    '-- 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.
    '-- intRow 추가
'    strItems = GetGetEquipExamCode_AU480(gEquip, pBarNo, intRow)
'
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = ""
'        'S 003401 0019          1013001918    E
'        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'
'        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'        'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E012" & ETX
'
'
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'
'    End If

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
        If Trim(GetText(vasID, i, colPID)) = pBarNo Then
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
    
    If Mid(pBarNo, 1, 2) = "QC" Then
        Call SetText(vasID, Format(Now, "yyyy-mm-dd"), intRow, 2)
        Call SetText(vasID, pBarNo, intRow, colPID)
        Call SetText(vasID, "QC", intRow, colPName)
    Else
        'Call SetText(vasID, pBarNo, intRow, colBarNo)             '2 Barcode
    End If
    
    Call vasActiveCell(vasID, intRow, colPID)
    
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                                '5,6,7,8
    
    gRow = intRow
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strQCResult  As String   '수신한 결과(QC)
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
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim strSndBuffer As String
    Dim intCol As Integer
    
    Dim strFunc       As String
    Dim strFunction   As String
    Dim strDate       As String
    Dim strSendData   As String
    Dim strSndMsg     As String
    Dim strExamCode() As String
    
    Dim i             As Integer
    Dim varBuffer As Variant
    
'    strRcvBuf = strBuffer
'
'    strType = Mid$(strRcvBuf, 1, 1)
'    If IsNumeric(strType) Then
'        strType = Mid$(strRcvBuf, 2, 1)
'    End If
    
    
    
    strBuffer = Replace(strBuffer, vbLf, "")
    varBuffer = Split(strBuffer, vbCr)
    
    If Trim(varBuffer(0)) = "EXP" Then
        Exit Sub
    End If
    
    For intCnt = 1 To UBound(varBuffer)
        strRcvBuf = varBuffer(intCnt)
        Select Case intCnt
            Case 17
                strBarNo = Trim(strRcvBuf)
                
                'MsgBox strBarNo
                gRow = 0
                '-- 포지션 번호로 바코드 번호 찾기
                For i = 1 To vasID.DataRowCnt
                    If Val(Trim(GetText(vasID, i, colOCnt))) = Val(strBarNo) Then
                        strBarNo = Trim(GetText(vasID, i, colPID))
                        'MsgBox strBarNo
                        gRow = i
                        Exit For
                    End If
                Next
                
                If strBarNo = "" Or gRow <= 0 Then Exit Sub
                
                With mResult
                    .BarNo = Trim(strBarNo)
                    '.Seq = Trim(strSeq)
                    '.RackNo = Trim(strRackNo)
                    '.TubePos = Trim(strTubePos)
                End With
            
                Call SetPatInfo(strBarNo)

                strState = "O"
                
                '-- 오른쪽 결과화면 초기화
                vasRes.MaxRows = 0
                        
            Case 18 To 49
                strIntBase = intCnt
                strResult = Trim(Mid(strRcvBuf, 1, 4))
                
                If Left(strResult, 1) = "." Then
                    strResult = "0" & strResult
                End If
                
                strFlag = Mid(strRcvBuf, 5, 1)
                
                If strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH "
                    SQL = SQL & "  FROM EQUIPEXAM"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                    
                    Res = GetDBSelectColumn(gLocal, SQL)
                    
                    '-- 오더 있을 경우
                    If Res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        '-- Work List
                        SetText vasID, "1", gRow, colCheckBox                  '체크
                        'SetText vasID, strResult, gRow, colA1c                  '결과
                        'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                        SetText vasID, "Result", gRow, colState                 '진행상태
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                        
                        
                        SetText vasRes, strResult, lsResRow, colResult          '결과
                        
                        vasRes.Row = lsResRow
                        vasRes.Col = colResult
                        vasRes.FontBold = False
                        vasRes.ForeColor = vbBlack
                        
                        
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        'SetText vasRes, strComm, lsResRow, 7                    'Flag
                        
                        '-- 로컬 저장
                        SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                        
                        lsResult_Buff = ""
                        strState = "R"
                        
                    '-- 오더 없을 경우
                    Else
                    
                              SQL = "Select examcode, examname, seqno ,REFLOW,REFHIGH "
                        SQL = SQL & "  From equipexam"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        SQL = SQL & "   and examtype = 'H' "
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '소수점 처리, 결과 형태 처리
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            'SetText vasID, "0", gRow, colCheckBox                  '체크
                            'SetText vasID, strResult, gRow, colA1c                  '결과
                            'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            SetText vasID, "Result", gRow, colState                 '진행상태
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                                                            
                            SetText vasRes, strResult, lsResRow, colResult          '결과
                            
                            vasRes.Row = lsResRow
                            vasRes.Col = colResult
                            vasRes.FontBold = False
                            vasRes.ForeColor = vbBlack
    
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                
                vasRes.RowHeight(-1) = 14
            
            Case 50
                
                If MnTransAuto.Checked = True And strState = "R" Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- 저장 성공
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        
                        SQL = " Update pat_res Set " & vbCrLf & _
                              "  transdt = '" & Format(Now, "yyyymmdd") & "', " & vbCrLf & _
                              "  transyn = '2' " & vbCrLf & _
                              "  Where barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' " & _
                              "    and examtype = 'H'"
                              
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                        SetText vasID, "0", gRow, colCheckBox
            
                    End If
                    strState = ""
                End If
        End Select
    Next
    
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
    Res = GetDBSelectColumn(gLocal, SQL)
    
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
        If IsNumeric(sEquipRes) Then
            If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
                sResFlag = ""
            ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
    
    gsFlag = sResFlag
    SetResult = sResult
    
End Function

' asRow1 = Work List
' asRow2 = 결과 List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Format(dtpToday, "yyyymmdd")


'    SQL = ""
'    SQL = SQL & "UPDATE PAT_RES SET "
'    SQL = SQL & " RESULT = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf
'    SQL = SQL & " RSLTDATE = '" & Format(CDate(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf
'    SQL = SQL & " WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, 4)) & "' " & vbCrLf
'    SQL = SQL & "   AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMTYPE = 'H'"

    SQL = ""
    SQL = "DELETE FROM PAT_RES " & vbCrLf
    SQL = SQL & "WHERE RSLTDATE = '" & Format(Now, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colPID)) & "' " & vbCrLf
    SQL = SQL & "  AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    SQL = SQL & "  AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID,EXAMTYPE) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'H',"
'    SQL = SQL & "'" & gIFUser & "')"
'
    
          SQL = " insert into pat_res("
    SQL = SQL & "Company,HospCode,ChartNo, "
    SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
    SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
    SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
    SQL = SQL & " values ("
    SQL = SQL & "'ACK',"
    SQL = SQL & "'41343051',"
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "',"
    SQL = SQL & "'QC',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Format(Now, "yyyymmdd") & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Format(Now, "yyyymmdd") & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'H')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function SetLocalDB_New(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Format(dtpToday, "yyyymmdd")

'    SQL = ""
'    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'" & gIFUser & "')"
    
    SQL = ""
    SQL = SQL & "UPDATE PAT_RES SET "
    SQL = SQL & " RESULT = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf
    SQL = SQL & " RSLTDATE = '" & Format(CDate(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf
    SQL = SQL & " WHERE CHARTNO = '" & Trim(GetText(vasID, asRow1, 5)) & "' " & vbCrLf
    SQL = SQL & "   AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    SQL = SQL & "   AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf
    SQL = SQL & "   AND EXAMTYPE = 'H'"
'    SQL = SQL & " WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, 4)) & "' " & vbCrLf
'    SQL = SQL & "   AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMTYPE = 'H'"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
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



Private Sub txtNum_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer

    If KeyAscii = 13 Then
        With vasID
            For intRow = .ActiveRow To .DataRowCnt
                SetText vasID, txtNum, intRow, colOCnt 'colSpecNo

                txtNum = Val(txtNum) + 1
            Next
        End With
    End If
End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 0 Then
        vasID.Value = 1
        Else
        vasID.Value = 0
        End If
    Next i
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim strChart As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    strChart = Trim(GetText(vasID, Row, 5))
    lsID = Trim(GetText(vasID, Row, 4))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    lblBarcode(0).Caption = strChart 'lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.EXAMNO, a.TRANSYN, b.reflow,b.refhigh, b.seqno " & vbCrLf
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b " & vbCrLf
    SQL = SQL & " WHERE a.EXAMID = b.EXAMCODE " & vbCrLf
    SQL = SQL & "   AND a.COMEXAMID = b.EQUIPCODE " & vbCrLf
    'SQL = SQL & "   AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND a.BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND a.CHARTNO = '" & strChart & "' " & vbCrLf
    SQL = SQL & "   AND a.EXAMTYPE = 'H' " & vbCrLf
    SQL = SQL & " GROUP BY a.EXAMNO, a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.TRANSYN, b.reflow,b.refhigh, b.seqno "
    SQL = SQL & " ORDER BY b.seqno * 10"
    Res = GetDBSelectVas_Ref(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    Else
        vasRes.RowHeight(-1) = 12
    End If

    vasRes.MaxRows = vasRes.DataRowCnt

End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsNM As String
    Dim lsPid As String
    Dim i As Integer

    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If

        lsID = Trim(GetText(vasID, iRow, colBarcode))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsNM = Trim(GetText(vasID, iRow, colPName))
        
        If MsgBox(lsNM & "을 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0

        vasID.MaxRows = vasID.MaxRows - 1
        
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
'    res = GetDBSelectColumn(gLocal, SQL)
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
    Dim lsBarcode As String
    Dim i As Integer
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsBarcode = Trim(GetText(vasRID, Row, 3))
    lsID = Trim(GetText(vasRID, Row, 4))
    lblChangeBar.Caption = lsID
    lblBarcode(1).Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
    lblPname(1).Caption = Trim(GetText(vasRID, Row, colPName))
    lblRrow.Caption = Row
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.EXAMNO, a.TRANSYN, b.reflow,b.refhigh " & vbCrLf
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b " & vbCrLf
    SQL = SQL & " WHERE a.EXAMID = b.EXAMCODE " & vbCrLf
    SQL = SQL & "   AND a.COMEXAMID = b.EQUIPCODE " & vbCrLf
    SQL = SQL & "   AND a.BARCODE = '" & lsBarcode & "' " & vbCrLf
    SQL = SQL & "   AND a.CHARTNO = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND a.EXAMTYPE = 'H' " & vbCrLf
    SQL = SQL & " GROUP BY a.EXAMNO, a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.TRANSYN, b.reflow,b.refhigh "
    
    Res = GetDBSelectVas_Ref(gLocal, SQL, vasRRes)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    Else
        vasRRes.RowHeight(-1) = 12
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

'Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasRID.ActiveRow
'
'    If KeyCode = 13 Then
'
'        If GetSampleInfoR(iRow) = -1 Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = ""
'        SQL = SQL & "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf
'        SQL = SQL & "  FROM PAT_RES " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' "
'        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
'        SQL = SQL & "   AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf
'        SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
'
'        Res = GetDBSelectVas(gLocal, SQL, vasTemp)
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'                SQL = ""
'                SQL = SQL & "INSERT INTO PAT_RES("
'                SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                            "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                            "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'                SQL = SQL & "VALUES("
'                SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'                SQL = SQL & "'" & gEquip & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colBarcode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colDISK)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPos)) & "', " & vbCrLf
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPID)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPName)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colSex)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colAge)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colEquipCode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamCode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colSeq)) & "', " & vbCrLf
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colMachResult)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colResult)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamName)) & "', "
'                SQL = SQL & "'0', "
'                SQL = SQL & "'" & gIFUser & "')"
'
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Next i
'
'            SQL = ""
'            SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
'            SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
'            SQL = SQL & "   AND BARCODE  = '" & lblChangeBar.Caption & "' " & vbCrLf
'            SQL = SQL & "   AND PID      = '" & lblChangePID.Caption & "' " & vbCrLf
'            SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'            SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'            SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'            Res = SendQuery(gLocal, SQL)
'
'            If Res = -1 Then
'                SaveQuery SQL
'                Exit Sub
'            End If
'
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'
'                SQL = ""
'                SQL = SQL & "UPDATE PAT_RES " & vbCrLf
'                SQL = SQL & "   SET RESULT    ='" & Trim(GetText(vasRRes, i, colResult)) & "' " & vbCrLf
'                SQL = SQL & " WHERE BARCODE   = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf
'                SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
'                SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf
'                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, colEquipCode)) & "' " & vbCrLf
'                SQL = SQL & "   AND PID       = '" & Trim(GetText(vasRID, iRow, colPID)) & "' " & vbCrLf
'                SQL = SQL & "   AND DISKNO    = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'                SQL = SQL & "   AND POSNO     = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'                SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Next i
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasRID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = ""
'        SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
'        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
'        SQL = SQL & "   AND PID      = '" & lsPid & "' " & vbCrLf
'        SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'        SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'        SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'        Res = SendQuery(gLocal, SQL)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasRID, iRow, iRow
'        vasRRes.MaxRows = 0
'
'    End If
'End Sub

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub



Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim pGrid_Point As Integer
    Dim sBarcode As String
    Dim sChartNo As String
    
    If Row = 0 Then Exit Sub
    
    With vasWorkList
        .Col = Col
        .Row = Row
        .Col = 5
        pGrid_Point = SeqSearch(vasID, Trim(.Text), 5)

        If pGrid_Point = 0 Then
            pGrid_Point = SeqNullSearch(vasID, Trim(.Text), 5)
            If pGrid_Point = 0 Then vasID.MaxRows = vasID.MaxRows + 1: pGrid_Point = vasID.MaxRows
            .RowHeight(-1) = 12
        End If
        
        .Row = Row: .Col = 4
        sBarcode = Trim(.Text)
        
        Call vasID.SetText(1, pGrid_Point, "1")
        Call vasID.SetText(4, pGrid_Point, .Text)
'        .Row = Row: .Col = 5
'        Call vasID.SetText(5, pGrid_Point, .Text)
'        .Row = Row: .Col = 6
'        Call vasID.SetText(6, pGrid_Point, .Text)
'        .Row = Row: .Col = 7
'        Call vasID.SetText(7, pGrid_Point, .Text)
'        .Row = Row: .Col = 8
'        Call vasID.SetText(8, pGrid_Point, .Text)
'        vasID.RowHeight(-1) = 12
    
        '바코드번호로 환자정보 불러오기
              SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
        SQL = SQL & vbCrLf & "  FROM PAT_RES "
        SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'H' "
        SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
        
    
        Res = GetDBSelectColumn(gLocal, SQL)
            
        If Res = 1 Then
            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
            SetText frmInterface.vasID, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
            SetText frmInterface.vasID, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
            SetText frmInterface.vasID, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
            SetText frmInterface.vasID, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
            
            SetText frmInterface.vasID, Trim(gReadBuf(4)), pGrid_Point, 12
            SetText frmInterface.vasID, Trim(gReadBuf(5)), pGrid_Point, 13
            SetText frmInterface.vasID, Trim(gReadBuf(6)), pGrid_Point, 14
            SetText frmInterface.vasID, Trim(gReadBuf(7)), pGrid_Point, 15
            SetText frmInterface.vasID, Trim(gReadBuf(8)), pGrid_Point, 16
            SetText frmInterface.vasID, Trim(gReadBuf(9)), pGrid_Point, 17
            SetText frmInterface.vasID, Trim(gReadBuf(10)), pGrid_Point, 18
            SetText frmInterface.vasID, Trim(gReadBuf(11)), pGrid_Point, 19
            vasID.RowHeight(-1) = 12
        End If
    
    End With
    
End Sub

'Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'
'    If Winsock1.State <> sckClosed Then
'        Winsock1.Close
'
'        Winsock1.Accept requestID
'        StatusBar1.Panels(2).Text = "장비에 접속되었습니다."
'    End If
'
'End Sub
'
'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'    Dim strText As String
'    Dim strTmp As String
'
'    Dim strLastSeq  As String
'    Dim strRcvSign  As String
'    Dim strSendAck  As String
'    Dim strRcvCnt   As String
'
'    Dim strNS       As String
'    Dim strNE       As String
'    Dim intNS       As Integer
'    Dim intNE       As Integer
'
'    Dim strSendData  As String
'    Dim varBuffers   As Variant
'    Dim i As Integer
'    Dim lngBufLen As Long
'    Dim BufChar     As String
'
'    Winsock1.GetData strText
'
'        strBuffer = strText
'
'    SetRawData "[Rx]" & strBuffer
'    StatusBar1.Panels(3).Text = strBuffer
'
'    strBuffer = Replace(strBuffer, vbLf, "")
'
'    strRecvData = Split(strBuffer, vbCr)
'
'    Call EditRcvData
'
'
'End Sub
Private Sub vasWorkList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsNM As String
    Dim lsPid As String
    Dim i As Integer

    iRow = vasWorkList.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasWorkList.DataRowCnt Then
            Exit Sub
        End If

        lsID = Trim(GetText(vasWorkList, iRow, colBarcode))
        lsPid = Trim(GetText(vasWorkList, iRow, colPID))
        lsNM = Trim(GetText(vasWorkList, iRow, colPName))
        
        If MsgBox(lsNM & "을 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow vasWorkList, iRow, iRow
        vasRes.MaxRows = 0

        vasWorkList.MaxRows = vasWorkList.MaxRows - 1
        
    End If
End Sub
