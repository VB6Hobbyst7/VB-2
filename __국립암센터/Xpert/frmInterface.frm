VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   Caption         =   "Xpert Interface Program"
   ClientHeight    =   9300
   ClientLeft      =   2685
   ClientTop       =   1200
   ClientWidth     =   15285
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   15285
   Begin VB.TextBox Text1 
      Height          =   1995
      Left            =   3930
      MultiLine       =   -1  'True
      TabIndex        =   61
      Top             =   4140
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "test"
      Height          =   375
      Left            =   11280
      TabIndex        =   60
      Top             =   5700
      Visible         =   0   'False
      Width           =   975
   End
   Begin FPSpread.vaSpread vasReport 
      Height          =   7065
      Left            =   420
      TabIndex        =   68
      Top             =   1980
      Visible         =   0   'False
      Width           =   9855
      _Version        =   393216
      _ExtentX        =   17383
      _ExtentY        =   12462
      _StockProps     =   64
      BorderStyle     =   0
      ColHeaderDisplay=   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   60
      SpreadDesigner  =   "frmInterface.frx":628A
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vasSortInfo 
      Height          =   2115
      Left            =   12600
      TabIndex        =   67
      Top             =   4140
      Visible         =   0   'False
      Width           =   1695
      _Version        =   393216
      _ExtentX        =   2990
      _ExtentY        =   3731
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
      MaxCols         =   1
      MaxRows         =   5
      SpreadDesigner  =   "frmInterface.frx":F157
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   675
      Left            =   1410
      TabIndex        =   65
      Top             =   2490
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   180
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vasLogRes 
      Height          =   1425
      Left            =   1440
      TabIndex        =   59
      Top             =   8100
      Visible         =   0   'False
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
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
      MaxCols         =   40
      SpreadDesigner  =   "frmInterface.frx":F4C8
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   4275
      Left            =   2130
      TabIndex        =   56
      Top             =   2610
      Visible         =   0   'False
      Width           =   11265
      _Version        =   393216
      _ExtentX        =   19870
      _ExtentY        =   7541
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      EditEnterAction =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":110C5
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9450
      Top             =   9990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "test_이거안씀"
      Height          =   375
      Left            =   6990
      TabIndex        =   50
      Top             =   7290
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtData 
      Height          =   555
      Left            =   1050
      TabIndex        =   49
      Top             =   7260
      Visible         =   0   'False
      Width           =   5925
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   60
      TabIndex        =   33
      Top             =   -60
      Width           =   15105
      Begin Threed.SSPanel SSPanel1 
         Height          =   855
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   1508
         _Version        =   131072
         ForeColor       =   8388736
         BackColor       =   16056319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    Xpert Interface"
         BevelOuter      =   0
         Alignment       =   1
         Begin FPSpread.vaSpread vasModuleCnt 
            Height          =   495
            Left            =   2280
            TabIndex        =   63
            Top             =   300
            Width           =   3915
            _Version        =   393216
            _ExtentX        =   6906
            _ExtentY        =   873
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
            MaxCols         =   4
            MaxRows         =   1
            ScrollBars      =   0
            SpreadDesigner  =   "frmInterface.frx":14FFB
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7320
            Picture         =   "frmInterface.frx":15392
            ScaleHeight     =   255
            ScaleWidth      =   285
            TabIndex        =   57
            Top             =   -30
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.ComboBox cboExam 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9540
            TabIndex        =   42
            Text            =   "전체선택"
            Top             =   -120
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.TextBox txtRemark 
            Height          =   435
            Left            =   2220
            TabIndex        =   39
            Top             =   1230
            Width           =   1545
         End
         Begin Xpert_국립암센터.MDButton cmdWorkList 
            Height          =   495
            Left            =   11310
            TabIndex        =   35
            Top             =   -210
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "접수조회"
         End
         Begin Xpert_국립암센터.MDButton cmdSch 
            Height          =   555
            Left            =   10530
            TabIndex        =   36
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과검색"
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   345
            Left            =   7290
            TabIndex        =   37
            Top             =   240
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
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
            Format          =   95289345
            CurrentDate     =   38584
         End
         Begin VB.CheckBox chkMode 
            Caption         =   "AUTO"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   585
            Left            =   9540
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   600
            Value           =   1  '확인
            Visible         =   0   'False
            Width           =   825
         End
         Begin Xpert_국립암센터.MDButton cmdClose 
            Height          =   555
            Left            =   13890
            TabIndex        =   51
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "종료"
         End
         Begin Xpert_국립암센터.MDButton cmdClear 
            Height          =   555
            Left            =   12750
            TabIndex        =   52
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "화면정리"
         End
         Begin Xpert_국립암센터.MDButton cmd_Trans 
            Height          =   555
            Left            =   11640
            TabIndex        =   53
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과전송"
         End
         Begin Xpert_국립암센터.MDButton cmdPrint 
            Height          =   555
            Left            =   9420
            TabIndex        =   66
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "출력"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Module별 검사횟수"
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
            Left            =   2280
            TabIndex        =   64
            Top             =   60
            Width           =   1965
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  '투명
            Caption         =   "사용자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   7740
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '투명
            Caption         =   "검사항목"
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
            Left            =   8550
            TabIndex        =   43
            Top             =   -60
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "_국립암센터 핵의학과"
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
            Left            =   690
            TabIndex        =   41
            Top             =   450
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Label Label6 
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
            Height          =   195
            Left            =   6330
            TabIndex        =   40
            Top             =   330
            Width           =   900
         End
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      OutBufferSize   =   1024
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   60
      Top             =   1590
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      Height          =   285
      Left            =   720
      TabIndex        =   32
      Top             =   1020
      Width           =   225
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   8265
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   14579
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   66
      Protect         =   0   'False
      SpreadDesigner  =   "frmInterface.frx":1591C
   End
   Begin VB.Frame frameSch 
      Height          =   8655
      Left            =   60
      TabIndex        =   27
      Top             =   930
      Visible         =   0   'False
      Width           =   15075
      Begin VB.CommandButton cmdSchClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   13440
         TabIndex        =   28
         Top             =   8220
         Width           =   1485
      End
      Begin FPSpread.vaSpread vasSch 
         Height          =   7815
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   14805
         _Version        =   393216
         _ExtentX        =   26114
         _ExtentY        =   13785
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         Protect         =   0   'False
         SpreadDesigner  =   "frmInterface.frx":1AB60
      End
   End
   Begin VB.TextBox txtBuff 
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   4965
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   3300
      Visible         =   0   'False
      Width           =   915
      _Version        =   393216
      _ExtentX        =   1614
      _ExtentY        =   1296
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
      SpreadDesigner  =   "frmInterface.frx":1C89B
   End
   Begin VB.Frame Frame1 
      Height          =   6435
      Left            =   3060
      TabIndex        =   7
      Top             =   1890
      Visible         =   0   'False
      Width           =   10995
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2520
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "◀"
         Height          =   465
         Left            =   240
         TabIndex        =   21
         Top             =   5190
         Width           =   645
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "▶"
         Height          =   465
         Left            =   900
         TabIndex        =   20
         Top             =   5190
         Width           =   645
      End
      Begin VB.TextBox txtRack 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2970
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtTube 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3390
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtResDate 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2070
         Width           =   2475
      End
      Begin VB.TextBox txtPName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1275
         Width           =   1545
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   825
         Width           =   1545
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   390
         Width           =   1545
      End
      Begin FPSpread.vaSpread vasRes1 
         Height          =   5865
         Left            =   2940
         TabIndex        =   8
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   10345
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":1CB3F
      End
      Begin FPSpread.vaSpread vasRes2 
         Height          =   5865
         Left            =   6840
         TabIndex        =   9
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   10345
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":1D1DF
      End
      Begin Threed.SSCommand cmdCloseDetail 
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         Top             =   5175
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         _Version        =   131072
         Caption         =   "닫기"
         ButtonStyle     =   2
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "W/L No"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2580
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   24
         Top             =   3030
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Tube"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   3450
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과시간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   1785
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "등록번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   885
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   450
         Width           =   840
      End
   End
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   30
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   705
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   9600
      Visible         =   0   'False
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   1085
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      Begin MSComDlg.CommonDialog cdFindFile 
         Left            =   10170
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Xpert_국립암센터.MDButton cmdResCall 
         Height          =   435
         Left            =   5550
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "결과파일 불러오기"
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   495
         Left            =   90
         TabIndex        =   44
         Top             =   90
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   873
         _Version        =   131072
         ForeColor       =   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   1
         Begin Xpert_국립암센터.MDButton cmdResDel 
            Height          =   345
            Left            =   3330
            TabIndex        =   48
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과삭제"
         End
         Begin VB.TextBox txtESeq 
            Appearance      =   0  '평면
            Height          =   315
            Left            =   1950
            TabIndex        =   46
            Top             =   90
            Width           =   1185
         End
         Begin VB.TextBox txtSSeq 
            Appearance      =   0  '평면
            Height          =   315
            Left            =   480
            TabIndex        =   45
            Top             =   90
            Width           =   1185
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1740
            TabIndex        =   47
            Top             =   120
            Width           =   225
         End
      End
      Begin Threed.SSPanel sspPort 
         Height          =   495
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Visible         =   0   'False
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   873
         _Version        =   131072
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   COBRA II 장비"
         BevelOuter      =   1
         Alignment       =   1
         Begin VB.Label lblCA 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "연결"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2430
            TabIndex        =   6
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lblCACom 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[COM1]9600,n,8,1"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2910
            TabIndex        =   5
            Top             =   150
            Width           =   1530
         End
      End
      Begin Xpert_국립암센터.MDButton cmdResPrint 
         Height          =   435
         Left            =   7920
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "결과출력"
      End
      Begin VB.Label lblIPState 
         Caption         =   "대기중"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13470
         TabIndex        =   62
         Top             =   210
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5730
         TabIndex        =   31
         Top             =   180
         Width           =   60
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일"
      Begin VB.Menu subClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu subN1 
         Caption         =   "-"
      End
      Begin VB.Menu subClose 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "연결"
      Begin VB.Menu subSendMode 
         Caption         =   "서버 결과 전송"
         Begin VB.Menu subSend1 
            Caption         =   "Auto"
         End
         Begin VB.Menu subSend2 
            Caption         =   "Manual"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "설정"
      Begin VB.Menu subCodeSet 
         Caption         =   "코드설정"
      End
      Begin VB.Menu subComSetup 
         Caption         =   "통신설정"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "검색"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheck = 1
Const colBarcode = 2
Const colPID = 3
Const colPName = 4
Const colReceNo = 5
Const colRack = 6
Const colPos = 7
Const colState = 8
Const colExamCode = 9
Const colExamName = 10
Const colResIC = 11
Const colResTarget = 12

'결과들
Const colResult = 13

Const colMTBPA = 16
Const colMTBPB = 17
Const colMTBPC = 18
Const colMTBPD = 19
Const colMTBPE = 20

Const colRifPA = 21
Const colRifPB = 22
Const colRifPC = 23
Const colRifPD = 24
Const colRifPE = 25

Const colRemark = 26

Const colMTBSPC = 27
Const colRifSPC = 28

Const colStartDate = 29
Const colEndDate = 30
Const colCartNo = 31
Const colReagentNo = 32
Const colExpDate = 33
Const colError1 = 34
Const colError2 = 35

Const colQC1Ct = 36
Const colQC2Ct = 37

Const colProAPt = 38
Const colProBPt = 39
Const colProCPt = 40
Const colProDPt = 41
Const colProEPt = 42
Const colSPCPt = 43
Const colQC1Pt = 44
Const colQC2Pt = 45

Const colProARes = 46
Const colProBRes = 47
Const colProCRes = 48
Const colProDRes = 49
Const colProERes = 50
Const colSPCRes = 51
Const colQC1Res = 52
Const colQC2Res = 53

Const colAssay = 54

Const colSexAge = 55
Const colMedDept = 56
Const colIO = 57
Const colTestDate = 58

Const colProACheck = 59
Const colProBCheck = 60
Const colProCCheck = 61
Const colProDCheck = 62
Const colProECheck = 63
Const colSPCCheck = 64
Const colQC1Check = 65
Const colQC2Check = 66

Dim gEquipCode As String


Const colReceDate = 14
Const colEquipCode = 15

Dim gCurRow As Long
Dim gMaxCol As Long

Dim iRow1 As Long
Dim iRow2 As Long
Dim iCol1 As Long
Dim iCol2 As Long

Dim PreRack As String
Dim PrePos As String
Dim PreRow As Long

Dim SelVas As Integer
Dim sSampleType As String
Dim strToxCheck As String


'''Public colState As String

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 20

    db_tmp = ""
    
    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.db = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("OPTION", "InsCode", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gInsCode = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("OPTION", "WkCode", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gWkCode = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Data", "WorkListExpire", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDays = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("FTP", "FTPServer", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gFTPConf.Server = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("FTP", "FTPPort", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gFTPConf.Port = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("FTP", "FTPUser", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gFTPConf.User = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("FTP", "FTPPassWD", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gFTPConf.Passwd = Trim(txtTemp)
    
    '-- Winsock 관련
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerIP = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerPort = Trim(txtTemp)
    
    GetSetup = True

End Function


Private Sub chkAll_Click()
    vasList.Row = -1
    vasList.Col = 1
    
    If chkAll.Value = 0 Then
        vasList.Value = 0
    Else
        vasList.Value = 1
    End If
End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
        SaveSetting "MEDIMATE", "COBRA", "SendMode", "1"
    Else
        chkMode.Caption = "Manual"
        SaveSetting "MEDIMATE", "COBRA", "SendMode", "0"
    End If
End Sub

Private Sub cmd_Trans_Click()
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim lsResult As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liRet As Integer
    Dim lsExamCode As String
    Dim sParam As String
    Dim sResFlag As String
    Dim sResult1 As String
    Dim strModule As String
    
    Dim srtComment As String
    
    
    If MsgBox(" " & vbCrLf & "검사 결과를 전송하시겠습니까?" & vbCrLf & " ", vbInformation + vbYesNo + vbDefaultButton2, "결과 전송 알림") = vbNo Then
        Exit Sub
    End If
    
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        
        If vasList.Value = 1 Then
            lsID = Trim(GetText(vasList, lRow, colBarcode))
            lsExamCode = Trim(GetText(vasList, lRow, colExamCode))
            
            If lsExamCode = "L2742" Then
                lsResult = Trim(GetText(vasList, lRow, colResult))
            Else
                lsResult = ""
                If InStr(1, GetText(vasList, lRow, colProARes), "NEG") > 0 Then
                    lsResult = "Toxin B : NEGATIVE"
                ElseIf InStr(1, GetText(vasList, lRow, colProARes), "POS") > 0 Then
                    lsResult = "Toxin B : POSITIVE"
                End If
                
                If GetText(vasList, lRow, colProBRes) = "NEG" Then
                    lsResult = lsResult & "/Binary Toxin : NEGATIVE"
                ElseIf GetText(vasList, lRow, colProBRes) = "POS" Then
                    lsResult = lsResult & "/Binary Toxin : POSITIVE"
                End If
                
                If GetText(vasList, lRow, colProCRes) = "NEG" Then
                    lsResult = lsResult & "/TcdC : NEGATIVE"
                ElseIf GetText(vasList, lRow, colProCRes) = "POS" Then
                    lsResult = lsResult & "/TcdC : POSITIVE"
                End If
                
                lsResult = lsResult & "//<결과해석>"
                lsResult = lsResult & "/1) Toxin B (+), Binary toxin (-), tcdC (-) : Toxin B를 분비하는 일반 C.diffcile 균주"
                lsResult = lsResult & "/2) Toxin B (+), Binary toxin 또는 tcdC 둘 중 하나 (+) : Toxin B를 분비하는 일반 C.diffcile 균주"
                lsResult = lsResult & "/   - Binary toxin (+) : toxin의 활성을 촉진하므로 적극적인 치료가 요구됨"
                lsResult = lsResult & "/   - tcdC (+) : toxin B의 분비가 약 23배 증가됨이 보고됨"
                lsResult = lsResult & "/3) Toxin B (+), Binary toxin (+), tcdC (+) : 고병원성 ribotype 027균주"
                lsResult = lsResult & "/   - Birnary toxin을 분비하면서 tcdC 유전자가 결손된 강한 독성을 보유한 C. difficile 균주"
                
            End If
            
            
            
            srtComment = Trim(GetText(vasList, lRow, colRemark))
            strModule = Trim(GetText(vasList, lRow, colRack))
            
            If lsResult <> "" And Len(lsID) > 10 Then
            
                
                
                sParam = ""
                sParam = sParam & "<Table>" & _
                        "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                        "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                        "<USERID><![CDATA[LIA]]></USERID>" & _
                        "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                        "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                        "<P0><![CDATA[" & lsID & "]]></P0>" & _
                        "<P1><![CDATA[" & lsExamCode & "]]></P1>" & _
                        "<P2><![CDATA[" & Replace(lsResult, "/", vbCrLf) & "]]></P2>" & _
                        "<P3><![CDATA[]]></P3>" & _
                        "<P4><![CDATA[" & gEquip & Mid(strModule, 2, 1) & "]]></P4>" & _
                        "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                        "<P6><![CDATA[]]></P6>" & _
                        "<P7><![CDATA[" & Replace(srtComment, "/", vbCrLf) & "]]></P7>" & _
                        "<P8><![CDATA[]]></P8>" & _
                        "<P9><![CDATA[]]></P9>" & _
                        "</Table>"
            
                sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
                
                Online_Result_Qry sParam
        
                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                SetText vasList, "완료", lRow, 8
                
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                
                SQL = "update pat_res set sendflag = 'C' where barcode = '" & lsID & "' and examcode = '" & lsExamCode & "'"
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                End If
            End If
            
        End If
    Next lRow
End Sub

Private Sub cmdClear_Click()
    subClear_Click
    ExamCount
End Sub

Private Sub cmdCode_Click()
    frmCode.Show 1
End Sub

Private Sub cmdComSetup_Click()
    frmConfig.Show 1
End Sub
Private Sub cmdClose_Click()
    subClose_Click
End Sub

Private Sub cmdCloseDetail_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdNext_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 + 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    txtRack = Trim(GetText(argSpread, lRow1, 6))
    txtTube = Trim(GetText(argSpread, lRow1, 7))
    txtEquip = Trim(GetText(argSpread, lRow1, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    lRow = 0
    For lCol = gResCol + 1 To gResCol + 35
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argRes = vasRes1
            Else
                Set argRes = vasRes2
            End If
            If lRow = 21 Then lRow = 1
            
            SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol

End Sub

Private Sub cmdPrev_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 - 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    txtRack = Trim(GetText(argSpread, lRow1, 6))
    txtTube = Trim(GetText(argSpread, lRow1, 7))
    txtEquip = Trim(GetText(argSpread, lRow1, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    lRow = 0
    For lCol = gResCol + 1 To gResCol + 35
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argRes = vasRes1
            Else
                Set argRes = vasRes2
            End If
            If lRow = 21 Then lRow = 1
            
            SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol

End Sub

Private Sub cmdPrint_Click()
    Dim i As Integer
    
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            vasList.Value = 0
            Call FN_ReportPtint(i)
            
        End If
    Next i
    
End Sub

Function FN_ReportPtint(argRow As Integer)
    'vasReport.PrintSmartPrint = True
    vasReport.PrintMarginLeft = 600
    vasReport.PrintMarginRight = 0
    vasReport.PrintMarginTop = 1000
    vasReport.PrintMarginBottom = 0
    'vasReport.BorderStyle = BorderStyleNone
    vasReport.PrintBorder = False
''''    vasReport.PrintGrid = False
    vasReport.PrintColor = True

    'vasReport.SetCellBorder 1, 36, 1, 36, 8, &HFFFFFFFF, 1
    Dim intRow As Integer
    Dim intResSplit As Integer
    Dim strRes1 As String
    Dim strRes2 As String
    ClearSpread vasSortInfo
    SetBackColor vasReport, 13, 13, 3, 10, 255, 255, 255
    SetBackColor vasReport, 15, 15, 3, 10, 255, 255, 255
    
'    vasReport.col = 3
'    vasReport.Row = 11
'    vasReport.BackColor = RGB(255, 255, 255)
'
'    vasReport.col = 3
'    vasReport.Row = 13
'    vasReport.BackColor = RGB(255, 255, 255)


    intRow = argRow
    intResSplit = InStr(1, Trim(GetText(vasList, intRow, colResult)), "/")
    If intResSplit > 0 Then
        strRes1 = Mid(Trim(GetText(vasList, intRow, colResult)), 1, intResSplit - 1)
        strRes2 = Mid(Trim(GetText(vasList, intRow, colResult)), intResSplit + 1)
    Else
        strRes1 = Trim(GetText(vasList, intRow, colResult))
        strRes2 = ""
    End If


    '이름 (성별 / 나이)
    SetText vasReport, Trim(GetText(vasList, intRow, colPName)) & "(" & Trim(GetText(vasList, intRow, colSexAge)) & ")", 4, 2
    '환자번호
    SetText vasReport, Trim(GetText(vasList, intRow, colPID)), 4, 5
    '진료과
    SetText vasReport, Trim(GetText(vasList, intRow, colIO)) & "-" & Trim(GetText(vasList, intRow, colMedDept)), 4, 8

    '검체명
    SetText vasReport, Trim(GetText(vasList, intRow, colPos)), 5, 2
    '바코드번호
    SetText vasReport, Trim(GetText(vasList, intRow, colBarcode)), 5, 5
    '접수일자
    SetText vasReport, Format(Trim(GetText(vasList, intRow, colReceDate)), "@@@@-@@-@@"), 5, 8


    '검사일자
    SetText vasReport, Format(Trim(GetText(vasList, intRow, colTestDate)), "@@@@-@@-@@"), 6, 2
    'W/L 번호
    SetText vasReport, Trim(GetText(vasList, intRow, colReceNo)), 6, 5




    'Asaay Info
    SetText vasReport, Trim(GetText(vasList, intRow, colAssay)), 10, 3

    'Test Result
    SetText vasReport, strRes1, 13, 3
    SetText vasReport, strRes2, 15, 3
    '색변환
    If InStr(1, strRes1, "MTB           : Not detected") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 0, 255, 50
    ElseIf InStr(1, strRes1, "MTB           : Detected") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 255, 0, 150
    ElseIf InStr(1, strRes1, "ERROR") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 255, 100, 50
    End If

    If InStr(1, strRes2, "RIF resistance: Not detected") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 0, 255, 50
    ElseIf InStr(1, strRes2, "RIF resistance: Detected") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 255, 0, 150
    ElseIf InStr(1, strRes2, "ERROR") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 255, 100, 50
    End If
    
    
    If InStr(1, strRes1, ": NEG") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 0, 255, 50
    ElseIf InStr(1, strRes1, ": POS") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 255, 0, 150
    ElseIf InStr(1, strRes1, "ERROR") > 0 Then
        SetBackColor vasReport, 13, 13, 3, 10, 255, 100, 50
    End If
    
    If InStr(1, strRes2, ": NEG") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 0, 255, 50
    ElseIf InStr(1, strRes2, ": POS") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 255, 0, 150
    ElseIf InStr(1, strRes2, "ERROR") > 0 Then
        SetBackColor vasReport, 15, 15, 3, 10, 255, 100, 50
    End If
    
    If Trim(GetText(vasList, intRow, colExamName)) = "" Then
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
    End If
    '검사명
'    Probe a
'    Probe B
'    Probe C
'    Probe D
'    Probe E
'    SPC
'    QC -1
'    QC -2
    If Trim(GetText(vasList, intRow, colExamCode)) = "L2742" Then
        SetText vasReport, Trim("Probe A"), 21, 1
        SetText vasReport, Trim("Probe B"), 22, 1
        SetText vasReport, Trim("Probe C"), 23, 1
        SetText vasReport, Trim("Probe D"), 24, 1
        SetText vasReport, Trim("Probe E"), 25, 1
        SetText vasReport, Trim("SPC"), 26, 1
        SetText vasReport, Trim("QC-1"), 27, 1
        SetText vasReport, Trim("QC-2"), 28, 1
        
        SetText vasReport, Trim("MTB    :"), 44, 2
        SetText vasReport, Trim("RIF    :"), 45, 2
        SetText vasReport, Trim(""), 46, 2
        SetText vasReport, Trim(""), 46, 3
        SetBackColor vasReport, 44, 44, 3, 10, 255, 255, 255
        SetBackColor vasReport, 45, 45, 3, 10, 255, 255, 255
        SetBackColor vasReport, 46, 46, 3, 10, 255, 255, 255
        
        SetText vasReport, Trim("※ MTB 양성 판정 기준 : Ct값 중 가장 최소값 2개의 차이 : < 2"), 47, 1
        SetText vasReport, Trim("※ RIF 내성 판정 기준 : Ct값 중 가장 큰수 - Ct값 중 가장 작은 수 : >4"), 48, 1
        
        'Ct, Pt, AnalRes
        'ProA
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProAPt)), 21, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProARes)), 21, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProACheck)), 21, 5
        'ProB
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPB)), 22, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProBPt)), 22, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProBRes)), 22, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProBCheck)), 22, 5
        'ProC
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPC)), 23, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProCPt)), 23, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProCRes)), 23, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProCCheck)), 23, 5
        'ProD
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPD)), 24, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProDPt)), 24, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProDRes)), 24, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProDCheck)), 24, 5
        'ProE
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPE)), 25, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProEPt)), 25, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProERes)), 25, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProECheck)), 25, 5
        'SPC
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBSPC)), 26, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCPt)), 26, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCRes)), 26, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCCheck)), 26, 5
        'QC1
        SetText vasReport, Trim(GetText(vasList, intRow, colQC1Ct)), 27, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colQC1Pt)), 27, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colQC1Res)), 27, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colQC1Check)), 27, 5
    
        'QC2
        SetText vasReport, Trim(GetText(vasList, intRow, colQC2Ct)), 28, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colQC2Pt)), 28, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colQC2Res)), 28, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colQC2Check)), 28, 5
        
    Else
        SetText vasReport, Trim("Toxine B"), 21, 1
        SetText vasReport, Trim("BinaryToxine"), 22, 1
        SetText vasReport, Trim("TcdC"), 23, 1
        SetText vasReport, Trim("SPC"), 24, 1
        SetText vasReport, Trim(""), 25, 1
        SetText vasReport, Trim(""), 26, 1
        SetText vasReport, Trim(""), 27, 1
        SetText vasReport, Trim(""), 28, 1
        
        SetText vasReport, Trim(""), 44, 2
        SetText vasReport, Trim(""), 45, 2
        
        SetText vasReport, Trim(""), 47, 1
        SetText vasReport, Trim(""), 48, 1
        
        
        'Ct, Pt, AnalRes
        'ProA
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 21, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProAPt)), 21, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProARes)), 21, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProACheck)), 21, 5
        'ProB
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPB)), 22, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProBPt)), 22, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProBRes)), 22, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProBCheck)), 22, 5
        'ProC
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBPC)), 23, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colProCPt)), 23, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colProCRes)), 23, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colProCCheck)), 23, 5
        'SPC
        SetText vasReport, Trim(GetText(vasList, intRow, colMTBSPC)), 24, 2
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCPt)), 24, 3
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCRes)), 24, 4
        SetText vasReport, Trim(GetText(vasList, intRow, colSPCCheck)), 24, 5
        
        
        'ProE
        SetText vasReport, "", 25, 2
        SetText vasReport, "", 25, 3
        SetText vasReport, "", 25, 4
        SetText vasReport, "", 25, 5
        'SPC
        
        SetText vasReport, "", 26, 2
        SetText vasReport, "", 26, 3
        SetText vasReport, "", 26, 4
        SetText vasReport, "", 26, 5
        'QC1
        
        SetText vasReport, "", 27, 2
        SetText vasReport, "", 27, 3
        SetText vasReport, "", 27, 4
        SetText vasReport, "", 27, 5
    
        'QC2
        
        SetText vasReport, "", 28, 2
        SetText vasReport, "", 28, 3
        SetText vasReport, "", 28, 4
        SetText vasReport, "", 28, 5
    
    
    
        
    End If
    
    
    
'    If Trim(GetText(vasList, intRow, colQC2Res)) = "NEG" Then
'        SetText vasReport, "PASS", 26, 5
'    Else
'        SetText vasReport, Trim(GetText(vasList, intRow, colQC2Res)), 26, 5
'    End If


    '이미지
    'SetText vasReport, Trim(GetText(vasList, intRow, colImage)), 16, 7
    'Select a single cell
    vasReport.Col = 6
    vasReport.Row = 21

    'Define cells as type PICTURE
    vasReport.CellType = 9
    vasReport.TypeHAlign = 2
    vasReport.TypeVAlign = 2
    'vasReport.TypePictMaintainScale = True
    vasReport.TypePictStretch = True
    Dim strImagePath As String
    Dim strImageName As String
    strImagePath = "C:\GeneXpert\Report\Result Graph\"
    
    '파일명 만들기
    strImageName = Trim(GetText(vasList, intRow, colBarcode)) & "_"
    strImageName = strImageName & Format(Trim(GetText(vasList, intRow, colStartDate)), "@@@@.@@.@@_@@.@@.@@")
                                                                    '       17051202733_2017.05.12_14.05.06_Primary Curve.jpg
    strImageName = strImageName & "_Primary Curve"
    
    'vasReport.TypePictPicture = LoadPicture(App.Path & "\" & Trim(GetText(vasList, intRow, colBarcode)) & ".bmp")
    If Dir(strImagePath & strImageName & ".jpg", vbDirectory) = strImageName & ".jpg" Then
        vasReport.TypePictPicture = LoadPicture("C:\GeneXpert\Report\Result Graph\" & strImageName & ".jpg")
        'vasReport.TypePictPicture = LoadPicture(App.Path & "\1.bmp")
    Else
        vasReport.TypePictPicture = LoadPicture("")
        'vasReport.TypePictPicture = LoadPicture(App.Path & "\2.bmp")
    End If


    'StarTime
    SetText vasReport, Format(Trim(GetText(vasList, intRow, colStartDate)), "@@@@-@@-@@ @@:@@:@@"), 33, 3
    'ENDTime
    SetText vasReport, Format(Trim(GetText(vasList, intRow, colEndDate)), "@@@@-@@-@@ @@:@@:@@"), 33, 8

    'ModuleNo
    SetText vasReport, Trim(GetText(vasList, intRow, colRack)), 34, 3

    'Catridge S/N
    SetText vasReport, Trim(GetText(vasList, intRow, colCartNo)), 34, 8

    'Reagent
    SetText vasReport, Trim(GetText(vasList, intRow, colReagentNo)), 35, 3

    'Expiration Date
    SetText vasReport, Format(Trim(GetText(vasList, intRow, colExpDate)), "@@@@-@@-@@"), 35, 8

    'Error1
    SetText vasReport, "     " & Mid(GetText(vasList, intRow, colError1) & GetText(vasList, intRow, colError2), 1, 100), 38, 1
    SetText vasReport, "     " & Mid(GetText(vasList, intRow, colError1) & GetText(vasList, intRow, colError2), 101), 39, 1

    'Error2
    SetText vasReport, "     " & Mid(GetText(vasList, intRow, colError1) & GetText(vasList, intRow, colError2), 1, 100), 40, 1
    SetText vasReport, "     " & Mid(GetText(vasList, intRow, colError1) & GetText(vasList, intRow, colError2), 101), 41, 1
    
    
    If Trim(GetText(vasList, intRow, colExamCode)) = "L2742" Then
        ResultRole intRow
    Else
        SetText vasReport, Trim(""), 44, 3
        SetText vasReport, Trim(""), 45, 3
        
        SetText vasReport, Trim("Toxin B    :"), 44, 2
        SetText vasReport, Trim("BinaryToxin:"), 45, 2
        SetText vasReport, Trim("TcdC       :"), 46, 2
        
        If InStr(1, GetText(vasReport, 21, 4), "NEG") > 0 Then
            SetText vasReport, "  NEGATIVE", 44, 3
            SetBackColor vasReport, 44, 44, 3, 10, 0, 255, 50
        ElseIf InStr(1, GetText(vasReport, 21, 4), "POS") > 0 Then
            SetText vasReport, "  POSITIVE", 44, 3
            SetBackColor vasReport, 44, 44, 3, 10, 255, 0, 150
            
        End If
        
        If InStr(1, GetText(vasReport, 22, 4), "NEG") > 0 Then
            SetText vasReport, "  NEGATIVE", 45, 3
            SetBackColor vasReport, 45, 45, 3, 10, 0, 255, 50
        ElseIf InStr(1, GetText(vasReport, 22, 4), "POS") > 0 Then
            SetText vasReport, "  POSITIVE", 45, 3
            SetBackColor vasReport, 45, 45, 3, 10, 255, 0, 150
        End If
        
        If InStr(1, GetText(vasReport, 23, 4), "NEG") > 0 Then
            SetText vasReport, "  NEGATIVE", 46, 3
            SetBackColor vasReport, 46, 46, 3, 10, 0, 255, 50
        ElseIf InStr(1, GetText(vasReport, 23, 4), "POS") > 0 Then
            SetText vasReport, "  POSITIVE", 46, 3
            SetBackColor vasReport, 46, 46, 3, 10, 255, 0, 150
        End If
        
    End If
''    'vasReport.PrintSmartPrint = True
''    vasReport.PrintFooter = "/fz" & "20"
''    vasReport.PrintFooter = "/fz" & 20 & "/r국립암센터 진단검사의학과 특수검사실/n/n/n"
    'vasReport.Action = ActionPrint
    'vasReport.PrintSmartPrint = True
    vasReport.Action = ActionPrint
End Function

Function ResultRole(argRow As Integer)
    Dim strMTB As String
    Dim strRif As String
    Dim strTemp As String
    
    vasSortInfo.SetText 1, 1, Trim(GetText(vasList, argRow, colMTBPA))
    vasSortInfo.SetText 1, 2, Trim(GetText(vasList, argRow, colMTBPB))
    vasSortInfo.SetText 1, 3, Trim(GetText(vasList, argRow, colMTBPC))
    vasSortInfo.SetText 1, 4, Trim(GetText(vasList, argRow, colMTBPD))
    vasSortInfo.SetText 1, 5, Trim(GetText(vasList, argRow, colMTBPE))
    vasSortInfo.Sort 1, 1, 1, 5, SortByRow, 1
    
    
    'MTB 결과
    strTemp = ""
    strTemp = CCur(Trim(GetText(vasSortInfo, 2, 1))) - CCur(Trim(GetText(vasSortInfo, 1, 1)))
    If strTemp = "0" Then strTemp = "0.0"
    
    strMTB = "  " & strTemp & " (" & Format(Trim(GetText(vasSortInfo, 2, 1)), "##0.0") & " - " & Format(Trim(GetText(vasSortInfo, 1, 1)), "##0.0") & " )"
    
    'RIF 결과
    strTemp = ""
    strTemp = CCur(Trim(GetText(vasSortInfo, 5, 1))) - CCur(Trim(GetText(vasSortInfo, 1, 1)))
    If strTemp = "0" Then strTemp = "0.0"
    
    strRif = "  " & strTemp & " (" & Format(Trim(GetText(vasSortInfo, 5, 1)), "##0.0") & " - " & Format(Trim(GetText(vasSortInfo, 1, 1)), "##0.0") & " )"
    
    
    'ResultInfo
    SetText vasReport, strMTB, 44, 3
    SetText vasReport, strRif, 45, 3
    
    
End Function



'''Private Sub cmdResCall_Click()
'''    Dim sFileName As String
'''
'''    cdFindFile.Filter = "All Files (*.*)|*.*|All Files (*.*)|*.*"
'''
'''
'''    cdFindFile.ShowOpen
'''
'''    sFileName = cdFindFile.Filename
'''
'''    If Trim(sFileName) = "" Then
'''        Exit Sub
'''    End If
'''
'''    ReadTxtFile vasLogRes, sFileName
'''End Sub

'''Private Sub ReadTxtFile(ByVal argSpread As vaSpread, argFilePath As String)
'''
'''    Dim FN
'''    Dim strLine As String
'''    Dim Buff As String
'''    Dim iRow As Long
'''    Dim i As Long
'''    Dim j As Long
'''    Dim x As Long
'''    Dim iCol As Long
'''
'''    Dim iRowCnt As Long
'''    Dim blResStart As Boolean
'''
'''    'ClearSpread ======================================
'''    argSpread.Row = 1
'''    argSpread.Col = 0
'''    argSpread.Row2 = 1
'''    argSpread.Col2 = 0
'''    argSpread.BlockMode = True
'''    argSpread.Action = 3
'''    argSpread.BlockMode = False
'''
'''
'''    iRow = 0
'''    FN = FreeFile
'''
'''    blResStart = False
'''
'''    Open argFilePath For Input As #FN
'''        Do While Not EOF(FN)
'''            iRow = argSpread.DataRowCnt + 1
'''            If iRow > argSpread.MaxRows Then
'''                argSpread.MaxRows = iRow
'''            End If
'''
'''            Line Input #1, strLine
'''            If InStr(1, strLine, "Result Information") > 0 Then
'''                blResStart = True
'''            End If
'''
'''            If blResStart = True Then
'''                If InStr(1, strLine, "=================================================================") > 0 Then
'''                Else
'''                    iCol = 0
'''                    For j = 1 To 40
'''                        x = InStr(1, strLine, vbTab)
'''                        If x > 0 Then
'''                            iCol = iCol + 1
'''
'''                            argSpread.SetText iCol, iRow, Mid(strLine, 1, x - 1)
'''                            strLine = Mid(strLine, x + 1)
'''
'''                        End If
'''                    Next
'''                End If
'''
'''            End If
'''
'''        Loop
'''    Close #FN
'''
'''    For i = 2 To argSpread.DataRowCnt
'''        m2000 i
'''    Next
'''End Sub

'''Sub m2000(argRow As Long)
'''    Dim sAllStr As String
'''    Dim sRowStr() As String
'''    Dim sRowCnt As Integer
'''    Dim sResCV As String
'''    Dim sResCopy As String
'''    Dim sResIU As String
'''    Dim sExamCode As String
'''    Dim sEquipCode As String
'''    Dim sExamName As String
'''    Dim sPos As String
'''    Dim sBarcode As String
'''    Dim i, j As Integer
'''    Dim sRowPart() As String
'''    Dim si As Integer
'''    Dim sResStart As Boolean
'''    Dim sExamSeq As String
'''    Dim sRow As Integer
'''    Dim x, y As Integer
'''    Dim sResIC As String
'''    Dim sResTG As String
'''    Dim sEquipResult As String
'''
'''
'''    sEquipCode = "HBV"
'''
''''''    If sResStart = True Then
''''''    sExamSeq = Trim(sRowPart(1))
'''    sPos = Trim(GetText(vasLogRes, argRow, 1))
'''
'''    sBarcode = Trim(GetText(vasLogRes, argRow, 4))
'''    sResIC = Trim(GetText(vasLogRes, argRow, 23))
'''    sResTG = Trim(GetText(vasLogRes, argRow, 24))
'''    sEquipResult = Trim(GetText(vasLogRes, argRow, 25))
'''    sResIU = Trim(GetText(vasLogRes, argRow, 26))
'''
'''
'''    If sResIU <> "" And sBarcode <> "" Then
'''        sRow = -1
'''
'''        If sRow = -1 Then
'''            sRow = vasList.DataRowCnt + 1
'''        End If
'''        If sRow > vasList.MaxRows Then
'''            vasList.MaxRows = sRow
'''        End If
'''
'''        SetText vasList, sBarcode, sRow, 2
'''        SetText vasList, sPos, sRow, 7
'''        SetText vasList, sResIC, sRow, 11
'''        SetText vasList, sResTG, sRow, 12
'''        SetText vasList, sEquipResult, sRow, 13
'''        SetText vasList, sResIU, sRow, 14
'''
'''        If IsNumeric(sBarcode) = True And Len(sBarcode) > 10 Then
'''            Get_Sample_Info sRow
'''        End If
'''
'''        SQL = "select examcode, examname from equipexam where equipcode = '" & sEquipCode & "'"
'''        res = db_select_Col(gLocal, SQL)
'''        sExamCode = Trim(gReadBuf(0))
'''        sExamName = Trim(gReadBuf(1))
'''
'''        SetText vasList, "결과", sRow, 8
'''        SetText vasList, sExamCode, sRow, 9
'''        SetText vasList, sExamName, sRow, 10
'''
'''        Save_Local_One sRow, "A"
'''
'''    End If
'''
'''End Sub


Private Sub cmdResDel_Click()
    Dim i, j As Integer
    Dim sDelS, sDelE As String
    
    If IsNumeric(txtSSeq.Text) = False Or IsNumeric(txtESeq.Text) = False Then
        Exit Sub
    End If
    
    sDelS = CInt(txtSSeq.Text)
    sDelE = CInt(txtESeq.Text)
    
    For i = sDelS To sDelE
    
        SQL = "delete from pat_res " & vbCrLf & _
              "where barcode = '" & Trim(GetText(vasList, i, 2)) & "' " & vbCrLf & _
              "and examcode = '" & Trim(GetText(vasList, i, 9)) & "' "
        res = SendQuery(gLocal, SQL)
        
    Next
    cmdSch_Click
    
End Sub

Private Sub cmdResPrint_Click()
    Dim iRow As Integer
    Dim jRow As Integer
    Dim kRow As Integer
    
    '환자정보관련
    Dim sRack As String
    Dim sPos As String
    Dim sSampleNo As String
    Dim sPID As String
    Dim sPName As String
    Dim sPSex As String
    Dim sPage As String
    Dim i As Integer
    
    Dim sExamName As String
    
    Dim sHead As String
    Dim sHead1 As String    '의뢰시간
    Dim sFoot As String
    Dim sSlip As String
    Dim sCurDate As String
    Dim sExamDate As String
    Dim sTitle As String
    Dim PageCnt As Integer
    Dim sDate As String
'    Dim sExamName As String
    Dim lRow As Integer
    Dim lsResult As String
    Dim sResult1 As String
    Dim sResFlag As String
    Dim sFName As String
    
    
    
On Error GoTo ErrGoto
    
    
    ClearSpread vasPrint
    
    SQL = "Select a.barcode, a.pname, a.receno,'', a.result, a.result_copy, a.result_iu, a.unit, a.recedate, a.posno " & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' and a.sendflag <> 'O' and a.pid <> '' " & vbCrLf & _
          "order by a.examcode, a.recedate, a.receno, a.posno "
    
    res = db_select_Vas(gLocal, SQL, vasPrint)
    
    SQL = "Select a.barcode, a.pname, a.receno,'', a.result, a.result_copy, a.result_iu, a.unit, a.recedate, a.posno " & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' and a.sendflag <> 'O' and a.pid = '' " & vbCrLf & _
          "order by a.examcode, a.recedate, a.receno, a.posno "
    
    res = db_select_Vas(gLocal, SQL, vasPrint)
    
    
    vasPrint.MaxRows = vasPrint.DataRowCnt
    sFName = Trim(GetText(vasPrint, 1, 8))
    For lRow = 1 To vasPrint.DataRowCnt
        
            lsResult = Trim(GetText(vasPrint, lRow, 7))
            If IsNumeric(lsResult) = True Then
                lsResult = Format(lsResult, "###,###,###,###,###")
            End If
            
            lsResult = lsResult & " IU/mL"
            sResult1 = Trim(GetText(vasPrint, lRow, 6))
            
            If IsNumeric(sResult1) = True Then
                sResult1 = CStr(CCur(sResult1) / 1000)
                sResult1 = Format(CCur(sResult1), "###,###,###,###,###.0")
                sResult1 = sResult1 & " X 10³copies/mL"
                
            Else
                sResFlag = Trim(Mid(sResult1, 1, 1))
                sResult1 = Trim(Mid(sResult1, 2))
                
                sResult1 = CStr(CCur(sResult1) / 1000)
                sResult1 = Format(CCur(sResult1), "###,###,###,###,###.0")
                sResult1 = sResFlag & " " & sResult1 & " X 10³copies/mL"
            End If
            
            lsResult = lsResult & " (" & sResult1 & ")"
            SetText vasPrint, lsResult, lRow, 4
            
    Next lRow
    
    
    
    If vasPrint.DataRowCnt = 0 Then
        MsgBox "출력할 결과가 없습니다."
        Exit Sub
    End If
    

'    i = InStr(1, cboExam.Text, " ")
'    If i > 0 Then
'        sExamName = Trim(Mid(cboExam.Text, i + 1))
'    End If

    
    
    sDate = Format(dtpExamDate.Value, "yyyy-mm-dd")
    
'    CommonDialog1.ShowPrinter
    
    PageCnt = vasPrint.PrintPageCount
    
    sCurDate = Format(CDate(Date), "yyyy/mm/dd") & "   " & Format(CDate(Time), "hh:mm:dd")
    
    
    sTitle = "Versant440 RESULT"
    
    sHead = "/fn""굴림체"" /fz""15"" /fb1 /fi0 /fu0 " & "/l" & "                          " & "" & sTitle & "" & "/n/n/n " & _
            "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "  FileName : " & sFName & "/fn""굴림체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "검사일자 : " & sDate & "       " & "/n" '& "/n/n" & _
            "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "  검사일자 : " & sDate & "/n"
    
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "  " & sCurDate & "/fn""굴림체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "        국립암센터   "
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot
    
    
    vasPrint.PrintOrientation = PrintOrientationPortrait
    
    vasPrint.PrintMarginTop = 800
    vasPrint.PrintMarginBottom = 600
    
    '현재 SS가 비대칭으로 출력함
    vasPrint.PrintMarginLeft = 660
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
    'vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
    
    '원하는 셀까지만 출력함
    vasPrint.Row = 1
    vasPrint.Row2 = vasPrint.DataRowCnt
    vasPrint.Col = 1
    vasPrint.Col2 = 11
    vasPrint.PrintType = PrintTypeCellRange
    
    vasPrint.PrintShadows = True
    
    vasPrint.Action = 13 'SS_ACTION_PRINT

ErrGoto:
    '사용자가 취소버튼을 눌렀습니다.
    Exit Sub
End Sub

Public Sub cmdSch_Click()
    Dim lRow, lCol As Long
    Dim lsID As String
    Dim liEquipCode As Integer
    Dim lsExamCode As String
    Dim lsAllExam  As String
    Dim i As Integer
    
    Dim rs_Res As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    ClearSpread vasList ', 1, 2
    
    'frameSch.Visible = True
    
    Me.MousePointer = 11
    
    If cboExam.Text = "전체선택" Then
        lsAllExam = "ALL"
    Else
        i = InStr(1, cboExam.Text, " ")
        If i > 0 Then
            lsAllExam = Trim(Mid(cboExam.Text, 1, i - 1))
        End If
    End If

    
    SQL = "Select '', a.barcode, a.pid, a.pname, a.receno, a.diskno, " & _
            "a.posno, a.sendflag, a.examcode, a.examname , a.result, a.result_copy, a.result_iu, a.recedate,''  " & _
            ", a.MTBA, a.MTBB, a.MTBC, a.MTBD, a.MTBE , a.RifA, a.RifB, a.RifC, a.RifD, a.RifE, a.MTBRemark " & _
            ",SPCMTB, SPCRif" & _
            ",StartDate, EndDate, CartNo, ReagentNo, ExpDate, Errorstring1, Errorstring2, "
    SQL = SQL & vbCrLf & "QC1, QC2,"
    
    SQL = SQL & vbCrLf & "ProAPt, ProBPt, ProCPt, ProDPt, ProEPt,"
    SQL = SQL & vbCrLf & "SPCPt, QC1Pt, QC2Pt, "
    
    SQL = SQL & vbCrLf & "ProARes, ProBRes, ProCRes, ProDRes, ProERes,"
    SQL = SQL & vbCrLf & "SPCRes, QC1Res, QC2Res, "
    
    SQL = SQL & vbCrLf & "Assay, "
    
    SQL = SQL & vbCrLf & "SEXAGE, MEDDEPT, PATIO, TESTDATE,"
    
    SQL = SQL & vbCrLf & "ProACK, ProBCK, ProCCK, ProDCK, ProECK,"
    SQL = SQL & vbCrLf & "SPCCK, QC1CK, QC2CK "
    
    SQL = SQL & vbCrLf & "from pat_res a"
    SQL = SQL & vbCrLf & "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' and sendflag <> 'O' and a.pid <> '' "
          If lsAllExam = "ALL" Then
        
          Else
            SQL = SQL & "and equipcode = '" & lsAllExam & "' "
          End If
    SQL = SQL & "order by a.examcode, a.recedate, a.receno, a.posno "
    
    res = db_select_Vas(gLocal, SQL, vasList)
'
    SQL = "Select '', a.barcode, a.pid, a.pname, a.receno, a.diskno, " & _
            "a.posno, a.sendflag, a.examcode, a.examname , a.result, a.result_copy, a.result_iu, a.recedate " & _
            ",'', a.MTBA, a.MTBB, a.MTBC, a.MTBD , a.MTBE , a.RifA, a.RifB, a.RifC, a.RifD, a.RifE, a.MTBRemark " & _
            ",SPCMTB, SPCRif" & _
            ",StartDate, EndDate, CartNo, ReagentNo, ExpDate, Errorstring1, Errorstring2, "
    SQL = SQL & vbCrLf & "QC1, QC2,"
    
    SQL = SQL & vbCrLf & "ProAPt, ProBPt, ProCPt, ProDPt, ProEPt,"
    SQL = SQL & vbCrLf & "SPCPt, QC1Pt, QC2Pt, "
    
    SQL = SQL & vbCrLf & "ProARes, ProBRes, ProCRes, ProDRes, ProERes,"
    SQL = SQL & vbCrLf & "SPCRes, QC1Res, QC2Res, "
    
    SQL = SQL & vbCrLf & "Assay, "
    
    SQL = SQL & vbCrLf & "SEXAGE, MEDDEPT, PATIO, TESTDATE,"
    
    SQL = SQL & vbCrLf & "ProACK, ProBCK, ProCCK, ProDCK, ProECK,"
    SQL = SQL & vbCrLf & "SPCCK, QC1CK, QC2CK "
    SQL = SQL & vbCrLf & "from pat_res a"
    SQL = SQL & vbCrLf & "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                         "  and a.equipno = '" & gEquip & "' and sendflag <> 'O' and a.pid = '' "
          If lsAllExam = "ALL" Then

          Else
            SQL = SQL & "and equipcode = '" & lsAllExam & "' "
          End If
    SQL = SQL & "order by a.examcode, a.recedate, a.receno, a.posno "

    res = db_select_Vas(gLocal, SQL, vasList, vasList.DataRowCnt + 1)
    
    
    
    For i = 1 To vasList.DataRowCnt
        Select Case Trim(GetText(vasList, i, 8))
        Case "A", "0"
            vasList.SetText 8, i, "결과"
'        Case "B"
'            vasList.SetText 8, i, "수신"
        Case "C", "1"
            vasList.SetText 8, i, "완료"
        Case Else
            vasList.SetText 8, i, ""
        End Select
    Next
    
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then GoTo ErrHandle
    
    lsID = "interface"
    lRow = 0
    
    Me.MousePointer = 0
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub

End Sub

Private Sub cmdSchClose_Click()
    frameSch.Visible = False
End Sub
'
'Private Sub cmdWorkList_Click()
'    frmWorkList.Show
'End Sub

Private Sub Command1_Click()
    Dim S As String
    Dim sPID As String
    Dim sSendData As String
    Dim sSndMessage As String
    Dim i As Integer
    Dim iRow As Integer
    Dim lResRow As Long
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResult As String
    
    
    For i = 1 To Len(txtData.Text)
    
    
        S = Mid(txtData, i, 1)
        
        Select Case S
          
        Case chrENQ
            Save_Raw_Data "[Rx" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                    
            gSndState = ""
            gENQFlag = 9
            
            gRecodeType = ""
'''            txtToday = Format(Date, "yyyy/mm/dd")
            
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
            gPreSpecID = ""
            gPreRow = 0
            
        Case chrACK
    
        Case chrSTX     '자료수신 시작
            txtBuff.Text = S
            
        Case chrETX
            txtBuff.Text = txtBuff.Text & S
        
        Case chrLF
            txtBuff.Text = txtBuff.Text & S
            Save_Raw_Data "[Rx" & Format(Time, "hh:mm:ss") & "]" & txtBuff.Text
            m2000 txtBuff.Text
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
        Case chrEOT     '자료수신 완료
            If gRecodeType = "R" Then
                gSndState = "R"
                
            ElseIf gRecodeType = "Q" Then
                gOrdRow = 0
                gPreMsg = chrENQ
                
                frmInterface.MSComm1.Output = chrENQ
                Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                        
                gSndState = "Q"
                gPreMsg = chrENQ
            End If
    
            gMsgFlag = ""
            gHeadRecode = ""
            txtBuff.Text = ""
                
        Case Else
            txtBuff.Text = txtBuff.Text & S
        End Select
    Next
    txtData = ""
    
End Sub

Sub Var_Clear()
    gOrderMessage = ""
    
    gBarCode = ""

    
    glRow = -1
End Sub


Sub m2000(argData As String)
    
'==========================================================
'2005/10/06 이상은
'결과가 3.0이면 3으로 저장됨 -> ccur에서 cstr로 바꿈
'RA, CRP 검사항목 QC인 경우, 결과 Positive인 경우 수치결과도 같이 할 것
'2009/03/17 이상은
'Anti-HBs 검사코드 l503h 이면 수치결과 그대로 전송되도록 할 것
'==========================================================

    Dim i As Integer
    Dim j As Integer
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim aCnt As Integer
    Dim bCnt As Integer
    
    Dim lsTemp As String

    Dim sDate As String
    Dim sGubun As String
    Dim sPID As String
    Dim sReceNo As String
'    Dim sSampleType As String
    Dim sSpecID As String
    Dim sTestID As String
    Dim sExamCode As String
    Dim sExamCode1 As String
    Dim sExamName As String
    Dim sSeq As String
    Dim sResClassCode As String
    Dim sFlag As String
    Dim sResult As String
    Dim sResult2 As String
    Dim sGiho As String
    Dim sExamDate As String
    Dim sResDateTime As String
    Dim sResComment As String
    
    
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sCnt As String
    Dim sPage As String
    
    Dim lsRefLow As String
    Dim lsRefHigh As String
    Dim lsRefRev As String
    Dim lsResDate As String
       
    
    Dim sExamCode_All As String
    Dim sPart_All As String
    Dim sBarCode As String
    Dim sBarCode1 As String
    Dim sBarcode2 As String
    Dim sOrDate As String
    
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lResRow As Long
    
    Dim jRow As Integer
    
    Dim sLen, sLen2 As String
    Dim aCount As Integer
    Dim iRCnt As Integer
    Dim lsLotNo As String
    
    Dim strSeq As String
    Dim strSeq1 As String
    Dim strSeq2 As String
'''    Dim sGiho As String
    
    Select Case Mid(argData, 3, 1)
    Case "H"    'Header
        gPreRow = -1
                
        Var_Clear
    Case "P"    'Patient
        gPatFlag = -1
        
    Case "O"    'Test Order
        aCount = aCount + 1
        
        iCnt = 0
        
        jCnt = 0
        
        For i = 1 To Len(argData)
            If Mid(argData, i, 1) = "|" Then
                iCnt = iCnt + 1
                Select Case iCnt
                Case 3  'PID
                    sLen = InStr(i + 1, argData, "|")
                    sPID = Mid(argData, i + 1, sLen - i - 1)
                    sSpecID = sPID
'                    strSeq = sPID
                    j = InStr(1, sSpecID, "^")
                    strSeq = Mid(sSpecID, j + 1)
                    
                    sSpecID = Left(sSpecID, j - 1)
                    
                    If sSpecID <> gSpecID Then
                        gSpecID = sSpecID
                    End If
                    
                    j = InStr(1, strSeq, "^")
                    If j > 0 Then
                        strSeq = Mid(strSeq, j + 1)
                        strSeq1 = Mid(strSeq, 1, 1)
                        strSeq2 = Mid(strSeq, 2)
                    End If
                    
                    
                    
                Case 11
                    sLen = InStr(i + 1, argData, "|")
                    If sLen > 0 Then
                        If Mid(argData, i + 1, sLen - i - 1) = "Q" Then
                            sSampleType = "Q"
                        Else
                            sSampleType = "P"
                        End If
                    Else
                        sSampleType = "P"
                    End If
                
                End Select
            End If
            
            If Mid(argData, i, 1) = "^" Then
                jCnt = jCnt + 1
                Select Case jCnt
                Case 5  'TestID
                    sLen = InStr(i + 1, argData, "^")
                    sTestID = Mid(argData, i + 1, sLen - i - 1)
                    gtestid = sTestID
                End Select
            End If
        Next i
        If sSampleType = "P" Then
        
            glRow = -1
            For lRow = 1 To vasList.DataRowCnt
                If Trim(GetText(vasList, lRow, colBarcode)) = gSpecID Then
                    glRow = lRow
                    
                    If gPatFlag = -1 Then
                        vasList_Click 2, glRow
                        
                        gPatFlag = 1
                        vasActiveCell vasList, glRow, 2
                    End If
    
                    Exit For
                End If
            Next lRow
            
            '2004/06/16 이상은========================================================
            'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
            If glRow = -1 Then  ' vaslist에 없는 검체의 결과가 나올 때 데이터 추가
                glRow = vasList.DataRowCnt + 1
                If glRow > vasList.MaxRows Then
                    vasList.MaxRows = glRow + 1
                End If
                vasActiveCell vasList, glRow, colBarcode
                SetText vasList, sSpecID, glRow, colBarcode
            End If
            '==========================================================================
            
            SetText vasList, strSeq, glRow, colPos
'''            SetText vasList, strSeq2, glRow, colSeq2
            
            If Trim(GetText(vasList, glRow, colPID)) = "" Then
                If Len(gSpecID) > 10 Then
                    Get_Sample_Info glRow
                End If
            End If
        
        End If
        
        gPreSpecID = sSpecID
        
        gPreRow = glRow
        
    Case "R"    'Result
        gRecodeType = "R"
        
'        SetText vaslist, "Result", glRow, colState
            
        iCnt = 0
    
        For i = 1 To Len(argData)
            If Mid(argData, i, 1) = "^" Then
                iCnt = iCnt + 1
                If iCnt = 8 Then
                    sFlag = Mid(argData, i + 1, 1)
                    Exit For
                ElseIf iCnt = 9 Then
'                    lsLotNo = Mid(argData, i + 1)
'                    lsLotNo = Mid(lsLotNo, 1, InStr(1, lsLotNo, "^") - 1)
'                    lsLotNo = Mid(argData, i + 1, 1)
                End If
            End If
        Next i
            
        sExamCode = ""
        sResClassCode = ""
        sExamName = ""
        sResult = ""
        
        If sFlag = "F" Then
            aCnt = 0
            
            For i = 1 To Len(argData)
                If Mid(argData, i, 1) = "|" Then
                   aCnt = aCnt + 1
                   Select Case aCnt
                   Case 3
                        sLen = InStr(i + 1, argData, "|")
                        sResult = Trim(Mid(argData, i + 1, sLen - i - 1))
                        sResult2 = Trim(sResult)
                        Dim sReceExamCode   As String
                        Dim sRv             As String
                        Dim i2              As Integer
                        
                        Clear_XML_Exam
                        sRv = Online_XML(gXml_S07, Trim(GetText(vasList, glRow, colBarcode)))
                        sReceExamCode = ""
                        
                        For i2 = 0 To UBound(gExam_Select)
                    
                            If sReceExamCode = "" Then
                                sReceExamCode = "'" & Trim(gExam_Select(i2).TST_CD) & "'"
                            Else
                                sReceExamCode = sReceExamCode & ",'" & Trim(gExam_Select(i2).TST_CD) & "'"
                            End If
                            
                        Next i2
                        
                        SQL = "select examcode, examname from equipexam "
                        SQL = SQL & "where equip = '" & gEquip & "' and equipcode = '" & gtestid & "'"
                        If sReceExamCode <> "" Then
                            SQL = SQL & " and examcode in ( " & sReceExamCode & ") "
                        End If
                        
                        res = db_select_Col(gLocal, SQL)
                        
                        If res > 0 Then
                            sExamCode = gReadBuf(0)
                            sExamName = gReadBuf(1)
                        End If
                        
'                        sGiho = ""
'                        If Not IsNumeric(Left(sResult, 1)) Then
'                            sGiho = Left(sResult, 1)
'                            sResult = Trim(Mid(sResult, 2))
'                            sResult2 = sResult
'                        End If
'
                    
                    Case 12      '결과
                        sLen = InStr(i + 1, argData, "|")
                        sResDateTime = Mid(argData, i + 1, sLen - i - 1)
                        lsResDate = Mid(sResDateTime, 9, 6)
                        
                    End Select
                End If
            Next i
        End If
        
        
        If gtestid <> "" And sResult <> "" Then
            If sSampleType = "P" Then
                sExamCode_All = ""
                sPart_All = ""
                sGiho = ""
                If Mid(sResult, 1, 1) = ">" Or Mid(sResult, 1, 1) = "<" Then
                    sGiho = Mid(sResult, 1, 1)
                    sResult = Trim(Mid(sResult, 2))
                End If
                
                    
                If IsNumeric(sResult) = True Then
                    sResult = Format(sResult, "###,###,###,###,###")
                End If
                
                If sGiho <> "" Then
                    sResult = sGiho & " " & sResult
                End If
                
                
                SetText vasList, sResult, glRow, colResult
'''                SetText vasList, sResult, glRow, colEquipResult
                                 
                SetText vasList, sExamCode, glRow, colExamCode
'                SetText vaslist, gtestid, glRow, colequipcode
                SetText vasList, sExamName, glRow, colExamName
                SetText vasList, gtestid, glRow, colEquipCode
                SetText vasList, "Result", glRow, colState
                
'''                SetText vasList, Proc_ResCode(sResult, sExamCode), glRow, colResCode
                
'''                txtTextResult = Proc_ResComment(Trim(GetText(vasList, glRow, colResCode)), sExamCode, sResult)
                
                Save_Local_One glRow, "0"

            End If
        End If
        
        gMsgFlag = ""
        gHeadRecode = ""
        txtBuff.Text = ""
        
    Case "Q"    'Request
        
    Case "L"    '자료수신 완료
'''        Patient_init
    End Select
End Sub

Private Sub Command2_Click()
    XPert_All Text1
    Text1 = ""
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    
    dtpExamDate.Value = Format(Date, "yyyy-mm-dd")
    
    gResCol = 8

    GetComSetup
    
    cn_Local_Flag = False
    cn_Server_Flag = False
    
    GetSetup
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
''    MSComm1.CommPort = CA_COM.ComPort
''    MSComm1.Settings = CA_COM.Speed & "," & CA_COM.Parity & "," & CA_COM.DataBit & "," & CA_COM.StopBit
''    If CA_COM.RTSEnable = "1" Then
''        MSComm1.RTSEnable = True
''    Else
''        MSComm1.RTSEnable = False
''    End If
''    If CA_COM.DTREnable = "1" Then
''        MSComm1.DTREnable = True
''    Else
''        MSComm1.DTREnable = False
''    End If
''    MSComm1.PortOpen = True
    
    If Trim(GetSetting("MEDIMATE", "COBRA", "SendMode", "0")) = "1" Then
        chkMode.Value = 1
        subSend1.Checked = True
        subSend2.Checked = False
    Else
        chkMode.Value = 0
        subSend1.Checked = False
        subSend2.Checked = True
    End If
    
    GetExamCode
    WinSock_Listen Winsock1
    If Not IsNumeric(gDays) Then
        gDays = 30
        
        WritePrivateProfileString "Data", "Days", gDays, App.Path & "\interface.ini"
        
    End If
    
'    SQL = "Delete from pat_res where examdate < '" & DateAdd("d", 0 - CInt(gDays), dtpExamDate.Value) & "' "
'    SendQuery gLocal, SQL
    
    txtBuff = ""
    txtData = ""
    
    gRCnt = 0
    
    lblUser.Caption = gIFUser
    
    vasList.MaxRows = 1
    
    cboExam.Clear
    
    SQL = "Select EquipCode, ExamName from EquipExam where Equip = '" & gEquip & "' and UseFlag = 1 "
    res = db_select_Combo_2(gLocal, SQL, cboExam)
    cboExam.AddItem "전체선택", 0
    cboExam.ListIndex = 0
    
    
    gRow = 0
    ExamCount
End Sub

Sub ExamCount()
    SQL = "SELECT COUNT(DISKNO)"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
    SQL = SQL & vbCrLf & "  AND DISKNO = 'B1'"
    res = db_select_Col(gLocal, SQL)
    
    If gReadBuf(0) = "" Then
        SetText vasModuleCnt, "0", 1, 1
    Else
        SetText vasModuleCnt, gReadBuf(0), 1, 1
    End If
    
    SQL = "SELECT COUNT(DISKNO)"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
    SQL = SQL & vbCrLf & "  AND DISKNO = 'B2'"
    res = db_select_Col(gLocal, SQL)
    
    If gReadBuf(0) = "" Then
        SetText vasModuleCnt, "0", 1, 2
    Else
        SetText vasModuleCnt, gReadBuf(0), 1, 2
    End If
    
    SQL = "SELECT COUNT(DISKNO)"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
    SQL = SQL & vbCrLf & "  AND DISKNO = 'B3'"
    res = db_select_Col(gLocal, SQL)
    
    If gReadBuf(0) = "" Then
        SetText vasModuleCnt, "0", 1, 3
    Else
        SetText vasModuleCnt, gReadBuf(0), 1, 3
    End If
    
    SQL = "SELECT COUNT(DISKNO)"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
    SQL = SQL & vbCrLf & "  AND DISKNO = 'B4'"
    res = db_select_Col(gLocal, SQL)
    
    If gReadBuf(0) = "" Then
        SetText vasModuleCnt, "0", 1, 4
    Else
        SetText vasModuleCnt, gReadBuf(0), 1, 4
    End If
    
End Sub


Sub GetExamCode()
    Dim AdoRs_Exam As ADODB.Recordset
    Dim lCol As Long
    Dim i As Integer
    Dim vWidth
    
'    ReDim gArrExam(0)
'    gArrExam(0) = ""
    
    'SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh, RSGubun " & CR & _
          "  From EquipExam " & CR & _
          " WHERE Equip = '" & gEquip & "' " & CR & _
          "   and UseFlag = 1 " & vbCrLf & _
          " Order by seqno "
    SQL = "SELECT EquipCode, ExamName, SeqNo, count(ExamName) " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "   and UseFlag = 1 " & vbCrLf & _
          " Group by EquipCode, ExamName, SeqNo " & vbCrLf & _
          " Order by SeqNo"
    
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
    If AdoRs_Exam Is Nothing Then
        ClearSpread vasList, 1, 1
    Else
        ClearSpread vasList, 1, 1
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        
        
        Do Until AdoRs_Exam.EOF
            lCol = lCol + 1
             
            AdoRs_Exam.MoveNext
        Loop
        
        ReDim gArrExam(lCol - gResCol, 2)
        
        vWidth = 58.75 / (lCol - gResCol)
        If CCur(vWidth) < 7.25 Then
            vWidth = 7.25
        End If
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        
        
        
        Do Until AdoRs_Exam.EOF
            lCol = lCol + 1
             
            'ReDim Preserve gArrExam(lCol - gResCol, 5)
            For i = 0 To 1
                If IsNull(AdoRs_Exam.Fields(i).Value) Then
                    gArrExam(lCol - gResCol, i + 1) = ""
                Else
                    gArrExam(lCol - gResCol, i + 1) = AdoRs_Exam.Fields(i).Value
                End If
            Next i
            
'            SetText vasList, AdoRs_Exam.Fields(1).Value, 0, lCol
'            vasList.ColWidth(lCol) = vWidth
            
            AdoRs_Exam.MoveNext
        Loop
    End If
    
    If lCol = 0 Then
        gMaxCol = gResCol + 1
    Else
        gMaxCol = lCol + 1
    End If
        
'    vasList.MaxCols = gMaxCol
'
'    vasList.ColWidth(gMaxCol) = 0
    
End Sub

Private Sub Form_Resize()
    
On Error GoTo errorCheck



    Exit Sub
errorCheck:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisConnect_Local
    DisConnect_Server
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
End Sub

Private Sub Label6_Click()
    If Text1.Visible = True Then
        Text1.Visible = False
        Command2.Visible = False
    Else
        Command2.Visible = True
        Text1.Visible = True
    End If
    
End Sub

Private Sub MSComm1_OnComm()
    Dim S As String
    Dim sPID As String
    Dim sSendData As String
    Dim sSndMessage As String
    Dim i As Integer
    Dim iRow As Integer
    Dim lResRow As Long
    Dim sExamCode As String
    Dim sExamName As String
    Dim sResult As String
    
    S = MSComm1.Input
    
    Select Case S
      
    Case chrENQ
        Save_Raw_Data "[Rx" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                
        gSndState = ""
        gENQFlag = 9
        
        gRecodeType = ""
'''        txtToday = Format(Date, "yyyy/mm/dd")
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrACK
        
        gPreSpecID = ""
        gPreRow = 0
        
    Case chrACK

    Case chrSTX     '자료수신 시작
        txtBuff.Text = S
        
    Case chrETX
        txtBuff.Text = txtBuff.Text & S
    
    Case chrLF
        txtBuff.Text = txtBuff.Text & S
        Save_Raw_Data "[Rx" & Format(Time, "hh:mm:ss") & "]" & txtBuff.Text
        m2000 txtBuff.Text
        MSComm1.Output = chrACK
        Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrACK
        
    Case chrEOT     '자료수신 완료
        If gRecodeType = "R" Then
            gSndState = "R"
            
        ElseIf gRecodeType = "Q" Then
            gOrdRow = 0
            gPreMsg = chrENQ
            
            frmInterface.MSComm1.Output = chrENQ
            Save_Raw_Data "[Tx" & Format(Time, "hh:mm:ss") & "]" & chrENQ
                    
            gSndState = "Q"
            gPreMsg = chrENQ
        End If

        gMsgFlag = ""
        gHeadRecode = ""
        txtBuff.Text = ""
            
    Case Else
        txtBuff.Text = txtBuff.Text & S
    End Select
End Sub

Private Sub Command3_Click()
'''
'''    'vasReport.PrintSmartPrint = True
'''    vasReport.PrintMarginLeft = 1200
'''    vasReport.PrintMarginTop = 1200
'''    'vasReport.BorderStyle = BorderStyleNone
'''    vasReport.PrintBorder = False
'''    vasReport.PrintGrid = False
'''    vasReport.PrintColor = True
'''
'''    'vasReport.SetCellBorder 1, 36, 1, 36, 8, &HFFFFFFFF, 1
'''    Dim intRow As Integer
'''    Dim intResSplit As Integer
'''    Dim strRes1 As String
'''    Dim strRes2 As String
'''
'''    intRow = 1
'''    intResSplit = InStr(1, Trim(GetText(vasList, intRow, colResult)), "/")
'''    If intResSplit > 0 Then
'''        strRes1 = Mid(Trim(GetText(vasList, intRow, colResult)), 1, intResSplit - 1)
'''        strRes2 = Mid(Trim(GetText(vasList, intRow, colResult)), intResSplit + 1)
'''    Else
'''        strRes1 = Trim(GetText(vasList, intRow, colResult))
'''        strRes2 = ""
'''    End If
'''
'''    '등록번호/이름
'''    SetText vasReport, Trim(GetText(vasList, intRow, colPID)), 4, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colPName)), 4, 8
'''    '바코드번호
'''    SetText vasReport, Trim(GetText(vasList, intRow, colBarcode)), 5, 2
'''    'Asaay Info
'''    SetText vasReport, Trim(GetText(vasList, intRow, colAssay)), 7, 3
'''
'''
'''    'Test Result
'''    SetText vasReport, strRes1, 9, 2
'''    SetText vasReport, strRes2, 10, 2
'''
'''    'Ct, Pt, AnalRes
'''    'ProA
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBPA)), 16, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProAPt)), 16, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProARes)), 16, 5
'''    'ProB
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBPB)), 17, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProBPt)), 17, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProBRes)), 17, 5
'''    'ProC
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBPC)), 18, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProCPt)), 18, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProCRes)), 18, 5
'''    'ProD
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBPD)), 19, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProDPt)), 19, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProDRes)), 19, 5
'''    'ProE
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBPE)), 20, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProEPt)), 20, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colProERes)), 20, 5
'''    'SPC
'''    SetText vasReport, Trim(GetText(vasList, intRow, colMTBSPC)), 21, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colSPCPt)), 21, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colSPCRes)), 21, 5
'''    'QC1
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC1Ct)), 22, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC1Pt)), 22, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC1Res)), 22, 5
'''    'QC2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC2Ct)), 23, 2
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC2Pt)), 23, 4
'''    SetText vasReport, Trim(GetText(vasList, intRow, colQC2Res)), 23, 5
'''
'''    '이미지
'''    'SetText vasReport, Trim(GetText(vasList, intRow, colImage)), 16, 7
'''    'Select a single cell
'''    vasReport.col = 7
'''    vasReport.Row = 16
'''
'''    'Define cells as type PICTURE
'''    vasReport.CellType = 9
'''    vasReport.TypeHAlign = 2
'''    vasReport.TypeVAlign = 2
'''    'vasReport.TypePictMaintainScale = True
'''    vasReport.TypePictStretch = True
'''
'''    'vasReport.TypePictPicture = LoadPicture(App.Path & "\" & Trim(GetText(vasList, intRow, colBarcode)) & ".bmp")
'''    If Dir(App.Path & "\" & Trim(GetText(vasList, intRow, colBarcode)) & ".jpg", vbDirectory) = Trim(GetText(vasList, intRow, colBarcode)) & ".jpg" Then
'''        vasReport.TypePictPicture = LoadPicture(App.Path & "\" & Trim(GetText(vasList, intRow, colBarcode)) & ".jpg")
'''        'vasReport.TypePictPicture = LoadPicture(App.Path & "\1.bmp")
'''    Else
'''        'vasReport.TypePictPicture = LoadPicture("")
'''        vasReport.TypePictPicture = LoadPicture(App.Path & "\2.bmp")
'''    End If
'''    'StarTime
'''    SetText vasReport, Trim(GetText(vasList, intRow, colStartDate)), 27, 3
'''    'ENDTime
'''    SetText vasReport, Trim(GetText(vasList, intRow, colEndDate)), 27, 8
'''
'''    'ModuleNo
'''    SetText vasReport, Trim(GetText(vasList, intRow, colRack)), 28, 3
'''
'''    'Catridge S/N
'''    SetText vasReport, Trim(GetText(vasList, intRow, colCartNo)), 29, 3
'''
'''    'Reagent
'''    SetText vasReport, Trim(GetText(vasList, intRow, colReagentNo)), 30, 3
'''
'''    'Expiration Date
'''    SetText vasReport, Trim(GetText(vasList, intRow, colExpDate)), 31, 3
'''
'''    'Error1
'''    SetText vasReport, "     " & GetText(vasList, intRow, colError1) & GetText(vasList, intRow, colError2), 33, 1
'''
'''
'''    'ResultInfo
'''    SetText vasReport, "36", 36, 4
'''    SetText vasReport, "37", 37, 4
'''
'''
'''
'''    vasReport.Action = ActionPrint
    
End Sub

Public Sub Res_Proc(asPath As String)
    Dim sFileName As String
    Dim sAllFile As String
    Dim FilNum
    Dim sTxtStr As String
    
    sFileName = asPath
    sAllFile = ""

    FilNum = FreeFile
    
    Open sFileName For Input As #FilNum   ' 입력을 위해 파일을 엽니다.

    Do While Not EOF(FilNum)
        Input #FilNum, sTxtStr
        sAllFile = sAllFile & sTxtStr & chrLF
    Loop

    Close #FilNum

    VerSant440 sAllFile

End Sub

Sub VerSant440(asVar As String)
    Dim sAllStr As String
    Dim sRowStr() As String
    Dim sRowCnt As Integer
    Dim sResCV As String
    Dim sResCopy As String
    Dim sResIU As String
    Dim sExamCode As String
    Dim sEquipCode As String
    Dim sExamName As String
    Dim sPos As String
    Dim sBarCode As String
    Dim i, j As Integer
    Dim sRowPart() As String
    Dim si As Integer
    Dim sResStart As Boolean
    Dim sExamSeq As String
    Dim sRow As Integer
    Dim X, y As Integer
    
    
    sAllStr = asVar
    
    sRowCnt = 1
    sEquipCode = "HBV"
    ReDim sRowStr(1)
    
    For i = 1 To Len(sAllStr)
        
        If Mid(sAllStr, i, 1) = chrLF Then
            sRowCnt = sRowCnt + 1
            ReDim Preserve sRowStr(sRowCnt)
            sRowStr(sRowCnt) = ""
        ElseIf Mid(sAllStr, i, 1) = chrLF Then
        Else
            sRowStr(sRowCnt) = sRowStr(sRowCnt) & Mid(sAllStr, i, 1)
        End If
        
    Next
    
    sResStart = False
    
    For i = 1 To sRowCnt
        si = 1
        ReDim sRowPart(15)
        For j = 1 To Len(sRowStr(i))
            If Mid(sRowStr(i), j, 1) = vbTab Then
                si = si + 1
                sRowPart(si) = ""
            Else
                sRowPart(si) = sRowPart(si) & Mid(sRowStr(i), j, 1)
            End If
        Next
        
        If sResStart = True Then
            sExamSeq = Trim(sRowPart(1))
            sPos = Trim(sRowPart(2))
            sBarCode = Trim(sRowPart(3))
            sResCV = Trim(sRowPart(7))
            sResCopy = Trim(sRowPart(9))
            sResIU = Trim(sRowPart(13))
            
            If sResCV <> "" And sBarCode <> "" Then
                sRow = -1
'                For x = 1 To vasList.DataRowCnt
'                    If Trim(GetText(vasList, x, 2)) = sBarcode Then
'                        sRow = x
'                        Exit For
'                    End If
'                Next
                If sRow = -1 Then
                    sRow = vasList.DataRowCnt + 1
                End If
                If sRow > vasList.MaxRows Then
                    vasList.MaxRows = sRow
                End If
                
                SetText vasList, sBarCode, sRow, 2
                SetText vasList, sPos, sRow, 7
'                SetText vasList, sExamSeq, sRow, 4
                SetText vasList, sResCV, sRow, 11
                SetText vasList, sResCopy, sRow, 12
                SetText vasList, sResIU, sRow, 13
                If sBarCode <> "" Then
                    Get_Sample_Info sRow
                End If
                
                SQL = "select examcode, examname from equipexam where equipcode = '" & sEquipCode & "'"
                res = db_select_Col(gLocal, SQL)
                sExamCode = Trim(gReadBuf(0))
                sExamName = Trim(gReadBuf(1))
                
                SetText vasList, "결과", sRow, 8
                SetText vasList, sExamCode, sRow, 9
                SetText vasList, sExamName, sRow, 10
                
                Save_Local_One sRow, "A"
                
            End If
            
        End If
        
        If Trim(sRowPart(1)) = "#" Then
            sResStart = True
            
        End If
        
        
    Next
    
'    MsgBox "test"
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsbarcode As String
Dim lsPID As String
Dim lsReceNo As String
Dim sRes As String


    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsbarcode = Trim(GetText(vasList, asRow, 2))   '샘플 바코드 번호
    
    sRes = Online_XML(gXml_S03, lsbarcode)
    SetText vasList, gPat_Info_Select.PT_NO, asRow, colPID
    SetText vasList, gPat_Info_Select.PT_NM, asRow, colPName
    
'    SetText vasList, gPat_Info_Select.SEX, asRow, colPSex
'    SetText vasList, gPat_Info_Select.AGE, asRow, colPAge
    SetText vasList, gPat_Info_Select.ACPTNO_1, asRow, 5
    SetText vasList, gPat_Info_Select.ACPT_DTETM, asRow, colReceDate
'    SetText vasList, gPat_Info_Select.SPC_CD_1, asRow, colReceno
    
    '검체코드 로 검체명 불러오기
    SQL = "SELECT SPCNAME"
    SQL = SQL & vbCrLf & "  FROM SPCCONFIG"
    SQL = SQL & vbCrLf & " WHERE SPCCODE = '" & gPat_Info_Select.SPC_CD_1 & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res < 1 Then
        SetText vasList, gPat_Info_Select.SPC_CD_1, asRow, colPos
    Else
        SetText vasList, gReadBuf(0), asRow, colPos
    End If
    
    '성별/나이 붙여서 저장, 진료과 표시, 입외 구분 추가
    SetText vasList, gPat_Info_Select.SEX & "/" & gPat_Info_Select.AGE, asRow, colSexAge
    SetText vasList, gPat_Info_Select.MEDDEPT, asRow, colMedDept
    SetText vasList, gPat_Info_Select.ORD_SITE, asRow, colIO
    
    Get_Sample_Info = 1
End Function

Sub PACKARD(asVar As String)
    Dim iStr As Integer
    Dim i As Integer
    Dim iCnt As Integer
    
    Dim j As Integer
    Dim K As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    
    Dim sSeqNo As String
    Dim sProtocol As String
    Dim sProtocol1 As String
    
    Dim sExamCode As String
    Dim sResult As String
    Dim sPoint As String
    Dim sTmpStr As String
    Dim sCPMRes As String
    Dim sExamType As String
    
    Dim sErr As String
    
    Dim iRow As Integer
    Dim lCol As Long
    
    Dim ii As Integer
    Dim jj As Integer
    
    Dim iPos As Integer

    If asVar = "" Then
        Exit Sub
    End If
    
    If Len(asVar) < 10 Then
        Exit Sub
    End If
    
    lsData = asVar
    
    gRow = -1
        
    iStr = 1
    i = 0
    iCnt = 0

    gRCnt = 0
    
    '콤마(,)로 순번, 계측치, 에러, 프로토콜, '', '', 결과, PAT/ID, CR+LF 등 구분함
    '1(S#),6(A:CPM),16(A:%ERR),프로토콜번호,A:%B(F),
    '콤마갯수가 8개이면 6번째 자리가 결과, 콤마갯수가 9개이면 7번째 자리가 결과임
    ii = 0
    For jj = 1 To Len(lsData)
        If Mid(lsData, jj, 1) = "," Then
            ii = ii + 1
        End If
    Next jj

    i = InStr(iStr, lsData, ",")
    
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        Select Case iCnt
'        Case 1      '순번
'            sSeqNo = lsTemp
            
        Case 3      'Error
            sErr = ""
            
            sErr = Trim(lsTemp)
        Case 4      '프로토콜
            sProtocol = Trim(lsTemp)
            
'            gRow = 0
'            For iRow = 1 To vasID.DataRowCnt
'                If sSeqNo = Trim(GetText(vasID, iRow, 0)) And sProtocol = Trim(GetText(vasID, iRow, colProtocol)) Then
'                    gRow = iRow
'
'                    Exit For
'                End If
'            Next iRow
            
        Case 6      '결과
            If ii = 6 Then
                sResult = ""
                
                sResult = Trim(lsTemp)
                sCPMRes = sResult
                
                
                If IsNumeric(sResult) Then
                    'sResult = Format(CCur(sResult), "#0.00")
                    
                    '소수점 처리
                    sPoint = "0"
                    SQL = " Select PointSize From EquipExam Where equip = '" & gEquip & "' " & CR & _
                          " And EquipCode = '" & Trim(sProtocol) & "' "
                    res = db_select_Col(gLocal, SQL)
                    sPoint = Trim(gReadBuf(0))
                    
                    If IsNumeric(sPoint) = True And IsNumeric(sResult) = True Then
                        If CInt(sPoint) > 0 Then
                            sTmpStr = "#0."
                            For i = 1 To CInt(sPoint)
                                sTmpStr = sTmpStr & "0"
                            Next i
                        ElseIf CInt(sPoint) = 0 Then
                            sTmpStr = "#0"
                        Else
                            sTmpStr = ""
                        End If
                        If Trim(sTmpStr) <> "" Then
                            sResult = Format(sResult, sTmpStr)
                        End If
                    End If
                End If
                
                j = InStr(1, Trim(lsData), "SAMPLE")
    
                If j > 0 Then
                    sSeqNo = Trim(Mid(lsData, j + 7))
                    
                    iPos = InStr(1, sSeqNo, Chr(13))
                    If iPos > 0 Then
                        sSeqNo = Mid(sSeqNo, 1, iPos - 1)
                    End If
                    sExamType = "Sample"
                Else
                    sExamType = "STD"
                End If
            End If
            
        Case 7      '결과
            If ii = 7 Then
                sResult = ""
                sResult = Trim(lsTemp)
                
                If IsNumeric(sResult) Then
                    'sResult = Format(CCur(sResult), "#0.00")
                    
                    '소수점 처리
                    sPoint = "0"
                    SQL = " Select PointSize From EquipExam Where equip = '" & gEquip & "' " & CR & _
                          " And EquipCode = '" & Trim(sProtocol) & "' "
                    res = db_select_Col(gLocal, SQL)
                    sPoint = Trim(gReadBuf(0))
                    
                    If IsNumeric(sPoint) = True And IsNumeric(sResult) = True Then
                        If CInt(sPoint) > 0 Then
                            sTmpStr = "#0."
                            For i = 1 To CInt(sPoint)
                                sTmpStr = sTmpStr & "0"
                            Next i
                        ElseIf CInt(sPoint) = 0 Then
                            sTmpStr = "#0"
                        Else
                            sTmpStr = ""
                        End If
                        If Trim(sTmpStr) <> "" Then
                            sResult = Format(sResult, sTmpStr)
                        End If
                    End If
                End If
             
                j = InStr(1, Trim(lsData), "SAMPLE")
    
                If j > 0 Then
                    sSeqNo = Trim(Mid(lsData, j + 7))
                    sExamType = "Sample"
                Else
                    sExamType = "STD"
                    
                End If
            End If
        End Select
    
        lsTemp = ""
        i = InStr(iStr, lsData, ",")
    Loop


'    If sSeqNo > vaslist.DataRowCnt Then
'        Exit Sub
'    End If
            
    '해당 프로토콜 위치찾기
    For i = 1 To UBound(gArrExam)
        If Trim(sProtocol) = gArrExam(i, 1) Then
            'k = gArrExam(i, 1)
            'lCol = (gArrExam(i, 1))
            
            K = gResCol + i
            lCol = K
            Exit For
        End If
    Next i
    
    sExamCode = ""
'    SQL = " Select examcode From equipexam where equip = '" & Trim(gEquip) & "' " & CR & _
'          " And equipcode = '" & Trim(sProtocol) & "'"
'    res = db_select_Col(gLocal, SQL)
'
'    sExamCode = Trim(gReadBuf(0))
    
'    SQL = "select barcode from pat_res where equip = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "and examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' and equipcode = '" & sProtocol & "'"
          
    If sExamType = "Sample" Then
        SQL = "update pat_res set result = '" & sResult & "', refvalue = '" & sCPMRes & "', sendflag = 'B' " & vbCrLf & _
          "where equipno = '" & Trim(gEquip) & "' and posno = " & sSeqNo & " " & vbCrLf & _
          "and examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' and equipcode = '" & sProtocol & "'"
    Else
    End If
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
    End If
    
    cmdSch_Click
    
    
'    lCol = -1
'    For i = gResCol + 1 To vasList.MaxCols
'        If Trim(GetText(vasList, 1, i)) = Trim(sExamCode) Then
'            lCol = i
'            Exit For
'        End If
'    Next i

    '결과 디스플레이
'    For iRow = 1 To vasList.DataRowCnt
'        If Trim(GetText(vasList, iRow, lCol)) = "*" And sSeqNo = Trim(GetText(vasList, iRow, 7)) Then
'            gRCnt = sSeqNo
'
'            SetText vasList, sResult, iRow, lCol
'            SetText vasList, "수신완료", iRow, 8
'
'            '로컬에 저장하기
'            sExamCode = ""
'            SQL = "Select ExamCode from EquipExam where equip = '" & gEquip & "' and Equipcode = '" & sProtocol & "' "
'            res = db_select_Col(gLocal, SQL)
'            sExamCode = Trim(gReadBuf(0))
'
'            If (sExamCode <> "") And sResult <> "" Then
'                SQL = "Delete FROM pat_res " & vbCrLf & _
'                      "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'                      "  AND equipcode = '" & sProtocol & "'" & vbCrLf & _
'                      "  AND barcode = '" & Trim(GetText(vasList, iRow, 2)) & "' "
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    SaveQuery SQL
'                    'Exit Function
'                End If
'
'                SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'                        "barcode, examtype, receno, " & _
'                        "pid, pname, pjumin, page, psex, " & _
'                        "resdate, seqno, diskno, posno, " & _
'                        "equipcode, examcode, " & _
'                        "result, sendflag, examname, " & _
'                        "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'                      "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'                      "'" & Trim(GetText(vasList, iRow, 2)) & "','" & Trim(GetText(vasList, iRow, 5)) & "', '" & Trim(GetText(vasList, iRow, 5)) & "', " & _
'                      "'" & Trim(GetText(vasList, iRow, 3)) & "', '" & Trim(GetText(vasList, iRow, 4)) & "', '', 0, '', " & _
'                      "'', '', '" & Trim(GetText(vasList, iRow, 6)) & "', '" & Trim(GetText(vasList, iRow, 7)) & "', " & vbCrLf & _
'                      "'" & sProtocol & "', '" & sExamCode & "', " & _
'                      "'" & sResult & "', 'B', '', " & vbCrLf & _
'                      "'', '', '', '', " & _
'                      "'', '' ) "
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    SaveQuery SQL
'                    'Exit Function
'                End If
'
'                Exit For
'            End If
'
'            Exit For
'        End If
'    Next iRow
'
'    If chkMode.Value = 1 And gRCnt = vasList.DataRowCnt Then
'        For iRow = 1 To vasList.DataRowCnt
'            sResult = ""
'            sResult = Trim(GetText(vasList, iRow, lCol))
'
'            If Set_EqpResultsql(sExamCode, sResult, "", Trim(GetText(vasList, iRow, 2)), gInsCode) = True Then
'                SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
'                SetText vasList, "완료", iRow, gResCol
'
'                vasList.Row = iRow
'                vasList.Col = 1
'                vasList.Value = 1
'
'                Update_Sample Trim(GetText(vasList, iRow, 2))
'                'DeleteWorkList lsID
'
'                SQL = "delete from pat_res " & vbCrLf & _
'                      "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'                      "  AND barcode = '" & Trim(GetText(vasList, iRow, 2)) & "' " & vbCrLf & _
'                      "  And equipcode = '" & sProtocol & "' " & vbCrLf & _
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Else
'                SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
'                SetText vasList, "실패", iRow, gResCol
'            End If
'        Next iRow
'
'        gRCnt = 0
'    End If
End Sub


Sub COBAS_Amplicor()
    Dim myVar As String
    Dim i, j, K, a As Long
    Dim lsData As String
    Dim lsTmp As String
    
    Dim lsRing As String
    Dim lsOrdDate As String
    Dim lsTube As String
    Dim lsOrdType As String
    Dim lsSpcID As String
    Dim lsEquipCode As String
    Dim lsTestType As String
    Dim lsResDate As String
    Dim lsQualRes As String
    Dim lsQualFlag As String
    Dim lsQualResRaw As String
    Dim lsQualQS As String
    Dim lsQuanRes As String
    Dim lsQuanFlag As String
    Dim lsQuanResRaw As String
    Dim lsQuanQS As String
    
    Dim lsID As String
    Dim lRow As Long
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResType As String
    Dim lsResPoint As String
    Dim lsRefLow As String
    Dim lsRefHigh As String
    Dim lsRes As String
    
    myVar = Trim(txtBuff)
    i = InStr(1, myVar, chrLF)
    Do While i > 0
        lsData = Left(myVar, i - 1)
        myVar = Mid(myVar, i + 1)
        
        Select Case Left(lsData, 2)
        'Result.Order processing
        Case "00"   'Result Selection
        Case "01"   'A-ring ID
            lsRing = Trim(Mid(lsData, 4, 6))
        Case "02"   'Order Date / Time
            lsOrdDate = Trim(Mid(lsData, 4))
        Case "03"   'Order Run Mode
        Case "04"   'A-tube Position
            lsTube = Trim(Mid(lsData, 4, 2))
        Case "05"   'Order Type
            lsOrdType = Trim(Mid(lsData, 4, 1))
        Case "06"   'Specimen Information
            lsSpcID = Trim(Mid(lsData, 4, 2))
        Case "07"   'Test ID
            lsEquipCode = Trim(Mid(lsData, 4, 3))
            
            lsExamCode = ""
            lsExamName = ""
            lsSeqNo = ""
            lsResType = ""
            lsResPoint = ""
            lsRefLow = ""
            lsRefHigh = ""
            
            lsRes = ""
            
            SQL = "Select EquipCode, ExamCode, ExamName, SeqNo, RSGubun, PointSize, RefLow, RefHigh " & vbCrLf & _
                  "from equipexam where equip = '" & gEquip & "' and EquipCode = '" & lsEquipCode & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = lsEquipCode Then
                lsExamCode = Trim(gReadBuf(1))
                lsExamName = Trim(gReadBuf(2))
                lsSeqNo = Trim(gReadBuf(3))
                lsResType = Trim(gReadBuf(4))
                lsResPoint = Trim(gReadBuf(5))
                lsRefLow = Trim(gReadBuf(6))
                lsRefHigh = Trim(gReadBuf(7))
            Else
                lsEquipCode = ""
            End If
        Case "08"   'Test Type
            lsTestType = Trim(Mid(lsData, 4, 1))
        Case "10"   'Result Date / Time
            lsResDate = Trim(Mid(lsData, 4))
'            If IsDate(lsResDate) Then
'                lsResDate = Format(lsResDate, "yyyy-mm-dd hh:nn:ss")
'            End If
            lsResDate = Mid(lsResDate, 7, 4) & "-" & Mid(lsResDate, 4, 2) & "-" & Mid(lsResDate, 1, 2) & Mid(lsResDate, 11)
        Case "11"   'Qualitative Result
            lsQualRes = Trim(Mid(lsData, 4, 1))
            Select Case lsQualRes
            
            Case "1"
                lsQualRes = "Positive"
            Case "2"
                lsQualRes = "Negative"
            Case "3"
                lsQualRes = "Trace"
            Case Else
                lsQualRes = ""
            End Select
            
            lsQualFlag = Trim(Mid(lsData, 6, 8))
            
            If lsResType = "T" Then
                lsRes = lsQualRes
                
                lRow = -1
                For lRow = 1 To vasList.DataRowCnt
                    If Trim(GetText(vasList, lRow, 6)) = lsRing And _
                       Trim(GetText(vasList, lRow, 7)) = lsTube Then
                        gRow = lRow
                        Exit For
                    End If
                Next lRow
                
                If lRow = -1 Then
                    SQL = "Select barcode, pid, pname, examtype, diskno, posno " & vbCrLf & _
                          "from pat_res " & vbCrLf & _
                          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
                          "  and diskno = '" & lsRing & "'   " & vbCrLf & _
                          "  and posno = '" & lsTube & "' " & vbCrLf & _
                          "  and equipcode = '" & lsEquipCode & "' "
                    res = db_select_Vas(gLocal, SQL, vasList, vasList.DataRowCnt + 1, 2)
                    If res = 0 Then
                        lRow = vasList.DataRowCnt + 1
                        gRow = lRow
                        
                        vasList.SetText 2, lRow, lsRing & "-" & lsTube
                    Else
                        lRow = vasList.DataRowCnt
                        gRow = lRow
                    End If
                Else
                    gRow = lRow
                End If
                'gRow = gRow + 1
                lRow = gRow
                
                vasList.SetText 6, lRow, lsRing
                vasList.SetText 7, lRow, lsTube
                vasList.SetText 8, lRow, "수신"
                
                For j = 1 To UBound(gArrExam)
                    If gArrExam(j, 1) = lsEquipCode Then
                    
                        SetText vasList, lsRes, gRow, gResCol + j

'                        If gArrExamRes(liEquipCode).RefFlag = "H" Then
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 255, 127, 0
'                        ElseIf gArrExamRes(liEquipCode).RefFlag = "L" Then
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 127, 255
'                        Else
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 0, 0
'                        End If

'                        Save_Local_One lRow, i, "A"
                        
                        SQL = "Delete FROM pat_res " & vbCrLf & _
                              "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                              "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                              "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
                        res = SendQuery(gLocal, SQL)
                        If res = -1 Then
                            SaveQuery SQL
                            'Exit Function
                        End If
                        
                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                              "barcode, examtype, receno, " & _
                              "pid, pname, pjumin, page, psex, " & _
                              "resdate, seqno, diskno, posno, " & _
                              "equipcode, examcode, " & _
                              "result, sendflag, examname, " & _
                              "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
                              "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                              "'" & Trim(GetText(vasList, lRow, 2)) & "','" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 5)) & "', " & _
                              "'" & Trim(GetText(vasList, lRow, 3)) & "', '" & Trim(GetText(vasList, lRow, 4)) & "', '', 0, '', " & _
                              "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & Trim(GetText(vasList, lRow, 7)) & "', " & vbCrLf & _
                              "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
                              "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
                              "'', '', '', '', " & _
                              "'', '' ) "
                        res = SendQuery(gLocal, SQL)
                        If res = -1 Then
                            SaveQuery SQL
                            'Exit Function
                        End If
                        
                        If chkMode.Value = 1 And Trim(GetText(vasList, lRow, 2)) <> "" And IsNumeric(Mid(Trim(GetText(vasList, lRow, 2)), 2)) = True Then
                            'res = Set_EqpResultsql(lsExamCode, lsRes, "", Trim(GetText(vasList, lRow, 2)), gInsCode)
                            If Set_EqpResultsql(lsExamCode, lsRes, "", Trim(GetText(vasList, lRow, 2)), gInsCode) Then
                                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                                SetText vasList, "완료", lRow, gResCol
                                
                                vasList.Row = lRow
                                vasList.Col = 1
                                vasList.Value = 1
                                
                                Update_Sample Trim(GetText(vasList, lRow, 2))
                                'DeleteWorkList lsID
                                SQL = "delete from pat_res " & vbCrLf & _
                                      "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                                      "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' " & vbCrLf & _
                                      "  And equipcode = '" & lsEquipCode & "' " & vbCrLf & _
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                    Exit Sub
                                End If
                                
                            Else
                                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                                SetText vasList, "실패", lRow, gResCol
                            End If
                        End If
                        
                        Exit For
                    End If
                Next j
                
            Else
                lsRes = ""
            End If
            
            PreRow = lRow
            PreRack = lsRing
            PrePos = lsTube
            
        Case "12"   'Qualitative Raw Data
        If lsResType = "T" And lsOrdType = "1" Then
            lsQualResRaw = Mid(lsData, 4, 5)
            Debug.Print lsQualResRaw
            lsTmp = ""
            
            For j = 1 To Len(lsQualResRaw)
                If IsNumeric(Mid(lsQualResRaw, j, 1)) Then
                    lsTmp = lsTmp & Mid(lsQualResRaw, j, 1)
                Else
                    lsTmp = lsTmp & "0"
                End If
            Next j
            
            lsQualResRaw = lsTmp
            lsQualResRaw = Left(lsQualResRaw, 2) & "." & Mid(lsQualResRaw, 3)
            If IsNumeric(lsQualResRaw) Then
                lsQualResRaw = Format(CCur(lsQualResRaw), "#0.000")
            End If
            If IsNumeric(lsRefHigh) Then
                If CCur(lsRefHigh) <= CCur(lsQualResRaw) Then
                    lsRes = "Positive"
                End If
            End If
            If IsNumeric(lsRefLow) Then
                If CCur(lsRefLow) > CCur(lsQualResRaw) Then
                    lsRes = "Negative"
                End If
            End If
            If IsNumeric(lsRefHigh) And IsNumeric(lsRefLow) Then
                If CCur(lsRefLow) <= CCur(lsQualResRaw) And CCur(lsRefHigh) > CCur(lsQualResRaw) Then
                    lsRes = "Trace"
                End If
            End If
            
            'Debug.Print lsQualResRaw
            'Debug.Print lsRes
            
            For j = 1 To UBound(gArrExam)
                If gArrExam(j, 1) = lsEquipCode Then
                    
                    SetText vasList, lsRes, lRow, gResCol + j

'                        If gArrExamRes(liEquipCode).RefFlag = "H" Then
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 255, 127, 0
'                        ElseIf gArrExamRes(liEquipCode).RefFlag = "L" Then
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 127, 255
'                        Else
'                            SetForeColor vasList, lRow, lRow, gResCol + j, gResCol + j, 0, 0, 0
'                        End If

'                        Save_Local_One lRow, i, "A"
                    SQL = "Delete FROM pat_res " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                          "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
                    res = SendQuery(gLocal, SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        'Exit Function
                    End If
                    
                    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                          "barcode, examtype, receno, " & _
                          "pid, pname, pjumin, page, psex, " & _
                          "resdate, seqno, diskno, posno, " & _
                          "equipcode, examcode, " & _
                          "result, sendflag, examname, " & _
                          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
                          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                          "'" & Trim(GetText(vasList, lRow, 2)) & "','" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 5)) & "', " & _
                          "'" & Trim(GetText(vasList, lRow, 3)) & "', '" & Trim(GetText(vasList, lRow, 4)) & "', '', 0, '', " & _
                          "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & Trim(GetText(vasList, lRow, 7)) & "', " & vbCrLf & _
                          "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
                          "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
                          "'', '', '', '', " & _
                          "'', '' ) "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        'Exit Function
                    End If
                    If chkMode.Value = 1 And Trim(GetText(vasList, lRow, 2)) <> "" And IsNumeric(Mid(Trim(GetText(vasList, lRow, 2)), 2)) = True Then
                        If Set_EqpResultsql(lsExamCode, lsRes, "", Trim(GetText(vasList, lRow, 2)), gInsCode) = True Then
                            SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                            SetText vasList, "완료", lRow, gResCol
                            
                            vasList.Row = lRow
                            vasList.Col = 1
                            vasList.Value = 1
                            
                            Update_Sample Trim(GetText(vasList, lRow, 2))
                            
                            SQL = "delete from pat_res " & vbCrLf & _
                                  "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                                  "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                                  "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' " & vbCrLf & _
                                  "  And equipcode = '" & lsEquipCode & "' " & vbCrLf & _
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                            
                        Else
                            SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                            SetText vasList, "실패", lRow, gResCol
                        End If
                    End If
                    
                    Exit For
                End If
            Next j
        End If
        Case "13"   'Quantitative Result
            lsQuanRes = Trim(Mid(lsData, 4, 10))
            
            lsQuanFlag = Trim(Mid(lsData, 15, 8))
        
            If lsResType <> "T" Then
                lsRes = lsQuanRes
                                            
                If IsNumeric(lsRes) Then
                    If Mid(lsQuanFlag, 5, 1) = "8" Then
                        lsRes = lsExamName & " : not detected"
                    Else
                        If IsNumeric(lsResType) Then
                            lsRes = Format(lsRes * CCur(lsResType), "#0.0000000")
                            lsTmp = ""
                            K = 0
                            For j = 1 To Len(lsRes)
                                If IsNumeric(Mid(lsRes, j, 1)) Then
                                    K = K + 1
                                    If K > 3 Then
                                        lsTmp = lsTmp & "0"
                                    Else
                                        lsTmp = lsTmp & Mid(lsRes, j, 1)
                                    End If
                                Else
                                    lsTmp = lsTmp & Mid(lsRes, j, 1)
                                End If
                            Next j
                            lsRes = lsTmp
                            lsRes = Format(lsRes, "0.00E+00")
                            
                            j = InStr(1, lsRes, "E")
                            If j > 0 Then
                                If Mid(lsRes, j + 1, 1) = "-" Then
                                    lsRes = Left(lsRes, j - 1) & "x10^" & Mid(lsRes, j + 1, 1) & CInt(Mid(lsRes, j + 2))
                                Else
                                    lsRes = Left(lsRes, j - 1) & "x10^" & CInt(Mid(lsRes, j + 2))
                                End If
                            End If
                        End If
                    End If
                Else
                    lsRes = lsExamName & " : not detected"
                End If
                    
                    
                lRow = -1
                For lRow = 1 To vasList.DataRowCnt
                    If Trim(GetText(vasList, lRow, 6)) = lsRing And _
                       Trim(GetText(vasList, lRow, 7)) = lsTube Then
                        gRow = lRow
                        Exit For
                    End If
                Next lRow
                
                If lRow = -1 Then
                    SQL = "Select barcode, pid, pname, examtype, diskno, posno " & vbCrLf & _
                          "from pat_res " & vbCrLf & _
                          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                          "  and equipno = '" & gEquip & "' " & vbCrLf & _
                          "  and diskno = '" & lsRing & "'   " & vbCrLf & _
                          "  and posno = '" & lsTube & "' " & vbCrLf & _
                          "  and equipcode = '" & lsEquipCode & "' "
                    res = db_select_Vas(gLocal, SQL, vasList, vasList.DataRowCnt + 1, 2)
                    If res = 0 Then
                        lRow = vasList.DataRowCnt + 1
                        gRow = lRow
                        
                        vasList.SetText 2, lRow, lsRing & "-" & lsTube
                    Else
                        lRow = vasList.DataRowCnt
                        gRow = lRow
                    End If
                Else
                    gRow = lRow
                End If

'                gRow = gRow + 1
                lRow = gRow

                vasList.SetText 6, lRow, lsRing
                vasList.SetText 7, lRow, lsTube
                vasList.SetText 8, lRow, "수신"

                For j = 1 To UBound(gArrExam)
                    If gArrExam(j, 1) = lsEquipCode Then
                    
                        SetText vasList, lsRes, gRow, gResCol + j
                        
                        SQL = "Delete FROM pat_res " & vbCrLf & _
                              "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                              "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                              "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
                        res = SendQuery(gLocal, SQL)
                        If res = -1 Then
                            SaveQuery SQL
                            'Exit Function
                        End If
                        
                        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                                "barcode, examtype, receno, " & _
                                "pid, pname, pjumin, page, psex, " & _
                                "resdate, seqno, diskno, posno, " & _
                                "equipcode, examcode, " & _
                                "result, sendflag, examname, " & _
                                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
                              "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                              "'" & Trim(GetText(vasList, lRow, 2)) & "','" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 5)) & "', " & _
                              "'" & Trim(GetText(vasList, lRow, 3)) & "', '" & Trim(GetText(vasList, lRow, 4)) & "', '', 0, '', " & _
                              "'" & lsResDate & "', '" & lsSeqNo & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & Trim(GetText(vasList, lRow, 7)) & "', " & vbCrLf & _
                              "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
                              "'" & lsRes & "', 'B', '" & lsExamName & "', " & vbCrLf & _
                              "'', '', '', '', " & _
                              "'', '' ) "
                        res = SendQuery(gLocal, SQL)
                        
                        If res = -1 Then
                            SaveQuery SQL
                            'Exit Function
                        End If
                        
                        If chkMode.Value = 1 And Trim(GetText(vasList, lRow, 2)) <> "" And IsNumeric(Mid(Trim(GetText(vasList, lRow, 2)), 2)) = True Then
                            If Set_EqpResultsql(lsExamCode, lsRes, "", Trim(GetText(vasList, lRow, 2)), gInsCode) = True Then
                                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                                SetText vasList, "완료", lRow, gResCol
                                
                                vasList.Row = lRow
                                vasList.Col = 1
                                vasList.Value = 1
                                
                                Update_Sample Trim(GetText(vasList, lRow, 2))
                                'DeleteWorkList lsID
                                
                                SQL = "delete from pat_res " & vbCrLf & _
                                      "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                                      "  AND barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' " & vbCrLf & _
                                      "  And equipcode = '" & lsEquipCode & "' " & vbCrLf & _
                                res = SendQuery(gLocal, SQL)
                                If res = -1 Then
                                    SaveQuery SQL
                                    Exit Sub
                                End If
                            Else
                                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                                SetText vasList, "실패", lRow, gResCol
                            End If
                        End If
                        
                        Exit For
                    End If
                Next j
                
            Else
                lsRes = ""
            End If
        
            PreRow = lRow
            PreRack = lsRing
            PrePos = lsTube
        
        Case "14"   'Quantitative Raw Data
        Case "15"   'Quantitative QS Raw Data
        Case "16"   'Quantitative Raw Data ID
        Case "17"   'Result Print/Send Status
        Case "20"   'Quantitative QS/Control Values
        Case "99"   'Result / Order Manipulation Response
        'Status processing
        Case "00"   'State Selection
        Case "41"   'A-ring Load
        Case "42"   'Reagent Load
        Case "43"   'Cassette Load
        Case "90"   'System Status
        Case "91"   'TC Status
        Case "92"   'DP Status
        Case "95"   'File Summary
        'Control processing
        Case "98"   'Protocol Software Version
        Case "99"   'General Response/Error Code
        End Select
        i = InStr(1, myVar, chrLF)
    Loop
End Sub

Sub SetResult(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As Integer
    Dim sTmp As String
    Dim sFormat As String
    
    If Not IsNumeric(gArrExamRes(aiRow).res) Then
        Exit Sub
    End If

    iFloat = gArrExam(aiItem, 5)

    If iFloat = 0 Then
        gArrExamRes(aiRow).res = CStr(CCur(gArrExamRes(aiRow).res))
    Else
        If IsNumeric(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat)) Then
            sTmp = CCur(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat))
        Else
            sTmp = "0"
        End If
        
        gArrExamRes(aiRow).res = sTmp & "." & Right(gArrExamRes(aiRow).res, iFloat)
        'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
        'Else
        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
        'End If

    End If

    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(gArrExam(aiItem, 6)) Then
        If CCur(gArrExam(aiItem, 6)) > gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "L"
        End If
    End If
    
    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(gArrExam(aiItem, 7)) Then
        If CCur(gArrExam(aiItem, 7)) < gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "H"
        End If
    End If

    iFloat = gArrExam(aiItem, 8)
    If IsNumeric(iFloat) Then
        If CInt(iFloat) = 0 Then
            sFormat = "#0"
        ElseIf CInt(iFloat) > 0 Then
            sFormat = ""
            sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
            sFormat = "0." & sFormat
        End If
        If IsNumeric(gArrExamRes(aiRow).res) Then
            gArrExamRes(aiRow).res = Format(CCur(gArrExamRes(aiRow).res), sFormat)
        End If
    End If

End Sub

Sub SetResult1(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As String
    Dim sTmp As String
    Dim sFormat As String
    
    If Not IsNumeric(gArrExamRes(aiRow).res) Then
        Exit Sub
    End If
    
    iFloat = Trim(GetText(vasTemp, aiItem, 5))
    If IsNumeric(gArrExamRes(aiRow).res) Then
        gArrExamRes(aiRow).res = Format(gArrExamRes(aiRow).res, "00000")
    End If
    If iFloat = 0 Then
        gArrExamRes(aiRow).res = CStr(CCur(gArrExamRes(aiRow).res))
    Else
        If IsNumeric(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat)) Then
            sTmp = CCur(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat))
        Else
            sTmp = "0"
        End If
        
        gArrExamRes(aiRow).res = sTmp & "." & Right(gArrExamRes(aiRow).res, iFloat)
        'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
        'Else
        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
        'End If
    End If
    
    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(GetText(vasTemp, aiItem, 6)) Then
        If CCur(GetText(vasTemp, aiItem, 6)) > gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "L"
        End If
    End If
    
    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(GetText(vasTemp, aiItem, 7)) Then
        If CCur(GetText(vasTemp, aiItem, 7)) < gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "H"
        End If
    End If

    iFloat = Trim(GetText(vasTemp, aiItem, 8))

    If IsNumeric(iFloat) Then
        If CInt(iFloat) = 0 Then
            sFormat = "#0"
        ElseIf CInt(iFloat) > 0 Then
            sFormat = ""
            sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
            sFormat = "0." & sFormat
        End If
        If IsNumeric(gArrExamRes(aiRow).res) Then
            gArrExamRes(aiRow).res = Format(CCur(gArrExamRes(aiRow).res), sFormat)
        End If
    End If

End Sub

Function Save_Local_One(ByVal asRow As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Date
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND examcode = '" & Trim(GetText(vasList, asRow, 9)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "insert into pat_res (examdate, equipno, barcode, " & vbCrLf & _
          "receno, pid, pname,  " & vbCrLf & _
          "resdate, seqno, posno,  " & vbCrLf & _
          "equipcode, examcode, examname,  " & vbCrLf & _
          "result, result_copy, result_iu, sendflag, unit, recedate," & vbCrLf & _
          "MTBA,MTBB,MTBC,MTBD,RifA,RifB,RifC,RifD,MTBRemark, " & vbCrLf & _
          "MTBE,RifE, " & vbCrLf & _
          "StartDate,EndDate, CartNo, " & vbCrLf & _
          "ReagentNo, ExpDate, ErrorString1, ErrorString2, " & vbCrLf & _
          "diskno, SPCMTB, SPCRif,"
    
    SQL = SQL & vbCrLf & "QC1, QC2,"
    
    SQL = SQL & vbCrLf & "ProAPt, ProBPt, ProCPt, ProDPt, ProEPt,"
    SQL = SQL & vbCrLf & "SPCPt, QC1Pt, QC2Pt, "
    
    SQL = SQL & vbCrLf & "ProARes, ProBRes, ProCRes, ProDRes, ProERes,"
    SQL = SQL & vbCrLf & "SPCRes, QC1Res, QC2Res, "
    
    SQL = SQL & vbCrLf & "Assay, "
    
    SQL = SQL & vbCrLf & "SEXAGE, MEDDEPT,"
    SQL = SQL & vbCrLf & "PATIO, TESTDATE,"
    
    SQL = SQL & vbCrLf & "ProACK, ProBCK, ProCCK, ProDCK, ProECK,"
    SQL = SQL & vbCrLf & "SPCCK, QC1CK, QC2CK) "
    
    
    SQL = SQL & vbCrLf & "values('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', '" & Trim(GetText(vasList, asRow, 2)) & "',  " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 5)) & "', '" & Trim(GetText(vasList, asRow, 3)) & "', '" & Trim(GetText(vasList, asRow, 4)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 14)) & "', '', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 15)) & "', '" & Trim(GetText(vasList, asRow, 9)) & "', '" & Trim(GetText(vasList, asRow, 10)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 13)) & "', '" & Trim(GetText(vasList, asRow, 12)) & "', '" & Trim(GetText(vasList, asRow, 13)) & "', '" & Trim(asSend) & "', '" & gResFileName & "', '" & Trim(GetText(vasList, asRow, 14)) & "',"
          
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colMTBPA)) & "','" & Trim(GetText(vasList, asRow, colMTBPB)) & "','" & Trim(GetText(vasList, asRow, colMTBPC)) & "','" & Trim(GetText(vasList, asRow, colMTBPD)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colRifPA)) & "','" & Trim(GetText(vasList, asRow, colRifPB)) & "','" & Trim(GetText(vasList, asRow, colRifPC)) & "','" & Trim(GetText(vasList, asRow, colRifPD)) & "','" & Trim(GetText(vasList, asRow, colRemark)) & "',"
    
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colMTBPE)) & "','" & Trim(GetText(vasList, asRow, colRifPE)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colStartDate)) & "','" & Trim(GetText(vasList, asRow, colEndDate)) & "','" & Trim(GetText(vasList, asRow, colCartNo)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colReagentNo)) & "','" & Trim(GetText(vasList, asRow, colExpDate)) & "','" & Trim(GetText(vasList, asRow, colError1)) & "','" & Trim(GetText(vasList, asRow, colError2)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colRack)) & "','" & Trim(GetText(vasList, asRow, colMTBSPC)) & "','" & Trim(GetText(vasList, asRow, colRifSPC)) & "',"
    
    '20140728 추가
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colQC1Ct)) & "','" & Trim(GetText(vasList, asRow, colQC2Ct)) & "',"
    
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colProAPt)) & "','" & Trim(GetText(vasList, asRow, colProBPt)) & "','" & Trim(GetText(vasList, asRow, colProCPt)) & "','" & Trim(GetText(vasList, asRow, colProDPt)) & "','" & Trim(GetText(vasList, asRow, colProEPt)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colSPCPt)) & "','" & Trim(GetText(vasList, asRow, colQC1Pt)) & "','" & Trim(GetText(vasList, asRow, colQC2Pt)) & "',"
    
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colProARes)) & "','" & Trim(GetText(vasList, asRow, colProBRes)) & "','" & Trim(GetText(vasList, asRow, colProCRes)) & "','" & Trim(GetText(vasList, asRow, colProDRes)) & "','" & Trim(GetText(vasList, asRow, colProERes)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colSPCRes)) & "','" & Trim(GetText(vasList, asRow, colQC1Res)) & "','" & Trim(GetText(vasList, asRow, colQC2Res)) & "',"
    
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colAssay)) & "',"
    
    '20140807 추가
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colSexAge)) & "','" & Trim(GetText(vasList, asRow, colMedDept)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colIO)) & "','" & Trim(GetText(vasList, asRow, colTestDate)) & "',"
    
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colProACheck)) & "','" & Trim(GetText(vasList, asRow, colProBCheck)) & "','" & Trim(GetText(vasList, asRow, colProCCheck)) & "','" & Trim(GetText(vasList, asRow, colProDCheck)) & "','" & Trim(GetText(vasList, asRow, colProECheck)) & "',"
    SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, asRow, colSPCCheck)) & "','" & Trim(GetText(vasList, asRow, colQC1Check)) & "','" & Trim(GetText(vasList, asRow, colQC2Check)) & "')"
    
    res = SendQuery(gLocal, SQL)
    'Save_Raw_Data "[QUERY]" & SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

'Function Save_Local_One(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = GetDateFull
'
'    sCnt = ""
'    SQL = "Delete FROM pat_res " & vbCrLf & _
'          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND equipcode = '" & gArrExamRes(aiIndex).EquipCode & "'" & vbCrLf & _
'          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
'
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
'            "barcode, examtype, receno, " & _
'            "pid, pname, pjumin, page, psex, " & _
'            "resdate, seqno, diskno, posno, " & _
'            "equipcode, examcode, " & _
'            "result, sendflag, examname, " & _
'            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
'          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
'          "'" & Trim(GetText(vasList, asRow, 3)) & "', '" & Trim(GetText(vasList, asRow, 4)) & "', '', 0, '', " & _
'          "'" & sExamDate & "', '" & gArrExamRes(aiIndex).SeqNo & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
'          "'" & gArrExamRes(aiIndex).EquipCode & "', '" & gArrExamRes(aiIndex).ExamCode & "', " & _
'          "'" & gArrExamRes(aiIndex).res & "', '" & asSend & "', '" & gArrExamRes(aiIndex).ExamName & "', " & vbCrLf & _
'          "'" & gArrExamRes(aiIndex).RefFlag & "', '', '', '', " & _
'          "'" & gArrExamRes(aiIndex).RefLow & " ~ " & gArrExamRes(aiIndex).RefHigh & "', '' ) "
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function

Function Save_Local_One_1(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExam(aiIndex, 1) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, examtype, receno, pid, " & _
          "pname, pjumin, page, psex, resdate, seqno, diskno, posno, " & _
          "equipcode, examcode, examtype, result, sendflag, examname, " & _
          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, 3)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '', " & _
          "0, '', " & _
          "'" & sExamDate & "', '" & gArrExam(aiIndex, 4) & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & gArrExam(aiIndex, 1) & "', '" & gArrExam(aiIndex, 2) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, gResCol + aiIndex)) & "', '" & asSend & "', '" & gArrExam(aiIndex, 3) & "', " & vbCrLf & _
          "'', '', " & _
          "'', '', " & _
          "'', '') "
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Update_Sample(ByVal asID As String)
    SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & asID & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function DeleteWorkList(ByVal asID As String)
    SQL = "Delete from WorkList where Barcode ='" & asID & "'"
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Public Function Set_EqpResultsql(ByVal Testcd As String, ByVal EqpRst As String, ByVal ErrDes As String, ByVal SPCID As String, ByVal INS_CODE As String) As Boolean
On Error GoTo errtrap
    'Set cmdSQL = New ADODB.Command
    Dim sDate As String
    
    sDate = GetDateFull
    
    DoSleep 5
    INS_CODE = gInsCode
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceResult_INSERT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(SPCID))
        .Parameters.Append .CreateParameter("@i_itemCode", adVarChar, adParamInput, 10, Trim(Testcd))
        .Parameters.Append .CreateParameter("@i_transTimestamp", adChar, adParamInput, 19, sDate)
        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(EqpRst))
        .Parameters.Append .CreateParameter("@i_instrumentCode", adChar, adParamInput, 2, Trim(INS_CODE))
        .Parameters.Append .CreateParameter("@i_errorDescription", adVarChar, adParamInput, 100, Trim(ErrDes))
        
        .Execute
    End With
    
    If cmdSQL("retval").Value = 2 Then
        Set_EqpResultsql = False
        MsgBox "결과전송 실패", vbInformation, "알림"
        'Set cmdSQL = Nothing
        Exit Function
    End If
    
    Set_EqpResultsql = True
    'Set cmdSQL = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    'Err.Raise Err.Number, Err.Description
End Function

Public Sub SaveRes(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    Open App.Path & "\Log\Res.log" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Private Sub Picture1_Click()
    frmUser.Show 0
End Sub

Private Sub subClear_Click()
    
    gRow = 0
    
    PreRack = ""
    PrePos = ""
    PreRow = 0
    
    txtBuff = ""
    
    gCurRow = -1
    ReDim gArrExamRes(0)
    GetExamCode
    
'vsSpread의 내용을 Clear 한다.
    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = vasList.MaxCols
    vasList.BlockMode = True
    vasList.Action = 3
    vasList.BackColor = RGB(255, 255, 255)
    vasList.ForeColor = RGB(0, 0, 0)
    vasList.BlockMode = False

    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = 1
    vasList.BlockMode = True
    vasList.Value = 0
    vasList.BlockMode = False
    
    vasList.MaxRows = 0
    
End Sub

Private Sub subClose_Click()
    Unload Me
End Sub

Private Sub subCodeSet_Click()
    frmCode.Show 1
End Sub


Private Sub subComSetup_Click()
    frmConfig.Show 1
End Sub


Private Sub subSend1_Click()
    subSend1.Checked = True
    subSend2.Checked = False
    
    chkMode.Value = 1
    SaveSetting "MEDIMATE", "COBRA", "SendMode", "1"
End Sub

Private Sub subSend2_Click()
    subSend1.Checked = False
    subSend2.Checked = True
    
    chkMode.Value = 0
    SaveSetting "MEDIMATE", "COBRA", "SendMode", "0"

End Sub

Private Sub Timer1_Timer()
        If MSComm1.CTSHolding = True Then
            lblCA.ForeColor = RGB(0, 255, 0)
        Else
            lblCA.ForeColor = RGB(0, 0, 255)
        End If
        
        If dtpExamDate <> Date Then
            dtpExamDate = Date
            
        End If
End Sub

Private Sub txtBuff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        COBAS_Amplicor
    End If
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    iRow1 = BlockRow
    iRow2 = BlockRow2
    iCol1 = BlockCol
    iCol2 = BlockCol2
    
    txtSSeq.Text = iRow1
    txtESeq.Text = iRow2
    
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        iRow1 = 0
        iRow2 = 0
'        Select Case Col
'        Case 2
'            vasSort vasList, Col
'        Case 3
'            vasSort vasList, Col
'        Case 4
'            vasSort vasList, Col, 2
'        Case 5
'            vasSort vasList, Col, 6
'        Case 6
'            vasSort vasList, Col, 5
'        Case 7
'            vasSort vasList, Col, 2
'        End Select
    Else
        iRow1 = Row
        iRow2 = Row
    End If
    
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim lRow, lCol As Long
'
'    If Row < 1 Or Row > vasList.DataRowCnt Then Exit Sub
'
'    txtID = ""
'    txtPID = ""
'    txtPName = ""
'    txtResDate = ""
'    ClearSpread vasRes1
'    ClearSpread vasRes2
'
'    txtID = Trim(GetText(vasList, Row, 2))
'    txtPID = Trim(GetText(vasList, Row, 3))
'    txtPName = Trim(GetText(vasList, Row, 4))
'    txtRack = Trim(GetText(vasList, Row, 6))
'    txtTube = Trim(GetText(vasList, Row, 7))
'    txtEquip = Trim(GetText(vasList, Row, 5))
'
'    SQL = "Select resdate from pat_res " & vbCrLf & _
'          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND barcode = '" & Trim(txtID) & "' "
'    res = db_select_Text(gLocal, SQL, txtResDate)
'
'    lCol = gResCol
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vasList, lRow, lCol)) <> "" Then
'
'            SetText vasRes1, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText vasRes1, Trim(GetText(vasList, lRow, lCol)), lRow, 3
'            SetText vasRes1, Trim(GetText(vasList, 0, lCol)), lRow, 2
'
'            vasList.Row = Row
'            vasList.Col = lCol
'            Select Case vasList.BackColor
'            Case RGB(255, 127, 0)
'                SetForeColor vasRes1, lRow, lRow, lCol, lCol, 255, 127, 0
'                SetText vasRes1, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor vasRes1, lRow, lRow, lCol, lCol, 0, 127, 255
'                SetText vasRes1, "▼", lRow, 4
'            Case Else
'                SetText vasRes1, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vasList, lRow, lCol)) <> "" Then
'
'            SetText vasRes2, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText vasRes2, Trim(GetText(vasList, lRow, lCol)), lRow, 3
'            SetText vasRes2, Trim(GetText(vasList, 0, lCol)), lRow, 2
'
'            vasList.Row = Row
'            vasList.Col = lCol
'            Select Case vasList.BackColor
'            Case RGB(255, 127, 0)
'                SetForeColor vasRes2, lRow, lRow, lCol, lCol, 255, 127, 0
'                SetText vasRes2, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor vasRes2, lRow, lRow, lCol, lCol, 0, 127, 255
'                SetText vasRes2, "▼", lRow, 4
'            Case Else
'                SetText vasRes2, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    Frame1.Visible = True
'
End Sub

Sub GetComSetup()
    Dim db_tmp As String * 100
    Dim lRow As Long
       
    lRow = 0
        
    db_tmp = ""
    Call GetPrivateProfileString("COM", "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.ComPort = Trim(txtTemp)
                                            
    Call GetPrivateProfileString("COM", "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.Speed = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.Parity = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.DataBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.StopBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.RTSEnable = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.DTREnable = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerPath = Trim(txtTemp)
    
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gIFUser = Trim(txtTemp)
    

End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
'''    Dim lRow, i, liEquipCode As Long
'''    Dim lsID As String
'''    Dim liRet As Integer
'''    'Dim lsID As String
'''    Dim lsResult As String
'''
'''
'''    If KeyCode = vbKeyReturn Then
'''        lRow = vasList.ActiveRow
'''
'''        SQL = "Select barcode, diskno, posno from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
'''        res = db_select_Col(gLocal, SQL)
'''        If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 2)) Then
'''            If MsgBox("입력하신 검체 [" & Trim(GetText(vasList, lRow, 2)) & "]는 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
'''                      " " & vbCrLf & _
'''                      "결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
'''                SetText vasList, Trim(GetText(vasList, lRow, gMaxCol + 4)), lRow, 2
'''                Exit Sub
'''            End If
'''        End If
'''
'''        lsID = Trim(GetText(vasList, lRow, 2))
'''
'''        SQL = "Update pat_res set barcode = '" & lsID & "' " & vbCrLf & _
'''              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'''              "  and barcode = '" & Trim(GetText(vasList, lRow, gMaxCol + 4)) & "'"
'''        res = SendQuery(gLocal, SQL)
'''
'''        liRet = 1
'''
'''        For liEquipCode = 1 To vasList.MaxCols
'''            lsResult = Trim(GetText(vasList, lRow, liEquipCode + gResCol))
'''
'''            For i = 1 To UBound(gArrExam)
'''                If CInt(gArrExam(i, 1)) = liEquipCode Then
'''                    If Trim(gArrExam(i, 2)) <> "" Then
'''
'''                        If Set_EqpResultsql(gArrExam(i, 2), lsResult, "", lsID, "78") Then
'''                        Else
'''                            liRet = -1
'''                        End If
'''                    End If
'''
'''                    Exit For
'''                End If
'''            Next i
'''        Next liEquipCode
'''
'''        If liRet = 1 Then
'''            SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
'''            SetText vasList, "완료", lRow, gResCol
'''
'''            vasList.Row = lRow
'''            vasList.Col = 1
'''            vasList.Value = 1
'''
'''            Update_Sample Trim(GetText(vasList, lRow, 2))
'''            DeleteWorkList Trim(GetText(vasList, lRow, 2))
'''
'''        Else
'''            SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
'''            SetText vasList, "실패", lRow, gResCol
'''        End If
'''
'''    End If
End Sub

Private Sub vasList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim i As Long
    
    If iRow1 < 0 And iRow2 < 0 Then
        iRow1 = Row
        iRow2 = Row
    End If
    
    For i = iRow1 To iRow2
        vasList.Row = i
        vasList.Col = 1
        vasList.Value = 1
    Next i
    vasList.BlockMode = False

End Sub

Private Sub vasSch_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lRow, lCol As Long
    Dim argSpread As vaSpread
    
    If Row < 1 Or Row > vasSch.DataRowCnt Then Exit Sub
    
    SelVas = 2
    
    If Row = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf Row = vasSch.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If vasSch.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(vasSch, Row, 2))
    txtPID = Trim(GetText(vasSch, Row, 3))
    txtPName = Trim(GetText(vasSch, Row, 4))
    txtRack = Trim(GetText(vasSch, Row, 6))
    txtTube = Trim(GetText(vasSch, Row, 7))
    txtEquip = Trim(GetText(vasSch, Row, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    
    
    'lCol = gResCol
    lRow = 0
    For lCol = gResCol + 1 To gResCol + 35
        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argSpread = vasRes1
            Else
                Set argSpread = vasRes2
            End If
            If lRow = 21 Then lRow = 1
            
            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
            
            vasSch.Row = Row
            vasSch.Col = lCol
            Select Case vasSch.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
                SetText argSpread, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
                SetText argSpread, "▼", lRow, 4
            Case Else
                SetText argSpread, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            vasActiveCell vasSch, lRow, lCol
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, lRow, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
    
    Frame1.Visible = True

End Sub

'WinSock Control ==============================================================================================================
Public Sub WinSock_Listen(argWinSock As Winsock)
    Dim sWinSockPort As String
    
    
    sWinSockPort = gDRDB_Parm.ServerPort
    
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    argWinSock.LocalPort = sWinSockPort
    argWinSock.Listen
    
'''    If EquipNum = 1 Then
'''        lblConnect1.Caption = "연결 대기중..."
'''    Else
'''        lblConnect2.Caption = "연결 대기중..."
'''    End If
    
End Sub

Private Sub Winsock1_Close()
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.LocalPort = gDRDB_Parm.ServerPort
    Winsock1.Listen
    
    frmInterface.Caption = "Xpert Interface Program : 연결 대기중...."
    'Xpert Interface Program
'''    lblConnect1.Caption = "연결 대기중..."
    
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Accept requestID
    frmInterface.Caption = "Xpert Interface Program : 연결[" & requestID & "]" & Winsock1.RemoteHostIP
'''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

'''Public Function HL7_Ack(argMSH As String) As String
'''    Dim strMSH As String
'''    Dim strACK As String
'''    Dim strDateTime As String
'''    Dim strSplit() As String
'''    Dim strSigNum As String
'''
'''    Dim i As Integer
'''    Dim j As Integer
'''
'''    strMSH = argMSH
'''    strSplit = Split(strMSH, "|")
'''
''''''    MSH|^~\&|cobas 8000||host||20130104114005||OUL^R22^REAL|31777||2.5||||AA||UNICODE UTF-8|
'''    strDateTime = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
'''
'''    strACK = Chr(11)
'''    strACK = strACK & "MSH|^~\&|" & Trim(strSplit(5)) & "||" & Trim(strSplit(3)) & "||" & strDateTime & "||ACK|" & CStr(gMSGSeq) & "||" & Trim(strSplit(3)) & "||||AA||" & Trim(strSplit(18)) & "|" & vbCr
'''    strACK = strACK & "MSA|AA|" & Trim(strSplit(10)) & "||" & vbCr
'''    strACK = strACK & Chr(28) & vbCr
'''    gMSGSeq = gMSGSeq + 1
'''
'''End Function



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String
    Dim strResData As String
    Dim strMsgSplit() As String
    Dim strACK As String
    Dim i As Integer
    Dim strMDateTime As String

    Winsock1.GetData sTmp

    
    Save_Raw_Data "[" & Format(Time, "hh:mm:ss") & "]" & sTmp
    
    txtData.Text = txtData.Text & sTmp
    
    If InStr(1, sTmp, chrLF) > 0 Then
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        Winsock1.SendData chrACK
    ElseIf InStr(1, sTmp, chrENQ) > 0 Then
        txtData.Text = sTmp
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        Winsock1.SendData chrACK
    ElseIf InStr(1, sTmp, chrEOT) > 0 Then
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtData.Text
        XPert_All txtData.Text
        txtData.Text = ""
        
    End If
    
    
    
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblIPState.Caption = "[Error]" & Number & " : " & Description
End Sub


Private Sub XPert_All(argString)
    Dim strData As String
    Dim strSub() As String
    Dim arrData
    Dim strTemp As String
    Dim intTempSeq As Integer
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    intTempSeq = 0
    strData = argString
    
    
    
    intTempSeq = InStr(1, strData, chrSTX)
    strTemp = Mid(strData, intTempSeq + 2)
    
    
    For j = 1 To 100
        i = InStr(1, strTemp, chrSTX)
        K = InStr(1, strTemp, chrETB)
        
        If K = 0 Then
            i = InStr(1, strTemp, chrETX)
            strTemp = Mid(strTemp, 1, i - 1)
            Exit For
        Else
            strTemp = Mid(strTemp, 1, K - 1) & Mid(strTemp, i + 2)
        End If
        
    Next j
    
    
    
'    strData = Replace(strData, chrENQ, "")
'
''    i = InStr(1, strData, chrSTX)
''
''    While i > 0
''        strData = Mid(strData, 1, i - 1) & Mid(strData, i + 2)
''        i = InStr(1, strData, chrSTX)
''    Wend
''
''    i = InStr(1, strData, chrETB)
''
''    While i > 0
''        strData = Mid(strData, 1, i - 1) & Mid(strData, i + 5)
''        i = InStr(1, strData, chrETB)
''    Wend
''
''    i = InStr(1, strData, chrETX)
''
''    While i > 0
''        strData = Mid(strData, 1, i - 1) & Mid(strData, i + 5)
''        i = InStr(1, strData, chrETX)
''    Wend
''
''    strData = Replace(strData, chrEOT, "")
    
    strSub = Split(strTemp, vbCr)
    
    For i = 1 To UBound(strSub)
        XPert strSub(i - 1)
    Next
    
    
End Sub


Private Sub XPert(argData As String)
    Dim strBarcode As String
    Dim strResult As String
    Dim strRealRes As String
    Dim strValue As String
    Dim strExamCode As String
    Dim strExamName As String
    Dim strEquipCode As String
    
    Dim strFlag As String
    Dim strSeqNo As String
    
    Dim arrModule() As String
    
    Dim strModule As String
    Dim srtComment As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strExpDate As String
    Dim strCartNo As String
    Dim strReagnetNo As String
    
    Dim strErrorComment As String
    
    
    Dim iDer As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim strSub() As String
    Dim strDel() As String
    Dim strData As String
    Dim strHeader As String
    Dim lsResRow As Integer
    
    Dim i2 As Integer
    Dim sReceExamCode As String
    Dim sRv             As String
    
    Dim strSend As String
    Dim sParam As String
    
    Dim intEquipCode As Integer
    
    
    strData = argData
    
    strSub = Split(strData, "|")
    
    strHeader = strSub(0)
    
    Select Case strHeader
    
    Case "H"
        strToxCheck = ""
    Case "P"
    Case "O"
        '장비코드를 변수를 초기화 한다.
        gEquipCode = ""
        
        strBarcode = Trim(strSub(2))
        
        gRow = -1
        For i = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, i, colBarcode)) = strBarcode Then
                gRow = i
                Exit For
            End If
        Next
        If gRow = -1 Then
            gRow = vasList.DataRowCnt + 1
            If gRow > vasList.MaxRows Then
                vasList.MaxRows = gRow
            End If
        End If
        
        SetText vasList, strBarcode, gRow, colBarcode
        If Trim(GetText(vasList, gRow, colPName)) = "" Then
            If Len(strBarcode) > 10 And IsNumeric(strBarcode) = True Then
                Clear_XML_Exam
                Get_Sample_Info gRow
            End If
        End If
        
        
        vasList_Click colBarcode, gRow
        
'        If Trim(strSub(15)) = "ORH" Then
'            SetText vasList, "Other", gRow, colPos
'        Else
'            SetText vasList, Trim(strSub(15)), gRow, colPos
'        End If
    Case "R"
        strEquipCode = Trim(strSub(2))
        strDel = Split(strEquipCode, "^")
        If UBound(strDel) < 6 Then
            Exit Sub
        End If
        
        Dim strSeq As String
        Dim strToxEquip As String
        strSeq = "0"
        strToxEquip = ""
        If Trim(strSub(1)) <> "1" And CCur(strSub(1)) < 14 And Trim(strDel(6)) = "Tox" Then
            strSeq = (CCur(Trim(strSub(1))) + 1) Mod 3 + 1
            
            If strSeq = 2 Then
                strToxEquip = "Ct"
            ElseIf strSeq = 3 Then
                strToxEquip = "EndPt"
            End If
            
        End If
        
        If UBound(strSub) > 12 Then
            arrModule = Split(Trim(strSub(13)), "^")
            strModule = arrModule(2)
            
            '모듈번호를 보기쉽게 변경한다.
            If strModule = "614414" Then
                strModule = "B1"
                SetText vasList, strModule, gRow, colRack
            ElseIf strModule = "614415" Then
                strModule = "B2"
                SetText vasList, strModule, gRow, colRack
            ElseIf strModule = "619205" Then
                strModule = "B3"
                SetText vasList, strModule, gRow, colRack
            ElseIf strModule = "633147" Then
                strModule = "B4"
                SetText vasList, strModule, gRow, colRack
            Else
                SetText vasList, strModule, gRow, colRack
            End If
            
            strStartDate = Trim(strSub(11))
            strEndDate = Trim(strSub(12))
            strCartNo = arrModule(3)
            strReagnetNo = arrModule(4)
            strExpDate = arrModule(5)
            
            SetText vasList, strStartDate, gRow, colStartDate
            SetText vasList, strEndDate, gRow, colEndDate
            SetText vasList, strCartNo, gRow, colCartNo
            SetText vasList, strReagnetNo, gRow, colReagentNo
            SetText vasList, strExpDate, gRow, colExpDate
            
            SetText vasList, Mid(strStartDate, 1, 8), gRow, colTestDate
            
        End If
        
        Dim intTox_B As Integer
        
        strEquipCode = Trim(strDel(3)) & "/" & Trim(strDel(6)) '& "/" & Trim(strDel(7))
        
        '버전이 업그레이드 될때마다 신호가 바뀌게 되므로 강제로 코드를 만들어줌.
        If strEquipCode = "3/Toxigenic C" Then
            strToxCheck = "diff/Tox"
        ElseIf strEquipCode = "2/027-NAP1-BI" Then
            strToxCheck = ""
            strEquipCode = "2/027"
        End If
        
        If strToxCheck = "diff/Tox" Then
            If Trim(strSub(1)) = "1" Then
                strEquipCode = "diff/Tox"
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
            Else
                intEquipCode = Trim(strSub(1)) Mod 3 - 1
                
                If Trim(strDel(6)) = "Toxin B" Then
                    strEquipCode = "diff/Tox"
                    strEquipCode = strEquipCode & "/Toxin B"
                ElseIf Trim(strDel(6)) = "Binary Toxin" Then
                    strEquipCode = "diff/Tox"
                    strEquipCode = strEquipCode & "/Binary Toxin"
                ElseIf Trim(strDel(6)) = "TcdC" Then
                strEquipCode = "diff/Tox"
                    strEquipCode = strEquipCode & "/TcdC"
                ElseIf Trim(strDel(6)) = "SPC" Then
                    strEquipCode = "diff/Tox"
                    strEquipCode = strEquipCode & "/SPC"
                End If
                
                If intEquipCode = 1 Then
                    strEquipCode = strEquipCode & "/AnalResult"
                ElseIf intEquipCode = -1 Then
                    strEquipCode = strEquipCode & "/Ct"
                ElseIf intEquipCode = 0 Then
                    strEquipCode = strEquipCode & "/EndPt"
                End If
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i = 1 Then
                    strRealRes = Mid(strRealRes, i + 1)
                ElseIf i > 1 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
            End If
        ElseIf strEquipCode = "2/027" Then
            If Trim(strSub(1)) = "14" Then
                strEquipCode = "027/027/"
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = "027 " & Mid(strRealRes, 1, i - 1)
                End If
            End If

        
        ElseIf Mid(strEquipCode, 1, 1) = "1" Or Mid(strEquipCode, 1, 1) = "2" Or Mid(strEquipCode, 1, 1) = "3" Then
            
            If Mid(strEquipCode, 1, 1) = "1" Then
                strEquipCode = "MTB" & Mid(strEquipCode, 2)
            ElseIf Mid(strEquipCode, 1, 1) = "2" Then
                strEquipCode = "QC" & Mid(strEquipCode, 2)
            ElseIf Mid(strEquipCode, 1, 1) = "3" Then
                strEquipCode = "RIF" & Mid(strEquipCode, 2)
            End If
            
            
            
            If Trim(strSub(1)) = "1" Then
                If strEquipCode = "MTB/MTB" Then
                    strEquipCode = "MTB/MTB/"
                Else
                    strEquipCode = strEquipCode
                End If
                
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
            ElseIf Trim(strSub(1)) = "20" Then
                If strEquipCode = "RIF/Rif Resistance" Then
                    strEquipCode = "Rif/Rif Resistance/"
                Else
                    strEquipCode = strEquipCode
                End If
                
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
                
                
                '결과가 없는경우 만들어준다.
                If strRealRes = "" Then
                    'strRealRes = "Rif Resistance DETECTED"
                End If
                
            Else
                'intEquipCode = Trim(strSub(1)) Mod 3 - 1
                strEquipCode = strEquipCode & "/" & Trim(strDel(7))
'''                If intEquipCode = 1 Then
'''                    strEquipCode = strEquipCode & "/"
'''                ElseIf intEquipCode = -1 Then
'''                    strEquipCode = strEquipCode & "/Ct"
'''                ElseIf intEquipCode = 0 Then
'''                    strEquipCode = strEquipCode & "/EndPt"
'''                End If
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i = 1 Then
                    strRealRes = Mid(strRealRes, i + 1)
                ElseIf i > 1 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
            End If
        
        Else
            If Trim(strDel(7)) = "Ct" Or Trim(strDel(7)) = "EndPt" Then
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = Mid(strRealRes, i + 1)
                End If
            Else
                strRealRes = Trim(strSub(3))
                i = InStr(1, strRealRes, "^")
                If i > 0 Then
                    strRealRes = Mid(strRealRes, 1, i - 1)
                End If
                
                If InStr(1, strRealRes, "MTB DETECTED") > 0 Then
                    If Len("MTB DETECTED") < Len(Trim(strRealRes)) Then
                        strRealRes = Mid(strRealRes, 1, 12) & " (" & Trim(Mid(strRealRes, 13)) & ")"
                    End If
                End If
            End If
        
        End If
        
        
        
        If InStr(1, strEquipCode, "diff/Tox") > 0 Then
            'Toxigenic C.diff POSITIVE
            If InStr(1, strRealRes, "POSITIVE") > 0 Then
                strResult = "Toxigenic C.diff : POS"
            ElseIf InStr(1, strRealRes, "NEGATIVE") > 0 Then
                strResult = "Toxigenic C.diff : NEG"
        
            ElseIf InStr(1, strRealRes, "027 PRESUMPTIVE POS") > 0 Then
                strResult = "027 PRESUMPTIVE  : POS"
            ElseIf InStr(1, strRealRes, "027 PRESUMPTIVE NEG") > 0 Then
                strResult = "027 PRESUMPTIVE  : NEG"
            
            ElseIf strRealRes = "0" Then
                strResult = "0.0"
            Else
                strResult = strRealRes
            End If
        
        ElseIf strEquipCode = "027/027/" Then
            If InStr(1, strRealRes, "027 PRESUMPTIVE POS") > 0 Then
                strResult = "027 PRESUMPTIVE  : POS"
            ElseIf InStr(1, strRealRes, "027 PRESUMPTIVE NEG") > 0 Then
                strResult = "027 PRESUMPTIVE  : NEG"
            
            ElseIf strRealRes = "0" Then
                strResult = "0.0"
            Else
                strResult = strRealRes
            End If

        ElseIf InStr(1, strEquipCode, "MTB/MTB/") > 0 Then
            '결과 형태 처리해야함 ==============================================================
            If InStr(1, strRealRes, "NOT DETECTED") > 0 Then
                strResult = Replace(strRealRes, "NOT DETECTED", "MTB           : Not detected") & "/" & "RIF resistance: Not detected"
            ElseIf InStr(1, strRealRes, "DETECTED") > 0 Then
                strResult = Replace(strRealRes, "DETECTED", "MTB           : Detected")
            ElseIf InStr(1, strRealRes, "Rif Resistance DETECTED") > 0 Then
                strResult = Replace(strRealRes, "Rif Resistance DETECTED", "RIF resistance: Detected")
            ElseIf InStr(1, strRealRes, "Rif Resistance NOT DETECTED") > 0 Then
                strResult = Replace(strRealRes, "Rif Resistance NOT DETECTED", "RIF resistance: Not detected")

            ElseIf strRealRes = "0" Then
                strResult = "0.0"
            Else
                strResult = strRealRes
            End If
        ElseIf InStr(1, strEquipCode, "Rif/Rif Resistance/") > 0 Then
            '결과 형태 처리해야함 ==============================================================
            If InStr(1, strRealRes, "NOT DETECTED") > 0 Then
                strResult = Replace(strRealRes, "NOT DETECTED", "RIF resistance: Not detected")
            ElseIf InStr(1, strRealRes, "DETECTED") > 0 Then
                strResult = Replace(strRealRes, "DETECTED", "RIF resistance: Detected")
            ElseIf strRealRes = "0" Then
                strResult = "0.0"
            Else
                strResult = strRealRes
            End If
        Else
            If strRealRes = "0" Then
                strResult = "0.0"
            Else
                strResult = strRealRes
            End If
        End If
            
        
        'MTB DETECTED (HIGH)/Rif Resistance DETECTED
        
        '===================================================================================
        
        
        
        Select Case strEquipCode '& "/" & Trim(strDel(7))
            
            Case "MTB/MTB/"
                strEquipCode = "MTB/MTB"
                gEquipCode = "MTB/MTB"
                SetText vasList, strResult, gRow, colResult
            Case "Rif/Rif Resistance/"
                strEquipCode = ""
                If strResult <> "" Then
                    SetText vasList, GetText(vasList, gRow, colResult) & "/" & strResult, gRow, colResult
                End If
            Case "MTB/Probe A/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPA
            Case "MTB/Probe B/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPB
            Case "MTB/Probe C/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPC
            Case "MTB/Probe D/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPD
            Case "MTB/Probe E/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPE
            
            Case "Rif/Probe A/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifPA
            Case "Rif/Probe B/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifPB
            Case "Rif/Probe C/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifPC
            Case "Rif/Probe D/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifPD
            Case "Rif/Probe E/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifPE
               
            Case "MTB/SPC/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBSPC
            Case "Rif/SPC/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colRifSPC
            
            '-----------------------------------------------------------
            '20140728 추가
            
            'QC Ct
            Case "QC/QC-1/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC1Ct
            Case "QC/QC-2/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC2Ct
            
            'EndPt
            Case "MTB/Probe A/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProAPt
            Case "MTB/Probe B/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProBPt
            Case "MTB/Probe C/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProCPt
            Case "MTB/Probe D/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProDPt
            Case "MTB/Probe E/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProEPt
            Case "MTB/SPC/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colSPCPt
            Case "QC/QC-1/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC1Pt
            Case "QC/QC-2/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC2Pt
            
            
            'Analyte result
            Case "MTB/Probe A/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProARes
            Case "MTB/Probe B/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProBRes
            Case "MTB/Probe C/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProCRes
            Case "MTB/Probe D/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProDRes
            Case "MTB/Probe E/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProERes
            Case "MTB/SPC/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colSPCRes
            Case "QC/QC-1/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC1Res
            Case "QC/QC-2/"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colQC2Res
            
            
            
            '/---------------- 20150911 추가
            'Toxigenic C
            Case "diff/Tox"
                strEquipCode = "TOX/TOX"
                SetText vasList, strResult, gRow, colResult
                        
            '027
            Case "027/027/"
                strEquipCode = ""
                If strResult <> "" Then
                    SetText vasList, GetText(vasList, gRow, colResult) & "/" & strResult, gRow, colResult
                End If
            
            
            'Toxin B
            Case "diff/Tox/Toxin B/AnalResult"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProARes
            Case "diff/Tox/Toxin B/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPA
            Case "diff/Tox/Toxin B/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProAPt
                
            'Binary Toxin
            Case "diff/Tox/Binary Toxin/AnalResult"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProBRes
            Case "diff/Tox/Binary Toxin/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPB
            Case "diff/Tox/Binary Toxin/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProBPt
                
            'TcdC
            Case "diff/Tox/TcdC/AnalResult"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProCRes
            Case "diff/Tox/TcdC/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBPC
            Case "diff/Tox/TcdC/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colProCPt

                
            
            'SPC
            Case "diff/Tox/SPC/AnalResult"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colSPCRes
            Case "diff/Tox/SPC/Ct"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colMTBSPC
            Case "diff/Tox/SPC/EndPt"
                strEquipCode = ""
                SetText vasList, strResult, gRow, colSPCPt
        
            
            Case Else
                strEquipCode = ""
        End Select
        
        
        
''        If Len(Trim(GetText(vasList, gRow, colBarcode))) >= 10 And IsNumeric(Trim(GetText(vasList, gRow, colBarcode))) = True And strEquipCode <> "" Then
''            Clear_XML_Exam
''            sRv = Online_XML(gXml_S07, Trim(GetText(vasList, gRow, colBarcode)))
''            sReceExamCode = ""
''
''            For i2 = 0 To UBound(gExam_Select)
''                If sReceExamCode = "" Then
''                    sReceExamCode = "'" & Trim(gExam_Select(i2).TST_CD) & "'"
''                Else
''                    sReceExamCode = sReceExamCode & ",'" & Trim(gExam_Select(i2).TST_CD) & "'"
''                End If
''            Next i2
''        Else
''            sReceExamCode = "''"
''        End If
''
''        Save_Raw_Data "[EXAMCODE]" & sReceExamCode
        
        If strEquipCode <> "" Then
            SQL = "select examcode, examname from equipexam "
            SQL = SQL & "where equip = '" & gEquip & "' and equipcode = '" & strEquipCode & "'"
            'If sReceExamCode <> "''" Then
            '    SQL = SQL & " and examcode in ( " & sReceExamCode & ") "
            'End If
            
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 Then
                strExamCode = Trim(gReadBuf(0))
                strExamName = Trim(gReadBuf(1))
                SetText vasList, strExamCode, gRow, colExamCode
                SetText vasList, strExamName, gRow, colExamName
                
                SetText vasList, strExamName, gRow, colAssay
                
                
            Else
                SetText vasList, "NotValue", gRow, colExamCode
                SetText vasList, "NotValue", gRow, colExamName
            End If
        End If
    Case "C"
        
        
        strDel = Split(Trim(strSub(3)), "^")
        If UBound(strDel) < 3 Then
            Exit Sub
        End If
        
        
        If InStr(1, strSub(3), "[Probe A] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colProACheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[Probe B] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colProBCheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[Probe C] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colProCCheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[Probe D] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colProDCheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[Probe E] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colProECheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[SPC] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colSPCCheck
            Exit Sub
        ElseIf InStr(1, strSub(3), "[QC-1] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colQC1Check
            Exit Sub
        ElseIf InStr(1, strSub(3), "[QC-2] probe check failed.") > 0 Then
            SetText vasList, "FAIL", gRow, colQC2Check
            Exit Sub
        End If
        
        
        
        
        strErrorComment = Trim(strDel(3))
        
        
        If Len(Trim(strErrorComment)) > 255 Then
            SetText vasList, Mid(strErrorComment, 1, 250), gRow, colError1
            SetText vasList, Mid(strErrorComment, 251), gRow, colError2
        Else
            SetText vasList, strErrorComment, gRow, colError1
        End If
        
        
      
    Case "L"
        MTBRemarkShcek
        
        '프로브 체크에
        For i = colProACheck To colQC2Check
            If GetText(vasList, gRow, i) = "" Then
                SetText vasList, "PASS", gRow, i
            End If
        Next i
        
        
        '결과를 다 받고 저장함
        Save_Local_One gRow, "1"
        
        
        SQL = ""
        SQL = SQL & vbCrLf & "UPDATE PAT_RES "
        SQL = SQL & vbCrLf & "   SET MTBC = '" & Trim(GetText(vasList, gRow, colMTBPC)) & "'"
        SQL = SQL & vbCrLf & " WHERE BARCODe  = '" & Trim(GetText(vasList, gRow, colBarcode)) & "'"
        res = SendQuery(gLocal, SQL)
        
        Call ExamCount
        If subSend1.Checked = True Then
            strBarcode = Trim(GetText(vasList, gRow, colBarcode))
            strExamCode = Trim(GetText(vasList, gRow, colExamCode))
            strResult = Trim(GetText(vasList, gRow, colResult))
            srtComment = Trim(GetText(vasList, gRow, colRemark))
            strModule = Trim(GetText(vasList, gRow, colRack))
            
            If strResult <> "" And Len(strBarcode) > 10 Then
            
                
                
                sParam = ""
                sParam = sParam & "<Table>" & _
                        "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                        "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                        "<USERID><![CDATA[LIA]]></USERID>" & _
                        "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                        "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                        "<P0><![CDATA[" & strBarcode & "]]></P0>" & _
                        "<P1><![CDATA[" & strExamCode & "]]></P1>" & _
                        "<P2><![CDATA[" & Replace(strResult, "/", vbCrLf) & "]]></P2>" & _
                        "<P3><![CDATA[]]></P3>" & _
                        "<P4><![CDATA[" & gEquip & Mid(strModule, 2, 1) & "]]></P4>" & _
                        "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                        "<P6><![CDATA[]]></P6>" & _
                        "<P7><![CDATA[" & Replace(srtComment, "/", vbCrLf) & "]]></P7>" & _
                        "<P8><![CDATA[]]></P8>" & _
                        "<P9><![CDATA[]]></P9>" & _
                        "</Table>"
            
                sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
                
                
                
                strSend = Online_Result_Qry(sParam)
        
                

                'If InStr(1, strSend, "<P0><![CDATA[" & strBarcode & "]]></P0>") > 0 Then
                'If InStr(1, strSend, strBarcode) > 0 Then
                    SetBackColor vasList, gRow, gRow, 1, colState, 202, 255, 112
                    SetText vasList, "Trans", gRow, colState
    
                    SQL = " Update pat_res Set " & vbCrLf & _
                          " sendflag = '2' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasList, gRow, colBarcode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
    
                        Exit Sub
                    End If
                    
                'Else
'                    SetForeColor vasList, gRow, gRow, 1, colState, 255, 0, 0
'                    SetText vasList, "Failed", gRow, colState
                'End If
            End If
        Else
            SetText vasList, "Result", gRow, colState
        End If
    
    End Select
    
End Sub

Private Sub MTBRemarkShcek()
    Dim strRemark As String
    
    Dim MTBProb As String
    Dim RifProb As String
    
    Dim MTBMin As Integer
    Dim ArrMTB(0 To 4)
    Dim RifMin As Integer
    Dim ArrRif(0 To 4)
    
    Dim i As Integer
    
    If gRow < 1 Then Exit Sub

    ArrMTB(0) = GetText(vasList, gRow, colMTBPA)
    ArrMTB(1) = GetText(vasList, gRow, colMTBPB)
    ArrMTB(2) = GetText(vasList, gRow, colMTBPC)
    ArrMTB(3) = GetText(vasList, gRow, colMTBPD)
    ArrMTB(4) = GetText(vasList, gRow, colMTBPE)
    
    ArrRif(0) = GetText(vasList, gRow, colRifPA)
    ArrRif(1) = GetText(vasList, gRow, colRifPB)
    ArrRif(2) = GetText(vasList, gRow, colRifPC)
    ArrRif(3) = GetText(vasList, gRow, colRifPD)
    ArrRif(4) = GetText(vasList, gRow, colRifPE)
    MTBMin = 9999
    RifMin = 9999
    
    strRemark = ""
    MTBProb = ""
    RifProb = ""
    
    For i = 0 To 4
        If MTBMin > ArrMTB(i) Then
            MTBProb = i
            MTBMin = ArrMTB(i)
        End If
        
        If RifMin > ArrRif(i) Then
            RifProb = i
            RifMin = ArrRif(i)
        End If
    Next i
    
    If InStr(1, UCase(GetText(vasList, gRow, colResult)), "MTB NOT DETECTED") > 0 Then
        
        strRemark = ""
    Else
'        If MTBProb = "0" Then
'            strRemark = "MTB Mutation Probe A"
'        ElseIf MTBProb = "1" Then
'            strRemark = "MTB Mutation Probe B"
'        ElseIf MTBProb = "2" Then
'            strRemark = "MTB Mutation Probe C"
'        ElseIf MTBProb = "3" Then
'            strRemark = "MTB Mutation Probe D"
'        End If
       
        If RifProb = "0" Then
            strRemark = strRemark & "Rif Mutation Probe A"
        ElseIf RifProb = "1" Then
            strRemark = strRemark & "Rif Mutation Probe B"
        ElseIf RifProb = "2" Then
            strRemark = strRemark & "Rif Mutation Probe C"
        ElseIf RifProb = "3" Then
            strRemark = strRemark & "Rif Mutation Probe D"
        ElseIf RifProb = "4" Then
            strRemark = strRemark & "Rif Mutation Probe E"
        End If
        
    End If
    
    If InStr(1, GetText(vasList, gRow, colResult), "ERROR") = 0 Then
        SetText vasList, strRemark, gRow, colRemark
    Else
        SetText vasList, "", gRow, colRemark
    End If
    
    
    
    
End Sub

