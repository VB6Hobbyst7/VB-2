VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SANSOFT LAB INTERFACE"
   ClientHeight    =   10110
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   15675
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
   Picture         =   "frmInterface.frx":554A
   ScaleHeight     =   10110
   ScaleWidth      =   15675
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   7365
      Left            =   6180
      TabIndex        =   11
      Top             =   2070
      Visible         =   0   'False
      Width           =   8055
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00F8E4D8&
         Caption         =   "세로"
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
         Index           =   1
         Left            =   3930
         TabIndex        =   42
         Top             =   690
         Width           =   705
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00F8E4D8&
         Caption         =   "가로"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   41
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.CheckBox chkWAll 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3480
         TabIndex        =   39
         Top             =   510
         Width           =   225
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1455
         Left            =   150
         TabIndex        =   23
         Top             =   1170
         Width           =   4725
         Begin VB.OptionButton optSaveResult 
            Caption         =   "장비"
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
            Index           =   0
            Left            =   990
            TabIndex        =   26
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton optSaveResult 
            Caption         =   "수정"
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
            Index           =   1
            Left            =   1740
            TabIndex        =   25
            Top             =   300
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtTemp 
            Height          =   435
            Left            =   3840
            TabIndex        =   24
            Top             =   690
            Width           =   645
         End
         Begin VB.Label lblBarcode 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   240
            TabIndex        =   43
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label5 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과적용"
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
            Left            =   90
            TabIndex        =   27
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3435
         Left            =   180
         TabIndex        =   16
         Top             =   2760
         Width           =   3705
         Begin FPSpread.vaSpread vasExcel 
            Height          =   1005
            Left            =   1860
            TabIndex        =   17
            Top             =   2220
            Visible         =   0   'False
            Width           =   1725
            _Version        =   393216
            _ExtentX        =   3043
            _ExtentY        =   1773
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
            SpreadDesigner  =   "frmInterface.frx":57CD
         End
         Begin FPSpread.vaSpread vasCode 
            Height          =   945
            Left            =   120
            TabIndex        =   18
            Top             =   2190
            Width           =   1665
            _Version        =   393216
            _ExtentX        =   2937
            _ExtentY        =   1667
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
            SpreadDesigner  =   "frmInterface.frx":59F3
         End
         Begin FPSpread.vaSpread vasTemp1 
            Height          =   945
            Left            =   1860
            TabIndex        =   19
            Top             =   1230
            Width           =   1725
            _Version        =   393216
            _ExtentX        =   3043
            _ExtentY        =   1667
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
            SpreadDesigner  =   "frmInterface.frx":5C19
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   975
            Left            =   120
            TabIndex        =   20
            Top             =   210
            Width           =   1695
            _Version        =   393216
            _ExtentX        =   2990
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
            SpreadDesigner  =   "frmInterface.frx":5E3F
         End
         Begin FPSpread.vaSpread vasResTemp 
            Height          =   1035
            Left            =   1860
            TabIndex        =   21
            Top             =   180
            Width           =   1695
            _Version        =   393216
            _ExtentX        =   2990
            _ExtentY        =   1826
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
            SpreadDesigner  =   "frmInterface.frx":6065
         End
         Begin FPSpread.vaSpread vasTemp 
            Height          =   975
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1695
            _Version        =   393216
            _ExtentX        =   2990
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
            SpreadDesigner  =   "frmInterface.frx":628B
         End
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   150
         TabIndex        =   15
         Top             =   210
         Width           =   2835
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
            Top             =   300
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   660
            Top             =   300
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1140
            Top             =   180
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":64B1
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6A4B
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6FE5
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":757F
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7E11
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7F6B
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":80C5
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   3900
         TabIndex        =   12
         Top             =   3570
         Width           =   1965
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   13
            Top             =   1290
            Width           =   1710
            _Version        =   393216
            _ExtentX        =   3016
            _ExtentY        =   1826
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
            SpreadDesigner  =   "frmInterface.frx":821F
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1665
            _Version        =   393216
            _ExtentX        =   2937
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
            SpreadDesigner  =   "frmInterface.frx":86FC
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8535
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   15495
      Begin FPSpread.vaSpread vasID 
         Height          =   7995
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   8475
         _Version        =   393216
         _ExtentX        =   14949
         _ExtentY        =   14102
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         ColsFrozen      =   16
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
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   1
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmInterface.frx":8922
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8010
         Left            =   8670
         TabIndex        =   2
         Top             =   180
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   14129
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
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmInterface.frx":94B1
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00F8E4D8&
      Height          =   645
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   15615
      TabIndex        =   3
      Top             =   0
      Width           =   15675
      Begin VB.CommandButton cmdIFClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면정리"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11130
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   90
         Width           =   1065
      End
      Begin VB.CommandButton cmdSL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "상세결과"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9990
         Style           =   1  '그래픽
         TabIndex        =   40
         Top             =   90
         Width           =   1125
      End
      Begin VB.CommandButton cmdWorkPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결과출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8850
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   90
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   16920
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   190447617
         CurrentDate     =   40457
      End
      Begin VB.CommandButton cmdExcelExport 
         BackColor       =   &H00FFFFFF&
         Caption         =   "엑셀출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7740
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.CommandButton cmdRsltSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결과조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6450
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   90
         Width           =   1215
      End
      Begin VB.Frame FrmCommTest 
         BackColor       =   &H00F8E4D8&
         Height          =   555
         Left            =   15900
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtRcv 
            Height          =   405
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   120
            Width           =   3465
         End
         Begin VB.CommandButton cmdCommTest 
            Caption         =   "받기"
            Height          =   375
            Left            =   3540
            TabIndex        =   29
            Top             =   150
            Width           =   615
         End
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   1680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   5040
         TabIndex        =   35
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
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
         Format          =   190447617
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   345
         Left            =   3570
         TabIndex        =   36
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
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
         Format          =   190447617
         CurrentDate     =   40248
      End
      Begin VB.Label lblTestDate 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Left            =   16020
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label12 
         Appearance      =   0  '평면
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
         Height          =   195
         Left            =   4890
         TabIndex        =   38
         Top             =   210
         Width           =   105
      End
      Begin VB.Label Label20 
         BackColor       =   &H00F8E4D8&
         Caption         =   "조회일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   37
         Top             =   210
         Width           =   795
      End
      Begin VB.Label lblMachNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "HemaVet950"
         BeginProperty Font 
            Name            =   "Segoe UI Historic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   780
         TabIndex        =   10
         Top             =   180
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   60
         Picture         =   "frmInterface.frx":9B38
         Top             =   90
         Width           =   2580
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   14670
         TabIndex        =   8
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   13785
         TabIndex        =   7
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   195
         Left            =   12930
         TabIndex        =   6
         Top             =   210
         Width           =   420
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15180
         Picture         =   "frmInterface.frx":B347
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14295
         Picture         =   "frmInterface.frx":B8D1
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   13440
         Picture         =   "frmInterface.frx":BE5B
         Top             =   180
         Width           =   240
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   9705
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2021-03-12"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "오전 10:19"
         EndProperty
      EndProperty
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
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnPrint 
         Caption         =   "인쇄"
         Begin VB.Menu MnPrintLand 
            Caption         =   "가로인쇄"
         End
         Begin VB.Menu MnPrintPort 
            Caption         =   "세로인쇄"
         End
      End
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
      Caption         =   "Send"
      Visible         =   0   'False
      Begin VB.Menu MnTransAuto 
         Caption         =   "Auto"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
      End
   End
   Begin VB.Menu MnMode 
      Caption         =   "Mode"
      Visible         =   0   'False
      Begin VB.Menu MnModeBarcode 
         Caption         =   "Barcode"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnModeWorkList 
         Caption         =   "WorkList"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gRow            As Long
Dim gsFlag          As String
Dim gChecked        As Boolean
Dim strBuffer       As String
Dim strRecvData()   As String

Dim strState        As String
Dim intPhase        As Integer
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer

Const STX   As String = ""
Const ETX   As String = ""
Const ENQ   As String = ""
Const ACK   As String = ""
Const NAK   As String = ""
Const EOT   As String = ""
Const ETB   As String = ""
Const FS    As String = ""
Const RS    As String = ""
Const GS    As String = ""

Private Sub chkWAll_Click()
    Dim iRow As Long
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 0
            Next iRow
        End If
    End With
End Sub

Private Sub SaveExcel(Filename As String, argSpread As vaSpread)
    Dim xlApp   As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim iRow    As Integer
    Dim iCol    As Integer
    
On Error Resume Next
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlApp.Quit

End Sub

Private Sub cmdCommTest_Click()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = txtRcv.Text
    
    lngBufLen = Len(Buffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)
        Select Case BufChar
            Case STX
                strBuffer = ""
            Case ETX
                dtpToday = Date
                
                DoEvents
                
                If gMach = "HemaVet950" Then
                    Call EditRcvData_HemaVet950
                    strBuffer = ""
                    
                ElseIf gMach = "AU10V" Then
                    Call EditRcvData_AU10V
                    strBuffer = ""
                    
                ElseIf gMach = "FDC7000" Or gMach = "FDC7000i" Or gMach = "NX500" Or gMach = "NX500i" Or gMach = "NX700i" Then
                    Call EditRcvData_NXSeries
                    strBuffer = ""
                Else
'                    Call EditRcvDataASTM
                    strBuffer = ""
                End If
            Case Else
                strBuffer = strBuffer & BufChar
                
        End Select
    Next i

End Sub

Private Sub cmdExcelExport_Click()
    Dim iRow        As Integer
    Dim j, k        As Integer
    Dim sFileName   As String
    Dim blnWrite    As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasPrint.MaxRows = vasID.MaxRows
    vasPrint.MaxCols = vasID.MaxCols
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
            
        If vasID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasID.MaxCols
                    If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                        If j = 3 Then k = 1
                        If j = 4 Then k = 2
                        If j = 5 Then k = 3
                        If j = 6 Then k = 4
                        
                        If j >= 17 Then
                            k = j - 12
                        End If

                        SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, k
                    End If
                Next
            End If

            For j = 1 To vasID.MaxCols
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                    If j = 3 Then k = 1
                    If j = 4 Then k = 2
                    If j = 5 Then k = 3
                    If j = 6 Then k = 4
                    If j >= 17 Then
                        k = j - 12
                    End If
                        
                    SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, k
                End If
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        Call SaveExcel(sFileName, vasPrint)
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Private Sub cmdIFClear_Click()
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    gRow = 0
    
End Sub

Private Sub cmdRsltSearch_Click()
    Dim RS          As ADODB.Recordset
    Dim strDate     As String
    Dim strSaveSeq  As String
    Dim strChart    As String
    Dim i           As Integer
    Dim blnSame     As Boolean
    Dim intCol      As Integer
    
    ClearSpread vasID
    ClearSpread vasRes

    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS 접수일자, BARCODE AS 바코드번호, CHARTNO AS 차트번호, PID AS 내원번호, PNAME AS 이름,PSEX AS 성별, PAGE AS 나이, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE,SEQNO "
    Set RS = cn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                    If blnSame = True Then
                        For intCol = colState + 1 To vasID.MaxCols
                            If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                                If Trim(RS.Fields("REFFLAG")) = "H" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbRed
                                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbBlue
                                End If
                                Exit For
                            End If
                        Next
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "0", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Format(Trim(RS("EXAMDATE")), "####-##-##"), .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("차트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("나이")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("INOUT")) & "", .MaxRows, colINOUT
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                        Case "0": SetText vasID, "에러", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 201, 112
                        Case "1": SetText vasID, "결과조회", .MaxRows, colState
                        Case "2": SetText vasID, "완료", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                        Case "3": SetText vasID, "수정", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 245, 112
                    End Select
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                            If Trim(RS.Fields("REFFLAG")) = "H" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbRed
                            ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbBlue
                            End If
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            RS.MoveNext
        Loop
    End If
    RS.Close
    vasID.RowHeight(-1) = 15
    
End Sub

Private Sub cmdSL_Click()
    
    If cmdSL.Caption = "전체결과" Then '▶◀
        cmdSL.Caption = "상세결과"
        vasID.Width = Frame1.Width - 200
    Else
        cmdSL.Caption = "전체결과"
        vasID.Width = Me.Width - vasRes.Width - 710
    End If

    Call Form_Resize
    
End Sub

Private Sub cmdWorkPrint_Click()
    Dim iRow        As Integer
    Dim i, j, k     As Integer
    Dim blnWrite    As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    
    vasPrint.MaxRows = vasID.MaxRows
    vasPrint.MaxCols = vasID.MaxCols
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
            
        If vasID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasID.MaxCols
                    If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                        If j = 3 Then k = 1
                        If j = 4 Then k = 2
                        If j = 5 Then k = 3
                        If j = 6 Then k = 4
                        
                        If j >= 17 Then
                            k = j - 12
                        End If

                        SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, k
                    End If
                Next
            End If

            For j = 1 To vasID.MaxCols
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                    If j = 3 Then k = 1
                    If j = 4 Then k = 2
                    If j = 5 Then k = 3
                    If j = 6 Then k = 4
                    If j >= 17 Then
                        k = j - 12
                    End If
                        
                    SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, k
                End If
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    Else
        If optPrint(0).Value = True Then
            For i = 6 To vasPrint.MaxCols
                vasPrint.ColWidth(i) = 5
            Next
            vasPrint.PrintOrientation = PrintOrientationLandscape '가로출력
            vasPrint.Action = 13
        Else
            vasPrint.PrintOrientation = PrintOrientationPortrait '세로출력
            vasPrint.Action = 13
        End If
        MsgBox "결과 출력완료", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If frmInterface.ScaleHeight = 0 Then Exit Sub
        
    If cmdSL.Caption = "전체결과" Then
        Frame1.Height = frmInterface.ScaleHeight - 1200
        vasID.Height = Frame1.Height - 300
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 7300
        Frame1.Top = Picture1.Top + Picture1.Height
        
        vasRes.Height = vasID.Height
        vasRes.Left = vasID.Width + 200
    Else
        Frame1.Height = frmInterface.ScaleHeight - 1200
        vasID.Height = Frame1.Height - 300
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 300
        Frame1.Top = Picture1.Top + Picture1.Height
    End If
    
    StatusBar1.Panels(3).Width = Frame1.Width - 8500
    
End Sub

Private Sub imgPort_DblClick()
    
    If FrmHideControl.Visible = True Then
        FrmHideControl.Visible = False
    Else
        FrmHideControl.Visible = True
    End If

End Sub

Private Sub Form_Load()
    Dim sDate As String
    
On Error GoTo Rst

    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    '화면 초기화
    Call cmdIFClear_Click
    
    '설정읽어오기(INI)
    Call GetSetup
    '장비명
    lblMachNm.Caption = gEquip
    '사용자
    frmInterface.StatusBar1.Panels(1).Text = gUserID
    '통신설정
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    If comEqp.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        frmInterface.StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If
    '로컬DB연결
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    '검사정보읽어오기
    Call GetExamCode
    '검사정보화면표시
    Call SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    If gPOri = "L" Then
        optPrint(0).Value = True
    Else 'If gPOri = "P" Then
        optPrint(1).Value = True
    End If
    
    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================

Exit Sub
Rst:
    If Err = 8002 Then
        MsgBox "통신포트를 확인하세요!", vbExclamation, "Communication"
        frmConfig.Show 1
        Call GetExamCode
    Else
        Resume Next
    End If
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 8
        Next
    End With
    
End Sub

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = ""
    SQL = SQL & "Select equipcode, examcode, examname, resprec, seqno   " & vbCrLf
    SQL = SQL & "  From EQPMASTER                                       " & vbCrLf
    SQL = SQL & " Where equipno = '" & gEquip & "'                      " & vbCrLf
    SQL = SQL & " Order by seqno * 10                                   " & vbCrLf
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 7)
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
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
    Next i
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

    DisConnect_Local
    Unload Me
    End

End Sub

Private Sub lblMachNm_DblClick()
    
    If dtpToday.Visible = True Then
        dtpToday.Visible = False
        lblTestDate.Visible = False
    Else
        dtpToday.Visible = True
        lblTestDate.Visible = True
    End If
    
End Sub

Private Sub lblPort_DblClick()
    
    If FrmCommTest.Visible = False Then
        FrmCommTest.Visible = True
    Else
        FrmCommTest.Visible = False
    End If

End Sub

Private Sub MnExamConfig_Click()
    frmTestSet.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnPrintLand_Click()

    vasID.PrintOrientation = PrintOrientationLandscape '가로출력
    vasID.Action = 13

End Sub

Private Sub MnPrintPort_Click()

    vasID.PrintOrientation = PrintOrientationPortrait '세로출력
    vasID.Action = 13

End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub comEqp_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
    Select Case comEqp.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            Buffer = comEqp.Input
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        strBuffer = ""
                    Case ETX
                        dtpToday = Date
                        DoEvents
                        If gMach = "HemaVet950" Or gMach = "HEMAVET950" Then
                            Call EditRcvData_HemaVet950
                            strBuffer = ""
                            
                        ElseIf gMach = "AU10V" Then
                            Call EditRcvData_AU10V
                            strBuffer = ""
                            
                        ElseIf gMach = "URISCAN" Then
                            Call EditRcvData_URISCAN
                            strBuffer = ""
                            
                        ElseIf gMach = "FDC7000" Or gMach = "FDC7000i" Or gMach = "NX500" Or gMach = "NX500i" Or gMach = "NX700i" Then
                            Call EditRcvData_NXSeries
                            strBuffer = ""
                        Else
'                            Call EditRcvDataASTM
                            strBuffer = ""
                        End If
                    Case Else
                        strBuffer = strBuffer & BufChar
                End Select
            Next i

        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
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

Private Sub SetPatInfo(ByVal pBarNo As String, Optional ByVal pRno As String, Optional ByVal pPno As String)
    Dim intRow      As Long
    
    intRow = -1
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- 장비수신정보 표시
    Call SetText(vasID, "1", intRow, colCheckBox)
    If pBarNo = "" Then
        Call SetText(vasID, mResult.PatNo, intRow, colBARCODE)
        Call SetText(vasID, mResult.PatNo, intRow, colCHARTNO)
    Else
        Call SetText(vasID, mResult.BarNo, intRow, colBARCODE)
        Call SetText(vasID, mResult.PatNo, intRow, colCHARTNO)
    End If
    Call SetText(vasID, mResult.TestTime, intRow, colHOSPDATE)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    Call SetText(vasID, mResult.TubePos, intRow, colPOSNO)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 현재 Row
    gRow = intRow
    
End Sub



'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_NXSeries()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarNo        As String   '수신한 바코드번호
    Dim strTestDt       As String
    Dim strTestTm       As String

    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment

    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String

    Dim lsExamCode      As String
    Dim lsExamName      As String
    Dim lsSeqNo         As String
    Dim lsResult_Buff   As String
    Dim lsExamDate      As String
    Dim lsEquipRes      As String
    Dim lsResRow        As String

    Dim i               As Integer
    Dim intCol          As Integer
    
    strRcvBuf = strBuffer
                
    strType = Mid$(strRcvBuf, 1, 1)
    If IsNumeric(strType) Then
        strType = Mid$(strRcvBuf, 2, 1)
    End If
    
    '12345678901234567890123456789012345678901234567890
    'NORMAL 2017-05-1523:596                         01GGT-P  =1        U/l   1  @         
    'NORMAL 2019-01-2415:191452474                   03IP-P   <0.1      mg/dl 01
    'NORMAL 2019-12-0415:1419           19 f8 3      04TG-P   =25       mg/dl 01           GPT-P  =29       U/l   01


    '-- Type1 일때 사용(오더요청)
    If UCase(strType) = "W" Then
        strBarNo = Trim(mGetP(strRcvBuf, 2, ","))
        
    ElseIf UCase(strType) = "N" Then
        strTestDt = Trim(Mid(strRcvBuf, 8, 10))
        strTestTm = Trim(Mid(strRcvBuf, 18, 5))
        
        strTestDt = strTestDt & Space(1) & strTestTm
        strSeq = Trim(Mid(strRcvBuf, 23, 13))
        strBarNo = Trim(Mid(strRcvBuf, 36, 13))
        strTubePos = Trim(Mid(strRcvBuf, 49, 2))
    
        If strBarNo = "" Then
            strBarNo = strSeq
        End If
        
        With mResult
            .TestTime = strTestDt
            .PatNo = strSeq
            .BarNo = strBarNo
            If IsNumeric(strTubePos) Then
                .TubePos = Val(strTubePos)
            Else
                .TubePos = strTubePos
            End If
            .RsltDate = Format(Now, "yyyymmdd")
            .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
        End With
        
        Call SetPatInfo(strBarNo)

        If gRow < 0 Then
            Exit Sub
        End If
                    
        For i = 51 To Len(strRcvBuf) Step 36
            strIntBase = Trim(Mid(strRcvBuf, i, 7))
            strResult = Trim(Mid(strRcvBuf, i + 8, 8))
            strComm = Trim(Mid(strRcvBuf, i + 16, 7))
        
            strResult = Replace(strResult, "=", "")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
        
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH"
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                
                If Res > 0 Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsSeqNo = Trim(gReadBuf(2))
                    strLow = Trim(gReadBuf(3))
                    strHigh = Trim(gReadBuf(4))
                    
                    lsResRow = vasRes.DataRowCnt + 1
                    If vasRes.MaxRows < lsResRow Then
                        vasRes.MaxRows = lsResRow
                    End If
                    
                    '소수점 처리, 결과 형태 처리
                    lsEquipRes = strResult
                    
                    If IsNumeric(strResult) Then
                        strResult = SetResult(strResult, strIntBase)
                    End If
                    
                    '단위포함
                    If gUnit = True Then
                        strResult = strResult & " " & strComm
                    End If
                    
                    lsResult_Buff = strResult
                    
                    '-- Work List
                    SetText vasID, "결과수신", gRow, colState                 '11 진행상태
                    
                    '-- vasID 에 표시
                    For intCol = colState + 1 To vasID.MaxCols
                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, strResult, gRow, intCol
                            Exit For
                        End If
                    Next
            
                    strJudge = ""
                    If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                        If IsNumeric(lsEquipRes) Then
                            If CCur(lsEquipRes) > CCur(strLow) And CCur(lsEquipRes) < CCur(strHigh) Then
                                strJudge = ""
                            ElseIf CCur(strHigh) <= CCur(lsEquipRes) Then
                                strJudge = "H"
                            ElseIf CCur(strLow) >= CCur(lsEquipRes) Then
                                strJudge = "L"
                            End If
                        End If
                    End If
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlue
                    Else
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlack
                    End If
    
                    '-- 결과 List
                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlue
                    Else
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlack
                    End If
                    
                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                    SetText vasRes, strJudge, lsResRow, colFLAG             '판정
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
            
            vasRes.RowHeight(-1) = 15
        Next
    
    End If
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_URISCAN()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarNo        As String   '수신한 바코드번호
    Dim strTestDt       As String
    Dim strTestTm       As String

    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment

    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String

    Dim lsExamCode      As String
    Dim lsExamName      As String
    Dim lsSeqNo         As String
    Dim lsResult_Buff   As String
    Dim lsExamDate      As String
    Dim lsEquipRes      As String
    Dim lsResRow        As String

    Dim i               As Integer
    Dim intCol          As Integer
    Dim Pos             As Integer
    
    Pos = InStr(strBuffer, "ID_NO")
    If Pos > 0 Then
        strBuffer = Replace(strBuffer, vbLf, "")
        strRecvData = Split(strBuffer, vbCr)
                
        '-- 검사시간
        strRcvBuf = strRecvData(0)
        strRcvBuf = Mid(strRcvBuf, 7)
        strTestDt = Trim(strRcvBuf)
                
                
        '-- ID 찾기
        strRcvBuf = strRecvData(1)
        strRcvBuf = mGetP(strRcvBuf, 2, ":")
        strRcvBuf = mGetP(strRcvBuf, 1, "-")
        strSeq = Trim(strRcvBuf)
        
        If strBarNo = "" Then
            strBarNo = strSeq
        End If
        
        With mResult
            .TestTime = strTestDt
            .PatNo = strSeq
            .BarNo = strBarNo
            If IsNumeric(strTubePos) Then
                .TubePos = Val(strTubePos)
            Else
                .TubePos = strTubePos
            End If
            .RsltDate = Format(Now, "yyyymmdd")
            .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
        End With
        
        Call SetPatInfo(strBarNo)

        If gRow < 0 Then
            Exit Sub
        End If
                    
        For i = 4 To UBound(strRecvData)
            strRcvBuf = strRecvData(i)
            strIntBase = Trim(Mid(strRcvBuf, 1, 3))
            strComm = Trim(Mid(strRcvBuf, 19))
            strResult = ""
    
            Select Case strIntBase
                Case "p.H", "pH", "S.G", "SG", "COL" '## 소숫점 포함 3자리
                    strResult = Trim$(Mid$(strRcvBuf, 4))
                    strResult = Replace(strResult, "mg/dl", "")
                    strResult = Replace(strResult, "RBC/ul", "")
                    strResult = Replace(strResult, "WBC/ul", "")
                    
                    strResult = Replace(strResult, "<", "")
                    strResult = Replace(strResult, ">", "")
                    strResult = Replace(strResult, "=", "")
                
                Case Else
                    strResult = Trim$(Mid$(strRcvBuf, 4, 7))
                    'strResult = Trim(Mid(strRcvBuf, 12))  '-- 정량
                    strResult = Replace(strResult, "mg/dl", "")
                    strResult = Replace(strResult, "RBC/ul", "")
                    strResult = Replace(strResult, "WBC/ul", "")
                    
                    strResult = Replace(strResult, "<", "")
                    strResult = Replace(strResult, ">", "")
                    strResult = Replace(strResult, "=", "")
            End Select
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH"
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                
                If Res > 0 Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsSeqNo = Trim(gReadBuf(2))
                    strLow = Trim(gReadBuf(3))
                    strHigh = Trim(gReadBuf(4))
                    
                    lsResRow = vasRes.DataRowCnt + 1
                    If vasRes.MaxRows < lsResRow Then
                        vasRes.MaxRows = lsResRow
                    End If
                    
                    '소수점 처리, 결과 형태 처리
                    lsEquipRes = strResult
                    
                    If IsNumeric(strResult) Then
                        strResult = SetResult(strResult, strIntBase)
                    End If
                    
                    '단위포함
                    If gUnit = True Then
                        strResult = strResult & " " & strComm
                    End If
                    
                    lsResult_Buff = strResult
                    
                    '-- Work List
                    SetText vasID, "결과수신", gRow, colState                 '11 진행상태
                    
                    '-- vasID 에 표시
                    For intCol = colState + 1 To vasID.MaxCols
                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, strResult, gRow, intCol
                            Exit For
                        End If
                    Next
            
                    strJudge = ""
                    If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                        If IsNumeric(lsEquipRes) Then
                            If CCur(lsEquipRes) > CCur(strLow) And CCur(lsEquipRes) < CCur(strHigh) Then
                                strJudge = ""
                            ElseIf CCur(strHigh) <= CCur(lsEquipRes) Then
                                strJudge = "H"
                            ElseIf CCur(strLow) >= CCur(lsEquipRes) Then
                                strJudge = "L"
                            End If
                        End If
                    End If
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlue
                    Else
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlack
                    End If
    
                    '-- 결과 List
                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlue
                    Else
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlack
                    End If
                    
                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                    SetText vasRes, strJudge, lsResRow, colFLAG             '판정
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
            
            vasRes.RowHeight(-1) = 15
        Next
        
    End If
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_AU10V()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
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
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    Dim strTestNo   As String
    
    varRcvBuf = Split(strBuffer, vbCrLf)
    
    'For i = 0 To UBound(varRcvBuf)
    strRcvBuf = strBuffer
    strType = Mid(strRcvBuf, 1, 1)
    
    Select Case strType
        Case "N"    '## Normal result
    
            strSeqNo = Trim$(Mid(strRcvBuf, 23, 13))
            strBarNo = Trim$(Mid(strRcvBuf, 36, 13))
            strTestDt = Trim$(Mid(strRcvBuf, 8, 10)) & " " & Trim$(Mid(strRcvBuf, 18, 5))
            strTestDt = Format(strTestDt, "yyyy-mm-dd")
            If strBarNo = "" Then
                strBarNo = strSeqNo
            End If
            'strDevice = Trim$(Mid(strRcvBuf, 49, 2))
            
            '-- 오른쪽 결과화면 초기화
            vasRes.MaxRows = 0
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '결과일자
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If

            strIntBase = Trim(Mid(strRcvBuf, 51, 7))
            strResult = Trim(Mid(strRcvBuf, 58, 10))
            strResult = Replace(strResult, "=", "")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = strResult & " " & Trim(Mid(strRcvBuf, 68, 5))
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
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
                    SetText vasID, "Result", gRow, colState                 '11 진행상태
                    
                    '-- vasID 에 표시
                    For intCol = colState + 1 To vasID.MaxCols
                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- 결과 List
                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                    SetText vasRes, strComm, lsResRow, 7                    'Flag
                    '-- 로컬 저장
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
        
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
        
        Case Else
        
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_HemaVet950()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarNo        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim strTestDt       As String
    Dim strTestTm       As String
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String
    
    Dim lsExamCode      As String
    Dim lsExamName      As String
    Dim lsSeqNo         As String
    Dim lsResult_Buff   As String
    Dim lsExamDate      As String
    Dim lsEquipRes      As String
    Dim lsResRow        As String
    
    Dim i               As Integer
    Dim intCol          As Integer
    Dim varRcvBuf       As Variant
    
    varRcvBuf = Split(strBuffer, vbCrLf)
    
    For i = 0 To UBound(varRcvBuf)
        strRcvBuf = varRcvBuf(i)
        
        If i = 0 Then
            '1,  0,HV   , 1,PATIENT 00,OTHER      , 2584,10/31/15,09:25:29,B
            strSeq = Trim$(mGetP(strRcvBuf, 2, ","))        'Seq Number
            strBarNo = Trim$(mGetP(strRcvBuf, 5, ","))      'Sample id
            strTubePos = Trim(mGetP(strRcvBuf, 7, ","))     'Test Number
            strTestDt = Trim$(mGetP(strRcvBuf, 8, ","))     'Date
            strTestTm = Trim$(mGetP(strRcvBuf, 9, ","))     'Time
            strTestDt = Format(strTestDt, "yyyy-mm-dd") & Space(1) & Format(strTestTm, "hh:mm:ss")
            
        
            If strBarNo = "" Then
                strBarNo = strTestDt
            End If
            
            With mResult
                .TestTime = strTestDt
                .PatNo = strSeq
                .BarNo = strBarNo
                If IsNumeric(strTubePos) Then
                    .TubePos = Val(strTubePos)
                Else
                    .TubePos = strTubePos
                End If
                .RsltDate = Format(Now, "yyyymmdd")
                .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
            End With
            
            Call SetPatInfo(strBarNo)
    
            If gRow < 0 Then
                Exit Sub
            End If
        Else
            strIntBase = Trim$(mGetP(varRcvBuf(i), 1, ","))
            strResult = Trim$(mGetP(varRcvBuf(i), 2, ","))
            strComm = Trim$(mGetP(varRcvBuf(i), 4, ","))
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH"
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                
                If Res > 0 Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsSeqNo = Trim(gReadBuf(2))
                    strLow = Trim(gReadBuf(3))
                    strHigh = Trim(gReadBuf(4))
                    
                    lsResRow = vasRes.DataRowCnt + 1
                    If vasRes.MaxRows < lsResRow Then
                        vasRes.MaxRows = lsResRow
                    End If
                    
                    '소수점 처리, 결과 형태 처리
                    lsEquipRes = strResult
                    
                    If IsNumeric(strResult) Then
                        strResult = SetResult(strResult, strIntBase)
                    End If
                    
                    '단위포함
                    If gUnit = True Then
                        strResult = strResult & " " & strComm
                    End If
                    
                    lsResult_Buff = strResult
                    
                    '-- Work List
                    SetText vasID, "결과수신", gRow, colState                 '11 진행상태
                    
                    '-- vasID 에 표시
                    For intCol = colState + 1 To vasID.MaxCols
                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, strResult, gRow, intCol
                            Exit For
                        End If
                    Next
            
                    strJudge = ""
                    If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                        If IsNumeric(lsEquipRes) Then
                            If CCur(lsEquipRes) > CCur(strLow) And CCur(lsEquipRes) < CCur(strHigh) Then
                                strJudge = ""
                            ElseIf CCur(strHigh) <= CCur(lsEquipRes) Then
                                strJudge = "H"
                            ElseIf CCur(strLow) >= CCur(lsEquipRes) Then
                                strJudge = "L"
                            End If
                        End If
                    End If
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlue
                    Else
                        vasID.Row = gRow
                        vasID.Col = intCol
                        vasID.ForeColor = vbBlack
                    End If
    
                    '-- 결과 List
                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                    SetText vasRes, strResult, lsResRow, colRESULT          '결과
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlue
                    Else
                        vasRes.Row = lsResRow
                        vasRes.Col = colRESULT
                        vasRes.ForeColor = vbBlack
                    End If
                    
                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                    SetText vasRes, strJudge, lsResRow, colFLAG             '판정
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
            
        End If
    Next
    
    vasRes.RowHeight(-1) = 15

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_FDC7000()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
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
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    
    strRcvBuf = strBuffer
    
    strType = mGetP(strRcvBuf, 1, ",")
            
    Select Case strType
        Case "R"
            '-- 오른쪽 결과화면 초기화
            vasRes.MaxRows = 0
            
            strTestDt = Trim(mGetP(strRcvBuf, 3, ","))
            strTestTm = Trim(mGetP(strRcvBuf, 4, ","))
            strSeqNo = Trim(mGetP(strRcvBuf, 5, ","))
            strBarNo = Trim(mGetP(strRcvBuf, 6, ","))
            
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '결과일자
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            
            For i = 13 To Len(strRcvBuf) Step 7
                strIntBase = Trim(mGetP(strRcvBuf, i, ","))
                strIntBase = mGetP(strIntBase, 1, "-")
                '-- 판정 , 단위 포함
                'strResult = Trim(mGetP(strRcvBuf, i + 1, ",")) & Trim(mGetP(strRcvBuf, i + 2, ","))
                strResult = Trim(mGetP(strRcvBuf, i + 1, ",")) & Mid(Trim(mGetP(strRcvBuf, i + 2, ",")), 1, 8)
                
                strResult = Replace(strResult, "=", "")
                strResult = Replace(strResult, "  ", " ")
                strResult = Replace(strResult, "  ", " ")
                strResult = Replace(strResult, "  ", " ")
            
                strResult = Trim(strResult)
                
                If strIntBase = "TCHO" Then
                    strTC = strResult
                    MsgBox "strTC:" & strTC
                End If
                
                If strIntBase = "TG" Then
                    strTG = strResult
                    MsgBox "strTG:" & strTG
                End If
                
                If strIntBase = "HDLC" Then
                    strHDL = strResult
                    MsgBox "strHDL:" & strHDL
                End If
                
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        
                        '-- vasID 에 표시
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            'LDL 계산식 적용
            If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                MsgBox "1"
                'blnLDLCal = False
                strIntBase = "LDL-CAL"
                strResult = strTC - ((strTG / 5) + strHDL)
                If strResult < 0 Then
                    strResult = "0"
                End If
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        
                        '-- vasID 에 표시
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
                
            End If
        
            
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_FDC7000i()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
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
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    
    strRcvBuf = strBuffer
    
    strType = Mid(strRcvBuf, 1, 1)
            
    Select Case strType
        Case "N"
            '-- 오른쪽 결과화면 초기화
            vasRes.MaxRows = 0
            
            'NORMAL 2002-03-0521:299            1512150228   03GGT-P  =67       U/l   01           CPK-P  =121      U/l   01           CRE-P  =1.0      mg/dl 01           
            'NORMAL 2002-03-0521:3110           1512150228   03BUN-P  =10.0     mg/dl 01           TBIL-P =0.5      mg/dl 01           LDH-P  =193      U/l   01           ALB-P  =5.0      g/dl  01           HDLC-P =42       mg/dl 01           TP-P   =7.6      g/dl  01           GLU-P  =76       mg/dl 01           GOT-P  =35       U/l   01           TG-P   =410      mg/dl 01H          GPT-P  =48       U/l   01H          TCHO-P =236      mg/dl 01H          ALP-P  =250      U/l   01           

            
            strTestDt = Trim(Mid(strRcvBuf, 8, 10))
            strTestTm = Trim(Mid(strRcvBuf, 18, 5))
            strSeqNo = Trim(Mid(strRcvBuf, 23, 5))
            strBarNo = Trim(Mid(strRcvBuf, 30, 16))
            
            If strBarNo = "" Then
                strBarNo = strSeqNo
            End If
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '결과일자
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            

            
            For i = 51 To Len(strRcvBuf) Step 36
                strIntBase = Trim(Mid(strRcvBuf, i, 7))
                strIntBase = mGetP(strIntBase, 1, "-")
                strResult = Trim(Mid(strRcvBuf, i + 8, 8))
            
                If strIntBase = "TCHO" Then
                    strTC = strResult
                End If
                
                If strIntBase = "TG" Then
                    strTG = strResult
                End If
                
                If strIntBase = "HDLC" Then
                    strHDL = strResult
                End If
            
            
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        
                        '-- vasID 에 표시
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            'LDL 계산식 적용
            If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                'blnLDLCal = False
                strIntBase = "LDL-CAL"
                strResult = strTC - ((strTG / 5) + strHDL)
                If strResult < 0 Then
                    strResult = "0"
                End If
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        
                        '-- vasID 에 표시
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
                
            End If
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTMi()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
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
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    strRcvBuf = strBuffer
    
    strType = Mid(strRcvBuf, 1, 1)
            
    Select Case strType
        Case "N"
            '-- 오른쪽 결과화면 초기화
            vasRes.MaxRows = 0
            
            'NORMAL 2002-03-0521:299            1512150228   03GGT-P  =67       U/l   01           CPK-P  =121      U/l   01           CRE-P  =1.0      mg/dl 01           
            'NORMAL 2002-03-0521:3110           1512150228   03BUN-P  =10.0     mg/dl 01           TBIL-P =0.5      mg/dl 01           LDH-P  =193      U/l   01           ALB-P  =5.0      g/dl  01           HDLC-P =42       mg/dl 01           TP-P   =7.6      g/dl  01           GLU-P  =76       mg/dl 01           GOT-P  =35       U/l   01           TG-P   =410      mg/dl 01H          GPT-P  =48       U/l   01H          TCHO-P =236      mg/dl 01H          ALP-P  =250      U/l   01           

            
            strTestDt = Trim(Mid(strRcvBuf, 8, 10))
            strTestTm = Trim(Mid(strRcvBuf, 18, 5))
            strSeqNo = Trim(Mid(strRcvBuf, 23, 5))
            strBarNo = Trim(Mid(strRcvBuf, 30, 16))
            
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '결과일자
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            
            For i = 51 To Len(strRcvBuf) Step 36
                strIntBase = Trim(Mid(strRcvBuf, i, 7))
                strIntBase = mGetP(strIntBase, 1, "-")
                strResult = Trim(Mid(strRcvBuf, i + 8, 8))
            
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        
                        '-- vasID 에 표시
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colRESULT          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

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
    
    SQL = "select resprec, reflow, refhigh from EQPMASTER where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
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
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
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
    Dim strSaveSeq As String
    
    'sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    sExamDate = Trim(GetText(vasID, asRow1, colEXAMDATE))
    If Trim(GetText(vasID, asRow1, colSAVESEQ)) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
   ' SetRawData "[SQL]" & SQL
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
    SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
    SQL = SQL & ", EXAMDATE"                        '검사일자"
    SQL = SQL & ", HOSPDATE"                        '병원접수일자"
    SQL = SQL & ", EQUIPNO"                         '장비코드"
    SQL = SQL & ", BARCODE" & vbCrLf
    SQL = SQL & ", EQUIPCODE"                       '검사채널"
    SQL = SQL & ", EXAMCODE"                        '병원검사코드"
    SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
    SQL = SQL & ", EXAMNAME"
    SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
    SQL = SQL & ", SAMPLETYPE"                      '검체유형"
    SQL = SQL & ", INOUT"                           '검체코드
    SQL = SQL & ", DISKNO"
    SQL = SQL & ", POSNO"
    SQL = SQL & ", EQUIPRESULT"                     '장비결과"
    SQL = SQL & ", RESULT" & vbCrLf                 '소수점적용결과"
    SQL = SQL & ", REFFLAG"
    SQL = SQL & ", REFVALUE"
    SQL = SQL & ", CHARTNO"
    SQL = SQL & ", PID"                             '병록번호(내원번호)"
    SQL = SQL & ", PNAME" & vbCrLf
    SQL = SQL & ", PSEX"
    SQL = SQL & ", PAGE"
    SQL = SQL & ", PJUMIN"
    SQL = SQL & ", PANICVALUE"
    SQL = SQL & ", DELTAVALUE" & vbCrLf
    SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
    SQL = SQL & ", SENDDATE"
    SQL = SQL & ", EXAMUID"
    SQL = SQL & ", HOSPITAL)" & vbCrLf
    SQL = SQL & " VALUES (" & vbCrLf
'    SQL = SQL & strSaveSeq
    SQL = SQL & Trim(GetText(vasID, asRow1, colSAVESEQ))
    SQL = SQL & ",'" & sExamDate
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colHOSPDATE))
    SQL = SQL & "','" & gEquip
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colBARCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEQUIPCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSUBCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
    SQL = SQL & "','"
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colINOUT))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))   '-- ORDERCODE 저장
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
'    SetRawData "[SQL]" & SQL
    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- 결과업데이트
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
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

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    If BlockRow <= 0 Then
        Exit Sub
    End If
    
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
    Dim RS          As ADODB.Recordset
    'Dim lsID        As String
    Dim strBarcode  As String
    Dim lngSaveseq  As Long
    Dim iRow        As Long
    
    If Row = 0 Then
        If Col = colCheckBox Then
            With vasID
                If gChecked = False Then
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = colCheckBox
                        .Value = 1
                    Next iRow
                    gChecked = True
                Else
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = colCheckBox
                        .Value = 0
                    Next iRow
                    gChecked = False
                End If
            End With
        Else
            With vasID
                .Col = 1: .Col2 = .MaxCols
                .Row = 2: .Row2 = .DataRowCnt
                .SortBy = 0
                .SortKey(1) = Col       '정렬키 열번호
    
                .SortKeyOrder(1) = SortKeyOrderAscending
        
                .Action = ActionSort
            End With
            Exit Sub
        End If
    End If
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    strBarcode = Trim(GetText(vasID, Row, colBARCODE))
    lblBarcode.Caption = strBarcode
    lngSaveseq = Trim(GetText(vasID, Row, colSAVESEQ))
    If Trim(GetText(vasID, Row, colSAVESEQ)) = "" Then
        Exit Sub
    End If
    
    frmInterface.StatusBar1.Panels(3).Text = strBarcode
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lngSaveseq & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        vasRes.MaxRows = 0
        Do Until RS.EOF
            With vasRes
                .MaxRows = .MaxRows + 1
                SetText vasRes, "0", .MaxRows, colCheckBox
                SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
                SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
                SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
                SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
                
                If Trim(RS.Fields("REFFLAG")) = "H" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbRed
                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbBlue
                End If
           
            End With
            RS.MoveNext
        Loop
    End If
    vasRes.RowHeight(-1) = 12
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow    As Long
    Dim iCol    As Long
    Dim strBarcode  As String
    Dim strPID      As String
    Dim lsTime  As String
'    Dim lsPid   As String
    'Dim lsSeq   As String
    Dim lngSaveseq  As Long
    Dim i       As Integer
    Dim strResult As String
    Dim blnModify As Boolean
    
    blnModify = False
    
    iRow = vasID.ActiveRow
    iCol = vasID.ActiveCol

    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        If iCol > colState Then
            Exit Sub
        End If
        strBarcode = Trim(GetText(vasID, iRow, colBARCODE))
        strPID = Trim(GetText(vasID, iRow, colPID))
        lngSaveseq = Trim(GetText(vasID, iRow, colSAVESEQ))

        If lngSaveseq = "" Then
            Exit Sub
        End If

        If MsgBox(lngSaveseq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCrLf
        SQL = SQL & "   AND PID     = '" & strPID & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lngSaveseq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        vasID.MaxRows = vasID.MaxRows - 1
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            
            strBarcode = Trim(GetText(vasID, iRow, colBARCODE))
            
            '-- 바코드 번호가 이전과 틀리다면 업데이트
            If strBarcode <> lblBarcode.Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "yyyymmdd") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & strBarcode & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",INOUT = '" & Trim(GetText(vasID, iRow, colINOUT)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode.Caption & "' "

                'SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True

            End If
        ElseIf iCol = colDISKNO Then
            
        
        Else
            Exit Sub
            vasID.Row = iRow
            vasID.Col = colState
            If Trim(vasID.Text) = "" Then
                Exit Sub
            End If

            '-- 결과만 수정했을 경우의 업데이트는 Delete >> Insert 순으로 한다.
            '-- Delete
                  SQL = "DELETE FROM PATRESULT "
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, iRow, colBARCODE)) & "' "

            Res = SendQuery(gLocal, SQL)
                
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

            '-- Insert
            For i = colState + 1 To vasID.MaxCols
                vasID.Row = iRow
                vasID.Col = i
                If Trim(vasID.Text) <> "" Then
                    '-- 결과 소수점 적용
                    strResult = SetResult(Trim(GetText(vasID, iRow, i)), gArrEquip(i - colState, 2))
                    '-- H/L 일때 색표시
                    If gsFlag = "L" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbBlue
                    ElseIf gsFlag = "H" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbRed
                    End If
                    vasID.Text = strResult

                    SQL = ""
                    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
                    SQL = SQL & "SAVESEQ, EXAMDATE, HOSPDATE, EQUIPNO, BARCODE" & vbCrLf
                    SQL = SQL & ", EQUIPCODE, EXAMCODE, EXAMSUBCODE, EXAMNAME, SEQNO" & vbCrLf
                    SQL = SQL & ", SAMPLETYPE, DISKNO, POSNO, EQUIPRESULT, RESULT" & vbCrLf
                    SQL = SQL & ", REFFLAG, REFVALUE, CHARTNO, PID, PNAME" & vbCrLf
                    SQL = SQL & ", PSEX, PAGE, PJUMIN, PANICVALUE, DELTAVALUE" & vbCrLf
                    SQL = SQL & ", SENDFLAG, SENDDATE, EXAMUID, HOSPITAL)" & vbCrLf
                    SQL = SQL & " VALUES (" & vbCrLf
                    SQL = SQL & Trim(GetText(vasID, iRow, colSAVESEQ))
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colEXAMDATE))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colHOSPDATE))
                    SQL = SQL & "','" & gEquip
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colBARCODE))
                    'equipcode , examcode, examname, resprec, seqno
                    SQL = SQL & "','" & gArrEquip(i - colState, 2) 'Trim(GetText(vasRes, asRow2, colEQUIPCODE))
                    SQL = SQL & "','" & gArrEquip(i - colState, 3) 'Trim(GetText(vasRes, asRow2, colEXAMCODE))
                    SQL = SQL & "','"                              'Trim(GetText(vasRes, asRow2, colSubCode))
                    SQL = SQL & "','" & gArrEquip(i - colState, 4) 'Trim(GetText(vasRes, asRow2, colEXAMNAME))
                    SQL = SQL & "','" & gArrEquip(i - colState, 6) 'Trim(GetText(vasRes, asRow2, colSeq))
                    SQL = SQL & "',''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colDISKNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPOSNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colMachResult))
                    SQL = SQL & "','" & strResult 'Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colRESULT))
                    SQL = SQL & "','" & gsFlag & "'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colCHARTNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPID))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPNAME))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPSEX))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPAGE))
                    SQL = SQL & "',''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'3'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gIFUser
                    SQL = SQL & "','')"

                    Res = SendQuery(gLocal, SQL)
                    SetText vasID, "수정", iRow, colState

                End If
            Next
            blnModify = True
        End If
    End If
    
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub

        vasID_Click colBARCODE, lRow
    End If
End Sub



