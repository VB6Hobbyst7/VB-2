VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19380
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15390
   ScaleWidth      =   28680
   WindowState     =   2  '최대화
   Begin VB.Frame fraHidden 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hidden"
      Height          =   2565
      Left            =   9960
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   5925
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H00ACFFEF&
         Caption         =   "저장포함"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4290
         TabIndex        =   71
         Top             =   270
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Frame fraVision 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   2145
         Begin VB.TextBox txtRCnt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   870
            TabIndex        =   68
            Text            =   "1"
            Top             =   30
            Width           =   495
         End
         Begin VB.CommandButton cmdGetRslt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "받기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1410
            Style           =   1  '그래픽
            TabIndex        =   67
            Top             =   30
            Width           =   615
         End
         Begin VB.TextBox txtLastSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   660
            TabIndex        =   66
            Text            =   "0"
            Top             =   30
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '투명
            Caption         =   "결과수"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   69
            Top             =   60
            Width           =   765
         End
      End
      Begin VB.Frame fraAccess 
         BackColor       =   &H00ACFFEF&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   3450
         TabIndex        =   60
         Top             =   1110
         Visible         =   0   'False
         Width           =   2175
         Begin VB.TextBox txtRackNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   285
            Left            =   0
            TabIndex        =   64
            Text            =   "1"
            Top             =   0
            Width           =   360
         End
         Begin VB.TextBox txtPosNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   285
            Left            =   360
            TabIndex        =   63
            Text            =   "1"
            Top             =   0
            Width           =   360
         End
         Begin VB.TextBox txtSeqNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   285
            Left            =   720
            TabIndex        =   62
            Text            =   "1"
            Top             =   0
            Width           =   360
         End
         Begin VB.CommandButton cmdOrder 
            BackColor       =   &H00C0E0FF&
            Caption         =   "오더전송"
            Height          =   285
            Left            =   1110
            Style           =   1  '그래픽
            TabIndex        =   61
            Top             =   30
            Width           =   1035
         End
      End
      Begin VB.ListBox lstComStatus 
         Height          =   420
         Left            =   660
         TabIndex        =   55
         Top             =   1980
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "워크조회"
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "결과조회"
         Height          =   315
         Left            =   2610
         TabIndex        =   12
         Top             =   270
         Width           =   1425
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   330
         Top             =   1440
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Label lblPatInfo 
         BackStyle       =   0  '투명
         Caption         =   "박검사"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   1170
         TabIndex        =   15
         Top             =   990
         Width           =   3465
      End
      Begin VB.Shape shpPatInfo 
         BorderColor     =   &H00FF0000&
         Height          =   1155
         Left            =   900
         Shape           =   4  '둥근 사각형
         Top             =   810
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   900
         Top             =   210
         Width           =   1545
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   2550
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '아래 맞춤
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   28620
      TabIndex        =   24
      Top             =   14130
      Visible         =   0   'False
      Width           =   28680
      Begin VB.CommandButton cmdRcvView 
         Caption         =   "V"
         Height          =   525
         Left            =   13410
         TabIndex        =   43
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "C"
         Height          =   525
         Left            =   12900
         TabIndex        =   34
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   33
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   32
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   31
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   30
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   29
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         Height          =   555
         Left            =   13950
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   60
         Width           =   3045
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   525
         Left            =   17010
         TabIndex        =   27
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "Rcv"
         Height          =   525
         Left            =   11910
         TabIndex        =   25
         Top             =   60
         Width           =   975
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6495
      Left            =   16440
      TabIndex        =   48
      Top             =   1380
      Width           =   6495
      _Version        =   393216
      _ExtentX        =   11456
      _ExtentY        =   11456
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      MaxCols         =   13
      MaxRows         =   20
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":0E42
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   60
      TabIndex        =   35
      Top             =   570
      Width           =   5895
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "전체"
         Height          =   180
         Index           =   3
         Left            =   5040
         TabIndex        =   75
         Top             =   720
         Width           =   705
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "접수"
         Height          =   180
         Index           =   2
         Left            =   4290
         TabIndex        =   74
         Top             =   720
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "입력"
         Height          =   180
         Index           =   1
         Left            =   3510
         TabIndex        =   73
         Top             =   720
         Width           =   705
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "최종"
         Height          =   180
         Index           =   0
         Left            =   2700
         TabIndex        =   72
         Top             =   720
         Width           =   705
      End
      Begin VB.CommandButton cmdMatch 
         BackColor       =   &H00C0E0FF&
         Caption         =   "M"
         Height          =   375
         Left            =   5340
         Style           =   1  '그래픽
         TabIndex        =   56
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00A7FAEB&
         Caption         =   "조회"
         Height          =   375
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   54
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00A7FAEB&
         Caption         =   "▶▶"
         Height          =   375
         Left            =   4470
         Style           =   1  '그래픽
         TabIndex        =   53
         Top             =   180
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   600
         TabIndex        =   36
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430273
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2190
         TabIndex        =   37
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430273
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2010
         TabIndex        =   39
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   38
         Top             =   210
         Width           =   480
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00404040&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   28680
      TabIndex        =   4
      Top             =   14805
      Width           =   28680
      Begin VB.Timer tmrConn 
         Left            =   18060
         Top             =   90
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   18960
         Top             =   90
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   18540
         Top             =   90
      End
      Begin VB.Timer tmrDBConn 
         Left            =   17610
         Top             =   90
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   19440
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A9A
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2034
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25CE
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B68
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33FA
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3554
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":36AE
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3808
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":40E2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgOff 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":49BC
         Top             =   30
         Width           =   480
      End
      Begin VB.Image imgOn 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":5286
         Top             =   30
         Width           =   480
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   7380
         Top             =   90
         Width           =   45
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   4770
         Top             =   90
         Width           =   45
      End
      Begin VB.Label lblIFStatus 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   11100
         TabIndex        =   45
         Top             =   150
         Width           =   5325
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   10920
         Top             =   90
         Width           =   5685
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3810
         Picture         =   "frmMain.frx":5B50
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   5805
         Picture         =   "frmMain.frx":60DA
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   7080
         Picture         =   "frmMain.frx":6664
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "통신연결"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   3930
         TabIndex        =   9
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "받는신호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   4995
         TabIndex        =   8
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보내는신호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   6120
         TabIndex        =   7
         Top             =   210
         Width           =   900
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6BEE
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6D38
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6E82
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Com1 연결성공"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   7530
         TabIndex        =   6
         Top             =   180
         Width           =   3255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2955
      End
      Begin VB.Label lblDBStatus 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "데이터베이스 연결성공"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   750
         TabIndex        =   5
         Top             =   180
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3180
         Top             =   90
         Width           =   7725
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00A5704B&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   28680
      TabIndex        =   0
      Top             =   0
      Width           =   28680
      Begin VB.ComboBox cboDoct 
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
         Left            =   14220
         Style           =   2  '드롭다운 목록
         TabIndex        =   77
         Top             =   90
         Width           =   1575
      End
      Begin VB.CheckBox chkAdd 
         Appearance      =   0  '평면
         BackColor       =   &H00ACFFEF&
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   20670
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   21480
         TabIndex        =   57
         Text            =   "123456789012345"
         Top             =   90
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTestNmSave 
         BackColor       =   &H00FFC0FF&
         Caption         =   "변경"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12690
         Style           =   1  '그래픽
         TabIndex        =   52
         Top             =   90
         Width           =   615
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "▷ 결과창 숨기기"
         Height          =   375
         Left            =   18420
         Style           =   1  '그래픽
         TabIndex        =   51
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0E0FF&
         Caption         =   "선택저장"
         Height          =   375
         Left            =   15900
         Style           =   1  '그래픽
         TabIndex        =   50
         ToolTipText     =   "선택한 결과를 EMR서버로 저장합니다"
         Top             =   60
         Width           =   1185
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFC0&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   17130
         Style           =   1  '그래픽
         TabIndex        =   49
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   60
         Width           =   1245
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11490
         TabIndex        =   44
         Text            =   "1234567890"
         Top             =   90
         Width           =   1185
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   1980
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9330
         TabIndex        =   10
         Text            =   "1234567890"
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "담당의 : "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   13530
         TabIndex        =   76
         Top             =   120
         Width           =   735
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  '투명
         Height          =   375
         Left            =   13440
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label lblRow 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   21180
         TabIndex        =   59
         Top             =   135
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  '투명
         Height          =   375
         Left            =   20610
         Top             =   60
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblHospInfo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "전남대학교병원 HITACHI 7020"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Top             =   120
         Width           =   4545
      End
      Begin VB.Label lblHospInfo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "전남대학교병원 HITACHI 7020"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   41
         Top             =   90
         Width           =   4515
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "검사자명 : "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   10620
         TabIndex        =   40
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblTestDate 
         BackStyle       =   0  '투명
         Caption         =   "1971-03-11"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7050
         TabIndex        =   3
         Top             =   105
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "검사자ID : "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   8460
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  '투명
         Height          =   375
         Left            =   8370
         Top             =   60
         Width           =   5025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "검사일자 :"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   6120
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '투명
         Height          =   375
         Left            =   6030
         Top             =   60
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '투명
         Height          =   375
         Left            =   90
         Top             =   60
         Width           =   5865
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   16440
      TabIndex        =   13
      Top             =   570
      Width           =   6525
      Begin VB.TextBox txtSA 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "1004"
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "1004"
         Top             =   450
         Width           =   2115
      End
      Begin VB.TextBox txtPatID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "1004"
         Top             =   150
         Width           =   1995
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "1004"
         Top             =   150
         Width           =   2115
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00FF8080&
         Height          =   270
         Left            =   3330
         Top             =   450
         Width           =   1005
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FF8080&
         Height          =   270
         Left            =   3330
         Top             =   150
         Width           =   1005
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00FF8080&
         Height          =   270
         Left            =   240
         Top             =   450
         Width           =   1005
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FF8080&
         Height          =   270
         Left            =   240
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "Sex/Age"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   23
         Top             =   510
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "이      름"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   510
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "병록번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   19
         Top             =   210
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   210
         Width           =   885
      End
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   6030
      TabIndex        =   46
      Top             =   630
      Width           =   16335
      _Version        =   393216
      _ExtentX        =   28813
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16645631
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":6FCC
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7305
      Left            =   60
      TabIndex        =   47
      Top             =   1260
      Visible         =   0   'False
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   12885
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":8EC7
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  조회 "
      Begin VB.Menu mnuResult 
         Caption         =   "▣ 결과 조회"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "▣ 워크 조회"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " 설정 "
      Begin VB.Menu mnuComm 
         Caption         =   "▣ 통신 설정"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "▣ 검사 설정"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "▣ 화면 설정"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServer 
         Caption         =   "▣ 서버 설정"
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "▣ 옵션 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "▣ 기관정보 설정"
      End
      Begin VB.Menu menuSep001 
         Caption         =   "-"
      End
      Begin VB.Menu menuUser 
         Caption         =   "▣ 사용자 설정"
      End
      Begin VB.Menu menuSep002 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " 옵션 "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "▷ 바코드 사용"
         WindowList      =   -1  'True
         Begin VB.Menu mnuBarcode 
            Caption         =   "바코드사용"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "순번사용"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "체크순"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "▷ 적용 결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "▷ 결과 전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "▷ EMR 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " 기타 "
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "원격지원(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sStartTime       As Date
Public sStartDate       As Date

Dim pDel                As Boolean
Dim strOldBarno         As String
Dim gMnuIdx             As Integer

Private Sub cmdEnd_Click()

    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then

        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If

        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        Unload Me

        End
    End If
    
End Sub

Private Sub chkAdd_Click()
    
    If chkAdd.Value = "1" Then
        lblRow.Visible = True
    Else
        lblRow.Visible = False
    End If
    
End Sub

Private Sub cmdAck_Click()
    
    txtSend.Text = txtSend.Text & ACK

End Sub

Private Sub cmdAll_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer

    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To spdOrder.MaxRows
                    spdOrder.Row = intORow
                    spdOrder.Col = colBARCODE
                    If strBarno = GetText(spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    spdOrder.MaxRows = spdOrder.MaxRows + 1
                    intRow = spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                    Next
            
                    '장비에서 오더요청이 안오는 배치오더용
                    Select Case gKUKDO.MACHNM
                        Case "ACCESS2"
                            Call SetTag(spdOrder, GetTag(spdWork, intWRow, colSTATE), intORow, colSTATE)
                            'Call SetText(spdOrder, GetText(spdWork, intWRow, colSTATE), intORow, colSTATE)
                            'Call SetToolTip(spdOrder, GetTag(spdWork, intWRow, colSTATE), intORow, colSTATE)
                        Case "PPC300N"
                            Call SetTag(spdOrder, GetTag(spdWork, intWRow, colSTATE), intORow, colSTATE)
                            'Call SetText(spdOrder, GetText(spdWork, intWRow, colSTATE), intORow, colSTATE)
                            'Call SetToolTip(spdOrder, GetTag(spdWork, intWRow, colSTATE), intORow, colSTATE)
                    End Select

                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To spdOrder.MaxCols
                            spdOrder.Row = 0
                            spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(spdOrder.Text) Then
                                .Row = spdOrder.MaxRows
                                Call SetText(spdOrder, "◇", spdOrder.MaxRows, intOCol)
                            End If
                        Next
                    Next

                    spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
        .MaxRows = 0
    End With
    
    
'    With spdWork
'        For intWRow = 1 To .MaxRows
'            Call spdWork_DblClick(colBARCODE, intWRow)
'            DoEvents
'        Next
'    End With
End Sub

Private Sub cmdClear_Click()

    Call frmClear
    
End Sub

Private Sub cmdEnq_Click()
    
    txtSend.Text = txtSend.Text & ENQ
    
End Sub

Private Sub cmdEot_Click()
    
    txtSend.Text = txtSend.Text & EOT

End Sub

Private Sub cmdEtx_Click()
    
    txtSend.Text = txtSend.Text & ETX

End Sub



Private Sub cmdHide_Click()
    
    spdResult.Visible = False
    fraPatInfo.Visible = False
    
    spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH + 500

End Sub

Private Sub cmdMatch_Click()
    Dim intWRow     As Integer
    Dim intWSrcRow  As Integer
    Dim intORow     As Integer
    Dim intOSrcRow  As Integer
    Dim blnSame     As Boolean
    Dim i           As Integer
    Dim intCnt      As Integer
    Dim varItems    As Variant
    Dim intItems    As Integer
    Dim intOCol     As Integer
    
    blnSame = False
    intCnt = 0
    
    For intWRow = 1 To spdWork.MaxRows
        If GetText(spdWork, intWRow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
            intWSrcRow = intWRow
        End If
    Next
    
    If intCnt = 0 Then
        Exit Sub
    End If
    
    
    If intCnt > 1 Then
        MsgBox "워크리스트에서 하나의 검체만 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    intCnt = 0
    
    For intORow = 1 To spdOrder.MaxRows
        If GetText(spdOrder, intORow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
            intOSrcRow = intORow
            blnSame = True
            'Exit For
        End If
    Next
    
    If blnSame = False Then
        MsgBox "결과리스트에서 대상 검체를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If intCnt > 1 Then
        MsgBox "결과리스트에서 하나의 검체만 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If blnSame = True Then
        For i = colCHECKBOX To colSTATE
            Call SetText(spdOrder, GetText(spdWork, intWSrcRow, i), intOSrcRow, i)
        Next
        
        varItems = GetText(spdWork, intWSrcRow, colITEMS)
        varItems = Split(varItems, "/")
        For intItems = 0 To UBound(varItems)
            For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                spdOrder.Row = 0
                spdOrder.Col = intOCol
                If varItems(intItems) = Trim(spdOrder.Text) Then
                    Call SetText(spdOrder, "◇", intOSrcRow, intOCol)
                End If
            Next
        Next
        
        '정보수정
        SQL = ""
        SQL = SQL & "UPDATE PATRESULT "
        SQL = SQL & "   SET HOSPDATE = '" & Trim(GetText(spdOrder, intOSrcRow, colBARCODE)) & "'   " & vbCrLf
        SQL = SQL & "     , BARCODE  = '" & Trim(GetText(spdOrder, intOSrcRow, colBARCODE)) & "'   " & vbCrLf
        SQL = SQL & "     , PID      = '" & Trim(GetText(spdOrder, intOSrcRow, colPID)) & "'       " & vbCrLf
        SQL = SQL & "     , CHARTNO  = '" & Trim(GetText(spdOrder, intOSrcRow, colCHARTNO)) & "'   " & vbCrLf
        SQL = SQL & "     , SPECIMEN = '" & Trim(GetText(spdOrder, intOSrcRow, colSPECIMEN)) & "'  " & vbCrLf
        SQL = SQL & "     , DEPT     = '" & Trim(GetText(spdOrder, intOSrcRow, colDEPT)) & "'      " & vbCrLf
        SQL = SQL & "     , INOUT    = '" & Trim(GetText(spdOrder, intOSrcRow, colINOUT)) & "'     " & vbCrLf
        SQL = SQL & "     , ERYN     = '" & Trim(GetText(spdOrder, intOSrcRow, colER)) & "'        " & vbCrLf
        SQL = SQL & "     , RETESTYN = '" & Trim(GetText(spdOrder, intOSrcRow, colRT)) & "'        " & vbCrLf
        SQL = SQL & "     , PNAME    = '" & Trim(GetText(spdOrder, intOSrcRow, colPNAME)) & "'     " & vbCrLf
        SQL = SQL & "     , PSEX     = '" & Trim(GetText(spdOrder, intOSrcRow, colPSEX)) & "'      " & vbCrLf
        SQL = SQL & "     , PAGE     = '" & Trim(GetText(spdOrder, intOSrcRow, colPAGE)) & "'      " & vbCrLf
        SQL = SQL & "     , DISKNO   = '" & Trim(GetText(spdOrder, intOSrcRow, colRACKNO)) & "'    " & vbCrLf
        SQL = SQL & "     , POSNO    = '" & Trim(GetText(spdOrder, intOSrcRow, colPOSNO)) & "'     " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO  = '" & gKUKDO.MACHCD & "'                                   " & vbCrLf
        SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, intOSrcRow, colEXAMDATE)) & "'  " & vbCrLf
        SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, intOSrcRow, colEXAMTIME)) & "'  " & vbCrLf
        SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intOSrcRow, colSAVESEQ)) & vbCrLf
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- 성공
        End If
    End If
End Sub

Private Sub cmdOrder_Click()
    Dim i As Integer
    
    strState = ""
    
    With spdOrder
        If .MaxRows > 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = colCHECKBOX
                If .Value = "1" And GetText(spdOrder, i, colSTATE) = "" Then
                    If MsgBox("준비된 오더를 전송하시겠습니까?", vbInformation + vbYesNo) = vbYes Then
                        Call SendData(ENQ)
                        strState = "Q"
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Next
        End If
    End With
    
End Sub



Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

'Private Sub cmdSave_Click()
'    Dim lRow As Long
'    Dim Res  As Integer
'
'    If spdOrder.MaxRows = 0 Then
'        Exit Sub
'    End If
'
'    If MsgBox("선택한 결과를 전송하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
'        With spdOrder
'            For lRow = 1 To .DataRowCnt
'                .Row = lRow
'                .Col = colCHECKBOX
'                If .Value = 1 Then
'                    Res = SaveTransData(lRow, spdOrder)
'
'                    If Res = -1 Then
'                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
'                        SetText spdOrder, "저장실패", lRow, colSTATE
'
'                              SQL = " UPDATE PATRESULT SET " & vbCrLf
'                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
'                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
'                        SQL = SQL & " WHERE EQUIPNO = '" & gKUKDO.MACHCD & "' " & vbCrLf
'                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
'
'                        If DBExec(AdoCn_Local, SQL) Then
'                            '-- 성공
'                        End If
'
'                    Else
'                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
'                        SetText spdOrder, "저장완료", lRow, colSTATE
'
'                              SQL = " UPDATE PATRESULT SET " & vbCrLf
'                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
'                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
'                        SQL = SQL & " WHERE EQUIPNO = '" & gKUKDO.MACHCD & "' " & vbCrLf
'                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
'
'                        If DBExec(AdoCn_Local, SQL) Then
'                            '-- 성공
'                        End If
'
'                    End If
'                    spdOrder.Row = lRow
'                    spdOrder.Col = colCHECKBOX
'                    spdOrder.Value = 0
'                End If
'            Next lRow
'        End With
'    End If
'
'End Sub

Private Sub cmdSearch_Click()
    Dim i       As Integer
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intSeq      As Integer
        
    'Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)

    If gDBTYPE = "99" Then
        spdOrder.MaxRows = 10
    End If
    
'    If gKUKDO.MACHNM = "ACCESS2" Then
'        intRackNo = txtRackNo.Text
'        intPosNo = txtPosNo.Text
'        intSeq = txtSeqNo.Text
'
'        With spdWork
'            For i = 1 To .MaxRows
'                Call SetText(spdWork, Format(intRackNo, "0"), i, colRACKNO)
'                Call SetText(spdWork, ((intPosNo Mod 11) + 1) - 1, i, colPOSNO)
'                Call SetText(spdWork, intSeq, i, colSEQNO)
'                intSeq = intSeq + 1
'                intPosNo = intPosNo + 1
'                If (intPosNo Mod 11) = 0 Then
'                    intRackNo = intRackNo + 1
'                    intPosNo = 1
'                End If
'
'                txtRackNo.Text = intRackNo
'                txtPosNo.Text = intPosNo
'                txtSeqNo.Text = intSeq
'            Next
'        End With
'    End If
    
End Sub

Private Sub cmdSend_Click()
    
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub

Private Sub cmdTestNmSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\KDBAR.ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\KDBAR.ini")

End Sub

Private Sub cmdView_Click()
    
    If cmdView.Caption = "결과창 보이기 ◁" Then
        cmdView.Caption = "▷ 결과창 숨기기"
'        cmdView.ForeColor = &HC000&
        
        spdResult.Visible = True
        fraPatInfo.Visible = True
            
        spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
        spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
        spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
        
        fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
    Else
        cmdView.Caption = "결과창 보이기 ◁"
'        cmdView.ForeColor = &HFF00FF
        
        spdResult.Visible = False
        fraPatInfo.Visible = False
        
        spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH + 500
    End If
    
End Sub


Private Sub comEQP_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
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

            pBuffer = comEqp.Input

            'SetRawData "[Rx]" & pBuffer
            SetRawData "" & pBuffer
            

'            Select Case UCase(gKUKDO.MACHNM)
'                        Case "PATHFAST":        Call Phase_Serial_PATHFAST
'
'                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
'                        Case "UROMETER720":     Call Phase_Serial_UROMETER720
'                        Case "XP300":           Call Phase_Serial_XP300
'                        Case "AU480":           Call Phase_Serial_AU480
'                        Case "UROMETER720":     Call Phase_Serial_UROMETER720
'                        Case "HORIBA":          Call Phase_Serial_HORIBA
'                        Case "ACCESS2":         Call Phase_Serial_ACCESS2
'                        Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
'                        Case "YUMIZEN":         Call Phase_Serial_YUMIZEN           '영인과학 HORIBA YUMIZEN H500
'                        Case "ISMART30":        Call Phase_Serial_ISMART30
'                        Case "STAGO":           Call Phase_Serial_STAGO
'                        'Case "KLITE":           Call Phase_Serial_KLITE
'                        'Case "INDIKO":          Call Phase_Serial_INDIKO
'
'            End Select

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

    If ERMsg$ <> "" Then
        lblIFStatus.Caption = ERMsg$
    End If
    
End Sub

Private Sub Command1_Click()

    spdOrder.MaxRows = spdOrder.MaxRows + 1
    spdOrder.RowHeight(-1) = 14
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call cmdEnd_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
        
        Close #1

        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    End If
    
End Sub



Private Sub SendData(ByVal pSendData As Variant)

    '-- 전송
    comEqp.Output = pSendData
    
    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    DoEvents
    
    '-- 로그기록
    Call SetRawData("[Tx]" & pSendData)

    '-- 상태표시
    ''Call SetCommStatus("S", pSendData, spdComStatus)
    'Call SetCommStatus("S", pSendData, lstComStatus)

End Sub




Private Sub frmClear()
    
    shpPatInfo.Visible = False
    lblPatInfo.Caption = ""
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    spdWork.MaxRows = 0
    
    txtBarcode.Text = ""
    txtPatID.Text = ""
    txtPName.Text = ""
    txtSA.Text = ""
    txtBarNum.Text = ""

    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    txtRackNo.Text = "1"
    txtPosNo.Text = "1"
    txtSeqNo.Text = "1"
        
End Sub


Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    Dim strIFStatus As String
    
On Error GoTo ErrHandle
    
'    Me.Width = 20940
'    Me.Height = 12585


    'Me.Caption = gKUKDO.MACHNM
    'Me.Caption = gKUKDO.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"

    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
        Me.TOP = gForm.TOP
        Me.LEFT = gForm.LEFT
        Me.WIDTH = gForm.WIDTH
        Me.HEIGHT = gForm.HEIGHT
    End If
    
    Me.Caption = gKUKDO.HOSPNM & Space$(5) & gKUKDO.MACHNM
    If gKUKDO.APIURL = "" Then
        Me.Caption = gKUKDO.HOSPNM & Space$(5) & gKUKDO.MACHNM
    Else
        Me.Caption = gKUKDO.HOSPNM & Space$(5) & gKUKDO.MACHNM & " ▶URL : " & gKUKDO.APIURL
    End If
    
    lblHospInfo(0).Caption = "◈ " & gKUKDO.PARTNM & " - " & gKUKDO.MACHNM & " 인터페이스"
    lblHospInfo(1).Caption = "◈ " & gKUKDO.PARTNM & " - " & gKUKDO.MACHNM & " 인터페이스"
    
    
    strIFStatus = ""
    If gKUKDO.BARUSE = "Y" Then
        strIFStatus = strIFStatus & "▣ 바코드사용"
    Else
        If gKUKDO.RSTTYPE = "1" Then
            strIFStatus = strIFStatus & "▣ 순번 맞춤"
        ElseIf gKUKDO.RSTTYPE = "2" Then
            strIFStatus = strIFStatus & "▣ R/P 맞춤"
        ElseIf gKUKDO.RSTTYPE = "3" Then
            strIFStatus = strIFStatus & "▣ 체크순"
        End If
    End If
    strIFStatus = strIFStatus & IIf(gKUKDO.SAVELIS = "Y", "  ▣ LIS결과", "  ▣ 장비결과")
    strIFStatus = strIFStatus & IIf(gKUKDO.SAVEAUTO = "Y", "  ▣ 자동전송", "  ▣ 수동전송")
    
    lblIFStatus.Caption = strIFStatus
    lblComStatus.Caption = ""
    
    If gWORKPOS = "M" Then
        spdWork.Visible = True
        fraWorkInfo.Visible = True
        
        cmdView.Visible = True
        
    Else
        spdWork.Visible = False
        fraWorkInfo.Visible = False
    
        cmdView.Visible = False
    
    End If


    Call CtlInitializing

    Call frmClear
    
    '-- Menu Set
    Call SetMenu

    '-- 컬럼보이기설정
    Call SetColumnView(spdOrder)

    '-- 통신열기
    'Call OpenCommunication

    lstComStatus.Clear
    
    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    txtTestID.Text = gKUKDO.USERID
    txtTestNm.Text = gKUKDO.USERNM
    lblPatInfo.Caption = ""

    imgNet1.ZOrder 0
    tmrDBConn.Interval = 500
    tmrDBConn.Enabled = True
    
    
    '-- 이전결과 삭제
    strTmp = Format$(DateAdd("d", -Val(gKUKDO.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gKUKDO.SAVEDAY + "일전 데이타를 삭제하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gKUKDO.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    
    If gKUKDO.MACHNM = "VISION" Then
        fraVision.Visible = True
    Else
        fraVision.Visible = False
    End If
    
'    If gKUKDO.MACHNM = "ACCESS2" Then
'        fraAccess.Visible = True
'    Else
'        fraAccess.Visible = False
'    End If

    If gKUKDO.DBCONCHK = "Y" Then
        tmrConn.Interval = 60000
        tmrConn.Enabled = True
    Else
        tmrConn.Enabled = False
    End If
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            
            Resume Next
        Else
            End
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & gKUKDO.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If

End Sub

Public Sub GetDoctList()
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    
    Dim strDoctor   As String
    
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DOCTOR", "DOCTOR", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    strDoctor = Trim(strSetUp1)
    
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DOCTOR", "DOCTORCOUNT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    J = Trim(strSetUp1)
    
    If IsNumeric(J) Then
        For i = 1 To J
            strSetup = "":    strSetUp1 = ""
            Call GetPrivateProfileString("DOCTOR", "DOCTOR" & CStr(i), "", strSetup, 100, App.PATH & "\KDBAR.ini")
            strSetUp1 = Trim(strSetup)
            strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
            'gMACHS(i) = Trim(strSetUp1)
            If strDoctor = strSetUp1 Then
                k = i
            End If
            cboDoct.AddItem mGetP(Trim(strSetUp1), 1, "|") & Space(20) & "|" & mGetP(Trim(strSetUp1), 2, "|")
        Next
    End If
    
    If k > 0 Then
        cboDoct.ListIndex = k - 1
    End If
    
End Sub

'
'Public Sub OpenCommunication()
'
'    If gComm.COMTYPE = "1" Then
'
'        comEqp.CommPort = gComm.COMPORT
'        comEqp.RTSEnable = gComm.RTSEnable
'        comEqp.DTREnable = gComm.DTREnable
'        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
'
'        If comEqp.PortOpen = False Then
'            comEqp.PortOpen = True
'        End If
'
'        If comEqp.PortOpen Then
'            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
'
'            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
'            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgOn.ZOrder 0
'           ' imgCom.Picture = imlStatus.ListImages("ON").ExtractIcon
'
'
'        Else
'            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'            imgOff.ZOrder 0
'
'           ' imgCom.Picture = imlStatus.ListImages("OFF").ExtractIcon
'
'        End If
'
'    ElseIf gComm.COMTYPE = "2" Then
'        'lblComStatus.Left = imgPort.Left + 500
'        'lblComStatus.Width = 6000
'        If gComm.TCPTYPE = "SERVER" Then
'            wSck.LocalPort = CInt(gComm.TCPPORT)
'            wSck.Listen
'
'
'            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 연결중.."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            'imgSend.Visible = False
'            'imgReceive.Visible = False
'            'lblSend.Visible = False
'            'lblRcv.Visible = False
'            imgOff.ZOrder 0
'
'        Else
'            wSck.Close
'            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
'
'            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 연결중..."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            'imgSend.Visible = False
'            'imgReceive.Visible = False
'            'lblSend.Visible = False
'            'lblRcv.Visible = False
'            imgOff.ZOrder 0
'
'        End If
'    ElseIf gComm.COMTYPE = "" Then
'
'    End If
'
'End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'Me.TOP = 0

    If gWORKPOS = "M" Then
        spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraWorkInfo.HEIGHT - 300
        
        If spdResult.Visible = True Then
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = fraWorkInfo.TOP + 80
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
            
            fraPatInfo.Visible = True
            fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            fraPatInfo.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - spdResult.HEIGHT - 300
        Else
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = fraWorkInfo.TOP + 80
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
            'spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            'spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
        
            'fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            'fraPatInfo.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - spdResult.HEIGHT - 300
        End If
    Else
        spdOrder.LEFT = 100
        spdOrder.TOP = fraPatInfo.TOP + 80
        spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200
        spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
        
        spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
        
        fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        fraPatInfo.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - spdResult.HEIGHT - 300

    End If
    
    If Me.WindowState = 2 Then
        'gForm.MAXYN = True
        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\KDBAR.ini")
    Else
        If Me.TOP < 0 Then
            Me.TOP = 0
        End If
        'gForm.MAXYN = False
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call WritePrivateProfileString("FORM", "MAXYN", "N", App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("FORM", "TOP", gForm.TOP, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("FORM", "LEFT", gForm.LEFT, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("FORM", "WIDTH", gForm.WIDTH, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("FORM", "HEIGHT", gForm.HEIGHT, App.PATH & "\KDBAR.ini")

    End If
    
End Sub

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Private Function GetPatTRestResult(ByVal asRow As Integer) As Integer
    Dim strBarno    As String
    Dim intSeq      As String
    Dim strExamDate As String
    Dim intRow   As Integer

On Error GoTo ErrHandle

    GetPatTRestResult = -1
    intRow = 0

    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = GetText(spdOrder, asRow, colEXAMDATE)
    strBarno = GetText(spdOrder, asRow, colBARCODE)
    
    If intSeq = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EQUIPCODE, EXAMNAME, EXAMCODE, EQUIPRESULT, RESULT, PREVRESULT, REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE = '" & strBarno & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                If AdoRs_Local.Fields("EXAMCODE").Value & "" = "" Then
                    Call SetText(frmMain.spdResult, "0", intRow, colCHECKBOX)
                Else
                    Call SetText(frmMain.spdResult, "1", intRow, colCHECKBOX)
                End If
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                If AdoRs_Local.Fields("REFJUDGE").Value & "" = "H" Then
                    .ForeColor = vbRed
                    .FontBold = True
                ElseIf AdoRs_Local.Fields("REFJUDGE").Value & "" = "L" Then
                    .ForeColor = vbBlue
                    .FontBold = True
                Else
                    .ForeColor = vbBlack
                    .FontBold = False
                End If
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("PREVRESULT").Value & "", intRow, colRPREVRESULT)
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
        GetPatTRestResult = 1
    End If

    AdoRs_Local.Close

Exit Function

ErrHandle:
    GetPatTRestResult = -1

    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gKUKDO.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Function

Private Sub imgPort_DblClick()
    
    If gComm.COMTYPE = "1" And comEqp.PortOpen = True Then
        
        If MsgBox("COMM PORT CLOSE?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
            comEqp.PortOpen = False
        End If
    ElseIf gComm.COMTYPE = "1" And comEqp.PortOpen = False Then
        
        If MsgBox("COMM PORT OPEN?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
            comEqp.CommPort = gComm.COMPORT
            comEqp.RTSEnable = gComm.RTSEnable
            comEqp.DTREnable = gComm.DTREnable
            comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
    
            If comEqp.PortOpen = False Then
                comEqp.PortOpen = True
            End If
        End If
    
    End If
    
    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
        
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    End If

End Sub



Private Sub menuUser_Click()
    
    frmMstUser.Show vbModal, frmMain

End Sub

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\KDBAR.ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\KDBAR.ini")

End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\KDBAR.ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\KDBAR.ini")

End Sub

Private Sub mnuComm_Click()
    
    frmConfig.Show

End Sub

Private Sub mnuComTest_Click()

End Sub

Private Sub mnuCommTest_Click()

    If picComm.Visible = True Then
        picComm.Visible = False
    Else
        picComm.Visible = True
    End If
    
End Sub

'Private Sub mnuDoctor_Click()
'
'    frmDoctor.Show
'
'End Sub

'Private Sub mnuEMRInfo_Click()
'
'    If InputBox("비밀번호 입력" & Space(5) & "hint:개발자oyh") = "dev0503" Then
'        frmEMRInfo.Show
'    End If
'
'End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\KDBAR.ini")

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then

        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If

        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\KDBAR.ini")

End Sub

'Private Sub mnuOpt_Click()
'
'    frmTestOptSet.Show vbModal
'
'End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\KDBAR.ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\KDBAR.ini")

End Sub


Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\KDBAR.ini")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\KDBAR.ini")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\KDBAR.ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\KDBAR.ini")
    
End Sub

Public Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    '-- 정렬
'    If Row = 0 Then
'        '-- 정렬 추가
'        Exit Sub
'    End If
    
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "1", i, colCHECKBOX)
            Next
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, Row, colCHECKBOX) = "1" Then
            Call SetText(spdOrder, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdOrder, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
    
    If Row = 0 Then
        Exit Sub
    End If
    
    '-- 환자정보표시
    shpPatInfo.Visible = True
    
    strPatInfo = ""
    strPatInfo = strPatInfo & "◈이    름: " & GetText(spdOrder, Row, colPNAME)
    If GetText(spdOrder, Row, colPSEX) <> "" Then
        strPatInfo = strPatInfo & " [" & GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE) & "] " & vbCrLf
    Else
        strPatInfo = strPatInfo & vbCrLf
    End If
    strPatInfo = strPatInfo & "◈검체번호: " & GetText(spdOrder, Row, colBARCODE) & vbCrLf
    strPatInfo = strPatInfo & "◈환자번호: " & GetText(spdOrder, Row, colPID) & vbCrLf
    
    lblPatInfo.Caption = strPatInfo
    
    txtBarcode.Text = GetText(spdOrder, Row, colBARCODE)
    txtPatID.Text = GetText(spdOrder, Row, colPID)
    txtPName.Text = GetText(spdOrder, Row, colPNAME)
    txtSA.Text = GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE)
    
    '-- 결과표시
    If GetPatTRestResult(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◇
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 15
                End If
            Next
        End With
    End If

    lblRow.Caption = Row
    
End Sub


Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    If Row = 0 Then
        If Col = colCHECKBOX Then
            If GetText(spdWork, 1, colCHECKBOX) = "1" Then
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "0", i, colCHECKBOX)
                Next
            Else
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "1", i, colCHECKBOX)
                Next
            End If
        Else
            '-- 정렬 추가
            Call SetSpreadSort(spdWork)
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    Dim strBarno_Work   As String
    'Dim strUritItems    As String
    
    If Row = 0 Then Exit Sub
    If Col <> colBARCODE Then
        Exit Sub
    End If
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            .MaxRows = .MaxRows + 1
            intRow = .MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
            Next
            
            '장비에서 오더요청이 안오는 배치오더용
            Select Case gKUKDO.MACHNM
                Case "ACCESS2"
                    Call SetTag(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                    'Call SetToolTip(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                Case "PPC300N"
                    Call SetTag(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                    'Call SetToolTip(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
            End Select
            
            varItems = GetText(spdWork, intWRow, colITEMS)
            varItems = Split(varItems, "/")
            For intItems = 0 To UBound(varItems)
                For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                    .Row = 0
                    .Col = intOCol
                    If varItems(intItems) = Trim(.Text) Then
                        .Row = intRow
                        Call SetText(spdOrder, "◇", intRow, intOCol)
                    End If
                Next
            Next
            
            
            
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
            .RowHeight(-1) = 15
        End If
    
    End With
End Sub

Private Sub tmrConn_Timer()
    Dim sqlRet          As Long
    Dim RS          As ADODB.Recordset
    
On Error GoTo ErrHandle
    If DbConnect_SQL = True Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("Select sysdate from DUAL", sqlRet)
        RS.Close
        
        ''Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmMain.lstComStatus)
    End If
Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gKUKDO.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    lblDBStatus.Caption = "데이터베이스 연결실패"
'    frmErrMsg.Show
    
End Sub

Private Sub tmrDBConn_Timer()

    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If
    
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub




Private Sub txtPosNo_KeyPress(KeyAscii As Integer)
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intRow      As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intRackNo = txtRackNo.Text
        intPosNo = txtPosNo.Text
        
        If Not IsNumeric(intPosNo) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intRackNo, intRow, colRACKNO)
                Call SetText(spdOrder, ((intPosNo Mod 11) + 1) - 1, intRow, colPOSNO)
                intPosNo = intPosNo + 1
                If (intPosNo Mod 11) = 0 Then
                    intRackNo = intRackNo + 1
                    intPosNo = 1
                End If
            Next
        End With
        
        txtRackNo.Text = intRackNo
        txtPosNo.Text = intPosNo
        
        'Call txtSeqNo_KeyPress(vbKeyReturn)
        
    End If
End Sub

Private Sub txtRackNo_KeyPress(KeyAscii As Integer)
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intRow      As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intRackNo = txtRackNo.Text
        intPosNo = txtPosNo.Text
        
        If Not IsNumeric(intRackNo) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            If .MaxRows = 0 Then
                Exit Sub
            End If
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intRackNo, intRow, colRACKNO)
                Call SetText(spdOrder, ((intPosNo Mod 11) + 1) - 1, intRow, colPOSNO)
                intPosNo = intPosNo + 1
                If (intPosNo Mod 11) = 0 Then
                    intRackNo = intRackNo + 1
                    intPosNo = 1
                End If
            Next
        End With
        
        txtRackNo.Text = intRackNo
        txtPosNo.Text = intPosNo
    
        'Call txtSeqNo_KeyPress(vbKeyReturn)
    
    End If
    
'    intRackNo = txtRackNo.Text
'    intPosNo = txtPosNo.Text
'    intSeq = txtSeqNo.Text
'
'    With spdWork
'        For i = 1 To .MaxRows
'            Call SetText(spdWork, Format(intRackNo, "0"), i, colRACKNO)
'            Call SetText(spdWork, ((intPosNo Mod 11) + 1) - 1, i, colPOSNO)
'            Call SetText(spdWork, intSeq, i, colSEQNO)
'            intSeq = intSeq + 1
'            intPosNo = intPosNo + 1
'            If (intPosNo Mod 11) = 0 Then
'                intRackNo = intRackNo + 1
'                intPosNo = 1
'            End If
'
'            txtRackNo.Text = intRackNo
'            txtPosNo.Text = intPosNo
'            txtSeqNo.Text = intSeq
'        Next
'    End With
    
End Sub

Private Sub txtSeqNo_KeyPress(KeyAscii As Integer)
    Dim intSeq  As Integer
    Dim intRow  As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intSeq = txtSeqNo.Text
        
        If Not IsNumeric(intSeq) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intSeq, intRow, colSEQNO)
                intSeq = intSeq + 1
            Next
        End With
        
        txtSeqNo.Text = intSeq
        
        'Call txtRackNo_KeyPress(vbKeyReturn)
    End If
    
End Sub



