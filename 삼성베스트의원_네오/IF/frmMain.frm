VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15960
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
   ScaleHeight     =   9180
   ScaleWidth      =   15960
   WindowState     =   2  '최대화
   Begin VB.Frame fraHidden 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hidden"
      Height          =   3975
      Left            =   9510
      TabIndex        =   11
      Top             =   4050
      Visible         =   0   'False
      Width           =   7005
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4200
         Top             =   2730
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraAccess 
         BackColor       =   &H00ACFFEF&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   690
         TabIndex        =   60
         Top             =   2640
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdOrder 
            BackColor       =   &H00C0E0FF&
            Caption         =   "오더전송"
            Height          =   285
            Left            =   1110
            Style           =   1  '그래픽
            TabIndex        =   64
            Top             =   30
            Width           =   1035
         End
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
            TabIndex        =   63
            Text            =   "1"
            Top             =   30
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
            TabIndex        =   62
            Text            =   "1"
            Top             =   30
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
            TabIndex        =   61
            Text            =   "1"
            Top             =   30
            Width           =   360
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
      ScaleWidth      =   15900
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   15960
      Begin VB.Frame fraVision 
         BackColor       =   &H00ACFFEF&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   6390
         TabIndex        =   67
         Top             =   150
         Width           =   2145
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
            TabIndex        =   70
            Text            =   "0"
            Top             =   30
            Visible         =   0   'False
            Width           =   255
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
            Left            =   1440
            Style           =   1  '그래픽
            TabIndex        =   69
            Top             =   30
            Width           =   615
         End
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
            Left            =   900
            TabIndex        =   68
            Text            =   "1"
            Top             =   30
            Width           =   495
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
            TabIndex        =   71
            Top             =   60
            Width           =   765
         End
      End
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
         Left            =   11940
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
         Format          =   137887745
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
         Format          =   137887745
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
      ScaleWidth      =   15960
      TabIndex        =   4
      Top             =   8595
      Width           =   15960
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
               Picture         =   "frmMain.frx":1A23
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FBD
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2557
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2AF1
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3383
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":34DD
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3637
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3791
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":406B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgOff 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":4945
         Top             =   30
         Width           =   480
      End
      Begin VB.Image imgOn 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":520F
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
         Picture         =   "frmMain.frx":5AD9
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   5805
         Picture         =   "frmMain.frx":6063
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   7080
         Picture         =   "frmMain.frx":65ED
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
         Picture         =   "frmMain.frx":6B77
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6CC1
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6E0B
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
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   15960
      TabIndex        =   0
      Top             =   0
      Width           =   15960
      Begin VB.Frame fraMicroPlate 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   4590
         TabIndex        =   65
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
         Begin VB.CommandButton cmdReceive 
            BackColor       =   &H00FFC0FF&
            Caption         =   "결과받기"
            Height          =   345
            Left            =   30
            Style           =   1  '그래픽
            TabIndex        =   66
            Top             =   0
            Width           =   1305
         End
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
         Left            =   6060
         TabIndex        =   58
         Top             =   120
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
         Left            =   6870
         TabIndex        =   57
         Text            =   "123456789012345"
         Top             =   90
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
         Left            =   15150
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
         Left            =   13950
         TabIndex        =   44
         Text            =   "1234567890"
         Top             =   90
         Width           =   1185
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   2940
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   2430
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
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
         Left            =   11790
         TabIndex        =   10
         Text            =   "1234567890"
         Top             =   90
         Width           =   1185
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
         Left            =   6570
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
         Left            =   6030
         Top             =   60
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
         Width           =   5475
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
         Width           =   5475
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
         Left            =   13080
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
         Left            =   9510
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
         Left            =   10920
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
         Left            =   10830
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
         Left            =   8580
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
         Left            =   8490
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
      Top             =   690
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16645631
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":6F55
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7275
      Left            =   60
      TabIndex        =   47
      Top             =   1290
      Visible         =   0   'False
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   12832
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
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":8C22
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
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMRInfo 
         Caption         =   "▣ 전산정보 설정"
         Visible         =   0   'False
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
                    Select Case gHOSP.MACHNM
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


Private Sub cmdGetRslt_Click()
    Dim strSendData As String
    Dim strFirstSeq  As String
    Dim strLastSeq  As String
'    Dim db_tmp As String * 100
    
On Error GoTo RST

    strFirstSeq = txtLastSeq.Text
    strFirstSeq = (strFirstSeq - 1) - (txtRCnt.Text - 1)
    
    strLastSeq = strFirstSeq + (txtRCnt.Text - 1)
    
    strSendData = "0" & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq & vbLf
    
    wSck.SendData strSendData
    SetRawData "[Tx]" & strSendData

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_cmdGetRslt_Click" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

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
        SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'                                   " & vbCrLf
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

Private Sub cmdRcv_Click()
    Dim i As Integer
    
    pBuffer = txtRcv.Text
    
'    If UCase(gHOSP.MACHNM) = "PPC300N" Then
'        If LEFT(pBuffer, 1) <> SB Then
'            pBuffer = SB & pBuffer
'        End If
'        If Right(pBuffer, 1) <> EB Then
'            pBuffer = pBuffer & EB
'        End If
'    ElseIf UCase(gHOSP.MACHNM) = "GENEXPERT" Then
'    Else
'        If LEFT(pBuffer, 1) <> STX Then
'            pBuffer = STX & pBuffer
'        End If
'        If Right(pBuffer, 1) <> ETX Then
'            pBuffer = pBuffer & ETX
'        End If
'    End If
    
    Select Case UCase(gHOSP.MACHNM)
        Case "HORIBA":          Call Phase_Serial_HORIBA
        Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
        
        Case "F200":            Call Phase_TCP_F200
        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
        Case "CA800":           Call Phase_Serial_CA800
        Case "XN1000":          Call Phase_Serial_XN1000
        Case "AU480":           Call Phase_Serial_AU480
        Case "GENEXPERT":       Call Phase_TCP_GENEXPERT
        Case "PPC300N":         Call Phase_TCP_PPC300N
        Case "UROMETER720":     Call Phase_Serial_UROMETER720
        Case "HORIBA":          Call Phase_Serial_HORIBA
        Case "ACCESS2":         Call Phase_Serial_ACCESS2
        Case "YUMIZEN":         Call Phase_Serial_YUMIZEN
        'Case "XP300":           Call Phase_Serial_XP300
        Case "XP300":           Call Phase_TCP_XP300
        Case "ISMART30":        Call Phase_Serial_ISMART30
        Case "STAGO":           Call Phase_Serial_STAGO
        Case "PATHFAST":        Call Phase_Serial_PATHFAST
        Case "VISION":          Call Phase_TCP_VISION
    
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdRcvView_Click()

    frmLogView.Show
    
End Sub

Private Sub cmdReceive_Click()
    Dim strInfoPath     As String
    Dim strRsltPath     As String
    Dim txtFilename     As String
    
    With CommonDialog1
        .CancelError = True
        
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly
        .InitDir = gComm.RSTPATH
        .Filter = "XML Files (info.xml)|*.xml|All Files (*.*)|*.*|"
        .FilterIndex = 1
        .Filename = ""
        .ShowOpen
        txtFilename = .Filename
    End With

    Screen.MousePointer = 11
    
    strInfoPath = txtFilename 'gComm.RSTPATH & "\" & "info.xml"
    
    Call DisplayNode_Info(strInfoPath)

    
    If UBound(strRecvData) > 1 Then
        Call SerialRcvData_MULTIPLATE
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
    
ErrHandler:
  ' 사용자가 [취소] 단추를 눌렀습니다.
Exit Sub
    
End Sub


Public Sub DisplayNode_Info(asPath As String)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    
    Dim i, J, k         As Integer
    Dim MsgType         As String
    
    On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    'xmlDoc.Load "D:\프로젝트\VB\__JC메디컴\새움병원_MCC\IF\XML"
    
    k = 0
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                Exit For
            End If
        Next
        
        Erase strRecvData
        intBufCnt = 1
        ReDim Preserve strRecvData(7)
        strRecvData(intBufCnt) = "H|" & xNode.childNodes.Item(0).baseName
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Runtime" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(8).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AreaUnderCurve" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(9).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Aggregation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(10).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Velocity" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(11).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AdditionalInformation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(12).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1


        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    MsgBox "파일오류"
    
    Exit Sub
    
End Sub
'
'
'Public Sub DisplayNode_Result(asPath As String)
'
'    Dim xmlDoc As New MSXML2.DOMDocument30
'    Dim nodeBook As IXMLDOMElement
'    Dim nodeId As IXMLDOMAttribute
'    Dim xNode As MSXML2.IXMLDOMNode
'    Dim namedNodeMap As IXMLDOMNamedNodeMap
'    Dim Child_Node As MSXML2.IXMLDOMNodeList
''    Dim MsgType As String
''    Dim strBuffer As String
''    Dim intRow As Long
''    Dim varBuffer As Variant
''    Dim blnQc     As Boolean
'    Dim i, J, k, m As Integer
'    Dim ii, jj, kk  As Integer
'    Dim strOData    As String
'    Dim strRData    As String
'
'    On Error GoTo ErrXML:
''    On Error Resume Next
'
'    Set xmlDoc = New MSXML2.DOMDocument30
'
'    xmlDoc.async = False
'    xmlDoc.Load asPath
'    'xmlDoc.Load "D:\프로젝트\VB\광주포유병리과의원\참고\Result.xml"
'
'    If (xmlDoc.parseError.errorCode <> 0) Then
'        Dim myErr
'        Set myErr = xmlDoc.parseError
'        MsgBox ("You have error " & myErr.reason)
'    Else
'        Set Child_Node = xmlDoc.childNodes
'        For Each xNode In Child_Node
'            If xNode.nodeType = NODE_ELEMENT Then
'                'MsgType = xNode.nodeName
'                'If MsgType = "testinfo" Then
'                    Exit For
'                'End If
'            End If
'        Next
'
'
'        ii = 0
'        jj = 0
'        kk = 0
'
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(8).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "2" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(9).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "3" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(10).childNodes.Item(1).nodeTypedValue
'
'        For i = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Length - 1
'            ii = ii + 1
'            For J = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Length - 1
'                For k = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Length - 1
'                    If k = 0 Then
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & CStr(ii) & "|"
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).Attributes.Item(k).baseName 'xNode.childNodes.Item(0).childNodes.Item(0).baseName
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "|"
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).Attributes.Item(k).nodeTypedValue 'xNode.childNodes.Item(0).childNodes.Item(0).nodeTypedValue
'                    End If
'
'                    For m = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Length - 1
'                        strRData = strRData & "" & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Item(m).baseName
'                        strRData = strRData & "" & "|"
'                        strRData = strRData & "" & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Item(m).nodeTypedValue
'                        strRData = strRData & "" & "|"
'                    Next
'                Next
'                'Debug.Print strRData
'                If strRData <> "" Then
'                    intBufCnt = intBufCnt + 1
'                    ReDim Preserve strRecvData(intBufCnt)
'                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & CStr(ii) & "|" & strRData
'                End If
'                strRData = ""
'
'            Next
'            'XXXCFTR -M3XXX
'        Next
'        Set Child_Node = Nothing
'
'    End If
'
'Exit Sub
'
'ErrXML:
'    MsgBox "파일오류"
'    Exit Sub
'
'End Sub

Private Sub cmdResult_Click()

    frmResult.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("선택한 결과를 전송하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
        With spdOrder
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = colCHECKBOX
                If .Value = 1 Then
                    Res = SaveTransData(lRow, spdOrder)
                    
                    If Res = -1 Then
                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdOrder, "저장실패", lRow, colSTATE
                    
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    
                    Else
                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdOrder, "저장완료", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                        
                    End If
                    spdOrder.Row = lRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub

Private Sub cmdSearch_Click()
    Dim i       As Integer
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intSeq      As Integer
        
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)

'    If gHOSP.MACHNM = "ACCESS2" Then
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
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

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

Private Sub cmdWork_Click()

    frmWorkList.Show vbModal
    
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
            

            Select Case UCase(gHOSP.MACHNM)
                        Case "HORIBA":          Call Phase_Serial_HORIBA
                        Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
                        Case "ACCESS2":         Call Phase_Serial_ACCESS2
                        
                        Case "XN1000":          Call Phase_Serial_XN1000
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                        Case "URINSCAN":        Call Phase_Serial_URINSCAN
                        Case "AVL9180":         Call Phase_Serial_AVL9180
                        Case "CA800_ASTM":      Call Phase_Serial_CA800_ASTM
                        Case "CA800":           Call Phase_Serial_CA800
                        Case "AU480":           Call Phase_Serial_AU480
                        Case "UROMETER720":     Call Phase_Serial_UROMETER720
                        Case "YUMIZEN":         Call Phase_Serial_YUMIZEN           '영인과학 HORIBA YUMIZEN H500
                        Case "UROMETER720":     Call Phase_Serial_UROMETER720
                        Case "XP300":           Call Phase_Serial_XP300
                        Case "ISMART30":        Call Phase_Serial_ISMART30
                        Case "STAGO":           Call Phase_Serial_STAGO
                        Case "PATHFAST":        Call Phase_Serial_PATHFAST
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                        'Case "KLITE":           Call Phase_Serial_KLITE
                        'Case "INDIKO":          Call Phase_Serial_INDIKO
                            
            End Select

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



Private Sub GetOrder_HITACHI7180(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    Dim strSndMsg   As String
    
    intRow = -1
    GetOrder = ""

    ''Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    ''Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        
        mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            'GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
            
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            'GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
        
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_HITACHI7020(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

'    'Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        mOrder.PID = Trim(GetText(frmMain.spdOrder, i, colCHARTNO))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarNo, intRow)
        
        '바코드를 사용하지 않을 경우에 사용한다.
        'If gHOSP.BARUSE <> "Y" Then
            mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.PID & Space(13), 13))
        'End If
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_STAGO(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_STAGO(gHOSP.MACHCD, pBarNo, intRow)
        
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_XN1000(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

'    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_XN1000(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_CA800(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    Dim SendBuf     As String
    
    intRow = -1
    GetOrder = ""

'    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_CA800(gHOSP.MACHCD, pBarNo, intRow)
        
        SendBuf = "S"
        SendBuf = SendBuf & "2"
        SendBuf = SendBuf & "21"
        SendBuf = SendBuf & "01"
        SendBuf = SendBuf & "01"
        SendBuf = SendBuf & "U"
        SendBuf = SendBuf & Format$(Date, "YYMMDD")
        SendBuf = SendBuf & Format$(Now, "HHMM")
        SendBuf = SendBuf & mOrder.RackNo
        SendBuf = SendBuf & mOrder.TubePos
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            SendBuf = SendBuf & Space(15)
            SendBuf = SendBuf & "C"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            SendBuf = SendBuf & Right(Space(15) & mOrder.BarNo, 15)
            SendBuf = SendBuf & "B"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        SendBuf = STX & SendBuf & ETX
        
        Call Sleep(500)
        
        Call SendData(SendBuf)

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_ACCESS2(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, pBarNo, intRow)
        
        mOrder.Order = strItems
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_YUMIZEN(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_YUMIZEN(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            'mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            'mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String

    intRow = -1

    '-- 1. 접수정보 조회
    With frmMain
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        strItems = ""
        mOrder.Order = ""
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarNo, intRow)

        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
            
        Else
            mOrder.NoOrder = False
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & strItems & ETX
        End If

        Call SendData(strOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

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

Private Sub SendWSckData(ByVal pSendData As Variant)

    '-- 전송
    wSck.SendData pSendData
    
'    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
'    If tmrSend.Enabled = False Then
'        tmrSend.Enabled = True
'    Else
'        tmrSend.Enabled = False
'        tmrSend.Enabled = True
'    End If
'    DoEvents
    
    '-- 로그기록
    Call SetRawData("[Tx]" & pSendData)

End Sub


Private Sub TCPRcvData_KLITE()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20190611090403||ORU^R01|TR14-009|P|2.4||||||ASCII
                    strHeader = mGetP(strRcvBuf, 10, "|")
                    strHeaderType = mGetP(strRcvBuf, 18, "|")
                    
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||" & strHeaderType & "|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|" & strHeader & "|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    'If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    'End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- 참고치
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- 남자참고치를 기본으로 한다
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- 결과Row 추가
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- 소수점 처리
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- 결과판정
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- 진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '-- 메인화면 결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '순번
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '장비채널
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS결과
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
                            
                            '-- 이전결과 조회
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '이전결과
                            
                            '-- H/L 색깔표시
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- 로컬 저장
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB에 결과저장
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_GENEXPERT()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    Dim strSend         As String
    
    Dim strMTB          As String
    Dim strRIF          As String
    Dim strCDIF         As String
    Dim str027          As String
    Dim strCarbaRPos    As String
    Dim strCarbaRNeg    As String
    
    Dim strMTBRIFCMT    As String
    Dim strCarbaRCMT    As String
    'Dim strCarbaRNeg    As String
    Dim strMachNum      As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = Mid(strRcvBuf, 1, 1)
            If IsNumeric(strType) Then
                strType = Mid(strRcvBuf, 2, 1)
            End If

            Select Case strType
                Case "H"
                    mResult.CARBAR_CMTCD = ""
                    mResult.MTBRIF_CMTCD = ""
                    mResult.CMNTCD = ""
                Case "P"
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    
                    If strBarno = "" Then
                        Exit Sub
                    End If
                
'''                    If Trim(strBarno) <> Trim(strOldBarno) Then
'''                        '-- 결과정보
'''                        With mResult
'''                            .BarNo = strBarno
'''                            .RsltDate = Format(Now, "yyyy-mm-dd")
'''                            .RsltTime = Format(Now, "hh:mm:ss")
'''                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'''                        End With
'''                    End If
'''
'''                    strOldBarno = strBarno
'''
'''                    '-- 결과환자정보
'''                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''
'''                    If gRow <= 0 Then
'''                        Exit Sub
'''                    End If
                
                
                Case "R"
                    'R|1|^MTB-RIF^^MTB^Xpert MTB-RIF Assay G4^6^MTB^|DETECTED HIGH^|||||F||<None>|20190819150251|20190819164413|Cepheid-642628D^820753^723731^785423426^24912^20210321
'A1: 699607
'A2: 699606
'A3: 699605
'A4: 699604
'B4: 723731
'B3: 723731
'B2: 723715
'B1: 722171
'
                    
                    strMachNum = mGetP(mGetP(strRcvBuf, 14, "|"), 3, "^")
                    mResult.EqpCd = "E13"
                    Select Case strMachNum
                        Case "699607", "699607", "699607", "699607"
                            mResult.EqpCd = "E13"
                        Case "723731", "723731", "723715", "722171"
                            mResult.EqpCd = "E14"
                    End Select
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                        
                    strOldBarno = strBarno
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strIntBase = mGetP(strRcvBuf, 3, "|")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = "" 'mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    
                    Call SetSQLData("RCV", strIntBase & ":" & strResult, "A")
                    
                    '-- MTB Ct값 찾기
'''                    If strIntBase = "^MTB-RIF^^MTB^^^Probe E^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 38 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
'''
'''                    '-- TOX Ct값 찾기
'''                    If strIntBase = "^G3^^Toxi^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 5 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
'''
'''                    '-- Carba-R 값 찾기
'''                    If strIntBase = "^Carba-R^^IMP1^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        'MTB
                        If strIntBase = "^MTB-RIF^^MTB^Xpert MTB-RIF Assay G4^6^MTB^" Then
                            strMTB = strResult
                        End If
                        'RIF
                        If strIntBase = "^MTB-RIF^^RIF^Xpert MTB-RIF Assay G4^6^Rif Resistance^" Then
                            strRIF = strResult
                        End If
                        
                        'Carba-R
                        'IMP
                        If strIntBase = "^Carba-R^^IMP1^Xpert Carba-R^2^IMP1^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "IMP1" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "IMP1" & "/"
                            End If
                        End If
                        'VIM
                        If strIntBase = "^Carba-R^^VIM^Xpert Carba-R^2^VIM^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "VIM" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "VIM" & "/"
                            End If
                        End If
                        'NDM
                        If strIntBase = "^Carba-R^^NDM^Xpert Carba-R^2^NDM^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "NDM" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "NDM" & "/"
                            End If
                        End If
                        'KPC
                        If strIntBase = "^Carba-R^^KPC^Xpert Carba-R^2^KPC^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "KPC" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "KPC" & "/"
                            End If
                        End If
                        'OXA48
                        If strIntBase = "^Carba-R^^OXA48^Xpert Carba-R^2^OXA48^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "OXA48" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "OXA48" & "/"
                            End If
                        End If
                        
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                        
                    End If

                    .spdResult.RowHeight(-1) = 15
                
                Case "L"
                    If strMTB = "NOT DETECTED" And strRIF = "" Then
                        strIntBase = "^MTB-RIF^^RIF^Xpert MTB-RIF Assay G4^6^Rif Resistance^"
                        strResult = "*"
                        
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    If strMTB = "NOT DETECTED" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되지 않았으나 결핵이 의심되면 타검사 결과를 확인하시기 바랍니다."
                    
                        mResult.MTBRIF_CMTCD = "TB2"
                    
                    ElseIf strMTB = "DETECTED VERY LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB1"
                        
                    ElseIf strMTB = "DETECTED LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB3"
                    
                    ElseIf strMTB = "DETECTED MEDIUM" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine
                    
                        mResult.MTBRIF_CMTCD = "TB4"
                    
                    ElseIf strMTB = "DETECTED HIGH" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine
                        
                        mResult.MTBRIF_CMTCD = "TB5"
                    
                    End If
            
                    If strRIF = "DETECTED" Then
                        If strMTB = "DETECTED VERY LOW" Then
                            mResult.MTBRIF_CMTCD = "RIF1"
                            
                        ElseIf strMTB = "DETECTED LOW" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF2"
                        
                        ElseIf strMTB = "DETECTED MEDIUM" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF3"
                        
                        ElseIf strMTB = "DETECTED HIGH" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF4"
                        
                        End If
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "Rifamin 내성연관 돌연변이가 검출되어 내성으로 판단됩니다." & vbNewLine
                    
                    End If
                    
                    mResult.MTBRIF_CMT = strMTBRIFCMT
                    
                    strMTB = ""
                    strRIF = ""
                    strMTBRIFCMT = ""
                    
                    If strCarbaRPos <> "" Then
                        strCarbaRPos = Mid(strCarbaRPos, 1, Len(strCarbaRPos) - 1)
                        strCarbaRPos = Replace(strCarbaRPos, "/", " ")
                        
                        strCarbaRCMT = ""
                        strCarbaRCMT = strCarbaRCMT & "[Comment]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "검출된 Carbapenemase 유전자형 : strCarbaRPos" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "환자의 검체에서 Carbapenemase 유전자가 검출되었습니다." & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "Carbapenemase-producing Enter obacteriaceae (CPE) 보균자로 판단됩니다." & vbNewLine
                        
                    Else
                        strCarbaRCMT = ""
                        strCarbaRCMT = strCarbaRCMT & "[Comment]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "검출된 Carbapenemase 유전자형 : 없음" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "본 검사는 KPC, NDM, VIM 및 OXA-48 이외의 검사에서 carbapenemase에 의해서 발생한 CRE나," & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "필요 시 CRE 선별배양검사(검사코드 : 40920)를 의뢰하시기 바랍니다." & vbNewLine
                        
                    End If
                    
                    mResult.CARBAR_CMT = strCarbaRCMT
                    strCarbaRNeg = ""
                    strCarbaRPos = ""
                    strCarbaRCMT = ""
                     
                    If mResult.MTBRIF_CMTCD <> "" Then
                        mResult.CMNTCD = mResult.MTBRIF_CMTCD
                    End If
            End Select
        Next
        
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_GENEXPERT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub



Private Sub TCPRcvData_PPC300N()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20190611090403||ORU^R01|TR14-009|P|2.4||||||ASCII
                    strHeader = mGetP(strRcvBuf, 10, "|")
                    strHeaderType = mGetP(strRcvBuf, 18, "|")
                    
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
'[Rx]MSH|^~\&|PKL|PKL PPC 300N|||20190807112436||ORU^R01|201908070001|p|2.3.1||||0||ASCII|||
'PID|1||||||||||||||||||||||||||||||
'OBR|1||201908070001|PKL^PKL PPC 300N||||||||||||||||||||||||||||||||||||||||||
'OBX|1|NM|1|CHOL|222|mg/dL|130.0-250.0|N|||F||0.232932|||Admin||
'
                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||" & strHeaderType & "|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|" & strHeader & "|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    'If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    'End If
                
                    strSeq = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strSeq) <> Trim(strOldBarno) Then
                        strOldBarno = strSeq
                        '-- 결과정보
                        With mResult
                            '.BarNo = strBarno
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    'strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strIntBase = mGetP(strRcvBuf, 5, "|")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    strIntResult = strResult
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If


                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_PPC300N" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_F200()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    Dim strSend         As String
    
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    
                Case "PID"
                    
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "OBR"

                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB & "MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|CA|{d4acc100-7cdd-45dd-bf26-83045c48fb0d}" & vbCr
                    strSend = strSend & EB & vbCr
                    
                    
                    Call SendWSckData(strSend)
                
                    strResultA = ""
                    strResultB = ""
                    strResultA_NTE = ""
                    strResultB_NTE = ""
                                
                
                Case "OBX"
                    strTmp = mGetP(strRcvBuf, 4, "|")
                    strTmp = mGetP(strTmp, 2, "^")
                    strIntBase = strTmp 'Mid(strTmp, 1, 11)
                    
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
'                    If strIntBase = "Influenza A" Then
'                        If UCase(strResult) = "NEGATIVE" Then
'                            strResultA = "NEGATIVE"
'                        End If
'
'                        If UCase(strResult) = "POSITIVE" Then
'                            strResultA = "POSITIVE"
'                        End If
'                    End If
'
'                    If strIntBase = "Influenza B" Then
'                        If UCase(strResult) = "NEGATIVE" Then
'                            strResultB = "NEGATIVE"
'                        End If
'
'                        If UCase(strResult) = "POSITIVE" Then
'                            strResultB = "POSITIVE"
'                        End If
'                    End If
                    
                    If strIntBase = "Influenza A" Then
                        strResultA = strResult
                    End If
                    
                    If strIntBase = "Influenza B" Then
                        strResultB = strResult
                    End If
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If

                    .spdResult.RowHeight(-1) = 15

'                Case "NTE"  'Device Information
'                    'NTE|1||Cut Off Index,Value=16.89
'
'                    strTmp = mGetP(strRcvBuf, 4, "|")
'
'                    If Mid(strTmp, 1, 6) <> "Device" Then
'                        If strIntBase = "Influenza A" Then
'                            'strResultA_NTE = mGetP(strTmp, 2, "=")
'                            'strResultA_NTE = Format(strResultA_NTE, "##0.00")
'                            'strResultA = strResultA & "(" & strResultA_NTE & ")"
'                            'strResult = strResultA
'
'                            strIntResult = mGetP(strTmp, 2, "=")
'                            strIntResult = Format(strIntResult, "##0.00")
'                            strResultA_NTE = strIntResult
'
'                        ElseIf strIntBase = "Influenza B" Then
'                            'strResultB_NTE = mGetP(strTmp, 2, "=")
'                            'strResultB_NTE = Format(strResultB_NTE, "##0.00")
'                            'strResultB = strResultB & "(" & strResultB_NTE & ")"
'                            'strResult = strResultB
'
'                            strIntResult = mGetP(strTmp, 2, "=")
'                            strIntResult = Format(strIntResult, "##0.00")
'                            strResultB_NTE = strIntResult
'
'                        End If
'                    End If
'

RST:
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If


                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
    
'        If strResultA <> "" And strResultB <> "" And strResultA_NTE <> "" And strResultB_NTE <> "" Then
'            strIntBase = "Influenza A/B"
'            If UCase(Mid(strResultA, 1, 8)) = "POSITIVE" And UCase(Mid(strResultB, 1, 8)) = "POSITIVE" Then
'                strResult = ""
'            ElseIf UCase(Mid(strResultA, 1, 8)) = "POSITIVE" And UCase(Mid(strResultB, 1, 8)) = "NEGATIVE" Then
'                'strResult = "A Positive"
'                strResult = "Positive" & "(type A:" & strResultA_NTE & ")"
'            ElseIf UCase(Mid(strResultA, 1, 8)) = "NEGATIVE" And UCase(Mid(strResultB, 1, 8)) = "POSITIVE" Then
'                'strResult = "B Positive"
'                strResult = "Positive" & "(type B:" & strResultB_NTE & ")"
'            ElseIf UCase(Mid(strResultA, 1, 8)) = "NEGATIVE" And UCase(Mid(strResultB, 1, 8)) = "NEGATIVE" Then
'                'strResult = "Negative"
'                If strResultA_NTE > strResultB_NTE Then
'                    strResult = "Negative" & "(" & strResultA_NTE & ")"
'                Else
'                    strResult = "Negative" & "(" & strResultB_NTE & ")"
'                End If
'            Else
'                strResult = ""
'            End If
'
'            strResultA = ""
'            strResultB = ""
'            strResultA_NTE = ""
'            strResultB_NTE = ""
'
'            GoTo RST
'
'        End If
        
        If strResultA <> "" And strResultB <> "" Then
            strIntBase = "Influenza A/B"
            If UCase(Mid(strResultA, 1, 8)) = "POSITIVE" And UCase(Mid(strResultB, 1, 8)) = "POSITIVE" Then
                strResult = ""
            ElseIf UCase(Mid(strResultA, 1, 8)) = "POSITIVE" And UCase(Mid(strResultB, 1, 8)) = "NEGATIVE" Then
                'strResult = "A Positive"
                strResult = "Positive"
                strIntResult = "type A:" & strResultA_NTE
            ElseIf UCase(Mid(strResultA, 1, 8)) = "NEGATIVE" And UCase(Mid(strResultB, 1, 8)) = "POSITIVE" Then
                'strResult = "B Positive"
                strResult = "Positive"
                strIntResult = "type B:" & strResultB_NTE
                
            ElseIf UCase(Mid(strResultA, 1, 8)) = "NEGATIVE" And UCase(Mid(strResultB, 1, 8)) = "NEGATIVE" Then
                strResult = "Negative"
                If strResultA_NTE > strResultB_NTE Then
                    strIntResult = strResultA_NTE
                Else
                    strIntResult = strResultB_NTE
                End If
            Else
                strResult = ""
            End If
            
            strResultA = ""
            strResultB = ""
            strResultA_NTE = ""
            strResultB_NTE = ""
            
            GoTo RST
        
        End If
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_PPC300N_OLD()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strTypeSeq      As String   '수신한 Record Type Seq
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    Dim strTemp         As String
    Dim strSend         As String
    Dim strOrder        As String
    Dim strLot          As String
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strTemp = mGetP(strRcvBuf, 4, "|")
            strType = mGetP(strTemp, 1, ";")
            strTypeSeq = mGetP(strTemp, 2, ";")
            
            Select Case strType
                Case "REQ"
                    If strTypeSeq = "1" Then
                        '오더요청 REQ;1
                        'Request information
                        'Start time         2010/11/01 00:00:00
                        'End time           2010/11/01 23:59:59
                        '<SB>|;^\|U8030|REQ;1|2010/11/01^00:00:00;2010/11/01^23:59:59|ASCII|<EB>

                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
''                        strSend = SB & "|;^\|LisDemo|" & "ASW;1||ASCII|" & EB
''                        SetRawData "[Tx]" & strSend
''                        wSck.SendData strSend
''                        '3.LIS Busy : <SB> |;^\|LISDEMO|ASW;1||ASCII|<EB>
''
''                        strSend = SB & "|;^\|LisDemo|" & "ASW;3||ASCII|" & EB
''                        SetRawData "[Tx]" & strSend
''                        wSck.SendData strSend
''                        '3.No Order : <SB> |;^\|LISDEMO|ASW;3||ASCII|<EB>
                        
                        strOrder = "" '5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP
                        intOrdCnt = 0
                        With spdOrder
                            For i = 1 To .MaxRows
                                .Row = i
                                .Col = colCHECKBOX
                                If .Value = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
                                    intOrdCnt = intOrdCnt + 1
                                    'strOrder = strOrder & GetText(spdOrder, i, colBARCODE) & ";" & GetText(spdOrder, i, colDEPT) & ";"
                                    
                                    strOrder = strOrder & GetText(spdOrder, i, colBARCODE) & ";"
                                    strOrder = strOrder & GetTag(spdOrder, i, colSTATE) & ";"
                                    
                                    Call SetText(spdOrder, "0", i, colCHECKBOX)
                                    Call SetText(spdOrder, "오더전송", i, colSTATE)
                                End If
                            Next
                        End With
                        
                        If strOrder = "" And intOrdCnt = 0 Then
                            strSend = SB & "|;^\|LisDemo|" & "ASW;3||ASCII|" & EB
                            SetRawData "[Tx]" & strSend
                            wSck.SendData strSend
                            '3.No Order : <SB> |;^\|LISDEMO|ASW;3||ASCII|<EB>
                        Else
                            strOrder = Mid(strOrder, 1, Len(strOrder) - 1)
                            strOrder = CStr(intOrdCnt) & ";" & strOrder
                            
                            strSend = SB & "|;^\|LisDemo|TRA;5|" & strOrder & "|ASCII|" & EB
                            SetRawData "[Tx]" & strSend
                            wSck.SendData strSend
                            '3.Order : SB & "|;^\|LisDemo|TRA;5|5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP|ASCII|" & EB
                        End If
                        
                        strState = "Q"
                        '<SB>|;^\|LisDemo|TRA;5|5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP|ASCII|<EB>
                    Else
                        '일반샘플 REQ;2
                        '샘플정보 REQ;3
                        'QC  정보 REQ;4
                        'Cal 정보 REQ;5
                        
                        'Request transferring results
                        '1.RCV  : <SB>|;^\|U8030|REQ;2|1234;2|ASCII|<EB>
                        strTemp = mGetP(strRcvBuf, 5, "|")
                        strBarno = Trim(mGetP(strTemp, 1, ";"))     'BarCode
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
                        strSend = SB & "|;^\|LisDemo|" & "ASW;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB> |;^\|LISDEMO|ASW;2||ASCII|<EB>
                    
                        If Trim(strBarno) <> Trim(strOldBarno) Then
                            strOldBarno = strBarno
                            '-- 결과정보
                            With mResult
                                .BarNo = strBarno
                                .RsltDate = Format(Now, "yyyy-mm-dd")
                                .RsltTime = Format(Now, "hh:mm:ss")
                                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            End With
                        End If
                        
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                        If gRow <= 0 Then
                            Exit Sub
                        End If
                        
                        strState = "O"
                        
                    End If
                    
                Case "ASK"
                    '--
                
                Case "TRA"
                    '1.RCV  : <SB>|;^\|U8030|TRA;2|1;201009200001;1234;;ALT;;43;;U/L;0;40;;;|ASCII|<EB>
                    strTemp = mGetP(strRcvBuf, 5, "|")
                    
                    '샘플정보
                    If strTypeSeq = "1" Then
                        '샘플정보 TRA;1
                        strSeq = mGetP(strTemp, 1, ";")
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;" & strSeq & "||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                    
                    '일반결과정보
                    ElseIf strTypeSeq = "2" Then
                        strIntBase = Trim(mGetP(strTemp, 5, ";"))   'Item
                        strResult = Trim(mGetP(strTemp, 7, ";"))    'Result
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;6|" & strBarno & ";" & strIntBase & "|ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                        
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    'QC 정보
                    ElseIf strTypeSeq = "3" Then
                        strIntBase = Trim(mGetP(strTemp, 1, ";"))   'Item
                        strLot = Trim(mGetP(strTemp, 4, ";"))   'Lot
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;" & strSeq & "||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                        
                    'Cal 정보
                    ElseIf strTypeSeq = "4" Then
                        strIntBase = Trim(mGetP(strTemp, 1, ";"))   'Item
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;7" & "|" & strLot & ";" & strIntBase & "|ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                    End If
                    
                    .spdResult.RowHeight(-1) = 15
                
                Case "END"
                    '1.RCV  : <SB>|;^\|U8030|END;1||ASCII|<EB>
                    
                    strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                    SetRawData "[Tx]" & strSend
                    wSck.SendData strSend
                    '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
                    strSend = SB & "|;^\|LisDemo|" & "REP;2||ASCII|" & EB
                    SetRawData "[Tx]" & strSend
                    wSck.SendData strSend
                    '2.SEND : <SB>|;^\|LisDemo|REP;2||ASCII| <EB>
            
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
            
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If

            End Select
        
        Next
    
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub



Private Sub SerialRcvData_HITACHI7180()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim strGFR          As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strCRP          As String
    Dim strRF           As String
    
    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    sFunc = Mid(strRcvBuf, 2, 40)
                    
                    With mOrder
                        .BarNo = strBarno
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                    End With
                    
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration 데이터는 무시함
'                    If UCase(strFunc) = "K" Or UCase(strFunc) = "L" Or UCase(strFunc) = "G" Or UCase(strFunc) = "H" Then
'                        '-- 장비로 전송
'                        Call SendData(SndMore)
'                        strState = ""
'                        Exit Sub
'                    End If
                    
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    If UCase(strFunc) = "F" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If
                    
                    strTC = ""
                    strTG = ""
                    strHDL = ""
                    
                    For ii = 51 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                        strComm = Trim(Mid(strRcvBuf, ii + 9, 1))
            
                        '-- CREA 결과저장
                        If Trim(strIntBase) = "2" Then
                            strGFR = ""
                            strResult = Format(strResult, "##0.00")
                            strCrea = strResult
                            
                            If mPatient.AGE <> "" And mPatient.SEX <> "" Then
                                If CCur(strResult) > 0 Then
                                    '18세 이상만 적용
                                    If IsNumeric(strCrea) And mPatient.AGE > 18 Then
                                        If mPatient.SEX = "M" Then
                                            strGFR = 186 * (strCrea ^ -1.154) * (mPatient.AGE ^ -0.203)
                                        ElseIf mPatient.SEX = "F" Then
                                            strGFR = 186 * (strCrea ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
                                        End If
                                        
                                        If strGFR <> "" Then
                                            strGFR = Format(strGFR, "##0.00")
                                            If strGFR <= 120 Then
                                                strGFR = Round(strGFR, 2)
                                            ElseIf strGFR > 120 Then
                                                strGFR = "> 120"
                                            End If
                                        End If
                                    End If
                                Else
                                    strGFR = "Error"
                                End If
                            End If
                        End If
                        
                        If strIntBase = "20" Then    'CRP
                            strCRP = strResult
                        End If
                        
                        If strIntBase = "21" Then    'RF
                            strRF = strResult
                        End If
                    
                        If strIntBase = "19" Then    'TCHO
                            strTC = strResult
                        End If

                        If strIntBase = "10" Then   'TG
                            strTG = strResult
                        End If

                        If strIntBase = "9" Then    'HDLC
                            strHDL = strResult
                        End If
                    
ReCal:
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                    Next
                    
                    'GA% = GA + (ALB - BCP)
                    
                    'LDL 계산
                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                        strIntBase = "95"
                        strResult = strTC - ((strTG / 5) + strHDL)
                        If strResult < 0 Then
                            strResult = "0"
                        End If
                        strTC = ""
                        strTG = ""
                        strHDL = ""
                        
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                    End If
                    
                    'CRP 정성
                    If strCRP <> "" Then
                        strIntBase = "87"
                        If strCRP < 0.5 Then
                            strResult = "Negative (" & strCRP & ")"
                        Else
                            strResult = "Positive (" & strCRP & ")"
                        End If
                        strCRP = ""
                        
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
                    
                    'RA Fact 정성
                    If strRF <> "" Then
                        strIntBase = "88"
                        If strCRP < 15 Then
                            strResult = "Negative (" & strRF & ")"
                        Else
                            strResult = "Positive (" & strRF & ")"
                        End If
                        strCRP = ""
                        
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
                    
                    '-- GFR 저장
                    If strGFR <> "" Then
                        strIntBase = "89"
                        strResult = strGFR
                        
                        strGFR = ""
                        
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    End If
    

                    
                    Call SendData(SndMore)
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SerialRcvData_H7180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_HITACHI7020()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
                    strFunc = Mid(strRcvBuf, 2, 1)              ' Function

                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                         Exit Sub
                    End If

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    If gHOSP.BARUSE = "Y" Then
                        '바코드 사용
                        sFunc = Mid(strRcvBuf, 2, 40)
                    Else
                        '바코드 미사용
                        sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    End If
                    
                    With mOrder
                        .BarNo = strBarno
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                    End With
                    
                    Call GetOrder_HITACHI7020(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    
                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration 데이터는 무시함
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    '## QC
                    If UCase(strFunc) = "F" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    ':N     1   1 300065626                    5  1    15   2    10   5     8   8    85  11   0.7

                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    mOrder.Seq = Trim(strSeq)
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If
                    
                    For ii = 45 To Len(strRcvBuf) Step 10
                        If strIntBase = "18" Then Stop
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        'strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 5))
                        strIntResult = strResult
                        
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                        'RA정량
                        If strIntBase = "20" Then
                            'RA정성
                            strIntBase = "99"
                            If IsNumeric(strResult) Then
                                If strResult > 15 Then
                                    strResult = "Positive"
                                Else
                                    strResult = "Negative"
                                End If
                            End If
                            
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                    Next
                    
                    Call SendData(SndMore)
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_H7020" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ResultProcess(ByVal pBarNo As String, ByVal pIntBase As String, ByVal pResult As String, ByVal pIntResult As String) As Boolean
    Dim RS_L            As ADODB.Recordset
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strSeqno        As String   '검사순번
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim strCheck        As String   '검사오더체크
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strRstType      As String
    Dim i               As Integer
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCol          As Integer  '결과컬럼 갯수
    
    ResultProcess = False
    
    strSeqno = ""
    strTestCode = ""
    strTestName = ""
    strAbbrName = ""
    intResPrecUse = -1
    intResPrec = -1
    strAMRResult = ""
    
    SQL = ""
    SQL = SQL & "SELECT EQPMASTER.TESTCODE,TESTNAME,ABBRNAME,EQPMASTER.SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC,RESTYPE   " & vbCrLf
    SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7               " & vbCrLf
    SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7              " & vbCrLf
    SQL = SQL & "     , AMRLimit8,  AMRLimit9,  AMRLimit10,  AMRLimit11,  AMRLimit12,  AMRLimit13,  AMRLimit14          " & vbCrLf
    SQL = SQL & "     , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14         " & vbCrLf
    SQL = SQL & "     , AMRINResult                                                                                     " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER , AMRMASTER                                                                           " & vbCrLf
    SQL = SQL & " WHERE EQPMASTER.EQUIPCD     = '" & gHOSP.MACHCD & "'                                                            " & vbCrLf
    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = '" & pIntBase & "'                                                                " & vbCrLf
    If gPatOrdCd <> "" Then
        SQL = SQL & "   AND EQPMASTER.TESTCODE in (" & gPatOrdCd & ") "
    End If
    SQL = SQL & "   AND EQPMASTER.EQUIPCD     = AMRMASTER.EQUIPCD                                                       " & vbCrLf
    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = AMRMASTER.RSLTCHANNEL                                                   " & vbCrLf
    SQL = SQL & "   AND EQPMASTER.TESTCODE    = AMRMASTER.TESTCODE                                                      "
    
    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        strSeqno = Trim(RS_L.Fields("SEQNO"))
        strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
        strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
        strAbbrName = Trim(RS_L.Fields("ABBRNAME")) & ""
        
        '-- 참고치
        If mPatient.SEX = "M" Then
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        ElseIf mPatient.SEX = "F" Then
            strLow = Trim(RS_L.Fields("REFFLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
        Else
            '-- 남자참고치를 기본으로 한다
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        End If
        
        '-- 소수점변환 사용여부와 변환자리수
        intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
        intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
                
        '사용결과
        strResType = Trim(RS_L.Fields("RESTYPE")) & ""
        
        '-- 0:수치,1:판정,2:수치/판정
        If strResType = 0 Then
            '-- 소수점 처리
            strMachResult = pIntResult
            If intResPrecUse = 1 Then
                For i = 0 To intResPrec
                    If i = 0 Then
                        strResType = "#0"
                    ElseIf i = 1 Then
                        strResType = strResType & ".0"
                    Else
                        strResType = strResType & "0"
                    End If
                Next
                pIntResult = Format(pIntResult, strResType)
            End If
                
            '-- AMR 적용 (수치형)
            If IsNumeric(pIntResult) Then
                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                    End If
                End If
                If strAMRResult = "" Then
                    strAMRResult = pIntResult
                End If
            End If
            
        ElseIf strResType = 1 Then
            '-- AMR 적용 (문자형)
            If pResult <> "" Then
                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                    End If
                End If
                'add
                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT8"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT9"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT10"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT11"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT12"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT13")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT13")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT13"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT14")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT14")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT14"))
                    End If
                End If
                If strAMRResult = "" Then
                    strAMRResult = pResult
                End If
            End If
            
        ElseIf strResType = 2 Then
            '-- 소수점 처리
            strMachResult = pIntResult
            If intResPrecUse = 1 Then
                For i = 0 To intResPrec
                    If i = 0 Then
                        strResType = "#0"
                    ElseIf i = 1 Then
                        strResType = strResType & ".0"
                    Else
                        strResType = strResType & "0"
                    End If
                Next
                pIntResult = Format(pIntResult, strResType)
            End If
                
            '-- AMR 적용 (수치형)
            If IsNumeric(pIntResult) Then
                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                    End If
                End If
                If strAMRResult = "" Then
                    strAMRResult = pIntResult
                End If
            End If
                                    
            '-- AMR 적용 (문자형)
            If pResult <> "" Then
                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                    End If
                End If
                'add
                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT8"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT9"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT10"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT11"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT12"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT13")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT13")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT13"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT14")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT14")) Then
                        strAMRResult = Trim(RS_L.Fields("AMRRESULT14"))
                    End If
                End If
                If strAMRResult = "" Then
                    strAMRResult = pResult
                End If
            End If
        End If
        
        '수치결과 포함
        '0:사용안함, 1:정성(정량), 2:정량(정성), 3:정성_정량, 4:정량_정성
        If strAMRResult <> "" Then
            If pIntResult <> "" Then
                If Trim(RS_L.Fields("AMRINResult") & "") = "0" Then
                    pResult = strAMRResult
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                    pResult = strAMRResult & "(" & pIntResult & ")"
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "2" Then
                    pResult = pIntResult & "(" & strAMRResult & ")"
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "3" Then
                    pResult = strAMRResult & " " & pIntResult
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "4" Then
                    pResult = pIntResult & " " & strAMRResult
                End If
            Else
                pResult = strAMRResult
            End If
        End If
    
        '--- 결과판정
        strJudge = ""
        If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
            If IsNumeric(pResult) Then
                If CCur(pResult) > CCur(strLow) And CCur(pResult) < CCur(strHigh) Then
                    strJudge = ""
                ElseIf CCur(strHigh) <= CCur(pResult) Then
                    strJudge = "H"
                ElseIf CCur(strLow) >= CCur(pResult) Then
                    strJudge = "L"
                End If
            End If
        End If
        
        With frmMain
            '-- 결과Row 추가
            intRstRow = .spdResult.DataRowCnt + 1
            If .spdResult.MaxRows < intRstRow Then
                .spdResult.MaxRows = intRstRow
            End If
    
            '-- 진행상태 표시("결과")
            SetText .spdOrder, "결과", gRow, colSTATE
    
            '-- 메인화면 결과값 표시
            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                If strAbbrName = gArrEQP(intCol - colSTATE, 6) Then
                    SetText .spdOrder, pResult, gRow, intCol
                    
                    '-- H/L 색깔표시
                    If strJudge = "H" Then
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbBlue
                    Else
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbBlack
                    End If
                    
                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
                    
                    Exit For
                End If
            Next
    
            '-- 결과 List
            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '순번
            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
            SetText .spdResult, pIntBase, intRstRow, colRCHANNEL              '장비채널
            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
            SetText .spdResult, pResult, intRstRow, colRLISRESULT             'LIS결과
            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
            
            '-- 이전결과 조회
            strPrevRslt = GetPrevResult(mResult.BarNo, pIntBase, strTestCode)
            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '이전결과
            
            '-- H/L 색깔표시
            If strJudge = "H" Then
                .spdResult.Row = intRstRow
                .spdResult.Col = colRLISRESULT
                .spdResult.ForeColor = vbRed
                .spdResult.FontBold = True
            ElseIf strJudge = "L" Then
                .spdResult.Row = intRstRow
                .spdResult.Col = colRLISRESULT
                .spdResult.ForeColor = vbBlue
                .spdResult.FontBold = True
            Else
                .spdResult.Row = intRstRow
                .spdResult.Col = colRLISRESULT
                .spdResult.ForeColor = vbBlack
                .spdResult.FontBold = False
            End If
            
            '-- 결과Count
            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                SetText .spdOrder, "1", gRow, colRCNT
            Else
                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
            End If
        End With
        
        '-- 로컬 저장
        Call SetLocalDB(gRow, intRstRow, "1", "")
        
        ResultProcess = True
    End If
    
End Function

Private Sub SerialRcvData_XN1000()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                    strWBC = ""
                    strNeut = ""
                
                Case "Q"    '## Request Information
                    '2Q|1|15^8^            1000001207^B||||20190904144851||||||N

                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_XN1000(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strBarno = Trim(mGetP(strTemp1, 3, "^"))
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    '7R|1|^^^^WBC^1|6.15|10*3/uL||N||F||||20190904083314

                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 5, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strIntResult = strTemp2
                        strResult = strTemp2
                    End If
                    
'                    If strIntBase = "WBC" And IsNumeric(strResult) Then
'                        strWBC = strResult * 1000
'                    End If
'
'                    If strIntBase = "NEUT%" And IsNumeric(strResult) Then
'                        strNeut = strResult / 100
'                    End If
                    
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

'                    If strWBC <> "" And strNeut <> "" Then
'                        ''ANC = (wbc * 1000 * neut%) / 100
'                        strIntBase = "ANC"
'                        strResult = (strWBC * strNeut)
'                        strResult = Format(strResult, "##0")
'                        strWBC = ""
'                        strNeut = ""
'                        GoTo RST
'                    End If
                    
                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_MULTIPLATE()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "O"
                    strBarno = mGetP(strRcvBuf, 2, "|")

                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(strRcvBuf, 2, "|")
                    strResult = mGetP(strRcvBuf, 3, "|")
                    '단위제거
                    strResult = mGetP(strResult, 1, " ")
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_CA800()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim sBC$, sLC$
    Dim strTemp$
    Dim strInfo$
    
    
On Error GoTo RST

    strRcvBuf = strRecvData(1) 'strBuffer
    
    Call SetSQLData("RCV", strRcvBuf, "A")
    
    sBC = Mid(strRcvBuf, 1, 2)
    sLC = Mid(strRcvBuf, 3, 1)
    
    With frmMain
        Select Case sBC
            'R2210101U0904191511000501     1000001216B           040      050      [Tx]
            Case "R1", "R2"
                strBarno = Trim(Mid(strRcvBuf, 26, 15))
                
                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                    .RackNo = Trim(Mid(strRcvBuf, 20, 4))
                    .TubePos = Trim(Mid(strRcvBuf, 24, 2))
                End With
                
                Call GetOrder_CA800(strBarno, gHOSP.RSTTYPE)
                
                strState = "Q"
                
            Case "D1"
                strBarno = Trim(Mid(strRcvBuf, 26, 15))
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = Trim(Mid(strRcvBuf, 20, 4))
                    .TubePos = Trim(Mid(strRcvBuf, 24, 2))
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                        
                strTemp = Mid(strRcvBuf, 53)
                
                For i = 1 To Len(strTemp) Step 9
                    strIntBase = Mid$(strTemp, i, 3)
                    strType = Mid$(strIntBase, 3, 1)
                    strResult = Trim$(Mid$(strTemp, i + 3, 5))
                    strFlag = Trim(Mid$(strTemp, i + 8, 1))
                    strInfo = GetInfo(strFlag)
                    
                    Select Case strType
                        Case "1"    '## Time
                            strResult = Trim$(Format$(strResult, "@@@@.@"))
                            If strFlag = "*" Or InStr(strResult, "*") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                        Case "2"    '## Activity percent/concentration
                            strResult = Trim$(Format$(strResult, "@@@.@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            ElseIf Mid$(strIntBase, 1, 2) = "04" Then
                                '   - PT %값이 100이상이면 의미없는 결과라서 "100이상"으로 결과저장
                                '     하는것으로 수정
                                strResult = IIf(Val(strResult) > 100, ">100", strResult)
                            End If
                        Case "3"    '## Ratio
                            strResult = Trim$(Format$(strResult, "@.@@@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                        Case "4"    '## INR
                            strResult = Trim$(Format$(strResult, "@@.@@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                    End Select
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                Next
                
                .spdResult.RowHeight(-1) = 15

                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
        End Select
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Result Flags에 대한 상세설명 조회
'-----------------------------------------------------------------------------'
Private Function GetInfo(ByVal pFlag As String)
    Dim strInfo     As String

    If pFlag = "" Then Exit Function

    Select Case pFlag
        Case "+":   strInfo = "Over the upper control limit"
        Case "-":   strInfo = "Under the lower control limit"
        Case "*":   strInfo = "Analysis error occurred, disparate data of mean data occurred, or Fbg was over analysis range."
        Case "!":   strInfo = "Coagulation time was obtained by re-dilution analysis."
        Case ">":   strInfo = "Over the upper report limit."
        Case "<":   strInfo = "Under the lower report limit."
    End Select

    GetInfo = strInfo
End Function

Private Sub SerialRcvData_CA800_ASTM()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_CA800(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim(mGetP(strTemp1, 3, "^"))

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    
                    strTemp2 = mGetP(strRcvBuf, 5, "|")
                    strFlag = mGetP(strRcvBuf, 8, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strResult = strTemp2
                        strIntResult = strTemp2
                    End If
                    
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XP300()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

   ' ReDim Preserve strRData(UBound(strRecvData))
    
   ' strRData = strRecvData
    
    strRData = Split(RcvBuffer, vbCr)
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = Trim(mGetP(strTemp1, 1, "^"))
                    strTubePos = Trim(mGetP(strTemp1, 2, "^"))
                    strBarno = Trim(mGetP(strRcvBuf, 3, "^"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    strResult = ""
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strIntResult = strTemp2
                    End If
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XP300" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_URINSCAN()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim Pos             As Integer
        
On Error GoTo RST
    
    With frmMain
        Pos = InStr(RcvBuffer, "ID_NO")
        If Pos > 0 Then
            RcvBuffer = Replace(RcvBuffer, vbLf, "")
            strRecvData = Split(RcvBuffer, vbCr)
            
            '-- SEQ 번호 찾기
            strRcvBuf = strRecvData(2)
            strRcvBuf = mGetP(strRcvBuf, 2, ":")
            strRcvBuf = mGetP(strRcvBuf, 1, "-")
            strSeq = Trim(strRcvBuf)
            
            strBarno = strSeq
            With mResult
                .BarNo = strBarno
                .TubePos = strSeq
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
                    
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            
            For intCnt = 4 To UBound(strRecvData)
                strRcvBuf = strRecvData(intCnt)
                
                strType = Trim(Mid$(strRcvBuf, 1, 3))
                strIntBase = strType
                strResult = ""

                Select Case strType
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
                
                        
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                .spdResult.RowHeight(-1) = 15
            Next
                
            '## DB에 결과저장
            If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                Res = SaveTransData(gRow, spdOrder)

                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "저장실패", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX

                          SQL = "Update PATRESULT Set                                                               " & vbCrLf
                    SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                    SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                    SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                    SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                    SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
                
                spdOrder.Row = gRow
                spdOrder.Col = colCHECKBOX
                spdOrder.Value = 0
            End If
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_URINSCAN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_AVL9180()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim Pos             As Integer
        
On Error GoTo RST
    
    With frmMain
        strRcvBuf = RcvBuffer
        
        Call SetSQLData("RCV", strRcvBuf, "A")
        
        If InStr(RcvBuffer, "Na=") > 0 Or InStr(RcvBuffer, "K =") > 0 Or InStr(RcvBuffer, "Cl=") > 0 Then
            strIntBase = Trim(Mid(RcvBuffer, 1, 2))
            strResult = Trim(Mid(RcvBuffer, 4, 5))
            
            If strIntBase = "Na" Then
                mResult.strNa = strResult
                mResult.strK = ""
                mResult.strCl = ""
            ElseIf strIntBase = "K" Then
                mResult.strK = strResult
                mResult.strCl = ""
            ElseIf strIntBase = "Cl" Then
                mResult.strCl = strResult
            End If
                
            If mResult.strCl <> "" Then
                With mResult
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
        
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                                
                For i = 1 To 3
                    strIntBase = ""
                    strResult = ""
                    Select Case i
                        Case 1: strIntBase = "Na": strResult = mResult.strNa
                        Case 2: strIntBase = "K":  strResult = mResult.strK
                        Case 3: strIntBase = "Cl": strResult = mResult.strCl
                    End Select
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                
                Next
                       
                mResult.strNa = ""
                mResult.strK = ""
                mResult.strCl = ""
                
                .spdResult.RowHeight(-1) = 14
                
                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)
    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
    
                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            End If
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AVL9180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_PATHFAST()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim(mGetP(strTemp1, 1, "^"))
                    strSeq = Trim(mGetP(strTemp1, 2, "^"))
                    
                    '-- 결과정보
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = strBarno
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_PATHFAST" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_VISION()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    strRData = Split(pBuffer, vbLf)
    
    With frmMain
        For intCnt = 0 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            Call SetSQLData("RCV", strRcvBuf, "A")
            
            If Len(strRcvBuf) > 20 Then
                strIntBase = "ESR"
                strSeq = mGetP(strRcvBuf, 1, vbTab)
                strBarno = mGetP(strRcvBuf, 7, vbTab)
                '-- 18도 사용
                strResult = mGetP(strRcvBuf, 10, vbTab)
                strIntResult = mGetP(strRcvBuf, 10, vbTab)
                'strResult = mGetP(strRcvBuf, 11, vbTab)

                '-- 결과정보
                With mResult
                    .BarNo = strBarno
                    .Seq = strSeq
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                        
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                .spdResult.RowHeight(-1) = 15

                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            End If
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_KLITE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ISMART30()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_YUMIZEN()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^289645146||ALL||||||||O<CR><ETX>F7<CR><LF
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))

                    With mOrder
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder(Trim$(strBarno), gHOSP.RSTTYPE)
                    'Call GetOrder_YUMIZEN(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_STAGO()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_STAGO(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = mGetP(strTemp1, 1, "^")
                    strSeq = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    strBarno = Replace(strBarno, "_", "1")
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .Seq = strSeq
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    Select Case strFlag
                        Case "F"    '## 정량
                            strIntResult = strIntResult
                        Case "I"    '## 정성
                            Select Case Mid$(strIntResult, 1, 1)
                                Case "N":   strResult = "Negative"
                                Case "G":   strResult = "GRAYZONE"
                                Case "R":   strResult = "Positive"
                                Case "P":   strResult = "Positive"
                            End Select
                    End Select
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_STAGO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ACCESS2()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^190807015||ALL||||||||O

                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_ACCESS2(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    '3O|1|190807015|^1403^1|^^^HCG5^1|||||||||||Serum||||||||||F
                    '4R|1|^^^HCG5^1|>1342.00|mIU/mL|0.00 to 5.00^normal|>|N|F||||20190807153839|511896
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    strRackNo = Format(strRackNo, "0000")
                    strTubePos = Format(strTubePos, "00")
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo_ACCESS2(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    '4R|1|^^^hLH^1|17.28|mIU/mL||N||F||||20190731123358|511896
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strIntResult = mGetP(strTemp1, 1, "^")
                    strResult = mGetP(strTemp1, 2, "^")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If strResult = "" Then
                        strResult = strIntResult
                    End If
'                    If strIntBase = "HBsAgV3" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 1 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    'HbsAb
'                    '4R|1|^^^HBAb3^1|0.7|mIU/mL||N||F||||20190415103432|510062
'                    ElseIf strIntBase = "HBAb3" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 10 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    'HCV
'                    '4R|1|^^^HCVPLUS^1|0.10^Non-React.|S/CO||N||F||||20190415103620|510062
'                    ElseIf strIntBase = "HCVPLUS" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 1 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    Else
'                        strResult = strIntResult
'                    End If
                    
                    
'                    Select Case strFlag
'                        Case "F"    '## 정량
'                            strResult = strIntResult
'                        Case "I"    '## 정성
'                            Select Case Mid$(strIntResult, 1, 1)
'                                Case "N":   strResult = "Negative"
'                                Case "G":   strResult = "GRAYZONE"
'                                Case "R":   strResult = "Positive"
'                                Case "P":   strResult = "Positive"
'                            End Select
'                    End Select
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ACCESS2" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_UROMETER720()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    strRData = Split(RcvBuffer, vbCrLf)
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            Select Case intCnt
                Case 3
                    strSeq = Mid(strRcvBuf, 10)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))
                    
                    '-- 결과정보
                    mOrder.Seq = strSeq
                    With mResult
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    strResult = ""
                    strIntResult = ""
                    strResult = Mid(strRcvBuf, 8, 4) '-- 정성
                    strResult = Trim(strResult)
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        strIntResult = Trim(Mid(strRcvBuf, 4))  '-- 정량
                        strIntResult = Replace(strIntResult, "mg/dl", "")
                        strIntResult = Replace(strIntResult, "RBC/ul", "")
                        strIntResult = Replace(strIntResult, "WBC/ul", "")
                        
                        strIntResult = Replace(strIntResult, "<", "")
                        strIntResult = Replace(strIntResult, ">", "")
                        strIntResult = Replace(strIntResult, "=", "")
                    End If
                    
                    '-- URO
'                    If strResult = "norm" Then
'                        strResult = "Negative"
'                    End If
'
'                    '-- NIT
'                    If strResult = "pos" Then
'                        strResult = "1+"
'                    End If
'
'                    Select Case Trim(strResult)
'                        Case "-":       strResult = "Negative"
'                        Case "+":       strResult = "Pos(1+)"
'                        Case "++":      strResult = "Pos(2+)"
'                        Case "+++":     strResult = "Pos(3+)"
'                        Case "++++":    strResult = "Pos(4+)"
'                        Case "+/-":     strResult = "Trace(±)"
'                    End Select
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_UROMETER720" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HORIBA()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    strRData = Split(RcvBuffer, vbCr)
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            Select Case intCnt
                Case 4
                    If InStr(strRcvBuf, "AUTO_SID") > 0 Then
                        strSeq = Mid(strRcvBuf, InStr(strRcvBuf, "AUTO_SID") + 8)
                    Else
                        strSeq = mGetP(strRcvBuf, 2, Space(1))
                        strSeq = Val(strSeq)
                    End If
                    
                    '-- 결과정보
                    mOrder.Seq = strSeq
                    strBarno = strSeq
                    With mResult
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case 9 To 27
                    strIntBase = Trim(Mid(strRcvBuf, 1, 2))
                    strResult = Trim(Mid(strRcvBuf, 3))
                    strResult = Replace(strResult, "h", "")
                    strResult = Replace(strResult, "H", "")
                    strResult = Replace(strResult, "l", "")
                    strResult = Replace(strResult, "L", "")
                    strResult = Replace(strResult, "s", "")
                    strResult = Replace(strResult, "S", "")
                    strResult = Replace(strResult, " ", "")
                    strIntResult = strResult
                    
                    If strIntBase = "'" Then
                        strIntBase = "|"
                    End If
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_HORIBA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-- 장비설정
'System / System Condition /
'   [Test Requisition]
'       Routine:  BARCODE
'   [S.ID Barcode]
'       Barcode Type    : Multi
'       Digits          : 10
'       Check Mode      : No(No Chk.Chr.)
'System / Format /
'   Sample ID   Digits  : 20
Private Sub SerialRcvData_AU480()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqno        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            
            strType = Mid$(strRcvBuf, 1, 2)

            Select Case strType
                Case "R "    '## Inquiry Order
                    'R 003201 0018          1013001917
                    'S 003201 0018          1013001917    E      13
                    
                    strBarno = Trim(Mid(strRcvBuf, 14, 20))
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = Mid(strRcvBuf, 9, 5)
                    End With
                    
                    Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        
                Case "D "    '## Result
                    'D 000103 0003          1908130030    E107  2.35  
                    
                    strBarno = Trim$(Mid$(strRcvBuf, 14, 20))
                    mResult.BarNo = strBarno
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = Mid(strRcvBuf, 3, 4)
                        .TubePos = Mid(strRcvBuf, 7, 2)
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    If strBarno = "" Then Exit Sub
    
                    strTmp = Mid$(strRcvBuf, 39)
                                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                
                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop
                
                    .spdResult.RowHeight(-1) = 15
                    

                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AU480" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub Phase_Serial_HITACHI7180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_HITACHI7180
                        
                    Case vbCr
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_HITACHI7020()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_HITACHI7020
                        
                    Case vbCr
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_UROMETER720()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)


    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case "~"
                        RcvBuffer = ""
                        intPhase = 2
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
            Case 2
            
                Select Case BufChar
                    Case "~"
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_UROMETER720
                        RcvBuffer = ""
                        intPhase = 1
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_URINSCAN()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        RcvBuffer = ""
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_URINSCAN
                        RcvBuffer = ""
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i
End Sub

Private Sub Phase_Serial_AVL9180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbLf
                Call SerialRcvData_AVL9180
                
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_XP300()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            'Call SendOrder_XP300
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_XP300
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_YUMIZEN()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_YUMIZEN
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_YUMIZEN
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub


Private Sub Phase_TCP_XP300()
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                intFrameNo = intFrameNo + 1
                RcvBuffer = RcvBuffer & BufChar
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
    If InStr(RcvBuffer, "L|1|N") > 0 Then
        intPhase = 1
        intBufCnt = 0
        
        Call SerialRcvData_XP300
        
        intFrameNo = 0
        
    End If
    
End Sub

Private Sub Phase_TCP_YUMIZEN()
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                intFrameNo = intFrameNo + 1
                RcvBuffer = RcvBuffer & BufChar
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
    If InStr(RcvBuffer, "L|1|N") > 0 Then
        intPhase = 1
        intBufCnt = 0
        
        Call SerialRcvData_YUMIZEN
        
        intFrameNo = 0
        
    End If
    
End Sub

Private Sub Phase_Serial_ISMART30()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            'Call SendOrder_XP300
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_ISMART30
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub SendOrder_STAGO()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            
            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                intSndPhase = 3
            Else
                intSndPhase = 2
            End If

            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|||" & mOrder.PID & "|^1^1^56|||19700505" & vbCr & ETX
            intSndPhase = 4
            intFrameNo = intFrameNo + 1

        Case 3  '## No Order
            strOutput = intFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            intSndPhase = 5

        Case 4  '## Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 4
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 5
                End If
            '## 남은 문자열이 있을때
            Else
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
            intFrameNo = intFrameNo + 1

        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1

        Case 6  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_XN1000()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800_ASTM()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            '<STX>1                   H|\^&|||HostName^^^^|||||CA-600<CR><ETX><CHK1><CHK2><CR><LF
            strOutput = intFrameNo & "H|\^&|||HostName^^^^|||||CA-600" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            '<STX>1                   H|\^&|||HostName^^^^|||||CA-600<CR><ETX><CHK1><CHK2><CR><LF
            strOutput = intFrameNo & "H|\^&|||HostName^^^^|||||CA-600" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


Private Sub SendOrder_ACCESS2()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer
    Dim intDestRow  As Integer

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|" & mOrder.PID & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## No Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                'strOutput = "O|1|" & mOrder.BarNo & "|" & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "|" & mOrder.Order & "|R||||||A||||" & "Serum"
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & "Serum"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            intSndPhase = 1
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)
    
End Sub

Private Sub SendOrder_ACCESS2_Batch()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer
    Dim intDestRow  As Integer
    Dim blnOrder    As Boolean
    Dim blnLast     As Boolean

    blnOrder = False
    blnLast = True
    
    With spdOrder
        If intSndPhase = 2 Or intSndPhase = 3 Then
            For intRow = 1 To .MaxRows
                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
                    mOrder.BarNo = Trim(GetText(spdOrder, intRow, colBARCODE))
                    mOrder.PID = Trim(GetText(spdOrder, intRow, colPID))
                    mOrder.RackNo = Trim(GetText(spdOrder, intRow, colRACKNO))
                    mOrder.TubePos = Trim(GetText(spdOrder, intRow, colPOSNO))
                    'mOrder.Order = Trim(GetText(spdOrder, intRow, colDEPT))
                    mOrder.Order = Trim(GetTag(spdOrder, intRow, colSTATE))
                    mOrder.DestRow = intRow
                    'blnOrder = True
                    'intDestRow = intRow
                    Exit For
                End If
            Next
'            For intRow = intDestRow + 1 To .MaxRows
'                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
'                    blnLast = False
'                    Exit For
'                End If
'            Next
        End If
    End With
    
    If blnOrder = True Then
        Select Case intSndPhase
            Case 1  '## Header
                strOutput = intFrameNo & "H|\^&|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
                intSndPhase = 2
                intFrameNo = intFrameNo + 1
                
            Case 2  '## Patient
                strOutput = intFrameNo & "P|1|" & mOrder.PID & vbCr & ETX
                intSndPhase = 3
                intFrameNo = intFrameNo + 1
    
            Case 3  '## No Order
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    'strOutput = "O|1|" & mOrder.BarNo & "|" & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "|" & mOrder.Order & "|R||||||A||||" & "Serum"
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & "Serum"
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
                intFrameNo = intFrameNo + 1
    
            Case 4  '## Termianator
                strOutput = intFrameNo & "L|1|N" & vbCr & ETX
                intSndPhase = 5
                intFrameNo = intFrameNo + 1
    
            Case 5  '## EOT
                strState = ""
                Call SendData(EOT)
                intFrameNo = 1
                intSndPhase = 1
                
                Call SetText(spdOrder, "0", mOrder.DestRow, colCHECKBOX)
                Call SetText(spdOrder, "오더전송", mOrder.DestRow, colSTATE)
                
                blnLast = True
                For intRow = mOrder.DestRow + 1 To spdOrder.MaxRows
                    If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
                        blnLast = False
                        Exit For
                    End If
                Next

                If blnLast = False Then
                    strState = "Q"
                    Call SendData(ENQ)
                End If
                Exit Sub
        End Select
    
        If intFrameNo = 8 Then
            intFrameNo = 0
        End If
    
        strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
        Call SendData(strOutput)
    End If
    
End Sub


Private Sub SendOrder_YUMIZEN()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer

    Select Case intSndPhase
        Case 1  '## Header
                                     'H|\^&|||HCM|||||||P|LIS2-A2|20150323160111<CR><ETX>51<CR><LF
            strOutput = intFrameNo & "H|\^&|||HCM|||||||P|LIS2-A2|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
                                     'P|1||2                 ||BOND^JAMES||19770526|M|||||<CR><ETX>24<CR><LF
            strOutput = intFrameNo & "P|1||" & mOrder.PID & "||^||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## Order
                                    '3O|1|289645146||^^^DIF|R|20150323160111|||||N||||||||||||||Q|||||<CR><ETX>C0<CR><LF
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N||||||||||||||Q|||||"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
                                    '4L|1|<CR><ETX>B9<CR><LF
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            
            With spdOrder
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = colBARCODE
                    If Trim(.Text) = mOrder.BarNo Then
                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
                        Exit For
                    End If
                Next
            End With
            
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)

End Sub

Private Sub Phase_Serial_STAGO()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_STAGO
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_XN1000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_XN1000
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr, vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_XN1000
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

    
End Sub

Private Sub Phase_Serial_CA800_ASTM()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_CA800
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr, vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_CA800
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

    
End Sub

Private Sub Phase_Serial_CA800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
                
            Case ETX
                Call Sleep(200)         '0.2 sec or More Delay
                
                Call SendData(ACK)
                
                Call SerialRcvData_CA800
            
            Case ACK
            
            Case NAK
            
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i


    
End Sub

Private Sub Phase_Serial_ACCESS2()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_ACCESS2
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_ACCESS2
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_PPC300N()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_STAGO
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_PATHFAST()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'
'        Select Case BufChar
'            Case ENQ
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'                comEqp.Output = ACK
'                SetRawData "[Tx]" & ACK
'            Case STX
'                intBufCnt = intBufCnt + 1
'                ReDim Preserve strRecvData(intBufCnt)
'
'            Case vbLf
'                comEqp.Output = ACK
'                SetRawData "[Tx]" & ACK
'
'            Case EOT
'                dtpToday.Value = Now
'                Call SerialRcvData_PATHFAST
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
'                            Call SendOrder_PATHFAST
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_PATHFAST
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_AU480()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            Case ETB
            Case ETX
                lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                
                Call SerialRcvData_AU480
                RcvBuffer = ""
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    
    Next i
    
End Sub

Private Sub Phase_Serial_HORIBA()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
                RcvBuffer = RcvBuffer & BufChar
            Case ETX
                lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                
                Call SerialRcvData_HORIBA
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub


Private Sub Phase_TCP_KLITE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_KLITE
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub

Private Sub Phase_TCP_F200()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_F200
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub


Private Sub Phase_TCP_PPC300N()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_PPC300N
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub

Private Sub Phase_TCP_GENEXPERT()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                        Call SendWSckData(ACK)
                End Select
            Case 2
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendWSckData(ACK)
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
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendWSckData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        intPhase = 1
                        Call TCPRcvData_GENEXPERT
                        
                End Select
        End Select
    Next i

End Sub


Private Sub Phase_TCP_VISION()
    Dim Buffer      As Variant
    Dim BufChar     As String
    'Dim lngBufLen   As Long
    Dim i           As Long
    
    Dim strBuffer   As String
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strRcvCnt   As String
    Dim strSendAck  As String
    
    Dim strNS       As String
    Dim strNE       As String
    Dim intNS       As Integer
    Dim intNE       As Integer
    
    Dim strSendData As String
    
    strRecvData = Split(pBuffer, vbLf)
    
    For i = 0 To UBound(strRecvData)
        strBuffer = strRecvData(i)
        If strBuffer = "" Then
            Exit For
        End If
        strLastSeq = mGetP(strBuffer, 1, vbTab)
        strRcvSign = mGetP(strBuffer, 2, vbTab)
        strSendAck = strLastSeq & vbTab & "ACK"
        
        Select Case UCase(strRcvSign)
            Case "RESULT"
                '2   RESULT  1   VC0111  2015-11-03T06:55:19Z    3   3   23.3    21  17  23.5625 24.8125 False   False
                '3   RESULT  2   VC0111  2015-11-03T06:55:19Z    4   4   24.0    96  84  23.5625 24.8125 False   False
                
                'RcvBuffer = strBuffer
                
                Call TCPRcvData_VISION
                strBuffer = ""
            
            Case "CONNECT"
                strSendData = strSendAck & vbLf
                
                wSck.SendData strSendData
                SetRawData "[Tx]" & strSendData
            
            Case "RESULTS"
                '결과요청
                strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))
                
                strNS = strRcvCnt
                strNE = mGetP(strBuffer, 4, vbTab)
                
                strNS = strNS - strNE
                strNE = strNS + strNE
                
                strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE & vbLf
                
                wSck.SendData strSendData
                SetRawData "[Tx]" & strSendData
                
                'Call WritePrivateProfileString("config", "LASTSEQ", strRcvCnt, App.PATH & "\Interface.ini")
                txtLastSeq.Text = strRcvCnt
                
                'blnResults = False
        End Select
    Next i


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


    'Me.Caption = gHOSP.MACHNM
    'Me.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"

    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
        Me.TOP = gForm.TOP
        Me.LEFT = gForm.LEFT
        Me.WIDTH = gForm.WIDTH
        Me.HEIGHT = gForm.HEIGHT
    End If
    
    Me.Caption = gHOSP.HOSPNM & Space$(5) & gHOSP.MACHNM
    lblHospInfo(0).Caption = "◈ " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " 인터페이스"
    lblHospInfo(1).Caption = "◈ " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " 인터페이스"
    
    strIFStatus = ""
    If gHOSP.BARUSE = "Y" Then
        strIFStatus = strIFStatus & "▣ 바코드사용"
    Else
        If gHOSP.RSTTYPE = "1" Then
            strIFStatus = strIFStatus & "▣ 순번 맞춤"
        ElseIf gHOSP.RSTTYPE = "2" Then
            strIFStatus = strIFStatus & "▣ R/P 맞춤"
        ElseIf gHOSP.RSTTYPE = "3" Then
            strIFStatus = strIFStatus & "▣ 체크순"
        End If
    End If
    strIFStatus = strIFStatus & IIf(gHOSP.SAVELIS = "Y", "  ▣ LIS결과", "  ▣ 장비결과")
    strIFStatus = strIFStatus & IIf(gHOSP.SAVEAUTO = "Y", "  ▣ 자동전송", "  ▣ 수동전송")
    
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

    '-- 컬럼헤더설정
'    Call SetColumnHeader(spdOrder)

    '-- 컬럼보이기설정
    Call SetColumnView(spdOrder)

    '-- 검사코드
    Call GetTestList

    Call GetTestListName

    '-- 검사명 보이기
    Call SetExamCode(spdOrder)

    '-- 통신열기
    Call OpenCommunication

    pDel = False

'    spdComStatus.MaxRows = 0
'    spdComStatus.Font.Bold = True
    
    lstComStatus.Clear
'    lstComStatus.FontBold = True
    
    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    txtTestID.Text = gHOSP.USERID
    txtTestNm.Text = gHOSP.USERNM
    lblPatInfo.Caption = ""

    imgNet1.ZOrder 0
    tmrDBConn.Interval = 500
    tmrDBConn.Enabled = True
    
    
    '-- 이전결과 삭제
    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gHOSP.SAVEDAY + "일전 데이타를 삭제하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    
    If gHOSP.MACHNM = "VISION" Then
        fraVision.Visible = True
    Else
        fraVision.Visible = False
    End If
    
    If gHOSP.MACHNM = "MICROPLATE" Then
        fraMicroPlate.Visible = True
    Else
        fraMicroPlate.Visible = False
    End If

'    If gHOSP.MACHNM = "ACCESS2" Then
'        fraAccess.Visible = True
'    Else
'        fraAccess.Visible = False
'    End If

    If gHOSP.DBCONCHK = "Y" Then
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
        strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If

End Sub
 
'
Public Sub OpenCommunication()

    If gComm.COMTYPE = "1" Then

        comEqp.CommPort = gComm.COMPORT
        comEqp.RTSEnable = gComm.RTSEnable
        comEqp.DTREnable = gComm.DTREnable
        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

        If comEqp.PortOpen = False Then
            comEqp.PortOpen = True
        End If

        If comEqp.PortOpen Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
            
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgOn.ZOrder 0
           ' imgCom.Picture = imlStatus.ListImages("ON").ExtractIcon

        
        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgOff.ZOrder 0
        
           ' imgCom.Picture = imlStatus.ListImages("OFF").ExtractIcon
        
        End If
        
    ElseIf gComm.COMTYPE = "2" Then
        'lblComStatus.Left = imgPort.Left + 500
        'lblComStatus.Width = 6000
        If gComm.TCPTYPE = "SERVER" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen
                
            
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 연결중.."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgSend.Visible = False
            'imgReceive.Visible = False
            'lblSend.Visible = False
            'lblRcv.Visible = False
            imgOff.ZOrder 0

        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
            
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 연결중..."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgSend.Visible = False
            'imgReceive.Visible = False
            'lblSend.Visible = False
            'lblRcv.Visible = False
            imgOff.ZOrder 0
        
        End If
    ElseIf gComm.COMTYPE = "" Then

    End If

End Sub


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
        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        If Me.TOP < 0 Then
            Me.TOP = 0
        End If
        'gForm.MAXYN = False
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call WritePrivateProfileString("FORM", "MAXYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "TOP", gForm.TOP, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "LEFT", gForm.LEFT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "WIDTH", gForm.WIDTH, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "HEIGHT", gForm.HEIGHT, App.PATH & "\INI\" & gMACH & ".ini")

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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
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



Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

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

Private Sub mnuEMRInfo_Click()
    
    If InputBox("비밀번호 입력" & Space(5) & "hint:개발자oyh") = "dev0503" Then
        frmEMRInfo.Show
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

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

Private Sub mnuHosp_Click()

    frmHospInfo.Show 'vbModal

End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuOpt_Click()
    
    frmTestOptSet.Show vbModal
    
End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show 'vbModal
    
End Sub

Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    
End Sub

Private Sub mnuTest_Click()
    
    frmTestSet.Show 'vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show 'vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show 'vbModal

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

Private Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    Dim intRow      As Integer
    Dim strSeq      As String
    
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdOrder, sRow, sCol)
    
    If KeyCode = vbKeyReturn Then
        If colBARCODE = sCol Then
            If GetSampleInfo(sRow, spdOrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdOrder, strSeq + 1, intRow, colSEQNO)
                    strSeq = strSeq + 1
                Next
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If strNewBarNo = "" Then
        
        End If
        
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdOrder, sRow, sRow
        spdOrder.MaxRows = spdOrder.MaxRows - 1
        spdResult.MaxRows = 0
    ElseIf KeyCode = vbKeyDown Then
        DoEvents
        If sRow = spdOrder.MaxRows Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow + 1)
        DoEvents
    ElseIf KeyCode = vbKeyUp Then
        DoEvents
        If sRow = 1 Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow - 1)
        DoEvents
        
    End If
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
            Call SetSpreadSort(spdWork, 0)
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
            Select Case gHOSP.MACHNM
                Case "ACCESS2"
                    Call SetTag(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                    'Call SetToolTip(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                Case "PPC300N"
                    Call SetTag(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
                    'Call SetToolTip(spdOrder, GetTag(spdWork, Row, colSTATE), intRow, colSTATE)
            End Select
            
            Call SetText(spdWork, GetSampleITEM(Row, spdWork), Row, colITEMS)
            
            DoEvents
            
            
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
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


Private Sub txtBarNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow As Integer
    
    If KeyCode = vbKeyReturn Then
        If chkAdd.Value = "1" Then
            With spdOrder
                sRow = lblRow.Caption
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                'Call spdOrder_KeyDown(13, 1)
                
                If GetSampleInfo(sRow, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                Else
                    '정보수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                    SQL = SQL & " ,HOSPDATE = '" & Trim(GetText(spdOrder, sRow, colHOSPDATE)) & "'" & vbCrLf
                    SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                    SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                    SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                    SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                    SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                    SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                    SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                    SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                    SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                    SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                    SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                    SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
            End With
        Else
            With spdOrder
                .MaxRows = .MaxRows + 1
                sRow = .MaxRows
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                
                If GetSampleInfo(.Row, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                Else
                    '정보수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                    SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                    SQL = SQL & " ,HOSPDATE = '" & Trim(GetText(spdOrder, sRow, colHOSPDATE)) & "'" & vbCrLf
                    SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                    SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                    SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                    SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                    SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                    SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                    SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                    SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                    SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                    SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                    SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                    lblRow.Caption = lblRow.Caption + 1
                End If
                
                Call spdActiveCell(spdOrder, .Row + 1, colBARCODE)
                
            End With
        End If
        
        txtBarNum.SelStart = 0
        txtBarNum.SelLength = Len(txtBarNum.Text)
    
    End If

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

Private Sub wSCK_Close()
        
    If gComm.TCPTYPE = "SERVER" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    End If

End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "SERVER" Then
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 포트 연결성공"
            imgOn.ZOrder 0
        Else
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트 연결성공"
            imgOn.ZOrder 0
        End If
    End If
            
End Sub

Private Sub wSCK_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    wSck.GetData strText
    SetRawData "[Rx]" & strText
    
    '-- 컴파일시 제외할 것!!
'    strText = Replace(strText, vbLf, "")
    pBuffer = strText
    
    If Len(pBuffer) > 0 Then
    
        Select Case UCase(gHOSP.MACHNM)
            Case "F200":            Call Phase_TCP_F200
            
            Case "GENEXPERT":       Call Phase_TCP_GENEXPERT
            
            Case "PPC300N":         Call Phase_TCP_PPC300N
            Case "VISION":          Call Phase_TCP_VISION
            Case "KLITE":           Call Phase_TCP_KLITE
            Case "XP300":           Call Phase_TCP_XP300
            Case "YUMIZEN":         Call Phase_TCP_YUMIZEN
        
        End Select
    End If

End Sub

