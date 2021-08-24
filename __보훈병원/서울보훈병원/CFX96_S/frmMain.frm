VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   -1530
   ClientWidth     =   21900
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
   ScaleHeight     =   11880
   ScaleWidth      =   21900
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   6345
      Left            =   8460
      TabIndex        =   14
      Top             =   4620
      Visible         =   0   'False
      Width           =   11325
      Begin VB.CommandButton cmdOrder 
         Caption         =   "오더전송"
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
         Left            =   1050
         TabIndex        =   67
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   4710
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.OptionButton optTest 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         Caption         =   "RP+PB"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   3
         Left            =   9390
         TabIndex        =   65
         Top             =   330
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton optTest 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         Caption         =   "PB6"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   2
         Left            =   8520
         TabIndex        =   64
         Top             =   330
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdTestNmSave 
         Caption         =   "저장"
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
         Left            =   3270
         TabIndex        =   58
         Top             =   3600
         Width           =   555
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1530
         TabIndex        =   54
         Text            =   "12121334133"
         Top             =   2790
         Width           =   1845
      End
      Begin VB.CheckBox chkAdd 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         Caption         =   "O"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   960
         TabIndex        =   53
         Top             =   2790
         Value           =   1  '확인
         Width           =   495
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "워크조회"
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResultSearch 
         Caption         =   "결과조회"
         Height          =   315
         Left            =   2610
         TabIndex        =   15
         Top             =   270
         Width           =   1425
      End
      Begin FPSpread.vaSpread spdCodeName 
         Height          =   4125
         Left            =   5100
         TabIndex        =   52
         Top             =   870
         Width           =   4245
         _Version        =   393216
         _ExtentX        =   7488
         _ExtentY        =   7276
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS PGothic"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmMain.frx":0E42
      End
      Begin FPSpread.vaSpread spdComStatus 
         Height          =   330
         Left            =   2700
         TabIndex        =   62
         Top             =   5160
         Visible         =   0   'False
         Width           =   3570
         _Version        =   393216
         _ExtentX        =   6297
         _ExtentY        =   582
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   3
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   2
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":10D9
         UserResize      =   0
         TextTip         =   2
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   990
         Top             =   4650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "검사자명 : "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1470
         TabIndex        =   59
         Top             =   3630
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   1380
         Top             =   3540
         Width           =   2505
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
         TabIndex        =   20
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
   Begin VB.CommandButton cmdSL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "▶"
      Height          =   345
      Left            =   90
      Style           =   1  '그래픽
      TabIndex        =   51
      Top             =   660
      Width           =   465
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   17505
      _Version        =   393216
      _ExtentX        =   30877
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
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":151F
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      TabIndex        =   41
      Top             =   660
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   5310
         Style           =   1  '그래픽
         TabIndex        =   47
         Top             =   210
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "워크조회"
         Height          =   375
         Left            =   4200
         Style           =   1  '그래픽
         TabIndex        =   46
         Top             =   210
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   42
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Format          =   132972545
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   43
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Format          =   132972545
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
         Left            =   2580
         TabIndex        =   45
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "접수일자 :"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   44
         Top             =   330
         Width           =   930
      End
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '아래 맞춤
      Height          =   2985
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   21840
      TabIndex        =   29
      Top             =   8310
      Visible         =   0   'False
      Width           =   21900
      Begin FPSpread.vaSpread spdExcel 
         Height          =   2835
         Left            =   60
         TabIndex        =   50
         Top             =   60
         Width           =   22245
         _Version        =   393216
         _ExtentX        =   39238
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
         SpreadDesigner  =   "frmMain.frx":348F
      End
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "C"
         Height          =   495
         Left            =   12930
         TabIndex        =   39
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   38
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   37
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   36
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   35
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   34
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         Height          =   555
         Left            =   13560
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   30
         Width           =   3435
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   525
         Left            =   17010
         TabIndex        =   32
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "Rcv"
         Height          =   525
         Left            =   11910
         TabIndex        =   30
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   17610
      TabIndex        =   16
      Top             =   660
      Width           =   5355
      Begin VB.TextBox txtSA 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "1004"
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1004"
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox txtPatID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "1004"
         Top             =   210
         Width           =   1485
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "1004"
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "S / A"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2670
         TabIndex        =   28
         Top             =   690
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
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   26
         Top             =   690
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
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2670
         TabIndex        =   24
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   22
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00404040&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   6
      Top             =   11295
      Width           =   21900
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
         Left            =   1770
         Top             =   -30
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
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3726
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3CC0
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":425A
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":47F4
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5086
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":51E0
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":533A
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAPI 
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
         Left            =   9300
         TabIndex        =   63
         Top             =   180
         Visible         =   0   'False
         Width           =   10665
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   16080
         Top             =   90
         Width           =   8175
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3900
         Picture         =   "frmMain.frx":5494
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   4845
         Picture         =   "frmMain.frx":5A1E
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   5760
         Picture         =   "frmMain.frx":5FA8
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   210
         Width           =   360
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
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
         Height          =   195
         Left            =   4335
         TabIndex        =   10
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
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
         Height          =   195
         Left            =   5220
         TabIndex        =   9
         Top             =   210
         Width           =   420
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":6532
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":667C
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":67C6
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
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
         Height          =   315
         Left            =   6240
         TabIndex        =   8
         Top             =   180
         Width           =   7905
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
         TabIndex        =   7
         Top             =   180
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3210
         Top             =   90
         Width           =   12885
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6795
      Left            =   17610
      TabIndex        =   4
      Top             =   1770
      Width           =   5325
      _Version        =   393216
      _ExtentX        =   9393
      _ExtentY        =   11986
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
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":6910
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00800000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   0
      Top             =   0
      Width           =   21900
      Begin VB.CommandButton cmdPrint 
         Caption         =   "화면인쇄"
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
         Left            =   10680
         TabIndex        =   68
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   150
         Width           =   1425
      End
      Begin VB.ComboBox cboUser 
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   66
         Text            =   "Combo1"
         Top             =   150
         Width           =   1905
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   61
         Text            =   "1004"
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "워크조회"
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
         Left            =   12210
         TabIndex        =   60
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   150
         Width           =   1425
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Height          =   495
         Left            =   8700
         TabIndex        =   55
         Top             =   0
         Width           =   1845
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            Caption         =   "COVID-19"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00FF80FF&
            Caption         =   "RP19"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   1
            Left            =   960
            TabIndex        =   56
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "결과받기"
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
         Left            =   13740
         TabIndex        =   48
         ToolTipText     =   "선택한 결과를 EMR서버로 저장합니다"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "화면정리"
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
         Left            =   15270
         TabIndex        =   19
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "선택저장"
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
         Left            =   16800
         TabIndex        =   18
         ToolTipText     =   "선택한 결과를 EMR서버로 저장합니다"
         Top             =   150
         Width           =   1425
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   13
         Text            =   "1004"
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdTestIDSave 
         Caption         =   "저장"
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
         Left            =   5970
         TabIndex        =   12
         Top             =   150
         Width           =   555
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   21060
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSComDlg.CommonDialog CFXFile 
         Left            =   20310
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   10620
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   12150
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   18270
         Top             =   90
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblRow 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   18360
         TabIndex        =   49
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   13680
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   15210
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   16740
         Top             =   90
         Width           =   1545
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
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1380
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2565
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "검사자ID/명 : "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2910
         TabIndex        =   3
         Top             =   180
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   2820
         Top             =   90
         Width           =   5835
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "검사일자 :"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   180
         Width           =   945
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7275
      Left            =   60
      TabIndex        =   40
      Top             =   1410
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
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":760C
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  파일  "
      Begin VB.Menu mnuHosp 
         Caption         =   "▷ 병원 정보"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "▷ EMR 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu04 
      Caption         =   "  조회  "
      Begin VB.Menu mnuResult 
         Caption         =   "▣ 결과 조회"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "▣ 워크 조회"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "  설정  "
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
      Begin VB.Menu mnuResultTrans 
         Caption         =   "▣ 결과 변환"
      End
      Begin VB.Menu mnuSep231 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComment 
         Caption         =   "▣ 코멘트 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep24 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "▣ 옵션 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " 검사옵션 "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "▷ 바코드 사용"
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
      Begin VB.Menu mnuSep14 
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
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " 원격지원 "
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

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Dim OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


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


Private Sub cboUser_Click()

    txtTestID.Text = mGetP(cboUser.Text, 2, "|")
    txtTestNm.Text = mGetP(cboUser.Text, 1, "|")

    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

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
                    Next
                    
                    spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
        '.MaxRows = 0
    End With
    
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

Private Sub cmdOrder_Click()
    Dim intRow      As Long
    Dim intCnt      As Integer
    Dim strBC       As String
    Dim strName     As String
    Dim strTest     As String
    Dim FileNumber
    Dim strOneStr   As String
    Dim strAllStr   As String
    Dim strFileName As String
    Dim intLen      As Integer
    Dim STM         As ADODB.Stream
    
    Dim intRows     As Integer
    Dim intColumn   As Integer
    
    strAllStr = ""
    intCnt = 0
    strFileName = "(" & gHOSP.MACHCD & ") " & gHOSP.MACHNM & " " & "Order_" & Format(Now, "yyyymmddhhmmss") & ".csv"
    
    '(M02) CFX96_CSV Order_20200309140545.csv
    
    'Row,Column,*Target Name,*Sample Name
    'A , 1, , 2020048504
    'B , 1, , 2010071380
    'C , 1, , 2010071379
    'D , 1, , 2020048635
    'E , 1, , 2010071327
    'F , 1, , 2010071326
    'G , 1, , 2010071329
    'H , 1, , 2010071328
    'A , 2, , 2020048766
    
    intRows = 64
    intColumn = 1
    
    With spdOrder
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                strBC = GetText(spdOrder, intRow, colBARCODE)
                'strBC = "12112"
                If strBC <> "" Then
                    
                    If Chr(intRows) = "I" Then
                        intRows = 65
                        intColumn = intColumn + 1
                    Else
                        intRows = intRows + 1
                        If intRows = 73 Then
                            intRows = 65
                            intColumn = intColumn + 1
                        End If
                    End If
                    

                    strOneStr = ""
                    strOneStr = strOneStr & Chr(intRows) & "," & CStr(intColumn) & ",," & strBC & vbCrLf
                        
                    strAllStr = strAllStr & strOneStr
                    
                    '-- 진행상태(Order) 표시
                    Call SetText(spdOrder, "오더전송", intRow, colSTATE)
                    Call SetText(spdOrder, "0", intRow, colCHECKBOX)
                
                End If
            End If
        Next
        
        strAllStr = "Row,Column,*Target Name,*Sample Name" & vbCrLf & strAllStr

        
'        'RP
'        ElseIf optTest(1).Value = True Then
'            For intRow = 1 To .MaxRows
'                .Row = intRow
'                .Col = colCHECKBOX
'                If .Value = "1" Then
'                    strBC = GetText(spdOrder, intRow, colBARCODE)
'                    strName = GetText(spdOrder, intRow, colPNAME)
'                    If strBC <> "" Then
'                        intCnt = intCnt + 1
'                        strOneStr = ""
'                        strOneStr = strOneStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) '5
'                        strOneStr = strOneStr & strBC & Space(20 - Len(strBC))              '20
'                        strOneStr = strOneStr & "RP1" & Space(12)                           '15
'                        strOneStr = strOneStr & "RP2" & Space(12)                           '15
'                        strOneStr = strOneStr & "RP3" & Space(12)                           '15
'                        strOneStr = strOneStr & Space(90)                                   '90
'                        strOneStr = strOneStr & strName & Space(15 - LenB(strName) - Len(strName))      '15
'
'                        'MsgBox LenB(strOneStr)
'                        strAllStr = strAllStr & strOneStr
'
'                        '-- 진행상태(Order) 표시
'                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
'                        Call SetText(spdOrder, "0", intRow, colCHECKBOX)
'
'                        'strFileName = "nimbus_rb6.lis"
'
'                    End If
'                End If
'            Next
'        'PB
'        ElseIf optTest(2).Value = True Then
'            For intRow = 1 To .MaxRows
'                .Row = intRow
'                .Col = colCHECKBOX
'                If .Value = "1" Then
'                    strBC = GetText(spdOrder, intRow, colBARCODE)
'                    strName = GetText(spdOrder, intRow, colPNAME)
'                    If strBC <> "" Then
'                        intCnt = intCnt + 1
'                        strOneStr = ""
'                        strOneStr = strOneStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) '5
'                        strOneStr = strOneStr & strBC & Space(20 - Len(strBC))              '20
'                        strOneStr = strOneStr & "PB" & Space(13)                           '15
'                        strOneStr = strOneStr & Space(15)                           '15
'                        strOneStr = strOneStr & Space(15)                           '15
'                        strOneStr = strOneStr & Space(90)                                   '90
'                        strOneStr = strOneStr & strName & Space(15 - LenB(strName) - Len(strName))      '15
'
'                        'MsgBox LenB(strOneStr)
'                        strAllStr = strAllStr & strOneStr
'
'                        '-- 진행상태(Order) 표시
'                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
'                        Call SetText(spdOrder, "0", intRow, colCHECKBOX)
'
'                        'strFileName = "nimbus_rb6.lis"
'
'                    End If
'                End If
'            Next
'        'RP + PB
'        ElseIf optTest(3).Value = True Then
'            For intRow = 1 To .MaxRows
'                .Row = intRow
'                .Col = colCHECKBOX
'                If .Value = "1" Then
'                    strBC = GetText(spdOrder, intRow, colBARCODE)
'                    strName = GetText(spdOrder, intRow, colPNAME)
'                    '2020.02.14 수정 : 이름이 길경우 오류 발생
'                    strName = Mid(strName, 1, 5)
'                    strTest = GetText(spdOrder, intRow, colPOSNO)
'                    If strBC <> "" Then
'                        intCnt = intCnt + 1
'                        If strTest = "RP" Then
'                            strOneStr = ""
'                            strOneStr = strOneStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) '5
'                            strOneStr = strOneStr & strBC & Space(20 - Len(strBC))              '20
'                            strOneStr = strOneStr & "RP1" & Space(12)                           '15
'                            strOneStr = strOneStr & "RP2" & Space(12)                           '15
'                            strOneStr = strOneStr & "RP3" & Space(12)                           '15
'                            strOneStr = strOneStr & Space(90)                                   '90
'                            strOneStr = strOneStr & strName & Space(15 - LenB(strName) - Len(strName))      '15
'
'                        ElseIf strTest = "PB" Then
'                            strOneStr = ""
'                            strOneStr = strOneStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) '5
'                            strOneStr = strOneStr & strBC & Space(20 - Len(strBC))              '20
'                            strOneStr = strOneStr & "PB" & Space(13)                           '15
'                            strOneStr = strOneStr & Space(15)                           '15
'                            strOneStr = strOneStr & Space(15)                           '15
'                            strOneStr = strOneStr & Space(90)                                   '90
'                            strOneStr = strOneStr & strName & Space(15 - LenB(strName) - Len(strName))      '15
'
'                        End If
'
'                        strAllStr = strAllStr & strOneStr
'
'                        '-- 진행상태(Order) 표시
'                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
'                        Call SetText(spdOrder, "0", intRow, colCHECKBOX)
'
'                    End If
'                End If
'            Next
'        End If
    End With

    If strAllStr <> "" Then
        'nimbus.lis 파일 만들기 ======================================================
        '-- OPEN
'        FileNumber = FreeFile
'        Open gComm.ORDPATH & "\" & strFileName For Output As FileNumber
'        Close FileNumber

        If Dir$(gComm.ORDPATH & "\" & strFileName, vbNormal) <> "" Then
            Kill gComm.ORDPATH & "\" & strFileName
        End If
                    
        '## 파일오픈
        Set STM = New ADODB.Stream
        
        STM.Open
        STM.Type = adTypeText
        STM.Charset = "utf-8"
        'strAllStr = Mid(strAllStr, 1, Len(strAllStr) - 4)
        
        STM.WriteText strAllStr
        STM.SaveToFile gComm.ORDPATH & "\" & strFileName, adSaveCreateNotExist
        STM.Close
        Set STM = Nothing
        
'        '-- WRITE
'        Open gComm.ORDPATH & "\" & strFileName For Append As FileNumber
'        Print #FileNumber, strAllStr;
'        Close FileNumber
        '=========================================================================
        MsgBox "오더 파일 생성 완료", vbOKOnly + vbInformation, Me.Caption
    Else
        MsgBox "전송할 오더가 없습니다", vbOKOnly + vbCritical, Me.Caption
    End If


End Sub

Private Sub cmdPrint_Click()
    
    spdOrder.PrintOrientation = PrintOrientationLandscape
    spdOrder.Action = ActionPrint

    MsgBox "출력 완료", vbOKOnly + vbInformation, Me.Caption
    
End Sub

Private Sub cmdRcv_Click()
        
    pBuffer = txtRcv.Text
    
    Select Case UCase(gHOSP.MACHNM)
                
        Case "AU680":           Call Phase_Serial_AU680
        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                    
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdResult_Click()
    Dim intRow      As Integer
    Dim intIDX      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strBuffer   As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim j           As Long
    Dim intCnt      As Integer
    Dim varTmp      As Variant
    Dim RESULTPATH   As String

    Dim strBarno As String
    Dim strOldBarno As String
    Dim strNewBarno As String

On Error GoTo ErrRoutine

'    Call Excel_Open
    
    Call getCFXExlData(0)

    Call Phase_File_CFX96
    
Exit Sub

ErrRoutine:

End Sub



Private Function ShowOpen(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
    
End Function


Private Sub getCFXExlData(ByVal intIDX As Integer)

    Dim x As Integer, Y As Boolean, z As Boolean
    Dim ListCount   As Integer, handle As Integer
    Dim List(10)    As String
    Dim intRow, intCol As Long
    Dim varTmp      As Variant
    Dim strExcel    As String
    Dim rv          As Integer
    Dim lRow        As Integer
    Dim lRow1       As Integer
    Dim sExamCode   As String
    Dim sExamName   As String
    Dim sEquipCode  As String
    Dim sItemCode   As String
    Dim strAge      As String
    Dim strSex      As String
    Dim strPtno     As String
    Dim strPtname   As String
    Dim strTmp      As String
    Dim sFile       As String
'    Dim intSheet    As Integer

    Dim idates1$, idates2$, iexamcode$
    Dim PT_NO$(), PATNAME$(), SEX$(), AGE$()
    Dim SPC_NO$(), gnl_item_cd$(), bl_gth_dte$()
    Dim dept$(), WD_NO$(), TST_CD$()
    Dim ispcno$

    Dim k As Integer
    Dim intRow1 As Integer
    
On Error Resume Next

    '-- 이전버전
    If intIDX = 0 Then
        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", gComm.RSTPATH)
        If sFile <> "" Then
            strExcel = sFile
            spdExcel.ScriptEnhanced = True
            x = spdExcel.IsExcelFile(strExcel)
            If x = 1 Then
                Y = spdExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
                If Y = True Then
                    z = spdExcel.ImportExcelSheet(handle, 0)
                    If z = True Then
                        'MsgBox "가져오기 성공"
                    Else
                        'MsgBox "가려오기 실패"
                    End If
                End If
            End If
            
'''            With vasExcel
'''                For intRow = 2 To .DataRowCnt Step 4
'''                    For intCol = 1 To 8
'''                        .GetText intCol, intRow, varTmp
'''
'''
'''                        If varTmp <> "" Then
'''                            Select Case intCol
'''                            Case 1
'''                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
'''                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
'''
'''
'''                                'rv = 1
'''                                If rv >= 1 Then
'''                                    '-- 환자정보
'''                                    vasID.MaxRows = vasID.MaxRows + 1
'''                                    vasID.RowHeight(-1) = 12
'''                                    lRow = vasID.MaxRows
'''                                    sExamName = Get_ExamName(tst_cd(0))
'''                                    sEquipCode = Get_EquipCode(tst_cd(0))
'''                                    sExamCode = tst_cd(0)
'''
'''                                    sItemCode = gnl_item_cd(0)
'''                                    strPtno = patname(0)
'''                                    strPtname = patname(0)
'''                                    strSex = Sex(0)
'''                                    strAge = Age(0)
'''
'''                                    SetText vasID, ispcno$, lRow, 2
'''                                    SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
'''                                    SetText vasID, pt_no(0), lRow, 6           '병록번호
'''                                    SetText vasID, patname(0), lRow, 7         '이름
'''                                    SetText vasID, Sex(0), lRow, 8            '성별
'''                                    SetText vasID, Age(0), lRow, 9            '나이
'''
'''                                    '-- 채널
'''                                    '.GetText 5, intRow, varTmp: sEquipCode = varTmp
'''                                    sEquipCode = "COVID19"
'''                                    'sExamCode = sItemCode 'Get_ExamCode(sEquipCode)
'''                                    'sExamName = Get_ExamName(sExamCode)
'''
'''                                          SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
'''                                    SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
'''                                    Res = db_select_Col(gLocal, SQL)
'''                                    '-- 결과
'''                                    .GetText 8, intRow, varTmp: strTmp = varTmp
''''                                    If strTmp = "-" Then
''''                                        strTmp = "Negative"
''''                                    ElseIf strTmp = "+" Then
''''                                        strTmp = "Positive(1+)"
''''                                    ElseIf strTmp = "++" Then
''''                                        strTmp = "Positive(2+)"
''''                                    ElseIf strTmp = "+++" Then
''''                                        strTmp = "Positive(3+)"
''''                                    ElseIf strTmp = "++++" Then
''''                                        strTmp = "Positive(4+)"
''''                                    'Else
''''                                    '    strTmp = "Positive"
''''                                    End If
'''
'''                                    If Res > 0 Then
'''                                        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'''                                              "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'''                                              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'''                                              "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
'''                                              "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
'''                                              "  and examcode= '" & sExamCode & "'"
'''                                        Res = SendQuery(gLocal, SQL)
'''                                    End If
'''
'''                                    SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
'''                                          "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
'''                                          "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
'''                                          "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
'''                                    Res = SendQuery(gLocal, SQL)
'''    '                                Debug.Print SQL
'''
'''                                    Exit For
''''''                                Else
''''''                                    '-- 환자정보
''''''                                    vasID.maxrows = vasID.maxrows + 1
''''''                                    vasID.RowHeight(-1) = 12
''''''                                    lRow = vasID.maxrows
''''''                                    SetText vasID, CStr(varTmp), lRow, 2
'''                                End If
''''                            Case "8"
'''                                    '-- 채널
''''                                    sEquipCode = "COVID19"
''''                                    sExamCode = Get_ExamCode(sEquipCode)
''''                                    sExamName = Get_ExamName(sExamCode)
''''
''''                                          SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
''''                                    SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
''''                                    res = db_select_Col(gLocal, SQL)
''''                                    '-- 결과
''''                                    .GetText 8, intRow, varTmp: strTmp = varTmp
'''''                                    If strTmp = "-" Then
'''''                                        strTmp = "Negative"
'''''                                    ElseIf strTmp = "+" Then
'''''                                        strTmp = "Positive(1+)"
'''''                                    ElseIf strTmp = "++" Then
'''''                                        strTmp = "Positive(2+)"
'''''                                    ElseIf strTmp = "+++" Then
'''''                                        strTmp = "Positive(3+)"
'''''                                    ElseIf strTmp = "++++" Then
'''''                                        strTmp = "Positive(4+)"
'''''                                    'Else
'''''                                    '    strTmp = "Positive"
'''''                                    End If
''''
''''
''''                                    If res > 0 And sExamCode <> "" Then
''''
''''                                        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
''''                                              "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
''''                                              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''''                                              "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
''''                                              "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
''''                                              "  and examcode= '" & sExamCode & "'"
''''                                        res = SendQuery(gLocal, SQL)
''''                                    End If
''''
''''                                    SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
''''                                          "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
''''                                          "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
''''                                          "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
''''                                    res = SendQuery(gLocal, SQL)
''''    '                                Debug.Print SQL
''''                                    Exit For
'''                            End Select
'''                        End If
'''                    Next
'''                Next
'''            End With
        End If
    
'''    Else
'''    '-- 신버전
'''        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", App.PATH)
'''        If sFile <> "" Then
'''            vasID.MaxRows = 0
'''            strExcel = sFile
'''            vasExcel.ScriptEnhanced = True
'''            x = vasExcel.IsExcelFile(strExcel)
'''            If x = 1 Then
'''                y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
'''                If y = True Then
'''                    '-- 2번째sheet
'''                    z = vasExcel.ImportExcelSheet(handle, 1)
'''                    If z = True Then
'''                        'MsgBox "가져오기 성공"
'''                        With vasExcel
'''                            For intRow = 3 To .DataRowCnt 'Step 6
'''                                .GetText 1, intRow, varTmp
'''                                If varTmp <> "" Then
'''                                    ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
'''                                    rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
'''                                   ' rv = 1
'''                                    If rv >= 1 Then
'''                                        '-- 환자정보
'''                                        vasID.MaxRows = vasID.MaxRows + 1
'''                                        vasID.RowHeight(-1) = 12
'''                                        lRow = vasID.MaxRows
'''                                        sExamName = Get_ExamName(tst_cd(0))
'''                                        sEquipCode = Get_EquipCode(tst_cd(0))
'''                                        sItemCode = gnl_item_cd(0)
'''                                        strPtno = patname(0)
'''                                        strPtname = patname(0)
'''                                        strSex = Sex(0)
'''                                        strAge = Age(0)
'''
'''                                        SetText vasID, ispcno$, lRow, 2
'''                                        SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
'''                                        SetText vasID, pt_no(0), lRow, 6           '병록번호
'''                                        SetText vasID, patname(0), lRow, 7         '이름
'''                                        SetText vasID, Sex(0), lRow, 8            '성별
'''                                        SetText vasID, Age(0), lRow, 9            '나이
'''                                    Else
'''                                        '-- 환자정보
'''                                        vasID.MaxRows = vasID.MaxRows + 1
'''                                        vasID.RowHeight(-1) = 12
'''                                        lRow = vasID.MaxRows
'''                                        SetText vasID, ispcno$, lRow, 2
'''                                    End If
'''
'''                                    intRow1 = intRow
'''
'''                                    For intCol = 6 To 18 Step 2
'''                                        .GetText intCol, intRow1, varTmp
'''
'''                                        'Debug.Print varTmp
'''
'''                                        If varTmp <> "" Then
'''                                            '-- Row = 1,2 채널
'''                                            .GetText intCol, 2, varTmp: sEquipCode = varTmp
'''                                            sExamCode = Get_ExamCode(sEquipCode)
'''                                            sExamName = Get_ExamName(sExamCode)
'''
''''                                                      SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
''''                                                SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
''''                                                res = db_select_Col(gLocal, SQL)
'''                                            '-- 결과
'''                                            .GetText intCol, intRow1, varTmp: strTmp = varTmp
'''
''''                                                Debug.Print varTmp
'''
'''                                            If strTmp = "-" Then
'''                                                strTmp = "Negative"
'''                                            ElseIf strTmp = "+" Then
'''                                                strTmp = "Positive" '"Positive(1+)"
'''                                            ElseIf strTmp = "++" Then
'''                                                strTmp = "Positive" '"Positive(2+)"
'''                                            ElseIf strTmp = "+++" Then
'''                                                strTmp = "Positive" '"Positive(3+)"
'''                                            ElseIf strTmp = "++++" Then
'''                                                strTmp = "Positive" '"Positive(4+)"
'''                                            'Else
'''                                            '    strTmp = "Positive"
'''                                            End If
'''                                            If strTmp <> "" Then
'''                                                If Res > 0 And sExamCode <> "" Then
'''                                                    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'''                                                          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'''                                                          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'''                                                          "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
'''                                                          "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
'''                                                          "  and examcode= '" & sExamCode & "'"
'''                                                    Res = SendQuery(gLocal, SQL)
'''                                                    Debug.Print SQL
'''                                                End If
'''                                                'If strTmp <> "" Then
'''                                                    SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
'''                                                          "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
'''                                                          "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
'''                                                          "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
'''                                                    Res = SendQuery(gLocal, SQL)
'''                                                    Debug.Print SQL
'''                                                'End If
'''                                            End If
'''
'''                                        End If
'''                                    Next
'''
'''                                    Call vasID_Click(2, vasID.MaxRows)
'''
'''                                End If
'''                            Next
'''                        End With
'''                    End If
'''                    '-- 3번째sheet
'''                    z = vasExcel.ImportExcelSheet(handle, 2)
'''                    If z = True Then
'''                        'MsgBox "가져오기 성공"
'''                        With vasExcel
'''                            For intRow = 3 To .DataRowCnt
'''                                .GetText 1, intRow, varTmp
'''                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
'''                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
'''                                'rv = 1
'''                                If rv >= 1 Then
'''                                    '-- 환자정보
'''                                    vasID.MaxRows = vasID.MaxRows + 1
'''                                    vasID.RowHeight(-1) = 12
'''                                    lRow = vasID.MaxRows
'''                                    sExamName = Get_ExamName(tst_cd(0))
'''                                    sEquipCode = Get_EquipCode(tst_cd(0))
'''                                    sItemCode = gnl_item_cd(0)
'''                                    strPtno = patname(0)
'''                                    strPtname = patname(0)
'''                                    strSex = Sex(0)
'''                                    strAge = Age(0)
'''
'''                                    SetText vasID, ispcno$, lRow, 2
'''                                    SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
'''                                    SetText vasID, pt_no(0), lRow, 6           '병록번호
'''                                    SetText vasID, patname(0), lRow, 7         '이름
'''                                    SetText vasID, Sex(0), lRow, 8            '성별
'''                                    SetText vasID, Age(0), lRow, 9            '나이
'''                                End If
'''
'''                                For intCol = 6 To 20 Step 2
'''                                    .GetText intCol, 2, varTmp
'''                                    If varTmp <> "" Then
'''                                        '-- Row = 1,2 채널
'''                                        .GetText intCol, 2, varTmp: sEquipCode = varTmp
'''                                        sExamCode = Get_ExamCode(sEquipCode)
'''                                        sExamName = Get_ExamName(sExamCode)
'''
'''                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
'''                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
'''                                        Res = db_select_Col(gLocal, SQL)
'''                                        '-- 결과
'''                                        .GetText intCol, intRow, varTmp: strTmp = varTmp
'''                                        If strTmp = "-" Then
'''                                            strTmp = "Negative"
'''                                        ElseIf strTmp = "+" Then
'''                                            strTmp = "Positive(1+)"
'''                                        ElseIf strTmp = "++" Then
'''                                            strTmp = "Positive(2+)"
'''                                        ElseIf strTmp = "+++" Then
'''                                            strTmp = "Positive(3+)"
'''                                        ElseIf strTmp = "++++" Then
'''                                            strTmp = "Positive(4+)"
'''                                        'Else
'''                                        '    strTmp = "Positive"
'''                                        End If
'''
'''                                        If Res > 0 Then
'''                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'''                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'''                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'''                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
'''                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
'''                                                  "  and examcode= '" & sExamCode & "'"
'''                                            Res = SendQuery(gLocal, SQL)
'''                                        End If
'''
'''                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
'''                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
'''                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
'''                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
'''                                        Res = SendQuery(gLocal, SQL)
'''
'''
'''                                    End If
'''                                Next
'''                            Next
'''                        End With
'''                    End If
'''
'''                End If
'''            End If
'''        End If
    End If

End Sub


Private Sub Excel_Open()
    Dim xlapp   As New Excel.Application
    Dim XLappWS As Worksheet
    Dim lngSCnt As Long
    Dim lngSColCnt(100) As Long
    Dim dummy       As String
    Dim strRowData  As Variant
    Dim lngRowCnt   As Long
    Dim chk_str     As String
    Dim dummy_max   As Long
    Dim lngTotColCnt   As Long
    Dim lngTotRowCnt   As Long
    Dim i, j, k     As Long
    
    lngTotColCnt = 0
    lngTotRowCnt = 0
    
    
    '엑셀 열기
    With CFXFile
        .InitDir = gComm.RSTPATH
        .Filename = "*.csv"
        '.Filter = "Resource CSV (*.CSV)|*.CSV|All File (*.*)|*.*|"
        .Filter = "Excel(*.csv)|*.csv|Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx"
        .DialogTitle = "CFX96 자료 읽어오기"
        .ShowOpen
    End With
    
    
    If CFXFile.FileTitle = "" Then
        Exit Sub
    End If
    
   ' MsgBox CFXFile.FileTitle


    xlapp.Workbooks.Open (Trim(CFXFile.Filename))
    
    lngSCnt = xlapp.Worksheets.Count
    
    'MsgBox "1:" & lngSCnt
    
    '-- 전체 워크시트 불러오기와서 '임시.txt' 파일로 저장
    For i = 1 To lngSCnt
        Set XLappWS = xlapp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlapp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        xlapp.ActiveWorkbook.SaveAs "C:\IF_CFX96\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
        
        
        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>엑셀 2000용
        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>엑셀 2003용
    Next i
    
    xlapp.Quit
    Set XLappWS = Nothing
    Set xlapp = Nothing
    
    'MsgBox "2"
    
    '-- 전체 엑셀의 MAX cols값 추출
    dummy_max = 0
    For i = 1 To lngSCnt
        If lngSColCnt(i) >= dummy_max Then
            dummy_max = lngSColCnt(i)
        End If
    Next i
    lngTotColCnt = dummy_max
    
    '전체 row값 추출
    For i = 1 To lngSCnt
'''        Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\IF_CFX96\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(Trim(dummy)) > 0 Then
                lngTotRowCnt = lngTotRowCnt + 1
            End If
        Wend
        Close #1
    Next i
    
    '-- 그리드 초기화
    spdExcel.MaxRows = 0
    spdExcel.MaxRows = lngTotRowCnt
    spdExcel.MaxCols = lngTotColCnt
    
    '-- 그리드에 출력
    For i = 1 To lngSCnt
        '''Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\IF_CFX96\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(chk_str) > 0 Then
                lngRowCnt = lngRowCnt + 1
                For j = 0 To UBound(strRowData)
                    Call spdExcel.SetText(j + 1, lngRowCnt, CStr(strRowData(j)))
                Next j
            End If
        Wend
        Close #1
    Next i

'    Call SpreadSheetSort(vasExcel, 6, 2)
    With spdExcel
        .Col = 1: .Col2 = .MaxCols
        .Row = 2: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = 6       '정렬키 열번호
        .SortKey(2) = 2       '정렬키 열번호

        '.SortKeyOrder(1) = SortKeyOrderAscending
        '.SortKeyOrder(2) = SortKeyOrderAscending

        .Action = ActionSort
    End With


'Dim SortKeys, SortKeyOrder As Variant
'
'    SortKeys = Array(6, 2)
'    SortKeyOrder = Array(6, 2)
'    ' Sort data in first five columns and rows by column 1 and 3
'    vasExcel.Sort 6, 2, 2, vasExcel.MaxRows, SS_SORT_BY_ROW, SortKeys, SortKeyOrder

End Sub

Private Sub cmdResultSearch_Click()

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
                .Col = 1
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
                    spdOrder.Col = 1
                    spdOrder.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)

End Sub

Private Sub cmdSend_Click()
    
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdSL_Click()
    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        'vasID.Width = 18285 '18075 '15225
        spdOrder.Width = Me.ScaleWidth - 200
        spdOrder.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - 300
        spdResult.Visible = False
        fraPatInfo.Visible = False
    Else
        cmdSL.Caption = "▶"
        spdOrder.Width = Me.ScaleWidth - spdResult.Width - 200
        spdOrder.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - 300
        spdResult.Visible = True
        fraPatInfo.Visible = True
    End If

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub

Private Sub cmdTestIDSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdTestNmSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdWork_Click()

    frmWorkList.Show vbModal
    
End Sub

Private Sub cmdWorkList_Click()
    
    frmWorkList.Show ' vbModal
    
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
                        Case "AU680":           Call Phase_Serial_AU680
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                            
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

Private Sub GetOrder_AU680(ByVal pBarNo As String, ByVal pType As String)

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

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
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
        strItems = GetEquipExamCode_AU680(gHOSP.MACHCD, pBarNo, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & ETX

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 장비로 전송
        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

        '-- 현재 Row
        gRow = intRow

    End With

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

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
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
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        If gHOSP.BARUSE = "N" Then
            mOrder.Function = Replace(mOrder.Function, String(gHOSP.BARLEN, "#"), Left(mOrder.BarNo & Space(gHOSP.BARLEN), gHOSP.BARLEN))
        End If
        
        '-- 검사채널로 장비오더 만들기
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

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
    
    '-- 로그기록
    Call SetRawData("[Tx]" & pSendData)

End Sub

Private Sub SerialRcvData_AU680()
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
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")

            strType = Mid$(strRcvBuf, 1, 2)

            Select Case strType
                Case "R "
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 9, 5)
                    strBarno = Trim(Mid(strRcvBuf, 14, gHOSP.BARLEN))
                    '-- 오더정보
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- 오더환자정보
                    Call GetOrder_AU680(Trim$(strBarno), gHOSP.RSTTYPE)

                Case "D "    '## Result
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 10, 4)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN))
                    
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
                        Exit Sub
                    End If
                    
                    If mResult.BarNo = "" Then
                        Exit Sub
                    End If

                    strTmp = Mid$(strRcvBuf, gHOSP.BARLEN + 19)
                    
                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                        
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                        
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- 검사마스터 정보 가져오기
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- 결과채널이 같고...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
                                    '-- 환자 처방정보 가져오기
'                                    If UBound(gPatTest) > 0 Then
'                                        For intOrdCnt = 1 To UBound(gPatTest)
'                                            For intTestCdCnt = 1 To UBound(gArrEQP)
'                                                '-- 검사코드도 같다면...
'                                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'                                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'                                                    strTestName = gArrEQP(intTestCdCnt, 5)
'                                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'                                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'                                                    '-- 남자참고치를 기본으로 한다
'                                                    strLow = gArrEQP(intTestCdCnt, 9)
'                                                    strHigh = gArrEQP(intTestCdCnt, 10)
'
'                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
'                                                        strCheck = "1"
'
'                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
'                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
'
'                                                        If mPatient.SEX = "M" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 9)
'                                                            strHigh = gArrEQP(intTestCdCnt, 10)
'                                                        ElseIf mPatient.SEX = "F" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 11)
'                                                            strHigh = gArrEQP(intTestCdCnt, 12)
'                                                        Else
'                                                            strLow = ""
'                                                            strHigh = ""
'                                                        End If
'                                                        strState = "R"
'                                                        blnSame = True
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            Next
'                                        Next
'                                    End If

                                    For intTestCdCnt = 1 To UBound(gArrEQP)
                                        '-- 검사코드도 같다면...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            '-- 남자참고치를 기본으로 한다
                                            strLow = gArrEQP(intTestCdCnt, 9)
                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                            
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        strCheck = "1"
                                                        
                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                        
                                                        If mPatient.SEX = "M" Then
                                                            strLow = gArrEQP(intTestCdCnt, 9)
                                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                                        ElseIf mPatient.SEX = "F" Then
                                                            strLow = gArrEQP(intTestCdCnt, 11)
                                                            strHigh = gArrEQP(intTestCdCnt, 12)
                                                        Else
                                                            strLow = ""
                                                            strHigh = ""
                                                        End If
                                                        strState = "R"
                                                        blnSame = True
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                    
                                    If blnSame = True Then
                                        Exit For
                                    End If
                                End If
                            Next

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
                                    Exit For
                                End If
                            Next

                            '-- 결과 List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '순번
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
                            '-- 처방이 있을때만 검사코드를 저장한다.
                            If strState = "R" Then
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
                                SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '검사코드
                                SetText .spdResult, "", intRstRow, colRSUBCD                    '검사코드SUB
                            End If
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
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop

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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub FileRcvData_CFX96_RV6()
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
    
    Dim strTarget       As String
    Dim strVarTest      As String
    Dim strLGrp         As String
    Dim strHGrp         As String
    Dim strTotFlag      As String
    
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    Dim intR            As Integer
    Dim intC            As Integer
    Dim strHospSite     As String
    Dim strHospGbn      As String

    Dim strRstRow       As String
    Dim strAbbrName     As String
    
On Error GoTo ErrHandle
    
    With frmMain
        For intR = 3 To spdExcel.MaxRows
            For intC = 1 To .spdExcel.MaxCols
                strIntBase = ""
                strResult = ""
                If intC = 4 Then
                    strBarno = GetText(.spdExcel, intR, intC)
                    strRackNo = GetText(.spdExcel, intR, intC - 1)
                    
                    If strBarno = "" Then
                        strBarno = strRackNo
                    End If
                    
                    If strOldBarno <> strBarno Then
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RackNo = strRackNo
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            '.SITE = strHospSite
                            'gHOSP.SITE = strHospSite
                        End With
                                    
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                ElseIf intC = 6 Or intC = 8 Or intC = 10 Or intC = 12 Or intC = 14 Or intC = 16 Or intC = 18 Then
                    strIntBase = GetText(.spdExcel, 2, intC)
                    strResult = GetText(.spdExcel, intR, intC)
                    
                    '-- 2019-12-02 수정 : 결과가 양성일 경우 Ct값을 같이 보여준다 ex) Positive(20.5)
                    strIntResult = GetText(.spdExcel, intR, intC + 1)
                    
                    If strResult = "-" Then
                        strResult = gCmnt.PB6NEG
                    ElseIf strResult = "+" Then
                        strResult = gCmnt.PB6POS
                        strResult = strResult & "(" & strIntResult & ")"
                    ElseIf strResult = "++" Then
                        strResult = gCmnt.PB6POS
                        strResult = strResult & "(" & strIntResult & ")"
                    ElseIf strResult = "+++" Then
                        strResult = gCmnt.PB6POS
                        strResult = strResult & "(" & strIntResult & ")"
                    ElseIf strResult = "++++" Then
                        strResult = gCmnt.PB6POS
                        strResult = strResult & "(" & strIntResult & ")"
                    ElseIf strResult = "+++++" Then
                        strResult = gCmnt.PB6POS
                        strResult = strResult & "(" & strIntResult & ")"
                    Else
                        strResult = ""
                    End If
                    
                    If strIntBase <> "" And strResult <> "" And strIntBase <> "Result" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
                        SQL = SQL & "      ,ABBRNAME "
                        SQL = SQL & "  FROM EQPMASTER                                                   " & vbCrLf
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                        " & vbCrLf
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                          " & vbCrLf
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")                             " & vbCrLf
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strAbbrName = Trim(RS_L.Fields("ABBRNAME"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("RESPRECUSE") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                'If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                If strAbbrName = Trim(gArrEQPNm(intCol - colSTATE, 6)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    If InStr(strResult, "Pos") > 0 Then
                                        .spdOrder.Row = gRow
                                        .spdOrder.Col = intCol
                                        .spdOrder.ForeColor = vbRed
                                        .spdOrder.FontSize = 10
                                        .spdOrder.FontBold = True
                                    ElseIf InStr(strResult, "Neg") > 0 Then
                                        .spdOrder.Row = gRow
                                        .spdOrder.Col = intCol
                                        .spdOrder.ForeColor = vbBlue
                                        .spdOrder.FontSize = 10
                                        .spdOrder.FontBold = True
                                    Else
                                        .spdOrder.Row = gRow
                                        .spdOrder.Col = intCol
                                        .spdOrder.ForeColor = vbBlack
                                        .spdOrder.FontSize = 10
                                        .spdOrder.FontBold = False
                                    End If
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS결과
                            
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
                            
                            'SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            'SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '참고치
                            
                            If InStr(UCase(strResult), "POS") > 0 Then
                                .spdResult.Row = strRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontSize = 10
                                .spdResult.FontBold = True
                            ElseIf InStr(UCase(strResult), "NEG") > 0 Then
                                .spdResult.Row = strRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontSize = 10
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = strRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontSize = 10
                                .spdResult.FontBold = False
                            End If
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        
                        '처방이 없으면
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
                            SQL = SQL & "      , ABBRNAME "
                            SQL = SQL & "  FROM EQPMASTER                                                   " & vbCrLf
                            SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                        " & vbCrLf
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                          " & vbCrLf
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strAbbrName = Trim(RS_L.Fields("ABBRNAME"))
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqno = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("RESPRECUSE") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    'If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    If strAbbrName = Trim(gArrEQPNm(intCol - colSTATE, 6)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        If InStr(strResult, "Pos") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbRed
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        ElseIf InStr(strResult, "Neg") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlue
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        Else
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlack
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = False
                                        End If
                                        
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS결과
                                
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
                                
                                'SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                'SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '참고치
                                
                                If InStr(UCase(strResult), "POS") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbRed
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                ElseIf InStr(UCase(strResult), "NEG") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlue
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                Else
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlack
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = False
                                End If
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                If strState = "R" Then
                                    strState = "R"
                                Else
                                    strState = ""
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        End If
                    End If
                
                End If
                
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
            Next
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "FileRcvData_CFX96" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub FileRcvData_CFX96_RP19()
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
    Dim strAbbrName     As String   '검사약어
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
    
    Dim strTarget       As String
    Dim strVarTest      As String
    Dim strLGrp         As String
    Dim strHGrp         As String
    Dim strTotFlag      As String
    
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    Dim intR            As Integer
    Dim intC            As Integer
    Dim strHospSite     As String
    Dim strHospGbn      As String

    Dim strRstRow       As String

On Error GoTo ErrHandle
    
    With frmMain
        For intR = 2 To spdExcel.MaxRows Step 2
            For intC = 3 To .spdExcel.MaxCols
                strIntBase = ""
                strResult = ""
                If intC = 3 Then
                    strRackNo = GetText(.spdExcel, intR, intC)
                    'mResult.RackNo = strRackNo
                    
                ElseIf intC = 4 Then
                    strBarno = GetText(.spdExcel, intR, intC)
                    If strBarno = "" Then
                        strBarno = strRackNo
                    End If
                    
                    If strOldBarno <> strBarno Then
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RackNo = strRackNo
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            '.SITE = strHospSite
                            'gHOSP.SITE = strHospSite
                        End With
                                    
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                'ElseIf intC = 8 Or intC = 12 Or intC = 14 Or intC = 16 Then
                ElseIf intC > 5 And intC < 20 Then
                    strIntBase = GetText(.spdExcel, intR, intC)
                    strResult = GetText(.spdExcel, intR + 1, intC)
                    
                    '-- 2019-12-02 수정 : 결과가 양성일 경우 Ct값을 같이 보여준다 ex) Positive(20.5)
                    strIntResult = GetText(.spdExcel, intR + 1, intC + 1)
                    
                    If strResult = "-" Then
                        strResult = gCmnt.RP19NEG
                    ElseIf strResult = "+" Then
                        strResult = gCmnt.RP19POS
                        If strIntBase = "MP" Or strIntBase = "LP" Or strIntBase = "BP" Or strIntBase = "CP" Or strIntBase = "HI" Or strIntBase = "SP" Then
                            strResult = strResult & "(" & strIntResult & ")"
                        End If
                    ElseIf strResult = "++" Then
                        strResult = gCmnt.RP19POS
                        If strIntBase = "MP" Or strIntBase = "LP" Or strIntBase = "BP" Or strIntBase = "CP" Or strIntBase = "HI" Or strIntBase = "SP" Then
                            strResult = strResult & "(" & strIntResult & ")"
                        End If
                    ElseIf strResult = "+++" Then
                        strResult = gCmnt.RP19POS
                        If strIntBase = "MP" Or strIntBase = "LP" Or strIntBase = "BP" Or strIntBase = "CP" Or strIntBase = "HI" Or strIntBase = "SP" Then
                            strResult = strResult & "(" & strIntResult & ")"
                        End If
                    ElseIf strResult = "++++" Then
                        strResult = gCmnt.RP19POS
                        If strIntBase = "MP" Or strIntBase = "LP" Or strIntBase = "BP" Or strIntBase = "CP" Or strIntBase = "HI" Or strIntBase = "SP" Then
                            strResult = strResult & "(" & strIntResult & ")"
                        End If
                    ElseIf strResult = "+++++" Then
                        strResult = gCmnt.RP19POS
                        If strIntBase = "MP" Or strIntBase = "LP" Or strIntBase = "BP" Or strIntBase = "CP" Or strIntBase = "HI" Or strIntBase = "SP" Then
                            strResult = strResult & "(" & strIntResult & ")"
                        End If
                    Else
                        strResult = ""
                    End If
                    
                    
                    
                    If strIntBase <> "" And strResult <> "" And strIntBase <> "Result" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
                        SQL = SQL & "       , ABBRNAME                                                  " & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER                                                   " & vbCrLf
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                        " & vbCrLf
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                          " & vbCrLf
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")                             " & vbCrLf
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
                            strTestName = Trim(RS_L.Fields("TESTNAME"))
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("RESPRECUSE") & "")
                            strAbbrName = Trim(RS_L.Fields("ABBRNAME") & "")
    
                            '-- 결과Row 추가
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                'gArrEQPNm(introw,6)
                                'If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                If strAbbrName = Trim(gArrEQPNm(intCol - colSTATE, 6)) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '순번
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                
                End If
                
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
            Next
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "FileRcvData_CFX96" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub FileRcvData_CFX96_COVID19()
    Dim RS_L            As ADODB.Recordset
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strOldSeq          As String   '수신한 Sequence
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
    Dim j               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTarget       As String
    Dim strVarTest      As String
    Dim strLGrp         As String
    Dim strHGrp         As String
    Dim strTotFlag      As String
    
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    Dim intR            As Integer
    Dim intC            As Integer
    Dim strHospSite     As String
    Dim strHospGbn      As String

    Dim strRstRow       As String
    Dim strAbbrName     As String
    
    Dim strEgene        As String
    Dim strRdRPgene     As String
    
On Error GoTo ErrHandle
    
    With frmMain
        For intR = 3 To spdExcel.DataRowCnt
            For intC = 3 To .spdExcel.DataColCnt      'MaxCols
                strIntBase = ""
                strResult = ""
                '이름(검사번호)
                If intC = 4 Then
                    strSeq = GetText(.spdExcel, intR, intC)
                    strRackNo = GetText(.spdExcel, intR, intC - 1)
                    
                    'If strBarno = "" Then
                    '    strBarno = strRackNo
                    'End If
                    
                    If strOldSeq <> strSeq Then
                        '-- 결과정보
                        With mResult
                            .Seq = strSeq
                           ' .BarNo = strBarno
                            .RackNo = strRackNo
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            '.SITE = strHospSite
                            'gHOSP.SITE = strHospSite
                        End With
                                    
                        '-- 결과환자정보
                        Call SetPatInfo(strSeq, gHOSP.RSTTYPE)
                    End If
                    
                    strOldSeq = strSeq
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                'ElseIf intC = 14 Then
                ElseIf intC = 7 Or intC = 9 Or intC = 11 Or intC = 14 Then
                        If intC = 14 Then
                            strIntBase = "COVID-19"
                            strResult = GetText(.spdExcel, intR, intC)
                            If strResult = "-" Then
                                strResult = "Negative"
                            End If
                            If strResult = "+" Then
                                strResult = "Positive"
                            End If
                            If strResult = "SARS-CoV-2" Then
                                strResult = "Positive"
                            End If
                        ElseIf intC = 7 Then
                            strIntBase = "E"
                            strResult = GetText(.spdExcel, intR, intC)
                        ElseIf intC = 9 Then
                            strIntBase = "RdRp"
                            strResult = GetText(.spdExcel, intR, intC)
                        ElseIf intC = 11 Then
                            strIntBase = "N"
                            strResult = GetText(.spdExcel, intR, intC)
                        End If
                        
                        If strIntBase <> "" And strResult <> "" And strIntBase <> "Result" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
                            SQL = SQL & "      ,ABBRNAME "
                            SQL = SQL & "  FROM EQPMASTER                                                   " & vbCrLf
                            SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                        " & vbCrLf
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                          " & vbCrLf
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")                             " & vbCrLf
                            End If
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqno = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("RESPRECUSE") & "")
                                strAbbrName = Trim(RS_L.Fields("ABBRNAME") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    'If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    If strAbbrName = Trim(gArrEQPNm(intCol - colSTATE, 6)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        If InStr(strResult, "Pos") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbRed
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        ElseIf InStr(strResult, "Neg") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlue
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        Else
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlack
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = False
                                        End If
                                        
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS결과
                                
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
                                
                                'SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                'SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '참고치
                                
                                If InStr(UCase(strResult), "POS") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbRed
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                ElseIf InStr(UCase(strResult), "NEG") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlue
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                Else
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlack
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = False
                                End If
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        End If
                    'Next
                End If
                
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
            Next
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "FileRcvData_CFX96" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub FileRcvData_CFX96_MTB()
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
    Dim j               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTarget       As String
    Dim strVarTest      As String
    Dim strLGrp         As String
    Dim strHGrp         As String
    Dim strTotFlag      As String
    
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    Dim intR            As Integer
    Dim intC            As Integer
    Dim strHospSite     As String
    Dim strHospGbn      As String

    Dim strRstRow       As String
    Dim strAbbrName     As String

On Error GoTo ErrHandle
    
    With frmMain
        For intR = 2 To spdExcel.MaxRows Step 3
            For intC = 3 To .spdExcel.MaxCols
                strIntBase = ""
                strResult = ""
                If intC = 3 Then
                    strBarno = GetText(.spdExcel, intR, intC)
                    strRackNo = GetText(.spdExcel, intR, intC + 1)
                    
                    If strBarno = "" Then
                        strBarno = strRackNo
                    End If
                    
                    If strOldBarno <> strBarno Then
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RackNo = strRackNo
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            '.SITE = strHospSite
                            'gHOSP.SITE = strHospSite
                        End With
                                    
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                ElseIf intC = 5 Then
                    For j = 0 To 1
                        strIntBase = GetText(.spdExcel, intR + j, intC)
                        strResult = GetText(.spdExcel, intR + j, intC + 1)
                        
                        If strResult = "-" Then
                            strResult = gCmnt.MTBNEG
                        ElseIf strResult = "+" Then
                            strResult = gCmnt.MTBPOS
                        ElseIf strResult = "++" Then
                            strResult = gCmnt.MTBPOS
                        ElseIf strResult = "+++" Then
                            strResult = gCmnt.MTBPOS
                        ElseIf strResult = "++++" Then
                            strResult = gCmnt.MTBPOS
                        ElseIf strResult = "+++++" Then
                            strResult = gCmnt.MTBPOS
                        Else
                            strResult = ""
                        End If
                        
                        If strIntBase <> "" And strResult <> "" And strIntBase <> "Result" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
                            SQL = SQL & "      ,ABBRNAME "
                            SQL = SQL & "  FROM EQPMASTER                                                   " & vbCrLf
                            SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                        " & vbCrLf
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                          " & vbCrLf
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")                             " & vbCrLf
                            End If
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
                                strTestName = Trim(RS_L.Fields("TESTNAME"))
                                strSeqno = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("RESPRECUSE") & "")
                                strAbbrName = Trim(RS_L.Fields("ABBRNAME") & "")
        
                                '-- 결과Row 추가
                                strRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < strRstRow Then
                                    .spdResult.MaxRows = strRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    'If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                    If strAbbrName = Trim(gArrEQPNm(intCol - colSTATE, 6)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        If InStr(strResult, "Pos") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbRed
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        ElseIf InStr(strResult, "Neg") > 0 Then
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlue
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = True
                                        Else
                                            .spdOrder.Row = gRow
                                            .spdOrder.Col = intCol
                                            .spdOrder.ForeColor = vbBlack
                                            .spdOrder.FontSize = 10
                                            .spdOrder.FontBold = False
                                        End If
                                        
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '순번
                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '장비채널
                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '장비결과
                                SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS결과
                                
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
                                
                                'SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
                                'SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '참고치
                                
                                If InStr(UCase(strResult), "POS") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbRed
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                ElseIf InStr(UCase(strResult), "NEG") > 0 Then
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlue
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = True
                                Else
                                    .spdResult.Row = strRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlack
                                    .spdResult.FontSize = 10
                                    .spdResult.FontBold = False
                                End If
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, strRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        End If
                    Next
                End If
                
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
            Next
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "FileRcvData_CFX96" & vbNewLine & vbNewLine
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
'    SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7               " & vbCrLf
'    SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7              " & vbCrLf
'    SQL = SQL & "     , AMRLimit8,  AMRLimit9,  AMRLimit10,  AMRLimit11,  AMRLimit12,  AMRLimit13,  AMRLimit14          " & vbCrLf
'    SQL = SQL & "     , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14         " & vbCrLf
'    SQL = SQL & "     , AMRINResult                                                                                     " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER                                                                            " & vbCrLf
    SQL = SQL & " WHERE EQPMASTER.EQUIPCD     = '" & gHOSP.MACHCD & "'                                                            " & vbCrLf
    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = '" & pIntBase & "'                                                                " & vbCrLf
    If gPatOrdCd <> "" Then
        SQL = SQL & "   AND EQPMASTER.TESTCODE in (" & gPatOrdCd & ") "
    End If
'    SQL = SQL & "   AND EQPMASTER.EQUIPCD     = AMRMASTER.EQUIPCD                                                       " & vbCrLf
'    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = AMRMASTER.RSLTCHANNEL                                                   " & vbCrLf
'    SQL = SQL & "   AND EQPMASTER.TESTCODE    = AMRMASTER.TESTCODE                                                      "
    
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
'            If IsNumeric(pIntResult) Then
'                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
'                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
'                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
'                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
'                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
'                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
'                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
'                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
'                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
'                    End If
'                End If
'            End If
            
            pResult = pIntResult
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
            End If
        End If
        
        '수치결과 포함
        '0:사용안함, 1:정성(정량), 2:정량(정성), 3:정성_정량, 4:정량_정성
        If strAMRResult <> "" Then
            If strIntResult <> "" Then
                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                    pResult = strAMRResult
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                    pResult = strAMRResult & "(" & strIntResult & ")"
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


Private Sub SerialRcvData_HITACHI7180()
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
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            'Call SetSQLData("RCV", strRcvBuf, "")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
                    'gHOSP.BARLEN = 13
                    If gHOSP.BARUSE = "Y" Then
                        strFunction = Mid(strRcvBuf, 2, 40)
                    Else
                        strFunction = Mid(strRcvBuf, 2, 12) & String(gHOSP.BARLEN, "#") & Mid(strRcvBuf, 27, 15)
                    End If
                    strBarno = Trim(Mid(strRcvBuf, 14, gHOSP.BARLEN))
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    '-- 오더정보
                    With mOrder
                        .Function = strFunction
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- 오더환자정보
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                    '## Control, Calibration 데이터는 무시함
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
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

                    If mResult.BarNo = "" Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If

                    strTmp = Mid$(strRcvBuf, 51)
                    
                    Do While Len(strTmp) >= 10
                        strIntBase = Trim(Mid$(strTmp, 1, 3))
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 9, 1)
                        
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                            
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- 검사마스터 정보 가져오기
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- 결과채널이 같고...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
'                                    '-- 환자 처방정보 가져오기
'                                    If UBound(gPatTest) > 0 Then
'                                        For intOrdCnt = 1 To UBound(gPatTest)
'                                            For intTestCdCnt = 1 To UBound(gArrEQP)
'                                                '-- 검사코드도 같다면...
'                                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'                                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'                                                    strTestName = gArrEQP(intTestCdCnt, 5)
'                                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'                                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'                                                    '-- 남자참고치를 기본으로 한다
'                                                    strLow = gArrEQP(intTestCdCnt, 9)
'                                                    strHigh = gArrEQP(intTestCdCnt, 10)
'
'                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
'                                                        strCheck = "1"
'
'                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
'                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
'
'                                                        If mPatient.SEX = "M" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 9)
'                                                            strHigh = gArrEQP(intTestCdCnt, 10)
'                                                        ElseIf mPatient.SEX = "F" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 11)
'                                                            strHigh = gArrEQP(intTestCdCnt, 12)
'                                                        Else
'                                                            strLow = ""
'                                                            strHigh = ""
'                                                        End If
'                                                        strState = "R"
'                                                        blnSame = True
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            Next
'                                        Next
'                                    End If
                                    For intTestCdCnt = 1 To UBound(gArrEQP)
                                        '-- 검사코드도 같다면...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            '-- 남자참고치를 기본으로 한다
                                            strLow = gArrEQP(intTestCdCnt, 9)
                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                            
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        strCheck = "1"
                                                        
                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                        
                                                        If mPatient.SEX = "M" Then
                                                            strLow = gArrEQP(intTestCdCnt, 9)
                                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                                        ElseIf mPatient.SEX = "F" Then
                                                            strLow = gArrEQP(intTestCdCnt, 11)
                                                            strHigh = gArrEQP(intTestCdCnt, 12)
                                                        Else
                                                            strLow = ""
                                                            strHigh = ""
                                                        End If
                                                        strState = "R"
                                                        blnSame = True
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                    
                                    If blnSame = True Then
                                        Exit For
                                    End If
                                End If
                            Next

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
                                    Exit For
                                End If
                            Next

                            '-- 결과 List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '순번
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
                            '-- 처방이 있을때만 검사코드를 저장한다.
                            If strState = "R" Then
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '검사코드
                            End If
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
                        End If
                        strTmp = Mid$(strTmp, 11)
                    Loop

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

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub Phase_Serial_AU680()
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
                    Case ETB
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_AU680
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub


Private Sub Phase_File_CFX96()
                        
    If spdExcel.MaxRows > 0 Then
        
        If optTest(0).Value = True Then
            'Call FileRcvData_CFX96_MTB
            Call FileRcvData_CFX96_COVID19
        ElseIf optTest(1).Value = True Then
            Call FileRcvData_CFX96_RP19
        ElseIf optTest(2).Value = True Then
            Call FileRcvData_CFX96_RV6
        ElseIf optTest(3).Value = True Then
            Call FileRcvData_CFX96_RP19
        End If
            
        
        lblRow.Caption = "1"
        txtBarNum.SetFocus
    
    End If

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

Private Sub frmClear()
    
    shpPatInfo.Visible = False
    lblPatInfo.Caption = ""
    
    spdWork.MaxRows = 0
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    spdExcel.MaxRows = 0
    
    txtBarcode.Text = ""
    txtPatID.Text = ""
    txtPName.Text = ""
    txtSA.Text = ""
    txtBarNum.Text = ""
    lblRow.Caption = 0
    
End Sub

Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    
On Error GoTo ErrHandle
    
    Me.Width = 20940
    Me.Height = 12585

    'Me.Caption = gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"

    Call CtlInitializing

    Call frmClear
    
    '-- Menu Set
    Call SetMenu

    '-- 컬럼보이기설정
    Call SetColumnView(spdOrder)

    '-- 검사코드
    Call GetTestList

    '-- 검사명(약어)
    Call GetTestListName

    '-- 검사코드별 검사명
    Call GetTestCode_Name

    '-- 검사명 보이기
    Call SetExamCode(spdOrder)

    '-- 통신열기
    Call OpenCommunication


'Public gUser() As String


    cboUser.Clear
    For intCnt = 1 To gUserCount
        ReDim Preserve gUser(intCnt) As String
        
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("USER", CStr(intCnt), "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gUser(intCnt) = Trim(strSetUp1)
        
        cboUser.AddItem gUser(intCnt)
    Next
    
''    'cboUser.AddItem "-- 사용자 --"
''    cboUser.AddItem "김은헤|103787"
''    cboUser.AddItem "김현지|111233"
''    cboUser.AddItem "구본상|105142"
''    cboUser.AddItem "정영호|106438"
''    cboUser.AddItem "김준|114516"
'''    cboUser.AddItem "배영란|183130"
''    cboUser.AddItem "정이진|186762"
''    cboUser.AddItem "방주영|188274"
    
    '111235 김근엽
    
'    cboUser.ListIndex = 0
    
    pDel = False

    spdComStatus.MaxRows = 0
    spdComStatus.Font.Bold = True
    
    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    txtTestID.Text = gHOSP.USERID
    txtTestNm.Text = gHOSP.USERNM
    lblPatInfo.Caption = ""
    lblAPI.Caption = "연결서버: " & gHOSP.APIURL
    
    dtpFrom.Value = Now
    dtpTo.Value = Now

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
    
'spdOrder.MaxRows = 30

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & " 연결실패"
            
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
            lblComStatus.Caption = "COM" & comEqp.CommPort & " 연결성공"
            
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & " 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    ElseIf gComm.COMTYPE = "2" Then
'        If gComm.TCPTYPE = "1" Then
'            wSck.LocalPort = CInt(gComm.TCPPORT)
'            wSck.Listen
'
'            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 연결..."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgSend.Visible = False
'            imgReceive.Visible = False
'            lblSend.Visible = False
'            lblRcv.Visible = False
'
'        Else
'            wSck.Close
'            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
'
'            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 연결..."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgSend.Visible = False
'            imgReceive.Visible = False
'            lblSend.Visible = False
'            lblRcv.Visible = False
'        End If
    ElseIf gComm.COMTYPE = "3" Then
        imgPort.Visible = False
        imgSend.Visible = False
        imgReceive.Visible = False
        lblPort.Visible = False
        lblSend.Visible = False
        lblRcv.Visible = False
        
        lblComStatus.Left = lblPort.Left
        lblComStatus.Width = lblComStatus.Width * 3
        'lblComStatus.Caption = "결과 경로: " & gComm.RSTPATH
        lblComStatus.Caption = "◈ 오더경로: " & gComm.ORDPATH & "   ◈ 결과경로: " & gComm.RSTPATH

    End If

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    Me.Top = 0

'    spdWork.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - fraWorkInfo.Height - 300
    
    spdOrder.Top = spdOrder.Top + 40
    spdOrder.Width = Me.ScaleWidth - spdResult.Width - 200
    spdOrder.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - 300
    
    spdResult.Left = spdOrder.Left + spdOrder.Width + 50
    spdResult.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - fraPatInfo.Height - 300

    fraPatInfo.Left = spdOrder.Left + spdOrder.Width + 50
    fraPatInfo.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - spdResult.Height - 300

    cmdSL.Top = spdOrder.Top + 10
    cmdSL.Left = spdOrder.Left + 10
    
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
    SQL = SQL & "SELECT DISTINCT SEQNO, EXAMNAME, EXAMCODE, RESULT, PREVRESULT, REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND MID(EXAMDATE,1,10) = '" & Mid(strExamDate, 1, 10) & "'" & vbCr
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
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
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
                
                If InStr(Trim(UCase(AdoRs_Local.Fields("RESULT").Value & "")), "POS") > 0 Then
                    .ForeColor = vbRed
                    .FontBold = True
                ElseIf InStr(Trim(UCase(AdoRs_Local.Fields("RESULT").Value & "")), "NEG") > 0 Then
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

Private Sub Label5_DblClick()

    If spdExcel.Visible = True Then
        spdExcel.Visible = False
        spdExcel.ZOrder 0
    Else
        spdExcel.Visible = True
        spdExcel.ZOrder 0
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
    
    frmConfig.Show vbModal

End Sub

Private Sub mnuComment_Click()

    frmComment.Show vbModal

End Sub

Private Sub mnuCommTest_Click()

    If picComm.Visible = True Then
        picComm.Visible = False
        picComm.ZOrder 0
    Else
        picComm.Visible = True
        picComm.ZOrder 0
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHosp_Click()

    frmHospInfo.Show vbModal

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
    
    frmResult.Show vbModal
    
End Sub

Private Sub mnuResultTrans_Click()
    
    frmResultTrans.Show vbModal

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
    
    frmTestSet.Show vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show ' vbModal

End Sub

Private Sub optTest_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        optTest(i).BackColor = &H800000
    Next
    
    optTest(Index).BackColor = &HFF80FF
    
    frmWorkList.optTest(Index).Value = True
    

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
    
    lblRow.Caption = Row
    
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

End Sub

Private Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarno As String
    Dim intRow      As Integer
    Dim strSeq      As String
    
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarno = GetText(spdOrder, sRow, sCol)
    
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
        If strNewBarno = "" Then
        
        End If
        
        If MsgBox(strNewBarno & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
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
    
    '-- 정렬
    If Row = 0 Then
        '-- 정렬 추가
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
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
            Next
            
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
            .RowHeight(-1) = 15
        End If
    
    End With
    
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
                .MaxRows = .MaxRows + 1
                sRow = .MaxRows
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                'Call spdOrder_KeyDown(13, 1)
                
                If GetSampleInfo(sRow, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                    txtBarNum.SelStart = 0
                    txtBarNum.SelLength = Len(txtBarNum.Text)
                    Exit Sub
                    
                Else
                    '정보수정
'                    SQL = ""
'                    SQL = SQL & "UPDATE PATRESULT SET "
'                    SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
'                    SQL = SQL & " ,HOSPDATE = '" & Trim(GetText(spdOrder, sRow, colHOSPDATE)) & "'" & vbCrLf
'                    SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
'                    SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
'                    SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
'                    SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
'                    SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
'                    SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
'                    SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
'                    SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
'                    SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
'                    SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
'                    SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
'                    SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
'                    SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
'                    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
'                    SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
'                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                    
'                    If DBExec(AdoCn_Local, SQL) Then
'                        '-- 성공
'                    End If
                End If
            End With
        Else
            With spdOrder
                sRow = lblRow.Caption ' .ActiveRow
'                If sRow < 1 Then
'                    Exit Sub
'                End If
                
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                
                'Call spdOrder_KeyDown(13, 1)
                
                If GetSampleInfo(.Row, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                    txtBarNum.SelStart = 0
                    txtBarNum.SelLength = Len(txtBarNum.Text)
                    Exit Sub
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

