VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   -1530
   ClientWidth     =   21900
   BeginProperty Font 
      Name            =   "????"
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
   StartUpPosition =   1  '?????? ???
   WindowState     =   2  '?ִ?ȭ
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   60
      TabIndex        =   2
      Top             =   630
      Width           =   16335
      _Version        =   393216
      _ExtentX        =   28813
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
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
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":0E42
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      TabIndex        =   45
      Top             =   660
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   5310
         Style           =   1  '?׷???
         TabIndex        =   51
         Top             =   210
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "??ũ??ȸ"
         Height          =   375
         Left            =   4200
         Style           =   1  '?׷???
         TabIndex        =   50
         Top             =   210
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   46
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139591681
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   47
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139591681
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '????
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2580
         TabIndex        =   49
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '????
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "???????? :"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   48
         Top             =   330
         Width           =   930
      End
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '?Ʒ? ????
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   21840
      TabIndex        =   33
      Top             =   10620
      Visible         =   0   'False
      Width           =   21900
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "C"
         Height          =   495
         Left            =   12930
         TabIndex        =   43
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   42
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   41
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   40
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   39
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   38
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         Height          =   555
         Left            =   13560
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   30
         Width           =   3435
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   525
         Left            =   17010
         TabIndex        =   36
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "Rcv"
         Height          =   525
         Left            =   11910
         TabIndex        =   34
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   16440
      TabIndex        =   20
      Top             =   660
      Width           =   6525
      Begin VB.TextBox txtSA 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "1004"
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "1004"
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtPatID 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "1004"
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1004"
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "S / A"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   32
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "??      ??"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   30
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "???Ϲ?ȣ"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   28
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "??ü??ȣ"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   26
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   2355
      Left            =   20250
      TabIndex        =   18
      Top             =   8580
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdWork 
         Caption         =   "??ũ??ȸ"
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResultSearch 
         Caption         =   "??????ȸ"
         Height          =   315
         Left            =   2610
         TabIndex        =   19
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lblPatInfo 
         BackStyle       =   0  '????
         Caption         =   "?ڰ˻?"
         BeginProperty Font 
            Name            =   "????ü"
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
         TabIndex        =   24
         Top             =   990
         Width           =   3465
      End
      Begin VB.Shape shpPatInfo 
         BorderColor     =   &H00FF0000&
         Height          =   1155
         Left            =   900
         Shape           =   4  '?ձ? ?簢??
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
   Begin VB.PictureBox picBottom 
      Align           =   2  '?Ʒ? ????
      BackColor       =   &H00404040&
      BorderStyle     =   0  '????
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   7
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
               Picture         =   "frmMain.frx":2B86
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3120
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":36BA
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3C54
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":44E6
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4640
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":479A
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread spdComStatus 
         Height          =   330
         Left            =   8010
         TabIndex        =   13
         Top             =   120
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
            Name            =   "????"
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
         SpreadDesigner  =   "frmMain.frx":48F4
         UserResize      =   0
         TextTip         =   2
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Label lblAPI 
         BackStyle       =   0  '????
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   11760
         TabIndex        =   54
         Top             =   180
         Width           =   6105
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   7980
         Top             =   90
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3900
         Picture         =   "frmMain.frx":4CFA
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   4845
         Picture         =   "frmMain.frx":5284
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   5760
         Picture         =   "frmMain.frx":580E
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "??Ʈ"
         BeginProperty Font 
            Name            =   "????"
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
         TabIndex        =   12
         Top             =   210
         Width           =   360
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "?۽?"
         BeginProperty Font 
            Name            =   "????"
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
         TabIndex        =   11
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "????"
         BeginProperty Font 
            Name            =   "????"
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
         TabIndex        =   10
         Top             =   210
         Width           =   420
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":5D98
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":5EE2
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":602C
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '????
         Caption         =   "Com1 ???Ἲ??"
         BeginProperty Font 
            Name            =   "???? ????"
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
         TabIndex        =   9
         Top             =   180
         Width           =   5235
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2955
      End
      Begin VB.Label lblDBStatus 
         Alignment       =   2  '??? ????
         BackStyle       =   0  '????
         Caption         =   "?????ͺ??̽? ???Ἲ??"
         BeginProperty Font 
            Name            =   "???? ????"
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
         TabIndex        =   8
         Top             =   180
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3210
         Top             =   90
         Width           =   8415
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6555
      Left            =   16440
      TabIndex        =   4
      Top             =   2010
      Width           =   6495
      _Version        =   393216
      _ExtentX        =   11456
      _ExtentY        =   11562
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
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
      SpreadDesigner  =   "frmMain.frx":6176
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '?? ????
      BackColor       =   &H00800000&
      BorderStyle     =   0  '????
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   0
      Top             =   0
      Width           =   21900
      Begin VB.CheckBox chkAdd 
         Appearance      =   0  '????
         BackColor       =   &H00800000&
         Caption         =   "O"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   9630
         TabIndex        =   57
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '??? ????
         BorderStyle     =   0  '????
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   56
         Text            =   "12121334133"
         Top             =   180
         Width           =   1845
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   19380
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "?????ޱ?"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13740
         TabIndex        =   53
         ToolTipText     =   "?????? ?????? EMR?????? ?????մϴ?"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "????????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12210
         TabIndex        =   52
         ToolTipText     =   "????ȭ???? ???? ?????ϴ?"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȭ??????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15270
         TabIndex        =   23
         ToolTipText     =   "????ȭ???? ???? ?????ϴ?"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "????????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   16800
         TabIndex        =   22
         ToolTipText     =   "?????? ?????? EMR?????? ?????մϴ?"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdTestNmSave 
         Caption         =   "????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8970
         TabIndex        =   17
         Top             =   150
         Width           =   555
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '????
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7410
         TabIndex        =   16
         Text            =   "1004"
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '????
         BeginProperty Font 
            Name            =   "????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4230
         TabIndex        =   15
         Text            =   "1004"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdTestIDSave 
         Caption         =   "????"
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   14
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
      Begin VB.Shape Shape13 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   18270
         Top             =   90
         Width           =   615
      End
      Begin VB.Label lblRow 
         Alignment       =   2  '??? ????
         BackStyle       =   0  '????
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   18360
         TabIndex        =   58
         Top             =   210
         Width           =   435
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   13680
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   12150
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
         BackStyle       =   0  '????
         Caption         =   "1971-03-11"
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   180
         Width           =   1365
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2865
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   6270
         Top             =   90
         Width           =   3315
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '????
         Caption         =   "?˻??ڸ? : "
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6390
         TabIndex        =   5
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '????
         Caption         =   "?˻???ID : "
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   180
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3120
         Top             =   90
         Width           =   3165
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '????
         Caption         =   "?˻????? :"
         BeginProperty Font 
            Name            =   "???? ????"
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
      TabIndex        =   44
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
         Name            =   "???? ????"
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
      SpreadDesigner  =   "frmMain.frx":6DBE
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdExcel 
      Height          =   7065
      Left            =   3180
      TabIndex        =   55
      Top             =   2820
      Visible         =   0   'False
      Width           =   11955
      _Version        =   393216
      _ExtentX        =   21087
      _ExtentY        =   12462
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
      SpreadDesigner  =   "frmMain.frx":8FBC
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ????  "
      Begin VB.Menu mnuHosp 
         Caption         =   "?? ???? ????"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "?? EMR ????"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "????"
      End
   End
   Begin VB.Menu mnuMenu04 
      Caption         =   "  ??ȸ  "
      Begin VB.Menu mnuResult 
         Caption         =   "?? ???? ??ȸ"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "?? ??ũ ??ȸ"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "  ????  "
      Begin VB.Menu mnuComm 
         Caption         =   "?? ???? ????"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "?? ?˻? ????"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "?? ȭ?? ????"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResultTrans 
         Caption         =   "?? ???? ??ȯ"
      End
      Begin VB.Menu mnuSep231 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComment 
         Caption         =   "?? ?ڸ?Ʈ ????"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep24 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "?? ?ɼ? ????"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " ?˻??ɼ? "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "?? ???ڵ? ????"
         Begin VB.Menu mnuBarcode 
            Caption         =   "???ڵ?????"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "????????"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "üũ??"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "?? ???? ????"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "?ڵ?"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "????"
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "?? ???? ????"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "????????"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS????"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " ???????? "
      Begin VB.Menu mnuHelp01 
         Caption         =   "????????(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "????????(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "????????(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "?????׽?Ʈ"
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

    If MsgBox("?????? ???????Դϴ?. ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, "???α׷? ????") = vbYes Then

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
                                    Call SetText(spdOrder, "??", spdOrder.MaxRows, intOCol)
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
    Dim FileNumber
    Dim strAllStr   As String

    strAllStr = ""
    intCnt = 0
    
    '''---???? ????---
    '''???ϸ??? nimbus.lis ?? ????
    ' "1    " & "BCD1               " & "RV16" & space(146) &
    ' "2    " & "BCD2               " & "RV16" & space(146) &
    '1    ]C114072901TC480257 RP1
    '2    2013-08-02/04655    RP1
    
    With spdOrder
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                strBC = GetText(spdOrder, intRow, colBARCODE)
                If strBC <> "" Then
                    intCnt = intCnt + 1
                    
                    strAllStr = strAllStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) & _
                                            strBC & Space(20 - Len(CStr(intCnt))) & _
                                            "RP1" & Space(147)  '175

                    '-- ????????(Order) ǥ??
                    Call SetText(spdOrder, "????????", intRow, colSTATE)
                    Call SetText(spdOrder, "0", intRow, colCHECKBOX)
                End If
            End If
        Next
    End With

    If strAllStr <> "" Then
        'nimbus.lis ???? ?????? ======================================================
        '-- OPEN
        FileNumber = FreeFile
        Open gComm.ORDPATH & "\nimbus_rb6.lis" For Output As FileNumber
        Close FileNumber
        '-- WRITE
        Open gComm.ORDPATH & "\nimbus_rb6.lis" For Append As FileNumber
        Print #FileNumber, strAllStr
        Close FileNumber
        '=========================================================================
        MsgBox "???? ???? ???? ?Ϸ?", vbOKOnly + vbInformation, Me.Caption
    Else
        MsgBox "?????? ?????? ?????ϴ?", vbOKOnly + vbCritical, Me.Caption
    End If


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
    Dim intIdx      As Integer
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

    Call Excel_Open

    Call Phase_File_CFX96
    
Exit Sub

ErrRoutine:

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
    
    
    '???? ????
    With CFXFile
        .InitDir = gComm.RSTPATH
        .Filename = "*.xls"
        '.Filter = "Resource CSV (*.CSV)|*.CSV|All File (*.*)|*.*|"
        .Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx"
        .DialogTitle = "CFX96 ?ڷ? ?о?????"
        .ShowOpen
    End With
    
    
    If CFXFile.FileTitle = "" Then
        Exit Sub
    End If
    
    xlapp.Workbooks.Open (Trim(CFXFile.Filename))
    
    lngSCnt = xlapp.Worksheets.Count
    
    '-- ??ü ??ũ??Ʈ ?ҷ??????ͼ? '?ӽ?.txt' ???Ϸ? ????
    For i = 1 To lngSCnt
        Set XLappWS = xlapp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlapp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ????
        xlapp.ActiveWorkbook.SaveAs "C:\IF_CFX96\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ????
        
        
        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>???? 2000??
        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>???? 2003??
    Next i
    
    xlapp.Quit
    Set XLappWS = Nothing
    Set xlapp = Nothing
    
    '-- ??ü ?????? MAX cols?? ????
    dummy_max = 0
    For i = 1 To lngSCnt
        If lngSColCnt(i) >= dummy_max Then
            dummy_max = lngSColCnt(i)
        End If
    Next i
    lngTotColCnt = dummy_max
    
    '??ü row?? ????
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
    
    '-- ?׸??? ?ʱ?ȭ
    spdExcel.MaxRows = 0
    spdExcel.MaxRows = lngTotRowCnt
    spdExcel.MaxCols = lngTotColCnt
    
    '-- ?׸??忡 ????
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
        .SortKey(1) = 6       '????Ű ????ȣ
        .SortKey(2) = 2       '????Ű ????ȣ

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
    
    If MsgBox("?????? ?????? ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, "????????") = vbYes Then
        With spdOrder
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                    Res = SaveTransData(lRow, spdOrder)
                    
                    If Res = -1 Then
                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdOrder, "????????", lRow, colSTATE
                    
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ????
                        End If
                    
                    Else
                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdOrder, "?????Ϸ?", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ????
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
            EVMsg$ = "CTS ???? ????"
        Case comEvDSR
            EVMsg$ = "DSR ???? ????"
        Case comEvCD
            EVMsg$ = "CD ???? ????"
        Case comEvRing
            EVMsg$ = "??ȭ ???? ?︮?? ??"
        Case comEvEOF
            EVMsg$ = "EOF ????"

        '???? ?޽???
        Case comBreak
            ERMsg$ = "?ߴ? ??ȣ ????"
        Case comCDTO
            ERMsg$ = "?ݼ??? ???? ?ð? ?ʰ?"
        Case comCTSTO
            ERMsg$ = "CTS ?ð? ?ʰ?"
        Case comDCB
            ERMsg$ = "DCB ?˻? ????"
        Case comDSRTO
            ERMsg$ = "DSR ?ð? ?ʰ?"
        Case comFrame
            ERMsg$ = "?????̹? ????"
        Case comOverrun
            ERMsg$ = "?и?Ƽ ????"
        Case comRxOver
            ERMsg$ = "???? ???? ?ʰ?"
        Case comRxParity
            ERMsg$ = "?и?Ƽ ????"
        Case comTxFull
            ERMsg$ = "???? ???ۿ? ?????? ????"
        Case Else
            ERMsg$ = "?? ?? ???? ???? ?Ǵ? ?̺?Ʈ"
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call cmdEnd_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If MsgBox("?????? ???????Դϴ?. ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, "???α׷? ????") = vbYes Then
        
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
    
    '-- 1. ???????? ??ȸ
    With frmMain
        Select Case pType
            '-- ???ڵ? ????
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
    

        '-- ???????忡?? ??ã????..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ???????????? ȭ??ǥ??
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ???????????? ??????
        .spdResult.MaxRows = 0

        '-- ?˻??? ???? ????????
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = GetEquipExamCode_AU680(gHOSP.MACHCD, pBarNo, intRow)

        '-- ?˻?ä?η? ???????? ??????
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & ETX

            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX

            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If

        '-- ?????? ????
        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

        '-- ???? Row
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
    
    '-- 1. ???????? ??ȸ
    With frmMain
        Select Case pType
            '-- ???ڵ? ????
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
    

        '-- ???????忡?? ??ã????..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ???????????? ȭ??ǥ??
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ???????????? ??????
        .spdResult.MaxRows = 0

        '-- ?˻??? ???? ????????
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        If gHOSP.BARUSE = "N" Then
            mOrder.Function = Replace(mOrder.Function, String(gHOSP.BARLEN, "#"), Left(mOrder.BarNo & Space(gHOSP.BARLEN), gHOSP.BARLEN))
        End If
        
        '-- ?˻?ä?η? ???????? ??????
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

        '-- ???? Row
        gRow = intRow

    End With

End Sub

Private Sub SendData(ByVal pSendData As Variant)

    '-- ????
    comEqp.Output = pSendData
    
    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    
    '-- ?αױ???
    Call SetRawData("[Tx]" & pSendData)

End Sub

Private Sub SerialRcvData_AU680()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strMachResult   As String   '?????? ????????
    Dim strResult       As String   '?????? ????(????)
    Dim strIntResult    As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    Dim strCheck        As String   '?˻?????üũ
    Dim strSeqno        As String   '?˻?????
    Dim strOrderCode    As String   'ó???ڵ?
    Dim strTestName     As String   '?˻??ڵ?
    Dim strTestCode     As String   '?˻??ڵ?
    Dim strTestCodeSub  As String   '?˻??ڵ?SUB
    Dim intResPrecUse   As Integer  '?Ҽ?????ȯ????
    Dim intResPrec      As Integer  '?Ҽ????ڸ???
    Dim strResType      As String   '?Ҽ?????ȯ????
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '????????
    Dim strPrevRslt     As String   '????????
    
    Dim intRstRow       As String   '???????????? ???? Row
    Dim intCnt          As Integer  '???? Frame ????
    Dim intCol          As Integer  '?????÷? ????
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
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
                    '-- ????????
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- ????ȯ??????
                    Call GetOrder_AU680(Trim$(strBarno), gHOSP.RSTTYPE)

                Case "D "    '## Result
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 10, 4)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN))
                    
                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ????ȯ??????
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
                            '-- ?˻縶???? ???? ????????
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ????ä???? ????...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
                                    '-- ȯ?? ó?????? ????????
'                                    If UBound(gPatTest) > 0 Then
'                                        For intOrdCnt = 1 To UBound(gPatTest)
'                                            For intTestCdCnt = 1 To UBound(gArrEQP)
'                                                '-- ?˻??ڵ嵵 ???ٸ?...
'                                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'                                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'                                                    strTestName = gArrEQP(intTestCdCnt, 5)
'                                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'                                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'                                                    '-- ????????ġ?? ?⺻???? ?Ѵ?
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
                                        '-- ?˻??ڵ嵵 ???ٸ?...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            '-- ????????ġ?? ?⺻???? ?Ѵ?
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

                            '-- ????Row ?߰?
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If

                            '-- ?Ҽ??? ó??
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
                            
                            '--- ????????
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

                            '-- ???????? ǥ??("????")
                            SetText .spdOrder, "????", gRow, colSTATE

                            '-- ????ȭ?? ?????? ǥ??
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next

                            '-- ???? List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '????
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó???ڵ?
                            '-- ó???? ???????? ?˻??ڵ带 ?????Ѵ?.
                            If strState = "R" Then
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '?˻??ڵ?
                                SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '?˻??ڵ?SUB
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '?˻??ڵ?
                                SetText .spdResult, "", intRstRow, colRSUBCD                    '?˻??ڵ?SUB
                            End If
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '?˻???
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '????ä??
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '????????
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS????
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '????
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '????ġ
                            '-- ???????? ??ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '????????
                            '-- H/L ????ǥ??
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
                            
                            '-- ???? ????
                            Call SetLocalDB(gRow, intRstRow, "1", "")

                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop

                    .spdResult.RowHeight(-1) = 15

                    '## DB?? ????????
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ???? ????
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "????????", gRow, colSTATE
                        Else
                            '-- ???? ????
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "?????Ϸ?", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub FileRcvData_CFX96_RV6()
    Dim RS_L            As ADODB.Recordset
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strMachResult   As String   '?????? ????????
    Dim strResult       As String   '?????? ????(????)
    Dim strIntResult    As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    Dim strCheck        As String   '?˻?????üũ
    Dim strSeqno        As String   '?˻?????
    Dim strOrderCode    As String   'ó???ڵ?
    Dim strTestName     As String   '?˻??ڵ?
    Dim strTestCode     As String   '?˻??ڵ?
    Dim strTestCodeSub  As String   '?˻??ڵ?SUB
    Dim intResPrecUse   As Integer  '?Ҽ?????ȯ????
    Dim intResPrec      As Integer  '?Ҽ????ڸ???
    Dim strResType      As String   '?Ҽ?????ȯ????
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '????????
    Dim strPrevRslt     As String   '????????
    
    Dim intRstRow       As String   '???????????? ???? Row
    Dim intCnt          As Integer  '???? Frame ????
    Dim intCol          As Integer  '?????÷? ????
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
        For intR = 3 To spdExcel.MaxRows
            For intC = 1 To .spdExcel.MaxCols
                strIntBase = ""
                strResult = ""
                If intC = 1 Then
                    strBarno = GetText(.spdExcel, intR, intC)
                    strRackNo = GetText(.spdExcel, intR, intC + 2)
                    
                    If strBarno = "" Then
                        strBarno = strRackNo
                    End If
                    
                    If strOldBarno <> strBarno Then
                        '-- ????????
                        With mResult
                            .BarNo = strBarno
                            .RackNo = strRackNo
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            '.SITE = strHospSite
                            'gHOSP.SITE = strHospSite
                        End With
                                    
                        '-- ????ȯ??????
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                ElseIf intC = 6 Or intC = 8 Or intC = 10 Or intC = 12 Or intC = 14 Or intC = 16 Or intC = 18 Then
                    strIntBase = GetText(.spdExcel, 2, intC)
                    strResult = GetText(.spdExcel, intR, intC)
                    
                    If strResult = "-" Then
                        strResult = gCmnt.NEG
                    ElseIf strResult = "+/-" Or strResult = "+-" Or strResult = "-/+" Or strResult = "-+" Or strResult = "-+" Or UCase(strResult) = "TRACE" Then
                        strResult = gCmnt.POS
                    ElseIf strResult = "+" Then
                        strResult = gCmnt.POS
                    ElseIf strResult = "++" Then
                        strResult = gCmnt.POS2
                    ElseIf strResult = "+++" Then
                        strResult = gCmnt.POS3
                    ElseIf strResult = "++++" Then
                        strResult = gCmnt.POS4
                    ElseIf strResult = "+++++" Then
                        strResult = gCmnt.POS5
                    Else
                        strResult = ""
                    End If
                    
                    If strIntBase <> "" And strResult <> "" And strIntBase <> "Result" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,RESPRECUSE,RESPREC " & vbCrLf
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
    
                            '-- ????Row ?߰?
                            strRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < strRstRow Then
                                .spdResult.MaxRows = strRstRow
                            End If
    
                            '?Ҽ??? ó??, ???? ???? ó??
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '???????? ǥ??("????")
                            SetText .spdOrder, "????", gRow, colSTATE
    
                            '?????? ǥ??
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
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
    
                            '-- ???? List
                            SetText .spdResult, strSeqno, strRstRow, colRSEQNO                '????
                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          'ó???ڵ?
                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '?˻??ڵ?
                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '?˻???
                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL            '????ä??
                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT      '????????
                            SetText .spdResult, strResult, strRstRow, colRLISRESULT           'LIS????
                            
                            '--- ????????
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
                            
                            'SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '????
                            'SetText .spdResult, Trim(RS_L.Fields("REFMLOW")) & "~" & Trim(RS_L.Fields("REFMHIGH")), strRstRow, colRREF          '????ġ
                            
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
                            
                            '-- ???? ????
                            SetLocalDB gRow, strRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                
                End If
                
                '## DB?? ????????
                If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)
        
                    If Res = -1 Then
                        '-- ???? ????
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "????????", gRow, colSTATE
                    Else
                        '-- ???? ????
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "?????Ϸ?", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "FileRcvData_CFX96" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ResultProcess(ByVal pBarNo As String, ByVal pIntBase As String, ByVal pResult As String, ByVal pIntResult As String) As Boolean
    Dim RS_L            As ADODB.Recordset
    Dim strOrderCode    As String   'ó???ڵ?
    Dim strTestName     As String   '?˻??ڵ?
    Dim strAbbrName     As String   '?˻??ڵ?
    Dim strSeqno        As String   '?˻?????
    Dim strTestCode     As String   '?˻??ڵ?
    Dim strTestCodeSub  As String   '?˻??ڵ?SUB
    Dim strCheck        As String   '?˻?????üũ
    Dim intResPrecUse   As Integer  '?Ҽ?????ȯ????
    Dim intResPrec      As Integer  '?Ҽ????ڸ???
    Dim strResType      As String   '?Ҽ?????ȯ????
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '????????
    Dim strPrevRslt     As String   '????????
    Dim strIntResult    As String   '?????? ????(????)
    Dim strMachResult   As String   '?????? ????????
    Dim strAMRResult    As String   '?????? ????(????)
    Dim strRstType      As String
    Dim i               As Integer
    Dim intRstRow       As String   '???????????? ???? Row
    Dim intCol          As Integer  '?????÷? ????
    
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
        
        '-- ????ġ
        If mPatient.SEX = "M" Then
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        ElseIf mPatient.SEX = "F" Then
            strLow = Trim(RS_L.Fields("REFFLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
        Else
            '-- ????????ġ?? ?⺻???? ?Ѵ?
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        End If
        
        '-- ?Ҽ?????ȯ ???뿩?ο? ??ȯ?ڸ???
        intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
        intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
                
        '????????
        strResType = Trim(RS_L.Fields("RESTYPE")) & ""
        
        '-- 0:??ġ,1:????,2:??ġ/????
        If strResType = 0 Then
            '-- ?Ҽ??? ó??
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
                
            '-- AMR ???? (??ġ??)
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
            '-- AMR ???? (??????)
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
            '-- ?Ҽ??? ó??
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
                
            '-- AMR ???? (??ġ??)
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
                                    
            '-- AMR ???? (??????)
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
        
        '??ġ???? ????
        '0:????????, 1:????(????), 2:????(????), 3:????_????, 4:????_????
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
    
        '--- ????????
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
            '-- ????Row ?߰?
            intRstRow = .spdResult.DataRowCnt + 1
            If .spdResult.MaxRows < intRstRow Then
                .spdResult.MaxRows = intRstRow
            End If
    
            '-- ???????? ǥ??("????")
            SetText .spdOrder, "????", gRow, colSTATE
    
            '-- ????ȭ?? ?????? ǥ??
            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                If strAbbrName = gArrEQP(intCol - colSTATE, 6) Then
                    SetText .spdOrder, pResult, gRow, intCol
                    
                    '-- H/L ????ǥ??
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
    
            '-- ???? List
            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '????
            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó???ڵ?
            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '?˻??ڵ?
            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '?˻??ڵ?SUB
            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '?˻???
            SetText .spdResult, pIntBase, intRstRow, colRCHANNEL              '????ä??
            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '????????
            SetText .spdResult, pResult, intRstRow, colRLISRESULT             'LIS????
            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '????
            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '????ġ
            
            '-- ???????? ??ȸ
            strPrevRslt = GetPrevResult(mResult.BarNo, pIntBase, strTestCode)
            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '????????
            
            '-- H/L ????ǥ??
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
            
            '-- ????Count
            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                SetText .spdOrder, "1", gRow, colRCNT
            Else
                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
            End If
        End With
        
        '-- ???? ????
        Call SetLocalDB(gRow, intRstRow, "1", "")
        
        ResultProcess = True
    
    End If
    
End Function


Private Sub SerialRcvData_HITACHI7180()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strMachResult   As String   '?????? ????????
    Dim strResult       As String   '?????? ????(????)
    Dim strIntResult    As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    Dim strCheck        As String   '?˻?????üũ
    Dim strSeqno        As String   '?˻?????
    Dim strOrderCode    As String   'ó???ڵ?
    Dim strTestName     As String   '?˻??ڵ?
    Dim strTestCode     As String   '?˻??ڵ?
    Dim strTestCodeSub  As String   '?˻??ڵ?SUB
    Dim intResPrecUse   As Integer  '?Ҽ?????ȯ????
    Dim intResPrec      As Integer  '?Ҽ????ڸ???
    Dim strResType      As String   '?Ҽ?????ȯ????
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '????????
    Dim strPrevRslt     As String   '????????
    
    Dim intRstRow       As String   '???????????? ???? Row
    Dim intCnt          As Integer  '???? Frame ????
    Dim intCol          As Integer  '?????÷? ????
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
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
                Case ">", "?", "@"      'ANY ????
                    
                    '-- ?????? ????
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ?????? ????
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
                    '-- ????????
                    With mOrder
                        .Function = strFunction
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- ????ȯ??????
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                    '## Control, Calibration ?????ʹ? ??????
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
                        '-- ?????? ????
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    '-- ????ȯ??????
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
                            
                        '-- ?˻縶???? ???? ????????
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- ?˻縶???? ???? ????????
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ????ä???? ????...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
'                                    '-- ȯ?? ó?????? ????????
'                                    If UBound(gPatTest) > 0 Then
'                                        For intOrdCnt = 1 To UBound(gPatTest)
'                                            For intTestCdCnt = 1 To UBound(gArrEQP)
'                                                '-- ?˻??ڵ嵵 ???ٸ?...
'                                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'                                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'                                                    strTestName = gArrEQP(intTestCdCnt, 5)
'                                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'                                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'                                                    '-- ????????ġ?? ?⺻???? ?Ѵ?
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
                                        '-- ?˻??ڵ嵵 ???ٸ?...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            '-- ????????ġ?? ?⺻???? ?Ѵ?
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

                            '-- ????Row ?߰?
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If

                            '-- ?Ҽ??? ó??
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
                            
                            '--- ????????
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

                            '-- ???????? ǥ??("????")
                            SetText .spdOrder, "????", gRow, colSTATE

                            '-- ????ȭ?? ?????? ǥ??
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next

                            '-- ???? List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '????
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó???ڵ?
                            '-- ó???? ???????? ?˻??ڵ带 ?????Ѵ?.
                            If strState = "R" Then
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '?˻??ڵ?
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '?˻??ڵ?
                            End If
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '?˻???
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '????ä??
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '????????
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS????
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '????
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '????ġ
                            '-- ???????? ??ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '????????
                            '-- H/L ????ǥ??
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
                            
                            '-- ???? ????
                            Call SetLocalDB(gRow, intRstRow, "1", "")

                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        strTmp = Mid$(strTmp, 11)
                    Loop

                    .spdResult.RowHeight(-1) = 15

                    '## DB?? ????????
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ???? ????
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "????????", gRow, colSTATE
                        Else
                            '-- ???? ????
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "?????Ϸ?", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
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
        Call FileRcvData_CFX96_RV6
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
    Me.Caption = gHOSP.MACHNM & Space$(5) & "?¢¢¢¢?     [?????? ???????̽?]     ?¢¢¢¢?"

    Call CtlInitializing

    Call frmClear
    
    '-- Menu Set
    Call SetMenu

    '-- ?÷????̱⼳??
    Call SetColumnView(spdOrder)

    '-- ?˻??ڵ?
    Call GetTestList

    Call GetTestListName

    '-- ?˻??? ???̱?
    Call SetExamCode(spdOrder)

    '-- ???ſ???
    Call OpenCommunication

    pDel = False

    spdComStatus.MaxRows = 0
    spdComStatus.Font.Bold = True
    
    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    txtTestID.Text = gHOSP.USERID
    txtTestNm.Text = gHOSP.USERNM
    lblPatInfo.Caption = ""
    lblAPI.Caption = "???Ἥ??: " & gHOSP.APIURL
    
    dtpFrom.Value = Now
    dtpTo.Value = Now

    imgNet1.ZOrder 0
    tmrDBConn.Interval = 500
    tmrDBConn.Enabled = True
    
    '-- ???????? ????
    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gHOSP.SAVEDAY + "???? ????Ÿ?? ?????Ͻðڽ??ϱ??", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("??Ʈ ??ȣ?? ?߸??Ǿ????ϴ?." & vbNewLine & vbNewLine & "   ???? ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & " ????????"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            
            Resume Next
        Else
            End
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
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
            lblComStatus.Caption = "COM" & comEqp.CommPort & " ???Ἲ??"
            
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & " ????????"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    ElseIf gComm.COMTYPE = "2" Then
'        If gComm.TCPTYPE = "1" Then
'            wSck.LocalPort = CInt(gComm.TCPPORT)
'            wSck.Listen
'
'            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ????..."
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
'            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ????..."
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
        lblComStatus.Caption = "???? ????: " & gComm.RSTPATH

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

    
End Sub

'???????̽? ȯ?ڼ??ý? ?????? ?˻??׸?/?????????ֱ?
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
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE = '" & strBarno & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ??????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
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
    
    frmWorkList.Show vbModal

End Sub

Public Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    '-- ????
'    If Row = 0 Then
'        '-- ???? ?߰?
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
    
    '-- ȯ??????ǥ??
    shpPatInfo.Visible = True
    
    strPatInfo = ""
    strPatInfo = strPatInfo & "????    ??: " & GetText(spdOrder, Row, colPNAME)
    If GetText(spdOrder, Row, colPSEX) <> "" Then
        strPatInfo = strPatInfo & " [" & GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE) & "] " & vbCrLf
    Else
        strPatInfo = strPatInfo & vbCrLf
    End If
    strPatInfo = strPatInfo & "?°?ü??ȣ: " & GetText(spdOrder, Row, colBARCODE) & vbCrLf
    strPatInfo = strPatInfo & "??ȯ?ڹ?ȣ: " & GetText(spdOrder, Row, colPID) & vbCrLf
    
    lblPatInfo.Caption = strPatInfo
    
    txtBarcode.Text = GetText(spdOrder, Row, colBARCODE)
    txtPatID.Text = GetText(spdOrder, Row, colPID)
    txtPName.Text = GetText(spdOrder, Row, colPNAME)
    txtSA.Text = GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE)
    
    '-- ????ǥ??
    If GetPatTRestResult(Row) = -1 Then
        '?????????? ???????? ?˻????? ?????ֱ?
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '??
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
                MsgBox "?Է??? ???ڵ忡?? ȯ???????? ã?? ???߽??ϴ?." & vbNewLine & " ???ڵ? ??ȣ?? Ȯ???ϼ???", vbOKOnly + vbCritical, Me.Caption
            Else
                '????????
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
                    '-- ????
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "???ڸ? ?Է??? ?????մϴ?"
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
        
        If MsgBox(strNewBarno & " ?? ?????ðڽ??ϱ??", vbInformation + vbYesNo, "?˸?") = vbNo Then
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
    
    '-- ????
    If Row = 0 Then
        '-- ???? ?߰?
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
                            Call SetText(spdOrder, "??", intRow, intOCol)
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
                    MsgBox "?Է??? ???ڵ忡?? ȯ???????? ã?? ???߽??ϴ?." & vbNewLine & " ???ڵ? ??ȣ?? Ȯ???ϼ???", vbOKOnly + vbCritical, Me.Caption
                    txtBarNum.SelStart = 0
                    txtBarNum.SelLength = Len(txtBarNum.Text)
                    Exit Sub
                    
                Else
                    '????????
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
'                        '-- ????
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
                    MsgBox "?Է??? ???ڵ忡?? ȯ???????? ã?? ???߽??ϴ?." & vbNewLine & " ???ڵ? ??ȣ?? Ȯ???ϼ???", vbOKOnly + vbCritical, Me.Caption
                    txtBarNum.SelStart = 0
                    txtBarNum.SelLength = Len(txtBarNum.Text)
                    Exit Sub
                Else
                    '????????
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
                        '-- ????
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
