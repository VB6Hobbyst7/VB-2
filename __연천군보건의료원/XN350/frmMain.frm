VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   21810
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
   ScaleHeight     =   9180
   ScaleWidth      =   21810
   WindowState     =   2  '?ִ?ȭ
   Begin FPSpreadADO.fpSpread spdWork1 
      Height          =   7575
      Left            =   10110
      TabIndex        =   48
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      _Version        =   524288
      _ExtentX        =   10398
      _ExtentY        =   13361
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   23
      OperationMode   =   2
      SpreadDesigner  =   "frmMain.frx":0E42
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      TabIndex        =   35
      Top             =   570
      Width           =   5895
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   690
         TabIndex        =   36
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   127270913
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   127270913
         CurrentDate     =   40457
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   375
         Left            =   3750
         TabIndex        =   58
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "??ȸ"
         CaptionChecked  =   "??ȸ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMain.frx":1C2F
         BackColor       =   12640511
         AlphaColor      =   12640511
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAll 
         Height          =   375
         Left            =   4770
         TabIndex        =   60
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "????"
         CaptionChecked  =   "????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         AlphaColor      =   12640511
         ImgOutLineSize  =   3
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
         Left            =   2100
         TabIndex        =   39
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '????
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "??ȸ?Ⱓ"
         BeginProperty Font 
            Name            =   "????"
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
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   2745
      Left            =   12270
      TabIndex        =   11
      Top             =   4470
      Visible         =   0   'False
      Width           =   6345
      Begin VB.TextBox txtLastSeq 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BeginProperty Font 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5220
         TabIndex        =   69
         Text            =   "0"
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdAll1 
         BackColor       =   &H00A7FAEB&
         Caption         =   "????"
         Height          =   375
         Left            =   120
         Style           =   1  '?׷???
         TabIndex        =   61
         Top             =   1140
         Width           =   975
      End
      Begin VB.CommandButton cmdMatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   1110
         Style           =   1  '?׷???
         TabIndex        =   59
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch1 
         BackColor       =   &H00A7FAEB&
         Caption         =   "??ȸ"
         Height          =   375
         Left            =   60
         Style           =   1  '?׷???
         TabIndex        =   57
         Top             =   630
         Width           =   1005
      End
      Begin VB.CommandButton cmdTestNmSave1 
         BackColor       =   &H00FFC0FF&
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
         Height          =   285
         Left            =   4110
         Style           =   1  '?׷???
         TabIndex        =   55
         Top             =   2010
         Width           =   615
      End
      Begin VB.CommandButton cmdClear1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ȭ??????"
         Height          =   375
         Left            =   2520
         Style           =   1  '?׷???
         TabIndex        =   54
         ToolTipText     =   "????ȭ???? ???? ?????ϴ?"
         Top             =   1950
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "????????"
         Height          =   375
         Left            =   1020
         Style           =   1  '?׷???
         TabIndex        =   53
         ToolTipText     =   "?????? ?????? EMR?????? ?????մϴ?"
         Top             =   1980
         Width           =   1395
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "??ũ??ȸ"
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "??????ȸ"
         Height          =   315
         Left            =   2610
         TabIndex        =   12
         Top             =   270
         Width           =   1425
      End
      Begin BHButton.BHImageButton cmdHide 
         Height          =   375
         Left            =   4170
         TabIndex        =   64
         Top             =   990
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "?? ????â ??????"
         CaptionChecked  =   "?? ????â ??????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSelected =   "frmMain.frx":1D89
         ForeColor       =   49152
         ImgOutLineSize  =   3
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
         TabIndex        =   15
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
      ScaleWidth      =   21810
      TabIndex        =   4
      Top             =   8595
      Width           =   21810
      Begin VB.ListBox lstComStatus 
         Height          =   420
         Left            =   15060
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   4785
      End
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
               Picture         =   "frmMain.frx":1FE5
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":257F
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B19
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":30B3
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3945
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3A9F
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3BF9
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3D53
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":462D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgOff 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":4F07
         Top             =   30
         Width           =   480
      End
      Begin VB.Image imgOn 
         Height          =   480
         Left            =   3270
         Picture         =   "frmMain.frx":57D1
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
         Left            =   11100
         TabIndex        =   49
         Top             =   150
         Width           =   3465
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   10920
         Top             =   90
         Width           =   3645
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   14580
         Top             =   90
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3810
         Picture         =   "frmMain.frx":609B
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   5805
         Picture         =   "frmMain.frx":6625
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   7080
         Picture         =   "frmMain.frx":6BAF
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "???ſ???"
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
         Left            =   3930
         TabIndex        =   9
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "?޴½?ȣ"
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
         Left            =   4995
         TabIndex        =   8
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "?????½?ȣ"
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
         Left            =   6120
         TabIndex        =   7
         Top             =   210
         Width           =   900
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":7139
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":7283
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":73CD
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         Alignment       =   2  '??? ????
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
   Begin VB.PictureBox picComm 
      Align           =   2  '?Ʒ? ????
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   21750
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   21810
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
   Begin FPSpreadADO.fpSpread spdResult 
      Height          =   7185
      Left            =   16440
      TabIndex        =   47
      Top             =   1380
      Width           =   6525
      _Version        =   524288
      _ExtentX        =   11509
      _ExtentY        =   12674
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   13
      MaxRows         =   10
      SpreadDesigner  =   "frmMain.frx":7517
   End
   Begin FPSpreadADO.fpSpread spdOrder 
      Height          =   7905
      Left            =   6090
      TabIndex        =   45
      Top             =   780
      Width           =   7845
      _Version        =   524288
      _ExtentX        =   13838
      _ExtentY        =   13944
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      SpreadDesigner  =   "frmMain.frx":7E4B
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '?? ????
      BackColor       =   &H00A5704B&
      BorderStyle     =   0  '????
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   21810
      TabIndex        =   0
      Top             =   0
      Width           =   21810
      Begin VB.Frame fraVision 
         BackColor       =   &H00ACFFEF&
         BorderStyle     =   0  '????
         Height          =   345
         Left            =   3780
         TabIndex        =   65
         Top             =   60
         Width           =   2145
         Begin VB.TextBox txtRCnt 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            TabIndex        =   66
            Text            =   "1"
            Top             =   30
            Width           =   495
         End
         Begin BHButton.BHImageButton cmdGetRslt 
            Height          =   300
            Left            =   1320
            TabIndex        =   68
            Top             =   30
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            Caption         =   "?ޱ?"
            CaptionChecked  =   "?????ޱ?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12640511
            AlphaColor      =   12640511
            ImgOutLineSize  =   3
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '????
            Caption         =   "??????"
            BeginProperty Font 
               Name            =   "????"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   67
            Top             =   90
            Width           =   765
         End
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   14640
         TabIndex        =   51
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "????????"
         CaptionChecked  =   "????????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMain.frx":8C22
         PictureSelected =   "frmMain.frx":8E7E
         BackColor       =   12640511
         AlphaColor      =   12640511
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   16140
         TabIndex        =   52
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "ȭ??????"
         CaptionChecked  =   "ȭ??????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMain.frx":90DA
         PictureSelected =   "frmMain.frx":9246
         ImgOutLineSize  =   3
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   4620
         TabIndex        =   46
         Top             =   150
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12420
         TabIndex        =   44
         Text            =   "1234567890"
         Top             =   90
         Width           =   1335
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   2880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   2400
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9930
         TabIndex        =   10
         Text            =   "1234567890"
         Top             =   90
         Width           =   1335
      End
      Begin BHButton.BHImageButton cmdTestNmSave 
         Height          =   300
         Left            =   13830
         TabIndex        =   56
         Top             =   90
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "????"
         CaptionChecked  =   "????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         AlphaColor      =   12640511
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdView 
         Height          =   375
         Left            =   17640
         TabIndex        =   63
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "?? ????â ??????"
         CaptionChecked  =   "?? ????â ??????"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSelected =   "frmMain.frx":94A2
         ForeColor       =   49152
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblHospInfo 
         Appearance      =   0  '????
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "???????б????? HITACHI 7020"
         BeginProperty Font 
            Name            =   "????"
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
         Appearance      =   0  '????
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "???????б????? HITACHI 7020"
         BeginProperty Font 
            Name            =   "????"
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
         BackStyle       =   0  '????
         Caption         =   "?˻??ڸ? : "
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11430
         TabIndex        =   40
         Top             =   120
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7320
         TabIndex        =   3
         Top             =   105
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '????
         Caption         =   "?˻???ID : "
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8940
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '???????? ????
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  '????
         Height          =   375
         Left            =   8760
         Top             =   60
         Width           =   5835
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '????
         Caption         =   "?˻????? :"
         BeginProperty Font 
            Name            =   "???? ????"
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
         Left            =   6240
         TabIndex        =   1
         Top             =   120
         Width           =   945
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '???????? ????
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '????
         Height          =   375
         Left            =   6030
         Top             =   60
         Width           =   2685
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '???????? ????
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '????
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "Sex/Age"
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "??      ??"
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '????
         Caption         =   "???Ϲ?ȣ"
         BeginProperty Font 
            Name            =   "????"
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
         Alignment       =   2  '??? ????
         Appearance      =   0  '????
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "??ü??ȣ"
         BeginProperty Font 
            Name            =   "????"
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
   Begin FPSpreadADO.fpSpread spdWork 
      Height          =   7905
      Left            =   60
      TabIndex        =   62
      Top             =   1290
      Width           =   5895
      _Version        =   524288
      _ExtentX        =   10398
      _ExtentY        =   13944
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "???? ????"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   23
      SpreadDesigner  =   "frmMain.frx":96FE
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ??ȸ ????  "
      Begin VB.Menu mnuResult 
         Caption         =   "?? ???? ??ȸ"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "?? ??ũ ??ȸ"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "????"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "???????̽? ???? "
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
      Caption         =   " ?˻? ?ɼ? "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "?? ???ڵ? ????"
         WindowList      =   -1  'True
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
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "?? ???? ????"
      End
      Begin VB.Menu mnuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "?? EMR ????"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " ??Ÿ ?޴? "
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


Private Sub cmdGetRslt_Click()
    Dim strSendData As String
    Dim strFirstSeq  As String
    Dim strLastSeq  As String
'    Dim db_tmp As String * 100
    
    
    strFirstSeq = txtLastSeq.Text
    strFirstSeq = (strFirstSeq - 1) - (txtRCnt.Text - 1)
    
    strLastSeq = strFirstSeq + (txtRCnt.Text - 1)
    
    strSendData = "0" & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq & vbLf
    
    wSck.SendData strSendData
    SetRawData "[Tx]" & strSendData
    
End Sub

Private Sub cmdHide_Click()
    
    spdResult.Visible = False
    fraPatInfo.Visible = False
    
    spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH + 500

End Sub

Private Sub cmdMatch_Click()
    Dim intWRow     As Integer
    Dim intORow     As Integer
    Dim blnSame     As Boolean
    Dim i           As Integer
    Dim intCnt      As Integer
    
    blnSame = False
    intCnt = 0
    
    For intWRow = 1 To spdWork.MaxRows
        If GetText(spdWork, intORow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
        End If
    Next
    
    If intCnt > 1 Then
        MsgBox "?? ?????? ?????ϼ???", vbOKOnly, Me.Caption
        Exit Sub
    End If
       
    For intORow = 1 To spdOrder.MaxRows
        If GetText(spdOrder, intORow, colCHECKBOX) = "1" Then
            blnSame = True
            Exit For
        End If
    Next
    
    If blnSame = False Then
        MsgBox "???? ??ü?? ?????ϼ???", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    
    If blnSame = True Then
        For i = colCHECKBOX To colSTATE
            Call SetText(spdOrder, GetText(spdWork, intWRow, i), intORow, i)
            
'            varItems = GetText(spdWork, intWRow, colITEMS)
'            varItems = Split(varItems, "/")
'            For intItems = 0 To UBound(varItems)
'                For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
'                    .Row = 0
'                    .Col = intOCol
'                    If varItems(intItems) = Trim(.Text) Then
'                        .Row = intRow
'                        Call SetText(spdOrder, "??", intRow, intOCol)
'                    End If
'                Next
'            Next
        Next
        
        '????????
        SQL = ""
        SQL = SQL & "UPDATE PATRESULT SET "
        SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, intORow, colBARCODE)) & "'" & vbCrLf
        SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, intORow, colPID)) & "'" & vbCrLf
        SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, intORow, colCHARTNO)) & "'" & vbCrLf
        SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, intORow, colSPECIMEN)) & "'" & vbCrLf
        SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, intORow, colDEPT)) & "'" & vbCrLf
        SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, intORow, colINOUT)) & "'" & vbCrLf
        SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, intORow, colER)) & "'" & vbCrLf
        SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, intORow, colRT)) & "'" & vbCrLf
        SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, intORow, colPNAME)) & "'" & vbCrLf
        SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, intORow, colPSEX)) & "'" & vbCrLf
        SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, intORow, colPAGE)) & "'" & vbCrLf
        SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, intORow, colRACKNO)) & "'" & vbCrLf
        SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, intORow, colPOSNO)) & "'" & vbCrLf
        SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
        SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, intORow, colEXAMDATE)) & "'" & vbCrLf
        SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, intORow, colEXAMTIME)) & "'" & vbCrLf
        SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intORow, colSAVESEQ)) & vbCrLf
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- ????
        End If
    End If
End Sub

Private Sub cmdRcv_Click()
        
    pBuffer = txtRcv.Text
    
    If LEFT(pBuffer, 1) <> STX Then
        pBuffer = STX & pBuffer
    End If
    If Right(pBuffer, 1) <> ETX Then
        pBuffer = pBuffer & ETX
    End If
    
    Select Case UCase(gHOSP.MACHNM)
        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
        Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
        Case "XP300":           Call Phase_Serial_XP300
        Case "KLITE":           Call Phase_TCP_KLITE
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdRcvView_Click()

    frmLogView.Show
    
End Sub

Private Sub cmdResult_Click()

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
                .Col = colCHECKBOX
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
                    spdOrder.Col = colCHECKBOX
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

Private Sub cmdTestNmSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdView_Click()
    
    If cmdView.Caption = "????â ???̱? ??" Then
        cmdView.Caption = "?? ????â ??????"
        cmdView.ForeColor = &HC000&
        
        spdResult.Visible = True
        fraPatInfo.Visible = True
            
        spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
        spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
        spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
        
        fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
    Else
        cmdView.Caption = "????â ???̱? ??"
        cmdView.ForeColor = &HFF00FF
        
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
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                        Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
                        Case "UROMETER720":     Call Phase_Serial_UROMETER720
                        Case "HORIBA":          Call Phase_Serial_HORIBA
                        Case "XP300":           Call Phase_Serial_XP300
                        Case "ISMART30":        Call Phase_Serial_ISMART30
                        Case "STAGO":           Call Phase_Serial_STAGO
                        
                        Case "XN350":           Call Phase_Serial_XN350
'                        Case "KLITE":           Call Phase_Serial_KLITE
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

Private Sub Command1_Click()

    spdOrder.MaxRows = spdOrder.MaxRows + 1
    spdOrder.RowHeight(-1) = 14
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

    ''Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    ''Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
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
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        
        '???ڵ带 ???????? ???? ???쿡 ?????Ѵ?.
        If gHOSP.BARUSE <> "Y" Then
            mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        End If
        
        '-- ?˻?ä?η? ???????? ??????
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
            
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- ???? Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_XN350(ByVal pBarNo As String, ByVal pType As String)

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

    ''Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    ''Call SetCommStatus("Q", pBarNo, frmMain.lstComStatus)
    
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
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = ""
        strItems = GetEquipExamCode_XN350(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- ?˻?ä?η? ???????? ??????
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If
        
        '-- ???? Row
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
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarNo, intRow)
        
        '???ڵ带 ???????? ???? ???쿡 ?????Ѵ?.
        If gHOSP.BARUSE <> "Y" Then
            mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        End If
        
        '-- ?˻?ä?η? ???????? ??????
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
            
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- ???? Row
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
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
        strItems = ""
        strItems = GetEquipExamCode_STAGO(gHOSP.MACHCD, pBarNo, intRow)
        
        
        '-- ?˻?ä?η? ???????? ??????
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- ????????(Order) ǥ??
            Call SetText(frmMain.spdOrder, "????????", intRow, colSTATE)
        End If

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
    DoEvents
    
    '-- ?αױ???
    Call SetRawData("[Tx]" & pSendData)

    '-- ????ǥ??
    ''Call SetCommStatus("S", pSendData, spdComStatus)
    'Call SetCommStatus("S", pSendData, lstComStatus)

End Sub


Private Sub TCPRcvData_KLITE()
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
    
    '?????? ????
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
                        '-- ????????
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ????ȯ??????
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- ???????̽? ????
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
                    
                    '-- ?˻縶???? ???? ????????
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
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
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
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
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
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
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
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub TCPRcvData_F200()
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
    
    '?????? ????
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    
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
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- ????????
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ????ȯ??????
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- ???????̽? ????
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||ASCII|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|TR03-025|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- ?˻縶???? ???? ????????
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
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
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
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
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
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
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
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HITACHI7180()
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
    Dim strAMRResult    As String   '?????? ????(????)
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
    Dim strAbbrName     As String   '?˻??ڵ?
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
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
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
                Case ">", "?", "@"      'ANY ????
                    
                    '-- ?????? ????
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ?????? ????
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    If gHOSP.BARUSE = "Y" Then
                        '???ڵ? ????
                        sFunc = Mid(strRcvBuf, 2, 40)
                    Else
                        '???ڵ? ?̻???
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

                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration ?????ʹ? ??????
                    If UCase(strFunc) = "K" Or UCase(strFunc) = "L" Or UCase(strFunc) = "G" Or UCase(strFunc) = "H" Then
                        '-- ?????? ????
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    If strFunc <> "@" And strFunc <> "M" Then
                        '## QC
                        If UCase(strFunc) = "F" Then
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
                        
                        strTC = ""
                        strTG = ""
                        strHDL = ""
                        
                        For ii = 44 To Len(strRcvBuf) Step 10
                            strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                            strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                            strComm = Trim(Mid(strRcvBuf, ii + 9, 1))
                        
                            If strIntBase = "6" Then    'TCHO
                                strTC = strResult
                            End If

                            If strIntBase = "13" Then   'TG
                                strTG = strResult
                            End If

                            If strIntBase = "11" Then    'HDLC
                                strHDL = strResult
                            End If
                        
ReCal:
                            strSeqno = ""
                            strTestCode = ""
                            strTestName = ""
                            strAbbrName = ""
                            intResPrecUse = -1
                            intResPrec = -1
                            strAMRResult = ""
                            
                            '-- ?˻縶???? ???? ????????
                            If strIntBase <> "" And strResult <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
                                SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
                                SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
                                SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
                                SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
                                SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                                      " & vbCrLf
                                If gPatOrdCd <> "" Then
                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                End If
                                
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
                                        
                                    '-- AMR ???? (??ġ??)
                                    If IsNumeric(strResult) Then
                                        If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                                            If CCur(strResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                                                strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                                            End If
                                        End If
                                        If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                                            If CCur(strResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                                                strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                                            End If
                                        End If
                                        If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                                            If CCur(strResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                                                strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                                            End If
                                        End If
                                        If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                                            If CCur(strResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                                                strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                                            End If
                                        End If
                                    End If
                                                            
                                    '-- AMR ???? (??????)
                                    If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                                        If strResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                                        End If
                                    End If
                                    If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                                        If strResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                                        End If
                                    End If
                                    If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                                        If strResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                                        End If
                                    End If
                                    
                                    If strAMRResult <> "" Then
                                        If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                                            If strIntResult <> "" Then
                                                strResult = strAMRResult & "(" & strIntResult & ")"
                                            Else
                                                strResult = strAMRResult
                                            End If
                                        Else
                                            strResult = strAMRResult
                                        End If
                                    End If
                                
                                    '--- ????????
                                    strJudge = ""
                                    If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                        If IsNumeric(strResult) Then
                                            If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                                strJudge = ""
                                            ElseIf CCur(strHigh) <= CCur(strResult) Then
                                                strJudge = "H"
                                            ElseIf CCur(strLow) >= CCur(strResult) Then
                                                strJudge = "L"
                                            End If
                                        End If
                                    End If
                                    
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
                                            SetText .spdOrder, strResult, gRow, intCol
                                            
                                            strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                            
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
                                    
                                    '-- ????Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                    
                                    '-- ???? ????
                                    Call SetLocalDB(gRow, intRstRow, "1", "")
                                    
                                    strState = "R"
                                    
                                End If
                            End If
                        Next
                        
                        'LDL ????
                        If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                            strIntBase = "99"
                            strResult = strTC - ((strTG / 5) + strHDL)
                            If strResult < 0 Then
                                strResult = "0"
                            End If
                            strTC = ""
                            strTG = ""
                            strHDL = ""
                            GoTo ReCal
                        End If
                        
                        Call SendData(SndMore)
                        
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
    
                                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                                If DBExec(AdoCn_Local, SQL) Then
                                    '-- ????
                                End If
                            End If
                            strState = ""
                        End If
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "SerialRcvData_H7180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_HITACHI7020()
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
    Dim strAMRResult    As String   '?????? ????(????)
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
    Dim strAbbrName     As String   '?˻??ڵ?
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
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
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
                Case ">", "?", "@"      'ANY ????
                    
                    '-- ?????? ????
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ?????? ????
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
                    strFunc = Mid(strRcvBuf, 2, 1)              ' Function

                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                         Exit Sub
                    End If

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    If gHOSP.BARUSE = "Y" Then
                        '???ڵ? ????
                        sFunc = Mid(strRcvBuf, 2, 40)
                    Else
                        '???ڵ? ?̻???
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

                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration ?????ʹ? ??????
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- ?????? ????
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    '## QC
                    If UCase(strFunc) = "F" Then
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

                    For ii = 51 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                    
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        strAbbrName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                        strAMRResult = ""
                        
                        '-- ?˻縶???? ???? ????????
                        If strIntBase <> "" And strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
                            SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
                            SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
                            SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
                            SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
                            SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                                      " & vbCrLf
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            End If
                            
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
                                    
                                '-- AMR ???? (??ġ??)
                                If IsNumeric(strResult) Then
                                    If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                                        If CCur(strResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                                        End If
                                    End If
                                    If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                                        If CCur(strResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                                        End If
                                    End If
                                    If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                                        If CCur(strResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                                        End If
                                    End If
                                    If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                                        If CCur(strResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                                            strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                                        End If
                                    End If
                                End If
                                                        
                                '-- AMR ???? (??????)
                                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                                    If strResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                                    If strResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                                    If strResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                                    End If
                                End If
                                
                                If strAMRResult <> "" Then
                                    If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                                        If strIntResult <> "" Then
                                            strResult = strAMRResult & "(" & strIntResult & ")"
                                        Else
                                            strResult = strAMRResult
                                        End If
                                    Else
                                        strResult = strAMRResult
                                    End If
                                End If
                            
                                '--- ????????
                                strJudge = ""
                                If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                    If IsNumeric(strResult) Then
                                        If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                            strJudge = ""
                                        ElseIf CCur(strHigh) <= CCur(strResult) Then
                                            strJudge = "H"
                                        ElseIf CCur(strLow) >= CCur(strResult) Then
                                            strJudge = "L"
                                        End If
                                    End If
                                End If
                                
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
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                        
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
                                
                                '-- ????Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                                '-- ???? ????
                                Call SetLocalDB(gRow, intRstRow, "1", "")
                                
                                strState = "R"
                                
                            End If
                        End If
                    Next
                    
                    Call SendData(SndMore)
                    
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_H7020" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ResultProcess(ByVal pBarNo As String, ByVal pIntBase As String, ByVal pResult As String) As Boolean
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
    SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
    SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
    SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
    SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL = '" & pIntBase & "'                                                      " & vbCrLf
    If gPatOrdCd <> "" Then
        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
    End If
    
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
        
        '-- ?Ҽ??? ó??
        strMachResult = pResult
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
            pResult = Format(pResult, strResType)
        End If
            
        '-- AMR ???? (??ġ??)
        If IsNumeric(pResult) Then
            If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                If CCur(pResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                    strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                End If
            End If
            If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                If CCur(pResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                    strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                End If
            End If
            If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                If CCur(pResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                    strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                End If
            End If
            If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                If CCur(pResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                    strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                End If
            End If
        End If
                                
        '-- AMR ???? (??????)
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
        
        If strAMRResult <> "" Then
            If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                If strIntResult <> "" Then
                    pResult = strAMRResult & "(" & strIntResult & ")"
                Else
                    pResult = strAMRResult
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
                    
                    If Trim(gArrEQP(intCol - colSTATE, 16)) <> "" Then
                        'PLIS:WA
                        strOrderCode = gArrEQP(intCol - colSTATE, 16)
                    End If
                    If Trim(gArrEQP(intCol - colSTATE, 17)) <> "" Then
                        'PLIS:ACCSEQ
                        strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
                    End If
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

Private Sub SerialRcvData_XP300()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strResult       As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    
    Dim intCnt          As Integer  '???? Frame ????
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
    
    '?????? ????
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
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = Trim(mGetP(strTemp1, 1, "^"))
                    strTubePos = Trim(mGetP(strTemp1, 2, "^"))
                    strBarno = Trim(mGetP(strRcvBuf, 3, "^"))

                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ????ȯ??????
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
                        '## ???????? ????
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ???????? ????
                        strResult = strTemp2
                    End If
                        
                    '-- ?˻?????ó?? ???μ???
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB?? ????????
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_XP300" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XN350()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strOrgBarno     As String   '?????? ???ڵ???ȣ
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strResult       As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    
    Dim intCnt          As Integer  '???? Frame ????
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
    
    '?????? ????
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
                    strOrgBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    'strBarno = Mid(strOrgBarno, 1, 11)
                    strBarno = Mid(strOrgBarno, 1, gHOSP.BARLEN)
                    
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .OrgBarNo = strOrgBarno
                    End With
                    
                    Call GetOrder_XN350(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = Trim(mGetP(strTemp1, 1, "^"))
                    strTubePos = Trim(mGetP(strTemp1, 2, "^"))
                    strOrgBarno = Trim(mGetP(strRcvBuf, 3, "^"))
                    strBarno = Mid(strOrgBarno, 1, gHOSP.BARLEN)
                    
                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ????ȯ??????
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
                        '## ???????? ????
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ???????? ????
                        strResult = strTemp2
                    End If
                        
                    '-- ?˻?????ó?? ???μ???
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB?? ????????
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_XN350" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_VISION()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strResult       As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    
    Dim intCnt          As Integer  '???? Frame ????
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
    
    '?????? ????
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
                '-- 18?? ????
                strResult = mGetP(strRcvBuf, 10, vbTab)
                'strResult = mGetP(strRcvBuf, 11, vbTab)

                '-- ????????
                With mResult
                    .BarNo = strBarno
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
                
                strState = "O"
                        
                '-- ?˻?????ó?? ???μ???
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                .spdResult.RowHeight(-1) = 15

                '## DB?? ????????
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_TCPRcvData_KLITE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ISMART30()
    '???? ???? ????
    Dim strRcvBuf       As String   '?????? Data
    Dim strType         As String   '?????? Record Type
    Dim strBarno        As String   '?????? ???ڵ???ȣ
    Dim strSeq          As String   '?????? Sequence
    Dim strRackNo       As String   '?????? Rack Or Disk No
    Dim strTubePos      As String   '?????? Tube Position
    Dim strIntBase      As String   '?????? ???????? ?˻???
    Dim strResult       As String   '?????? ????(????)
    Dim strQCResult     As String   '?????? ????(QC)
    Dim strFlag         As String   '?????? Abnormal Flag
    Dim strComm         As String   '?????? Comment
    
    '?????? ????
    
    Dim intCnt          As Integer  '???? Frame ????
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
    
    '?????? ????
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
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = Trim(mGetP(strTemp1, 1, "^"))
                    strTubePos = Trim(mGetP(strTemp1, 2, "^"))
                    strBarno = Trim(mGetP(strRcvBuf, 3, "^"))

                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ????ȯ??????
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
                        '## ???????? ????
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ???????? ????
                        strResult = strTemp2
                    End If
                        
                    '-- ?˻?????ó?? ???μ???
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB?? ????????
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ????
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_STAGO()
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
    Dim strAMRResult    As String   '?????? ????(????)
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
    Dim strAbbrName     As String   '?˻??ڵ?
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
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '?????? ????
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

                    '-- ????????
                    With mResult
                        .BarNo = strBarno
                        .Seq = strSeq
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ????ȯ??????
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    Select Case strFlag
                        Case "F"    '## ????
                            strResult = strIntResult
                        Case "I"    '## ????
                            Select Case Mid$(strIntResult, 1, 1)
                                Case "N":   strResult = "Negative"
                                Case "G":   strResult = "GRAYZONE"
                                Case "R":   strResult = "Positive"
                                Case "P":   strResult = "Positive"
                            End Select
                    End Select
                        
                    strSeqno = ""
                    strTestCode = ""
                    strTestName = ""
                    strAbbrName = ""
                    intResPrecUse = -1
                    intResPrec = -1
                    strAMRResult = ""
                    
                    '-- ?˻縶???? ???? ????????
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
                        SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
                        SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
                        SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                                      " & vbCrLf
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
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
                                
                            '-- AMR ???? (??ġ??)
                            If IsNumeric(strResult) Then
                                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                                    If CCur(strResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                                    If CCur(strResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                                    If CCur(strResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                                    If CCur(strResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                                    End If
                                End If
                            End If
                                                    
                            '-- AMR ???? (??????)
                            If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                                End If
                            End If
                            
                            If strAMRResult <> "" Then
                                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                                    If strIntResult <> "" Then
                                        strResult = strAMRResult & "(" & strIntResult & ")"
                                    Else
                                        strResult = strAMRResult
                                    End If
                                Else
                                    strResult = strAMRResult
                                End If
                            End If
                        
                            '--- ????????
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If IsNumeric(strResult) Then
                                    If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                        strJudge = ""
                                    ElseIf CCur(strHigh) <= CCur(strResult) Then
                                        strJudge = "H"
                                    ElseIf CCur(strLow) >= CCur(strResult) Then
                                        strJudge = "L"
                                    End If
                                End If
                            End If
                            
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
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
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
                            
                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                            '-- ???? ????
                            Call SetLocalDB(gRow, intRstRow, "1", "")
                            
                            strState = "R"
                            
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_STAGO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_UROMETER720()
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
    Dim strAMRResult    As String   '?????? ????(????)
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
    Dim strAbbrName     As String   '?˻??ڵ?
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
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
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
                    
                    '-- ????????
                    mOrder.Seq = strSeq
                    With mResult
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

                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    strResult = Mid(strRcvBuf, 8, 4) '-- ????
                    strResult = Trim(strResult)
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        strResult = Trim(Mid(strRcvBuf, 4))  '-- ????
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
                    '-- URO
                    If strResult = "norm" Then
                        strResult = "-"
                    End If
                    
                    '-- NIT
                    If strResult = "pos" Then
                        strResult = "+"
                    End If
            
                    Select Case Trim(strResult)
                        Case "-":       strResult = "negative"
                        Case "+":       strResult = "1positive"
                        Case "++":      strResult = "2positive"
                        Case "+++":     strResult = "3positive"
                        Case "++++":    strResult = "4positive"
                        Case "+/-":     strResult = "trace"
                    End Select
                    
                    strSeqno = ""
                    strTestCode = ""
                    strTestName = ""
                    strAbbrName = ""
                    intResPrecUse = -1
                    intResPrec = -1
                    strAMRResult = ""
                    
                    '-- ?˻縶???? ???? ????????
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
                        SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
                        SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
                        SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                                      " & vbCrLf
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
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
                                
                            '-- AMR ???? (??ġ??)
                            If IsNumeric(strResult) Then
                                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                                    If CCur(strResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                                    If CCur(strResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                                    If CCur(strResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                                    If CCur(strResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                                    End If
                                End If
                            End If
                                                    
                            '-- AMR ???? (??????)
                            If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                                End If
                            End If
                            
                            If strAMRResult <> "" Then
                                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                                    If strIntResult <> "" Then
                                        strResult = strAMRResult & "(" & strIntResult & ")"
                                    Else
                                        strResult = strAMRResult
                                    End If
                                Else
                                    strResult = strAMRResult
                                End If
                            End If
                        
                            '--- ????????
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If IsNumeric(strResult) Then
                                    If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                        strJudge = ""
                                    ElseIf CCur(strHigh) <= CCur(strResult) Then
                                        strJudge = "H"
                                    ElseIf CCur(strLow) >= CCur(strResult) Then
                                        strJudge = "L"
                                    End If
                                End If
                            End If
                            
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
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
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
                            
                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                            '-- ???? ????
                            Call SetLocalDB(gRow, intRstRow, "1", "")
                            
                            strState = "R"
                            
                        End If
                    End If
                    
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_UROMETER720" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HORIBA()
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
    Dim strAMRResult    As String   '?????? ????(????)
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
    Dim strAbbrName     As String   '?˻??ڵ?
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
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '?????? ????
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
                Case 4
                    If InStr(strRcvBuf, "AUTO_SID") > 0 Then
                        strSeq = Mid(strRcvBuf, InStr(strRcvBuf, "AUTO_SID") + 8)
                    End If
                    
                    '-- ????????
                    mOrder.Seq = strSeq
                    With mResult
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

                Case 9 To 27
                    strIntBase = Trim(Mid(strRcvBuf, 1, 2))
                    strResult = Trim(Mid(strRcvBuf, 3))
                    strResult = Replace(strResult, "h", "")
                    strResult = Replace(strResult, "H", "")
                    strResult = Replace(strResult, "l", "")
                    strResult = Replace(strResult, "L", "")
                    strResult = Replace(strResult, " ", "")
        
                    If strIntBase = "'" Then
                        strIntBase = "|"
                    End If
                    
                    strSeqno = ""
                    strTestCode = ""
                    strTestName = ""
                    strAbbrName = ""
                    intResPrecUse = -1
                    intResPrec = -1
                    strAMRResult = ""
                    
                    '-- ?˻縶???? ???? ????????
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,ABBRNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC   " & vbCrLf
                        SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7       " & vbCrLf
                        SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7      " & vbCrLf
                        SQL = SQL & "     , AMRINResult                                                                             " & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER                                                                               " & vbCrLf
                        SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                                    " & vbCrLf
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                                      " & vbCrLf
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
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
                                
                            '-- AMR ???? (??ġ??)
                            If IsNumeric(strResult) Then
                                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                                    If CCur(strResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT1"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                                    If CCur(strResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT2"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                                    If CCur(strResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT3"))
                                    End If
                                End If
                                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                                    If CCur(strResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                                        strAMRResult = Trim(RS_L.Fields("AMRRESULT4"))
                                    End If
                                End If
                            End If
                                                    
                            '-- AMR ???? (??????)
                            If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT5"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT6"))
                                End If
                            End If
                            If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                                If strResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                                    strAMRResult = Trim(RS_L.Fields("AMRRESULT7"))
                                End If
                            End If
                            
                            If strAMRResult <> "" Then
                                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                                    If strIntResult <> "" Then
                                        strResult = strAMRResult & "(" & strIntResult & ")"
                                    Else
                                        strResult = strAMRResult
                                    End If
                                Else
                                    strResult = strAMRResult
                                End If
                            End If
                        
                            '--- ????????
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If IsNumeric(strResult) Then
                                    If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                        strJudge = ""
                                    ElseIf CCur(strHigh) <= CCur(strResult) Then
                                        strJudge = "H"
                                    ElseIf CCur(strLow) >= CCur(strResult) Then
                                        strJudge = "L"
                                    End If
                                End If
                            End If
                            
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
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
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
                            
                            '-- ????Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                            '-- ???? ????
                            Call SetLocalDB(gRow, intRstRow, "1", "")
                            
                            strState = "R"
                            
                        End If
                    End If
                    
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

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "_SerialRcvData_HORIBA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
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

'-----------------------------------------------------------------------------'
'   ???? : ???????? ????
'-----------------------------------------------------------------------------'
Private Sub SendOrder_XN350()
    Dim strOutput   As String     '?۽??? ??????

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
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.OrgBarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q" & vbCr & ETX
                intSndPhase = 4
            Else
                '## ???? ??????
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.OrgBarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ???? ???ڿ??? ??????
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


Private Sub Phase_Serial_XN350()
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
                            Call SendOrder_XN350
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
                        
                        Call SerialRcvData_XN350
                        
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
    Dim strOutput   As String     '?۽??? ??????

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            
            '## ???????? ?????? ?Ǵ??Ͽ? SndPhase????
            If mOrder.NoOrder = True Then
                '## ?????????? ???°???
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
            '## ???? ??????
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
            '## ???? ???ڿ??? ??????
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
                        Call SerialRcvData_STAGO
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
                '??????û
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

    dtpFrom.Value = Now
    dtpTo.Value = Now
        
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
    'Me.Caption = gHOSP.MACHNM & Space$(5) & "?¢¢¢¢?     [?????? ???????̽?]     ?¢¢¢¢?"

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
    lblHospInfo(0).Caption = "?? " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " ???????̽?"
    lblHospInfo(1).Caption = "?? " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " ???????̽?"
    
    strIFStatus = ""
    strIFStatus = strIFStatus & IIf(gHOSP.SAVEAUTO = "Y", "?? ?ڵ?????????", "?? ????????????")
    If gHOSP.BARUSE = "Y" Then
        strIFStatus = strIFStatus & "  ?? ???ڵ?????"
    Else
        If gHOSP.RSTTYPE = "1" Then
            strIFStatus = strIFStatus & "  ?? ???? ????"
        ElseIf gHOSP.RSTTYPE = "2" Then
            strIFStatus = strIFStatus & "  ?? R/P ????"
        ElseIf gHOSP.RSTTYPE = "3" Then
            strIFStatus = strIFStatus & "  ?? üũ??"
        End If
    End If
    
    lblIFStatus.Caption = strIFStatus
    lblComStatus.Caption = ""
    
    If gWORKPOS = "M" Then
        spdWork.Visible = True
        fraWorkInfo.Visible = True
        
        cmdView.Visible = True
        cmdHide.Visible = True
        
    Else
        spdWork.Visible = False
        fraWorkInfo.Visible = False
    
        cmdView.Visible = False
        cmdHide.Visible = False
    
    End If


    Call CtlInitializing

    Call frmClear
    
    '-- Menu Set
    Call SetMenu

    '-- ?÷?????????
'    Call SetColumnHeader(spdOrder)

    '-- ?÷????̱⼳??
    Call SetColumnView(spdOrder)

    '-- ?˻??ڵ?
    Call GetTestList

    'Call GetTestListName

    '-- ?˻??? ???̱?
    Call SetExamCode(spdOrder)

    '-- ???ſ???
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
    
    If gHOSP.MACHNM = "VISION" Then
        fraVision.Visible = True
    Else
        fraVision.Visible = False
    End If
    'tmrConn.Interval = 60000
    'tmrConn.Enabled = True
    
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("??Ʈ ??ȣ?? ?߸??Ǿ????ϴ?." & vbNewLine & vbNewLine & "   ???? ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ????????"
            
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
            lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ???Ἲ??"
            
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgOn.ZOrder 0
           ' imgCom.Picture = imlStatus.ListImages("ON").ExtractIcon

        
        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ????????"
            
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
                
            
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ??????.."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False

        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
            
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ??????..."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
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
    SQL = SQL & "SELECT DISTINCT SEQNO, EQUIPCODE, EXAMNAME, EXAMCODE, EQUIPRESULT, RESULT, PREVRESULT, REFJUDGE" & vbCr
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
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
        lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ???Ἲ??"
        
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ????????"
        
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

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("?????? ???????Դϴ?. ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, "???α׷? ????") = vbYes Then

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
    
    '-- ????
'    If Row = 0 Then
'        '-- ???? ?߰?
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
        If strNewBarNo = "" Then
        
        End If
        
        If MsgBox(strNewBarNo & " ?? ?????ðڽ??ϱ??", vbInformation + vbYesNo, "?˸?") = vbNo Then
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
    strErrMsg = strErrMsg & "??    ġ : " & gHOSP.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    lblDBStatus.Caption = "?????ͺ??̽? ????????"
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

Private Sub wSCK_Close()
        
    If gComm.TCPTYPE = "SERVER" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ??Ʈ ???Ἲ??"
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ??Ʈ ???Ἲ??"
    End If

End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "SERVER" Then
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ??Ʈ ???Ἲ??"
        Else
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ??Ʈ ???Ἲ??"
        End If
    End If
            
End Sub

Private Sub wSCK_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    wSck.GetData strText
    SetRawData "[Rx]" & strText
    
    '-- ?????Ͻ? ?????? ??!!
'    strText = Replace(strText, vbLf, "")
    pBuffer = strText
    
    If Len(pBuffer) > 0 Then
    
        Select Case UCase(gHOSP.MACHNM)
            Case "KLITE":                    Call Phase_TCP_KLITE
            Case "VISION":                   Call Phase_TCP_VISION
        
        End Select
    End If

End Sub

