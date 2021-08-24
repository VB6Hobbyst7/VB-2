VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPrtICE 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ICE Box 출력 "
   ClientHeight    =   12735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   22560
   WindowState     =   2  '최대화
   Begin VB.TextBox txtComm 
      Appearance      =   0  '평면
      Height          =   5955
      Left            =   24120
      MultiLine       =   -1  'True
      TabIndex        =   76
      Top             =   1500
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   11205
      Left            =   90
      TabIndex        =   9
      Top             =   1050
      Width           =   25185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   10785
         Left            =   8160
         TabIndex        =   10
         Top             =   210
         Width           =   16785
         Begin VB.TextBox txtMatCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12870
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   113
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtExMonth 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11820
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   112
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13380
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   111
            Top             =   360
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12330
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   110
            Top             =   360
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtPrtPPBoxNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11250
            MaxLength       =   5
            TabIndex        =   108
            Top             =   360
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txtCompCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10770
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   107
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.Frame fraLotNo2 
            BackColor       =   &H00FFFFFF&
            Caption         =   " Lot 2 "
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8085
            Left            =   7380
            TabIndex        =   98
            Top             =   1710
            Visible         =   0   'False
            Width           =   8175
            Begin VB.TextBox txtInBarcode2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4410
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   128
               Top             =   5730
               Width           =   2520
            End
            Begin VB.Frame fra500e2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               Height          =   525
               Left            =   30
               TabIndex        =   121
               Top             =   7320
               Visible         =   0   'False
               Width           =   6975
               Begin VB.CommandButton cmd500ePrt2 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "출력"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5730
                  Style           =   1  '그래픽
                  TabIndex        =   125
                  Top             =   60
                  Width           =   825
               End
               Begin VB.TextBox txtPartsID2 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1500
                  TabIndex        =   124
                  Top             =   60
                  Width           =   1800
               End
               Begin VB.CommandButton cmdUnV2 
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   6630
                  TabIndex        =   123
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   315
               End
               Begin VB.TextBox txtQty2 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   4740
                  TabIndex        =   122
                  Top             =   60
                  Width           =   930
               End
               Begin VB.Label lblComp 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  '단일 고정
                  Caption         =   "PartsID"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Index           =   20
                  Left            =   120
                  TabIndex        =   127
                  Top             =   60
                  Width           =   1350
               End
               Begin VB.Label lblComp 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  '단일 고정
                  Caption         =   "수량"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Index           =   19
                  Left            =   3360
                  TabIndex        =   126
                  Top             =   60
                  Width           =   1350
               End
            End
            Begin VB.TextBox txtICEBoxNo2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4410
               MaxLength       =   5
               TabIndex        =   119
               Top             =   5310
               Width           =   1530
            End
            Begin VB.Frame fra408a2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '없음
               Height          =   525
               Left            =   60
               TabIndex        =   114
               Top             =   7350
               Visible         =   0   'False
               Width           =   6975
               Begin VB.CommandButton cmdUnvisible2 
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   6630
                  TabIndex        =   117
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   315
               End
               Begin VB.TextBox txtTopPrtNo2 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1500
                  TabIndex        =   116
                  Top             =   60
                  Width           =   4200
               End
               Begin VB.CommandButton cmdTopPrint2 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "출력"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5730
                  Style           =   1  '그래픽
                  TabIndex        =   115
                  Top             =   60
                  Width           =   825
               End
               Begin VB.Label lblComp 
                  Alignment       =   2  '가운데 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  '단일 고정
                  Caption         =   "출력 바코드"
                  BeginProperty Font 
                     Name            =   "맑은 고딕"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Index           =   17
                  Left            =   120
                  TabIndex        =   118
                  Top             =   60
                  Width           =   1350
               End
            End
            Begin VB.TextBox txtProdOrderDt2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00D0E0E0&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1650
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   104
               Top             =   300
               Width           =   1500
            End
            Begin VB.TextBox txtSlittingNo2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00D0E0E0&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5130
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   103
               Top             =   720
               Width           =   1500
            End
            Begin VB.TextBox txtLotNo2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00D0E0E0&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1650
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   101
               Top             =   720
               Width           =   1500
            End
            Begin FPSpread.vaSpread spdPrtReelDetail2 
               Height          =   3435
               Left            =   120
               TabIndex        =   99
               Top             =   1230
               Width           =   6885
               _Version        =   393216
               _ExtentX        =   12144
               _ExtentY        =   6059
               _StockProps     =   64
               ColsFrozen      =   8
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   15
               MaxRows         =   20
               RetainSelBlock  =   0   'False
               ScrollBarMaxAlign=   0   'False
               ScrollBars      =   2
               ScrollBarShowMax=   0   'False
               ShadowColor     =   16775150
               SpreadDesigner  =   "frmPrtICE.frx":0000
               ScrollBarTrack  =   3
               ShowScrollTips  =   3
            End
            Begin FPSpread.vaSpread spdScan2 
               Height          =   2505
               Left            =   150
               TabIndex        =   100
               Top             =   4710
               Width           =   2625
               _Version        =   393216
               _ExtentX        =   4630
               _ExtentY        =   4419
               _StockProps     =   64
               ColsFrozen      =   8
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "맑은 고딕"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GrayAreaBackColor=   16777215
               GridColor       =   15921919
               GridShowVert    =   0   'False
               MaxCols         =   19
               MaxRows         =   20
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               ScrollBars      =   2
               ShadowColor     =   16774120
               SpreadDesigner  =   "frmPrtICE.frx":0B3E
               ScrollBarTrack  =   3
               ShowScrollTips  =   3
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "내부바코드"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   21
               Left            =   3030
               TabIndex        =   129
               Top             =   5730
               Width           =   1350
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "ICE Box No"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   18
               Left            =   3030
               TabIndex        =   120
               Top             =   5310
               Width           =   1350
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "Slitting 작업번호"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   16
               Left            =   3300
               TabIndex        =   106
               Top             =   720
               Width           =   1800
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "생산일자"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   105
               Top             =   300
               Width           =   1500
            End
            Begin VB.Label lblLotNo2 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "Lot No"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   102
               Top             =   720
               Width           =   1500
            End
         End
         Begin VB.TextBox txtPPMaxTot 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   78
            Top             =   5910
            Width           =   810
         End
         Begin VB.TextBox txtReelQTY 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   77
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   63
            Top             =   360
            Width           =   1500
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   9075
            Left            =   7170
            ScaleHeight     =   9015
            ScaleWidth      =   9285
            TabIndex        =   30
            Top             =   1440
            Width           =   9345
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   15
               Left            =   900
               TabIndex        =   62
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   14
               Left            =   870
               TabIndex        =   61
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   13
               Left            =   690
               TabIndex        =   60
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   12
               Left            =   390
               TabIndex        =   59
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   11
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   10
               Left            =   0
               TabIndex        =   57
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   9
               Left            =   0
               TabIndex        =   56
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   8
               Left            =   0
               TabIndex        =   55
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   0
               TabIndex        =   54
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   5
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   15
               Left            =   0
               TabIndex        =   47
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   14
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   13
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   12
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   11
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   10
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   9
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   8
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   5
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Image imgBar2 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtICE.frx":1C7F
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.Image imgQrBar 
               Height          =   750
               Left            =   660
               Picture         =   "frmPrtICE.frx":5C4C
               Stretch         =   -1  'True
               Top             =   3000
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Image imgBar1 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtICE.frx":C284
               Stretch         =   -1  'True
               Top             =   1950
               Visible         =   0   'False
               Width           =   2685
            End
         End
         Begin VB.TextBox txtSlittingNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   29
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtCompNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   28
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtProdOrderDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   27
            Top             =   360
            Width           =   1500
         End
         Begin VB.TextBox txtPackNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   26
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtProdNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   25
            Top             =   360
            Width           =   1485
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   735
            Left            =   210
            TabIndex        =   20
            Top             =   9930
            Width           =   6915
            Begin VB.CommandButton cmdBC 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "BC"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   990
               Style           =   1  '그래픽
               TabIndex        =   84
               Top             =   150
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.CommandButton cmdErrClear 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   60
               Style           =   1  '그래픽
               TabIndex        =   83
               Top             =   150
               Width           =   675
            End
            Begin VB.CommandButton cmdView 
               BackColor       =   &H00E0E0E0&
               Caption         =   "보기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   3090
               Style           =   1  '그래픽
               TabIndex        =   24
               Top             =   150
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdPrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   4470
               Style           =   1  '그래픽
               TabIndex        =   23
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "닫기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   5580
               Style           =   1  '그래픽
               TabIndex        =   22
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdMakeBar 
               BackColor       =   &H00E0E0E0&
               Caption         =   "내부바코드출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1560
               Style           =   1  '그래픽
               TabIndex        =   21
               Top             =   150
               Visible         =   0   'False
               Width           =   1515
            End
         End
         Begin VB.TextBox txtICEBoxNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            MaxLength       =   5
            TabIndex        =   19
            Top             =   5010
            Width           =   1530
         End
         Begin VB.TextBox txtReelBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   18
            Text            =   "2X2707R0202001P10110000"
            Top             =   5460
            Width           =   4980
         End
         Begin VB.CheckBox chAutoPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동출력"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5310
            TabIndex        =   17
            Top             =   5910
            Value           =   1  '확인
            Width           =   1065
         End
         Begin VB.PictureBox picSide 
            BackColor       =   &H00FFFFFF&
            Height          =   3855
            Left            =   7170
            ScaleHeight     =   3795
            ScaleWidth      =   5775
            TabIndex        =   16
            Top             =   6660
            Width           =   5835
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   0
               Left            =   210
               Picture         =   "frmPrtICE.frx":10251
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   1
               Left            =   2520
               Picture         =   "frmPrtICE.frx":1421E
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   2
               Left            =   180
               Picture         =   "frmPrtICE.frx":181EB
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   3
               Left            =   2490
               Picture         =   "frmPrtICE.frx":1C1B8
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   4
               Left            =   180
               Picture         =   "frmPrtICE.frx":20185
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   5
               Left            =   2490
               Picture         =   "frmPrtICE.frx":24152
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   6
               Left            =   180
               Picture         =   "frmPrtICE.frx":2811F
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   7
               Left            =   2490
               Picture         =   "frmPrtICE.frx":2C0EC
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   8
               Left            =   180
               Picture         =   "frmPrtICE.frx":300B9
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   9
               Left            =   2490
               Picture         =   "frmPrtICE.frx":34086
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   10
               Left            =   180
               Picture         =   "frmPrtICE.frx":38053
               Stretch         =   -1  'True
               Top             =   3030
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   11
               Left            =   2490
               Picture         =   "frmPrtICE.frx":3C020
               Stretch         =   -1  'True
               Top             =   3030
               Visible         =   0   'False
               Width           =   2205
            End
         End
         Begin VB.TextBox txtMaxTot 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   15
            Top             =   5910
            Width           =   930
         End
         Begin VB.TextBox txtScanCount 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   6630
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            Top             =   5460
            Width           =   480
         End
         Begin VB.CheckBox chkReelPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reel출력"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   13620
            TabIndex        =   13
            Top             =   930
            Value           =   1  '확인
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtMsg 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   3015
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   6900
            Width           =   4245
         End
         Begin VB.TextBox txtInBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            Top             =   5010
            Width           =   2520
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   3435
            Left            =   240
            TabIndex        =   64
            Top             =   1440
            Width           =   6885
            _Version        =   393216
            _ExtentX        =   12144
            _ExtentY        =   6059
            _StockProps     =   64
            ColsFrozen      =   8
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   15
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            ShadowColor     =   16775150
            SpreadDesigner  =   "frmPrtICE.frx":3FFED
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin FPSpread.vaSpread spdScan 
            Height          =   2985
            Left            =   4500
            TabIndex        =   82
            Top             =   6900
            Width           =   2625
            _Version        =   393216
            _ExtentX        =   4630
            _ExtentY        =   5265
            _StockProps     =   64
            ColsFrozen      =   8
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   15921919
            GridShowVert    =   0   'False
            MaxCols         =   19
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   2
            ShadowColor     =   16774120
            SpreadDesigner  =   "frmPrtICE.frx":40B2B
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Frame fra500e 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   525
            Left            =   120
            TabIndex        =   90
            Top             =   6330
            Visible         =   0   'False
            Width           =   6975
            Begin VB.CommandButton cmdLabelPrint 
               BackColor       =   &H00E0E0E0&
               Caption         =   "전용라벨출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   130
               Top             =   60
               Width           =   1425
            End
            Begin VB.TextBox txtQty 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3840
               TabIndex        =   96
               Top             =   60
               Width           =   810
            End
            Begin VB.CommandButton cmdUnV 
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6630
               TabIndex        =   93
               Top             =   90
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txtPartsID 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1500
               TabIndex        =   92
               Top             =   60
               Width           =   1620
            End
            Begin VB.CommandButton cmd500ePrt 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4680
               Style           =   1  '그래픽
               TabIndex        =   91
               Top             =   60
               Width           =   825
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "수량"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   13
               Left            =   3150
               TabIndex        =   95
               Top             =   60
               Width           =   660
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "PartsID"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   94
               Top             =   60
               Width           =   1350
            End
         End
         Begin VB.Frame fra408a 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   525
            Left            =   120
            TabIndex        =   85
            Top             =   6330
            Visible         =   0   'False
            Width           =   6975
            Begin VB.CommandButton cmdTopPrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5730
               Style           =   1  '그래픽
               TabIndex        =   88
               Top             =   60
               Width           =   825
            End
            Begin VB.TextBox txtTopPrtNo 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1500
               TabIndex        =   87
               Top             =   60
               Width           =   4200
            End
            Begin VB.CommandButton cmdUnvisible 
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6630
               TabIndex        =   86
               Top             =   90
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "출력 바코드"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   89
               Top             =   60
               Width           =   1350
            End
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "출력 Box No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   9720
            TabIndex        =   109
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "3 매"
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
            Left            =   6420
            TabIndex        =   81
            Top             =   5970
            Width           =   435
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PP당 Reel수량"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   2580
            TabIndex        =   80
            Top             =   5910
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Lot No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   6360
            TabIndex        =   74
            Top             =   360
            Width           =   1800
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "박스당 PP수량"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   73
            Top             =   5910
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "포장코드"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   72
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "생산일자"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   3300
            TabIndex        =   71
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Slitting 작업번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   6360
            TabIndex        =   70
            Top             =   810
            Width           =   1800
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "고객사"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   3300
            TabIndex        =   69
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "ICE Box No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   67
            Top             =   5010
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PP Box 스캔"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   66
            Top             =   5460
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "내부바코드"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   3210
            TabIndex        =   65
            Top             =   5010
            Width           =   1350
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   10665
         Left            =   210
         TabIndex        =   75
         Top             =   300
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   18812
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   25
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtICE.frx":41C6C
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   25185
      Begin VB.TextBox txtCustomLabel 
         Height          =   315
         Left            =   15330
         TabIndex        =   131
         Top             =   300
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtTopPrtVal 
         Height          =   270
         Left            =   13950
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   300
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox chkYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "작업완료포함조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6750
         Style           =   1  '그래픽
         TabIndex        =   2
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   9210
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1650
         TabIndex        =   4
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   137756673
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   3750
         TabIndex        =   5
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   137756673
         CurrentDate     =   43884
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   9870
         Top             =   330
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
               Picture         =   "frmPrtICE.frx":4308E
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":43628
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":43BC2
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":4415C
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":449EE
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":44B48
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":44CA2
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":44DFC
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":456D6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblstrPrtLabelName 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   17580
         TabIndex        =   79
         Top             =   390
         Width           =   2265
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtICE.frx":45FB0
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '투명
         Caption         =   "Com1 연결실패"
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
         Height          =   285
         Left            =   11685
         TabIndex        =   8
         Top             =   420
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 생산일자 "
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
         Left            =   330
         TabIndex        =   7
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "~"
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
         Left            =   3450
         TabIndex        =   6
         Top             =   420
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPrtICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmPrtICEBox.frm
'   작성자  : 오세원
'   내  용  : ICE Box 라벨출력
'   작성일  : 2020-03-02
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'


Private Sub cmd500ePrt_Click()
    Dim strOutput   As String
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strAFont    As String
    Dim strData(0)  As String
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strFooter = "^XZ" & vbLf
    strAFont = "^A0N,75,45"
    
    If txtQty.Text <> "" Then
        strOutput = ""
        strOutput = strOutput & "^FO 100,100^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Grade    " & txtProdNm.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strOutput & "^FO 100,300^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Parts ID    " & txtPartsID.Text
        strOutput = strOutput & "^FS" & vbLf

        strOutput = strOutput & "^FO 600,280^CI26"
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtPartsID.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strOutput & "^FO 100,500^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Quantity (Reel)   " & txtQty.Text
        strOutput = strOutput & "^FS" & vbLf

        strOutput = strOutput & "^FO 600,480^CI26"
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtQty.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strHeader & strOutput & strFooter
        
        comEqp.Output = strOutput
    
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
        
        strData(0) = txtTopPrtVal & ETX & strOutput
    
        '트래킹 저장
        Call SetPackTrack("", txtInBarcode.Text, strData)
        
    
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
End Sub

Private Sub cmd500ePrt2_Click()
    Dim strOutput   As String
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strAFont    As String
    Dim strData(0)  As String
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strFooter = "^XZ" & vbLf
    strAFont = "^A0N,75,45"
    
    If txtQty2.Text <> "" Then
        strOutput = ""
        strOutput = strOutput & "^FO 100,100^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Grade    " & txtProdNm.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strOutput & "^FO 100,300^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Parts ID    " & txtPartsID2.Text
        strOutput = strOutput & "^FS" & vbLf

        strOutput = strOutput & "^FO 600,280^CI26"
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtPartsID2.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strOutput & "^FO 100,500^CI26"
        strOutput = strOutput & strAFont
        strOutput = strOutput & "^FD" & "Quantity (Reel)   " & txtQty2.Text
        strOutput = strOutput & "^FS" & vbLf

        strOutput = strOutput & "^FO 600,480^CI26"
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtQty2.Text
        strOutput = strOutput & "^FS" & vbLf
        
        strOutput = strHeader & strOutput & strFooter
        
        comEqp.Output = strOutput
    
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
        
        strData(0) = txtTopPrtVal & ETX & strOutput
    
        '트래킹 저장
        Call SetPackTrack("", txtInBarcode2.Text, strData)
        
    
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
End Sub

Private Sub cmdBC_Click()

    If spdScan.Visible = False Then
        spdScan.Visible = True
        spdScan.ZOrder 0
    Else
        spdScan.Visible = False
    End If
    
End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReel.MaxRows = 0
    spdPrtReelDetail.MaxRows = 0
    spdScan.MaxRows = 0
    spdPrtReelDetail2.MaxRows = 0
    spdScan2.MaxRows = 0
    
    'spdRegOrderDetail.MaxRows = 0
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now


    For i = 0 To 15
        'barReel.Visible = False
        'barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    For i = 0 To 11
        imgPpBar(i).Visible = False
    Next


    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
'    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    'txtReelQTY.Text = ""
    
    txtICEBoxNo.Text = ""
    txtPrtPPBoxNo.Text = ""
    txtReelBarcode.Text = ""
    txtMaxTot.Text = "0"
    txtScanCount.Text = "0"

    chkReelPrint.Value = "0"
    txtMsg.Text = ""
    txtInBarcode.Text = ""
    txtInBarcode2.Text = ""
    txtScanCount.Text = "0"

    lblstrPrtLabelName.Caption = ""
    
    txtLotNo2.Text = ""
    fraLotNo2.Visible = False
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub


' 바코드 리스트 가져옴
Private Sub GetOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)
    Dim strLabelType    As String
    
    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 3)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_NO").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReel, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReel, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReel, AdoRs.Fields("ROOL_INFO").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReel, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 10)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 12)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 13)
                Call SetText(spdPrtReel, AdoRs.Fields("CLOSE_YN").Value & "", .MaxRows, 14)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LABEL_CD").Value & "", .MaxRows, 15)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

' Reel 작업 리스트 가져옴
Private Sub GetReelOrderList_PP(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)

    Dim strLabelType    As String

    Set AdoRs = Get_OrderList_PP(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)

    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1

                Call SetText(spdPrtReel, "1", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
'                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_PRT_NAME").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReel, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 8)
                'Call SetText(spdPrtReel, AdoRs.Fields("ROOL_INFO").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 10)
                Call SetText(spdPrtReel, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 12)
                'Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_VIEW").Value & "", .MaxRows, 13)
                Call SetText(spdPrtReel, AdoRs.Fields("CLOSE_YN").Value & "", .MaxRows, 14)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LABEL_CD").Value & "", .MaxRows, 15)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                'Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With

            AdoRs.MoveNext
        Loop

    End If

    AdoRs.Close

End Sub


Private Sub cmdErrClear_Click()
    
    txtMsg.Text = ""
    
End Sub

Private Sub cmdMakeBar_Click()
'    Dim strAFont    As String
'    Dim strOutput   As String
'    Dim strBarcode  As String
'    Dim strHeader   As String
'
'    If spdScan.MaxRows = 0 Then
'        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    If txtInBarcode.Text = "" Then
'        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    strAFont = "^A0N,35,25"
'
'    strHeader = "^XA" & vbLf
'    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'    strHeader = strHeader & "^PON^FS" & vbLf
'    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'    strHeader = strHeader & "^MD9"
'
'    strOutput = ""
'    strOutput = strOutput & "^FO" & "20,100"
'    strOutput = strOutput & "^CI26"
'    strOutput = strOutput & "^BC" & "N,60,Y,N,N"
'    strOutput = strOutput & "^FD" & txtInBarcode.Text
'    strOutput = strOutput & "^FS" & vbLf
'
'
'    strOutput = strHeader & strOutput & "^XZ" & vbLf
'    comEqp.Output = strOutput

    Dim strPrtData  As String
    
    strPrtData = GetMakeInBar(txtInBarcode.Text)
    
    If strPrtData <> "" Then
        comEqp.Output = strPrtData
    End If

End Sub

'내부 관리용 바코드 출력한다
Private Function GetMakeInBar(pBarcode As String) As String
    Dim strAFont    As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strHeader   As String
    
    GetMakeInBar = ""
    
    If pBarcode = "" Then
        Exit Function
    End If
    
    strAFont = "^AJN,50,30"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^MD9"
    
    strOutput = ""
    
    If lblstrPrtLabelName.Caption = "P0003" Then
        strOutput = strOutput & "^FO500,50^CI26" & strAFont & "^FD국도화학 내부관리용코드^FS" & vbLf
        strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & pBarcode
    Else
        strOutput = strOutput & "^FO100,100^CI26" & strAFont & "^FD국도화학 내부관리용코드^FS" & vbLf
        strOutput = strOutput & "^FO100,200^CI26^BY3,3,100^BC^FD" & pBarcode
    End If
    strOutput = strOutput & "^FS" & vbLf
    strOutput = strHeader & strOutput & "^XZ" & vbLf
    
    GetMakeInBar = strOutput

End Function

'^XA : Opening Bracket으로 Format의 시작을 알린다.
'^FO : 인쇄 할 항목의 인쇄 할 위치(X축,Y축)를 정의한다.
'^A1 : 폰트설정
'      o => 회전 : N(0),R(90),I(180),B(270)
'      h => 높이 : 20
'      w => 넓이 : 25
'^BY : 바코드 사이즈 조정
'      w => module width        : 1 ~ 10    (Default : 2)
'      r => wide bar to narrow  : 2.0 ~ 3.0 (Default : 3)
'      h => bar code height     : 10
'^BQ : QR Code Bar Code

'^BC : Code 128 Bar Code (Subsets A, B, and C)  ==> Format ^BCo,h,f,g,e,m
'       o = orientation Accepted Values:
'           N = Normal
'           R = rotated 90 degrees (clockwise)
'           I = inverted 180 degrees
'           B = read from bottom up, 270 degrees
'           Default Value: current ^FW value
'       h = bar code height (in dots)
'           Accepted Values: 1 to 32000
'           Default Value: value set by ^BY
'       f = print interpretation line
'           Accepted Values: y (yes) Or N(NO)
'           Default Value: y
'           The interpretation line can be printed in any font by
'               placing the font command before the bar code command.
'       g = print interpretation line above code
'           Accepted Values: y (yes) Or N(NO)
'           Default Value: N
'       e = UCC check digit Accepted Values: Y (turns on) or N (turns off)
'           Mod 103 check digit is always there. It cannot be turned on or off. Mod 10 and 103 appear together with e turned on.
'           Default Value: N

'-----------------------------------------------------------------------------------------------------------------------
'  COMMAND SUMMARY
'-----------------------------------------------------------------------------------------------------------------------
' ^XA : start of label format
' ^XZ : end   of label format
' ^LH : set label home position
' ^FD : start of field data
' ^FS : end   of fiels data
' ^B3 : select Code39 bar code
' ^FO : set field position
' ^PQ : set print quantity  (^PQ10)
' ^B  : set barcode type
' ^BY : set barcode style   (code 128)
'-----------------------------------------------------------------------------------------------------------------------
' ^BY2,3,70                       ^BCN,70,Y,N,N
'
'  2 -> 가는선 font               ^BC -> code 128
'  3 -> 굵은선 font                70 -> barcode 높이
' 70 -> barcode 높이               Y -> barcode 하단에 문자열 출력
'                                  N -> barcode 상단에 문자열 출력
'                                  N -> check digit 표시유무
'-----------------------------------------------------------------------------------------------------------------------
' ^A  (FONT type)
'    font type       matrix     interchar gap     baseline
'                                 (in dots)       (in dots)
'        A             9X5            1               7
'        B            11X7            2               11
'       C,D           18X10           2               14
'        E            28X15           5               23
'        F            26X13           3               21
'        G            60X40           8               15
'        H            21X13           6                21
'        GS           24X24                       3XHEIGHT/4
'        0 DEFAULT    15X12                       3XHEIGHT/4
'
' ^A0,N,26,22 : D TYPE 높이 26, 폭 22dot
'
' ^CI  (Change International Font/Encoding)
' 26 = Multibyte Asian Encodings with ASCII Transparency a And c

        'asc("~") : R_7E
        '? Hex(126)
            
        'asc("℃") : R_A1C9
         '?Hex(126)
            
                                'strData = Replace(strData, "~", "_7E")
                                'strData = Replace(strData, "℃", "_A1C9")
            


Private Sub cmdPrint_Click()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            Call ICEPrint_203C
            Exit Sub
            
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            
                            '^FO800,300^CI26^BY4,4,120^BC^BCB,120,Y,N,N^FD2Y0P08A0000031P1000800^FS

                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & strFooter
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            Call ICEPrint_408G
            Exit Sub
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
'                        intReelCnt = 0
'                        With spdScan
'                            For m = 1 To .MaxRows
'                                .Row = m
'                                .Col = 1
'                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
'                            Next
'                        End With
                        
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo.Text = strBarcode
            
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strData = Mid(strData, 1, 21) & Format(CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput

            blnPrint = True
            
'            txtComm.Text = strOutput
'            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0

''        With spdScan
''            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
''                'For intCnt = 1 To .MaxRows
''                strXPos = 100: strYPos = 100
''
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 200
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 100
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            Else
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 100
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 200
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            End If
''        End With

'        For j = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(j)
'        Next
'
'        strOutput = strHeader & strOutput & "^XZ" & vbLf
'        comEqp.Output = strOutput
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

'''            strAFont = "^AJN,50,30"
'''            strHeader = "^XA" & vbLf
'''            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'''            strHeader = strHeader & "^PON^FS" & vbLf
'''            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'''            strHeader = strHeader & "^MD9"
'''
'''            strOutput = ""
'''            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
'''            strOutput = strOutput & "^FS" & vbLf
'''            strOutput = strHeader & strOutput & "^XZ" & vbLf
'''
'''            comEqp.Output = strOutput
        
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            'intCnt = intCnt + 1
            
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If
    
End Sub


Private Sub ICEPrint_500E_350()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            
                            '^FO800,300^CI26^BY4,4,120^BC^BCB,120,Y,N,N^FD2Y0P08A0000031P1000800^FS

                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & strFooter
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
'                        intReelCnt = 0
'                        With spdScan
'                            For m = 1 To .MaxRows
'                                .Row = m
'                                .Col = 1
'                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
'                            Next
'                        End With
                        
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo.Text = strBarcode
            
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strData = Mid(strData, 1, 21) & Format(CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput

            blnPrint = True
            
'            txtComm.Text = strOutput
'            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0

''        With spdScan
''            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
''                'For intCnt = 1 To .MaxRows
''                strXPos = 100: strYPos = 100
''
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 200
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 100
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            Else
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 100
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 200
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            End If
''        End With

'        For j = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(j)
'        Next
'
'        strOutput = strHeader & strOutput & "^XZ" & vbLf
'        comEqp.Output = strOutput
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

'''            strAFont = "^AJN,50,30"
'''            strHeader = "^XA" & vbLf
'''            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'''            strHeader = strHeader & "^PON^FS" & vbLf
'''            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'''            strHeader = strHeader & "^MD9"
'''
'''            strOutput = ""
'''            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
'''            strOutput = strOutput & "^FS" & vbLf
'''            strOutput = strHeader & strOutput & "^XZ" & vbLf
'''
'''            comEqp.Output = strOutput
        
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            'intCnt = intCnt + 1
            
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If

End Sub

Private Sub ICEPrint_500B_350()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            
                            '^FO800,300^CI26^BY4,4,120^BC^BCB,120,Y,N,N^FD2Y0P08A0000031P1000800^FS

                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & strFooter
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
'                        intReelCnt = 0
'                        With spdScan
'                            For m = 1 To .MaxRows
'                                .Row = m
'                                .Col = 1
'                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
'                            Next
'                        End With
                        
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo.Text = strBarcode
            
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strData = Mid(strData, 1, 21) & Format(CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput

            blnPrint = True
            
'            txtComm.Text = strOutput
'            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0

''        With spdScan
''            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
''                'For intCnt = 1 To .MaxRows
''                strXPos = 100: strYPos = 100
''
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 200
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 100
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            Else
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 100
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 200
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            End If
''        End With

'        For j = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(j)
'        Next
'
'        strOutput = strHeader & strOutput & "^XZ" & vbLf
'        comEqp.Output = strOutput
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

'''            strAFont = "^AJN,50,30"
'''            strHeader = "^XA" & vbLf
'''            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'''            strHeader = strHeader & "^PON^FS" & vbLf
'''            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'''            strHeader = strHeader & "^MD9"
'''
'''            strOutput = ""
'''            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
'''            strOutput = strOutput & "^FS" & vbLf
'''            strOutput = strHeader & strOutput & "^XZ" & vbLf
'''
'''            comEqp.Output = strOutput
        
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            'intCnt = intCnt + 1
            
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If

End Sub

Private Sub ICEPrint_500B_270()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            
                            '^FO800,300^CI26^BY4,4,120^BC^BCB,120,Y,N,N^FD2Y0P08A0000031P1000800^FS

                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & strFooter
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
'                        intReelCnt = 0
'                        With spdScan
'                            For m = 1 To .MaxRows
'                                .Row = m
'                                .Col = 1
'                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
'                            Next
'                        End With
                        
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo.Text = strBarcode
            
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strData = Mid(strData, 1, 21) & Format(CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput

            blnPrint = True
            
'            txtComm.Text = strOutput
'            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0

''        With spdScan
''            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
''                'For intCnt = 1 To .MaxRows
''                strXPos = 100: strYPos = 100
''
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 200
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 100
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            Else
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 100
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 200
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            End If
''        End With

'        For j = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(j)
'        Next
'
'        strOutput = strHeader & strOutput & "^XZ" & vbLf
'        comEqp.Output = strOutput
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

'''            strAFont = "^AJN,50,30"
'''            strHeader = "^XA" & vbLf
'''            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'''            strHeader = strHeader & "^PON^FS" & vbLf
'''            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'''            strHeader = strHeader & "^MD9"
'''
'''            strOutput = ""
'''            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
'''            strOutput = strOutput & "^FS" & vbLf
'''            strOutput = strHeader & strOutput & "^XZ" & vbLf
'''
'''            comEqp.Output = strOutput
        
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            'intCnt = intCnt + 1
            
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If

End Sub


Private Sub ICEPrint_400E_270()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            
                            '^FO800,300^CI26^BY4,4,120^BC^BCB,120,Y,N,N^FD2Y0P08A0000031P1000800^FS

                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & strFooter
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
'                        intReelCnt = 0
'                        With spdScan
'                            For m = 1 To .MaxRows
'                                .Row = m
'                                .Col = 1
'                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
'                            Next
'                        End With
                        
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo.Text = strBarcode
            
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strData = Mid(strData, 1, 21) & Format(CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
'            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
'            strOutput = ""
'            For j = 0 To UBound(strPrtData)
'                strOutput = strOutput & strPrtData(j)
'            Next
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
'            '3번 출력 수정...
'            For i = 1 To 3
'                comEqp.Output = strOutput
'            Next
'
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput

            blnPrint = True
            
'            txtComm.Text = strOutput
'            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0

''        With spdScan
''            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
''                'For intCnt = 1 To .MaxRows
''                strXPos = 100: strYPos = 100
''
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''                ReDim Preserve strPrtData(i) As String
''                strPrtData(i) = ""
''                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
''                strPrtData(i) = strPrtData(i) & "^CI26"
''                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
''                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                i = i + 1
''
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 200
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 100
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            Else
''                For intCnt = 1 To .MaxRows
''                    .Row = intCnt
''                    .Col = 1
''
''                    If intCnt = 1 Then
''                        strXPos = 100: strYPos = 100
''                    Else
''                        If intCnt Mod 2 = 0 Then
''                            strXPos = strPlusXPos:  strYPos = strYPos
''                        Else
''                            strXPos = 100:          strYPos = strYPos + 200
''                        End If
''                    End If
''
''                    ReDim Preserve strPrtData(i) As String
''                    strPrtData(i) = ""
''                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
''                    strPrtData(i) = strPrtData(i) & "^CI26"
''                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
''                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
''                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
''                    i = i + 1
''                Next
''            End If
''        End With

'        For j = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(j)
'        Next
'
'        strOutput = strHeader & strOutput & "^XZ" & vbLf
'        comEqp.Output = strOutput
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

'''            strAFont = "^AJN,50,30"
'''            strHeader = "^XA" & vbLf
'''            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'''            strHeader = strHeader & "^PON^FS" & vbLf
'''            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'''            strHeader = strHeader & "^MD9"
'''
'''            strOutput = ""
'''            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
'''            strOutput = strOutput & "^FS" & vbLf
'''            strOutput = strHeader & strOutput & "^XZ" & vbLf
'''
'''            comEqp.Output = strOutput
        
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            'intCnt = intCnt + 1
            
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If

End Sub



Private Sub ICEPrint_203C()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
'    strAFont = "^A0N,75,45"
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    '207도 회전
    strAFont = "^A0B,75,45"
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    

    With spdPrtReelDetail
        For intRow = 1 To .MaxRows
            '좌측 출력
            strType = GetText(spdPrtReelDetail, intRow, 1)
            strData = GetText(spdPrtReelDetail, intRow, 3)
            strXPos = GetText(spdPrtReelDetail, intRow, 4)
            strYPos = GetText(spdPrtReelDetail, intRow, 5)
            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
            strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
            strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
            
            Select Case strRot
                Case "0":   strRot = "N"
                Case "90":  strRot = "R"
                Case "180": strRot = "I"
                Case "270": strRot = "B"
            End Select
            
            If strType = "바코드" Then
                If Mid(strBarType, 1, 1) = "1" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                    strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                    strPrtData(i) = strPrtData(i) & "^BQ"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
                strBarcode = strData
            Else
                If strType = "LotNo" Then
                    strLot = mGetP(strData, 1, "(")
                    strLot = strLot & "(" & strLotSub & ")"
                    
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
            End If
        Next
    End With
        

    blnPrint = True
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        strInBarcode = txtInBarcode.Text

        '내부 관리용 바코드 출력한다
        If strInBarcode <> "" Then
            strInBarData = GetMakeInBar(strInBarcode)
            If strInBarData <> "" Then
                'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                If lblstrPrtLabelName.Caption <> "P0003" Then
                    '출력
                    comEqp.Output = strInBarData
                    
                    Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                    
                    '재출력용
                    strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                End If
            End If
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
    End If
    
End Sub

Private Sub ICEPrint_408G()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    
    '폰트
    strHeader = strHeader & "^MD8"
    '270 도 회전
    strAFont = "^A0B,75,45"
    strBarcode = ""
    
    With spdPrtReelDetail
        For intRow = 1 To .MaxRows
            strType = GetText(spdPrtReelDetail, intRow, 1)
            strData = GetText(spdPrtReelDetail, intRow, 3)
            strXPos = GetText(spdPrtReelDetail, intRow, 4)
            strYPos = GetText(spdPrtReelDetail, intRow, 5)
            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
            strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
            strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
            
            '바코드데이터 만들기
            If strType = "Material code" Then
                strBarcode = strData & "S2"
            ElseIf strType = "수량" Then
                strReelCnt = strData
                strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                intReelCnt = strReelCnt
                
            ElseIf strType = "생산일자" Then
                strBarcode = strBarcode & Get_YMD("Y4", Year(strData))
                strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                strBarcode = strBarcode & Get_YMD("D2", Day(strData))
                strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                strBarcode = strBarcode & Format(intReelCnt, "0000")
                strBarcode = strBarcode & "00"
            End If
                                    
            Select Case strRot
                Case "0":   strRot = "N"
                Case "90":  strRot = "R"
                Case "180": strRot = "I"
                Case "270": strRot = "B"
            End Select
            
            If strType = "바코드" Then
                strLotSub = Format(intPrt, "0000")
                strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                If Mid(strBarType, 1, 1) = "1" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                    strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
                strBarcode = strData
            Else
                ReDim Preserve strPrtData(i) As String
                strPrtData(i) = ""
                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                strPrtData(i) = strPrtData(i) & "^CI26"
                strPrtData(i) = strPrtData(i) & strAFont
                strPrtData(i) = strPrtData(i) & "^FH"
                If strNamePrt = "Y" Then
                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                Else
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                End If
                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                i = i + 1
            End If
        Next
    End With
    
    '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
    txtTopPrtNo.Text = strBarcode
    
    blnPrint = True
    

   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
    
    'intCnt = intCnt + 1
       
               
    If blnPrint = True Then
        Dim intMaxNo    As Integer
        Dim strDate     As String
        
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text

            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        
        txtCustomLabel.Text = ""
        For i = 1 To spdScan.MaxRows
            txtCustomLabel.Text = txtCustomLabel.Text & GetText(spdScan, i, 1) & "|"
        Next
        
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    
    End If
    
    If fraLotNo2.Visible = True Then
        Call IcePrint2
        If lblstrPrtLabelName.Caption <> "P0003" And lblstrPrtLabelName.Caption <> "P0006" And lblstrPrtLabelName.Caption <> "P0007" And lblstrPrtLabelName.Caption <> "P0010" Then
            fraLotNo2.Visible = False
        End If
    
    End If

End Sub





Private Sub IcePrint2()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim m           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    Dim intReelCnt  As Integer
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    '-- 재출력용
    Dim strPrintVal()    As Variant
    
    Dim strInBarData    As String
    Dim strInBarcode    As String
    Dim strReelCnt      As String
    
    Dim intMaxNo    As Integer
    Dim strDate     As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal

    blnPrint = False
    strBarcode = ""
    strOutput = ""
    intCnt = 0
    i = 0
    intReelCnt = 0
    
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"

            With spdPrtReelDetail2
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail2, intRow, 1)
                    strData = GetText(spdPrtReelDetail2, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail2, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail2, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail2, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail2, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail2, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail2, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail2, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        ''strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY4,4,120^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,120,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            blnPrint = True
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0B,75,45"
            strBarcode = ""
            
            With spdPrtReelDetail2
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail2, intRow, 1)
                    strData = GetText(spdPrtReelDetail2, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail2, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail2, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail2, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail2, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail2, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail2, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail2, intRow, 10)
                    
                    '바코드데이터 만들기
                    If strType = "Material code" Then
                        strBarcode = strData & "S2"
                    ElseIf strType = "수량" Then
                        strReelCnt = strData
                        strReelCnt = Trim(Replace(strReelCnt, "Reel", ""))
                        strReelCnt = Trim(Replace(strReelCnt, "REEL", ""))
                        intReelCnt = strReelCnt
                        
                    ElseIf strType = "생산일자" Then
                        'MAX NO 찾기
                        Set AdoRs = Get_MAX_NO(strData, gPackTrack.PRODCD, "I")
                        If AdoRs.RecordCount = 0 Then
                            intMaxNo = 1
                        Else
                            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
                            intMaxNo = intMaxNo + 1
                        End If
                        AdoRs.Close
                        
                        txtICEBoxNo2.Text = intMaxNo
                        
                        strBarcode = strBarcode & Get_YMD("Y7", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D4", Day(strData))
                        strBarcode = strBarcode & Format(intMaxNo, "000")
                        strBarcode = strBarcode & Format(intReelCnt, "0000")
                        strBarcode = strBarcode & "00"
                        
                    End If
                                            
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & strAFont
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
            
            '바코드 데이터 만들기 (라벨에 바코드는 안찍히나 따로 바코드만 한장을 찍음)
            txtTopPrtNo2.Text = strBarcode
            blnPrint = True
        
        Case "P0004", "P0005", "P0008", "P0009"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail2
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail2, intRow, 1)
                    strData = GetText(spdPrtReelDetail2, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail2, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail2, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail2, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail2, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail2, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail2, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail2, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            blnPrint = True
        Case "P0006", "P0007", "P0010"
            '폰트
            strHeader = strHeader & "^MD8"
            '270 도 회전
            strAFont = "^A0R,75,45"
            
            With spdPrtReelDetail2
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail2, intRow, 1)
                    strData = GetText(spdPrtReelDetail2, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail2, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail2, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail2, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail2, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail2, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail2, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail2, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        strData = Mid(strData, 1, 21) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 25)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData
                            'strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & strAFont
                            If strType = "바코드값" Then
                                strPrtData(i) = strPrtData(i) & "^A0R,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            blnPrint = True
    End Select
   
    strOutput = ""
    For j = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(j)
    Next
        
    strOutput = strHeader & strOutput & strFooter
    
    '3번 출력 수정...
    For i = 1 To 3
        comEqp.Output = strOutput
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    Next
    
    '트래킹용
    ReDim Preserve strTrackBC(intCnt)
    strTrackBC(intCnt) = strBarcode
    
    '재출력용
    ReDim Preserve strPrintVal(intCnt)
    strPrintVal(intCnt) = strOutput
    txtTopPrtVal.Text = strOutput
               
    If blnPrint = True Then
        
        strDate = Format(txtProdOrderDt2.Text, "yyyymmdd")
        
        gPackTrack.ORDERDT = strDate
        
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "I")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "I", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "I", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'ICE박스 B + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode2.Text = "B" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        strInBarcode = ""
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Then
            strInBarcode = txtInBarcode.Text
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    'TP408A는 보류한다  =>  보류가 아니라 제외 됨
                    If lblstrPrtLabelName.Caption <> "P0003" Then
                        '출력
                        comEqp.Output = strInBarData
                        
                        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                        
                        '재출력용
                        strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    End If
                End If
            End If
        End If
        
        '트래킹 저장
        Call SetPackTrack2(strBarcode, strInBarcode, strPrintVal)
    
        'txtICEBoxNo.Text = intMaxNo + 1
        'txtScanCount.Text = "0"
        spdScan2.MaxRows = 0
        
'        For intCnt = 0 To 9
'            imgPpBar(intCnt).Visible = False
'        Next
    
    End If
    
End Sub


Private Sub SetPackTrack(ByVal pICEBarcode As String, ByVal pPPINBarcode As String, ByVal pICEPrtVal As Variant)
    Dim intCnt      As Integer
    Dim intMaxNo    As Integer
    
    With spdScan
        For intCnt = 1 To .MaxRows
            .Row = intCnt
            .Col = 1
            gPackTrack.ORDERDT = Format(txtProdOrderDt.Text, "yyyymmdd")     'Key
            gPackTrack.PRODCD = txtProdCd.Text                               'Key
            gPackTrack.REELBAR = ""
            gPackTrack.PPBAR = GetText(spdScan, intCnt, 3) ' .Text
            gPackTrack.ICEBAR = pICEBarcode
            gPackTrack.PPBARIN = GetText(spdScan, intCnt, 5)
            gPackTrack.ICEBARIN = pPPINBarcode 'txtInBarcode.Text  'txtPPBoxNo.Text
            gPackTrack.LOTNO = txtLotNo.Text
            gPackTrack.REELPRTID = ""
            gPackTrack.REELPRTDT = ""
            gPackTrack.PPPRTID = ""
            gPackTrack.PPPRTDT = ""
            gPackTrack.ICEPRTID = gKUKDO.USERID
            gPackTrack.ICEPRTDT = ""
            
            '재출력용
            gPackTrack.REELVAL = ""
            gPackTrack.PPVAL = ""
            '제일 처음 PP바코드에 저장한다.
            If intCnt = 1 Then
                gPackTrack.ICEVAL = pICEPrtVal(0)
            'ElseIf intCnt = 2 Then
            '    gPackTrack.ICEVAL = pICEPrtVal(1)
            Else
                gPackTrack.ICEVAL = ""
            End If
            
            '트래킹 저장
            '-- PP Box 는 Insert 없음
            If Set_Pack_Track("UP", "I") Then
            End If
        Next
    End With

End Sub

Private Sub SetPackTrack2(ByVal pICEBarcode As String, ByVal pPPINBarcode As String, ByVal pICEPrtVal As Variant)
    Dim intCnt      As Integer
    Dim intMaxNo    As Integer
    
    With spdScan2
        For intCnt = 1 To .MaxRows
            .Row = intCnt
            .Col = 1
            gPackTrack.ORDERDT = Format(txtProdOrderDt2.Text, "yyyymmdd")     'Key
            gPackTrack.PRODCD = txtProdCd.Text                               'Key
            gPackTrack.REELBAR = ""
            gPackTrack.PPBAR = GetText(spdScan2, intCnt, 3) ' .Text
            gPackTrack.ICEBAR = pICEBarcode
            gPackTrack.PPBARIN = GetText(spdScan2, intCnt, 5)
            gPackTrack.ICEBARIN = pPPINBarcode 'txtInBarcode.Text  'txtPPBoxNo.Text
            gPackTrack.LOTNO = txtLotNo2.Text
            gPackTrack.REELPRTID = ""
            gPackTrack.REELPRTDT = ""
            gPackTrack.PPPRTID = ""
            gPackTrack.PPPRTDT = ""
            gPackTrack.ICEPRTID = gKUKDO.USERID
            gPackTrack.ICEPRTDT = ""
            
            '재출력용
            gPackTrack.REELVAL = ""
            gPackTrack.PPVAL = ""
            '제일 처음 PP바코드에 저장한다.
            If intCnt = 1 Then
                gPackTrack.ICEVAL = pICEPrtVal(0)
            'ElseIf intCnt = 2 Then
            '    gPackTrack.ICEVAL = pICEPrtVal(1)
            Else
                gPackTrack.ICEVAL = ""
            End If
            
            '트래킹 저장
            '-- PP Box 는 Insert 없음
            If Set_Pack_Track("UP", "I") Then
            End If
        Next
    End With

End Sub

Private Sub cmdSearch_Click()
    Dim strFromDt    As String
    Dim strToDt      As String
    Dim strYN        As String
    
    strFromDt = Format(dtpFromDate, "yyyymmdd")
    strToDt = Format(dtpToDate, "yyyymmdd")
    
    Call cmdClear_Click
    
    Call GetReelOrderList_PP(strFromDt, strToDt, "", "", "I")

End Sub


Private Sub cmdTopPrint_Click()
    Dim strOutput   As String
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strAFont   As String
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strFooter = "^XZ" & vbLf
    strAFont = "^A0N,75,45"
    
    If txtTopPrtNo.Text <> "" Then
        strOutput = ""
        strOutput = strOutput & "^FO 100,100"
        strOutput = strOutput & "^CI26"
'        strOutput = strOutput & strAFont
'        strOutput = strOutput & "^FD" & txtTopPrtNo.Text
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtTopPrtNo.Text
        strOutput = strOutput & "^FS" & vbLf
        strOutput = strHeader & strOutput & strFooter
        
        comEqp.Output = strOutput
    
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    
'        If lblstrPrtLabelName.Caption = "P0003" Then
'            fra408a.Visible = True
'        Else
'            fra408a.Visible = False
'        End If
    End If

End Sub

Private Sub cmdTopPrint2_Click()
    Dim strOutput   As String
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strAFont   As String
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strFooter = "^XZ" & vbLf
    strAFont = "^A0N,75,45"
    
    If txtTopPrtNo.Text <> "" Then
        strOutput = ""
        strOutput = strOutput & "^FO 100,100"
        strOutput = strOutput & "^CI26"
        strOutput = strOutput & "^BC" & "N,100,Y,N,N"
        strOutput = strOutput & "^FD" & txtTopPrtNo2.Text
        strOutput = strOutput & "^FS" & vbLf
        strOutput = strHeader & strOutput & strFooter
        
        comEqp.Output = strOutput
    
        Call SetPrtData("ICEBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
    End If
End Sub

Private Sub cmdUnV_Click()

    fra500e.Visible = False
    
End Sub

Private Sub cmdUnV2_Click()

    fra500e2.Visible = False

End Sub

Private Sub cmdUnvisible_Click()

    fra408a.Visible = False
    
End Sub

Private Sub cmdUnvisible2_Click()

    fra408a2.Visible = False
    
End Sub

Private Sub cmdView_Click()
    
    If txtComm.Visible = False Then
        txtComm.Visible = True
    Else
        txtComm.Visible = False
    End If
    
End Sub




Private Sub cmdLabelPrint_Click()
    Dim strPlusXPos As String
    Dim strXPos     As String
    Dim strYPos     As String
    Dim strAFont    As String
    Dim strHeader   As String
    Dim strFooter   As String
    Dim strPrtData() As String
    Dim i           As Integer
    Dim intCnt      As Integer
    Dim strOutput   As String
    Dim strVal()    As String
    
    strVal = Split(txtCustomLabel.Text, "|")
    txtCustomLabel.Text = ""
    
    Erase strPrtData
    i = 0
    strPlusXPos = 630
    strAFont = "^A0N,75,45"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    
    strFooter = "^XZ" & vbLf
    
    strXPos = 100: strYPos = 100

    i = i + 1
    ReDim Preserve strPrtData(i) As String
    strPrtData(i) = ""
    strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
    strPrtData(i) = strPrtData(i) & "^CI26"
    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
    strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
    i = i + 1

    For intCnt = 1 To UBound(strVal)
        If intCnt = 1 Then
            strXPos = 100: strYPos = 200
        Else
            If intCnt Mod 2 = 0 Then
                strXPos = strPlusXPos:  strYPos = strYPos
            Else
                strXPos = 100:          strYPos = strYPos + 100
            End If
        End If

        ReDim Preserve strPrtData(i) As String
        strPrtData(i) = ""
        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
        strPrtData(i) = strPrtData(i) & "^CI26"
        strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
        strPrtData(i) = strPrtData(i) & "^FD" & strVal(intCnt - 1)
        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
        i = i + 1
    Next

    For i = 1 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(i)
    Next
    
    strOutput = strHeader & strOutput & strFooter
    
    comEqp.Output = strOutput

    Call SetPrtData("ICEBOX" & "_Custom_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")

End Sub

Private Sub Form_Load()
    
    Unload frmPrtReel
    Unload frmPrtPP
    Unload frmPrtReprint

    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle

'    frmPrtReel.comEqp.PortOpen = False
'    frmPrtPP.comEqp.PortOpen = False
    
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
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            
            Resume Next
        Else
            
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & "Public Sub OpenCommunication()" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If


End Sub

'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdPrtReel
        Call SetText(spdPrtReel, "선택", 0, 1):              .ColWidth(1) = 0
        Call SetText(spdPrtReel, "Lot No", 0, 2):            .ColWidth(2) = 12
        Call SetText(spdPrtReel, "제조일자", 0, 3):          .ColWidth(3) = 10
        Call SetText(spdPrtReel, "공정No", 0, 4):            .ColWidth(4) = 0
        Call SetText(spdPrtReel, "제품코드", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReel, "제품명", 0, 6):            .ColWidth(6) = 12
        Call SetText(spdPrtReel, "포장코드", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReel, "메모", 0, 8):              .ColWidth(8) = 0
        Call SetText(spdPrtReel, "작업내용설명", 0, 9):      .ColWidth(9) = 0 'Roll정보
        Call SetText(spdPrtReel, "길이", 0, 10):             .ColWidth(10) = 4
        Call SetText(spdPrtReel, "SLT No", 0, 11):           .ColWidth(11) = 0
        Call SetText(spdPrtReel, "수량", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReel, "고객사", 0, 13):           .ColWidth(13) = 10
        Call SetText(spdPrtReel, "작업완료여부", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReel, "라벨코드", 0, 15):         .ColWidth(15) = 10
        Call SetText(spdPrtReel, "입력자", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReel, "입력일시", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReel, "수정자", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReel, "수정일시", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdPrtReelDetail
        Call SetText(spdPrtReelDetail, "항목", 0, 1):          .ColWidth(1) = 15
        Call SetText(spdPrtReelDetail, "순서", 0, 2):          .ColWidth(2) = 4
        Call SetText(spdPrtReelDetail, "내용", 0, 3):          .ColWidth(3) = 35
        Call SetText(spdPrtReelDetail, "X", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdPrtReelDetail, "Y", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReelDetail, "", 0, 6):            .ColWidth(6) = 0
        Call SetText(spdPrtReelDetail, "", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReelDetail, "", 0, 8):          .ColWidth(8) = 0
        Call SetText(spdPrtReelDetail, "", 0, 9):          .ColWidth(9) = 0
        Call SetText(spdPrtReelDetail, "", 0, 10):      .ColWidth(10) = 0
        Call SetText(spdPrtReelDetail, "", 0, 11):           .ColWidth(11) = 0
        Call SetText(spdPrtReelDetail, "", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReelDetail, "", 0, 13):           .ColWidth(13) = 0
        Call SetText(spdPrtReelDetail, "", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReelDetail, "", 0, 15):         .ColWidth(15) = 0
        Call SetText(spdPrtReelDetail, "", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReelDetail, "", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReelDetail, "", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReelDetail, "", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdPrtReelDetail2
        Call SetText(spdPrtReelDetail2, "항목", 0, 1):          .ColWidth(1) = 15
        Call SetText(spdPrtReelDetail2, "순서", 0, 2):          .ColWidth(2) = 4
        Call SetText(spdPrtReelDetail2, "내용", 0, 3):          .ColWidth(3) = 35
        Call SetText(spdPrtReelDetail2, "X", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdPrtReelDetail2, "Y", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 6):            .ColWidth(6) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 8):          .ColWidth(8) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 9):          .ColWidth(9) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 10):      .ColWidth(10) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 11):           .ColWidth(11) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 13):           .ColWidth(13) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 15):         .ColWidth(15) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReelDetail2, "", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    
    With spdScan
        Call SetText(spdScan, "스캔바코드", 0, 1):          .ColWidth(1) = 18
        Call SetText(spdScan, "REEL바코드", 0, 2):          .ColWidth(2) = 0
        Call SetText(spdScan, "PP바코드", 0, 3):            .ColWidth(3) = 0
        Call SetText(spdScan, "ICE바코드", 0, 4):           .ColWidth(4) = 0
        Call SetText(spdScan, "PP바코드_IN", 0, 5):         .ColWidth(5) = 0
        Call SetText(spdScan, "ICE바코드_IN", 0, 6):        .ColWidth(6) = 0
        .MaxRows = 0
        .MaxCols = 6
        .RowHeight(-1) = 16
    End With
    
    With spdScan2
        Call SetText(spdScan2, "스캔바코드", 0, 1):          .ColWidth(1) = 18
        Call SetText(spdScan2, "REEL바코드", 0, 2):          .ColWidth(2) = 0
        Call SetText(spdScan2, "PP바코드", 0, 3):            .ColWidth(3) = 0
        Call SetText(spdScan2, "ICE바코드", 0, 4):           .ColWidth(4) = 0
        Call SetText(spdScan2, "PP바코드_IN", 0, 5):         .ColWidth(5) = 0
        Call SetText(spdScan2, "ICE바코드_IN", 0, 6):        .ColWidth(6) = 0
        .MaxRows = 0
        .MaxCols = 6
        .RowHeight(-1) = 16
    End With
    
'    With spdRegOrderDetail
'        Call SetText(spdRegOrderDetail, "제조일자", 0, 1):        .ColWidth(1) = 0
'        Call SetText(spdRegOrderDetail, "순번", 0, 2):            .ColWidth(2) = 0
'        Call SetText(spdRegOrderDetail, "제품코드", 0, 3):        .ColWidth(3) = 0
'        Call SetText(spdRegOrderDetail, "SLT No", 0, 4):          .ColWidth(4) = 0
'        Call SetText(spdRegOrderDetail, "일련번호", 0, 5):        .ColWidth(5) = 8
'        Call SetText(spdRegOrderDetail, "SLT내용", 0, 6):         .ColWidth(6) = 28
'        Call SetText(spdRegOrderDetail, "시작번호", 0, 7):        .ColWidth(7) = 10
'        Call SetText(spdRegOrderDetail, "끝번호", 0, 8):          .ColWidth(8) = 10
'        Call SetText(spdRegOrderDetail, "", 0, 9):                .ColWidth(9) = 0
'        Call SetText(spdRegOrderDetail, "No", 0, 10):             .ColWidth(10) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 11):               .ColWidth(11) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 12):               .ColWidth(12) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 13):               .ColWidth(13) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 14):               .ColWidth(14) = 0
'        Call SetText(spdRegOrderDetail, "사용여부", 0, 15):       .ColWidth(15) = 0
'        Call SetText(spdRegOrderDetail, "입력자", 0, 16):         .ColWidth(16) = 0
'        Call SetText(spdRegOrderDetail, "입력일시", 0, 17):       .ColWidth(17) = 0
'        Call SetText(spdRegOrderDetail, "수정자", 0, 18):         .ColWidth(18) = 0
'        Call SetText(spdRegOrderDetail, "수정일시", 0, 19):       .ColWidth(19) = 0
'
'        .MaxRows = 0
'    End With
    
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtLotNo.Text = ""
    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
'    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    
    txtICEBoxNo.Text = ""
    txtPrtPPBoxNo.Text = ""
    txtReelBarcode.Text = ""
    txtMaxTot.Text = "0"
    txtScanCount.Text = "0"

    chkReelPrint.Value = "0"
    txtMsg.Text = ""
    
    txtLotNo2.Text = ""
    fraLotNo2.Visible = False

    'txtReelQTY.Text = ""
    
    gSORT = 0

End Sub



Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pAdoRS1      As ADODB.Recordset
    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
    Dim j               As Integer
    Dim strPrtSide      As String
    Dim strBarCd        As String
    Dim strDate         As String
    Dim strQty          As String
    Dim strProdLabelCd  As String
    Dim strProdCd       As String
    Dim strCompCd       As String
    Dim strTmp          As String
    Dim strGu           As String
    Dim strContents     As String
    Dim strLeft         As String
    Dim strLotNo        As String
    Dim strLotNoFull    As String
    '자재코드
    Dim strBarData      As String
    Dim strProdName     As String
    Dim strProdLen      As String
    Dim strProdMatCd    As String
    Dim strExMonth      As String
    Dim strProdSize     As String
    Dim strChimeiCd     As String
    Dim strVendorCd     As String
    Dim strProdLineFA   As String
    Dim strProdSlitFA   As String
    Dim strContYN       As String
    Dim strPcnNO        As String
    Dim strMaxTot       As String
    Dim strPPMaxTot     As String
    Dim strExDate       As String
    Dim strProdTemp     As String
    Dim strPrtLabelName As String
    Dim strProdCalLen   As String
    Dim intMaxNo    As Integer
    
    If Row = 0 Then
        If Col = 1 Then
            If GetText(spdPrtReel, 1, 1) = "1" Then
                For i = 1 To spdPrtReel.DataRowCnt
                    Call SetText(spdPrtReel, "0", i, 1)
                Next
            Else
                For i = 1 To spdPrtReel.DataRowCnt
                    Call SetText(spdPrtReel, "1", i, 1)
                Next
            End If
        Else
            '-- 정렬 추가
            Call SetSpreadSort(spdPrtReel, 0)
        End If
        Exit Sub
    End If
    
    For i = 0 To 15
        'barReel.Visible = False
        'barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    For i = 0 To 11
        imgPpBar(i).Visible = False
    Next
    
    txtReelBarcode.Text = ""
    txtInBarcode.Text = ""
    txtInBarcode2.Text = ""
    txtScanCount.Text = "0"
    spdScan.MaxRows = 0
    
    imgBar1.Visible = False
    imgBar2.Visible = False
    imgQrBar.Visible = False
    
    'strPrtLabelName = GetText(spdPrtReel, Row, 6) & "|" & GetText(spdPrtReel, Row, 10)
    strPrtLabelName = GetText(spdPrtReel, Row, 5)
    lblstrPrtLabelName.Caption = strPrtLabelName
    
    strDate = GetText(spdPrtReel, Row, 3)
    txtProdOrderDt.Text = strDate
    strProdCd = GetText(spdPrtReel, Row, 5)
    txtProdCd.Text = strProdCd
    txtProdNm.Text = GetText(spdPrtReel, Row, 6)
    strProdLen = GetText(spdPrtReel, Row, 10)
    txtProdLen.Text = strProdLen
    strProdLen = strProdLen * 100 '미터를 cm으로 변환
'    txtProdPosNo.Text = GetText(spdPrtReel, Row, 4)
    txtPackNm.Text = GetText(spdPrtReel, Row, 7)
    txtReelQTY.Text = GetText(spdPrtReel, Row, 12)
    txtSlittingNo.Text = GetText(spdPrtReel, Row, 11)
    txtCompNm.Text = GetText(spdPrtReel, Row, 13)
    strLotNo = GetText(spdPrtReel, Row, 2)
    txtLotNo.Text = strLotNo
    strProdLabelCd = GetText(spdPrtReel, Row, 15)
    strCompCd = GetText(spdPrtReel, Row, 18)
    txtCompCd.Text = strCompCd
    strQty = txtReelQTY.Text
    txtTopPrtVal.Text = ""
    txtPartsID.Text = ""
    txtQty.Text = ""
    
    
    txtLotNo2.Text = ""
    spdPrtReelDetail2.MaxRows = 0
    spdScan2.MaxRows = 0
    fraLotNo2.Visible = False
    
    
    gPackTrack.PRODCD = strProdCd   '5자리
    gPackTrack.LOTNO = strLotNo
    gPackTrack.ORDERDT = strDate    '8자리
              
    'MAX NO 찾기
    Set AdoRs = Get_MAX_NO(gPackTrack.ORDERDT, gPackTrack.PRODCD, "I")
    If AdoRs.RecordCount = 0 Then
        intMaxNo = 1
    Else
        intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
        intMaxNo = intMaxNo + 1
    End If
    AdoRs.Close
    
    txtICEBoxNo.Text = intMaxNo

    With spdPrtReelDetail
        .MaxRows = 0
    End With

    Set AdoRs = Get_LabelDetail(strProdLabelCd, "I")
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        '-- 제품정보 찾아오기
        SQL = ""
        SQL = SQL & "SELECT                                 " & vbCrLf
        SQL = SQL & "  M.PROD_NAME                          " & vbCrLf
        SQL = SQL & ", M.PROD_LENGTH                        " & vbCrLf
        SQL = SQL & ", M.PROD_MATERIAL_CD                   " & vbCrLf
        SQL = SQL & ", M.EXPIR_MONTH                        " & vbCrLf
        SQL = SQL & ", M.PROD_STOR_TEMP                     " & vbCrLf
        SQL = SQL & ", M.PROD_SIZE                          " & vbCrLf
        SQL = SQL & ", M.PROD_CHIMEI_PN                     " & vbCrLf
        SQL = SQL & ", M.VENDER_CD                          " & vbCrLf
        SQL = SQL & ", M.PROD_LINE_FA                       " & vbCrLf
        SQL = SQL & ", M.PROD_SLIT_FA                       " & vbCrLf
        SQL = SQL & ", M.PROD_CONTROL_YN                    " & vbCrLf
        SQL = SQL & ", M.PROD_PCN_NO                        " & vbCrLf
        SQL = SQL & ", L.PROD_MAX_TOT                       " & vbCrLf
        SQL = SQL & ", L.LABEL_PRT_SIDE                     " & vbCrLf
        SQL = SQL & ", (SELECT PROD_MAX_TOT                 " & vbCrLf
        SQL = SQL & "     FROM LBL_LABEL_MASTER             " & vbCrLf
        SQL = SQL & "    WHERE PROD_CD = '" & strProdCd & "'" & vbCrLf
        SQL = SQL & "      AND COMP_CD = '" & strCompCd & "'" & vbCrLf
        SQL = SQL & "      AND PROD_LABEL_TYPE = 'P') AS PPMAXCNT " & vbCrLf
        SQL = SQL & "  FROM LBL_M_PROD M                    " & vbCrLf
        SQL = SQL & "     , LBL_LABEL_MASTER L              " & vbCrLf
        SQL = SQL & " WHERE M.PROD_CD  = '" & strProdCd & "'" & vbCrLf
        SQL = SQL & "   AND M.COMP_CD  = '" & strCompCd & "'" & vbCrLf
        SQL = SQL & "   AND M.USED_YN  = 'Y'                " & vbCrLf
        SQL = SQL & "   AND M.PROD_CD = L.PROD_CD           " & vbCrLf
        SQL = SQL & "   AND M.COMP_CD = L.COMP_CD           " & vbCrLf
        SQL = SQL & "   AND L.PROD_LABEL_TYPE = 'I'         " & vbCrLf
   
        Set pAdoRS2 = New ADODB.Recordset
        Call GetRecordset(AdoCn, SQL, pAdoRS2, "")
        If Not pAdoRS2 Is Nothing Then
            If Not pAdoRS2.EOF Then
                'strBarData = strBarData & Trim(pAdoRS2("PROD_MATERIAL_CD") & "")
                strBarData = pAdoRS2("PROD_MATERIAL_CD") & ""
                txtMatCd.Text = strBarData
                strProdName = pAdoRS2("PROD_NAME") & ""
                strProdLen = pAdoRS2("PROD_LENGTH") & ""
                strProdMatCd = pAdoRS2("PROD_MATERIAL_CD") & ""
                strExMonth = pAdoRS2("EXPIR_MONTH") & ""
                txtExMonth.Text = strExMonth
                strExDate = DateAdd("m", strExMonth, strDate)
                strExDate = DateAdd("d", -1, strExDate)
                strProdTemp = pAdoRS2("PROD_STOR_TEMP") & ""
                strProdSize = pAdoRS2("PROD_SIZE") & ""
                strChimeiCd = pAdoRS2("PROD_CHIMEI_PN") & ""
                strVendorCd = pAdoRS2("VENDER_CD") & ""
                strProdLineFA = pAdoRS2("PROD_LINE_FA") & ""
                strProdSlitFA = pAdoRS2("PROD_SLIT_FA") & ""
                strContYN = pAdoRS2("PROD_CONTROL_YN") & ""
                strPcnNO = pAdoRS2("PROD_PCN_NO") & ""
                strMaxTot = pAdoRS2("PROD_MAX_TOT") & ""
                txtMaxTot.Text = strMaxTot
                strPPMaxTot = pAdoRS2("PPMAXCNT") & ""    'PP BOX 수량
                txtPPMaxTot = strPPMaxTot
                strPrtSide = pAdoRS2("LABEL_PRT_SIDE") & ""
                If strPrtSide = "Y" Then
                    picSide.Visible = True
                    chkReelPrint.Value = "1"
                Else
                    picSide.Visible = False
                    chkReelPrint.Value = "0"
                End If
            End If
        End If
        pAdoRS2.Close
        Set pAdoRS2 = Nothing
        '-- 제품정보 찾아오기
        
        Do Until AdoRs.EOF
            With spdPrtReelDetail
                .MaxRows = .MaxRows + 1
                strGu = AdoRs.Fields("LABEL_ITEM_GU").Value & ""
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 1)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 3)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_GU").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_FONT").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReelDetail, txtSlittingNo.Text, .MaxRows, 9)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "", .MaxRows, 10)
    
                '============================== 제품별로 분기 ==============================
                Select Case strPrtLabelName
                    '============== TP203C(ACF) ======================================================================
                    Case "P0001", "P0002"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "I")
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드":       ' strBarData = strBarData
                                        Case "유효기간_년":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                                        Case "유효기간_월":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                                        Case "유효기간_일":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                                        Case "N/A":             strBarData = strBarData & "0000"
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        Case "제품수량":        strBarData = strBarData & Format(CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text), "00000")
                                    End Select
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Provider" Then
                            Call SetText(spdPrtReelDetail, gKUKDO.COMPNM, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Storage Temperature" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Production Date" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Total Quantity/Length" Then
                            strContents = CCur(strMaxTot) & "Reels/" & CCur(strProdLen) * CCur(strMaxTot) * 100 & "cm"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material Code" Then
                            Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                            strContents = strChimeiCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Chimei P/N" Then
                            Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                            strContents = strChimeiCd
                            
                        End If
                        
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 4365

                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                            End If
                            strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If
                        strContents = ""
                    
                    '============== CF-TP408A ======================================================================
                    Case "P0003"

                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "I")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드"
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "ACF고객사":       strBarData = strBarData & "K"
                                        Case "제품순번":        strBarData = strBarData & "0001"
                                        Case "제품길이":        strBarData = strBarData & Format(strProdLen, "0000")
                                        Case "고객사내용":      strBarData = strBarData & "00"
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            strContents = CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 3000
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            imgBar2.Visible = True
                            imgBar2.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar2.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar2.WIDTH = 2000
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                        
                    '============== CF-TP400E ======================================================================
                    Case "P0004", "P0005", "P0008", "P0009"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "I")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드"
                                        Case "Vendor코드":          strBarData = strBarData & strVendorCd
                                        Case "제조라인공장":        strBarData = strBarData & strProdLineFA
                                        Case "Sliting공장":         strBarData = strBarData & strProdSlitFA
                                        Case "관리선이탈여부":      strBarData = strBarData & strContYN
                                        Case "PCN차수":             strBarData = strBarData & strPcnNO
                                        Case "포장일_년":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "포장일_월":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "포장일_일":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "대BOX포장순서":       strBarData = strBarData & "001"
                                        Case "대BOX내 REEL총길이":  strBarData = strBarData & Format(CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) * 270, "00000")
                                        Case "유효기간":            strBarData = strBarData & strExMonth
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            strProdMatCd = Mid(strProdMatCd, 1, 4) & "-" & Mid(strProdMatCd, 5)
                            
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            strContents = CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "00" & ")"
                            'strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgQrBar.Visible = True
                            imgQrBar.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgQrBar.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgQrBar.WIDTH = 1000
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                                                        
                    '============== CF-TP400E ======================================================================
                    Case "P0006", "P0007", "P0010"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "I")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "ACF":         strBarData = "C"
                                        Case "생산LOT":     strBarData = strBarData & strLotNo
                                        Case "_":           strBarData = strBarData & "_"
                                        Case "P/N":         strBarData = strBarData & "101"
                                        'Case "유효기간":    strBarData = strBarData & Format(strExDate, "yyyymmdd")
                                    
                                    
                                        Case "자재코드":            strBarData = Replace(strBarData, "-", "")
                                        Case "Vendor코드":          strBarData = strBarData & strVendorCd
                                        Case "제조라인공장":        strBarData = strBarData & strProdLineFA
                                        Case "Sliting공장":         strBarData = strBarData & strProdSlitFA
                                        Case "관리선이탈여부":      strBarData = strBarData & strContYN
                                        Case "PCN차수":             strBarData = strBarData & strPcnNO
                                        Case "포장일_년":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "포장일_월":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "포장일_일":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "대BOX포장순서":       strBarData = strBarData & "001"
                                        Case "대BOX내 REEL총길이":  strBarData = strBarData & Format(CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) * 270, "00000")
                                        Case "유효기간":            strBarData = strBarData & strExMonth
                                    
'                                        Case "자재코드"
'                                        Case "Vendor코드":          strBarData = strBarData & strVendorCd
'                                        Case "제조라인공장":        strBarData = strBarData & strProdLineFA
'                                        Case "Sliting공장":         strBarData = strBarData & strProdSlitFA
'                                        Case "관리선이탈여부":      strBarData = strBarData & strContYN
'                                        Case "PCN차수":             strBarData = strBarData & strPcnNO
'                                        Case "포장일_년":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
'                                        Case "포장일_월":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
'                                        Case "포장일_일":           strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
'                                        Case "대BOX포장순서":       strBarData = strBarData & "001"
'                                        Case "대BOX내 REEL총길이":  strBarData = strBarData & Format(CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) * 270, "00000")
'                                        Case "유효기간":            strBarData = strBarData & strExMonth
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            strContents = CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                            strContents = strMaxTot & " Reel"
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            strProdMatCd = Mid(strProdMatCd, 1, 4) & "-" & Mid(strProdMatCd, 5)
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
'                            barReel.Alignment = bcACenter
'                            barReel.Style = msSCode128B
'                            barReel.Visible = True
'                            barReel.Caption = strContents
'                            barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
'                            barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
'                            barReel.WIDTH = 4365
                        
'                            imgBar1.Visible = True
'                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
'                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
'                            imgBar1.WIDTH = 3000
                        
                            imgQrBar.Visible = True
                            imgQrBar.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgQrBar.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgQrBar.WIDTH = 1000
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
'                            barPart.Alignment = bcACenter
'                            barPart.Style = msSCode128B
'                            barPart.Visible = True
'                            barPart.Caption = strContents
'                            barPart.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
'                            barPart.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
'                            barPart.WIDTH = 3000
                        
                            imgBar2.Visible = True
                            imgBar2.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar2.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar2.WIDTH = 2000
                        
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '항목명 출력여부
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                                                        
                                                        
                End Select
                
                AdoRs.MoveNext
            End With
        Loop
    End If
    
    AdoRs.Close


    If lblstrPrtLabelName.Caption = "P0003" Then
        fra408a.Visible = True
    ElseIf lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
        fra500e.Visible = True
    Else
        fra408a.Visible = False
        fra500e.Visible = False
    End If
    
    txtReelBarcode.SetFocus

'    If strPrtLabelName = "P0006" Or strPrtLabelName = "P0007" Then
'        For i = 1 To 12
'            If i = 1 Then
'                txtReelBarcode.Text = strProdMatCd
'            Else
'                txtReelBarcode.Text = strLotNo
'
'            End If
'            Call txtReelBarcode_KeyPress(vbKeyReturn)
'
'        Next
'    End If
'
'    txtReelBarcode.SetFocus

'    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub




Private Sub txtReelBarcode_KeyPress(KeyAscii As Integer)
    Dim pAdoRS1         As ADODB.Recordset
    Dim pAdoRS2         As ADODB.Recordset
    Dim pAdoRS3         As ADODB.Recordset
    Dim i               As Integer
    Dim strDate         As String
    Dim strProdLen      As String
    Dim strContents     As String
    Dim strASReelCnt    As String
    Dim strData         As Variant
    Dim lngData         As Long
    Dim strBarData      As String
    Dim strBarDate      As String
    Dim strBarExDate    As String
    Dim strExMonth      As String
    Dim strVendorCd     As String
    Dim strProdLineFA   As String
    Dim strProdSlitFA   As String
    Dim strContYN       As String
    Dim strPcnNO        As String
    Dim intMaxNo        As Integer
    
    strBarData = ""
    
    If KeyAscii = vbKeyReturn Then
        If txtReelBarcode.Text <> "" Then
            txtReelBarcode.Text = UCase(txtReelBarcode.Text)
            If Len(gPackTrack.ORDERDT) = 10 Then
                strDate = Format(gPackTrack.ORDERDT, "yyyymmdd")
            Else
                strDate = gPackTrack.ORDERDT
            End If
            '트래킹 테이블에 존재하는지 체크
            Set AdoRs = Get_Pack_Track(strDate, gPackTrack.PRODCD, "", txtReelBarcode.Text, "")
        
            If AdoRs.RecordCount = 0 Then
                '스캔한 바코드가 다른 날짜에 있는지 찾는다. Lot번호 가져오기
                Set pAdoRS1 = Get_Pack_Track_LotNo2(gPackTrack.PRODCD, "", txtReelBarcode.Text, "")
                If pAdoRS1 Is Nothing Then
                    '등록된 정보 없음
                    txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 발행된 PP Box 바코드가 아닙니다." & vbCrLf
                    txtReelBarcode.SelStart = 0
                    txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                    Exit Sub
                Else
                    If pAdoRS1.EOF Then
                        '등록된 정보 없음
                        txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 발행된 PP Box 바코드가 아닙니다." & vbCrLf
                        txtReelBarcode.SelStart = 0
                        txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                        Exit Sub
                    End If
                    
                    If txtLotNo2.Text <> "" And txtLotNo2.Text <> pAdoRS1.Fields("PROD_LOT_NO").Value & "" Then
                        txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 Lot No가 다릅니다." & vbCrLf
                        txtMsg.Text = txtMsg.Text & "허용되는 Lot 수량은 박스당 최대 2개까지입니다." & vbCrLf
                        Exit Sub
                    End If
                    fraLotNo2.Visible = True
                    
                    txtLotNo2.Text = pAdoRS1.Fields("PROD_LOT_NO").Value & ""
                    txtSlittingNo2.Text = pAdoRS1.Fields("SLITING_NO").Value & ""
                    txtProdOrderDt2.Text = Format(pAdoRS1.Fields("PROD_ORDER_DT").Value & "", "####-##-##")
                    strBarDate = txtProdOrderDt2.Text
                    strExMonth = txtExMonth.Text
                    
                    txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 Lot No가 다릅니다." & vbCrLf

                    spdScan2.MaxRows = spdScan2.MaxRows + 1
                    Call SetText(spdScan2, txtReelBarcode.Text, spdScan2.MaxRows, 1)
                    Call SetText(spdScan2, pAdoRS1.Fields("PROD_REEL_BAR").Value & "", spdScan2.MaxRows, 2)
                    Call SetText(spdScan2, pAdoRS1.Fields("PROD_PP_BAR").Value & "", spdScan2.MaxRows, 3)
                    Call SetText(spdScan2, pAdoRS1.Fields("PROD_ICE_BAR").Value & "", spdScan2.MaxRows, 4)
                    Call SetText(spdScan2, pAdoRS1.Fields("PROD_PP_BAR_IN").Value & "", spdScan2.MaxRows, 5)
                    Call SetText(spdScan2, pAdoRS1.Fields("PROD_ICE_BAR_IN").Value & "", spdScan2.MaxRows, 6)
                    
                    '=========== LotNo2 =====================================================================
                    If lblstrPrtLabelName.Caption = "P0003" Then
                        fra408a2.Visible = True
                    ElseIf lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
                        fra500e2.Visible = True
                    Else
                        fra408a2.Visible = False
                        fra500e2.Visible = False
                    End If
                    
                    Set pAdoRS3 = Get_ProdMaster(lblstrPrtLabelName.Caption, txtCompCd.Text)
                    
                    strVendorCd = pAdoRS3.Fields("VENDER_CD").Value & ""
                    strProdLineFA = pAdoRS3.Fields("PROD_LINE_FA").Value & ""
                    strProdSlitFA = pAdoRS3.Fields("PROD_SLIT_FA").Value & ""
                    strContYN = pAdoRS3.Fields("PROD_CONTROL_YN").Value & ""
                    strPcnNO = pAdoRS3.Fields("PROD_PCN_NO").Value & ""
                    
                    pAdoRS3.Close
                    
                    'MAX NO
                    Set AdoRs = Get_MAX_NO(strBarDate, lblstrPrtLabelName.Caption, "I")
                    If AdoRs.RecordCount = 0 Then
                        'INSERT
                        intMaxNo = 1
                        If Set_MAX_NO("IN", "I", intMaxNo) Then
                        End If
                    Else
                        'UPDATE
                        intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
                        intMaxNo = intMaxNo + 1
                        If Set_MAX_NO("UP", "I", intMaxNo) Then
                        End If
                    End If
                    AdoRs.Close
                    
                    If spdPrtReelDetail2.MaxRows = 0 Then
                        Call spdPrtReelDetail.SetSelection(1, 1, spdPrtReelDetail.MaxCols, spdPrtReelDetail.MaxRows)
                        '클립보드 카피
                        spdPrtReelDetail.ClipboardCopy
                        spdPrtReelDetail2.MaxRows = spdPrtReelDetail.MaxRows
                        spdPrtReelDetail2.MaxCols = spdPrtReelDetail.MaxCols
                        spdPrtReelDetail2.BlockMode = True
                        spdPrtReelDetail2.Col = 1
                        spdPrtReelDetail2.Row = 1
                        spdPrtReelDetail2.Col2 = spdPrtReelDetail2.MaxCols
                        spdPrtReelDetail2.Row2 = spdPrtReelDetail2.Row2
                        spdPrtReelDetail2.ClipboardPaste
                        spdPrtReelDetail2.BlockMode = False
                        
                        With spdPrtReelDetail2
                            .Row = 1:       .Row2 = .MaxRows
                            .Col = 1:       .Col2 = .MaxCols
                            .BlockMode = True
                            .Action = ActionClearText
                            .Clip = Clipboard.GetText()
                            .ClipboardPaste
                            .RowHeight(-1) = 12
                            .BlockMode = False
                            Clipboard.Clear
                        End With
                        
                        strASReelCnt = pAdoRS1.RecordCount
                    End If
                    
                    For i = 1 To spdPrtReelDetail2.MaxRows
                        Select Case lblstrPrtLabelName.Caption
                            Case "P0001", "P0002"
                                '제품수량/길이
                                If GetText(spdPrtReelDetail2, i, 1) = "Total Quantity/Length" Then
                                    If spdScan2.MaxRows > 1 Then
                                        strASReelCnt = GetText(spdPrtReelDetail2, i, 3)
                                        strASReelCnt = mGetP(strASReelCnt, 1, "/")
                                        strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                        strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                        strASReelCnt = CCur(strASReelCnt) + CCur(pAdoRS1.RecordCount)
                                    Else
                                        strASReelCnt = pAdoRS1.RecordCount
                                    End If
                                    strContents = CCur(strASReelCnt) & "Reels/" & CCur(strASReelCnt) * (CCur((txtProdLen.Text)) * 100) & "cm"
                                    Call SetText(spdPrtReelDetail2, strContents, i, 3)
                                
                                '바코드
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "바코드" Then
                                    strBarExDate = DateAdd("m", strExMonth, strBarDate)
                                    strBarExDate = DateAdd("d", -1, strBarExDate)
                                    Set pAdoRS2 = Get_BarDetail_Prt(lblstrPrtLabelName.Caption, txtCompCd.Text, "I")
                                    If pAdoRS2 Is Nothing Then
                                        '등록된 정보 없음
                                    Else
                                        strBarData = ""
                                        Do Until pAdoRS2.EOF
                                            Select Case pAdoRS2.Fields("BAR_ITEM_NAME").Value & ""
                                                Case "자재코드":        strBarData = txtMatCd.Text
                                                Case "유효기간_년":     strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Year(strBarExDate))
                                                Case "유효기간_월":     strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strBarExDate))
                                                Case "유효기간_일":     strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Day(strBarExDate))
                                                Case "N/A":             strBarData = strBarData & "0000"
                                                Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Year(strBarDate))
                                                Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strBarDate))
                                                Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Day(strBarDate))
                                                Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo2.Text, "00")
                                                Case "제품수량":        strBarData = strBarData & Format(CCur(strASReelCnt), "00000")
                                            End Select
                                            pAdoRS2.MoveNext
                                        Loop
                                        pAdoRS2.Close
                                    End If
                                    Call SetText(spdPrtReelDetail2, strBarData, i, 3)
                                '유효기간
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "Expiration Date" Then
                                    strBarExDate = DateAdd("m", strExMonth, strBarDate)
                                    strBarExDate = DateAdd("d", -1, strBarExDate)
                                    Call SetText(spdPrtReelDetail2, strBarExDate, i, 3)
                                End If
                            Case "P0003"
                                fra408a2.Visible = True
                                '제품수량
                                If GetText(spdPrtReelDetail2, i, 1) = "수량" Then
                                    If spdScan2.MaxRows > 1 Then
                                        strASReelCnt = GetText(spdPrtReelDetail2, i, 3)
                                        strASReelCnt = mGetP(strASReelCnt, 1, "/")
                                        strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                        strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                        strASReelCnt = CCur(strASReelCnt) + CCur(pAdoRS1.RecordCount)
                                    Else
                                        strASReelCnt = pAdoRS1.RecordCount
                                    End If
                                    strContents = CCur(strASReelCnt) & " Reel"
                                    Call SetText(spdPrtReelDetail2, strContents, i, 3)
                                'LotNo
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "LotNo" Then
                                    Call SetText(spdPrtReelDetail2, txtLotNo2.Text, i, 3)
                                '유효기간
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "생산일자" Then
                                    Call SetText(spdPrtReelDetail2, strBarDate, i, 3)
                                '생산일자
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "유효기간" Then
                                    strBarExDate = DateAdd("m", strExMonth, strBarDate)
                                    strBarExDate = DateAdd("d", -1, strBarExDate)
                                    Call SetText(spdPrtReelDetail2, strBarExDate, i, 3)
                                End If
                            Case "P0004", "P0005", "P0008", "P0009"
                                '제품수량
                                If GetText(spdPrtReelDetail2, i, 1) = "수량" Then
                                    If spdScan2.MaxRows > 1 Then
                                        strASReelCnt = GetText(spdPrtReelDetail2, i, 3)
                                        strASReelCnt = mGetP(strASReelCnt, 1, "/")
                                        strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                        strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                        strASReelCnt = CCur(strASReelCnt) + CCur(pAdoRS1.RecordCount)
                                    Else
                                        strASReelCnt = pAdoRS1.RecordCount
                                    End If
                                    strContents = CCur(strASReelCnt) & " Reel"
                                    Call SetText(spdPrtReelDetail2, strContents, i, 3)
                                'LotNo
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "LotNo" Then
                                    Call SetText(spdPrtReelDetail2, txtLotNo2.Text, i, 3)
                                '유효기간
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "생산일자" Then
                                    Call SetText(spdPrtReelDetail2, strBarDate, i, 3)
                                '생산일자
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "유효기간" Then
                                    strBarExDate = DateAdd("m", strExMonth, strBarDate)
                                    strBarExDate = DateAdd("d", -1, strBarExDate)
                                    Call SetText(spdPrtReelDetail2, strBarExDate, i, 3)
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "바코드" Then
                                    Set pAdoRS2 = Get_BarDetail_Prt(lblstrPrtLabelName.Caption, txtCompCd.Text, "I")
                                    If pAdoRS2 Is Nothing Then
                                        '등록된 정보 없음
                                    Else
                                        Do Until pAdoRS2.EOF
                                            Select Case pAdoRS2.Fields("BAR_ITEM_NAME").Value & ""
                                                Case "자재코드":            strBarData = txtMatCd.Text
                                                Case "Vendor코드":          strBarData = strBarData & strVendorCd
                                                Case "제조라인공장":        strBarData = strBarData & strProdLineFA
                                                Case "Sliting공장":         strBarData = strBarData & strProdSlitFA
                                                Case "관리선이탈여부":      strBarData = strBarData & strContYN
                                                Case "PCN차수":             strBarData = strBarData & strPcnNO
                                                Case "포장일_년":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Year(strBarDate))
                                                Case "포장일_월":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strBarDate))
                                                Case "포장일_일":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Day(strBarDate))
                                                Case "대BOX포장순서":       strBarData = strBarData & Format(intMaxNo, "000")
                                                Case "대BOX내 REEL총길이":  strBarData = strBarData & Format(CCur(strASReelCnt) * 270, "00000")
                                                Case "유효기간":            strBarData = strBarData & strExMonth
                                            End Select
                                
                                            pAdoRS2.MoveNext
                                        Loop
                                        pAdoRS2.Close
                                        txtICEBoxNo2.Text = intMaxNo
                                    End If
                                    Call SetText(spdPrtReelDetail2, strBarData, i, 3)
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "바코드값" Then
                                    Call SetText(spdPrtReelDetail2, strBarData, i, 3)
                                End If
                            Case "P0006", "P0007", "P0010"
                                '제품수량
                                If GetText(spdPrtReelDetail2, i, 1) = "수량" Then
                                    If spdScan2.MaxRows > 1 Then
                                        strASReelCnt = GetText(spdPrtReelDetail2, i, 3)
                                        strASReelCnt = mGetP(strASReelCnt, 1, "/")
                                        strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                        strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                        strASReelCnt = CCur(strASReelCnt) + CCur(pAdoRS1.RecordCount)
                                    Else
                                        strASReelCnt = pAdoRS1.RecordCount
                                    End If
                                    strContents = CCur(strASReelCnt) & " Reel"
                                    Call SetText(spdPrtReelDetail2, strContents, i, 3)
                                    txtQty2.Text = strASReelCnt
                                    
                                'LotNo
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "LotNo" Then
                                    Call SetText(spdPrtReelDetail2, txtLotNo2.Text, i, 3)
                                '유효기간
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "생산일자" Then
                                    Call SetText(spdPrtReelDetail2, strBarDate, i, 3)
                                '생산일자
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "유효기간" Then
                                    strBarExDate = DateAdd("m", strExMonth, strBarDate)
                                    strBarExDate = DateAdd("d", -1, strBarExDate)
                                    Call SetText(spdPrtReelDetail2, strBarExDate, i, 3)
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "바코드" Then
                                    Set pAdoRS2 = Get_BarDetail_Prt(lblstrPrtLabelName.Caption, txtCompCd.Text, "I")
                                    If pAdoRS2 Is Nothing Then
                                        '등록된 정보 없음
                                    Else
                                        Do Until pAdoRS2.EOF
                                            Select Case pAdoRS2.Fields("BAR_ITEM_NAME").Value & ""
                                                Case "자재코드":            strBarData = txtMatCd.Text
                                                Case "Vendor코드":          strBarData = strBarData & strVendorCd
                                                Case "제조라인공장":        strBarData = strBarData & strProdLineFA
                                                Case "Sliting공장":         strBarData = strBarData & strProdSlitFA
                                                Case "관리선이탈여부":      strBarData = strBarData & strContYN
                                                Case "PCN차수":             strBarData = strBarData & strPcnNO
                                                Case "포장일_년":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Year(strBarDate))
                                                Case "포장일_월":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strBarDate))
                                                Case "포장일_일":           strBarData = strBarData & Get_YMD(pAdoRS2.Fields("LABEL_ITEM_TYPE").Value & "", Day(strBarDate))
                                                Case "대BOX포장순서":       strBarData = strBarData & Format(intMaxNo, "000")
                                                Case "대BOX내 REEL총길이":  strBarData = strBarData & Format(CCur(strASReelCnt) * 350, "00000")
                                                Case "유효기간":            strBarData = strBarData & strExMonth
                                            End Select
                                
                                            pAdoRS2.MoveNext
                                        Loop
                                        pAdoRS2.Close
                                        txtICEBoxNo2.Text = intMaxNo
                                        
                                    End If
                                    Call SetText(spdPrtReelDetail2, strBarData, i, 3)
                                ElseIf GetText(spdPrtReelDetail2, i, 1) = "바코드값" Then
                                    Call SetText(spdPrtReelDetail2, strBarData, i, 3)
                                ElseIf GetText(spdPrtReelDetail, i, 1) = "Material code" Then
                                    txtPartsID2.Text = GetText(spdPrtReelDetail, i, 3)
                                End If
                        End Select
                    Next
                    
                    txtReelBarcode.Text = ""
                    
                    Exit Sub
                    '=========== LotNo2 =====================================================================
                End If
                
            Else
                If AdoRs.Fields("PROD_ICE_BAR").Value & "" <> "" Then
                    txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 트래킹된 PP Box 바코드입니다." & vbCrLf
                    txtReelBarcode.SelStart = 0
                    txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                    Exit Sub
                End If
            End If
            
            With spdScan
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    If .Text = txtReelBarcode.Text Then
                        'MsgBox txtReelBarcode.Text & "와 같은 바코드가 있습니다.", vbOKOnly + vbInformation, Me.Caption
                        txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "와 같은 바코드가 있습니다." & vbCrLf
                        txtReelBarcode.SelStart = 0
                        txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                        Exit Sub
                    End If
                Next
                
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdScan, txtReelBarcode.Text, .MaxRows, 1)
                Call SetText(spdScan, AdoRs.Fields("PROD_REEL_BAR").Value & "", .MaxRows, 2)
                Call SetText(spdScan, AdoRs.Fields("PROD_PP_BAR").Value & "", .MaxRows, 3)
                Call SetText(spdScan, AdoRs.Fields("PROD_ICE_BAR").Value & "", .MaxRows, 4)
                Call SetText(spdScan, AdoRs.Fields("PROD_PP_BAR_IN").Value & "", .MaxRows, 5)
                Call SetText(spdScan, AdoRs.Fields("PROD_ICE_BAR_IN").Value & "", .MaxRows, 6)
                
                'PP Box당 갖고있는 Reel 수량
                txtPPMaxTot.Text = AdoRs.RecordCount
            
            End With
            
            'LotNo1
            imgPpBar(i - 1).Visible = True
            txtScanCount.Text = txtScanCount.Text + 1
            txtReelBarcode.Text = ""
            
            For i = 1 To spdPrtReelDetail.MaxRows
                Select Case lblstrPrtLabelName.Caption
                    Case "P0001", "P0002"
                        '제품수량/길이
                        If GetText(spdPrtReelDetail, i, 1) = "Total Quantity/Length" Then
                            If txtScanCount.Text > 1 Then
                                strASReelCnt = GetText(spdPrtReelDetail, i, 3)
                                strASReelCnt = mGetP(strASReelCnt, 1, "/")
                                strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                strASReelCnt = CCur(strASReelCnt) + CCur(txtPPMaxTot.Text)
                            Else
                                strASReelCnt = txtPPMaxTot.Text
                            End If
                            strContents = CCur(strASReelCnt) & "Reels/" & CCur(strASReelCnt) * (CCur((txtProdLen.Text)) * 100) & "cm"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                        '바코드
                        If GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 17)
                            strContents = strProdLen & Format(CCur(strASReelCnt), "00000")
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case "P0003"
                        '제품수량/길이
                        If GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            If txtScanCount.Text > 1 Then
                                strASReelCnt = GetText(spdPrtReelDetail, i, 3)
                                strASReelCnt = Replace(strASReelCnt, "Reels", "")
                                strASReelCnt = Replace(strASReelCnt, "Reel", "")
                                strASReelCnt = CCur(strASReelCnt) + CCur(txtPPMaxTot.Text)
                            Else
                                strASReelCnt = txtPPMaxTot.Text
                            End If
                            strContents = CCur(strASReelCnt) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case "P0004", "P0005", "P0008", "P0009"
                        If GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            '제일 처음 온다는 조건하에..
                            If txtScanCount.Text > 1 Then
                                strASReelCnt = GetText(spdPrtReelDetail, i, 3)
                                strASReelCnt = Trim(Replace(strASReelCnt, "Reel", ""))
                                strASReelCnt = CCur(strASReelCnt) + CCur(txtPPMaxTot.Text)
                            Else
                                strASReelCnt = txtPPMaxTot.Text
                            End If
                            strContents = CCur(strASReelCnt) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 21) & Format(txtICEBoxNo.Text, "000")
                            strContents = strProdLen & Format(CCur(strASReelCnt) * 270, "00000") & Right(GetText(spdPrtReelDetail, i, 3), 1)
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드값" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 21) & Format(txtICEBoxNo.Text, "000")
                            strContents = strProdLen & Format(CCur(strASReelCnt) * 270, "00000") & Right(GetText(spdPrtReelDetail, i, 3), 1)
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case "P0006", "P0007", "P0010"
                        If GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            '제일 처음 온다는 조건하에..
                            If txtScanCount.Text > 1 Then
                                strASReelCnt = GetText(spdPrtReelDetail, i, 3)
                                strASReelCnt = Trim(Replace(strASReelCnt, "Reel", ""))
                                strASReelCnt = CCur(strASReelCnt) + CCur(txtPPMaxTot.Text)
                            Else
                                strASReelCnt = txtPPMaxTot.Text
                            End If
                            txtQty.Text = strASReelCnt
                            strContents = CCur(strASReelCnt) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 21) & Format(txtICEBoxNo.Text, "000")
                            strContents = strProdLen & Format(CCur(strASReelCnt) * 350, "00000") & Right(GetText(spdPrtReelDetail, i, 3), 1)
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드값" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 21) & Format(txtICEBoxNo.Text, "000")
                            strContents = strProdLen & Format(CCur(strASReelCnt) * 350, "00000") & Right(GetText(spdPrtReelDetail, i, 3), 1)
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "Material code" Then
                            txtPartsID.Text = GetText(spdPrtReelDetail, i, 3)
                        End If
                End Select
            Next
            
            If txtMaxTot.Text = txtScanCount.Text Then
                If chAutoPrint.Value = "1" Then
                    Call cmdPrint_Click
                End If
            End If
            
        End If
    End If
    
End Sub

