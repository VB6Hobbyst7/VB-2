VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Begin VB.Form frmPrtReel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reel 라벨 출력"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   19920
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   8205
      Left            =   90
      TabIndex        =   16
      Top             =   1050
      Width           =   19395
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   7725
         Left            =   8160
         TabIndex        =   17
         Top             =   300
         Width           =   10995
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   4185
            Left            =   5580
            ScaleHeight     =   4125
            ScaleWidth      =   5055
            TabIndex        =   37
            Top             =   1440
            Width           =   5115
            Begin BarcodLib.Barcod barReel 
               Height          =   555
               Left            =   300
               TabIndex        =   70
               Top             =   300
               Visible         =   0   'False
               Width           =   4365
               _Version        =   65543
               _ExtentX        =   7699
               _ExtentY        =   979
               _StockProps     =   75
               Caption         =   "2X2707R0202001P10110000"
               BarWidth        =   0
               Direction       =   0
               Style           =   7
               UPCNotches      =   3
               Alignment       =   2
               Extension       =   ""
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
               Index           =   15
               Left            =   900
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               Index           =   7
               Left            =   0
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
               Index           =   6
               Left            =   0
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
               Index           =   5
               Left            =   0
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
               Index           =   4
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
               Index           =   3
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
               Index           =   2
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
               Index           =   1
               Left            =   0
               TabIndex        =   55
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
               TabIndex        =   54
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
               TabIndex        =   53
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
               TabIndex        =   52
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
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
               Index           =   8
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
               Index           =   7
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
               Index           =   6
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
               Index           =   5
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
               Index           =   4
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
               Index           =   3
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
               Index           =   2
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
               Index           =   1
               Left            =   0
               TabIndex        =   40
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
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
         End
         Begin VB.TextBox txtP2To 
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
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   11
            Top             =   6720
            Width           =   1000
         End
         Begin VB.TextBox txtP2From 
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
            Left            =   4890
            MaxLength       =   5
            TabIndex        =   10
            Top             =   6720
            Width           =   1000
         End
         Begin VB.TextBox txtP1To 
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
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   9
            Top             =   6300
            Width           =   1000
         End
         Begin VB.TextBox txtP1From 
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
            Left            =   4890
            MaxLength       =   5
            TabIndex        =   8
            Top             =   6300
            Width           =   1000
         End
         Begin VB.TextBox txtReelQTY 
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
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   7
            Top             =   6330
            Width           =   1500
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
            Left            =   9390
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   30
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtProdPosNo 
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
            Left            =   9390
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   29
            Top             =   360
            Width           =   1305
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
            Left            =   5580
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   28
            Top             =   810
            Width           =   2205
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
            Left            =   5580
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   27
            Top             =   360
            Width           =   2205
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
            Width           =   2205
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
            TabIndex        =   19
            Top             =   360
            Width           =   2205
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   795
            Left            =   8250
            TabIndex        =   18
            Top             =   6570
            Width           =   2475
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
               Left            =   150
               Style           =   1  '그래픽
               TabIndex        =   12
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
               Left            =   1260
               Style           =   1  '그래픽
               TabIndex        =   13
               Top             =   150
               Width           =   1095
            End
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   4125
            Left            =   240
            TabIndex        =   6
            Top             =   1440
            Width           =   5325
            _Version        =   393216
            _ExtentX        =   9393
            _ExtentY        =   7276
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
            SpreadDesigner  =   "frmPrtReel.frx":0000
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label Label4 
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
            Left            =   5910
            TabIndex        =   36
            Top             =   6750
            Width           =   195
         End
         Begin VB.Label lblP2 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "P2"
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
            Left            =   3840
            TabIndex        =   35
            Top             =   6720
            Width           =   1005
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "2 매"
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
            Left            =   7290
            TabIndex        =   34
            Top             =   6810
            Width           =   435
         End
         Begin VB.Label Label2 
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
            Left            =   5910
            TabIndex        =   33
            Top             =   6330
            Width           =   195
         End
         Begin VB.Label lblP1 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "P1"
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
            Left            =   3840
            TabIndex        =   32
            Top             =   6300
            Width           =   1005
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Reel 수량"
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
            TabIndex        =   31
            Top             =   6330
            Width           =   1500
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
            TabIndex        =   25
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
            Left            =   4050
            TabIndex        =   24
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "공정 No"
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
            Left            =   7860
            TabIndex        =   23
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Slitting No"
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
            Left            =   7860
            TabIndex        =   22
            Top             =   810
            Width           =   1500
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
            Left            =   4050
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   360
            Width           =   1500
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   7635
         Left            =   210
         TabIndex        =   5
         Top             =   390
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   13467
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
         MaxCols         =   19
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtReel.frx":0B4C
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
      Width           =   19425
      Begin MSCommLib.MSComm comEqp 
         Left            =   9210
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1650
         TabIndex        =   1
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
         Format          =   60293121
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   3750
         TabIndex        =   2
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
         Format          =   60293121
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
               Picture         =   "frmPrtReel.frx":1ABE
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":2058
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":25F2
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":2B8C
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":341E
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":3578
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":36D2
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":382C
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":4106
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtReel.frx":49E0
         Top             =   420
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11685
         TabIndex        =   71
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   420
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPrtReel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmPrtReel.frm
'   작성자  : 오세원
'   내  용  : Reel 라벨출력
'   작성일  : 2020-02-24
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReel.MaxRows = 0
    spdPrtReelDetail.MaxRows = 0
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    
    txtReelQTY.Text = ""
    txtP1From.Text = ""
    txtP1To.Text = ""
    txtP2From.Text = ""
    txtP2To.Text = ""
    
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
Private Sub GetReelOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)

    Dim strLabelType    As String

    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)

    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1

                Call SetText(spdPrtReel, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 3)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_NO").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReel, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReel, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReel, AdoRs.Fields("ROOL_INFO").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReel, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 10)
                'Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_VIEW").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 12)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 13)
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

Private Sub cmdPrint_Click()
    Dim intPrt      As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strPrtData()   As String
    Dim strType     As String
    Dim strXPos     As String
    Dim strXPos2    As String
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
    Dim J           As Integer
    Dim k           As Integer
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    i = 0
       
    strPlusXPos = 630
    'strAFont = "^ABN,24,12"
    strAFont = "^A1N,25,15"
    'strAFont = "^A0N,35,35"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^MD9"
    
'asc("~") : R_7E
'? Hex(126)
    
'asc("℃") : R_A1C9
 '?Hex(126)
    
    For intPrt = CCur(txtP1From.Text) To CCur(txtP1To.Text)
        strOutput = ""
        Erase strPrtData
        i = 0
        With spdPrtReelDetail
            For k = 1 To 2
                If k = 1 Then
                    For intRow = 1 To .MaxRows
                        '좌측 출력
                        strType = GetText(spdPrtReelDetail, intRow, 1)
                        strXPos = GetText(spdPrtReelDetail, intRow, 4)
                        strYPos = GetText(spdPrtReelDetail, intRow, 5)
                        strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                        strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                        strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                        strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                        strData = GetText(spdPrtReelDetail, intRow, 3)
                        
                        strData = Replace(strData, "~", "_7E")
                        strData = Replace(strData, "℃", "_A1C9")
                        
                        Select Case strRot
                            Case "0":   strRot = "N"
                            Case "90":  strRot = "R"
                            Case "180": strRot = "I"
                            Case "270": strRot = "B"
                        End Select
                        
                        If strType = "바코드" Then
                            strLotSub = "P" & strSlt & Format(intPrt, "00")
                            strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                            If Mid(strBarType, 1, 1) = "1" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
'                                strPrtData(i) = strPrtData(i) & "^A1" & strRot & ",25" & "," & "12"
                                'strPrtData(i)  = strPrtData(i)  & "^BY" & "1.5,4,60"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                
                                strXPos2 = strXPos
                                i = i + 1
                            Else
                                ReDim Preserve strPrtData(i) As String
                                'strPrtData(i)  = strPrtData(i)  & "^FO100,100^BQN,2,10^FDNM,AAC-42^FS" & vbLf
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
'                                strPrtData(i) = strPrtData(i) & "^A1" & strRot & "25" & "," & "12"
                                strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                strPrtData(i) = strPrtData(i) & "^BQ"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            End If
                        Else
                            If strType = "SDI ACF Lot" Then
                                ReDim Preserve strPrtData(i) As String
                                strLot = mGetP(strData, 1, "(")
                                strLot = strLot & "(" & strLotSub & ")"
                                
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont
                                strPrtData(i) = strPrtData(i) & "^FH"
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            Else
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont
                                strPrtData(i) = strPrtData(i) & "^FH"
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            End If
                        End If
                    Next
                Else
                    strXPos2 = strXPos - strXPos2
                    strXPos = CCur(strPlusXPos) + CCur(strXPos)         '일반
                    strXPos2 = CCur(strPlusXPos) + CCur(strXPos2)       'Reel 바코드

                    For intRow = 1 To .MaxRows
                        strType = GetText(spdPrtReelDetail, intRow, 1)
                        strData = GetText(spdPrtReelDetail, intRow, 3)
                        'strXPos = GetText(spdPrtReelDetail, intRow, 4)
                        strYPos = GetText(spdPrtReelDetail, intRow, 5)
                        strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                        strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                        strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                        strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                        strData = GetText(spdPrtReelDetail, intRow, 3)
                        
                        strData = Replace(strData, "~", "_7E")
                        strData = Replace(strData, "℃", "_A1C9")
                        
                        Select Case strRot
                            Case "0":   strRot = "N"
                            Case "90":  strRot = "R"
                            Case "180": strRot = "I"
                            Case "270": strRot = "B"
                        End Select
                        
                        If strType = "바코드" Then
                            strLotSub = "P" & strSlt & Format(intPrt, "00")
                            strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                            If Mid(strBarType, 1, 1) = "1" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos2 & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
'                                strPrtData(i) = strPrtData(i) & "^A1" & strRot & ",25" & "," & "12"
                                'strPrtData(i)  = strPrtData(i)  & "^BY" & "1.5,4,60"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            Else
                                ReDim Preserve strPrtData(i) As String
                                'strPrtData(i)  = strPrtData(i)  & "^FO100,100^BQN,2,10^FDNM,AAC-42^FS" & vbLf
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos2 & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
'                                strPrtData(i) = strPrtData(i) & "^A1" & strRot & "25" & "," & "12"
                                strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                strPrtData(i) = strPrtData(i) & "^BQ"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            End If
                        Else
                            If strType = "SDI ACF Lot" Then
                                strLot = mGetP(strData, 1, "(")
                                strLot = strLot & "(" & strLotSub & ")"
                                
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont '"^AD" & strRot & ",25" & "," & "12"
                                strPrtData(i) = strPrtData(i) & "^FH"
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            Else
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont '"^AD" & strRot & ",25" & "," & "12"
                                strPrtData(i) = strPrtData(i) & "^FH"
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            End If
                        End If
                    Next
                End If
            Next
        End With
    
        strOutput = ""
        For J = 0 To UBound(strPrtData)
            strOutput = strOutput & strPrtData(J)
        Next
        
        strOutput = strHeader & strOutput & "^XZ" & vbLf
        comEqp.Output = strOutput
        strOutput = ""
        If intPrt = 3 Then Exit Sub
    Next
    
'        With spdPrtReelDetail
'            For intRow = 1 To .MaxRows
'                '좌측 출력
'                strType = GetText(spdPrtReelDetail, intRow, 1)
'                strData = GetText(spdPrtReelDetail, intRow, 3)
'                strXPos = GetText(spdPrtReelDetail, intRow, 4)
'                strYPos = GetText(spdPrtReelDetail, intRow, 5)
'                strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
'                strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
'                strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
'                Select Case strRot
'                    Case "0":   strRot = "N"
'                    Case "90":  strRot = "R"
'                    Case "180": strRot = "I"
'                    Case "270": strRot = "B"
'                End Select
'                If strType = "바코드" Then
'                    If Mid(strBarType, 1, 1) = "1" Then
'                        strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                        strOutput = strOutput & "^CI26"
'                        strOutput = strOutput & "^A1" & strRot & ",25" & "," & "12"
'                        'strOutput = strOutput & "^BY" & "1.5,4,60"
'                        strOutput = strOutput & "^BC" & "N,60,Y,N,N"
'                        strOutput = strOutput & "^FD" & strData
'                        strOutput = strOutput & "^FS" & vbLf
'                    Else
'                        'strOutput = strOutput & "^FO100,100^BQN,2,10^FDNM,AAC-42^FS" & vbLf
'                        strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                        strOutput = strOutput & "^CI26"
'                        strOutput = strOutput & "^A1" & strRot & "25" & "," & "12"
'                        strOutput = strOutput & "^BY" & "2,4,50"
'                        strOutput = strOutput & "^BQ"
'                        strOutput = strOutput & "^FD" & strData
'                        strOutput = strOutput & "^FS" & vbLf
'                    End If
'                Else
'                    strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                    strOutput = strOutput & "^CI26"
'                    strOutput = strOutput & "^A1" & strRot & ",25" & "," & "12"
'                    strOutput = strOutput & "^FD" & strType & " : " & strData
'                    strOutput = strOutput & "^FS" & vbLf
'                End If
'
'                strPlusXPos = 600
'                strXPos = CCur(strXPos) + CCur(strPlusXPos)
'
'                '우측 출력
'                strType = GetText(spdPrtReelDetail, intRow, 1)
'                strData = GetText(spdPrtReelDetail, intRow, 3)
'                strXPos = GetText(spdPrtReelDetail, intRow, 4)
'                strYPos = GetText(spdPrtReelDetail, intRow, 5)
'                strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
'                strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
'                strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
'                Select Case strRot
'                    Case "0":   strRot = "N"
'                    Case "90":  strRot = "R"
'                    Case "180": strRot = "I"
'                    Case "270": strRot = "B"
'                End Select
'                If strType = "바코드" Then
'                    If Mid(strBarType, 1, 1) = "1" Then
'                        strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                        strOutput = strOutput & "^CI26"
'                        strOutput = strOutput & "^A1" & strRot & ",25" & "," & "12"
'                        'strOutput = strOutput & "^BY" & "1.5,4,60"
'                        strOutput = strOutput & "^BC" & "N,60,Y,N,N"
'                        strOutput = strOutput & "^FD" & strData
'                        strOutput = strOutput & "^FS" & vbLf
'                    Else
'                        'strOutput = strOutput & "^FO100,100^BQN,2,10^FDNM,AAC-42^FS" & vbLf
'                        strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                        strOutput = strOutput & "^CI26"
'                        strOutput = strOutput & "^A1" & strRot & "25" & "," & "12"
'                        strOutput = strOutput & "^BY" & "2,4,50"
'                        strOutput = strOutput & "^BQ"
'                        strOutput = strOutput & "^FD" & strData
'                        strOutput = strOutput & "^FS" & vbLf
'                    End If
'                Else
'                    strOutput = strOutput & "^FO" & strXPos & "," & strYPos
'                    strOutput = strOutput & "^CI26"
'                    strOutput = strOutput & "^A1" & strRot & ",25" & "," & "12"
'                    strOutput = strOutput & "^FD" & strType & " : " & strData
'                    strOutput = strOutput & "^FS" & vbLf
'                End If
'
'
'            Next
'
'            strOutput = strHeader & strOutput & "^XZ" & vbLf
'
''      strOutput = ""
''strOutput = strOutput & "^XA^prb^by,3,3,350^fs"
''strOutput = strOutput & "^SEE:UHANGUL.DAT^FS"
''strOutput = strOutput & "^CWJ,E:KFONT3.FNT^FS"
''strOutput = strOutput & "^FO240,350^Aj,50,50^FD서울 구로구 구로동 abcd^FS"
''strOutput = strOutput & "^FO240,390^Aj,30,30^FD010-5789-0357^FS"
''strOutput = strOutput & "^FO260,630^Aj,50,40^FD 한영유지만 ^FS"
''strOutput = strOutput & "^FO230,680^bqn,2,100^FD02) dhdhhdhhd597-5972^FS"
''strOutput = strOutput & "^FO230,680^bc^FD02) dhdhhdhhd597-5972^FS"
''strOutput = strOutput & "^PQ1"
''strOutput = strOutput & "^XZ"
'
'            comEqp.Output = strOutput
'
'        End With
        
'    Next
    
    
'    Call SetPackTrack
    
End Sub

'Private Sub SetPackTrack()
'    Dim intRow      As Integer
'    Dim intCol      As Integer
'    Dim intItemNo   As Integer
'
'
'    '-- Insert / Update 찾아오기
'    Set AdoRs = Get_BarMaster(gBarMaster.BARCD)
'
'    If AdoRs.RecordCount = 0 Then
'        'INSERT
'        If Set_Bar_Master("IN") Then
'            '상세내용 저장
'            For intRow = 1 To spdRegBarDetail.DataRowCnt
'                If Set_Bar_Detail("IN", intRow) Then
'                    Call cmdSearch_Click
'                End If
'            Next
'        End If
'    Else
'        'UPDATE
'        If Set_Bar_Master("UP") Then
'            If Set_Bar_Detail("DEL", intRow) Then
'                '상세내용 저장
'                For intRow = 1 To spdRegBarDetail.DataRowCnt
'                    If Set_Bar_Detail("IN", intRow) Then
'                        Call cmdSearch_Click
'                    End If
'                Next
'            End If
'        End If
'    End If
'
'
'End Sub


Private Sub cmdSearch_Click()
    Dim strFromDt    As String
    Dim strToDt      As String
    
    strFromDt = Format(dtpFromDate, "yyyymmdd")
    strToDt = Format(dtpToDate, "yyyymmdd")
    
    Call cmdClear_Click
    
    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub


Private Sub Form_Load()

    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle

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
        Call SetText(spdPrtReel, "생산일자", 0, 1):          .ColWidth(1) = 10
        Call SetText(spdPrtReel, "제품코드", 0, 2):          .ColWidth(2) = 0
        Call SetText(spdPrtReel, "제품명", 0, 3):            .ColWidth(3) = 10
        Call SetText(spdPrtReel, "길이", 0, 4):              .ColWidth(4) = 6
        Call SetText(spdPrtReel, "일련번호", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReel, "골정No", 0, 6):            .ColWidth(6) = 0
        Call SetText(spdPrtReel, "포장코드", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReel, "Reel수량", 0, 8):          .ColWidth(8) = 8
        Call SetText(spdPrtReel, "Roll정보", 0, 9):          .ColWidth(9) = 0
        Call SetText(spdPrtReel, "SLT No", 0, 10):           .ColWidth(10) = 6
        Call SetText(spdPrtReel, "고객사", 0, 11):           .ColWidth(11) = 10
        Call SetText(spdPrtReel, "메모", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReel, "Lot No", 0, 13):           .ColWidth(13) = 0
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
        Call SetText(spdPrtReelDetail, "순서", 0, 2):          .ColWidth(2) = 5
        Call SetText(spdPrtReelDetail, "내용", 0, 3):          .ColWidth(3) = 21
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
    
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    
    txtReelQTY.Text = ""
    txtP1From.Text = ""
    txtP1To.Text = ""
    txtP2From.Text = ""
    txtP2To.Text = ""
    gSORT = 0

End Sub

Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim pAdoRS1      As ADODB.Recordset
'    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
    Dim strPrtSide      As String
    Dim strBarCd        As String
    Dim strDate         As String
    Dim strQty          As String
    Dim strProdLabelCd  As String
    Dim strProdCd       As String
    Dim strCompCd       As String
    Dim strTmp          As String
    Dim strBarData      As String
    Dim strExMonth      As String
    Dim strExDate       As String
    Dim strProdLen      As String
    Dim strProdSize     As String
    Dim strProdTemp     As String
    Dim strGu           As String
    Dim strContents     As String
    Dim strLeft         As String
    Dim strLotNo        As String
    Dim strLotNoFull    As String
    Dim strChimeiCd     As String
    
    If Row = 0 Then
        Call SetSpreadSort(spdPrtReel)
        Exit Sub
    End If
    
    strDate = GetText(spdPrtReel, Row, 1)
    txtProdOrderDt.Text = strDate
    strProdCd = GetText(spdPrtReel, Row, 2)
    txtProdNm.Text = GetText(spdPrtReel, Row, 3)
    strProdLen = GetText(spdPrtReel, Row, 4)
    strProdLen = strProdLen * 100 '미터를 cm으로 변환
    txtProdPosNo.Text = GetText(spdPrtReel, Row, 6)
    txtPackNm.Text = GetText(spdPrtReel, Row, 7)
    txtReelQTY.Text = GetText(spdPrtReel, Row, 8)
    txtSlittingNo.Text = GetText(spdPrtReel, Row, 10)
    txtCompNm.Text = GetText(spdPrtReel, Row, 11)
    strLotNo = GetText(spdPrtReel, Row, 13)
    strProdLabelCd = GetText(spdPrtReel, Row, 15)
    strCompCd = GetText(spdPrtReel, Row, 18)
    strQty = txtReelQTY.Text
    
    txtP1From.Text = 1
    txtP1To.Text = Format((strQty / 2), "00")
    txtP2From.Text = 1
    txtP2To.Text = strQty - txtP1To.Text
    
                    
    gPackTrack.PRODCD = strProdCd   '5자리
    gPackTrack.LOTNO = strLotNo
    gPackTrack.ORDERDT = strDate    '8자리
              
    With spdPrtReelDetail
        .MaxRows = 0
    End With

    Set AdoRs = Get_LabelDetail(strProdLabelCd, "R")
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReelDetail
                .MaxRows = .MaxRows + 1
                strGu = AdoRs.Fields("LABEL_ITEM_GU").Value & ""
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 1)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReelDetail, "", .MaxRows, 3)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_GU").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_FONT").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReelDetail, txtSlittingNo.Text, .MaxRows, 9)
    
                '-- 바코드 등록 정보 찾아오기
                If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                    strBarData = ""
                    strBarData = GetBarcode(strProdCd, strCompCd, "R", strDate, strProdLen)
                    
                    'Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                    
'                    If pAdoRS1 Is Nothing Then
'                        '등록된 정보 없음
'                    Else
'                        strBarData = ""
'
'
'                        Do Until pAdoRS1.EOF
'                            strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
'                            Select Case strTmp
'                                Case "자재코드"
'                                    SQL = ""
'                                    SQL = SQL & "SELECT PROD_MATERIAL_CD, EXPIR_MONTH, PROD_SIZE, PROD_STOR_TEMP, PROD_CHIMEI_PN   " & vbCrLf
'                                    SQL = SQL & "  FROM LBL_M_PROD                      " & vbCrLf
'                                    SQL = SQL & " WHERE PROD_CD  = '" & strProdCd & "'  " & vbCrLf
'                                    SQL = SQL & "   AND COMP_CD  = '" & strCompCd & "'  " & vbCrLf
'                                    SQL = SQL & "   AND USED_YN  = 'Y'                  " & vbCrLf
'                                    Set pAdoRS2 = New ADODB.Recordset
'                                    Call GetRecordset(AdoCn, SQL, pAdoRS2, "")
'                                    If Not pAdoRS2 Is Nothing Then
'                                        If Not pAdoRS2.EOF Then
'                                            strBarData = strBarData & Trim(pAdoRS2("PROD_MATERIAL_CD") & "") '& "|"
'                                            strExMonth = pAdoRS2("EXPIR_MONTH") & ""
'                                            strExDate = DateAdd("m", strExMonth, strDate)
'                                            strExDate = DateAdd("d", -1, strExDate)
'                                            strProdSize = pAdoRS2("PROD_SIZE") & ""
'                                            strProdTemp = pAdoRS2("PROD_STOR_TEMP") & ""
'                                            strChimeiCd = pAdoRS2("PROD_CHIMEI_PN") & ""
'                                        End If
'                                    End If
'                                    pAdoRS2.Close
'                                    Set pAdoRS2 = Nothing
'
'                                Case "유효기간_년"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
'
'                                Case "유효기간_월"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
'
'                                Case "유효기간_일"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
'
'                                Case "제조일_년"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
'
'                                Case "제조일_월"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
'
'                                Case "제조일_일"
'                                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
'
'                                Case "Slitting순번"
'                                    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
'
'                                Case "Product No"
'                                    strBarData = strBarData & "P" & Format(txtSlittingNo.Text, "0") & "01"
'
'                                Case "제품길이"
'                                    strBarData = strBarData & strProdLen '& "|"
'
'                            End Select
'
'                            pAdoRS1.MoveNext
'                        Loop
'                        pAdoRS1.Close
'                    End If
                    Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                    strContents = strBarData
                
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
                
                ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Storage Temperature" Then
                    Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                    strContents = strProdTemp
                
                ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Production Date" Then
                    Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                    strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                
                ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                    Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                    strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")

                ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "SDI ACF Lot" Then
                    If strLotNo = "" Then
                        strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                    End If
                    
                    strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                    
                    strLotNo = strLotNo & strLotNoFull
                    
                    Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                    strContents = strLotNo
                
                ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material Code" Then
                    Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                    strContents = strChimeiCd
                
                End If
                
                strLeft = 0
                If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                    barReel.Alignment = bcACenter
                    barReel.Style = msSCode128B
                    barReel.Visible = True
                    barReel.Caption = strContents
                    barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                    barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                    barReel.WIDTH = 4365
                Else
'                    If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
'                        '항목명 출력여부
'                        If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
'                            strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 10
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 10
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
'                        Else
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
'                            strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 10
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 10
'                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
'                        End If
'                    Else
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
'                    End If
                End If
                    
                strContents = ""
                
                AdoRs.MoveNext
            
            End With
        Loop
    End If
    
    AdoRs.Close

'    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub


Private Function GetBarcode(ByVal pProdCd As String, ByVal pCompCd As String, ByVal pBarGu As String, ByVal pDate As String, ByVal pProdLen As String) As String
    Dim pAdoRS1     As ADODB.Recordset
    Dim pAdoRS2     As ADODB.Recordset
    Dim strBarData  As String
    Dim strTmp      As String
    Dim strExMonth  As String
    Dim strExDate   As String
    Dim strProdSize As String
    Dim strProdTemp As String
    Dim strChimeiCd As String
'    Dim strDate     As String
    
    Set pAdoRS1 = Get_BarDetail_Prt(pProdCd, pCompCd, pBarGu)
    
    If pAdoRS1 Is Nothing Then
        '등록된 정보 없음
    Else
        strBarData = ""
        Do Until pAdoRS1.EOF
            strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
            Select Case strTmp
                Case "자재코드"
                    SQL = ""
                    SQL = SQL & "SELECT PROD_MATERIAL_CD, EXPIR_MONTH, PROD_SIZE, PROD_STOR_TEMP, PROD_CHIMEI_PN   " & vbCrLf
                    SQL = SQL & "  FROM LBL_M_PROD                      " & vbCrLf
                    SQL = SQL & " WHERE PROD_CD  = '" & pProdCd & "'  " & vbCrLf
                    SQL = SQL & "   AND COMP_CD  = '" & pCompCd & "'  " & vbCrLf
                    SQL = SQL & "   AND USED_YN  = 'Y'                  " & vbCrLf
                    Set pAdoRS2 = New ADODB.Recordset
                    Call GetRecordset(AdoCn, SQL, pAdoRS2, "")
                    If Not pAdoRS2 Is Nothing Then
                        If Not pAdoRS2.EOF Then
                            strBarData = strBarData & Trim(pAdoRS2("PROD_MATERIAL_CD") & "") '& "|"
                            strExMonth = pAdoRS2("EXPIR_MONTH") & ""
                            strExDate = DateAdd("m", strExMonth, pDate)
                            strExDate = DateAdd("d", -1, strExDate)
                            strProdSize = pAdoRS2("PROD_SIZE") & ""
                            strProdTemp = pAdoRS2("PROD_STOR_TEMP") & ""
                            strChimeiCd = pAdoRS2("PROD_CHIMEI_PN") & ""
                        End If
                    End If
                    pAdoRS2.Close
                    Set pAdoRS2 = Nothing
                    
                Case "유효기간_년"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                    
                Case "유효기간_월"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                
                Case "유효기간_일"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                
                Case "제조일_년"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(pDate))
                
                Case "제조일_월"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(pDate))
                
                Case "제조일_일"
                    strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(pDate))
                
                Case "Slitting순번"
                    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                    
                Case "Product No"
                    strBarData = strBarData & "P" & Format(txtSlittingNo.Text, "0") & "01"
                
                Case "제품길이"
                    strBarData = strBarData & pProdLen
                
            End Select
            
            pAdoRS1.MoveNext
        Loop
        pAdoRS1.Close
    End If
    
    GetBarcode = strBarData
    
End Function
