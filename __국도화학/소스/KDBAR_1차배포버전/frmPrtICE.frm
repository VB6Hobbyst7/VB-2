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
   ClientWidth     =   21165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   21165
   WindowState     =   2  '최대화
   Begin VB.TextBox txtComm 
      Appearance      =   0  '평면
      Height          =   5955
      Left            =   20490
      MultiLine       =   -1  'True
      TabIndex        =   82
      Top             =   1770
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   9765
      Left            =   90
      TabIndex        =   9
      Top             =   1050
      Width           =   20115
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   8595
         Left            =   8160
         TabIndex        =   10
         Top             =   330
         Width           =   11775
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
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   84
            Top             =   5880
            Width           =   900
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
            TabIndex        =   83
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
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
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   68
            Top             =   360
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
            TabIndex        =   67
            Top             =   360
            Width           =   1500
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   3105
            Left            =   6390
            ScaleHeight     =   3045
            ScaleWidth      =   5055
            TabIndex        =   34
            Top             =   1440
            Width           =   5115
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
               Index           =   14
               Left            =   870
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
               Index           =   13
               Left            =   690
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
               Index           =   12
               Left            =   390
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
               Index           =   11
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
               Index           =   10
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
               Index           =   9
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
               Index           =   8
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
               Index           =   7
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
               Index           =   6
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
               Index           =   5
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
               Index           =   4
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
               Index           =   3
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
               Index           =   2
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
               Index           =   1
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
               Index           =   15
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
               Index           =   14
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
               Index           =   13
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
               Index           =   12
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
               Index           =   11
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
               Index           =   10
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
               Index           =   9
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
               Index           =   8
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
               Index           =   7
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
               Index           =   6
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
               Index           =   5
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
               Index           =   4
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
               Index           =   3
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
               Index           =   2
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
               Index           =   1
               Left            =   0
               TabIndex        =   37
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
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Image imgBar2 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtICE.frx":0000
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.Image imgQrBar 
               Height          =   750
               Left            =   660
               Picture         =   "frmPrtICE.frx":3FCD
               Stretch         =   -1  'True
               Top             =   3000
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Image imgBar1 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtICE.frx":A605
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   360
            Width           =   1485
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   735
            Left            =   1140
            TabIndex        =   24
            Top             =   7710
            Width           =   5145
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
               Left            =   4020
               Style           =   1  '그래픽
               TabIndex        =   28
               Top             =   150
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
               Left            =   270
               Style           =   1  '그래픽
               TabIndex        =   27
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
               Left            =   2910
               Style           =   1  '그래픽
               TabIndex        =   26
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
               Left            =   1380
               Style           =   1  '그래픽
               TabIndex        =   25
               Top             =   150
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
            Left            =   1590
            MaxLength       =   5
            TabIndex        =   23
            Top             =   4980
            Width           =   1080
         End
         Begin VB.TextBox txtReelBarcode 
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
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   22
            Text            =   "2X2707R0202001P10110000"
            Top             =   5430
            Width           =   4140
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
            Left            =   3690
            TabIndex        =   21
            Top             =   5880
            Value           =   1  '확인
            Width           =   1185
         End
         Begin VB.PictureBox picSide 
            BackColor       =   &H00FFFFFF&
            Height          =   3855
            Left            =   6390
            ScaleHeight     =   3795
            ScaleWidth      =   5055
            TabIndex        =   20
            Top             =   4620
            Width           =   5115
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   0
               Left            =   210
               Picture         =   "frmPrtICE.frx":E5D2
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   1
               Left            =   2520
               Picture         =   "frmPrtICE.frx":1259F
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   2
               Left            =   180
               Picture         =   "frmPrtICE.frx":1656C
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   3
               Left            =   2490
               Picture         =   "frmPrtICE.frx":1A539
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   4
               Left            =   180
               Picture         =   "frmPrtICE.frx":1E506
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   5
               Left            =   2490
               Picture         =   "frmPrtICE.frx":224D3
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   6
               Left            =   180
               Picture         =   "frmPrtICE.frx":264A0
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   7
               Left            =   2490
               Picture         =   "frmPrtICE.frx":2A46D
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   8
               Left            =   180
               Picture         =   "frmPrtICE.frx":2E43A
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   9
               Left            =   2490
               Picture         =   "frmPrtICE.frx":32407
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   10
               Left            =   180
               Picture         =   "frmPrtICE.frx":363D4
               Stretch         =   -1  'True
               Top             =   3030
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   11
               Left            =   2490
               Picture         =   "frmPrtICE.frx":3A3A1
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
            Left            =   1590
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   19
            Top             =   5880
            Width           =   1050
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
            Left            =   1890
            MaxLength       =   5
            TabIndex        =   18
            Top             =   4470
            Visible         =   0   'False
            Width           =   1050
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
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            Top             =   5430
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
            Height          =   405
            Left            =   4890
            TabIndex        =   15
            Top             =   5880
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
            Height          =   1455
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   6330
            Width           =   6075
         End
         Begin VB.CommandButton cmdErrClear 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "C"
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
            Left            =   210
            Style           =   1  '그래픽
            TabIndex        =   13
            Top             =   7830
            Width           =   375
         End
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
            Left            =   600
            Style           =   1  '그래픽
            TabIndex        =   12
            Top             =   7830
            Width           =   465
         End
         Begin VB.TextBox txtInBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
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
            Left            =   4500
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            Top             =   4980
            Width           =   1740
         End
         Begin FPSpread.vaSpread spdScan 
            Height          =   1455
            Left            =   210
            TabIndex        =   16
            Top             =   6330
            Visible         =   0   'False
            Width           =   6075
            _Version        =   393216
            _ExtentX        =   10716
            _ExtentY        =   2566
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
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   2
            ShadowColor     =   16774120
            SpreadDesigner  =   "frmPrtICE.frx":3E36E
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   3135
            Left            =   240
            TabIndex        =   69
            Top             =   1440
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   5530
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
            SpreadDesigner  =   "frmPrtICE.frx":3F30F
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
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
            TabIndex        =   80
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
            Left            =   210
            TabIndex        =   79
            Top             =   5880
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            Left            =   210
            TabIndex        =   73
            Top             =   4980
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
            Left            =   210
            TabIndex        =   72
            Top             =   5430
            Width           =   1350
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
            Left            =   330
            TabIndex        =   71
            Top             =   4470
            Visible         =   0   'False
            Width           =   1500
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
            Left            =   3120
            TabIndex        =   70
            Top             =   4980
            Width           =   1350
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   8505
         Left            =   210
         TabIndex        =   81
         Top             =   390
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   15002
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
         SpreadDesigner  =   "frmPrtICE.frx":3FE27
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
      Width           =   20145
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
         Format          =   131006465
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
         Format          =   131006465
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
               Picture         =   "frmPrtICE.frx":41223
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":417BD
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":41D57
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":422F1
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":42B83
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":42CDD
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":42E37
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":42F91
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtICE.frx":4386B
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
         TabIndex        =   85
         Top             =   390
         Width           =   2265
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtICE.frx":44145
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
    txtScanCount.Text = "0"

    lblstrPrtLabelName.Caption = ""
    
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
    Dim strAFont    As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strHeader   As String
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    If txtInBarcode.Text = "" Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strAFont = "^A0N,35,25"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^MD9"
    
    strOutput = ""
    strOutput = strOutput & "^FO" & "20,100"
    strOutput = strOutput & "^CI26"
    strOutput = strOutput & "^BC" & "N,60,Y,N,N"
    strOutput = strOutput & "^FD" & txtInBarcode.Text
    strOutput = strOutput & "^FS" & vbLf

            
    strOutput = strHeader & strOutput & "^XZ" & vbLf
    comEqp.Output = strOutput

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
    Dim J           As Integer
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
    
    Erase strPrtData
    Erase strTrackBC
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
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0001", "P0002"
            '폰트
            strHeader = strHeader & "^MD9"
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
                        strLotSub = "P" & strSlt & Format(intPrt, "00")
                        strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BC" & "B,100,Y,N,N"
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
            
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            strOutput = strHeader & strOutput & "^XZ" & vbLf
            '3번 출력 수정...
            For i = 1 To 3
                comEqp.Output = strOutput
            Next
            
            ReDim Preserve strTrackBC(intCnt)
            strTrackBC(intCnt) = strBarcode
            intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
            strOutput = ""
        
            '국도화학 내부관리용코드
        
        Case "P0003"
            '폰트
            strHeader = strHeader & "^MD18"
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
                    ElseIf strType = "생산일자" Then
                        strBarcode = strBarcode & Get_YMD("Y7", Year(strData))
                        strBarcode = strBarcode & Get_YMD("M2", MONTH(strData))
                        strBarcode = strBarcode & Get_YMD("D4", Day(strData))
                        strBarcode = strBarcode & Format(txtICEBoxNo.Text, "000")
                        
                        'Reel 수량 찾아오기
                        intReelCnt = 0
                        With spdScan
                            For m = 1 To .MaxRows
                                .Row = m
                                .Col = 1
                                intReelCnt = intReelCnt + Get_PPReelCount(Format(strData, "yyyymmdd"), "P0003", .Text)
                            Next
                        End With
                        
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
'                        strBarcode = strData
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
        
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
            '3번 출력 수정...
            For i = 1 To 3
                comEqp.Output = strOutput
            Next
            
            ReDim Preserve strTrackBC(intCnt)
            strTrackBC(intCnt) = strBarcode
            intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
            strOutput = ""
        
        Case "P0004", "P0005"
            
            '폰트
            strHeader = strHeader & "^MD9"
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
                        'strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        
                        strData = Mid(strData, 1, 24) & Format(100 + CCur(txtICEBoxNo.Text), "000") & Mid(strData, 28, 3)
                        
                        If Mid(strBarType, 1, 1) = "1" Then
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
        
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
            '3번 출력 수정...
            For i = 1 To 3
                comEqp.Output = strOutput
            Next
            
            ReDim Preserve strTrackBC(intCnt)
            strTrackBC(intCnt) = strBarcode
            intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
            strOutput = ""
            
        Case "P0006", "P0007"
            
            '폰트
            strHeader = strHeader & "^MD9"
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
                        strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
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
            
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            strOutput = strHeader & strOutput & "^XZ" & vbLf
            
            '3번 출력 수정...
            For i = 1 To 3
                comEqp.Output = strOutput
            Next

            blnPrint = True
            
            txtComm.Text = strOutput
            strOutput = ""
    End Select
   
    '-- PP Box라벨(바코드) 출력
    If chkReelPrint.Value = "1" Then
        Erase strPrtData
        i = 0
        
        With spdScan
            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Then
                'For intCnt = 1 To .MaxRows
                strXPos = 100: strYPos = 100
                
                i = 0
                ReDim Preserve strPrtData(i) As String
                strPrtData(i) = ""
                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                strPrtData(i) = strPrtData(i) & "^CI26"
                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                i = i + 1
                ReDim Preserve strPrtData(i) As String
                strPrtData(i) = ""
                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
                strPrtData(i) = strPrtData(i) & "^CI26"
                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                i = i + 1
                
                For intCnt = 1 To .MaxRows
                    .Row = intCnt
                    .Col = 1
                    
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
                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Next
            Else
                For intCnt = 1 To .MaxRows
                    .Row = intCnt
                    .Col = 1
                    
                    If intCnt = 1 Then
                        strXPos = 100: strYPos = 100
                    Else
                        If intCnt Mod 2 = 0 Then
                            strXPos = strPlusXPos:  strYPos = strYPos
                        Else
                            strXPos = 100:          strYPos = strYPos + 200
                        End If
                    End If
                    
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Next
            End If
        End With
        
        For J = 0 To UBound(strPrtData)
            strOutput = strOutput & strPrtData(J)
        Next
        
        strOutput = strHeader & strOutput & "^XZ" & vbLf
        comEqp.Output = strOutput
        
        ReDim Preserve strTrackBC(intCnt)
        strTrackBC(intCnt) = strBarcode
        intCnt = intCnt + 1
        
        blnPrint = True
        txtComm.Text = txtComm.Text & vbCrLf & strOutput
        strOutput = ""
        
    End If
   
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
    
        'ICE박스 M + 200302(제조일자) + 001(박스번호) (001번부터 시작)
        txtInBarcode.Text = "M" & Mid(strDate, 3, 6) & Format(intMaxNo, "000")
        
        If lblstrPrtLabelName.Caption = "P0003" Then
            'strBarcode = txtInBarcode.Text
            strAFont = "^AJN,50,30"
            strHeader = "^XA" & vbLf
            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
            strHeader = strHeader & "^PON^FS" & vbLf
            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
            strHeader = strHeader & "^MD9"
            strOutput = ""
            strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & strBarcode
            strOutput = strOutput & "^FS" & vbLf
            strOutput = strHeader & strOutput & "^XZ" & vbLf
            comEqp.Output = strOutput
        End If
        
        Call SetPackTrack(strBarcode)
    
        txtICEBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 9
            imgPpBar(intCnt).Visible = False
        Next
    End If
    
    
End Sub

Private Sub SetPackTrack(ByVal pICEBarcode As String)
    Dim intCnt      As Integer
    Dim intMaxNo    As Integer
    
    With spdScan
        For intCnt = 1 To .MaxRows
            .Row = intCnt
            .Col = 1
            gPackTrack.ORDERDT = Format(txtProdOrderDt.Text, "yyyymmdd")     'Key
            gPackTrack.PRODCD = txtProdCd.Text                               'Key
            gPackTrack.REELBAR = ""
            gPackTrack.PPBAR = .Text
            gPackTrack.ICEBAR = pICEBarcode
            gPackTrack.PPBARIN = ""
            gPackTrack.ICEBARIN = txtInBarcode.Text  'txtPPBoxNo.Text
            gPackTrack.LOTNO = txtLotNo.Text
            gPackTrack.REELPRTID = ""
            gPackTrack.REELPRTDT = ""
            gPackTrack.PPPRTID = ""
            gPackTrack.PPPRTDT = ""
            gPackTrack.ICEPRTID = gKUKDO.USERID
            gPackTrack.ICEPRTDT = ""
            
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


Private Sub cmdView_Click()
    
    If txtComm.Visible = False Then
        txtComm.Visible = True
    Else
        txtComm.Visible = False
    End If
    
End Sub

Private Sub Form_Load()

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
        Call SetText(spdPrtReel, "선택", 0, 1):              .ColWidth(1) = 4
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
        Call SetText(spdPrtReelDetail, "항목", 0, 1):          .ColWidth(1) = 10
        Call SetText(spdPrtReelDetail, "순서", 0, 2):          .ColWidth(2) = 4
        Call SetText(spdPrtReelDetail, "내용", 0, 3):          .ColWidth(3) = 32
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
    
    With spdScan
        Call SetText(spdScan, "바코드", 0, 1):          .ColWidth(1) = 25
        .MaxRows = 0
        .MaxCols = 1
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
    
    'txtReelQTY.Text = ""
    
    gSORT = 0

End Sub


Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pAdoRS1      As ADODB.Recordset
    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
    Dim J               As Integer
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
    
    txtInBarcode.Text = ""
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
    strQty = txtReelQTY.Text
    
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
                strProdName = pAdoRS2("PROD_NAME") & ""
                strProdLen = pAdoRS2("PROD_LENGTH") & ""
                strProdMatCd = pAdoRS2("PROD_MATERIAL_CD") & ""
                strExMonth = pAdoRS2("EXPIR_MONTH") & ""
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
                    Case "P0004", "P0005"
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
                                        Case "대BOX내 REEL총길이":  strBarData = strBarData & "16200"
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
                    Case "P0006", "P0007"
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
                                        Case "유효기간":    strBarData = strBarData & Format(strExDate, "yyyymmdd")
                                    
                                    
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
                                        Case "대BOX내 REEL총길이":  strBarData = strBarData & "16200"
                                        Case "유효기간":            strBarData = strBarData & strExMonth
                                    
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
                            Call SetText(spdPrtReelDetail, strMaxTot, .MaxRows, 3)
                            strContents = strMaxTot & " Reel"
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
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
    Dim i As Integer
    Dim strDate     As String
    Dim strProdLen  As String
    Dim strContents As String
    
    If KeyAscii = vbKeyReturn Then
        If txtReelBarcode.Text <> "" Then
            
            If Len(gPackTrack.ORDERDT) = 10 Then
                strDate = Format(gPackTrack.ORDERDT, "yyyymmdd")
            Else
                strDate = gPackTrack.ORDERDT
            End If
            '트래킹 테이블에 존재하는지 체크
            Set AdoRs = Get_Pack_Track(strDate, gPackTrack.PRODCD, "", txtReelBarcode.Text, "")
        
            If AdoRs.RecordCount = 0 Then
                'MsgBox txtReelBarcode.Text & "는 발행된 Reel 바코드가 아닙니다.", vbOKOnly + vbInformation, Me.Caption
                txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 발행된 PP Box 바코드가 아닙니다." & vbCrLf
                txtReelBarcode.SelStart = 0
                txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                Exit Sub
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
                .Row = .MaxRows
                .Col = 1
                .Text = txtReelBarcode.Text
            End With
            
            imgPpBar(i - 1).Visible = True
            txtScanCount.Text = txtScanCount.Text + 1
            If txtMaxTot.Text = txtScanCount.Text Then
                If chAutoPrint.Value = "1" Then
                    Call cmdPrint_Click
                End If
            End If
            txtReelBarcode.Text = ""
            
            For i = 1 To spdPrtReelDetail.MaxRows
                Select Case lblstrPrtLabelName.Caption
                    Case "P0001", "P0002"
                        '제품수량/길이
                        If GetText(spdPrtReelDetail, i, 1) = "Total Quantity/Length" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = mGetP(strProdLen, 2, "/")
                            strProdLen = Replace(strProdLen, "cm", "")
                            strContents = CCur(txtScanCount.Text) * CCur(txtMaxTot.Text) & "Reels/" & CCur(strProdLen) * CCur(txtScanCount.Text) & "cm"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                        '바코드
                        If GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 17)
                            strContents = strProdLen & Format(CCur(txtScanCount.Text) * CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text), "00000")
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case Else
                        If GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            strContents = CCur(txtScanCount.Text) * CCur(txtMaxTot.Text) * CCur(txtPPMaxTot.Text) & " Reel"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strProdLen = GetText(spdPrtReelDetail, i, 3)
                            strProdLen = Mid(strProdLen, 1, 21) & Format(txtICEBoxNo.Text, "000")
                            strContents = strProdLen & Format(CCur(txtScanCount.Text) * 270 * CCur(txtPPMaxTot.Text), "00000") & Right(GetText(spdPrtReelDetail, i, 3), 1)
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                End Select
            Next
            
        End If
    End If
    
End Sub

